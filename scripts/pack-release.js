'use strict';

/**
 * Собирает релизный архив (release/) только из основных файлов сборки:
 *   WorkSpace.exe, README.txt и app\** (Electron + app.asar).
 *
 * Настройки, ключи и сессии в архив не попадают ни при каких условиях:
 * они лежат в %APPDATA%\WorkSpace и создаются приложением при первом запуске.
 * Любой посторонний файл в папке сборки (ai-settings.json, *.local.ps1,
 * gdrive.json, workspace-session.json, .env*, логи) отбрасывается по списку
 * из scripts/release-rules.js — он общий с build-portable.js.
 *
 * Целостность: содержимое архива проверяется по обязательным файлам, рядом
 * с zip пишется <имя>.zip.sha256.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const { isForbidden, readZipEntries, sha256File } = require('./release-rules');
const pkg = require('../package.json');

const root = path.resolve(__dirname, '..');
const buildDir = path.join(root, 'WorkSpace');
// Архив в release/, а не в dist/: vite build запускается с emptyOutDir:true
// и стирал dist целиком вместе с уже собранным релизом.
const outDir = path.join(root, 'release');
const outFile = path.join(outDir, `WorkSpace-${pkg.version}-win-x64.zip`);
// Упаковка идёт во временный файл: неудачный запуск не должен уничтожать
// предыдущий релиз (раньше outFile удалялся до запуска архиватора).
// Суффикс .zip обязателен: резервный Compress-Archive другие расширения
// отвергает (".tmp — неподдерживаемый формат файла архива").
const tempFile = `${outFile}.tmp.zip`;

const ROOT_FILES = ['WorkSpace.exe', 'README.txt'];

// Без этих файлов архив бесполезен: раньше проверялся только непустой размер,
// и архив без app.asar (или с чужим загрузчиком) уезжал на сайт.
// Пути указаны относительно верхней папки WorkSpace/ внутри zip.
const REQUIRED_ENTRIES = ['WorkSpace.exe', 'app/WorkSpace.exe', 'app/resources/app.asar'];

function collect() {
  const files = [];
  for (const name of ROOT_FILES) {
    const full = path.join(buildDir, name);
    if (!fs.existsSync(full)) throw new Error(`Нет обязательного файла сборки: WorkSpace\\${name}`);
    files.push(path.join('WorkSpace', name));
  }

  const appDir = path.join(buildDir, 'app');
  if (!fs.existsSync(appDir)) throw new Error('Нет папки сборки: WorkSpace\\app');

  const walk = (dir) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      const relative = path.relative(root, full);
      if (isForbidden(relative)) continue;
      if (entry.isDirectory()) walk(full);
      else if (entry.isFile()) files.push(relative);
    }
  };
  walk(appDir);
  return files;
}

function find7z() {
  const candidates = [
    'C:\\Program Files\\7-Zip\\7z.exe',
    'C:\\Program Files (x86)\\7-Zip\\7z.exe',
    path.join(process.env.ProgramFiles || '', '7-Zip', '7z.exe'),
  ];
  return candidates.find((candidate) => candidate && fs.existsSync(candidate)) || null;
}

function packWith7z(files, sevenZip) {
  const listFile = path.join(os.tmpdir(), `workspace-pack-${process.pid}.txt`);
  // BOM + -scsUTF-8: без них 7-Zip читает список в OEM-кодировке консоли и
  // теряет кириллические имена файлов.
  fs.writeFileSync(listFile, `\uFEFF${files.join('\r\n')}`, 'utf8');
  try {
    execFileSync(sevenZip, ['a', '-tzip', '-mx=5', '-scsUTF-8', tempFile, `@${listFile}`], {
      cwd: root,
      stdio: 'inherit',
    });
  } finally {
    try {
      fs.unlinkSync(listFile);
    } catch {
      // временный список уже удалён
    }
  }
}

/**
 * Строковый литерал PowerShell в одинарных кавычках: ' внутри значения
 * удваивается. Без экранирования апостроф в пути (os.tmpdir() бывает с
 * «O'Brien») закрывал строку и превращал остаток пути в код.
 */
const psQuote = (value) => `'${String(value).replace(/'/g, "''")}'`;

function packWithPowerShell(files) {
  // Запасной путь без 7-Zip: складываем отобранные файлы во временную папку
  // с тем же префиксом WorkSpace\ и архивируем её.
  const stage = fs.mkdtempSync(path.join(os.tmpdir(), 'workspace-stage-'));
  try {
    for (const relative of files) {
      const target = path.join(stage, relative);
      fs.mkdirSync(path.dirname(target), { recursive: true });
      fs.copyFileSync(path.join(root, relative), target);
    }
    execFileSync(
      'powershell',
      [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-Command',
        `Compress-Archive -Path ${psQuote(path.join(stage, 'WorkSpace'))} -DestinationPath ${psQuote(tempFile)} -CompressionLevel Optimal -Force`,
      ],
      { stdio: 'inherit' }
    );
  } finally {
    fs.rmSync(stage, { recursive: true, force: true });
  }
}

/**
 * Проверка содержимого готового архива. Верхняя папка WorkSpace/ отбрасывается:
 * в zip она есть всегда (список файлов и Compress-Archive кладут её внутрь),
 * а сравниваем пути относительно неё.
 */
async function verifyArchive(file) {
  const entries = await readZipEntries(file);
  const relative = entries.map((entry) =>
    entry.startsWith('WorkSpace/') ? entry.slice('WorkSpace/'.length) : entry
  );
  const missing = REQUIRED_ENTRIES.filter((required) => !relative.includes(required));
  if (missing.length > 0) {
    throw new Error(
      `В архиве нет обязательных файлов (${missing.join(', ')}): проверьте, что сборка WorkSpace/ полная.`
    );
  }
  return entries.length;
}

async function main() {
  if (!fs.existsSync(buildDir)) {
    throw new Error('Нет папки WorkSpace: сначала выполните npm run build:portable');
  }

  const files = collect();
  fs.mkdirSync(outDir, { recursive: true });
  try { fs.unlinkSync(tempFile); } catch {}
  try {
    const sevenZip = find7z();
    if (sevenZip) packWith7z(files, sevenZip);
    else packWithPowerShell(files);

    const sizeBytes = fs.statSync(tempFile).size;
    if (!sizeBytes) throw new Error('Архив получился пустым');
    const entryCount = await verifyArchive(tempFile);
    // Предыдущий релиз заменяется только после проверки содержимого.
    fs.renameSync(tempFile, outFile);

    // Формат sha256sum, чтобы сумму можно было проверить командой
    // sha256sum -c без распаковки.
    const digest = sha256File(outFile);
    const shaFile = `${outFile}.sha256`;
    fs.writeFileSync(shaFile, `${digest}  ${path.basename(outFile)}\n`, 'utf8');

    const size = (sizeBytes / (1024 * 1024)).toFixed(1);
    console.log(`Готово: ${path.relative(root, outFile)} (${size} МБ, ${files.length} файлов, ${entryCount} записей в архиве)`);
    console.log(`sha256:  ${digest}  (${path.relative(root, shaFile)})`);
    console.log('В архив не входят: ai-settings.json, secrets*.ps1, gdrive.json, workspace-session.json, .env*, логи.');
  } finally {
    try { fs.unlinkSync(tempFile); } catch {}
  }
}

main().catch((err) => {
  console.error(`[pack-release] ${err.message}`);
  process.exit(1);
});
