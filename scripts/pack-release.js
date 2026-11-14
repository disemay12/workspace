'use strict';

/**
 * Собирает релизный архив (release/) только из основных файлов сборки:
 *   WorkSpace.exe, README.txt и app\** (Electron + app.asar).
 *
 * Настройки, ключи и сессии в архив не попадают ни при каких условиях:
 * они лежат в %APPDATA%\WorkSpace и создаются приложением при первом запуске.
 * Любой посторонний файл в папке сборки (ai-settings.json, *.local.ps1,
 * gdrive.json, workspace-session.json, логи) отбрасывается по списку.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
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
const FORBIDDEN = [
  /(^|[\\/])ai-settings\.json$/i,
  /(^|[\\/])gdrive\.json$/i,
  /(^|[\\/])workspace-session\.json$/i,
  /\.local\.ps1$/i,
  /\.log$/i,
  /\.ws-tmp$/i,
  /(^|[\\/])secrets[^\\/]*$/i,
];

const forbidden = (relative) => FORBIDDEN.some((re) => re.test(relative));

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
      if (forbidden(relative)) continue;
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
  fs.writeFileSync(listFile, files.join('\r\n'), 'utf8');
  try {
    execFileSync(sevenZip, ['a', '-tzip', '-mx=5', tempFile, `@${listFile}`], {
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
        `Compress-Archive -Path '${path.join(stage, 'WorkSpace')}' -DestinationPath '${tempFile}' -CompressionLevel Optimal -Force`,
      ],
      { stdio: 'inherit' }
    );
  } finally {
    fs.rmSync(stage, { recursive: true, force: true });
  }
}

function main() {
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
    // Предыдущий релиз заменяется только после проверки размера.
    fs.renameSync(tempFile, outFile);

    const size = (sizeBytes / (1024 * 1024)).toFixed(1);
    console.log(`Готово: ${path.relative(root, outFile)} (${size} МБ, ${files.length} файлов)`);
    console.log('В архив не входят: ai-settings.json, secrets*.ps1, gdrive.json, workspace-session.json, логи.');
  } finally {
    try { fs.unlinkSync(tempFile); } catch {}
  }
}

main();
