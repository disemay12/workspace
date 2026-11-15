'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execSync, execFileSync } = require('child_process');
const { findForbidden, sha256File } = require('./release-rules');

const rootDir = path.resolve(__dirname, '..');
const pkg = require(path.join(rootDir, 'package.json'));
const outDir = path.join(rootDir, 'WorkSpace');
// Сборка идёт во временный каталог рядом и подменяет готовую папку в конце:
// rmSync(outDir) в начале оставлял WorkSpace/ полуразобранным при любом сбое,
// а pack-release затем упаковывал смесь старых и новых файлов.
const tmpOutDir = path.join(rootDir, `WorkSpace.tmp-${process.pid}`);
const appDir = path.join(tmpOutDir, 'app');
const resourcesDir = path.join(appDir, 'resources');
const electronDist = path.join(rootDir, 'node_modules', 'electron', 'dist');

console.log('=== Сборка портативного приложения WorkSpace ===\n');

/**
 * Удаление только внутри системного tmp: путь приходит из JSON дочернего
 * скрипта, и рекурсивное удаление не должно уметь задеть рабочий каталог.
 */
function removeInsideTmp(target) {
  const resolved = path.resolve(String(target));
  const tmpRoot = path.resolve(os.tmpdir());
  // realpath + toLowerCase: Windows не различает регистр, а ссылки в TEMP
  // раскрывались в путь, который сравнение по префиксу не узнавало.
  const realTarget = safeRealpath(resolved);
  const realTmp = safeRealpath(tmpRoot);
  const within = (base, child) => {
    const normalizedBase = (base.endsWith(path.sep) ? base : base + path.sep).toLowerCase();
    return child.toLowerCase() === base.toLowerCase() || child.toLowerCase().startsWith(normalizedBase);
  };
  if (realTarget === realTmp || !within(realTmp, realTarget)) {
    throw new Error(`Отказ удалять ${resolved}: путь вне ${tmpRoot}`);
  }
  fs.rmSync(realTarget, { recursive: true, force: true });
}

function safeRealpath(target) {
  try {
    return fs.realpathSync.native(target);
  } catch {
    return path.resolve(target);
  }
}

function copyFileIfChanged(src, dst) {
  const source = fs.statSync(src);
  try {
    const target = fs.statSync(dst);
    // size+mtime: по одному размеру файл, перезаписанный с тем же числом байт,
    // считался актуальным, и в app/ оставалась старая версия.
    if (source.size === target.size && Math.abs(source.mtimeMs - target.mtimeMs) < 1) return;
  } catch {
    // Назначения нет — копируем ниже.
  }
  fs.mkdirSync(path.dirname(dst), { recursive: true });
  fs.copyFileSync(src, dst);
  // copyFileSync ставит текущее время: без переноса mtime источника сравнение
  // выше не совпало бы никогда.
  fs.utimesSync(dst, source.atime, source.mtime);
}

function copyDirRecursive(src, dst, filter) {
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    if (filter && !filter(entry.name)) continue;
    const srcPath = path.join(src, entry.name);
    const dstPath = path.join(dst, entry.name);
    if (entry.isDirectory()) {
      copyDirRecursive(srcPath, dstPath, filter);
    } else {
      copyFileIfChanged(srcPath, dstPath);
    }
  }
}

/** Резолв CLI asar: при плоском node_modules жёсткий путь не совпадал. */
function resolveAsarCli() {
  try {
    return require.resolve('@electron/asar/bin/asar.js');
  } catch {
    throw new Error(
      'Модуль @electron/asar не найден. Он объявлен в devDependencies — выполните npm install.'
    );
  }
}

/** Компилятор C#: ищем в каталоге Windows, а не по двум захардкоженным путям. */
function findCsc() {
  const winDir = process.env.WINDIR || process.env.SystemRoot || 'C:\\Windows';
  const candidates = [
    path.join(winDir, 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe'),
    path.join(winDir, 'Microsoft.NET', 'Framework', 'v4.0.30319', 'csc.exe'),
  ];
  const found = candidates.find((candidate) => fs.existsSync(candidate));
  if (!found) {
    throw new Error(
      `Компилятор C# не найден (${candidates.join(' или ')}). ` +
        'Установите .NET Framework 4.x или соберите загрузчик вручную из scripts/launcher/Launcher.cs.'
    );
  }
  return found;
}

/** Подмена каталога сборки: старый уводим в .old-<pid>, временный переименовываем. */
function swapInPlace() {
  const backupDir = `${outDir}.old-${process.pid}`;
  const hadPrevious = fs.existsSync(outDir);
  if (hadPrevious) fs.renameSync(outDir, backupDir);
  try {
    fs.renameSync(tmpOutDir, outDir);
  } catch (err) {
    // Возвращаем прежнюю сборку: сбой подмены не должен её терять.
    if (hadPrevious) {
      try {
        fs.renameSync(backupDir, outDir);
      } catch {}
    }
    throw err;
  }
  if (hadPrevious) {
    try {
      fs.rmSync(backupDir, { recursive: true, force: true });
    } catch (err) {
      console.warn(`[build-portable] Старая сборка осталась в ${backupDir}: ${err.message}`);
    }
  }
}

/**
 * Сверка Electron до сборки: версия из node_modules должна совпадать с
 * package-lock.json, а распакованный dist — с версией пакета. Раньше рантайм
 * копировался вообще без проверки, и расхождение (dist от прошлой версии
 * после npm install) выяснялось уже у пользователя.
 */
function verifyElectronDist() {
  const pkgFile = path.join(rootDir, 'node_modules', 'electron', 'package.json');
  if (!fs.existsSync(pkgFile)) {
    throw new Error('Нет node_modules/electron: выполните npm ci.');
  }
  const installed = JSON.parse(fs.readFileSync(pkgFile, 'utf8')).version;
  const locked = JSON.parse(fs.readFileSync(path.join(rootDir, 'package-lock.json'), 'utf8')).packages?.[
    'node_modules/electron'
  ]?.version;
  if (!locked) {
    throw new Error('В package-lock.json нет записи node_modules/electron: обновите lock-файл.');
  }
  if (installed !== locked) {
    throw new Error(
      `Electron в node_modules (${installed}) не совпадает с package-lock.json (${locked}): выполните npm ci.`
    );
  }

  const distVersionFile = path.join(electronDist, 'version');
  if (!fs.existsSync(distVersionFile) || !fs.existsSync(path.join(electronDist, 'electron.exe'))) {
    throw new Error('node_modules/electron/dist не распакован: выполните node node_modules/electron/install.js.');
  }
  // dist/version пишет установщик Electron: другое значение означает, что
  // бинарник остался от предыдущей версии пакета.
  const distVersion = fs.readFileSync(distVersionFile, 'utf8').trim();
  if (distVersion !== installed) {
    throw new Error(
      `node_modules/electron/dist от версии ${distVersion}, а пакет electron — ${installed}: удалите dist и выполните npm ci.`
    );
  }
  return installed;
}

/**
 * Манифест рядом с артефактами: фиксирует версию рантайма (иначе её видно
 * только в node_modules) и хеш архива, если он уже лежит в release/.
 */
function writeBuildManifest(electronVersion) {
  const releaseDir = path.join(rootDir, 'release');
  fs.mkdirSync(releaseDir, { recursive: true });
  const zipName = `WorkSpace-${pkg.version}-win-x64.zip`;
  const zipPath = path.join(releaseDir, zipName);
  const hasArchive = fs.existsSync(zipPath);
  const manifest = {
    appVersion: pkg.version,
    electronVersion,
    archive: hasArchive ? zipName : null,
    archiveSha256: hasArchive ? sha256File(zipPath) : null,
    builtAt: new Date().toISOString(),
  };
  const manifestPath = path.join(releaseDir, 'build-manifest.json');
  fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
  return manifestPath;
}

function main() {
  // 0. Сверка рантайма до пятиминутной сборки
  console.log('[0/5] Сверка node_modules/electron...');
  const electronVersion = verifyElectronDist();
  console.log(`      Electron ${electronVersion}: dist и package-lock.json совпадают`);

  // 1. Сборка Vite/Electron (иконка генерируется prebuild-хуком npm)
  console.log('[1/5] Компиляция фронтенда и главного процесса (npm run build)...');
  execSync('npm run build', { cwd: rootDir, stdio: 'inherit' });

  // 2. Подготовка файлов для app.asar
  console.log('\n[2/5] Подготовка файлов пакета (scripts/stage-portable.js)...');
  const stageOutput = execFileSync(
    process.execPath,
    [path.join(rootDir, 'scripts', 'stage-portable.js')],
    { cwd: rootDir, encoding: 'utf8' }
  );
  const { stage } = JSON.parse(stageOutput.trim());
  console.log(`      Файлы подготовлены во временной папке: ${stage}`);

  // 3. Упаковка в app.asar
  console.log('\n[3/5] Упаковка в resources/app.asar...');
  const tempAsar = path.join(os.tmpdir(), `ws-app-${process.pid}.asar`);
  const asarCli = resolveAsarCli();
  try {
    // Секреты ищутся в stage ДО asar pack: внутри app.asar удалять уже нечего,
    // поэтому прежний список путей после упаковки был бесполезной защитой.
    const leaks = findForbidden(stage);
    if (leaks.length > 0) {
      // Находку не вычищаем молча: упавшая сборка заметна, а «почищенный»
      // перед упаковкой ключ — нет (он уже лежит в рабочем дереве).
      throw new Error(
        `В пакет попали запрещённые файлы (${leaks.join(', ')}). ` +
          'Уберите их из источников сборки; правила — scripts/release-rules.js.'
      );
    }

    // asar запускается как JS-файл под node, без shell: путь к stage приходит из
    // mkdtempSync и может содержать пробелы — склейка в shell-строку на этом ломалась.
    execFileSync(process.execPath, [asarCli, 'pack', stage, tempAsar], { stdio: 'inherit' });

    // 4. Развертывание среды Electron в WorkSpace/app
    console.log('\n[4/5] Развертывание среды Electron в WorkSpace/app...');
    fs.rmSync(tmpOutDir, { recursive: true, force: true });
    fs.mkdirSync(resourcesDir, { recursive: true });

    // Копируем бинарные компоненты Electron (без дефолтного electron.exe и resources)
    copyDirRecursive(electronDist, appDir, (name) => name !== 'electron.exe' && name !== 'resources');

    // Копируем исполняемый файл Electron как WorkSpace.exe внутри app/
    copyFileIfChanged(path.join(electronDist, 'electron.exe'), path.join(appDir, 'WorkSpace.exe'));

    // Копируем свежий app.asar
    fs.copyFileSync(tempAsar, path.join(resourcesDir, 'app.asar'));

    // Очищаем старые бэкапы app.asar.bak-*
    for (const file of fs.readdirSync(resourcesDir)) {
      if (file.startsWith('app.asar.bak-')) {
        try { fs.unlinkSync(path.join(resourcesDir, file)); } catch {}
      }
    }

    // README релиза: одинаковый текст для архива и для папки сборки.
    const readmeSrc = path.join(rootDir, 'scripts', 'release-readme.txt');
    if (fs.existsSync(readmeSrc)) {
      fs.copyFileSync(readmeSrc, path.join(tmpOutDir, 'README.txt'));
      fs.copyFileSync(readmeSrc, path.join(appDir, 'README.txt'));
    }
  } finally {
    // Удаление не должно перетирать основную ошибку сборки и не оставлять мусор
    // ни на успехе, ни на сбое.
    try {
      removeInsideTmp(stage);
    } catch (err) {
      console.warn(`[build-portable] Временный stage не удалён: ${err.message}`);
    }
    try {
      fs.unlinkSync(tempAsar);
    } catch {}
  }

  // 5. Компиляция оконного загрузчика WorkSpace.exe
  console.log('\n[5/5] Компиляция загрузчика WorkSpace.exe (C#)...');
  const csc = findCsc();

  const launcherSrc = path.join(rootDir, 'scripts', 'launcher', 'Launcher.cs');
  const launcherOut = path.join(tmpOutDir, 'WorkSpace.exe');
  const iconPath = path.join(rootDir, 'build', 'icon.ico');
  const iconArgs = fs.existsSync(iconPath) ? [`/win32icon:${iconPath}`] : [];
  execFileSync(
    csc,
    ['/nologo', '/target:winexe', '/optimize+', ...iconArgs, `/out:${launcherOut}`, launcherSrc],
    { stdio: 'inherit' }
  );

  const launcherSize = (fs.statSync(launcherOut).size / 1024).toFixed(1);
  const asarSize = (fs.statSync(path.join(resourcesDir, 'app.asar')).size / (1024 * 1024)).toFixed(2);

  // Готовую папку подменяем целиком: до этого момента WorkSpace/ не тронута.
  swapInPlace();

  const manifestPath = writeBuildManifest(electronVersion);

  console.log('\n========================================================');
  console.log('  СБОРКА УСПЕШНО ЗАВЕРШЕНА В ПАПКУ WorkSpace');
  console.log('========================================================');
  console.log(`Папка сборки: ${outDir}`);
  console.log(`Загрузчик:    ${path.join(outDir, 'WorkSpace.exe')} (${launcherSize} KB)`);
  console.log(`Приложение:   ${path.join(outDir, 'app', 'WorkSpace.exe')}`);
  console.log(`Бандл asar:   ${path.join(outDir, 'app', 'resources', 'app.asar')} (${asarSize} MB)`);
  console.log(`Манифест:     ${manifestPath} (Electron ${electronVersion})`);
  console.log('========================================================\n');
}

try {
  main();
} catch (err) {
  // Предыдущая сборка цела: убираем только временный каталог.
  try {
    fs.rmSync(tmpOutDir, { recursive: true, force: true });
  } catch {}
  console.error(`\n[build-portable] Сборка не удалась: ${err.message}`);
  process.exit(1);
}
