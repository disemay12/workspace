'use strict';

const fs = require('fs');
const path = require('path');
const { execSync, execFileSync } = require('child_process');

const rootDir = path.resolve(__dirname, '..');
const outDir = path.join(rootDir, 'WorkSpace');
const appDir = path.join(outDir, 'app');
const resourcesDir = path.join(appDir, 'resources');
const electronDist = path.join(rootDir, 'node_modules', 'electron', 'dist');

console.log('=== Сборка портативного приложения WorkSpace ===\n');

// 1. Сборка Vite/Electron (иконка генерируется prebuild-хуком npm)
console.log('[1/5] Компиляция фронтенда и главного процесса (npm run build)...');
execSync('npm run build', { cwd: rootDir, stdio: 'inherit' });

// 2. Подготовка файлов для app.asar
console.log('\n[2/5] Подготовка файлов пакета (scripts/stage-portable.js)...');
const stageOutput = execFileSync(process.execPath, [path.join(rootDir, 'scripts', 'stage-portable.js')], {
  cwd: rootDir,
  encoding: 'utf8',
});
const { stage } = JSON.parse(stageOutput.trim());
console.log(`      Файлы подготовлены во временной папке: ${stage}`);

// 3. Упаковка в app.asar
console.log('\n[3/5] Упаковка в resources/app.asar...');
const tempAsar = path.join(process.env.TEMP || 'C:/Temp', 'ws-app.asar');
const asarBin = path.join(rootDir, 'node_modules', '.bin', process.platform === 'win32' ? 'asar.cmd' : 'asar');
execSync(`"${asarBin}" pack "${stage}" "${tempAsar}"`, { stdio: 'inherit' });

// 4. Подготовка структуры папки WorkSpace/app и копирование среды Electron
console.log('\n[4/5] Развертывание среды Electron в WorkSpace/app...');
fs.mkdirSync(resourcesDir, { recursive: true });

function copyFileIfChanged(src, dst) {
  if (!fs.existsSync(dst) || fs.statSync(src).size !== fs.statSync(dst).size) {
    fs.mkdirSync(path.dirname(dst), { recursive: true });
    fs.copyFileSync(src, dst);
  }
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

// Копируем бинарные компоненты Electron (без дефолтного electron.exe и resources)
copyDirRecursive(electronDist, appDir, (name) => name !== 'electron.exe' && name !== 'resources');

// Копируем исполняемый файл Electron как WorkSpace.exe внутри app/
copyFileIfChanged(path.join(electronDist, 'electron.exe'), path.join(appDir, 'WorkSpace.exe'));

// Копируем свежий app.asar
fs.copyFileSync(tempAsar, path.join(resourcesDir, 'app.asar'));
try { fs.unlinkSync(tempAsar); } catch {}

// Очищаем старые бэкапы app.asar.bak-*
for (const file of fs.readdirSync(resourcesDir)) {
  if (file.startsWith('app.asar.bak-')) {
    try { fs.unlinkSync(path.join(resourcesDir, file)); } catch {}
  }
}

// Никакие локальные секреты (secrets.local.ps1, ai-settings.json) НЕ должны попадать в сборку
const sensitiveFiles = [
  path.join(appDir, 'secrets.local.ps1'),
  path.join(outDir, 'secrets.local.ps1'),
  path.join(appDir, 'ai-settings.json'),
  path.join(outDir, 'ai-settings.json'),
];
for (const file of sensitiveFiles) {
  if (fs.existsSync(file)) {
    try { fs.unlinkSync(file); } catch {}
  }
}

// 5. Компиляция оконного загрузчика WorkSpace.exe
console.log('\n[5/5] Компиляция загрузчика WorkSpace.exe (C#)...');
const cscPaths = [
  'C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe',
  'C:\\Windows\\Microsoft.NET\\Framework\\v4.0.30319\\csc.exe',
];
const csc = cscPaths.find((p) => fs.existsSync(p));
if (!csc) {
  throw new Error('Компилятор csc.exe не найден в системе.');
}

const launcherSrc = path.join(rootDir, 'scripts', 'launcher', 'Launcher.cs');
const launcherOut = path.join(outDir, 'WorkSpace.exe');
const iconPath = path.join(rootDir, 'build', 'icon.ico');
const iconArgs = fs.existsSync(iconPath) ? [`/win32icon:${iconPath}`] : [];
execFileSync(csc, ['/nologo', '/target:winexe', '/optimize+', ...iconArgs, `/out:${launcherOut}`, launcherSrc], {
  stdio: 'inherit',
});

const launcherSize = (fs.statSync(launcherOut).size / 1024).toFixed(1);
const asarSize = (fs.statSync(path.join(resourcesDir, 'app.asar')).size / (1024 * 1024)).toFixed(2);

console.log('\n========================================================');
console.log('  СБОРКА УСПЕШНО ЗАВЕРШЕНА В ПАПКУ WorkSpace');
console.log('========================================================');
console.log(`Папка сборки: ${outDir}`);
console.log(`Загрузчик:    ${launcherOut} (${launcherSize} KB)`);
console.log(`Приложение:   ${path.join(appDir, 'WorkSpace.exe')}`);
console.log(`Бандл asar:   ${path.join(resourcesDir, 'app.asar')} (${asarSize} MB)`);
console.log('========================================================\n');
