'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');

const srcRoot = path.resolve(__dirname, '..');
// Уникальный каталог в системном tmp: фиксированное имя позволяло двум запускам
// сборки одновременно писать и удалять один и тот же stage.
const stage = fs.mkdtempSync(path.join(os.tmpdir(), 'ws-app-stage-'));

function copyFile(from, to) {
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(from, to);
}

function copyDir(from, to, skipName) {
  for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
    if (skipName && skipName(entry.name)) continue;
    const source = path.join(from, entry.name);
    const target = path.join(to, entry.name);
    const stat = fs.lstatSync(source);
    if (stat.isSymbolicLink()) continue;
    if (stat.isDirectory()) {
      copyDir(source, target, skipName);
    } else {
      copyFile(source, target);
    }
  }
}

copyDir(path.join(srcRoot, 'dist'), path.join(stage, 'dist'), (name) => name === 'dist');

const keep = new Set([
  'main.js',
  'preload.js',
  'watcher.js',
  'terminal.js',
  'aiProxy.js',
  'aiSettings.js',
  'docText.js',
  'share.js',
  'gdrive.js',
  'googleAuth.js',
  'video.js',
  'youtube.js',
  'imageHost.js',
  'secrets.js',
  'contextMenu.js',
  'updater.js',
  // Импортируется main.ts, googleAuth.ts и imageHost.ts; без него пакет
  // падал бы на старте (проверка полноты ниже ловит такие пропуски).
  'safeOpen.js',
  // Границы путей: main.ts, video.ts, imageHost.ts.
  'fsGuard.js',
  // Контракт ответа update:check: updater.ts (main) и проверочный стенд.
  'updateInfo.js',
]);

for (const entry of fs.readdirSync(path.join(srcRoot, 'dist-electron'), { withFileTypes: true })) {
  if (entry.isFile() && keep.has(entry.name)) {
    copyFile(
      path.join(srcRoot, 'dist-electron', entry.name),
      path.join(stage, 'dist-electron', entry.name)
    );
  }
}

const skipDocs = (name) =>
  name === '.github' ||
  name === 'test' ||
  name === 'tests' ||
  /\.(md|markdown|map)$/i.test(name);

copyDir(path.join(srcRoot, 'node_modules', 'iconv-lite'), path.join(stage, 'node_modules', 'iconv-lite'), skipDocs);
copyDir(path.join(srcRoot, 'node_modules', 'safer-buffer'), path.join(stage, 'node_modules', 'safer-buffer'), skipDocs);

/**
 * Контроль полноты пакета: список keep — ручной allowlist, поэтому новый модуль
 * electron/ молча выпадал из asar и приложение падало уже у пользователя.
 * Обход BFS по require('./X') от main.js и preload.js: проверка только main.js
 * не видела модули, которые подключают googleAuth.js/imageHost.js/video.js.
 */
const entryPoints = ['main.js', 'preload.js'];
const visited = new Set();
const queue = [...entryPoints];
const missing = [];
while (queue.length > 0) {
  const name = queue.shift();
  if (visited.has(name)) continue;
  visited.add(name);
  const base = path.join(stage, 'dist-electron', name);
  const file = fs.existsSync(base) ? base : `${base}.js`;
  if (!fs.existsSync(file)) {
    missing.push(name);
    break;
  }
  const source = fs.readFileSync(file, 'utf-8');
  for (const match of source.matchAll(/require\(\s*['"]\.\/([\w.-]+)['"]\s*\)/g)) {
    const dependency = match[1].replace(/\.js$/, '');
    if (!visited.has(dependency)) queue.push(dependency);
  }
}
if (missing.length > 0) {
  throw new Error(
    `В пакет не попал модуль, который требуется из dist-electron/${entryPoints.join('/')}: ${missing.join(', ')}. ` +
      'Добавьте его в список keep в scripts/stage-portable.js.'
  );
}

const { version } = require(path.join(srcRoot, 'package.json'));

fs.writeFileSync(
  path.join(stage, 'package.json'),
  JSON.stringify({ name: 'workspace', version, main: 'dist-electron/main.js' }, null, 2)
);

let count = 0;
let bytes = 0;
(function walk(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full);
    else {
      count += 1;
      bytes += fs.statSync(full).size;
    }
  }
})(stage);

console.log(JSON.stringify({ stage, count, bytes }));
