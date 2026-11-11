'use strict';

const fs = require('fs');
const path = require('path');

const srcRoot = path.resolve(__dirname, '..');
const stage = path.join(process.env.TEMP || 'C:/Temp', 'ws-app-stage');

fs.rmSync(stage, { recursive: true, force: true });
fs.mkdirSync(stage, { recursive: true });

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
