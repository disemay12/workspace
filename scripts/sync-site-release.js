/**
 * Синхронизация сайта с последним релизом.
 *
 * docs/index.html держит версию, ссылки на ассет и размер архива статически —
 * при выпуске новой версии их надо обновить, иначе посетители скачивают старый
 * билд. Скрипт берёт версию из package.json и размер готового архива из dist/.
 *
 * Запуск: npm run site:sync (после pack:release).
 */
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const version = pkg.version;
const siteFile = path.join(root, 'docs', 'index.html');

if (!fs.existsSync(siteFile)) {
  console.error('docs/index.html не найден');
  process.exit(1);
}

const zipName = `WorkSpace-${version}-win-x64.zip`;
const zipPath = path.join(root, 'dist', zipName);
const sizeMb = fs.existsSync(zipPath)
  ? Math.round(fs.statSync(zipPath).size / (1024 * 1024))
  : null;

let html = fs.readFileSync(siteFile, 'utf8');
const before = html;

// Версия в JSON-LD и подписях, ссылки на ассет и его размер.
html = html.replace(
  /releases\/download\/v[\d.]+\/WorkSpace-[\d.]+-win-x64\.zip/g,
  `releases/download/v${version}/${zipName}`
);
html = html.replace(/"softwareVersion":\s*"[\d.]+"/g, `"softwareVersion": "${version}"`);
html = html.replace(/WorkSpace-[\d.]+-win-x64\.zip/g, zipName);
html = html.replace(/(<span class="hero__note">)[\d.]+(, архив)/g, `$1${version}$2`);
if (sizeMb !== null) {
  html = html.replace(/data-count="\d+" data-suffix=" МБ"/g, `data-count="${sizeMb}" data-suffix=" МБ"`);
  html = html.replace(/(\d+) МБ<\/span>, win-x64/g, `${sizeMb} МБ</span>, win-x64`);
  html = html.replace(/(WorkSpace-[\d.]+-win-x64\.zip, )\d+ МБ/g, `$1${sizeMb} МБ`);
}

if (html === before) {
  console.log(`Сайт уже указывает на ${version}${sizeMb !== null ? ` (${sizeMb} МБ)` : ''}`);
} else {
  fs.writeFileSync(siteFile, html, 'utf8');
  console.log(`Сайт обновлён: версия ${version}${sizeMb !== null ? `, архив ${sizeMb} МБ` : ', размер архива не найден — пропущен'}`);
}
