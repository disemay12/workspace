/**
 * Синхронизация сайта с последним релизом.
 *
 * docs/index.html держит версию, ссылки на ассет и размер архива статически —
 * при выпуске новой версии их надо обновить, иначе посетители скачивают старый
 * билд. Скрипт берёт версию из package.json и размер готового архива из release/
 *
 * Запуск: npm run site:sync (после pack:release).
 * Без собранного архива (локальный прогон): node scripts/sync-site-release.js --allow-missing-archive.
 */
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
// Локальный прогон без release/ не должен выглядеть как успешная синхронизация
// в CI: код возврата 1 остаётся, флаг снимает только отсутствие архива.
const allowMissingArchive = process.argv.includes('--allow-missing-archive');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const version = pkg.version;
const siteFile = path.join(root, 'docs', 'index.html');

if (!fs.existsSync(siteFile)) {
  console.error('docs/index.html не найден');
  process.exit(1);
}

const zipName = `WorkSpace-${version}-win-x64.zip`;
const zipPath = path.join(root, 'release', zipName);
const sizeMb = fs.existsSync(zipPath)
  ? Math.round(fs.statSync(zipPath).size / (1024 * 1024))
  : null;

let html = fs.readFileSync(siteFile, 'utf8');
const before = html;
let replaced = 0;
const notFound = [];

/**
 * Замена с подсчётом: «html === before» одинаково верно и когда всё уже
 * актуально, и когда шаблон в разметке не найден — раньше эти случаи
 * склеивались в ложное «сайт уже указывает на версию».
 */
function apply(re, build, label) {
  const matches = html.match(re);
  if (!matches) {
    notFound.push(label);
    return;
  }
  html = html.replace(re, (...args) => build(...args));
  replaced += matches.length;
}

// Версия в JSON-LD и подписях, ссылки на ассет и его размер.
apply(
  /releases\/download\/v[\d.]+\/WorkSpace-[\d.]+-win-x64\.zip/g,
  () => `releases/download/v${version}/${zipName}`,
  'ссылка на релиз'
);
apply(/"softwareVersion":\s*"[\d.]+"/g, () => `"softwareVersion": "${version}"`, 'softwareVersion');
apply(/WorkSpace-[\d.]+-win-x64\.zip/g, () => zipName, 'имя архива');
apply(/(<span class="hero__note">)[\d.]+(, архив)/g, (_m, head, tail) => `${head}${version}${tail}`, 'версия в подписи');
if (sizeMb !== null) {
  apply(/data-count="\d+" data-suffix=" МБ"/g, () => `data-count="${sizeMb}" data-suffix=" МБ"`, 'размер в счётчике');
  apply(/(\d+) МБ<\/span>, win-x64/g, () => `${sizeMb} МБ</span>, win-x64`, 'размер в подписи');
  apply(/(WorkSpace-[\d.]+-win-x64\.zip, )\d+ МБ/g, (_m, head) => `${head}${sizeMb} МБ`, 'размер рядом с архивом');
}

if (replaced === 0) {
  console.error(
    `Сайт НЕ обновлён: в docs/index.html не найдено ни одного шаблона (${notFound.join(', ')}). ` +
      'Проверьте разметку — возможно, изменились классы или формат ссылки.'
  );
  process.exit(1);
}

// Частичное обновление — незавершённая синхронизация: часть шаблонов в
// разметке не найдена, и сайт остался бы с прежней версией или размером.
if (notFound.length > 0) {
  console.error(
    `Сайт обновлён частично: не найдены шаблоны (${notFound.join(', ')}). ` +
      'Проверьте разметку docs/index.html.'
  );
  process.exitCode = 1;
}

if (html === before) {
  console.log(`Сайт уже указывает на ${version}${sizeMb !== null ? ` (${sizeMb} МБ)` : ''}`);
} else {
  fs.writeFileSync(siteFile, html, 'utf8');
  console.log(`Сайт обновлён: версия ${version}${sizeMb !== null ? `, архив ${sizeMb} МБ` : ', размер архива не найден — пропущен'}`);
}
if (sizeMb === null) {
  if (allowMissingArchive) {
    // Архив ещё не собран: размеры могли остаться от прошлой версии.
    console.warn(`Архив ${zipName} не найден в release/ — размер на сайте не обновлён (--allow-missing-archive).`);
  } else {
    console.error(
      `Архив ${zipName} не найден в release/: размер на сайте не обновлён. ` +
        'Сначала выполните npm run pack:release или запустите с --allow-missing-archive.'
    );
    process.exitCode = 1;
  }
}
