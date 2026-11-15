/**
 * Синхронизация сайта с последним релизом.
 *
 * docs/index.html держит версию, ссылки на ассет и размер архива статически —
 * при выпуске новой версии их надо обновить, иначе посетители скачивают старый
 * билд. Скрипт берёт версию из package.json и размер готового архива из release/
 *
 * Запуск: npm run site:sync (после pack:release).
 * Проверка без записи (CI): node scripts/sync-site-release.js --check —
 * расхождение с package.json даёт код 1, файл не меняется.
 */
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const checkOnly = process.argv.includes('--check');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const version = pkg.version;
const siteFile = path.join(root, 'docs', 'index.html');

if (!fs.existsSync(siteFile)) {
  console.error('docs/index.html не найден');
  process.exit(1);
}

const zipName = `WorkSpace-${version}-win-x64.zip`;
const zipPath = path.join(root, 'release', zipName);
// Наличие архива проверяется ДО любых замен: иначе ссылка и версия
// переписывались на zip, которого нет, а размер оставался от прошлого выпуска.
const hasArchive = fs.existsSync(zipPath);
const sizeMb = hasArchive ? Math.round(fs.statSync(zipPath).size / (1024 * 1024)) : null;

if (!hasArchive && !checkOnly) {
  console.error(
    `Архив ${zipName} не найден в release/: сайт не изменён — ссылка вела бы на несуществующий файл. ` +
      'Сначала выполните npm run pack:release; сверка без записи — node scripts/sync-site-release.js --check.'
  );
  process.exit(1);
}

let html = fs.readFileSync(siteFile, 'utf8');
const before = html;
const notFound = [];
const changed = [];

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
  const next = html.replace(re, (...args) => build(...args));
  if (next !== html) changed.push(label);
  html = next;
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

// Частичное обновление — незавершённая синхронизация: часть шаблонов в
// разметке не найдена, поэтому не пишем вовсе (сайт остался бы с прежней
// версией или размером от прошлого выпуска).
if (notFound.length > 0) {
  console.error(
    `Сайт не обновлён: в docs/index.html не найдены шаблоны (${notFound.join(', ')}). ` +
      'Проверьте разметку — возможно, изменились классы или формат ссылки.'
  );
  process.exit(1);
}

if (checkOnly) {
  if (changed.length > 0) {
    console.error(
      `docs/index.html расходится с package.json ${version}: не обновлено (${changed.join(', ')}). ` +
        'Выполните npm run site:sync.'
    );
    process.exit(1);
  }
  if (sizeMb === null) {
    // В CI релизного архива нет: release/ не коммитится, сверить нечем.
    // Версия и ссылки уже сошлись — этого достаточно для проверки.
    console.warn(`Размер архива не сверен: release/${zipName} отсутствует.`);
  }
  console.log(`Сайт соответствует package.json: версия ${version}${sizeMb !== null ? `, архив ${sizeMb} МБ` : ''}`);
  process.exit(0);
}

if (changed.length === 0) {
  console.log(`Сайт уже указывает на ${version}${sizeMb !== null ? ` (${sizeMb} МБ)` : ''}`);
  process.exit(0);
}

// Запись только после полного успеха: версия, ссылка и размер обновляются
// вместе, иначе сайт уезжает с версией на несуществующий zip.
fs.writeFileSync(siteFile, html, 'utf8');
console.log(`Сайт обновлён: версия ${version}, архив ${sizeMb} МБ`);
