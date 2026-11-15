'use strict';

/**
 * Общие правила релизных скриптов: что не должно попадать в сборку и как
 * проверять целостность артефактов.
 *
 * Список запрещённого был скопирован в pack-release.js и build-portable.js и
 * разошёлся: build-portable искал файлы после asar pack, то есть там, где
 * удалять уже нечего. Держим правила в одном месте, чтобы оба пути сборки
 * отбраковывали одинаково.
 */

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const FORBIDDEN = [
  /(^|[\\/])ai-settings\.json$/i,
  /(^|[\\/])gdrive\.json$/i,
  /(^|[\\/])workspace-session\.json$/i,
  // secrets*.json и secrets*.ps1 в любой вложенной папке, но не
  // dist-electron/secrets.js: скомпилированный electron/secrets.ts входит
  // в пакет намеренно.
  /(^|[\\/])secrets[^\\/]*\.(json|ya?ml|ps1|bat|cmd|txt|ini|conf|pem|key|env)$/i,
  /\.local\.ps1$/i,
  // .env, .env.local, .env.production — ключи не должны уехать в архив.
  /(^|[\\/])\.env(\.[^\\/]+)?$/i,
  /\.log$/i,
  /\.ws-tmp$/i,
];

const isForbidden = (relativePath) => FORBIDDEN.some((re) => re.test(relativePath));

/**
 * Рекурсивный поиск запрещённых файлов. Ошибки чтения не глотаются: каталог,
 * который не удалось прочитать, мог скрывать ключ — это повод упасть, а не
 * молча пропустить.
 */
function findForbidden(dir, base = dir) {
  const found = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    const relative = path.relative(base, full);
    if (isForbidden(relative)) {
      found.push(relative);
      continue;
    }
    if (entry.isDirectory()) found.push(...findForbidden(full, base));
  }
  return found;
}

/** sha256 потоком: релизный архив весит ~160 МБ, целиком в память его не читаем. */
function sha256File(file) {
  const hash = crypto.createHash('sha256');
  const fd = fs.openSync(file, 'r');
  try {
    const chunk = Buffer.allocUnsafe(1 << 20);
    let read = fs.readSync(fd, chunk, 0, chunk.length, null);
    while (read > 0) {
      hash.update(chunk.subarray(0, read));
      read = fs.readSync(fd, chunk, 0, chunk.length, null);
    }
  } finally {
    fs.closeSync(fd);
  }
  return hash.digest('hex');
}

/**
 * Список файлов zip для проверки содержимого архива. jszip подключается
 * лениво: build-portable.js нужны только правила и sha256.
 */
async function readZipEntries(file) {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  // Разделители приводим к /: 7-Zip пишет прямые, Compress-Archive — тоже
  // прямые, но список файлов собирается через path.join и может принести
  // обратные.
  return Object.keys(zip.files).map((name) => name.replace(/\\/g, '/'));
}

module.exports = { FORBIDDEN, isForbidden, findForbidden, sha256File, readZipEntries };
