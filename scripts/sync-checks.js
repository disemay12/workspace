/**
 * Переносит скомпилированные проверочные сценарии в dist-electron/.
 *
 * Зачем это нужно отдельным шагом: tsc компилирует electron/checks/*.ts в
 * dist-electron/electron/checks/*.js, а сценарии запускаются из dist-electron/
 * (так исторически сложилось, и run-all-checks.js ищет их именно там).
 *
 * Сам run-all-checks.js тоже лежит в этом каталоге, поэтому он не может
 * скопировать себя сам до запуска: в dist-electron/run-all-checks.js оставалась
 * старая версия — вместе со старым списком обработчиков IPC. В итоге правки в
 * проверках не применялись. Этот скрипт выполняется ДО запуска раннера и
 * разрывает замкнутый круг.
 */
'use strict';

const fs = require('fs');
const path = require('path');

// Скрипт лежит в scripts/, поэтому корень проекта — на уровень выше.
const ROOT = path.resolve(__dirname, '..');
const NESTED = path.join(ROOT, 'dist-electron', 'electron', 'checks');
const FLAT = path.join(ROOT, 'dist-electron');

if (!fs.existsSync(NESTED)) {
  console.error('[sync-checks] Каталог не найден: ' + NESTED);
  console.error('[sync-checks] Сначала выполните: npm run checks:build');
  process.exit(1);
}

/** Копирует один файл, сообщая о неудаче вместо молчаливого пропуска. */
function copy(from, to) {
  try {
    fs.copyFileSync(from, to);
    return true;
  } catch (err) {
    console.error('[sync-checks] Не удалось скопировать ' + path.basename(from) + ': ' + err.message);
    return false;
  }
}

let copied = 0;
const failed = [];

for (const entry of fs.readdirSync(NESTED)) {
  if (!entry.endsWith('.js')) continue;
  if (copy(path.join(NESTED, entry), path.join(FLAT, entry))) copied++;
  else failed.push(entry);
}

if (failed.length > 0) {
  console.error('[sync-checks] Не скопированы: ' + failed.join(', '));
  console.error('[sync-checks] Закройте запущенные экземпляры приложения и повторите.');
  process.exit(1);
}

console.log('[sync-checks] Скопировано файлов: ' + copied);
