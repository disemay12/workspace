/**
 * Запуск Electron с обходом неработоспособного Chromium-sandbox.
 *
 * Chromium не может инициализировать sandbox, когда в Windows отключён UAC
 * (HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA = 0).
 * В этом режиме процесс получает полный токен администратора, инициализация
 * sandbox падает, и Electron завершается с STATUS_BREAKPOINT (0x80000003) ещё
 * до загрузки кода приложения — Vite поднимается, а окно не появляется.
 *
 * Обход добавляется только в этом состоянии: при включённом UAC sandbox
 * остаётся рабочим и флаг не передаётся.
 *
 * Использование:
 *   node scripts/run-electron.js dist-electron/main.js [доп. аргументы...]
 */
'use strict';

const { spawnSync, execFileSync } = require('child_process');

const REG_PATH =
  'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System';

/**
 * @returns {boolean} true, если Chromium sandbox на этой машине работоспособен.
 */
function isSandboxUsable() {
  if (process.platform !== 'win32') return true;
  // Явное требование окружения уважаем всегда.
  if (process.env.ELECTRON_DISABLE_SANDBOX) return false;

  try {
    const out = execFileSync('reg', ['query', REG_PATH, '/v', 'EnableLUA'], {
      encoding: 'utf8',
      windowsHide: true,
    });
    const m = /EnableLUA\s+REG_DWORD\s+0x([0-9a-fA-F]+)/.exec(out);
    if (m) return parseInt(m[1], 16) !== 0;
  } catch {
    // Политику прочитать не удалось — считаем sandbox рабочим и не мешаем запуску.
  }
  return true;
}

let electronPath;
try {
  // Под обычным Node пакет electron экспортирует путь к своему бинарнику.
  electronPath = require('electron');
} catch {
  console.error('[run-electron] Пакет electron не найден. Выполните: npm install');
  process.exit(1);
}

// Do not let the parent shell force Electron into its Node-only mode.
// In that mode Chromium switches (for example --no-sandbox) are rejected.
delete process.env.ELECTRON_RUN_AS_NODE;

const appArgs = process.argv.slice(2);
if (appArgs.length === 0) {
  console.error(
    '[run-electron] Укажите точку входа, например: ' +
      'node scripts/run-electron.js dist-electron/main.js'
  );
  process.exit(1);
}

const flags = [];
if (!isSandboxUsable()) {
  flags.push('--no-sandbox');
  console.warn(
    '[run-electron] UAC отключён (EnableLUA=0): Chromium sandbox недоступен, ' +
      'запускаю Electron с --no-sandbox.'
  );
  console.warn(
    '[run-electron] Чтобы вернуть sandbox, включите UAC ' +
      '(EnableLUA=1) и перезагрузитесь.'
  );
}

const result = spawnSync(electronPath, flags.concat(appArgs), {
  stdio: 'inherit',
});

if (result.error) {
  console.error('[run-electron] Не удалось запустить Electron: ' + result.error.message);
  process.exit(1);
}
process.exit(result.status === null ? 1 : result.status);
