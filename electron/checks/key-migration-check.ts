/**
 * Перенос ключей API из localStorage в постоянное хранилище профиля.
 *
 * Проверяется критическая ветка src/state/settings.ts (migrateLegacyKeys):
 * до исправления ключ шифрования прежней версии стирался при первом чтении,
 * а в профиль переносился только ключ АКТИВНОГО провайдера — остальные записи
 * после перезапуска расшифровать было нечем, и ключи терялись молча.
 *
 * Сценарий: профиль с пустым ai-settings.json, в localStorage — две записи
 * `enc:` и legacy-ключ. После перезапуска окно обязано перенести ОБА ключа
 * в файл профиля, стереть legacy-ключ и не потерять активного провайдера.
 */
const { app, BrowserWindow, ipcMain, safeStorage } = require('electron');
const path = require('path');
const fs = require('fs');
const { createCheckSandbox } = require('./sandbox');
const { registerCommonCheckIpc } = require('./common-ipc');
const { registerAiSettings } = require('./aiSettings');

const ROOT = (() => {
  let dir = __dirname;
  for (let i = 0; i < 4; i++) {
    if (fs.existsSync(path.join(dir, 'package.json')) && fs.existsSync(path.join(dir, 'dist'))) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return path.join(__dirname, '..');
})();

const SANDBOX = createCheckSandbox('keymig', ROOT);
const PROFILE = path.join(SANDBOX, 'profile');
const SETTINGS_FILE = path.join(PROFILE, 'ai-settings.json');
const RESULT_FILE = path.join(ROOT, 'key-migration-check.json');
// Профиль подменяется до app.whenReady: aiSettings читает app.getPath('userData').
app.setPath('userData', PROFILE);

const OPENAI_KEY = 'sk-openai-legacy-7f3a1c9d';
const DEEPSEEK_KEY = 'sk-deepseek-legacy-2b6e4a80';

const out: any = { steps: [], errors: [] };
const step = (name: string, value: any) => out.steps.push({ name, ...value });

function writeResult() {
  try {
    fs.writeFileSync(RESULT_FILE, JSON.stringify(out, null, 2), 'utf-8');
  } catch {
    // Каталог проекта недоступен — процесс всё равно завершится с кодом 1.
  }
}

function fail(reason: unknown) {
  out.errors.push(String(reason));
  writeResult();
  app.exit(1);
}

process.on('unhandledRejection', fail);
process.on('uncaughtException', fail);
setTimeout(() => fail('watchdog: сценарий превысил лимит времени'), 120_000);

app.disableHardwareAcceleration();

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

function preloadPath(): string {
  const flat = path.join(__dirname, 'preload.js');
  if (fs.existsSync(flat)) return flat;
  return path.join(__dirname, '..', 'preload.js');
}

/** Расшифровка значения из профиля: тот же формат, что пишет electron/aiSettings.ts. */
function openStored(value: string): string {
  if (typeof value !== 'string') return '';
  try {
    if (value.startsWith('safe:')) return safeStorage.decryptString(Buffer.from(value.slice(5), 'base64'));
    if (value.startsWith('b64:')) return Buffer.from(value.slice(4), 'base64').toString('utf8');
  } catch {
    return '';
  }
  return '';
}

function readSettings(): any {
  try {
    return JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8'));
  } catch {
    return null;
  }
}

async function main() {
  fs.mkdirSync(PROFILE, { recursive: true });
  // Файл профиля без ключей: провайдер уже выбран пользователем.
  fs.writeFileSync(
    SETTINGS_FILE,
    JSON.stringify({ provider: 'openai', baseUrl: 'https://api.openai.com/v1', model: 'gpt-4o' }),
    'utf-8'
  );

  await app.whenReady();
  // Боевые обработчики aiSettings — до заглушек common-ipc: заглушка не должна
  // перехватить канал (второй handler для канала допускается и логируется).
  registerAiSettings();
  // Каналы, которых нет в common-ipc: приложение спрашивает их при старте, а
  // без обработчика в отчёт попадает шум «No handler registered».
  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('fs:listDirectory', () => ({ success: true, items: [] }));
  ipcMain.handle('system:getInitialExternalPath', () => null);
  // Кнопки заголовка окна: приложение спрашивает их сразу после монтирования.
  ipcMain.handle('window:isMaximized', () => false);
  ipcMain.handle('window:maximize', () => false);
  ipcMain.handle('window:minimize', () => undefined);
  ipcMain.handle('window:getAlwaysOnTop', () => ({ alwaysOnTop: false }));
  registerCommonCheckIpc({ sandbox: SANDBOX });

  const win = new BrowserWindow({
    show: false,
    width: 1200,
    height: 800,
    webPreferences: {
      preload: preloadPath(),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
    },
  });

  const js = (expr: string) => win.webContents.executeJavaScript(`(async () => { ${expr} })()`);
  const loaded = () =>
    new Promise<void>((resolve) => {
      win.webContents.once('did-finish-load', () => resolve());
    });

  let nextLoad = loaded();
  await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
  await nextLoad;

  // Состояние «старой версии»: две зашифрованные записи и ключ шифрования рядом.
  const seeded = await js(`
    const encode = (bytes) => btoa(String.fromCharCode(...bytes));
    const rawKey = crypto.getRandomValues(new Uint8Array(32));
    const key = await crypto.subtle.importKey('raw', rawKey, 'AES-GCM', false, ['encrypt', 'decrypt']);
    const seal = async (text) => {
      const iv = crypto.getRandomValues(new Uint8Array(12));
      const data = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, new TextEncoder().encode(text)));
      return 'enc:' + encode(iv) + ':' + encode(data);
    };
    localStorage.setItem('workspace_api_keys', JSON.stringify({ openai: await seal(${JSON.stringify(OPENAI_KEY)}), deepseek: await seal(${JSON.stringify(DEEPSEEK_KEY)}) }));
    localStorage.setItem('workspace_key_encryption_v1', encode(rawKey));
    localStorage.removeItem('workspace_api_keys_migrated_v2');
    localStorage.setItem('workspace_ai_settings', JSON.stringify({ provider: 'openai' }));
    return localStorage.getItem('workspace_key_encryption_v1') !== null;
  `);

  // Перезапуск окна запускает startup-эффект настроек, в котором живёт перенос.
  nextLoad = loaded();
  win.webContents.reload();
  await nextLoad;

  const deadline = Date.now() + 25_000;
  let settings = readSettings();
  while (Date.now() < deadline) {
    settings = readSettings();
    if (settings?.keys?.openai && settings?.keys?.deepseek) break;
    await sleep(250);
  }

  const keys = settings?.keys ?? {};
  const openaiPlain = openStored(keys.openai);
  const deepseekPlain = openStored(keys.deepseek);
  const storage = await js(`
    return {
      legacyKey: localStorage.getItem('workspace_key_encryption_v1'),
      store: localStorage.getItem('workspace_api_keys'),
      migrated: localStorage.getItem('workspace_api_keys_migrated_v2'),
    };
  `);

  step('seededLegacyState', {
    seeded: seeded === true,
    plainKeys: 2,
  });
  step('profileKeys', {
    fileHasBothKeys: Boolean(keys.openai) && Boolean(keys.deepseek),
    plaintextMatches: openaiPlain === OPENAI_KEY && deepseekPlain === DEEPSEEK_KEY,
    // Диагностика, а не утверждение: без DPAPI ключ ляжет префиксом b64:,
    // и перенос всё равно обязан сохранить оба ключа.
    openaiKeyLength: openaiPlain.length,
    deepseekKeyLength: deepseekPlain.length,
    encryptionPrefix: String(keys.openai || '').slice(0, 5),
  });
  step('cleanup', {
    legacyKeyRemoved: storage.legacyKey === null,
    localStoreCleared: storage.store === null,
    migrationFlagSet: storage.migrated === '1',
  });
  step('activeProviderPreserved', {
    providerRestored: settings?.provider === 'openai',
    modelKept: settings?.model === 'gpt-4o',
    baseUrlKept: settings?.baseUrl === 'https://api.openai.com/v1',
  });

  const failed = out.steps.some((entry: any) =>
    Object.entries(entry).some(([key, value]) => key !== 'name' && value === false)
  );
  if (failed) {
    writeResult();
    app.exit(1);
    return;
  }
  writeResult();
  if (!win.isDestroyed()) win.destroy();
  app.exit(0);
}

void main().catch(fail);
