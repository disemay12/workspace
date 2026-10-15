/**
 * Постоянные настройки AI.
 *
 * localStorage в собранном приложении привязан к адресу файла страницы и
 * обнуляется после перезапуска. Ключ хранится в обычном файле.
 *
 * Сначала — рядом с приложением (папка WorkSpace): этот процесс может её
 * писать, и файл переносится вместе с папкой на другой компьютер.
 * Если папка только для чтения — в профиле пользователя.
 * Открытый ключ на диск не пишется: он шифруется средствами Windows.
 */
import { app, ipcMain, safeStorage } from 'electron';
import * as fs from 'fs';
import * as path from 'path';

interface StoredSettings {
  provider?: string;
  baseUrl?: string;
  model?: string;
  temperature?: number;
  maxTokens?: number;
  systemPrompt?: string;
  activePresetId?: string;
  /** Зашифрованные ключи по провайдерам. Открытый текст сюда не пишется. */
  keys?: Record<string, string>;
}

/** Места хранения от предпочтительного к запасному. */
function candidatePaths(): string[] {
  const paths: string[] = [];
  if (app.isPackaged) {
    // execPath — это WorkSpace\app\WorkSpace.exe, файл кладём на уровень выше.
    paths.push(path.resolve(path.dirname(process.execPath), '..', 'ai-settings.json'));
  }
  paths.push(path.join(app.getPath('userData'), 'ai-settings.json'));
  return paths;
}

function readFrom(target: string): StoredSettings | null {
  try {
    const parsed = JSON.parse(fs.readFileSync(target, 'utf8')) as StoredSettings;
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch {
    return null;
  }
}

function read(): StoredSettings {
  for (const target of candidatePaths()) {
    const data = readFrom(target);
    if (data) return data;
  }
  return {};
}

/** Пишет в первое доступное место. Бросает ошибку, только если не вышло нигде. */
function write(data: StoredSettings): void {
  const payload = JSON.stringify(data);
  let lastError: unknown;
  for (const target of candidatePaths()) {
    try {
      fs.mkdirSync(path.dirname(target), { recursive: true });
      const tmp = target + '.tmp';
      fs.writeFileSync(tmp, payload, 'utf8');
      fs.renameSync(tmp, target);
      return;
    } catch (err) {
      lastError = err;
    }
  }
  throw lastError instanceof Error ? lastError : new Error('Не удалось сохранить настройки');
}

/** Шифрует ключ средствами ОС. Если шифрование недоступно — кодирует в base64. */
function seal(plain: string): string {
  if (safeStorage.isEncryptionAvailable()) {
    return 'safe:' + safeStorage.encryptString(plain).toString('base64');
  }
  return 'b64:' + Buffer.from(plain, 'utf8').toString('base64');
}

function open(stored: string): string {
  try {
    if (stored.startsWith('safe:')) {
      return safeStorage.decryptString(Buffer.from(stored.slice(5), 'base64'));
    }
    if (stored.startsWith('b64:')) {
      return Buffer.from(stored.slice(4), 'base64').toString('utf8');
    }
  } catch {
    return '';
  }
  return '';
}

export function registerAiSettings(): void {
  ipcMain.handle('aiSettings:load', () => {
    const data = read();
    const keys: Record<string, string> = {};
    for (const [provider, stored] of Object.entries(data.keys ?? {})) {
      const value = open(stored);
      if (value) keys[provider] = value;
    }
    const { keys: _omit, ...settings } = data;
    return { settings, keys };
  });

  ipcMain.handle(
    'aiSettings:save',
    (_event, patch: Partial<StoredSettings> & { apiKey?: string }) => {
      const data = read();
      const { apiKey, keys: _ignore, ...rest } = patch;
      Object.assign(data, rest);
      if (typeof apiKey === 'string' && patch.provider) {
        data.keys = data.keys ?? {};
        if (apiKey.trim()) data.keys[patch.provider] = seal(apiKey.trim());
        else delete data.keys[patch.provider];
      }
      write(data);
      return { success: true };
    }
  );
}
