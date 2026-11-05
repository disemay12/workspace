/**
 * Постоянные настройки AI.
 *
 * localStorage в собранном приложении привязан к адресу файла страницы и
 * обнуляется после перезапуска. Ключ хранится в обычном файле.
 *
 * Файл живёт только в профиле пользователя (%APPDATA%\WorkSpace): папка
 * релиза перезаписывается при обновлении, и настройки в ней терялись бы.
 * Если старая версия положила ai-settings.json рядом с exe — переносим его
 * один раз при первом запуске. Открытый ключ на диск не пишется: он
 * шифруется средствами Windows.
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
  modelsCache?: Record<string, string[]>;
  /** Зашифрованные ключи по провайдерам. Открытый текст сюда не пишется. */
  keys?: Record<string, string>;
}

/** Базовые значения: файл создаётся при первом запуске, чтобы UI был не пустым. */
const DEFAULT_SETTINGS: StoredSettings = {
  temperature: 0.7,
  maxTokens: 4096,
};

/** Единственное место записи: профиль пользователя. */
function settingsPath(): string {
  return path.join(app.getPath('userData'), 'ai-settings.json');
}

/** Файлы старой версии рядом с exe — только источник для разовой миграции. */
function legacyPaths(): string[] {
  const list = [path.join(path.dirname(process.execPath), 'ai-settings.json')];
  if (app.isPackaged) {
    // execPath — это WorkSpace\app\WorkSpace.exe, старый файл лежал на уровень выше.
    list.push(path.resolve(path.dirname(process.execPath), '..', 'ai-settings.json'));
  }
  return list;
}

function readFrom(target: string): StoredSettings | null {
  try {
    // Notepad сохраняет UTF-8 с BOM — JSON.parse на таком файле падает.
    const raw = fs.readFileSync(target, 'utf8').replace(/^\uFEFF/, '');
    const parsed = JSON.parse(raw) as StoredSettings;
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch {
    return null;
  }
}

function writeTo(target: string, data: StoredSettings): void {
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const tmp = target + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(data), 'utf8');
  fs.renameSync(tmp, target);
}

function read(): StoredSettings {
  const target = settingsPath();
  const own = readFrom(target);
  if (own) return own;

  // Старая версия держала файл рядом с exe — переносим его в профиль.
  for (const legacy of legacyPaths()) {
    const data = readFrom(legacy);
    if (!data) continue;
    try {
      writeTo(target, data);
    } catch {
      // Перенести не удалось (профиль только для чтения) — работаем с копией.
    }
    return data;
  }

  // Ничего нет: создаём файл с базовыми значениями.
  try {
    writeTo(target, DEFAULT_SETTINGS);
  } catch {
    // Не удалось — вернём значения в память, сохранение попробует ещё раз.
  }
  return { ...DEFAULT_SETTINGS };
}

/** Пишет только в профиль пользователя: папка релиза при обновлении меняется. */
function write(data: StoredSettings): void {
  try {
    writeTo(settingsPath(), data);
  } catch (err) {
    throw err instanceof Error ? err : new Error('Не удалось сохранить настройки');
  }
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
    /** Ключи, которые не удалось расшифровать (профиль/компьютер сменились). */
    const keyErrors: string[] = [];
    for (const [provider, stored] of Object.entries(data.keys ?? {})) {
      const value = open(stored);
      if (value) keys[provider] = value;
      else keyErrors.push(provider);
    }
    const { keys: _omit, ...settings } = data;
    return { settings, keys, keyErrors };
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
