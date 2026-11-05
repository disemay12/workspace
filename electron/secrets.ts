/**
 * Импорт локальных секретов из secrets.local.ps1.
 *
 * Файл содержит ключи и НЕ попадает в репозиторий (см. .gitignore).
 * Читается один раз в главном процессе; значения отдаются рендереру
 * только по явному запросу пользователя (кнопка «Импортировать из
 * secrets.local.ps1» в настройках), чтобы ключ не оказывался в UI
 * без ведома владельца.
 */
import { app } from 'electron';
import * as fs from 'fs';
import * as path from 'path';

export interface LocalSecrets {
  found: boolean;
  filePath: string;
  baseUrl?: string;
  apiKey?: string;
  textModel?: string;
  imageModel?: string;
  /** Есть ли в файле ключи, не относящиеся к AI (Notion, Telegram). */
  hasOtherSecrets: boolean;
  error?: string;
}

/** Разбирает `$env:NAME = "value"` из PowerShell-скрипта. */
function parseEnvAssignments(content: string): Record<string, string> {
  const result: Record<string, string> = {};
  // Поддерживаем одинарные и двойные кавычки, а также значения без кавычек
  const re = /^\s*\$env:([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\r\n#]+))/gm;
  let match: RegExpExecArray | null;
  while ((match = re.exec(content)) !== null) {
    const name = match[1];
    const value = (match[2] ?? match[3] ?? match[4] ?? '').trim();
    if (value) result[name] = value;
  }
  return result;
}

/** Ищет secrets.local.ps1: профиль пользователя, затем проект/папка exe. */
function locateSecretsFile(): string | null {
  const candidates: string[] = [];

  // 1. Профиль пользователя: единственное место, которое переживает обновление.
  candidates.push(path.join(app.getPath('userData'), 'secrets.local.ps1'));

  // 2. Рядом с исполняемым файлом / текущим рабочим каталогом
  candidates.push(path.join(process.cwd(), 'secrets.local.ps1'));

  // 2. В каталоге приложения (dev: корень проекта; prod: рядом с ресурсами)
  candidates.push(path.join(__dirname, '..', 'secrets.local.ps1'));
  candidates.push(path.join(__dirname, '..', '..', 'secrets.local.ps1'));

  // 3. Подъём вверх по дереву от cwd (до 5 уровней)
  let dir = process.cwd();
  for (let i = 0; i < 5; i++) {
    candidates.push(path.join(dir, 'secrets.local.ps1'));
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }

  for (const candidate of candidates) {
    try {
      if (fs.existsSync(candidate) && fs.statSync(candidate).isFile()) {
        return candidate;
      }
    } catch {
      // каталог недоступен — пропускаем
    }
  }
  return null;
}

/**
 * Создаёт шаблон в профиле пользователя, если файла нет нигде: пользователь
 * не должен вручную придумывать путь к файлу секретов.
 */
function ensureTemplate(): string | null {
  const target = path.join(app.getPath('userData'), 'secrets.local.ps1');
  const template = [
    '# WorkSpace: локальные секреты. Файл создан автоматически при первом запуске.',
    '# Впишите значения и перезапустите WorkSpace. Файл не публикуйте.',
    '# $env:API_BASE_URL = "https://api.deepseek.com/v1"',
    '# $env:API_KEY = ""',
    '# $env:TEXT_MODEL = "deepseek-chat"',
    '# $env:IMAGE_MODEL = ""',
    '',
  ].join('\r\n');
  try {
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, template, 'utf8');
    return target;
  } catch {
    return null;
  }
}

/** Существующий файл секретов или созданный шаблон; null — записать некуда. */
function ensureSecretsFile(): string | null {
  return locateSecretsFile() ?? ensureTemplate();
}

export const secretsImporter = {
  /**
   * Гарантирует, что файл секретов есть: возвращает существующий или создаёт
   * шаблон в профиле пользователя.
   */
  ensure: ensureSecretsFile,
  /**
   * Читает локальный файл секретов.
   * API-ключ возвращается только при includeKey = true.
   */
  read(includeKey = false): LocalSecrets {
    const filePath = ensureSecretsFile();

    if (!filePath) {
      return {
        found: false,
        filePath: '',
        hasOtherSecrets: false,
        error: 'Файл secrets.local.ps1 не найден рядом с проектом',
      };
    }

    try {
      const content = fs.readFileSync(filePath, 'utf-8');
      const env = parseEnvAssignments(content);

      const baseUrl = env.API_BASE_URL;
      const apiKey = env.API_KEY;
      const textModel = env.TEXT_MODEL;
      const imageModel = env.IMAGE_MODEL;

      const hasOtherSecrets = Boolean(env.INTEGRATION_TOKEN_NOTION || env.TELEGRAM_BOT_TOKEN);

      return {
        found: true,
        filePath,
        baseUrl,
        // Ключ отдаём только по явному запросу
        apiKey: includeKey ? apiKey : apiKey ? '***' : undefined,
        textModel,
        imageModel,
        hasOtherSecrets,
      };
    } catch (err) {
      return {
        found: true,
        filePath,
        hasOtherSecrets: false,
        error: `Не удалось прочитать файл: ${(err as Error).message}`,
      };
    }
  },
};
