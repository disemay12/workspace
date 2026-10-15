import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

/** Объединение классов Tailwind с корректным разрешением конфликтов. */
export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(inputs));
}

/** Форматирование размера файла в человекочитаемый вид. */
export function formatBytes(bytes: number): string {
  if (!bytes || bytes < 0) return '0 Б';
  const units = ['Б', 'КБ', 'МБ', 'ГБ', 'ТБ'];
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / Math.pow(1024, i);
  return `${value >= 100 || i === 0 ? Math.round(value) : value.toFixed(1)} ${units[i]}`;
}

/** Относительное время («5 мин назад») для списков файлов. */
export function formatRelativeTime(iso: string | null): string {
  if (!iso) return '';
  const then = new Date(iso).getTime();
  if (Number.isNaN(then)) return '';
  const diff = Date.now() - then;
  const min = Math.floor(diff / 60000);
  if (min < 1) return 'только что';
  if (min < 60) return `${min} мин назад`;
  const hours = Math.floor(min / 60);
  if (hours < 24) return `${hours} ч назад`;
  const days = Math.floor(hours / 24);
  if (days < 30) return `${days} дн назад`;
  return new Date(iso).toLocaleDateString('ru-RU');
}

/** Разбиение пути на компоненты (поддержка Windows и POSIX). */
export function splitPath(fullPath: string): string[] {
  return fullPath.split(/[\\/]+/).filter(Boolean);
}

/** Имя файла без каталога. */
export function baseName(fullPath: string): string {
  const parts = splitPath(fullPath);
  return parts.length ? parts[parts.length - 1] : fullPath;
}

/** Каталог файла. */
export function dirName(fullPath: string): string {
  const idx = Math.max(fullPath.lastIndexOf('\\'), fullPath.lastIndexOf('/'));
  return idx > 0 ? fullPath.slice(0, idx) : fullPath;
}

/** Расширение в нижнем регистре без точки. */
export function extName(fullPath: string): string {
  const name = baseName(fullPath);
  const idx = name.lastIndexOf('.');
  return idx > 0 ? name.slice(idx + 1).toLowerCase() : '';
}

/** Соединение пути с учётом разделителя родителя. */
export function joinPath(dir: string, name: string): string {
  const sep = dir.includes('\\') ? '\\' : '/';
  return dir.endsWith(sep) ? `${dir}${name}` : `${dir}${sep}${name}`;
}

/** Короткий уникальный идентификатор. */
export function uid(prefix = 'id'): string {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
}

/**
 * Извлекает блоки кода ```lang ... ``` из Markdown-ответа модели.
 * Используется для кнопки «Вставить в редактор».
 */
export function extractCodeBlocks(markdown: string): { lang: string; code: string }[] {
  const blocks: { lang: string; code: string }[] = [];
  const re = /```([\w+-]*)\r?\n([\s\S]*?)```/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(markdown)) !== null) {
    blocks.push({ lang: match[1] || '', code: match[2].replace(/\s+$/, '') });
  }
  return blocks;
}

/** Убирает Markdown-обёртку ``` из ответа модели, если она есть. */
export function stripCodeFence(text: string): string {
  const trimmed = text.trim();
  const fence = trimmed.match(/^```[\w+-]*\r?\n([\s\S]*?)```$/);
  return fence ? fence[1].replace(/\s+$/, '') : trimmed;
}
