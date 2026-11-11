/**
 * Чистая логика проводника: сортировка, скрытые файлы, плоский список
 * видимых строк, разрешение конфликтов имён, правила drag&drop и журнал
 * отмены файловых операций.
 *
 * Вынесено из `FileTree.tsx`, чтобы это можно было проверять в model-check
 * на синтетическом дереве, без React и файловой системы.
 */
import type { FileItem } from '../types';

export type ExplorerSortBy = 'name' | 'mtime' | 'size' | 'type';

export interface ExplorerSort {
  by: ExplorerSortBy;
  dir: 'asc' | 'desc';
  foldersFirst: boolean;
}

export interface ExplorerViewOptions {
  sort: ExplorerSort;
  showHidden: boolean;
  filter?: string;
}

/** Служебные имена, которые по умолчанию скрыты в дереве. */
export const HIDDEN_NAMES = new Set([
  '.git',
  '.svn',
  '.hg',
  'node_modules',
  '.idea',
  '.vscode',
  '__pycache__',
  '.DS_Store',
  'Thumbs.db',
  'desktop.ini',
  '.cache',
  '.next',
  'dist-electron',
]);

const HIDDEN_SUFFIXES = ['.tmp', '.log', '.bak', '.swp'];

export function isHiddenName(name: string, showHidden: boolean): boolean {
  if (showHidden) return false;
  if (HIDDEN_NAMES.has(name)) return true;
  if (name.startsWith('.') && name !== '..') return true;
  const lower = name.toLowerCase();
  return HIDDEN_SUFFIXES.some((suffix) => lower.endsWith(suffix));
}

function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot > 0 ? name.slice(dot + 1).toLowerCase() : '';
}

/** Компаратор строк: регистронезависимо, с числами «по-человечески». */
export function compareNames(a: string, b: string): number {
  return a.localeCompare(b, 'ru', { numeric: true, sensitivity: 'base' });
}

const timeOf = (item: FileItem): number => {
  if (!item.mtime) return 0;
  const parsed = Date.parse(item.mtime);
  return Number.isFinite(parsed) ? parsed : 0;
};

export function compareItems(a: FileItem, b: FileItem, sort: ExplorerSort): number {
  if (sort.foldersFirst && a.isDirectory !== b.isDirectory) {
    return a.isDirectory ? -1 : 1;
  }
  let result = 0;
  if (sort.by === 'size') result = (a.size || 0) - (b.size || 0);
  else if (sort.by === 'mtime') result = timeOf(a) - timeOf(b);
  else if (sort.by === 'type') result = compareNames(extensionOf(a.name), extensionOf(b.name));
  if (result === 0) result = compareNames(a.name, b.name);
  return sort.dir === 'desc' ? -result : result;
}

export function sortItems(items: FileItem[], sort: ExplorerSort): FileItem[] {
  return [...items].sort((a, b) => compareItems(a, b, sort));
}

/** Совпадение по фильтру: имя, расширение или тип файла. */
export function matchesFilter(item: FileItem, filter: string): boolean {
  const query = filter.trim().toLowerCase();
  if (!query) return true;
  if (item.name.toLowerCase().includes(query)) return true;
  if (item.extension.toLowerCase().includes(query)) return true;
  return Boolean(item.type && item.type.toLowerCase().includes(query));
}

/** Дети папки после фильтра/сортировки/скрытия. */
export function childItems(
  tree: Record<string, FileItem[]>,
  dirPath: string,
  options: ExplorerViewOptions
): FileItem[] {
  const items = tree[dirPath] ?? [];
  const visible = items.filter(
    (item) => !isHiddenName(item.name, options.showHidden) && matchesFilter(item, options.filter ?? '')
  );
  return sortItems(visible, options.sort);
}

/**
 * Плоский список видимых строк в порядке отрисовки: используется и для
 * клавиатурного фокуса, и для виртуализации.
 */
export function flattenVisible(
  tree: Record<string, FileItem[]>,
  expanded: Set<string> | string[],
  rootPath: string,
  options: ExplorerViewOptions,
  depth = 0
): FileItem[] {
  const expandedSet = expanded instanceof Set ? expanded : new Set(expanded);
  const result: FileItem[] = [];
  for (const item of childItems(tree, rootPath, options)) {
    result.push({ ...item, depth } as FileItem & { depth: number });
    if (item.isDirectory && expandedSet.has(item.path)) {
      result.push(...flattenVisible(tree, expandedSet, item.path, options, depth + 1));
    }
  }
  return result;
}

/** Следующее свободное имя при вставке: «a.txt» → «a (2).txt». */
export function resolvePasteName(
  fileName: string,
  existing: Set<string> | string[],
  mode: 'copy' | 'move' = 'copy'
): string {
  const taken = existing instanceof Set ? existing : new Set(existing);
  const lowerTaken = new Set([...taken].map((name) => name.toLowerCase()));
  if (mode === 'move' || !lowerTaken.has(fileName.toLowerCase())) return fileName;

  const dot = fileName.lastIndexOf('.');
  const stem = dot > 0 ? fileName.slice(0, dot) : fileName;
  const ext = dot > 0 ? fileName.slice(dot) : '';
  for (let index = 2; index < 1000; index += 1) {
    const candidate = `${stem} (${index})${ext}`;
    if (!lowerTaken.has(candidate.toLowerCase())) return candidate;
  }
  return `${stem} (${Date.now()})${ext}`;
}

export interface DropSource {
  path: string;
  name: string;
  isDirectory: boolean;
}

export interface DropTarget {
  path: string;
  isDirectory: boolean;
}

export interface DropVerdict {
  ok: boolean;
  reason?: string;
}

/** Правила переноса: нельзя в себя, в своего потомка и файл на файл. */
export function canDrop(source: DropSource, target: DropTarget): DropVerdict {
  if (!target.isDirectory) return { ok: false, reason: 'Цель — файл, а не папка' };
  if (source.path === target.path) return { ok: false, reason: 'Это та же папка' };
  const parent = source.path.replace(/[\\/][^\\/]+$/, '');
  if (parent === target.path) return { ok: false, reason: 'Файл уже здесь' };
  const normalizedSource = source.path.replace(/\//g, '\\').toLowerCase() + '\\';
  const normalizedTarget = target.path.replace(/\//g, '\\').toLowerCase();
  if (source.isDirectory && normalizedTarget.startsWith(normalizedSource)) {
    return { ok: false, reason: 'Нельзя переместить папку внутрь себя' };
  }
  return { ok: true };
}

export type ExplorerOpKind = 'move' | 'rename' | 'create' | 'copy';

export interface ExplorerOp {
  kind: ExplorerOpKind;
  from: string;
  to: string;
  /** Для create/copy: что удалить при откате. */
  createdAt?: string;
}

/** Журнал отмены: только обратимые операции, лимит на размер. */
export class ExplorerHistory {
  private entries: ExplorerOp[] = [];

  constructor(private readonly limit = 50) {}

  push(op: ExplorerOp): void {
    this.entries.push(op);
    if (this.entries.length > this.limit) this.entries.shift();
  }

  get size(): number {
    return this.entries.length;
  }

  peek(): ExplorerOp | undefined {
    return this.entries[this.entries.length - 1];
  }

  /** Снимает последнюю операцию и возвращает её для выполнения отката. */
  pop(): ExplorerOp | undefined {
    return this.entries.pop();
  }

  clear(): void {
    this.entries = [];
  }

  /** Обратный ход: куда переместить и что удалить. */
  static invert(op: ExplorerOp): { moveFrom?: string; moveTo?: string; remove?: string } {
    if (op.kind === 'move' || op.kind === 'rename') return { moveFrom: op.to, moveTo: op.from };
    return { remove: op.createdAt ?? op.to };
  }
}

/**
 * Массовое переименование по шаблону.
 *
 * Подстановки: `{name}` — исходное имя без расширения, `{ext}` — расширение,
 * `{n}` — порядковый номер (по умолчанию с ведущим нулём для серии > 9).
 * Расширение сохраняется, если шаблон его не задаёт.
 */
export function applyRenamePattern(
  names: string[],
  pattern: string,
  startIndex = 1
): { from: string; to: string }[] {
  const template = pattern.trim();
  if (!template) return [];
  const pad = String(names.length + startIndex - 1).length;
  return names.map((name, index) => {
    const dot = name.lastIndexOf('.');
    const stem = dot > 0 ? name.slice(0, dot) : name;
    const ext = dot > 0 ? name.slice(dot + 1) : '';
    const number = String(index + startIndex).padStart(pad, '0');
    let next = template
      .replace(/\{name\}/gi, stem)
      .replace(/\{ext\}/gi, ext)
      .replace(/\{n(?::(\d+))?\}/gi, (_match, width) =>
        width ? String(index + startIndex).padStart(Number(width), '0') : number
      );
    if (ext && !next.toLowerCase().endsWith('.' + ext.toLowerCase())) next = `${next}.${ext}`;
    return { from: name, to: next };
  });
}

export function formatBytes(bytes?: number): string {
  const value = Number(bytes) || 0;
  if (value <= 0) return '';
  if (value < 1024) return `${value} Б`;
  if (value < 1024 * 1024) return `${Math.round(value / 1024)} КБ`;
  if (value < 1024 * 1024 * 1024) return `${(value / (1024 * 1024)).toFixed(1)} МБ`;
  return `${(value / (1024 * 1024 * 1024)).toFixed(2)} ГБ`;
}

export function formatDate(mtime?: string | null): string {
  if (!mtime) return '';
  const parsed = Date.parse(mtime);
  if (!Number.isFinite(parsed)) return '';
  const date = new Date(parsed);
  const now = new Date();
  const sameDay =
    date.getDate() === now.getDate() &&
    date.getMonth() === now.getMonth() &&
    date.getFullYear() === now.getFullYear();
  const time = date.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
  if (sameDay) return time;
  return date.toLocaleDateString('ru-RU', { day: '2-digit', month: '2-digit', year: '2-digit' });
}
