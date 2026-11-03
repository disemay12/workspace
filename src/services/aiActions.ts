import * as XLSX from 'xlsx';
import { htmlToDocxBase64 } from '../lib/docx';
import { markdownToDocumentHtml } from '../lib/aiDocument';

/* ============================================================
   Типы действий AI по управлению программой WorkSpace
   ============================================================ */

export type CellValue = string | number | boolean | null;

export type AppActionKind =
  | 'create_table'
  | 'import_table'
  | 'edit_table'
  | 'create_file'
  | 'create_document'
  | 'open_file'
  | 'open_folder'
  | 'create_folder'
  | 'rename_path'
  | 'delete_path'
  | 'insert_into_editor'
  | 'replace_document'
  | 'open_settings'
  | 'toggle_panel'
  | 'select_tab'
  | 'close_tab'
  | 'save'
  | 'notify'
  | 'run_command';

export interface AICreateTablePayload {
  action: 'create_table';
  name?: string;
  path?: string;
  sheetName?: string;
  columns?: string[];
  rows?: CellValue[][];
  format?: 'xlsx' | 'csv';
}

/** Импорт данных из файла (Markdown/CSV) в новую или существующую таблицу. */
export interface AIImportTablePayload {
  action: 'import_table';
  /** Исходный файл: Markdown, CSV/TSV или Excel. */
  source: string;
  /** Целевой xlsx (если не задан — рядом с исходником). */
  target?: string;
  sheetName?: string;
}

/** Изменение существующей таблицы: замена или добавление строк. */
export interface AIEditTablePayload {
  action: 'edit_table';
  path?: string;
  name?: string;
  sheetName?: string;
  mode: 'replace' | 'append';
  columns?: string[];
  rows?: CellValue[][];
}

export interface AICreateFilePayload {  action: 'create_file';
  name?: string;
  path?: string;
  content?: string;
  encoding?: string;
}

export interface AICreateDocumentPayload {
  action: 'create_document';
  name?: string;
  path?: string;
  title?: string;
  content?: string;
}

export interface AIOpenFilePayload {
  action: 'open_file';
  path?: string;
  name?: string;
}

export interface AIOpenFolderPayload {
  action: 'open_folder';
  path?: string;
}

export interface AICreateFolderPayload {
  action: 'create_folder';
  name?: string;
  path?: string;
}

export interface AIRenamePathPayload {
  action: 'rename_path';
  path: string;
  newName?: string;
  newPath?: string;
}

export interface AIDeletePathPayload {
  action: 'delete_path';
  path: string;
}

export interface AIInsertEditorPayload {
  action: 'insert_into_editor';
  text: string;
}

export interface AIReplaceDocPayload {
  action: 'replace_document';
  text: string;
}

export type SettingsTab = 'provider' | 'model' | 'prompt' | 'keys' | 'integrations' | 'system';

export interface AIOpenSettingsPayload {
  action: 'open_settings';
  tab?: SettingsTab;
}

export type PanelName = 'sidebar' | 'ai' | 'terminal' | 'search';

export interface AITogglePanelPayload {
  action: 'toggle_panel';
  panel: PanelName;
  state?: 'open' | 'close' | 'toggle';
}

export interface AISelectTabPayload {
  action: 'select_tab';
  path?: string;
  name?: string;
  index?: number;
}

export interface AICloseTabPayload {
  action: 'close_tab';
  path?: string;
  name?: string;
}

export interface AISavePayload {
  action: 'save';
  all?: boolean;
}

export interface AINotifyPayload {
  action: 'notify';
  message: string;
}

export interface AIRunCommandPayload {
  action: 'run_command';
  command: string;
  cwd?: string;
}

export type AIActionPayload =
  | AICreateTablePayload
  | AIImportTablePayload
  | AIEditTablePayload
  | AICreateFilePayload
  | AICreateDocumentPayload
  | AIOpenFilePayload
  | AIOpenFolderPayload
  | AICreateFolderPayload
  | AIRenamePathPayload
  | AIDeletePathPayload
  | AIInsertEditorPayload
  | AIReplaceDocPayload
  | AIOpenSettingsPayload
  | AITogglePanelPayload
  | AISelectTabPayload
  | AICloseTabPayload
  | AISavePayload
  | AINotifyPayload
  | AIRunCommandPayload;

export interface ParsedMarkdownTable {
  columns: string[];
  rows: string[][];
  rawMarkdown: string;
}

/** Результат выполнения действия в приложении. */
export interface AppActionResult {
  ok: boolean;
  message: string;
  /** Путь созданного/открытого файла — для кнопки «Открыть». */
  path?: string;
}

/** Действия, которые выполняются только после явного подтверждения в чате. */
export const CONFIRMATION_ACTIONS = new Set<AppActionKind>(['delete_path', 'run_command']);

export function requiresConfirmation(action: AIActionPayload): boolean {
  return CONFIRMATION_ACTIONS.has(action.action);
}

/* ============================================================
   Разбор действий из ответа модели
   ============================================================ */

const ACTION_ALIASES: Record<string, AppActionKind> = {
  create_table: 'create_table',
  create_excel: 'create_table',
  createtable: 'create_table',
  excel: 'create_table',
  table: 'create_table',
  spreadsheet: 'create_table',
  import_table: 'import_table',
  import_excel: 'import_table',
  file_to_table: 'import_table',
  convert_to_table: 'import_table',
  md_to_excel: 'import_table',
  edit_table: 'edit_table',
  update_table: 'edit_table',
  append_table: 'edit_table',
  write_table: 'edit_table',
  fill_table: 'edit_table',
  write_excel: 'edit_table',
  write_xlsx: 'edit_table',
  update_excel: 'edit_table',
  fill_excel: 'edit_table',
  create_sheet: 'edit_table',
  create_file: 'create_file',
  createfile: 'create_file',
  new_file: 'create_file',
  write_file: 'create_file',
  file: 'create_file',
  create_document: 'create_document',
  create_doc: 'create_document',
  create_word: 'create_document',
  createword: 'create_document',
  word: 'create_document',
  document: 'create_document',
  open_file: 'open_file',
  openfile: 'open_file',
  open: 'open_file',
  open_folder: 'open_folder',
  openfolder: 'open_folder',
  open_workspace: 'open_folder',
  create_folder: 'create_folder',
  createfolder: 'create_folder',
  new_folder: 'create_folder',
  mkdir: 'create_folder',
  rename_path: 'rename_path',
  rename_file: 'rename_path',
  rename: 'rename_path',
  delete_path: 'delete_path',
  delete_file: 'delete_path',
  remove: 'delete_path',
  delete: 'delete_path',
  insert_into_editor: 'insert_into_editor',
  insert_text: 'insert_into_editor',
  insert: 'insert_into_editor',
  replace_document: 'replace_document',
  replace_text: 'replace_document',
  replace: 'replace_document',
  open_settings: 'open_settings',
  settings: 'open_settings',
  toggle_panel: 'toggle_panel',
  panel: 'toggle_panel',
  select_tab: 'select_tab',
  switch_tab: 'select_tab',
  tab: 'select_tab',
  close_tab: 'close_tab',
  close: 'close_tab',
  save: 'save',
  save_file: 'save',
  save_all: 'save',
  notify: 'notify',
  message: 'notify',
  run_command: 'run_command',
  run_terminal: 'run_command',
  terminal: 'run_command',
  command: 'run_command',
};

const ACTION_FENCE_LANGS = new Set([
  '',
  'json',
  'workspace',
  'workspaceaction',
  'workspaceactions',
  'action',
  'actions',
  'appaction',
  'appactions',
  'workspaceappaction',
]);

function normalizeKind(raw: unknown): AppActionKind | null {
  if (typeof raw !== 'string') return null;
  const key = raw.trim().toLowerCase().replace(/[\s-]+/g, '_');
  return ACTION_ALIASES[key] ?? ACTION_ALIASES[key.replace(/_/g, '')] ?? null;
}

function asString(value: unknown): string | undefined {
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  return undefined;
}

function asStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value.map((item) => asString(item) ?? '').filter((item) => item.length > 0);
}

function parseJsonLoose(raw: string): unknown {
  const attempts = [raw, raw.replace(/,\s*([\]}])/g, '$1')];
  for (const attempt of attempts) {
    try {
      return JSON.parse(attempt);
    } catch {
      // пробуем следующий вариант
    }
  }
  return null;
}

/**
 * Разбирает действие, even если ответ модели оборван по лимиту токенов:
 * достаёт имя действия, поля и все полностью сформированные строки массива.
 */
function parseActionTolerant(raw: string): unknown {
  const direct = parseJsonLoose(raw.trim());
  if (direct) return direct;
  const action =
    /"(?:action|type)"\s*:\s*"([^"]+)"/.exec(raw)?.[1] ??
    /"action"\s*:\s*"([^"]+)"/.exec(raw)?.[1];
  if (!action) return null;
  const body: Record<string, unknown> = { action };
  for (const key of ['path', 'file', 'source', 'target', 'sheetName', 'sheet', 'name', 'title', 'mode']) {
    const match = new RegExp(`"${key}"\\s*:\\s*"([^"]*)"`).exec(raw);
    if (match) body[key] = match[1];
  }
  const rows: unknown[] = [];
  for (const match of raw.matchAll(/\[[^\[\]]*\]/g)) {
    const value = parseJsonLoose(match[0]);
    if (Array.isArray(value)) rows.push(value);
  }
  if (rows.length > 0) body.data = rows;
  return Object.keys(body).length > 1 ? body : null;
}

function normalizeRows(raw: unknown): CellValue[][] {
  if (!Array.isArray(raw)) return [];
  const rows: CellValue[][] = [];
  for (const row of raw) {
    if (Array.isArray(row)) {
      rows.push(
        row.map((cell) => {
          if (cell === null || cell === undefined) return null;
          if (typeof cell === 'string' || typeof cell === 'number' || typeof cell === 'boolean') {
            return cell;
          }
          return JSON.stringify(cell);
        })
      );
    } else if (row !== null && row !== undefined) {
      rows.push([asString(row) ?? String(row)]);
    }
  }
  return rows;
}

function rowsFromObjects(raw: unknown, columns: string[]): CellValue[][] {
  if (!Array.isArray(raw)) return [];
  const cols =
    columns.length > 0
      ? columns
      : raw[0] && typeof raw[0] === 'object' && !Array.isArray(raw[0])
        ? Object.keys(raw[0] as Record<string, unknown>)
        : [];
  if (cols.length === 0) return normalizeRows(raw);
  return raw.map((row) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return [];
    const record = row as Record<string, unknown>;
    return cols.map((col) => {
      const value = record[col];
      if (value === null || value === undefined) return null;
      if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
        return value;
      }
      return JSON.stringify(value);
    });
  });
}

function normalizeAction(input: unknown): AIActionPayload | null {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return null;
  const obj = input as Record<string, unknown>;

  let kind = normalizeKind(obj.action ?? obj.type ?? obj.kind);
  let body = obj;

  // Формат { "create_file": { ... } }
  if (!kind) {
    const keys = Object.keys(obj);
    if (keys.length === 1) {
      const single = normalizeKind(keys[0]);
      const nested = obj[keys[0]];
      if (single && nested && typeof nested === 'object' && !Array.isArray(nested)) {
        kind = single;
        body = nested as Record<string, unknown>;
      }
    }
  }

  if (!kind) return null;

  const name =
    asString(body.name) ||
    asString(body.fileName) ||
    asString(body.filename) ||
    asString(body.file) ||
    undefined;
  const path = asString(body.path) || asString(body.filePath) || asString(body.target) || asString(body.file);
  const content =
    asString(body.content) ??
    asString(body.text) ??
    asString(body.body) ??
    asString(body.code) ??
    asString(body.data);

  switch (kind) {
    case 'create_table': {
      const columns = asStringArray(body.columns ?? body.headers ?? body.fields);
      const rawRows = body.rows ?? body.data ?? body.records ?? body.items;
      const rows =
        Array.isArray(rawRows) && rawRows[0] && typeof rawRows[0] === 'object' && !Array.isArray(rawRows[0])
          ? rowsFromObjects(rawRows, columns)
          : normalizeRows(rawRows);
      const finalColumns =
        columns.length > 0
          ? columns
          : rawRows && Array.isArray(rawRows) && rawRows[0] && typeof rawRows[0] === 'object' && !Array.isArray(rawRows[0])
            ? Object.keys(rawRows[0] as Record<string, unknown>)
            : [];
      const format =
        body.format === 'csv' || name?.toLowerCase().endsWith('.csv')
          ? 'csv'
          : body.format === 'xlsx' || body.format === 'excel'
            ? 'xlsx'
            : undefined;
      if (!name && !path && finalColumns.length === 0 && rows.length === 0) return null;
      return {
        action: 'create_table',
        name: name || (format === 'csv' ? 'Таблица.csv' : 'Таблица.xlsx'),
        path,
        sheetName: asString(body.sheetName) || asString(body.sheet) || 'Лист 1',
        columns: finalColumns,
        rows,
        format,
      };
    }

    case 'import_table': {
      const source =
        asString(body.source) ||
        asString(body.from) ||
        asString(body.input) ||
        asString(body.file) ||
        path ||
        name;
      if (!source) return null;
      return {
        action: 'import_table',
        source,
        target:
          asString(body.target) ||
          asString(body.output) ||
          asString(body.to) ||
          asString(body.destination),
        sheetName: asString(body.sheetName) || asString(body.sheet),
      };
    }

    case 'edit_table': {
      const columns = asStringArray(body.columns ?? body.headers ?? body.fields);
      const rawRows =
        body.rows ?? body.data ?? body.records ?? body.items ?? body.values;
      const rows =
        Array.isArray(rawRows) &&
        rawRows[0] &&
        typeof rawRows[0] === 'object' &&
        !Array.isArray(rawRows[0])
          ? rowsFromObjects(rawRows, columns)
          : normalizeRows(rawRows);
      const rawMode = (asString(body.mode) || asString(body.operation) || 'replace').toLowerCase();
      if (!path && !name && rows.length === 0 && columns.length === 0) return null;
      return {
        action: 'edit_table',
        path,
        name,
        sheetName: asString(body.sheetName) || asString(body.sheet),
        mode: rawMode === 'append' ? 'append' : 'replace',
        columns,
        rows,
      };
    }

    case 'create_file': {      const finalContent = content ?? '';
      if (!name && !path && !finalContent) return null;
      return {
        action: 'create_file',
        name: name || path?.split(/[\\/]/).pop() || 'Новый файл.txt',
        path,
        content: finalContent,
        encoding: asString(body.encoding),
      };
    }

    case 'create_document': {
      const finalContent = content ?? '';
      const title = asString(body.title);
      if (!name && !path && !finalContent && !title) return null;
      return {
        action: 'create_document',
        name: name || (title ? `${title}.docx` : 'Документ.docx'),
        path,
        title,
        content: finalContent,
      };
    }

    case 'open_file': {
      if (!path && !name) return null;
      return { action: 'open_file', path: path || name, name };
    }

    case 'open_folder':
      return { action: 'open_folder', path: path || undefined };

    case 'create_folder': {
      if (!name && !path) return null;
      return { action: 'create_folder', name, path };
    }

    case 'rename_path': {
      if (!path) return null;
      const newPath = asString(body.newPath) || asString(body.to);
      const newName = asString(body.newName) || asString(body.name);
      if (!newPath && !newName) return null;
      return { action: 'rename_path', path, newPath, newName };
    }

    case 'delete_path':
      if (!path) return null;
      return { action: 'delete_path', path };

    case 'insert_into_editor': {
      const text = asString(body.text) ?? content;
      if (typeof text !== 'string' || !text.trim()) return null;
      return { action: 'insert_into_editor', text };
    }

    case 'replace_document': {
      const text = asString(body.text) ?? content;
      if (typeof text !== 'string' || !text.trim()) return null;
      return { action: 'replace_document', text };
    }

    case 'open_settings': {
      const tab = (asString(body.tab) || 'provider').toLowerCase();
      const allowed: SettingsTab[] = ['provider', 'model', 'prompt', 'keys', 'integrations', 'system'];
      return { action: 'open_settings', tab: allowed.includes(tab as SettingsTab) ? (tab as SettingsTab) : 'provider' };
    }

    case 'toggle_panel': {
      const rawPanel = (asString(body.panel) || asString(body.name) || 'ai').toLowerCase();
      const panel: PanelName =
        rawPanel === 'sidebar' || rawPanel === 'explorer' || rawPanel === 'files'
          ? 'sidebar'
          : rawPanel === 'terminal'
            ? 'terminal'
            : rawPanel === 'search'
              ? 'search'
              : 'ai';
      const rawState = (asString(body.state) || asString(body.visible) || 'toggle').toLowerCase();
      const state = rawState === 'open' || rawState === 'close' ? rawState : 'toggle';
      return { action: 'toggle_panel', panel, state };
    }

    case 'select_tab': {
      const index = typeof body.index === 'number' ? body.index : undefined;
      if (!path && !name && index === undefined) return null;
      return { action: 'select_tab', path, name, index };
    }

    case 'close_tab':
      return { action: 'close_tab', path, name };

    case 'save':
      return { action: 'save', all: body.all === true || body.scope === 'all' };

    case 'notify': {
      const message = asString(body.message) ?? content;
      if (!message) return null;
      return { action: 'notify', message };
    }

    case 'run_command': {
      const command = asString(body.command) ?? asString(body.text) ?? asString(body.line);
      if (!command) return null;
      return { action: 'run_command', command, cwd: asString(body.cwd) };
    }

    default:
      return null;
  }
}

function collect(value: unknown, out: AIActionPayload[]): void {
  if (Array.isArray(value)) {
    for (const item of value) collect(item, out);
    return;
  }
  if (!value || typeof value !== 'object') return;
  const obj = value as Record<string, unknown>;

  if (Array.isArray(obj.actions)) {
    collect(obj.actions, out);
    return;
  }

  const normalized = normalizeAction(obj);
  if (normalized) out.push(normalized);
}

/** Возвращает диапазоны всех ```-блоков, чтобы не искать действия внутри примеров кода. */
function codeFenceRanges(text: string): { start: number; end: number }[] {
  const ranges: { start: number; end: number }[] = [];
  const re = /```[\s\S]*?```/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(text)) !== null) {
    ranges.push({ start: match.index, end: match.index + match[0].length });
  }
  return ranges;
}

function insideRanges(pos: number, ranges: { start: number; end: number }[]): boolean {
  return ranges.some((range) => pos >= range.start && pos < range.end);
}

/** Находит завершающую скобку JSON-объекта с учётом строк и экранирования. */
function matchBrace(text: string, start: number): number {
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === '\\') escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') depth++;
    else if (ch === '}') {
      depth--;
      if (depth === 0) return i;
    }
  }
  return -1;
}

interface JsonSpan {
  start: number;
  end: number;
  value: unknown;
}

/** Ищет JSON-объекты действий вне обычных ```-блоков. */
function findBareActionSpans(text: string, skip: { start: number; end: number }[]): JsonSpan[] {
  const spans: JsonSpan[] = [];
  for (let i = 0; i < text.length; i++) {
    if (text[i] !== '{') continue;
    if (insideRanges(i, skip)) continue;
    const end = matchBrace(text, i);
    if (end === -1) continue;
    const raw = text.slice(i, end + 1);
    if (!/"(?:action|actions|type)"\s*:/.test(raw)) continue;
    const value = parseJsonLoose(raw);
    if (!value) continue;
    spans.push({ start: i, end: end + 1, value });
    i = end;
  }
  return spans;
}

/**
 * Извлекает все действия из ответа модели.
 *
 * Поддерживаются:
 * - ```workspace-action { ... } ``` (а также action/json/app-action);
 * - `<workspace-action>...</workspace-action>`;
 * - массив действий или объект `{ "actions": [...] }`;
 * - «голый» JSON с полем action вне блоков кода.
 */
export function parseAIActions(content: string): AIActionPayload[] {
  if (!content) return [];

  const actions: AIActionPayload[] = [];
  const fenceRe = /```[ \t]*([^\r\n`]*)\r?\n([\s\S]*?)```/g;
  let match: RegExpExecArray | null;

  while ((match = fenceRe.exec(content)) !== null) {
    const lang = (match[1] || '').trim().toLowerCase().replace(/[\s_-]+/g, '');
    const raw = (match[2] || '').trim();
    if (!raw) continue;
    const looksLikeAction = /"(?:action|actions|type)"\s*:/.test(raw);
    if (!ACTION_FENCE_LANGS.has(lang) && !looksLikeAction) continue;
    if (!looksLikeAction) continue;
    const parsed = parseActionTolerant(raw);
    if (parsed) collect(parsed, actions);
  }

  const tagRe = /<workspace[-_]action>\s*([\s\S]*?)\s*<\/workspace[-_]action>/gi;
  while ((match = tagRe.exec(content)) !== null) {
    const parsed = parseActionTolerant((match[1] || '').trim());
    if (parsed) collect(parsed, actions);
  }

  // Оборванный ответ: открывающий тег без закрывающего (лимит токенов).
  const unclosedTag = /<workspace[-_]action>\s*([\s\S]*)$/i.exec(content);
  if (unclosedTag && /"(?:action|actions|type)"\s*:/.test(unclosedTag[1])) {
    const parsed = parseActionTolerant(unclosedTag[1]);
    if (parsed) collect(parsed, actions);
  }

  // Оборванный ```workspace-action блок без закрывающих ```
  const unclosedFence = /```[ \t]*(?:json|workspace[-_]?action|action)\s*\r?\n([\s\S]*)$/i.exec(content);
  if (unclosedFence && /"(?:action|actions|type)"\s*:/.test(unclosedFence[1])) {
    const parsed = parseActionTolerant(unclosedFence[1]);
    if (parsed) collect(parsed, actions);
  }

  // «Голый» JSON — только вне блоков кода, чтобы не выполнять примеры.
  for (const span of findBareActionSpans(content, codeFenceRanges(content))) {
    collect(span.value, actions);
  }

  // Убираем дубли, возникающие при пересечении способов разбора.
  const seen = new Set<string>();
  const unique: AIActionPayload[] = [];
  for (const action of actions) {
    const signature = JSON.stringify(action);
    if (seen.has(signature)) continue;
    seen.add(signature);
    unique.push(action);
  }
  return unique;
}

/**
 * Удаляет технические блоки действий из отображаемого пользователю текста.
 */
export function stripActionBlocks(content: string): string {
  if (!content) return '';

  let cleaned = content.replace(/```[ \t]*([^\r\n`]*)\r?\n([\s\S]*?)```/g, (full, lang, body) => {
    if (!/"(?:action|actions)"\s*:/.test(body)) return full;
    const key = String(lang || '').trim().toLowerCase().replace(/[\s_-]+/g, '');
    if (ACTION_FENCE_LANGS.has(key)) return '';
    return full;
  });

  cleaned = cleaned.replace(/<workspace[-_]action>\s*[\s\S]*?\s*<\/workspace[-_]action>/gi, '');

  // Удаляем «голый» JSON действий, оставшийся вне блоков кода.
  const remainingFences = codeFenceRanges(cleaned);
  const spans = findBareActionSpans(cleaned, remainingFences).sort((a, b) => b.start - a.start);
  for (const span of spans) {
    cleaned = cleaned.slice(0, span.start) + cleaned.slice(span.end);
  }

  return cleaned.replace(/\n{3,}/g, '\n\n').trim();
}

/* ============================================================
   Парсинг Markdown-таблиц (резервный механизм)
   ============================================================ */

/**
 * Находит таблицы в стандартной Markdown-разметке (| Заголовок 1 | Заголовок 2 |).
 */
export function detectMarkdownTables(content: string): ParsedMarkdownTable[] {
  if (!content) return [];
  const results: ParsedMarkdownTable[] = [];

  // Разбиваем на строки и ищем группы строк таблицы
  const lines = content.split(/\r?\n/);
  let currentTableLines: string[] = [];

  const processTableLines = (tblLines: string[]) => {
    if (tblLines.length < 2) return;
    const cleanRows = tblLines.map((l) =>
      l
        .trim()
        .replace(/^\|/, '')
        .replace(/\|$/, '')
        .split('|')
        .map((c) => c.trim())
    );

    // Вторая строка обычно разделитель |---|---|
    const hasSeparator = cleanRows.length > 1 && cleanRows[1].every((c) => /^:?-+:?$/.test(c));
    if (hasSeparator) {
      const columns = cleanRows[0];
      const rows = cleanRows.slice(2);
      if (columns.length > 0 && rows.length > 0) {
        results.push({
          columns,
          rows,
          rawMarkdown: tblLines.join('\n'),
        });
      }
    }
  };

  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.startsWith('|') && trimmed.endsWith('|') && trimmed.includes('|')) {
      currentTableLines.push(trimmed);
    } else {
      if (currentTableLines.length > 0) {
        processTableLines(currentTableLines);
        currentTableLines = [];
      }
    }
  }

  if (currentTableLines.length > 0) {
    processTableLines(currentTableLines);
  }

  return results;
}

/* ============================================================
   Генераторы файлов таблиц (.xlsx и .csv)
   ============================================================ */

/**
 * Преобразует текстовые числа из отчётов в числа Excel.
 * «214 999» → 214999, «2 347» → 2347, «12,5» → 12.5.
 */
export function normalizeTableValue(value: CellValue): CellValue {
  if (typeof value !== 'string') return value;
  const text = value.trim();
  if (!text || text.startsWith('=')) return value;
  // Числа с пробелами-разделителями разрядов
  if (/^-?\d{1,3}(?:[ \u00A0\u202F]\d{3})+(?:[.,]\d+)?$/.test(text)) {
    const num = Number(text.replace(/[ \u00A0\u202F]/g, '').replace(',', '.'));
    return Number.isFinite(num) ? num : value;
  }
  // Обычные целые/дробные без ведущих нулей
  if (/^-?(?:0|[1-9]\d*)(?:[.,]\d+)?$/.test(text)) {
    const num = Number(text.replace(',', '.'));
    return Number.isFinite(num) ? num : value;
  }
  return value;
}

/**
 * Создаёт бинарный base64-буфер файла Excel (.xlsx) с поддержкой формул.
 */
export function generateExcelBase64(
  columns: string[] = [],
  rows: (string | number | boolean | null)[][] = [],
  sheetName = 'Лист 1'
): string {
  const wb = XLSX.utils.book_new();

  // Подготавливаем матрицу данных
  const aoa: any[][] = [];
  if (columns.length > 0) {
    aoa.push(columns);
  }
  for (const r of rows) {
    aoa.push(r.map((cell) => normalizeTableValue(cell)));
  }

  if (aoa.length === 0) {
    aoa.push(['']);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Обработка формул в ячейках
  const rowOffset = columns.length > 0 ? 1 : 0;
  rows.forEach((row, rIdx) => {
    row.forEach((cell, cIdx) => {
      if (typeof cell === 'string' && cell.startsWith('=')) {
        const cellRef = XLSX.utils.encode_cell({ r: rIdx + rowOffset, c: cIdx });
        // В файле .xlsx формулы хранятся с запятой между аргументами
        const formula = cell.slice(1).replace(/;/g, ',');
        ws[cellRef] = { t: 'n', f: formula };
      }
    });
  });

  // Автоматическая ширина колонок
  const colWidths = (columns.length > 0 ? columns : (rows[0] || [])).map((col, cIdx) => {
    let maxLen = String(col ?? '').length;
    for (const r of rows) {
      const val = String(r[cIdx] ?? '');
      if (val.length > maxLen) maxLen = val.length;
    }
    return { wch: Math.min(Math.max(maxLen + 3, 10), 40) };
  });
  ws['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Лист 1');
  return XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
}

/**
 * Создаёт текстовое содержимое формата CSV.
 */
export function generateCsvContent(
  columns: string[] = [],
  rows: (string | number | boolean | null)[][] = []
): string {
  const escapeCell = (c: any): string => {
    const str = String(c ?? '');
    if (/[",\n\r]/.test(str)) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  };

  const lines: string[] = [];
  if (columns.length > 0) {
    lines.push(columns.map(escapeCell).join(','));
  }
  for (const r of rows) {
    lines.push(r.map(escapeCell).join(','));
  }

  return lines.join('\r\n');
}

/**
 * Подбирает имя файла, избегая коллизий.
 */
export function sanitizeFileName(name: string, defaultExt = '.xlsx'): string {
  let cleaned = (name || '').trim().replace(/[\\/:*?"<>|]/g, '_');
  if (!cleaned) {
    cleaned = `Таблица_${new Date().toISOString().slice(0, 10)}`;
  }
  if (!/\.[a-zA-Z0-9]+$/.test(cleaned)) {
    cleaned += defaultExt;
  }
  return cleaned;
}

/**
 * Создаёт бинарный base64-буфер файла Word (.docx) из текста или Markdown.
 */
export async function generateDocxBase64(content = '', title?: string): Promise<string> {
  const fullContent = title && !content.startsWith('#')
    ? `# ${title}\n\n${content}`
    : (content || '<p></p>');
  const html = markdownToDocumentHtml(fullContent);
  return await htmlToDocxBase64(html);
}
