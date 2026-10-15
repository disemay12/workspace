import * as XLSX from 'xlsx';

/* ============================================================
   Типы действий AI по управлению программой WorkSpace
   ============================================================ */

export interface AICreateTablePayload {
  action: 'create_table';
  name?: string;
  sheetName?: string;
  columns?: string[];
  rows?: (string | number | boolean | null)[][];
  format?: 'xlsx' | 'csv';
}

export interface AICreateFilePayload {
  action: 'create_file';
  name?: string;
  path?: string;
  content?: string;
}

export interface AIOpenFilePayload {
  action: 'open_file';
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

export type AIActionPayload =
  | AICreateTablePayload
  | AICreateFilePayload
  | AIOpenFilePayload
  | AIInsertEditorPayload
  | AIReplaceDocPayload;

export interface ParsedMarkdownTable {
  columns: string[];
  rows: string[][];
  rawMarkdown: string;
}

/* ============================================================
   Парсер действий из ответа модели
   ============================================================ */

/**
 * Извлекает все блоки действий из текста сообщения AI.
 * Поддерживает блоки:
 * ```workspace-action
 * { "action": "create_table", ... }
 * ```
 * а также JSON-блоки со свойством "action" и теги <workspace-action>.
 */
export function parseAIActions(content: string): AIActionPayload[] {
  if (!content) return [];
  const actions: AIActionPayload[] = [];

  // 1. Поиск ```workspace-action ... ```
  const actionBlockRegex = /```(?:workspace-action|workspace_action)\s*\n([\s\S]*?)\n```/gi;
  let match: RegExpExecArray | null;

  while ((match = actionBlockRegex.exec(content)) !== null) {
    const raw = match[1].trim();
    const parsed = tryParseJson(raw);
    if (parsed && isValidAction(parsed)) {
      actions.push(parsed);
    }
  }

  // 2. Поиск тегов <workspace-action> ... </workspace-action>
  const tagRegex = /<workspace[-_]action>\s*([\s\S]*?)\s*<\/workspace[-_]action>/gi;
  while ((match = tagRegex.exec(content)) !== null) {
    const raw = match[1].trim();
    const parsed = tryParseJson(raw);
    if (parsed && isValidAction(parsed)) {
      actions.push(parsed);
    }
  }

  // 3. Если не найдено выше, ищем блок ```json ... ``` с "action": "create_table" и т.п.
  if (actions.length === 0) {
    const jsonBlockRegex = /```(?:json)?\s*\n(\{\s*"action"\s*:\s*"(?:create_table|create_file|open_file|insert_into_editor|replace_document)"[\s\S]*?\})\n```/gi;
    while ((match = jsonBlockRegex.exec(content)) !== null) {
      const raw = match[1].trim();
      const parsed = tryParseJson(raw);
      if (parsed && isValidAction(parsed)) {
        actions.push(parsed);
      }
    }
  }

  return actions;
}

/**
 * Проверяет, является ли объект валидным действием AI.
 */
function isValidAction(obj: any): obj is AIActionPayload {
  if (!obj || typeof obj !== 'object') return false;
  const action = obj.action;
  if (!action || typeof action !== 'string') return false;

  switch (action) {
    case 'create_table':
      return Boolean(Array.isArray(obj.columns) || Array.isArray(obj.rows) || obj.name);
    case 'create_file':
      return Boolean(obj.name || obj.path || typeof obj.content === 'string');
    case 'open_file':
      return Boolean(obj.path || obj.name);
    case 'insert_into_editor':
    case 'replace_document':
      return typeof obj.text === 'string';
    default:
      return false;
  }
}

/**
 * Безопасный парсинг JSON с устранением частых огрехов (висячие запятые).
 */
function tryParseJson(raw: string): any {
  try {
    return JSON.parse(raw);
  } catch {
    try {
      // Удаляем завершающие запятые перед закрывающими скобками
      const cleaned = raw.replace(/,\s*([\]}])/g, '$1');
      return JSON.parse(cleaned);
    } catch {
      return null;
    }
  }
}

/**
 * Удаляет технические блоки действий из отображаемого пользователю текста.
 */
export function stripActionBlocks(content: string): string {
  if (!content) return '';
  let cleaned = content
    .replace(/```(?:workspace-action|workspace_action)\s*\n[\s\S]*?\n```/gi, '')
    .replace(/<workspace[-_]action>\s*[\s\S]*?\s*<\/workspace[-_]action>/gi, '')
    .trim();

  return cleaned;
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
    aoa.push(r);
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
