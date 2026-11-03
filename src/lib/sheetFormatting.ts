/**
 * Форматирование листов таблицы: стили ячеек, ширины/высоты, объединения,
 * закрепление областей. Метаданные хранятся отдельно от значений и
 * записываются в xlsx как пользовательские свойства документа, поэтому
 * не ломают совместимость с Excel.
 */

export type NumberFormat = 'general' | 'number' | 'currency' | 'percent' | 'date' | 'text';
export type HAlign = 'left' | 'center' | 'right';
export type VAlign = 'top' | 'middle' | 'bottom';

export interface CellBorders {
  top?: boolean;
  right?: boolean;
  bottom?: boolean;
  left?: boolean;
}

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  fill?: string;
  fontSize?: number;
  align?: HAlign;
  valign?: VAlign;
  wrap?: boolean;
  numberFormat?: NumberFormat;
  borders?: CellBorders;
}

/** Проверка данных: выпадающий список значений. */
export interface CellValidation {
  type: 'list';
  values: string[];
  allowBlank?: boolean;
}

/** Примечание к ячейке. */
export interface CellComment {
  text: string;
  author?: string;
}

export type ConditionalOperator =
  | 'greater'
  | 'less'
  | 'equal'
  | 'notEqual'
  | 'between'
  | 'contains'
  | 'empty'
  | 'notEmpty'
  | 'scale';

/** Условное форматирование диапазона. */
export interface ConditionalRule {
  id: string;
  range: string;
  operator: ConditionalOperator;
  value?: string;
  value2?: string;
  fill?: string;
  color?: string;
  bold?: boolean;
}

export type ChartType = 'bar' | 'line' | 'pie';

export type PivotAggregation = 'sum' | 'count' | 'average' | 'min' | 'max';

/** Описание сводной таблицы, построенной внутри WorkSpace. */
export interface PivotSpec {
  sourceSheet: string;
  range: string;
  rowCol: number;
  valueCol: number;
  aggregation: PivotAggregation;
}

/** Диаграмма на листе (строится по диапазону). */
export interface ChartSpec {
  id: string;
  type: ChartType;
  range: string;
  title: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface SheetMeta {
  /** Стили по адресу A1 → стиль. */
  cells: Record<string, CellStyle>;
  /** Ширина столбцов в px. */
  colWidths: Record<number, number>;
  /** Высота строк в px. */
  rowHeights: Record<number, number>;
  /** Объединённые диапазоны в формате A1:B2. */
  merges: string[];
  /** Сколько строк закреплено сверху. */
  freezeRows: number;
  /** Сколько столбцов закреплено слева. */
  freezeCols: number;
  /** Проверка данных по адресу A1. */
  validations: Record<string, CellValidation>;
  /** Примечания по адресу A1. */
  comments: Record<string, CellComment>;
  /** Правила условного форматирования. */
  conditional: ConditionalRule[];
  /** Диаграммы листа. */
  charts: ChartSpec[];
  /** Сводная, построенная на этом листе (для кнопки «Обновить»). */
  pivotSpec?: PivotSpec;
}

export type WorkbookMeta = Record<string, SheetMeta>;

export interface GridRange {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export const DEFAULT_ROW_HEIGHT = 28;
export const DEFAULT_COL_WIDTH = 104;
export const MIN_COL_WIDTH = 48;
export const MAX_COL_WIDTH = 480;
export const MIN_ROW_HEIGHT = 20;
export const MAX_ROW_HEIGHT = 240;

const BORDER_COLOR = '#8b98ab';

export function toColumnName(index: number): string {
  let name = '';
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

export function cellAddress(row: number, col: number): string {
  return `${toColumnName(col)}${row + 1}`;
}

export function parseAddress(address: string): { row: number; col: number } | null {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(address.trim());
  if (!match) return null;
  let col = 0;
  for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(match[2]) - 1, col: col - 1 };
}

export function parseRange(range: string): GridRange | null {
  const [startRaw, endRaw] = range.split(':');
  const start = parseAddress(startRaw);
  const end = parseAddress(endRaw || startRaw);
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

export function rangeToAddress(range: GridRange): string {
  return `${cellAddress(range.startRow, range.startCol)}:${cellAddress(range.endRow, range.endCol)}`;
}

export function emptySheetMeta(): SheetMeta {
  return {
    cells: {},
    colWidths: {},
    rowHeights: {},
    merges: [],
    freezeRows: 0,
    freezeCols: 0,
    validations: {},
    comments: {},
    conditional: [],
    charts: [],
  };
}

export function emptyWorkbookMeta(): WorkbookMeta {
  return {};
}

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function normalizeStyle(value: unknown): CellStyle | null {
  if (!isObject(value)) return null;
  const style: CellStyle = {};
  if (value.bold === true) style.bold = true;
  if (value.italic === true) style.italic = true;
  if (value.underline === true) style.underline = true;
  if (value.strike === true) style.strike = true;
  if (typeof value.color === 'string') style.color = value.color;
  if (typeof value.fill === 'string') style.fill = value.fill;
  if (typeof value.fontSize === 'number') style.fontSize = value.fontSize;
  if (value.align === 'left' || value.align === 'center' || value.align === 'right') {
    style.align = value.align;
  }
  if (value.valign === 'top' || value.valign === 'middle' || value.valign === 'bottom') {
    style.valign = value.valign;
  }
  if (value.wrap === true) style.wrap = true;
  if (
    value.numberFormat === 'general' ||
    value.numberFormat === 'number' ||
    value.numberFormat === 'currency' ||
    value.numberFormat === 'percent' ||
    value.numberFormat === 'date' ||
    value.numberFormat === 'text'
  ) {
    style.numberFormat = value.numberFormat;
  }
  if (isObject(value.borders)) {
    const borders: CellBorders = {};
    if (value.borders.top === true) borders.top = true;
    if (value.borders.top === false) borders.top = false;
    if (value.borders.right === true) borders.right = true;
    if (value.borders.right === false) borders.right = false;
    if (value.borders.bottom === true) borders.bottom = true;
    if (value.borders.bottom === false) borders.bottom = false;
    if (value.borders.left === true) borders.left = true;
    if (value.borders.left === false) borders.left = false;
    style.borders = borders;
  }
  return Object.keys(style).length > 0 ? style : null;
}

/** Безопасно разбирает метаданные из файла (повреждённые значения игнорируются). */
export function normalizeWorkbookMeta(raw: unknown): WorkbookMeta {
  if (!isObject(raw)) return {};
  const result: WorkbookMeta = {};
  for (const [name, value] of Object.entries(raw)) {
    if (!isObject(value)) continue;
    const meta = emptySheetMeta();
    if (isObject(value.cells)) {
      for (const [address, style] of Object.entries(value.cells)) {
        const normalized = normalizeStyle(style);
        if (normalized && parseAddress(address)) meta.cells[address] = normalized;
      }
    }
    if (isObject(value.colWidths)) {
      for (const [col, width] of Object.entries(value.colWidths)) {
        const index = Number(col);
        if (Number.isInteger(index) && typeof width === 'number' && width > 0) {
          meta.colWidths[index] = Math.min(MAX_COL_WIDTH, Math.max(MIN_COL_WIDTH, width));
        }
      }
    }
    if (isObject(value.rowHeights)) {
      for (const [row, height] of Object.entries(value.rowHeights)) {
        const index = Number(row);
        if (Number.isInteger(index) && typeof height === 'number' && height > 0) {
          meta.rowHeights[index] = Math.min(MAX_ROW_HEIGHT, Math.max(MIN_ROW_HEIGHT, height));
        }
      }
    }
    if (Array.isArray(value.merges)) {
      const unique = new Set<string>();
      for (const item of value.merges) {
        if (typeof item === 'string' && parseRange(item)) unique.add(item);
      }
      meta.merges = [...unique];
    }
    if (typeof value.freezeRows === 'number' && value.freezeRows > 0) {
      meta.freezeRows = Math.min(20, Math.floor(value.freezeRows));
    }
    if (typeof value.freezeCols === 'number' && value.freezeCols > 0) {
      meta.freezeCols = Math.min(20, Math.floor(value.freezeCols));
    }
    if (isObject(value.validations)) {
      for (const [address, raw] of Object.entries(value.validations)) {
        if (!isObject(raw) || !Array.isArray(raw.values)) continue;
        const values = raw.values.filter((item): item is string => typeof item === 'string');
        if (values.length > 0 && parseAddress(address)) {
          meta.validations[address] = { type: 'list', values };
        }
      }
    }
    if (isObject(value.comments)) {
      for (const [address, raw] of Object.entries(value.comments)) {
        if (!isObject(raw) || typeof raw.text !== 'string' || !parseAddress(address)) continue;
        meta.comments[address] = {
          text: raw.text,
          author: typeof raw.author === 'string' ? raw.author : undefined,
        };
      }
    }
    if (Array.isArray(value.conditional)) {
      for (const raw of value.conditional) {
        if (!isObject(raw) || typeof raw.range !== 'string' || !parseRange(raw.range)) continue;
        if (typeof raw.id !== 'string') continue;
        const operators: ConditionalOperator[] = [
          'greater',
          'less',
          'equal',
          'notEqual',
          'between',
          'contains',
          'empty',
          'notEmpty',
          'scale',
        ];
        if (!operators.includes(raw.operator as ConditionalOperator)) continue;
        meta.conditional.push({
          id: raw.id,
          range: raw.range,
          operator: raw.operator as ConditionalOperator,
          value: typeof raw.value === 'string' ? raw.value : undefined,
          value2: typeof raw.value2 === 'string' ? raw.value2 : undefined,
          fill: typeof raw.fill === 'string' ? raw.fill : undefined,
          color: typeof raw.color === 'string' ? raw.color : undefined,
          bold: raw.bold === true,
        });
      }
    }
    if (Array.isArray(value.charts)) {
      for (const raw of value.charts) {
        if (!isObject(raw) || typeof raw.range !== 'string' || !parseRange(raw.range)) continue;
        if (typeof raw.id !== 'string') continue;
        if (raw.type !== 'bar' && raw.type !== 'line' && raw.type !== 'pie') continue;
        meta.charts.push({
          id: raw.id,
          type: raw.type,
          range: raw.range,
          title: typeof raw.title === 'string' ? raw.title : 'Диаграмма',
          x: typeof raw.x === 'number' ? raw.x : 40,
          y: typeof raw.y === 'number' ? raw.y : 40,
          width: typeof raw.width === 'number' ? raw.width : 420,
          height: typeof raw.height === 'number' ? raw.height : 260,
        });
      }
    }
    if (isObject(value.pivotSpec)) {
      const spec = value.pivotSpec;
      if (
        typeof spec.sourceSheet === 'string' &&
        typeof spec.range === 'string' &&
        parseRange(spec.range) &&
        typeof spec.rowCol === 'number' &&
        typeof spec.valueCol === 'number' &&
        (spec.aggregation === 'sum' ||
          spec.aggregation === 'count' ||
          spec.aggregation === 'average' ||
          spec.aggregation === 'min' ||
          spec.aggregation === 'max')
      ) {
        meta.pivotSpec = {
          sourceSheet: spec.sourceSheet,
          range: spec.range,
          rowCol: spec.rowCol,
          valueCol: spec.valueCol,
          aggregation: spec.aggregation,
        };
      }
    }
    result[name] = meta;
  }
  return result;
}

/**
 * Кодирует метаданные для пользовательского свойства xlsx.
 * Base64 обязателен: SheetJS искажает не-ASCII символы в Custprops.
 */
export function encodeWorkbookMeta(meta: WorkbookMeta): string {
  const json = JSON.stringify(meta);
  try {
    return `b64:${btoa(unescape(encodeURIComponent(json)))}`;
  } catch {
    return json;
  }
}

/** Декодирует метаданные; поддерживает старый формат с чистым JSON. */
export function decodeWorkbookMeta(raw: unknown): WorkbookMeta {
  if (typeof raw !== 'string' || !raw) return {};
  let json = raw;
  if (raw.startsWith('b64:')) {
    try {
      json = decodeURIComponent(escape(atob(raw.slice(4))));
    } catch {
      return {};
    }
  }
  try {
    return normalizeWorkbookMeta(JSON.parse(json));
  } catch {
    return {};
  }
}

export function cloneWorkbookMeta(meta: WorkbookMeta): WorkbookMeta {
  const next: WorkbookMeta = {};
  for (const [name, sheet] of Object.entries(meta)) {
    next[name] = {
      cells: Object.fromEntries(
        Object.entries(sheet.cells).map(([address, style]) => [
          address,
          { ...style, borders: style.borders ? { ...style.borders } : undefined },
        ])
      ),
      colWidths: { ...sheet.colWidths },
      rowHeights: { ...sheet.rowHeights },
      merges: [...sheet.merges],
      freezeRows: sheet.freezeRows,
      freezeCols: sheet.freezeCols,
      validations: Object.fromEntries(
        Object.entries(sheet.validations).map(([address, validation]) => [
          address,
          { type: 'list' as const, values: [...validation.values], allowBlank: validation.allowBlank },
        ])
      ),
      comments: Object.fromEntries(
        Object.entries(sheet.comments).map(([address, comment]) => [address, { ...comment }])
      ),
      conditional: sheet.conditional.map((rule) => ({ ...rule })),
      charts: sheet.charts.map((chart) => ({ ...chart })),
      pivotSpec: sheet.pivotSpec ? { ...sheet.pivotSpec } : undefined,
    };
  }
  return next;
}

export function ensureSheetMeta(meta: WorkbookMeta, sheet: string): SheetMeta {
  return meta[sheet] ?? emptySheetMeta();
}

export function styleAt(meta: SheetMeta, row: number, col: number): CellStyle | undefined {
  return meta.cells[cellAddress(row, col)];
}

/** Применяет частичный стиль ко всем ячейкам диапазона. */
export function applyStyleToRange(
  meta: WorkbookMeta,
  sheet: string,
  range: GridRange,
  patch: Partial<CellStyle>
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      const address = cellAddress(row, col);
      const merged: CellStyle = { ...target.cells[address], ...patch };
      if (patch.borders) {
        merged.borders = { ...target.cells[address]?.borders, ...patch.borders };
      }
      if (Object.keys(merged).length === 0) delete target.cells[address];
      else target.cells[address] = merged;
    }
  }
  return next;
}

/** Границы выбранного диапазона: внешние рёбра или все ячейки. */
export function borderPatchForRange(
  range: GridRange,
  row: number,
  col: number,
  mode: 'all' | 'outer' | 'none'
): CellBorders {
  if (mode === 'none') return { top: false, right: false, bottom: false, left: false };
  if (mode === 'all') return { top: true, right: true, bottom: true, left: true };
  return {
    top: row === range.startRow,
    bottom: row === range.endRow,
    left: col === range.startCol,
    right: col === range.endCol,
  };
}

/** Снимает пользовательское форматирование с диапазона. */
export function clearStyleRange(
  meta: WorkbookMeta,
  sheet: string,
  range: GridRange
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      delete target.cells[cellAddress(row, col)];
    }
  }
  return next;
}

export function applyBordersToRange(
  meta: WorkbookMeta,
  sheet: string,
  range: GridRange,
  mode: 'all' | 'outer' | 'none'
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      const address = cellAddress(row, col);
      const merged: CellStyle = {
        ...target.cells[address],
        borders: borderPatchForRange(range, row, col, mode),
      };
      target.cells[address] = merged;
    }
  }
  return next;
}

export function setColWidth(meta: WorkbookMeta, sheet: string, col: number, width: number): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  target.colWidths[col] = Math.round(Math.min(MAX_COL_WIDTH, Math.max(MIN_COL_WIDTH, width)));
  return next;
}

export function setRowHeight(meta: WorkbookMeta, sheet: string, row: number, height: number): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  target.rowHeights[row] = Math.round(Math.min(MAX_ROW_HEIGHT, Math.max(MIN_ROW_HEIGHT, height)));
  return next;
}

/* ---------------- Объединение ячеек ---------------- */

/** Убирает пересекающиеся объединения и добавляет новое. */
export function addMerge(meta: WorkbookMeta, sheet: string, range: GridRange): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  const overlaps = (a: GridRange, b: GridRange) =>
    a.startRow <= b.endRow && a.endRow >= b.startRow && a.startCol <= b.endCol && a.endCol >= b.startCol;
  target.merges = target.merges.filter((item) => {
    const parsed = parseRange(item);
    return parsed ? !overlaps(parsed, range) : false;
  });
  if (range.startRow !== range.endRow || range.startCol !== range.endCol) {
    target.merges.push(rangeToAddress(range));
  }
  return next;
}

export function removeMergesInRange(meta: WorkbookMeta, sheet: string, range: GridRange): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  target.merges = target.merges.filter((item) => {
    const parsed = parseRange(item);
    if (!parsed) return false;
    const intersects =
      parsed.startRow <= range.endRow &&
      parsed.endRow >= range.startRow &&
      parsed.startCol <= range.endCol &&
      parsed.endCol >= range.startCol;
    return !intersects;
  });
  return next;
}

export interface MergeInfo {
  anchors: Map<string, { rowSpan: number; colSpan: number }>;
  covered: Set<string>;
}

export function buildMergeInfo(merges: string[]): MergeInfo {
  const anchors = new Map<string, { rowSpan: number; colSpan: number }>();
  const covered = new Set<string>();
  for (const item of merges) {
    const range = parseRange(item);
    if (!range) continue;
    const rowSpan = range.endRow - range.startRow + 1;
    const colSpan = range.endCol - range.startCol + 1;
    if (rowSpan === 1 && colSpan === 1) continue;
    anchors.set(cellAddress(range.startRow, range.startCol), { rowSpan, colSpan });
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        if (row === range.startRow && col === range.startCol) continue;
        covered.add(cellAddress(row, col));
      }
    }
  }
  return { anchors, covered };
}

/* ---------------- Закрепление ---------------- */

export function setFreeze(meta: WorkbookMeta, sheet: string, rows: number, cols: number): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  target.freezeRows = Math.max(0, rows);
  target.freezeCols = Math.max(0, cols);
  return next;
}

/* ---------------- Сдвиги при удалении строк/столбцов ---------------- */

function shiftAddress(address: string, rowDelta: number, colDelta: number): string | null {
  const parsed = parseAddress(address);
  if (!parsed) return null;
  const row = parsed.row + rowDelta;
  const col = parsed.col + colDelta;
  if (row < 0 || col < 0) return null;
  return cellAddress(row, col);
}

function shiftRange(range: string, rowDelta: number, colDelta: number): string | null {
  const parsed = parseRange(range);
  if (!parsed) return null;
  const startRow = parsed.startRow + rowDelta;
  const startCol = parsed.startCol + colDelta;
  const endRow = parsed.endRow + rowDelta;
  const endCol = parsed.endCol + colDelta;
  if (endRow < 0 || endCol < 0) return null;
  return rangeToAddress({
    startRow: Math.max(0, startRow),
    startCol: Math.max(0, startCol),
    endRow: Math.max(0, endRow),
    endCol: Math.max(0, endCol),
  });
}

/** Пересчитывает адреса после удаления одной строки. */
export function removeRowFromMeta(meta: WorkbookMeta, sheet: string, row: number): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  const cells: Record<string, CellStyle> = {};
  for (const [address, style] of Object.entries(target.cells)) {
    const parsed = parseAddress(address);
    if (!parsed || parsed.row === row) continue;
    const shifted = shiftAddress(address, parsed.row > row ? -1 : 0, 0);
    if (shifted) cells[shifted] = style;
  }
  target.cells = cells;
  const heights: Record<number, number> = {};
  for (const [index, height] of Object.entries(target.rowHeights)) {
    const r = Number(index);
    if (r === row) continue;
    heights[r > row ? r - 1 : r] = height;
  }
  target.rowHeights = heights;
  // Проверки данных и примечания переезжают вместе со строками
  const validations: Record<string, CellValidation> = {};
  for (const [address, validation] of Object.entries(target.validations)) {
    const parsed = parseAddress(address);
    if (!parsed || parsed.row === row) continue;
    const shifted = shiftAddress(address, parsed.row > row ? -1 : 0, 0);
    if (shifted) validations[shifted] = validation;
  }
  target.validations = validations;
  const comments: Record<string, CellComment> = {};
  for (const [address, comment] of Object.entries(target.comments)) {
    const parsed = parseAddress(address);
    if (!parsed || parsed.row === row) continue;
    const shifted = shiftAddress(address, parsed.row > row ? -1 : 0, 0);
    if (shifted) comments[shifted] = comment;
  }
  target.comments = comments;
  target.merges = target.merges
    .map((item) => {
      const parsed = parseRange(item);
      if (!parsed) return null;
      if (row < parsed.startRow) return shiftRange(item, -1, 0);
      if (row > parsed.endRow) return item;
      // Удалённая строка внутри объединения — сжимаем диапазон или снимаем его
      if (parsed.startRow === parsed.endRow) return null;
      return rangeToAddress({
        startRow: parsed.startRow,
        endRow: parsed.endRow - 1,
        startCol: parsed.startCol,
        endCol: parsed.endCol,
      });
    })
    .filter((item): item is string => Boolean(item));
  if (target.freezeRows > row) target.freezeRows -= 1;
  return next;
}

/** Пересчитывает адреса после удаления одного столбца. */
export function removeColumnFromMeta(meta: WorkbookMeta, sheet: string, col: number): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  const cells: Record<string, CellStyle> = {};
  for (const [address, style] of Object.entries(target.cells)) {
    const parsed = parseAddress(address);
    if (!parsed || parsed.col === col) continue;
    const shifted = shiftAddress(address, 0, parsed.col > col ? -1 : 0);
    if (shifted) cells[shifted] = style;
  }
  target.cells = cells;
  const widths: Record<number, number> = {};
  for (const [index, width] of Object.entries(target.colWidths)) {
    const c = Number(index);
    if (c === col) continue;
    widths[c > col ? c - 1 : c] = width;
  }
  target.colWidths = widths;
  const validations: Record<string, CellValidation> = {};
  for (const [address, validation] of Object.entries(target.validations)) {
    const parsed = parseAddress(address);
    if (!parsed || parsed.col === col) continue;
    const shifted = shiftAddress(address, 0, parsed.col > col ? -1 : 0);
    if (shifted) validations[shifted] = validation;
  }
  target.validations = validations;
  const comments: Record<string, CellComment> = {};
  for (const [address, comment] of Object.entries(target.comments)) {
    const parsed = parseAddress(address);
    if (!parsed || parsed.col === col) continue;
    const shifted = shiftAddress(address, 0, parsed.col > col ? -1 : 0);
    if (shifted) comments[shifted] = comment;
  }
  target.comments = comments;
  target.merges = target.merges
    .map((item) => {
      const parsed = parseRange(item);
      if (!parsed) return null;
      if (col < parsed.startCol) return shiftRange(item, 0, -1);
      if (col > parsed.endCol) return item;
      if (parsed.startCol === parsed.endCol) return null;
      return rangeToAddress({
        startRow: parsed.startRow,
        endRow: parsed.endRow,
        startCol: parsed.startCol,
        endCol: parsed.endCol - 1,
      });
    })
    .filter((item): item is string => Boolean(item));
  if (target.freezeCols > col) target.freezeCols -= 1;
  return next;
}

/** Переставляет стили строк вместе с сортировкой значений. */
export function reorderRowsInMeta(
  meta: WorkbookMeta,
  sheet: string,
  startRow: number,
  endRow: number,
  order: number[]
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  const cells: Record<string, CellStyle> = {};
  for (const [address, style] of Object.entries(target.cells)) {
    const parsed = parseAddress(address);
    if (!parsed || parsed.row < startRow || parsed.row > endRow) {
      cells[address] = style;
      continue;
    }
    const original = parsed.row - startRow;
    const newIndex = order.indexOf(original);
    if (newIndex === -1) continue;
    cells[cellAddress(startRow + newIndex, parsed.col)] = style;
  }
  target.cells = cells;
  const heights: Record<number, number> = {};
  for (const [index, height] of Object.entries(target.rowHeights)) {
    const r = Number(index);
    if (r < startRow || r > endRow) {
      heights[r] = height;
      continue;
    }
    const newIndex = order.indexOf(r - startRow);
    if (newIndex !== -1) heights[startRow + newIndex] = height;
  }
  target.rowHeights = heights;
  const validations: Record<string, CellValidation> = {};
  for (const [address, validation] of Object.entries(target.validations)) {
    const parsed = parseAddress(address);
    if (!parsed) continue;
    if (parsed.row < startRow || parsed.row > endRow) {
      validations[address] = validation;
      continue;
    }
    const newIndex = order.indexOf(parsed.row - startRow);
    if (newIndex !== -1) validations[cellAddress(startRow + newIndex, parsed.col)] = validation;
  }
  target.validations = validations;
  const comments: Record<string, CellComment> = {};
  for (const [address, comment] of Object.entries(target.comments)) {
    const parsed = parseAddress(address);
    if (!parsed) continue;
    if (parsed.row < startRow || parsed.row > endRow) {
      comments[address] = comment;
      continue;
    }
    const newIndex = order.indexOf(parsed.row - startRow);
    if (newIndex !== -1) comments[cellAddress(startRow + newIndex, parsed.col)] = comment;
  }
  target.comments = comments;
  return next;
}

/* ---------------- Отображение значений ---------------- */

const numberFormatter = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 2 });
const currencyFormatter = new Intl.NumberFormat('ru-RU', {
  style: 'currency',
  currency: 'RUB',
  maximumFractionDigits: 2,
});

function parseNumeric(value: string): number | null {
  const normalized = value.replace(/\s/g, '').replace(',', '.');
  if (!normalized || !/^-?\d+(\.\d+)?$/.test(normalized)) return null;
  const num = Number(normalized);
  return Number.isFinite(num) ? num : null;
}

/** Форматирует отображаемое значение с учётом числового формата ячейки. */
export function formatCellValue(value: string, format?: NumberFormat): string {
  if (!format || format === 'general' || format === 'text' || value === '') return value;
  const num = parseNumeric(value);
  if (num === null) return value;
  switch (format) {
    case 'number':
      return numberFormatter.format(num);
    case 'currency':
      return currencyFormatter.format(num);
    case 'percent':
      return `${numberFormatter.format(num * 100)}%`;
    case 'date': {
      // Excel-серийная дата: 1 = 01.01.1900
      if (num > 20000 && num < 80000) {
        const date = new Date(Date.UTC(1899, 11, 30) + num * 86400000);
        return date.toLocaleDateString('ru-RU');
      }
      const parsed = new Date(value);
      return Number.isNaN(parsed.getTime()) ? value : parsed.toLocaleDateString('ru-RU');
    }
    default:
      return value;
  }
}

/** Код числового формата Excel для записи в файл. */
export function numberFormatCode(format?: NumberFormat): string | undefined {
  switch (format) {
    case 'number':
      return '#,##0.00';
    case 'currency':
      return '#,##0.00 ₽';
    case 'percent':
      return '0%';
    case 'date':
      return 'dd.mm.yyyy';
    case 'text':
      return '@';
    default:
      return undefined;
  }
}

/** CSS-свойства ячейки из стиля. */
export function styleToCss(style?: CellStyle): Record<string, string | number> {
  if (!style) return {};
  const css: Record<string, string | number> = {};
  if (style.bold) css.fontWeight = 700;
  if (style.italic) css.fontStyle = 'italic';
  if (style.underline || style.strike) {
    css.textDecoration = [style.underline ? 'underline' : '', style.strike ? 'line-through' : '']
      .filter(Boolean)
      .join(' ');
  }
  if (style.color) css.color = style.color;
  if (style.fill) css.backgroundColor = style.fill;
  if (style.fontSize) css.fontSize = `${style.fontSize}px`;
  if (style.align) css.textAlign = style.align;
  if (style.valign) css.verticalAlign = style.valign;
  if (style.wrap) {
    css.whiteSpace = 'pre-wrap';
    css.wordBreak = 'break-word';
  }
  if (style.borders) {
    const side = (on?: boolean) => (on === true ? `1px solid ${BORDER_COLOR}` : on === false ? 'none' : undefined);
    const top = side(style.borders.top);
    const right = side(style.borders.right);
    const bottom = side(style.borders.bottom);
    const left = side(style.borders.left);
    if (top) css.borderTop = top;
    if (right) css.borderRight = right;
    if (bottom) css.borderBottom = bottom;
    if (left) css.borderLeft = left;
  }
  return css;
}

/** Смещения строк для виртуализации с переменной высотой. */
export function buildRowOffsets(rowHeights: Record<number, number>, count: number): number[] {
  const offsets = new Array<number>(count + 1);
  offsets[0] = 0;
  for (let row = 0; row < count; row++) {
    offsets[row + 1] = offsets[row] + (rowHeights[row] || DEFAULT_ROW_HEIGHT);
  }
  return offsets;
}

/** Индекс первой строки, попадающей в видимую область. */
export function findRowAtOffset(offsets: number[], offset: number): number {
  let low = 0;
  let high = offsets.length - 1;
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    if (offsets[mid] <= offset) low = mid + 1;
    else high = mid;
  }
  return Math.max(0, low - 1);
}

/* ---------------- Проверка данных ---------------- */

export function setValidationForRange(
  meta: WorkbookMeta,
  sheet: string,
  range: GridRange,
  values: string[]
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      target.validations[cellAddress(row, col)] = { type: 'list', values: [...values] };
    }
  }
  return next;
}

export function clearValidationForRange(
  meta: WorkbookMeta,
  sheet: string,
  range: GridRange
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      delete target.validations[cellAddress(row, col)];
    }
  }
  return next;
}

/* ---------------- Примечания ---------------- */

export function setCommentForCell(
  meta: WorkbookMeta,
  sheet: string,
  row: number,
  col: number,
  text: string | null
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  const address = cellAddress(row, col);
  if (text && text.trim()) target.comments[address] = { text: text.trim() };
  else delete target.comments[address];
  return next;
}

/* ---------------- Условное форматирование ---------------- */

export function addConditionalRule(
  meta: WorkbookMeta,
  sheet: string,
  rule: ConditionalRule
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  target.conditional = target.conditional.filter((item) => item.id !== rule.id).concat(rule);
  return next;
}

export function removeConditionalRule(
  meta: WorkbookMeta,
  sheet: string,
  id: string
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  target.conditional = target.conditional.filter((rule) => rule.id !== id);
  return next;
}

export function ruleContainsCell(rule: ConditionalRule, row: number, col: number): boolean {
  const range = parseRange(rule.range);
  if (!range) return false;
  return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol;
}

function numericValue(value: string): number | null {
  const normalized = value.replace(/\s/g, '').replace(',', '.');
  if (normalized === '' || !/^-?\d+(\.\d+)?$/.test(normalized)) return null;
  const num = Number(normalized);
  return Number.isFinite(num) ? num : null;
}

/** Проверяет условие правила для отображаемого значения ячейки. */
export function ruleMatches(rule: ConditionalRule, value: string): boolean {
  const text = value.trim();
  const lower = text.toLowerCase();
  const target = (rule.value || '').trim();
  switch (rule.operator) {
    case 'empty':
      return text === '';
    case 'notEmpty':
      return text !== '';
    case 'contains':
      return target !== '' && lower.includes(target.toLowerCase());
    case 'greater':
    case 'less':
    case 'between': {
      const current = numericValue(text);
      const first = numericValue(target);
      if (current === null || first === null) return false;
      if (rule.operator === 'greater') return current > first;
      if (rule.operator === 'less') return current < first;
      const second = numericValue(rule.value2 || '');
      return second !== null && current >= Math.min(first, second) && current <= Math.max(first, second);
    }
    case 'equal':
      return lower === target.toLowerCase();
    case 'notEqual':
      return lower !== target.toLowerCase();
    case 'scale':
      return numericValue(text) !== null;
    default:
      return false;
  }
}

/** CSS-патч условного форматирования (без цветовой шкалы). */
export function conditionalCss(rule: ConditionalRule): Record<string, string | number> {
  const css: Record<string, string | number> = {};
  if (rule.fill) css.backgroundColor = rule.fill;
  if (rule.color) css.color = rule.color;
  if (rule.bold) css.fontWeight = 700;
  return css;
}

/** Цвет цветовой шкалы: 0 — красный, 0.5 — жёлтый, 1 — зелёный. */
export function scaleColor(ratio: number): string {
  const clamped = Math.max(0, Math.min(1, ratio));
  const stops: [number, [number, number, number]][] = [
    [0, [254, 202, 202]],
    [0.5, [254, 240, 138]],
    [1, [187, 247, 208]],
  ];
  const upperIndex = clamped <= 0.5 ? 1 : 2;
  const lower = stops[upperIndex - 1];
  const upper = stops[upperIndex];
  const local = (clamped - lower[0]) / Math.max(0.0001, upper[0] - lower[0]);
  const channels = lower[1].map((value, index) =>
    Math.round(value + (upper[1][index] - value) * local)
  );
  return `rgb(${channels[0]}, ${channels[1]}, ${channels[2]})`;
}

/* ---------------- Диаграммы ---------------- */

export function addChartSpec(meta: WorkbookMeta, sheet: string, chart: ChartSpec): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  target.charts = target.charts.concat(chart);
  return next;
}

export function updateChartSpec(
  meta: WorkbookMeta,
  sheet: string,
  id: string,
  patch: Partial<ChartSpec>
): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  target.charts = target.charts.map((chart) => (chart.id === id ? { ...chart, ...patch } : chart));
  return next;
}

export function removeChartSpec(meta: WorkbookMeta, sheet: string, id: string): WorkbookMeta {
  const next = cloneWorkbookMeta(meta);
  const target = next[sheet] ?? emptySheetMeta();
  next[sheet] = target;
  target.charts = target.charts.filter((chart) => chart.id !== id);
  return next;
}

/** Диапазон данных вокруг ячейки (до первой пустой строки/столбца). */
export function detectUsedRange(grid: (string | number | boolean | null)[][], row: number, col: number): GridRange {
  const rows = grid.length;
  const cols = grid[0]?.length ?? 0;
  const isFilled = (r: number, c: number) => String(grid[r]?.[c] ?? '').trim() !== '';
  let startRow = row;
  let endRow = row;
  let startCol = col;
  let endCol = col;
  if (!isFilled(row, col)) {
    // Пустая ячейка — берём весь занятый прямоугольник листа
    startRow = 0;
    startCol = 0;
    endRow = rows - 1;
    endCol = cols - 1;
  } else {
    while (startRow > 0 && Array.from({ length: cols }, (_, c) => isFilled(startRow - 1, c)).some(Boolean)) startRow--;
    while (endRow < rows - 1 && Array.from({ length: cols }, (_, c) => isFilled(endRow + 1, c)).some(Boolean)) endRow++;
    while (startCol > 0 && Array.from({ length: rows }, (_, r) => isFilled(r, startCol - 1)).some(Boolean)) startCol--;
    while (endCol < cols - 1 && Array.from({ length: rows }, (_, r) => isFilled(r, endCol + 1)).some(Boolean)) endCol++;
  }
  // Обрезаем пустые края
  while (endRow > startRow && !Array.from({ length: cols }, (_, c) => isFilled(endRow, c)).some(Boolean)) endRow--;
  while (endCol > startCol && !Array.from({ length: rows }, (_, r) => isFilled(r, endCol)).some(Boolean)) endCol--;
  return { startRow, endRow, startCol, endCol };
}
