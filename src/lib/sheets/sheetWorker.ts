/**
 * Воркер разбора таблиц: CSV и XLSX больших размеров.
 *
 * Работает в отдельном потоке, поэтому интерфейс не замирает на разборе книги
 * в десятки мегабайт. Файл читает главный поток (`fs:readBinarySlice`) — воркер
 * сам запрашивает куски и потому не держит в памяти base64-копию целиком.
 *
 * Обмен: `{ id, kind, ... }` → `{ id, ok, result | error }`, плюс встречные
 * запросы `{ id, kind: 'read' }` из воркера в клиент.
 */
import * as XLSX from 'xlsx';
import { fromExcelFormula, toExcelFormula } from '../formulaEngine';

export interface SheetMeta {
  name: string;
  rows: number;
  cols: number;
}

export interface CellEdit {
  r: number;
  c: number;
  v: string | number | null;
}

export type WorkerRequest =
  | { id: number; kind: 'openCsv'; path: string; name: string; encoding: string; chunkSize?: number }
  | { id: number; kind: 'openXlsx'; bytes: ArrayBuffer; name: string; ext: string }
  | { id: number; kind: 'getWindow'; sheet: string; start: number; count: number }
  | { id: number; kind: 'setCells'; sheet: string; edits: CellEdit[] }
  | { id: number; kind: 'save' }
  | { id: number; kind: 'meta' }
  | { id: number; kind: 'close' };

export type WorkerResponse =
  | { id: number; ok: true; result?: unknown }
  | { id: number; ok: false; error: string }
  | { id: number; kind: 'read'; path: string; offset: number; length: number }
  | { id: number; kind: 'progress'; phase: string; done: number; total: number };

type Row = (string | number | boolean | null)[];

/* ============================================================
   CSV: потоковый разбор с индексом строк
   ============================================================ */

class CsvTable {
  /** Смещения начала строк в байтах файла: текст целиком не держим. */
  private rowStarts: number[] = [0];
  private rowCount = 0;
  private colCount = 0;
  /** Состояние «внутри кавычек» переходит между кусками файла. */
  private inQuotes = false;
  private length = 0;
  private lastByte = -1;
  private endsWithNewline = false;
  /** Файл был в CRLF — сохраняем те же переводы строк. */
  private crlf = false;
  /** Правки поверх файла: применяются при чтении окна и при сохранении. */
  private readonly edits = new Map<number, Map<number, string | number | boolean | null>>();

  constructor(
    private readonly delimiter: string,
    private readonly path: string,
    private readonly encoding: string
  ) {}

  /**
   * Индексирует кусок файла, не копируя его в память: окно строк читается
   * с диска заново, поэтому файл в сотни мегабайт не остаётся в памяти воркера.
   */
  append(chunk: Uint8Array): void {
    for (let index = 0; index < chunk.length; index++) {
      const byte = chunk[index];
      // 0x22 и 0x0A совпадают в UTF-8 и windows-1251 — сканируем байты.
      if (byte === 0x22) this.inQuotes = !this.inQuotes;
      else if (byte === 0x0a && !this.inQuotes) {
        const prev = index > 0 ? chunk[index - 1] : this.lastByte;
        if (prev === 0x0d) this.crlf = true;
        this.rowStarts.push(this.length + index + 1);
        this.rowCount++;
      }
    }
    this.length += chunk.length;
    if (chunk.length > 0) this.lastByte = chunk[chunk.length - 1];
  }

  /** Учитывает последнюю строку без завершающего перевода строки. */
  finish(): void {
    this.endsWithNewline = this.length === 0 || this.lastByte === 0x0a;
    if (this.length > 0 && !this.endsWithNewline) this.rowCount++;
  }

  get rows(): number {
    return this.rowCount;
  }

  get cols(): number {
    return Math.max(1, this.colCount);
  }

  get charset(): string {
    return this.encoding;
  }

  setCols(count: number): void {
    this.colCount = Math.max(1, count);
  }

  /** Правка ячейки: файл на диске до сохранения не трогаем. */
  edit(row: number, col: number, value: string | number | boolean | null): void {
    let rowEdits = this.edits.get(row);
    if (!rowEdits) {
      rowEdits = new Map();
      this.edits.set(row, rowEdits);
    }
    rowEdits.set(col, value);
  }

  /** Окно строк: нужный байтовый диапазон файла плюс накопленные правки. */
  async window(start: number, count: number): Promise<Row[]> {
    const from = Math.max(0, start);
    const to = Math.min(this.rowCount, from + Math.max(0, count));
    if (to <= from) return [];
    const begin = this.startOffset(from);
    const end = to >= this.rowCount ? this.length : this.startOffset(to);
    const rows = await this.parseRange(begin, end);
    for (let index = 0; index < rows.length; index += 1) {
      const rowEdits = this.edits.get(from + index);
      if (!rowEdits) continue;
      const row = rows[index];
      for (const [col, value] of rowEdits) {
        while (row.length <= col) row.push('');
        row[col] = value;
      }
    }
    return rows;
  }

  /**
   * CSV в байтах: пачка строк кодируется сразу и уходит в список кусков,
   * поэтому полного текста файла в памяти не живёт.
   */
  async toBytes(): Promise<Uint8Array> {
    const encoder = new TextEncoder();
    const eol = this.crlf ? '\r\n' : '\n';
    const parts: Uint8Array[] = [];
    // TextEncoder умеет только UTF-8, поэтому не-UTF-8 источник помечаем BOM:
    // без него повторное открытие файла как windows-1251 даёт кракозябры.
    if (!/^utf-?8/i.test(this.charset)) parts.push(new Uint8Array([0xef, 0xbb, 0xbf]));
    const batch = 4096;
    for (let start = 0; start < this.rowCount; start += batch) {
      const rows = await this.window(start, batch);
      if (rows.length === 0) break;
      const last = start + rows.length >= this.rowCount;
      const text =
        rows
          .map((row) => row.map((cell) => csvField(cell, this.delimiter)).join(this.delimiter))
          .join(eol) + (!last || this.endsWithNewline ? eol : '');
      parts.push(encoder.encode(text));
    }
    let total = 0;
    for (const part of parts) total += part.byteLength;
    const result = new Uint8Array(total);
    let at = 0;
    for (const part of parts) {
      result.set(part, at);
      at += part.byteLength;
    }
    return result;
  }

  private startOffset(row: number): number {
    if (row <= 0) return 0;
    return this.rowStarts[Math.min(row, this.rowStarts.length - 1)] ?? this.length;
  }

  private async parseRange(begin: number, end: number): Promise<Row[]> {
    if (end <= begin) return [];
    const bytes = await readBytes(this.path, begin, end - begin);
    const slice = new TextDecoder(this.encoding, { fatal: false }).decode(bytes);
    const rows: Row[] = [];
    let row: Row = [];
    let field = '';
    let inQuotes = false;

    const pushField = () => {
      const value = field;
      field = '';
      if (value === '') row.push('');
      else if (/^-?\d+(\.\d+)?$/.test(value)) row.push(Number(value));
      else row.push(value);
    };

    for (let index = 0; index < slice.length; index++) {
      const char = slice[index];
      if (inQuotes) {
        if (char === '"') {
          if (slice[index + 1] === '"') {
            field += '"';
            index++;
          } else {
            inQuotes = false;
          }
        } else {
          field += char;
        }
        continue;
      }
      if (char === '"') {
        inQuotes = true;
      } else if (char === this.delimiter) {
        pushField();
      } else if (char === '\r') {
        // игнорируем: границей строки служит \n
      } else if (char === '\n') {
        pushField();
        rows.push(row);
        row = [];
      } else {
        field += char;
      }
    }
    if (field.length > 0 || row.length > 0) {
      pushField();
      rows.push(row);
    }
    return rows;
  }
}

function detectDelimiter(sample: string): string {
  const candidates = [',', ';', '\t', '|'];
  let best = ',';
  let bestCount = -1;
  const lines = sample.split(/\r?\n/).slice(0, 20);
  for (const candidate of candidates) {
    let total = 0;
    for (const line of lines) {
      // Считаем разделители вне кавычек.
      let inQuotes = false;
      for (const char of line) {
        if (char === '"') inQuotes = !inQuotes;
        else if (char === candidate && !inQuotes) total++;
      }
    }
    if (total > bestCount) {
      bestCount = total;
      best = candidate;
    }
  }
  return best;
}

/** Поле CSV: кавычки только когда без них поле не прочитать обратно. */
function csvField(value: string | number | boolean | null, delimiter: string): string {
  if (value === null || value === undefined) return '';
  const text = typeof value === 'string' ? value : String(value);
  if (!text.includes(delimiter) && !/[\r\n"]/.test(text)) return text;
  return '"' + text.replace(/"/g, '""') + '"';
}

/* ============================================================
   XLSX: книга целиком в памяти воркера
   ============================================================ */

/** Пределы листа Excel: !ref может описывать их целиком, не имея ни одной ячейки. */
const MAX_SHEET_ROWS = 1_048_576;
const MAX_SHEET_COLS = 16_384;

/**
 * Ячейка листа: SheetJS с `dense: true` держит строки массивом массивов,
 * разреженный лист — по адресам вроде 'A1'.
 */
function cellAt(worksheet: XLSX.WorkSheet, r: number, c: number): XLSX.CellObject | undefined {
  if (Array.isArray(worksheet)) {
    const row = (worksheet as unknown as (XLSX.CellObject[] | undefined)[])[r];
    return row?.[c];
  }
  return worksheet[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
}

function setCellAt(
  worksheet: XLSX.WorkSheet,
  r: number,
  c: number,
  cell: XLSX.CellObject | undefined
): void {
  if (Array.isArray(worksheet)) {
    const rows = worksheet as unknown as (XLSX.CellObject[] | undefined)[];
    while (rows.length <= r) rows.push([]);
    const row = rows[r] ?? (rows[r] = []);
    if (cell) row[c] = cell;
    else delete row[c];
    return;
  }
  const address = XLSX.utils.encode_cell({ r, c });
  if (cell) (worksheet as Record<string, unknown>)[address] = cell;
  else delete (worksheet as Record<string, unknown>)[address];
}

/** Расширяет !ref: пустой или битый ref — не ошибка, лист начинается с этой ячейки. */
function touchRange(worksheet: XLSX.WorkSheet, r: number, c: number): void {
  const base = worksheet['!ref'];
  if (!base) {
    worksheet['!ref'] = XLSX.utils.encode_cell({ r, c });
    return;
  }
  try {
    const range = XLSX.utils.decode_range(base);
    if (r < range.s.r) range.s.r = r;
    if (c < range.s.c) range.s.c = c;
    if (r > range.e.r) range.e.r = r;
    if (c > range.e.c) range.e.c = c;
    worksheet['!ref'] = XLSX.utils.encode_range(range);
  } catch {
    worksheet['!ref'] = XLSX.utils.encode_cell({ r, c });
  }
}

/** Правка ячейки в самой книге: формула хранится в формате файла Excel. */
function writeCell(worksheet: XLSX.WorkSheet, edit: CellEdit): void {
  const value = edit.v;
  if (typeof value === 'string' && value.startsWith('=')) {
    // В файле аргументы через запятую; кэш пуст, иначе Excel показал бы текст формулы.
    setCellAt(worksheet, edit.r, edit.c, {
      t: 'n',
      f: toExcelFormula(value).slice(1),
    });
    return;
  }
  if (value === null || value === '') {
    setCellAt(worksheet, edit.r, edit.c, undefined);
    return;
  }
  if (typeof value === 'number') {
    setCellAt(worksheet, edit.r, edit.c, { t: 'n', v: value });
    return;
  }
  if (typeof value === 'boolean') {
    setCellAt(worksheet, edit.r, edit.c, { t: 'b', v: value });
    return;
  }
  setCellAt(worksheet, edit.r, edit.c, { t: 's', v: String(value) });
}

class XlsxTable {
  private workbook: XLSX.WorkBook;
  private order: string[] = [];

  constructor(workbook: XLSX.WorkBook) {
    this.workbook = workbook;
    this.order = workbook.SheetNames.slice();
  }

  get names(): string[] {
    return this.order;
  }

  private worksheet(sheet: string): XLSX.WorkSheet | undefined {
    return this.workbook.Sheets[sheet];
  }

  /**
   * Размеры листа по фактическим данным: после форматирования целого столбца
   * Excel пишет !ref="A1:XFD1048576", и окно 4000 строк разворачивалось в 65 млн ячеек.
   */
  private dims(sheet: string): { rows: number; cols: number } {
    const worksheet = this.worksheet(sheet);
    if (!worksheet) return { rows: 0, cols: 1 };
    const limit = this.refLimit(worksheet);
    let rows = 0;
    let cols = 1;
    if (Array.isArray(worksheet)) {
      const dense = worksheet as unknown as (XLSX.CellObject[] | undefined)[];
      for (let r = 0; r < dense.length; r += 1) {
        const row = dense[r];
        if (!row || row.length === 0) continue;
        // Хвост из undefined длина массива тоже считает, а данных в нём нет.
        let last = row.length - 1;
        while (last >= 0 && !row[last]) last -= 1;
        if (last < 0) continue;
        if (r + 1 > rows) rows = r + 1;
        if (last + 1 > cols) cols = last + 1;
      }
    } else {
      const sparse = worksheet as Record<string, unknown>;
      for (const key of Object.keys(sparse)) {
        if (!key || key.charCodeAt(0) === 0x21 /* ! */) continue;
        try {
          const cell = XLSX.utils.decode_cell(key);
          if (!Number.isFinite(cell.r) || !Number.isFinite(cell.c)) continue;
          if (cell.r + 1 > rows) rows = cell.r + 1;
          if (cell.c + 1 > cols) cols = cell.c + 1;
        } catch {
          // Ключ не адрес ячейки — на размеры листа не влияет.
        }
      }
    }
    return {
      rows: Math.min(rows, limit.rows, MAX_SHEET_ROWS),
      cols: Math.min(Math.max(1, cols), limit.cols, MAX_SHEET_COLS),
    };
  }

  /** Верхняя граница из !ref; ref на весь лист о данных не говорит, поэтому отбрасывается. */
  private refLimit(worksheet: XLSX.WorkSheet): { rows: number; cols: number } {
    const full = { rows: MAX_SHEET_ROWS, cols: MAX_SHEET_COLS };
    const ref = worksheet['!ref'];
    if (!ref) return full;
    try {
      const range = XLSX.utils.decode_range(ref);
      if (
        range.s.r === 0 &&
        range.s.c === 0 &&
        range.e.r >= MAX_SHEET_ROWS - 1 &&
        range.e.c >= MAX_SHEET_COLS - 1
      ) {
        return full;
      }
      return { rows: Math.max(1, range.e.r + 1), cols: Math.max(1, range.e.c + 1) };
    } catch {
      return full;
    }
  }

  meta(): SheetMeta[] {
    return this.order.map((name) => {
      const { rows, cols } = this.dims(name);
      return { name, rows, cols };
    });
  }

  window(sheet: string, start: number, count: number): Row[] {
    const worksheet = this.worksheet(sheet);
    const from = Math.max(0, start);
    const to = from + Math.max(0, count);
    if (!worksheet || to <= from) return [];
    const { rows, cols } = this.dims(sheet);
    const result: Row[] = [];
    for (let r = from; r < Math.min(to, rows); r += 1) {
      const row: Row = [];
      for (let c = 0; c < cols; c += 1) row.push(cellValue(cellAt(worksheet, r, c)));
      result.push(row);
    }
    return result;
  }

  /** Правки применяем к исходной книге: она же уходит в файл при сохранении. */
  setCells(sheet: string, edits: CellEdit[]): void {
    let worksheet = this.worksheet(sheet);
    if (!worksheet) {
      worksheet = {};
      this.workbook.Sheets[sheet] = worksheet;
    }
    for (const edit of edits) {
      if (!Number.isFinite(edit.r) || !Number.isFinite(edit.c)) continue;
      if (edit.r < 0 || edit.c < 0) continue;
      writeCell(worksheet, edit);
      touchRange(worksheet, edit.r, edit.c);
    }
  }

  /**
   * Сохранение: пишем ту же книгу, а не собранную заново из матрицы.
   * Пересборка через book_new/aoa_to_sheet теряла !merges, ширины и высоты
   * ячеек, число-форматы, примечания и закрепление.
   *
   * ВАЖНО: CE не переносит при записи оформление ячеек и диаграммы (стили
   * пишет injectStylesIntoXlsx), поэтому редактор сохраняет через
   * flushWindowed → syncWorkbook, а не этим путём.
   */
  toBytes(ext: string, vbaraw?: unknown): ArrayBuffer {
    const bookVBA = ext === 'xlsm';
    if (vbaraw) (this.workbook as { vbaraw?: unknown }).vbaraw = vbaraw;
    return XLSX.write(this.workbook, {
      bookType: bookVBA ? 'xlsm' : 'xlsx',
      type: 'array',
      bookVBA,
    }) as ArrayBuffer;
  }
}

function cellValue(cell: XLSX.CellObject | undefined): string | number | boolean | null {
  if (!cell) return null;
  if (typeof cell.f === 'string' && cell.f) return '=' + fromExcelFormula(cell.f);
  if (cell.v === undefined || cell.v === null) return null;
  if (typeof cell.v === 'number' || typeof cell.v === 'boolean' || typeof cell.v === 'string') {
    return cell.v;
  }
  return String(cell.v);
}

/* ============================================================
   Состояние воркера
   ============================================================ */

let csv: CsvTable | null = null;
let csvName = '';
let xlsx: XlsxTable | null = null;
let vbaRaw: unknown = null;
/** Расширение открытой книги: от него зависит формат записи макросов. */
let sheetExt: 'xlsx' | 'xlsm' = 'xlsx';
/** Поколение открытия: устаревшее открытие не коммитит состояние после await. */
let openGeneration = 0;

const scope = self as unknown as {
  postMessage: (message: WorkerResponse, transfer?: Transferable[]) => void;
  onmessage: ((event: MessageEvent<WorkerRequest | { id: number; kind: 'readResult'; bytes: ArrayBuffer; eof: boolean }>) => void) | null;
};

function reply(id: number, result?: unknown, transfer: Transferable[] = []): void {
  // Без transfer большие байты копируются структурным клонированием.
  scope.postMessage({ id, ok: true, result }, transfer);
}

function fail(id: number, error: unknown): void {
  scope.postMessage({ id, ok: false, error: error instanceof Error ? error.message : String(error) });
}

/** Запрос куска файла у главного потока. */
let readSequence = 1_000_000;
/** Ответ может не прийти вовсе — тогда окно иначе висело бы в «Загрузке» вечно. */
const READ_TIMEOUT_MS = 30_000;
interface ReadResult {
  bytes: ArrayBuffer;
  eof: boolean;
  total: number;
}
interface PendingRead {
  resolve: (value: ReadResult) => void;
  reject: (error: Error) => void;
  timer: ReturnType<typeof setTimeout>;
}
const pendingReads = new Map<number, PendingRead>();

function requestRead(path: string, offset: number, length: number): Promise<ReadResult> {
  const id = readSequence++;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingReads.delete(id);
      reject(new Error(`Главный поток не ответил на чтение ${length} байт (${path})`));
    }, READ_TIMEOUT_MS);
    pendingReads.set(id, { resolve, reject, timer });
    scope.postMessage({ id, kind: 'read', path, offset, length });
  });
}

/** Отклоняет все ожидающие чтения: после close ответа уже не будет. */
function rejectPendingReads(error: Error): void {
  for (const [, entry] of pendingReads) {
    clearTimeout(entry.timer);
    entry.reject(error);
  }
  pendingReads.clear();
}

/**
 * Читает ровно запрошенный диапазон файла. Оконный CSV держит только
 * смещения строк, поэтому каждое окно перечитывается с диска кусками.
 */
async function readBytes(path: string, offset: number, length: number): Promise<Uint8Array> {
  if (length <= 0) return new Uint8Array(0);
  const parts: Uint8Array[] = [];
  let done = 0;
  while (done < length) {
    const chunk = await requestRead(path, offset + done, length - done);
    const bytes = new Uint8Array(chunk.bytes);
    // Пустой ответ — потеря данных, а не конец файла: молча обрезать нельзя.
    if (bytes.byteLength === 0) {
      throw new Error(`Файл не прочитан: получено ${done} из ${length} байт (${path})`);
    }
    parts.push(bytes);
    done += bytes.byteLength;
    if (chunk.eof) break;
  }
  if (done < length) {
    throw new Error(`Файл не прочитан: получено ${done} из ${length} байт (${path})`);
  }
  if (parts.length === 1) return parts[0];
  const result = new Uint8Array(done);
  let at = 0;
  for (const part of parts) {
    result.set(part, at);
    at += part.byteLength;
  }
  return result;
}

async function openCsv(request: Extract<WorkerRequest, { kind: 'openCsv' }>): Promise<SheetMeta[]> {
  if (csv || xlsx) throw new Error('Таблица уже открыта: сначала закройте текущую');
  // Поколение до первого await: параллельный openXlsx/close отменяет коммит ниже.
  const generation = ++openGeneration;
  const chunkSize = request.chunkSize ?? 8 * 1024 * 1024;
  const encoding = request.encoding === 'utf-8-bom' ? 'utf-8' : request.encoding;
  // Первый кусок нужен, чтобы определить разделитель по живым данным.
  const head = await requestRead(request.path, 0, Math.min(chunkSize, 256 * 1024));
  const headBytes = new Uint8Array(head.bytes);
  const headText = new TextDecoder(encoding, { fatal: false }).decode(headBytes);
  const table = new CsvTable(detectDelimiter(headText), request.path, encoding);
  table.append(headBytes);

  let offset = headBytes.byteLength;
  const total = head.total || offset;
  let eof = head.eof;
  while (!eof) {
    const chunk = await requestRead(request.path, offset, chunkSize);
    const bytes = new Uint8Array(chunk.bytes);
    if (bytes.byteLength === 0) break;
    table.append(bytes);
    offset += bytes.byteLength;
    eof = chunk.eof;
    scope.postMessage({ id: 0, kind: 'progress', phase: 'read', done: offset, total });
  }
  table.finish();
  // Ширина листа — по первой строке, без материализации всего файла.
  const [firstRow] = await table.window(0, 1);
  table.setCols(firstRow?.length ?? 1);
  if (generation !== openGeneration) throw new Error('Открытие таблицы отменено другим запросом');
  // Книга Excel и CSV одновременно не живут: иначе запросы уходят не туда.
  csv = table;
  csvName = request.name || 'Лист 1';
  return [{ name: csvName, rows: table.rows, cols: table.cols }];
}

function openXlsx(request: Extract<WorkerRequest, { kind: 'openXlsx' }>): SheetMeta[] {
  if (csv || xlsx) throw new Error('Таблица уже открыта: сначала закройте текущую');
  // Отменяет незавершённый openCsv: его таблица уже не актуальна.
  openGeneration += 1;
  const workbook = XLSX.read(new Uint8Array(request.bytes), {
    type: 'array',
    dense: true,
    // Стили не читаем: CE не пишет их обратно, а ячейки со стилем только раздувают память.
    cellStyles: false,
    cellNF: false,
    cellText: false,
    sheetStubs: false,
    bookVBA: request.ext === 'xlsm',
  });
  vbaRaw = (workbook as { vbaraw?: unknown }).vbaraw ?? null;
  sheetExt = request.ext === 'xlsm' ? 'xlsm' : 'xlsx';
  xlsx = new XlsxTable(workbook);
  return xlsx.meta();
}

scope.onmessage = (event) => {
  const data = event.data as
    | WorkerRequest
    | { id: number; kind: 'readResult'; bytes: ArrayBuffer; eof: boolean; total: number };

  if (data.kind === 'readResult') {
    const entry = pendingReads.get(data.id);
    if (entry) {
      pendingReads.delete(data.id);
      clearTimeout(entry.timer);
      entry.resolve({ bytes: data.bytes, eof: data.eof, total: data.total });
    }
    return;
  }

  const request = data as WorkerRequest;
  try {
    switch (request.kind) {
      case 'openCsv':
        void openCsv(request).then((meta) => reply(request.id, meta)).catch((err) => fail(request.id, err));
        break;
      case 'openXlsx':
        reply(request.id, openXlsx(request));
        break;
      case 'meta':
        if (xlsx) reply(request.id, xlsx.meta());
        else if (csv) reply(request.id, [{ name: csvName, rows: csv.rows, cols: csv.cols }]);
        else reply(request.id, []);
        break;
      case 'getWindow':
        if (xlsx) reply(request.id, xlsx.window(request.sheet, request.start, request.count));
        else if (csv) {
          const table = csv;
          void table
            .window(request.start, request.count)
            .then((rows) => reply(request.id, rows))
            .catch((err) => fail(request.id, err));
        } else fail(request.id, 'Таблица не открыта');
        break;
      case 'setCells':
        if (xlsx) {
          xlsx.setCells(request.sheet, request.edits);
          reply(request.id);
        } else if (csv) {
          // Файл на диске до сохранения не трогаем: правки живут в индексе.
          for (const edit of request.edits) csv.edit(edit.r, edit.c, edit.v);
          reply(request.id);
        } else fail(request.id, 'Таблица не открыта');
        break;
      case 'save': {
        if (xlsx) {
          const bytes = xlsx.toBytes(sheetExt, vbaRaw);
          reply(request.id, bytes, [bytes]);
          break;
        }
        if (csv) {
          const table = csv;
          void table
            .toBytes()
            .then((bytes) => {
              const buffer = bytes.buffer as ArrayBuffer;
              reply(request.id, buffer, [buffer]);
            })
            .catch((err) => fail(request.id, err));
          break;
        }
        fail(request.id, 'Таблица не открыта');
        break;
      }
      case 'close':
        // Незавершённый openCsv не должен вернуть состояние после закрытия.
        openGeneration += 1;
        csv = null;
        xlsx = null;
        vbaRaw = null;
        rejectPendingReads(new Error('Таблица закрыта'));
        reply(request.id);
        break;
      default: {
        const unknownRequest = request as unknown as { id: number; kind: string };
        fail(unknownRequest.id, `Неизвестная команда: ${unknownRequest.kind}`);
      }
    }
  } catch (err) {
    fail(request.id, err);
  }
};
