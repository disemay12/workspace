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
import { fromExcelFormula } from '../formulaEngine';

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
        this.rowStarts.push(this.length + index + 1);
        this.rowCount++;
      }
    }
    this.length += chunk.length;
    if (chunk.length > 0) this.lastByte = chunk[chunk.length - 1];
  }

  /** Учитывает последнюю строку без завершающего перевода строки. */
  finish(): void {
    if (this.length > 0 && this.lastByte !== 0x0a) this.rowCount++;
  }

  get rows(): number {
    return this.rowCount;
  }

  get cols(): number {
    return Math.max(1, this.colCount);
  }

  setCols(count: number): void {
    this.colCount = Math.max(1, count);
  }

  /** Окно строк: нужный байтовый диапазон файла. */
  async window(start: number, count: number): Promise<Row[]> {
    const from = Math.max(0, start);
    const to = Math.min(this.rowCount, from + Math.max(0, count));
    if (to <= from) return [];
    const begin = this.startOffset(from);
    const end = to >= this.rowCount ? this.length : this.startOffset(to);
    return this.parseRange(begin, end);
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

/* ============================================================
   XLSX: книга целиком в памяти воркера
   ============================================================ */

/** Пределы листа Excel: !ref может описывать их целиком, не имея ни одной ячейки. */
const MAX_SHEET_ROWS = 1_048_576;
const MAX_SHEET_COLS = 16_384;

type DenseRows = (XLSX.CellObject[] | undefined)[];

/**
 * Раскладка листа: SheetJS с `dense: true` держит строки массивом массивов —
 * в самом листе или в `!data` (версии библиотеки различаются), разреженный
 * лист — по адресам вроде 'A1'. Всё прочее читать не умеем и молчать нельзя:
 * пустой лист выглядел как потеря данных.
 */
function sheetLayout(
  worksheet: XLSX.WorkSheet
): { kind: 'dense'; rows: DenseRows } | { kind: 'sparse'; cells: Record<string, unknown> } {
  if (Array.isArray(worksheet)) {
    return { kind: 'dense', rows: worksheet as unknown as DenseRows };
  }
  const data = (worksheet as Record<string, unknown>)['!data'];
  if (Array.isArray(data)) return { kind: 'dense', rows: data as unknown as DenseRows };
  if (typeof worksheet === 'object' && worksheet !== null) {
    return { kind: 'sparse', cells: worksheet as Record<string, unknown> };
  }
  throw new Error('Неизвестный формат листа XLSX');
}

function cellAt(worksheet: XLSX.WorkSheet, r: number, c: number): XLSX.CellObject | undefined {
  const layout = sheetLayout(worksheet);
  if (layout.kind === 'dense') return layout.rows[r]?.[c];
  return layout.cells[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
}

class XlsxTable {
  private workbook: XLSX.WorkBook;
  private order: string[] = [];
  /** Размеры листов: считается один раз на книгу, сбрасывается только загрузкой новой. */
  private readonly dimCache = new Map<string, { rows: number; cols: number }>();

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
   *
   * Перебор всех ячеек кэшируется: прокрутка листа спрашивает окна подряд, и
   * полный скан на каждый запрос давал O(окно × ячейки). Состав ячеек после
   * загрузки книги не меняется, поэтому кэш живёт вместе с книгой.
   */
  private dims(sheet: string): { rows: number; cols: number } {
    const cached = this.dimCache.get(sheet);
    if (cached) return cached;
    const worksheet = this.worksheet(sheet);
    if (!worksheet) return { rows: 0, cols: 1 };
    const limit = this.refLimit(worksheet);
    let rows = 0;
    let cols = 1;
    const layout = sheetLayout(worksheet);
    if (layout.kind === 'dense') {
      const dense = layout.rows;
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
      for (const key of Object.keys(layout.cells)) {
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
    const result = {
      rows: Math.min(rows, limit.rows, MAX_SHEET_ROWS),
      cols: Math.min(Math.max(1, cols), limit.cols, MAX_SHEET_COLS),
    };
    this.dimCache.set(sheet, result);
    return result;
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
  });
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
      case 'close':
        // Незавершённый openCsv не должен вернуть состояние после закрытия.
        openGeneration += 1;
        csv = null;
        xlsx = null;
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
