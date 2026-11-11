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
  private text = '';
  private rowStarts: number[] = [];
  private rowCount = 0;
  private colCount = 0;
  /** Состояние «внутри кавычек» переходит между кусками файла. */
  private inQuotes = false;

  constructor(private readonly delimiter: string) {}

  /** Добавляет расшифрованный кусок текста и индексирует завершённые строки. */
  append(chunk: string): void {
    const base = this.text.length;
    this.text += chunk;
    for (let index = 0; index < chunk.length; index++) {
      const char = chunk[index];
      if (char === '"') this.inQuotes = !this.inQuotes;
      else if (char === '\n' && !this.inQuotes) {
        this.rowStarts.push(base + index + 1);
        this.rowCount++;
      }
    }
  }

  /** Учитывает последнюю строку без завершающего перевода строки. */
  finish(): void {
    if (this.text.length > 0 && !this.text.endsWith('\n')) this.rowCount++;
  }

  get rows(): number {
    return this.rowCount;
  }

  get cols(): number {
    if (this.colCount > 0) return this.colCount;
    // оценка по первой строке
    const first = this.parseRange(0, 1)[0] ?? [];
    this.colCount = Math.max(1, first.length);
    return this.colCount;
  }

  private rowStart(row: number): number {
    if (row <= 0) return 0;
    return this.rowStarts[Math.min(row - 1, this.rowStarts.length - 1)] ?? this.text.length;
  }

  /** Разбор диапазона строк: без создания промежуточного массива всей таблицы. */
  window(start: number, count: number): Row[] {
    const from = Math.max(0, start);
    const to = Math.min(this.rowCount, from + Math.max(0, count));
    return this.parseRange(from, to - from);
  }

  private parseRange(start: number, count: number): Row[] {
    const begin = this.rowStart(start);
    const end = start + count >= this.rowCount ? this.text.length : this.rowStart(start + count);
    if (end <= begin) return [];
    const slice = this.text.slice(begin, end);
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

interface XlsxSheetData {
  rows: Row[];
  built: boolean;
}

class XlsxTable {
  private workbook: XLSX.WorkBook;
  private sheets = new Map<string, XlsxSheetData>();
  private order: string[] = [];

  constructor(workbook: XLSX.WorkBook) {
    this.workbook = workbook;
    this.order = workbook.SheetNames.slice();
  }

  get names(): string[] {
    return this.order;
  }

  private ensure(sheet: string): XlsxSheetData {
    const cached = this.sheets.get(sheet);
    if (cached?.built) return cached;
    const worksheet = this.workbook.Sheets[sheet];
    const data: XlsxSheetData = { rows: [], built: true };
    if (worksheet) {
      const dense = (worksheet as { '!data'?: XLSX.CellObject[][] })['!data'];
      if (dense && Array.isArray(dense)) {
        for (let r = 0; r < dense.length; r++) {
          const sourceRow = dense[r] ?? [];
          const row: Row = [];
          for (let c = 0; c < sourceRow.length; c++) {
            row.push(cellValue(sourceRow[c]));
          }
          data.rows.push(row);
        }
      } else {
        // Разреженный лист: собираем диапазон из !ref.
        const ref = worksheet['!ref'] ?? 'A1';
        const range = XLSX.utils.decode_range(ref);
        for (let r = range.s.r; r <= range.e.r; r++) {
          const row: Row = [];
          for (let c = range.s.c; c <= range.e.c; c++) {
            row.push(cellValue(worksheet[XLSX.utils.encode_cell({ r, c })]));
          }
          data.rows.push(row);
        }
      }
    }
    this.sheets.set(sheet, data);
    return data;
  }

  meta(): SheetMeta[] {
    return this.order.map((name) => {
      const data = this.ensure(name);
      const cols = data.rows.reduce((max, row) => Math.max(max, row.length), 1);
      return { name, rows: data.rows.length, cols };
    });
  }

  window(sheet: string, start: number, count: number): Row[] {
    const data = this.ensure(sheet);
    return data.rows.slice(Math.max(0, start), Math.max(0, start) + Math.max(0, count));
  }

  setCells(sheet: string, edits: CellEdit[]): void {
    const data = this.ensure(sheet);
    for (const edit of edits) {
      while (data.rows.length <= edit.r) data.rows.push([]);
      const row = data.rows[edit.r];
      while (row.length <= edit.c) row.push('');
      row[edit.c] = edit.v;
    }
  }

  toBytes(ext: string, vbaraw?: unknown): ArrayBuffer {
    const workbook = XLSX.utils.book_new();
    for (const name of this.order) {
      const data = this.ensure(name);
      const aoa = data.rows.map((row) => row.map((cell) => (cell === null ? '' : cell)));
      const sheet = XLSX.utils.aoa_to_sheet(aoa.length > 0 ? aoa : [[]]);
      // Формулы: значение начинается с «=» — пишем как формулу с кэшем.
      for (let r = 0; r < aoa.length; r++) {
        for (let c = 0; c < aoa[r].length; c++) {
          const value = aoa[r][c];
          if (typeof value !== 'string' || !value.startsWith('=')) continue;
          const address = XLSX.utils.encode_cell({ r, c });
          const existing = sheet[address];
          (sheet as Record<string, unknown>)[address] = {
            t: 's',
            f: toExcelFormula(value).slice(1),
            v: existing?.v ?? value,
          };
        }
      }
      XLSX.utils.book_append_sheet(workbook, sheet, name);
    }
    if (ext === 'xlsm' && vbaraw) {
      (workbook as { vbaraw?: unknown }).vbaraw = vbaraw;
    }
    const output = XLSX.write(workbook, {
      bookType: ext === 'xlsm' ? 'xlsm' : 'xlsx',
      type: 'array',
      bookVBA: ext === 'xlsm',
    }) as ArrayBuffer;
    return output;
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

const scope = self as unknown as {
  postMessage: (message: WorkerResponse) => void;
  onmessage: ((event: MessageEvent<WorkerRequest | { id: number; kind: 'readResult'; bytes: ArrayBuffer; eof: boolean }>) => void) | null;
};

function reply(id: number, result?: unknown): void {
  scope.postMessage({ id, ok: true, result });
}

function fail(id: number, error: unknown): void {
  scope.postMessage({ id, ok: false, error: error instanceof Error ? error.message : String(error) });
}

/** Запрос куска файла у главного потока. */
let readSequence = 1_000_000;
interface ReadResult {
  bytes: ArrayBuffer;
  eof: boolean;
  total: number;
}
const pendingReads = new Map<number, (value: ReadResult) => void>();

function requestRead(path: string, offset: number, length: number): Promise<ReadResult> {
  const id = readSequence++;
  return new Promise((resolve) => {
    pendingReads.set(id, resolve);
    scope.postMessage({ id, kind: 'read', path, offset, length });
  });
}

async function openCsv(request: Extract<WorkerRequest, { kind: 'openCsv' }>): Promise<SheetMeta[]> {
  const chunkSize = request.chunkSize ?? 8 * 1024 * 1024;
  const decoder = new TextDecoder(
    request.encoding === 'windows-1251' ? 'windows-1251' : request.encoding,
    { fatal: false }
  );
  // Первый кусок нужен, чтобы определить разделитель по живым данным.
  let chunk = await requestRead(request.path, 0, Math.min(chunkSize, 256 * 1024));
  const headText = decoder.decode(chunk.bytes, { stream: true });
  const table = new CsvTable(detectDelimiter(headText));
  table.append(headText);

  let offset = chunk.bytes.byteLength;
  const total = chunk.total || offset;
  while (!chunk.eof) {
    chunk = await requestRead(request.path, offset, chunkSize);
    if (chunk.bytes.byteLength === 0) break;
    table.append(decoder.decode(chunk.bytes, { stream: true }));
    offset += chunk.bytes.byteLength;
    scope.postMessage({ id: 0, kind: 'progress', phase: 'read', done: offset, total });
  }
  table.append(decoder.decode());
  table.finish();
  csv = table;
  csvName = request.name || 'Лист 1';
  return [{ name: csvName, rows: table.rows, cols: table.cols }];
}

function openXlsx(request: Extract<WorkerRequest, { kind: 'openXlsx' }>): SheetMeta[] {
  const workbook = XLSX.read(new Uint8Array(request.bytes), {
    type: 'array',
    dense: true,
    cellStyles: false,
    cellNF: false,
    cellText: false,
    sheetStubs: false,
    bookVBA: request.ext === 'xlsm',
  });
  vbaRaw = (workbook as { vbaraw?: unknown }).vbaraw ?? null;
  xlsx = new XlsxTable(workbook);
  return xlsx.meta();
}

scope.onmessage = (event) => {
  const data = event.data as
    | WorkerRequest
    | { id: number; kind: 'readResult'; bytes: ArrayBuffer; eof: boolean; total: number };

  if (data.kind === 'readResult') {
    const resolver = pendingReads.get(data.id);
    if (resolver) {
      pendingReads.delete(data.id);
      resolver({ bytes: data.bytes, eof: data.eof, total: data.total });
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
        else if (csv) reply(request.id, csv.window(request.start, request.count));
        else fail(request.id, 'Таблица не открыта');
        break;
      case 'setCells':
        if (xlsx) {
          xlsx.setCells(request.sheet, request.edits);
          reply(request.id);
        } else if (csv) {
          fail(request.id, 'Правка CSV большого размера пока не поддерживается');
        } else fail(request.id, 'Таблица не открыта');
        break;
      case 'save': {
        if (!xlsx) {
          fail(request.id, 'Сохранение доступно только для книги Excel');
          break;
        }
        const bytes = xlsx.toBytes('xlsx', vbaRaw);
        reply(request.id, bytes);
        break;
      }
      case 'close':
        csv = null;
        xlsx = null;
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
