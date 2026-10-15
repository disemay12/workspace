/**
 * Движок формул электронных таблиц.
 *
 * Вычисляет формулы Excel/Google Sheets над сеткой листов: арифметику,
 * ссылки на ячейки и диапазоны, ссылки на другие листы, встроенные функции
 * (математические, статистические, логические, текстовые, даты, поиск)
 * и корректные значения ошибок (#DIV/0!, #N/A, #VALUE! и т.д.).
 *
 * Устройство: лексер -> парсер (рекурсивный спуск) -> вычислитель по AST.
 * Кэш значений на уровне книги с разрешением циклических ссылок.
 */

/* ============================================================
   Значения и ошибки
   ============================================================ */

export type CellPrimitive = string | number | boolean | null;

/** Ошибка вычисления в виде значения ячейки (#DIV/0! и т.п.). */
export interface FormulaError {
  __error: string;
}

export const ERRORS = {
  DIV0: '#DIV/0!',
  NA: '#N/A',
  NAME: '#NAME?',
  NULL: '#NULL!',
  NUM: '#NUM!',
  REF: '#REF!',
  VALUE: '#VALUE!',
  CIRCULAR: '#CIRC!',
} as const;

export function isError(v: unknown): v is FormulaError {
  return typeof v === 'object' && v !== null && typeof (v as FormulaError).__error === 'string';
}

const err = (code: string): FormulaError => ({ __error: code });

/** Сетка листа: строки x столбцы. Ячейка может содержать строку с формулой. */
export type Grid = (string | number | boolean | null)[][];

/** Книга: имя листа -> сетка. */
export type Workbook = { [sheetName: string]: Grid };

/* ============================================================
   Адресация
   ============================================================ */

/** Индекс столбца (0 → A, 26 → AA). */
export function columnName(index: number): string {
  let name = '';
  let temp = Math.max(0, Math.floor(index));
  while (temp >= 0) {
    name = String.fromCharCode((temp % 26) + 65) + name;
    temp = Math.floor(temp / 26) - 1;
  }
  return name;
}

/** Буквенное обозначение столбца в индекс (A → 0, AA → 26). */
export function columnIndex(letters: string): number {
  let idx = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    idx = idx * 26 + (code - 64);
  }
  return idx - 1;
}

/** Адрес ячейки: "A1", "$B$3", "Лист2!C5". */
export interface CellAddress {
  sheet?: string;
  col: number;
  row: number;
  /** Абсолютные ссылки сохраняются для корректного копирования. */
  absCol?: boolean;
  absRow?: boolean;
}

/** Разбор ссылки на ячейку. Возвращает null, если строка не является адресом. */
export function parseCellRef(ref: string, defaultSheet?: string): CellAddress | null {
  const m = ref.match(/^(?:'([^']+)'|([A-Za-zА-Яа-яЁё0-9_ ]+))!\s*(\$?)([A-Za-z]+)(\$?)(\d+)$/);
  if (m) {
    return {
      sheet: m[1] ?? m[2],
      absCol: m[3] === '$',
      col: columnIndex(m[4]),
      absRow: m[5] === '$',
      row: parseInt(m[6], 10) - 1,
    };
  }
  const simple = ref.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
  if (simple) {
    return {
      sheet: defaultSheet,
      absCol: simple[1] === '$',
      col: columnIndex(simple[2]),
      absRow: simple[3] === '$',
      row: parseInt(simple[4], 10) - 1,
    };
  }
  return null;
}

export function formatCellRef(addr: CellAddress): string {
  return `${columnName(addr.col)}${addr.row + 1}`;
}

/* ============================================================
   Лексер
   ============================================================ */

type TokenType =
  | 'number'
  | 'string'
  | 'ref'
  | 'range'
  | 'ident'
  | 'op'
  | 'lparen'
  | 'rparen'
  | 'comma'
  | 'eof';

interface Token {
  type: TokenType;
  value: string;
  /** Для range: начало и конец. */
  range?: { start: CellAddress; end: CellAddress };
}

/** Простое имя: имя функции, TRUE/FALSE или имя листа без кавычек. */
const IDENT_RE = /^[A-Za-zА-Яа-яЁё_\\][A-Za-zА-Яа-яЁё0-9_.\\]*/;
/** Экранированное имя листа в апострофах: 'Мой лист'. */
const QUOTED_SHEET_RE = /^'(?:[^']|'')+'/;
/** Ссылка на ячейку без имени листа: A1, $B$3. */
const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})(?![\w(])/;
/** Число: 1, 1.5, .5, 1e3. */
const NUMBER_RE = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/;

/** Ссылка на ячейку вместе с необязательным именем листа. */
function matchCellWithSheet(
  input: string
): { text: string; addr: CellAddress } | null {
  let sheet: string | undefined;
  let offset = 0;

  const quoted = input.match(QUOTED_SHEET_RE);
  if (quoted) {
    const sheetName = quoted[0].slice(1, -1);
    // После имени листа обязателен '!'
    if (input[quoted[0].length] !== '!') return null;
    sheet = sheetName.replace(/''/g, "'");
    offset = quoted[0].length + 1;
  } else {
    const ident = input.match(IDENT_RE);
    if (ident && input[ident[0].length] === '!') {
      sheet = ident[0];
      offset = ident[0].length + 1;
    }
  }

  const cell = input.slice(offset).match(CELL_RE);
  if (!cell) return null;

  const addr: CellAddress = {
    sheet,
    absCol: input[offset] === '$',
    col: columnIndex(cell[1]),
    absRow: input[offset + (input[offset] === '$' ? 1 + cell[1].length : cell[1].length)] === '$',
    row: parseInt(cell[2], 10) - 1,
  };
  if (addr.col < 0 || addr.row < 0) return null;

  return { text: input.slice(0, offset + cell[0].length), addr };
}

export function tokenize(input: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;

  while (i < input.length) {
    const ch = input[i];
    const rest = input.slice(i);

    // Пробелы
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // Строковый литерал в двойных кавычках ("abc""def" — экранирование)
    if (ch === '"') {
      let value = '';
      i++;
      while (i < input.length) {
        if (input[i] === '"') {
          if (input[i + 1] === '"') {
            value += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        value += input[i++];
      }
      tokens.push({ type: 'string', value });
      continue;
    }

    // Число — только если это не хвост ссылки (B2 разбирается как ссылка ниже,
    // потому что буква идёт первой и перехватывается веткой ссылок).
    if (/\d/.test(ch) || (ch === '.' && /\d/.test(input[i + 1] ?? ''))) {
      const numMatch = rest.match(NUMBER_RE);
      if (numMatch) {
        tokens.push({ type: 'number', value: numMatch[0] });
        i += numMatch[0].length;
        continue;
      }
    }

    // Ссылка на ячейку (с именем листа или без) и, возможно, диапазон A1:B2
    const cellMatch = matchCellWithSheet(rest);
    if (cellMatch) {
      const after = rest.slice(cellMatch.text.length);
      const rangeSep = after.match(/^\s*:\s*/);
      if (rangeSep) {
        // Конец диапазона: с именем листа или без (наследует лист начала)
        const endMatch = matchCellWithSheet(after.slice(rangeSep[0].length));
        if (endMatch) {
          const endAddr = endMatch.addr;
          if (!endAddr.sheet) endAddr.sheet = cellMatch.addr.sheet;
          const consumed = cellMatch.text.length + rangeSep[0].length + endMatch.text.length;
          tokens.push({
            type: 'range',
            value: input.slice(i, i + consumed),
            range: { start: cellMatch.addr, end: endAddr },
          });
          i += consumed;
          continue;
        }
      }
      tokens.push({ type: 'ref', value: cellMatch.text });
      i += cellMatch.text.length;
      continue;
    }

    // Идентификатор: имя функции, TRUE/FALSE или неизвестное имя
    const identMatch = rest.match(IDENT_RE) ?? rest.match(QUOTED_SHEET_RE);
    if (identMatch) {
      tokens.push({ type: 'ident', value: identMatch[0].replace(/^'|'$/g, '') });
      i += identMatch[0].length;
      continue;
    }

    // Операторы
    if (ch === '(') {
      tokens.push({ type: 'lparen', value: ch });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: 'rparen', value: ch });
      i++;
      continue;
    }
    if (ch === ',' || ch === ';') {
      // В русской локали разделитель аргументов — ';'
      tokens.push({ type: 'comma', value: ',' });
      i++;
      continue;
    }
    if (ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '^' || ch === '&') {
      tokens.push({ type: 'op', value: ch });
      i++;
      continue;
    }
    if (ch === '=' || ch === '<' || ch === '>') {
      let op = ch;
      if ((ch === '<' || ch === '>') && input[i + 1] === '=') {
        op = `${ch}=`;
        i++;
      } else if (ch === '<' && input[i + 1] === '>') {
        op = '<>';
        i++;
      }
      tokens.push({ type: 'op', value: op });
      i++;
      continue;
    }
    if (ch === '%') {
      tokens.push({ type: 'op', value: '%' });
      i++;
      continue;
    }

    // Неизвестный символ — считаем частью имени, чтобы не сломать разбор
    tokens.push({ type: 'ident', value: ch });
    i++;
  }

  tokens.push({ type: 'eof', value: '' });
  return tokens;
}

/* ============================================================
   AST и парсер
   ============================================================ */

export type Node =
  | { kind: 'number'; value: number }
  | { kind: 'string'; value: string }
  | { kind: 'bool'; value: boolean }
  | { kind: 'error'; value: string }
  | { kind: 'ref'; addr: CellAddress; raw: string }
  | { kind: 'range'; start: CellAddress; end: CellAddress; raw: string }
  | { kind: 'name'; value: string }
  | { kind: 'unary'; op: string; operand: Node }
  | { kind: 'percent'; operand: Node }
  | { kind: 'binary'; op: string; left: Node; right: Node }
  | { kind: 'call'; name: string; args: Node[] };

class Parser {
  private tokens: Token[];
  private pos = 0;

  constructor(tokens: Token[]) {
    this.tokens = tokens;
  }

  private peek(): Token {
    return this.tokens[this.pos];
  }

  private next(): Token {
    return this.tokens[this.pos++];
  }

  private expect(type: TokenType): Token {
    const t = this.peek();
    if (t.type !== type) {
      throw new Error(`Ожидался ${type}, получен ${t.type} ("${t.value}")`);
    }
    return this.next();
  }

  parse(): Node {
    if (this.peek().type === 'eof') {
      return { kind: 'string', value: '' };
    }
    const node = this.parseComparison();
    return node;
  }

  private parseComparison(): Node {
    let left = this.parseConcat();
    while (
      this.peek().type === 'op' &&
      ['=', '<>', '<', '>', '<=', '>='].includes(this.peek().value)
    ) {
      const op = this.next().value;
      const right = this.parseConcat();
      left = { kind: 'binary', op, left, right };
    }
    return left;
  }

  private parseConcat(): Node {
    let left = this.parseAdditive();
    while (this.peek().type === 'op' && this.peek().value === '&') {
      this.next();
      const right = this.parseAdditive();
      left = { kind: 'binary', op: '&', left, right };
    }
    return left;
  }

  private parseAdditive(): Node {
    let left = this.parseMultiplicative();
    while (this.peek().type === 'op' && (this.peek().value === '+' || this.peek().value === '-')) {
      const op = this.next().value;
      const right = this.parseMultiplicative();
      left = { kind: 'binary', op, left, right };
    }
    return left;
  }

  private parseMultiplicative(): Node {
    let left = this.parseExponent();
    while (this.peek().type === 'op' && (this.peek().value === '*' || this.peek().value === '/')) {
      const op = this.next().value;
      const right = this.parseExponent();
      left = { kind: 'binary', op, left, right };
    }
    return left;
  }

  private parseExponent(): Node {
    const left = this.parseUnary();
    if (this.peek().type === 'op' && this.peek().value === '^') {
      this.next();
      // Право-ассоциативный: 2^3^2 = 2^(3^2)
      const right = this.parseExponent();
      return { kind: 'binary', op: '^', left, right };
    }
    return left;
  }

  private parseUnary(): Node {
    const t = this.peek();
    if (t.type === 'op' && (t.value === '-' || t.value === '+')) {
      this.next();
      const operand = this.parseUnary();
      return { kind: 'unary', op: t.value, operand };
    }
    return this.parsePostfix();
  }

  private parsePostfix(): Node {
    let node = this.parsePrimary();
    while (this.peek().type === 'op' && this.peek().value === '%') {
      this.next();
      node = { kind: 'percent', operand: node };
    }
    return node;
  }

  private parsePrimary(): Node {
    const t = this.peek();

    if (t.type === 'number') {
      this.next();
      return { kind: 'number', value: parseFloat(t.value) };
    }

    if (t.type === 'string') {
      this.next();
      return { kind: 'string', value: t.value };
    }

    if (t.type === 'ref') {
      this.next();
      const addr = parseCellRef(t.value);
      if (!addr) return { kind: 'error', value: ERRORS.REF };
      return { kind: 'ref', addr, raw: t.value };
    }

    if (t.type === 'range' && t.range) {
      this.next();
      return {
        kind: 'range',
        start: t.range.start,
        end: t.range.end,
        raw: t.value,
      };
    }

    if (t.type === 'lparen') {
      this.next();
      const node = this.parseComparison();
      this.expect('rparen');
      return node;
    }

    if (t.type === 'ident') {
      this.next();
      const upper = t.value.toUpperCase();

      if (upper === 'TRUE' || upper === 'ИСТИНА') return { kind: 'bool', value: true };
      if (upper === 'FALSE' || upper === 'ЛОЖЬ') return { kind: 'bool', value: false };

      // Имя листа как ссылка на весь лист: Лист1!A:A / Лист!1:1 обрабатываются
      // в лексере; здесь — вызов функции.
      if (this.peek().type === 'lparen') {
        this.next(); // (
        const args: Node[] = [];
        if (this.peek().type !== 'rparen') {
          args.push(this.parseComparison());
          while (this.peek().type === 'comma') {
            this.next();
            // Пустой аргумент: SUM(A1;;B1)
            if (this.peek().type === 'comma' || this.peek().type === 'rparen') {
              args.push({ kind: 'error', value: ERRORS.NULL });
              continue;
            }
            args.push(this.parseComparison());
          }
        }
        this.expect('rparen');
        return { kind: 'call', name: upper, args };
      }

      return { kind: 'name', value: t.value };
    }

    // Неожиданный токен — не ломаем вычисление цепочки, возвращаем ошибку
    this.next();
    return { kind: 'error', value: ERRORS.NAME };
  }
}

export function parseFormula(formula: string): Node {
  const body = formula.startsWith('=') ? formula.slice(1) : formula;
  return new Parser(tokenize(body)).parse();
}

/* ============================================================
   Приведение типов
   ============================================================ */

/** Текст -> число. Пустая строка -> 0. Нечисловой текст -> null. */
function toNumber(v: unknown): number | null {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return null;
  const s = String(v).trim();
  if (s === '') return 0;
  // Поддержка запятой как десятичного разделителя и процентов
  const cleaned = s.replace(/\s/g, '').replace(',', '.');
  if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(cleaned)) {
    return parseFloat(cleaned);
  }
  const pct = cleaned.match(/^([+-]?(?:\d+\.?\d*|\.\d+))%$/);
  if (pct) return parseFloat(pct[1]) / 100;
  return null;
}

/** Строгое числовое приведение: текст, не являющийся числом, — ошибка. */
function toNumberStrict(v: unknown): number | FormulaError {
  if (isError(v)) return v;
  const n = toNumber(v);
  if (n === null) return err(ERRORS.VALUE);
  return n;
}

function toText(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return v.__error;
  if (typeof v === 'number') {
    // Как в Excel: без лишних нулей, с разумной точностью
    if (Number.isInteger(v)) return String(v);
    return String(Math.round(v * 1e10) / 1e10);
  }
  return String(v);
}

function toBool(v: unknown): boolean | FormulaError {
  if (isError(v)) return v;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === undefined) return false;
  const s = String(v).trim().toUpperCase();
  if (s === 'TRUE') return true;
  if (s === 'FALSE') return false;
  const n = toNumber(v);
  if (n !== null) return n !== 0;
  return err(ERRORS.VALUE);
}

/** Сравнение значений по правилам Excel: числа < текст. */
function compareValues(a: unknown, b: unknown): number {
  const ae = isError(a);
  const be = isError(b);
  if (ae || be) return 0;

  const aEmpty = a === null || a === undefined || a === '';
  const bEmpty = b === null || b === undefined || b === '';
  if (aEmpty && bEmpty) return 0;
  if (aEmpty) return -1;
  if (bEmpty) return 1;

  const an = typeof a === 'boolean' ? (a ? 1 : 0) : null;
  const bn = typeof b === 'boolean' ? (b ? 1 : 0) : null;
  const aNum = an !== null ? an : typeof a === 'number' ? a : toNumber(a);
  const bNum = bn !== null ? bn : typeof b === 'number' ? b : toNumber(b);

  const aIsNum = aNum !== null;
  const bIsNum = bNum !== null;

  if (aIsNum && bIsNum) return aNum < bNum ? -1 : aNum > bNum ? 1 : 0;
  if (aIsNum && !bIsNum) return -1;
  if (!aIsNum && bIsNum) return 1;

  const as = String(a).toLowerCase();
  const bs = String(b).toLowerCase();
  return as < bs ? -1 : as > bs ? 1 : 0;
}

/* ============================================================
   Вычислитель
   ============================================================ */

export interface EvalContext {
  book: Workbook;
  /** Активный лист — для ссылок без имени листа. */
  sheet: string;
  /** Кэш вычисленных значений: "лист!A1" -> значение. */
  cache: Map<string, CellPrimitive | FormulaError>;
  /** Стек вычисляемых ячеек — обнаружение циклов. */
  stack: Set<string>;
  /** Счётчик шагов — защита от лавинообразных вычислений. */
  budget: { steps: number };
}

const MAX_STEPS = 2_000_000;

/** Значение ячейки с учётом формул и кэша. */
function cellValue(ctx: EvalContext, sheetName: string | undefined, col: number, row: number): CellPrimitive | FormulaError {
  if (col < 0 || row < 0) return err(ERRORS.REF);
  const sheet = sheetName || ctx.sheet;
  const key = `${sheet}!${col},${row}`;

  if (ctx.cache.has(key)) return ctx.cache.get(key)!;

  const grid = ctx.book[sheet];
  if (!grid) return err(ERRORS.REF);

  const raw = grid[row]?.[col];

  if (raw === undefined || raw === null) {
    ctx.cache.set(key, null);
    return null;
  }

  if (typeof raw === 'string' && raw.startsWith('=')) {
    if (ctx.stack.has(key)) {
      // Циклическая ссылка
      return err(ERRORS.CIRCULAR);
    }
    if (ctx.budget.steps++ > MAX_STEPS) {
      return err(ERRORS.CIRCULAR);
    }
    ctx.stack.add(key);
    let result: CellPrimitive | FormulaError;
    try {
      result = evaluateFormula(raw, { ...ctx, sheet });
    } catch {
      result = err(ERRORS.VALUE);
    } finally {
      ctx.stack.delete(key);
    }
    ctx.cache.set(key, result);
    return result;
  }

  if (typeof raw === 'boolean') {
    ctx.cache.set(key, raw);
    return raw;
  }

  ctx.cache.set(key, raw);
  return raw;
}

/** Разворачивает узел в список значений (диапазоны — поэлементно). */
function resolveValues(node: Node, ctx: EvalContext): (CellPrimitive | FormulaError)[] {
  if (node.kind === 'range') {
    const sheet = node.start.sheet || node.end.sheet || ctx.sheet;
    const r1 = Math.min(node.start.row, node.end.row);
    const r2 = Math.max(node.start.row, node.end.row);
    const c1 = Math.min(node.start.col, node.end.col);
    const c2 = Math.max(node.start.col, node.end.col);
    const values: (CellPrimitive | FormulaError)[] = [];
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        values.push(cellValue(ctx, sheet, c, r));
      }
    }
    return values;
  }
  return [evalNode(node, ctx)];
}

/** Разворачивает узел в «плоский» список для агрегатных функций. */
function flattenValues(node: Node, ctx: EvalContext): (CellPrimitive | FormulaError)[] {
  if (node.kind === 'range') return resolveValues(node, ctx);
  const v = evalNode(node, ctx);
  // Результат поэлементной арифметики над диапазонами (формула массива)
  // разворачивается, чтобы SUM/AVERAGE увидели каждое значение.
  if (Array.isArray(v)) return v as (CellPrimitive | FormulaError)[];
  return [v];
}

/** Числовые значения для агрегатов — текст игнорируется (как в Excel). */
function numericArgs(nodes: Node[], ctx: EvalContext, skipText = true): number[] {
  const out: number[] = [];
  for (const node of nodes) {
    const values = flattenValues(node, ctx);
    const isRange = node.kind === 'range';
    for (const v of values) {
      if (isError(v)) continue;
      if (typeof v === 'number') {
        out.push(v);
      } else if (!skipText || !isRange) {
        // Прямо указанный текстовый аргумент Excel пытается преобразовать
        const n = toNumber(v);
        if (n !== null) out.push(n);
      }
    }
  }
  return out;
}

/** Первая ошибка среди аргументов — пробрасывается наружу. */
function firstError(values: (CellPrimitive | FormulaError)[]): FormulaError | null {
  for (const v of values) if (isError(v)) return v;
  return null;
}

/* ---------- Реализации функций ---------- */

type FnImpl = (args: Node[], ctx: EvalContext) => CellPrimitive | FormulaError;

/** Аргумент как скаляр. */
function argValue(args: Node[], i: number, ctx: EvalContext): CellPrimitive | FormulaError {
  if (i >= args.length) return null;
  return evalNode(args[i], ctx);
}

/** Аргумент как число. */
function argNum(args: Node[], i: number, ctx: EvalContext, fallback = 0): number | FormulaError {
  if (i >= args.length) return fallback;
  const v = evalNode(args[i], ctx);
  if (isError(v)) return v;
  const n = toNumberStrict(v);
  return n;
}

function argText(args: Node[], i: number, ctx: EvalContext, fallback = ''): string | FormulaError {
  if (i >= args.length) return fallback;
  const v = evalNode(args[i], ctx);
  if (isError(v)) return v;
  return toText(v);
}

function argBool(args: Node[], i: number, ctx: EvalContext, fallback = false): boolean | FormulaError {
  if (i >= args.length) return fallback;
  const v = evalNode(args[i], ctx);
  if (isError(v)) return v;
  return toBool(v);
}

/** Значения диапазона/аргумента с сохранением ошибок. */
function argValues(args: Node[], i: number, ctx: EvalContext): (CellPrimitive | FormulaError)[] {
  if (i >= args.length) return [];
  return flattenValues(args[i], ctx);
}

/** Числовые значения аргументов. */
function argNums(args: Node[], i: number, ctx: EvalContext, skipText = true): number[] {
  const list: Node[] = args.slice(i);
  return numericArgs(list, ctx, skipText);
}

/**
 * Разбор критерия для SUMIF/COUNTIF/AVERAGEIF и т.п.
 * Поддерживает ">10", "<=5", "<>текст", "текст", 42, "*abc*".
 */
function makeCriteria(raw: unknown): (v: CellPrimitive | FormulaError) => boolean {
  if (isError(raw)) return () => false;

  if (typeof raw === 'number') return (v) => !isError(v) && toNumber(v) === raw;
  if (typeof raw === 'boolean') return (v) => !isError(v) && toBool(v) === raw;

  const s = String(raw ?? '');
  const m = s.match(/^(<=|>=|<>|=|<|>)\s*(.*)$/);
  if (m) {
    const op = m[1];
    const operand = m[2].trim();
    const operandNum = operand === '' ? null : toNumber(operand);
    const operandBool = /^(true|false)$/i.test(operand) ? operand.toLowerCase() === 'true' : null;

    return (v) => {
      if (isError(v)) return false;
      if (op === '<>' || op === '=') {
        const equal =
          operandNum !== null
            ? toNumber(v) === operandNum && !(v === null || v === '')
            : operandBool !== null
              ? toBool(v) === operandBool
              : String(v ?? '').toLowerCase() === operand.toLowerCase();
        return op === '=' ? equal : !equal;
      }
      const vn = toNumber(v);
      const target = operandNum ?? toNumber(operand);
      if (target === null || vn === null) {
        // Сравнение строк
        const c = String(v ?? '').toLowerCase().localeCompare(operand.toLowerCase());
        switch (op) {
          case '<':
            return c < 0;
          case '>':
            return c > 0;
          case '<=':
            return c <= 0;
          case '>=':
            return c >= 0;
        }
        return false;
      }
      switch (op) {
        case '<':
          return vn < target;
        case '>':
          return vn > target;
        case '<=':
          return vn <= target;
        case '>=':
          return vn >= target;
      }
      return false;
    };
  }

  // Шаблон с подстановочными знаками * и ?
  if (s.includes('*') || s.includes('?')) {
    const rx = new RegExp(
      `^${s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')}$`,
      'i'
    );
    return (v) => !isError(v) && rx.test(String(v ?? ''));
  }

  const num = toNumber(s);
  return (v) => {
    if (isError(v)) return false;
    if (num !== null) return toNumber(v) === num;
    return String(v ?? '').toLowerCase() === s.toLowerCase();
  };
}

/** Разбор диапазона для функций работы с массивами. */
function rangeBounds(node: Node): { c1: number; r1: number; c2: number; r2: number; sheet: string } | null {
  if (node.kind !== 'range') return null;
  return {
    sheet: node.start.sheet || node.end.sheet || '',
    r1: Math.min(node.start.row, node.end.row),
    r2: Math.max(node.start.row, node.end.row),
    c1: Math.min(node.start.col, node.end.col),
    c2: Math.max(node.start.col, node.end.col),
  };
}

const FUNCTIONS: Record<string, FnImpl> = {
  /* ---------- Математические ---------- */
  SUM: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    return nums.reduce((s, v) => s + v, 0);
  },

  SUMIF: (args, ctx) => {
    const rangeVals = argValues(args, 0, ctx);
    const criteriaRaw = argValue(args, 1, ctx);
    if (isError(criteriaRaw)) return criteriaRaw;
    const test = makeCriteria(criteriaRaw);
    const sumVals = args.length > 2 ? argValues(args, 2, ctx) : rangeVals;
    let total = 0;
    for (let i = 0; i < rangeVals.length; i++) {
      if (isError(rangeVals[i])) continue;
      if (test(rangeVals[i])) {
        const n = toNumber(sumVals[i]);
        if (n !== null) total += n;
      }
    }
    return total;
  },

  SUMIFS: (args, ctx) => {
    const sumVals = argValues(args, 0, ctx);
    let total = 0;
    const pairs: { values: (CellPrimitive | FormulaError)[]; test: (v: any) => boolean }[] = [];
    for (let i = 1; i + 1 < args.length; i += 2) {
      const criteriaRaw = argValue(args, i + 1, ctx);
      if (isError(criteriaRaw)) return criteriaRaw;
      pairs.push({ values: argValues(args, i, ctx), test: makeCriteria(criteriaRaw) });
    }
    for (let i = 0; i < sumVals.length; i++) {
      if (isError(sumVals[i])) continue;
      if (pairs.every((p) => p.test(p.values[i]))) {
        const n = toNumber(sumVals[i]);
        if (n !== null) total += n;
      }
    }
    return total;
  },

  SUMPRODUCT: (args, ctx) => {
    const arrays = args.map((_a, i) => argNums(args, i, ctx));
    if (arrays.length === 0) return err(ERRORS.VALUE);
    const len = Math.max(...arrays.map((a) => a.length));
    let total = 0;
    for (let i = 0; i < len; i++) {
      let product = 1;
      for (const arr of arrays) product *= arr[i] ?? 0;
      total += product;
    }
    return total;
  },

  PRODUCT: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    if (nums.length === 0) return 0;
    return nums.reduce((s, v) => s * v, 1);
  },

  ABS: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    return isError(n) ? n : Math.abs(n);
  },

  ROUND: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    const d = argNum(args, 1, ctx, 0);
    if (isError(d)) return d;
    const f = Math.pow(10, d);
    return Math.round((n + Number.EPSILON * Math.sign(n)) * f) / f;
  },

  ROUNDUP: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    const d = argNum(args, 1, ctx, 0);
    if (isError(d)) return d;
    const f = Math.pow(10, d);
    return (n < 0 ? -Math.ceil(-n * f) : Math.ceil(n * f)) / f;
  },

  ROUNDDOWN: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    const d = argNum(args, 1, ctx, 0);
    if (isError(d)) return d;
    const f = Math.pow(10, d);
    return (n < 0 ? -Math.floor(-n * f) : Math.floor(n * f)) / f;
  },

  INT: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    return isError(n) ? n : Math.floor(n);
  },

  TRUNC: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    const d = argNum(args, 1, ctx, 0);
    if (isError(d)) return d;
    const f = Math.pow(10, d);
    return Math.trunc(n * f) / f;
  },

  MOD: (args, ctx) => {
    const a = argNum(args, 0, ctx, 0);
    if (isError(a)) return a;
    const b = argNum(args, 1, ctx, 0);
    if (isError(b)) return b;
    if (b === 0) return err(ERRORS.DIV0);
    return a - b * Math.floor(a / b);
  },

  POWER: (args, ctx) => {
    const a = argNum(args, 0, ctx, 0);
    if (isError(a)) return a;
    const b = argNum(args, 1, ctx, 0);
    if (isError(b)) return b;
    const r = Math.pow(a, b);
    return Number.isFinite(r) ? r : err(ERRORS.NUM);
  },

  SQRT: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    return n < 0 ? err(ERRORS.NUM) : Math.sqrt(n);
  },

  EXP: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    return isError(n) ? n : Math.exp(n);
  },

  LN: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    return n <= 0 ? err(ERRORS.NUM) : Math.log(n);
  },

  LOG: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    const base = argNum(args, 1, ctx, 10);
    if (isError(base)) return base;
    if (n <= 0 || base <= 0) return err(ERRORS.NUM);
    return Math.log(n) / Math.log(base);
  },

  LOG10: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    return n <= 0 ? err(ERRORS.NUM) : Math.log10(n);
  },

  SIGN: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    return isError(n) ? n : Math.sign(n);
  },

  PI: () => Math.PI,

  RAND: () => Math.random(),

  RANDBETWEEN: (args, ctx) => {
    const lo = argNum(args, 0, ctx, 0);
    if (isError(lo)) return lo;
    const hi = argNum(args, 1, ctx, 1);
    if (isError(hi)) return hi;
    return Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) + Math.ceil(lo);
  },

  CEILING: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    const sig = argNum(args, 1, ctx, 1);
    if (isError(sig)) return sig;
    if (sig === 0) return 0;
    return Math.ceil(n / sig) * sig;
  },

  FLOOR: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    const sig = argNum(args, 1, ctx, 1);
    if (isError(sig)) return sig;
    if (sig === 0) return err(ERRORS.DIV0);
    return Math.floor(n / sig) * sig;
  },

  MROUND: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    const mult = argNum(args, 1, ctx, 1);
    if (isError(mult)) return mult;
    if (mult === 0) return 0;
    return (n < 0 ? -1 : 1) * Math.round(Math.abs(n) / Math.abs(mult)) * Math.abs(mult);
  },

  /* ---------- Статистические ---------- */
  AVERAGE: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    if (nums.length === 0) return err(ERRORS.DIV0);
    return nums.reduce((s, v) => s + v, 0) / nums.length;
  },

  AVERAGEA: (args, ctx) => {
    const nums = argNums(args, 0, ctx, false);
    if (nums.length === 0) return err(ERRORS.DIV0);
    return nums.reduce((s, v) => s + v, 0) / nums.length;
  },

  AVERAGEIF: (args, ctx) => {
    const rangeVals = argValues(args, 0, ctx);
    const criteriaRaw = argValue(args, 1, ctx);
    if (isError(criteriaRaw)) return criteriaRaw;
    const test = makeCriteria(criteriaRaw);
    const avgVals = args.length > 2 ? argValues(args, 2, ctx) : rangeVals;
    let total = 0;
    let count = 0;
    for (let i = 0; i < rangeVals.length; i++) {
      if (isError(rangeVals[i])) continue;
      if (test(rangeVals[i])) {
        const n = toNumber(avgVals[i]);
        if (n !== null) {
          total += n;
          count++;
        }
      }
    }
    return count === 0 ? err(ERRORS.DIV0) : total / count;
  },

  COUNT: (args, ctx) => {
    let count = 0;
    for (const node of args) {
      const values = flattenValues(node, ctx);
      for (const v of values) {
        if (isError(v)) continue;
        if (typeof v === 'number') count++;
        else if (node.kind !== 'range' && toNumber(v) !== null && String(v ?? '') !== '') count++;
      }
    }
    return count;
  },

  COUNTA: (args, ctx) => {
    let count = 0;
    for (const node of args) {
      const values = flattenValues(node, ctx);
      for (const v of values) {
        if (isError(v)) {
          count++;
          continue;
        }
        if (v !== null && v !== undefined && String(v) !== '') count++;
      }
    }
    return count;
  },

  COUNTBLANK: (args, ctx) => {
    const values = argValues(args, 0, ctx);
    let count = 0;
    for (const v of values) {
      if (v === null || v === undefined || String(v) === '') count++;
    }
    return count;
  },

  COUNTIF: (args, ctx) => {
    const rangeVals = argValues(args, 0, ctx);
    const criteriaRaw = argValue(args, 1, ctx);
    if (isError(criteriaRaw)) return criteriaRaw;
    const test = makeCriteria(criteriaRaw);
    return rangeVals.filter((v) => !isError(v) && test(v)).length;
  },

  COUNTIFS: (args, ctx) => {
    const first = argValues(args, 0, ctx);
    const pairs: { values: (CellPrimitive | FormulaError)[]; test: (v: any) => boolean }[] = [];
    for (let i = 0; i + 1 < args.length; i += 2) {
      const criteriaRaw = argValue(args, i + 1, ctx);
      if (isError(criteriaRaw)) return criteriaRaw;
      pairs.push({ values: argValues(args, i, ctx), test: makeCriteria(criteriaRaw) });
    }
    if (pairs.length === 0) return 0;
    let count = 0;
    for (let i = 0; i < first.length; i++) {
      if (pairs.every((p) => !isError(p.values[i]) && p.test(p.values[i]))) count++;
    }
    return count;
  },

  MIN: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    return nums.length === 0 ? 0 : Math.min(...nums);
  },

  MAX: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    return nums.length === 0 ? 0 : Math.max(...nums);
  },

  MEDIAN: (args, ctx) => {
    const nums = argNums(args, 0, ctx).slice().sort((a, b) => a - b);
    if (nums.length === 0) return err(ERRORS.NUM);
    const mid = Math.floor(nums.length / 2);
    return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
  },

  MODE: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    const counts = new Map<number, number>();
    for (const n of nums) counts.set(n, (counts.get(n) ?? 0) + 1);
    let best: number | null = null;
    let bestCount = 1;
    for (const [n, c] of counts) {
      if (c > bestCount) {
        best = n;
        bestCount = c;
      }
    }
    return best === null ? err(ERRORS.NA) : best;
  },

  STDEV: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    if (nums.length < 2) return err(ERRORS.DIV0);
    const mean = nums.reduce((s, v) => s + v, 0) / nums.length;
    return Math.sqrt(nums.reduce((s, v) => s + (v - mean) ** 2, 0) / (nums.length - 1));
  },

  STDEVP: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    if (nums.length < 1) return err(ERRORS.DIV0);
    const mean = nums.reduce((s, v) => s + v, 0) / nums.length;
    return Math.sqrt(nums.reduce((s, v) => s + (v - mean) ** 2, 0) / nums.length);
  },

  VAR: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    if (nums.length < 2) return err(ERRORS.DIV0);
    const mean = nums.reduce((s, v) => s + v, 0) / nums.length;
    return nums.reduce((s, v) => s + (v - mean) ** 2, 0) / (nums.length - 1);
  },

  VARP: (args, ctx) => {
    const nums = argNums(args, 0, ctx);
    if (nums.length < 1) return err(ERRORS.DIV0);
    const mean = nums.reduce((s, v) => s + v, 0) / nums.length;
    return nums.reduce((s, v) => s + (v - mean) ** 2, 0) / nums.length;
  },

  LARGE: (args, ctx) => {
    // Только первый аргумент — данные; второй — номер позиции.
    const nums = numericArgs(args.slice(0, 1), ctx, false).sort((a, b) => b - a);
    const k = argNum(args, 1, ctx, 1);
    if (isError(k)) return k;
    if (k < 1 || k > nums.length) return err(ERRORS.NUM);
    return nums[k - 1];
  },

  SMALL: (args, ctx) => {
    const nums = numericArgs(args.slice(0, 1), ctx, false).sort((a, b) => a - b);
    const k = argNum(args, 1, ctx, 1);
    if (isError(k)) return k;
    if (k < 1 || k > nums.length) return err(ERRORS.NUM);
    return nums[k - 1];
  },

  RANK: (args, ctx) => {
    const value = argNum(args, 0, ctx, 0);
    if (isError(value)) return value;
    const nums = argNums(args.slice(1, 2), 0, ctx);
    const order = args.length > 2 ? argNum(args, 2, ctx, 0) : 0;
    if (isError(order)) return order;
    const sorted = nums.slice().sort((a, b) => (order ? a - b : b - a));
    const idx = sorted.indexOf(value);
    return idx === -1 ? err(ERRORS.NA) : idx + 1;
  },

  PERCENTILE: (args, ctx) => {
    // Только первый аргумент — данные; второй — доля.
    const nums = numericArgs(args.slice(0, 1), ctx, false).sort((a, b) => a - b);
    const p = argNum(args, 1, ctx, 0);
    if (isError(p)) return p;
    if (nums.length === 0 || p < 0 || p > 1) return err(ERRORS.NUM);
    const pos = (nums.length - 1) * p;
    const lo = Math.floor(pos);
    const hi = Math.ceil(pos);
    return lo === hi ? nums[lo] : nums[lo] + (nums[hi] - nums[lo]) * (pos - lo);
  },

  /* ---------- Логические ---------- */
  IF: (args, ctx) => {
    const cond = argBool(args, 0, ctx, false);
    if (isError(cond)) return cond;
    if (cond) return args.length > 1 ? evalNode(args[1], ctx) : true;
    return args.length > 2 ? evalNode(args[2], ctx) : false;
  },

  IFERROR: (args, ctx) => {
    const v = argValue(args, 0, ctx);
    if (isError(v)) return args.length > 1 ? evalNode(args[1], ctx) : '';
    return v;
  },

  IFNA: (args, ctx) => {
    const v = argValue(args, 0, ctx);
    if (isError(v) && v.__error === ERRORS.NA) return args.length > 1 ? evalNode(args[1], ctx) : '';
    return v;
  },

  AND: (args, ctx) => {
    let sawValue = false;
    for (const node of args) {
      for (const v of flattenValues(node, ctx)) {
        if (isError(v)) return v;
        if (node.kind === 'range' && (v === null || v === '')) continue;
        const b = toBool(v);
        if (isError(b)) return b;
        sawValue = true;
        if (!b) return false;
      }
    }
    return sawValue ? true : err(ERRORS.VALUE);
  },

  OR: (args, ctx) => {
    let sawValue = false;
    for (const node of args) {
      for (const v of flattenValues(node, ctx)) {
        if (isError(v)) return v;
        if (node.kind === 'range' && (v === null || v === '')) continue;
        const b = toBool(v);
        if (isError(b)) return b;
        sawValue = true;
        if (b) return true;
      }
    }
    return sawValue ? false : err(ERRORS.VALUE);
  },

  XOR: (args, ctx) => {
    let count = 0;
    for (const node of args) {
      for (const v of flattenValues(node, ctx)) {
        if (isError(v)) return v;
        const b = toBool(v);
        if (isError(b)) return b;
        if (b) count++;
      }
    }
    return count % 2 === 1;
  },

  NOT: (args, ctx) => {
    const b = argBool(args, 0, ctx, false);
    return isError(b) ? b : !b;
  },

  TRUE: () => true,
  FALSE: () => false,

  /* ---------- Текстовые ---------- */
  CONCAT: (args, ctx) => {
    let out = '';
    for (const node of args) {
      for (const v of flattenValues(node, ctx)) {
        if (isError(v)) return v;
        out += toText(v);
      }
    }
    return out;
  },

  CONCATENATE: (args, ctx) => {
    let out = '';
    for (let i = 0; i < args.length; i++) {
      const t = argText(args, i, ctx, '');
      if (isError(t)) return t;
      out += t;
    }
    return out;
  },

  TEXTJOIN: (args, ctx) => {
    const delim = argText(args, 0, ctx, '');
    if (isError(delim)) return delim;
    const ignoreEmpty = argBool(args, 1, ctx, true);
    if (isError(ignoreEmpty)) return ignoreEmpty;
    const parts: string[] = [];
    for (let i = 2; i < args.length; i++) {
      for (const v of flattenValues(args[i], ctx)) {
        if (isError(v)) return v;
        const t = toText(v);
        if (ignoreEmpty && t === '') continue;
        parts.push(t);
      }
    }
    return parts.join(delim);
  },

  LEN: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    return isError(t) ? t : t.length;
  },

  LEFT: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    if (isError(t)) return t;
    const n = argNum(args, 1, ctx, 1);
    if (isError(n)) return n;
    return t.slice(0, Math.max(0, n));
  },

  RIGHT: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    if (isError(t)) return t;
    const n = argNum(args, 1, ctx, 1);
    if (isError(n)) return n;
    return n <= 0 ? '' : t.slice(-n);
  },

  MID: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    if (isError(t)) return t;
    const start = argNum(args, 1, ctx, 1);
    if (isError(start)) return start;
    const len = argNum(args, 2, ctx, 0);
    if (isError(len)) return len;
    if (start < 1 || len < 0) return err(ERRORS.VALUE);
    return t.slice(start - 1, start - 1 + len);
  },

  UPPER: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    return isError(t) ? t : t.toUpperCase();
  },

  LOWER: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    return isError(t) ? t : t.toLowerCase();
  },

  PROPER: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    if (isError(t)) return t;
    return t.replace(/\p{L}+/gu, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase());
  },

  TRIM: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    return isError(t) ? t : t.trim().replace(/\s+/g, ' ');
  },

  CLEAN: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    // eslint-disable-next-line no-control-regex
    return isError(t) ? t : t.replace(/[\x00-\x1F\x7F]/g, '');
  },

  REPT: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    if (isError(t)) return t;
    const n = argNum(args, 1, ctx, 1);
    if (isError(n)) return n;
    if (n < 0) return err(ERRORS.VALUE);
    if (n * t.length > 200000) return err(ERRORS.NUM);
    return t.repeat(Math.floor(n));
  },

  SUBSTITUTE: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    if (isError(t)) return t;
    const oldText = argText(args, 1, ctx, '');
    if (isError(oldText)) return oldText;
    const newText = argText(args, 2, ctx, '');
    if (isError(newText)) return newText;
    if (args.length > 3) {
      const nth = argNum(args, 3, ctx, 1);
      if (isError(nth)) return nth;
      let count = 0;
      let idx = 0;
      let out = '';
      while (true) {
        const found = t.indexOf(oldText, idx);
        if (found === -1 || oldText === '') {
          out += t.slice(idx);
          break;
        }
        count++;
        out += t.slice(idx, found);
        out += count === nth ? newText : oldText;
        idx = found + oldText.length;
      }
      return out;
    }
    return oldText === '' ? t : t.split(oldText).join(newText);
  },

  REPLACE: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    if (isError(t)) return t;
    const start = argNum(args, 1, ctx, 1);
    if (isError(start)) return start;
    const len = argNum(args, 2, ctx, 0);
    if (isError(len)) return len;
    const newText = argText(args, 3, ctx, '');
    if (isError(newText)) return newText;
    if (start < 1 || len < 0) return err(ERRORS.VALUE);
    return t.slice(0, start - 1) + newText + t.slice(start - 1 + len);
  },

  FIND: (args, ctx) => {
    const needle = argText(args, 0, ctx, '');
    if (isError(needle)) return needle;
    const hay = argText(args, 1, ctx, '');
    if (isError(hay)) return hay;
    const start = argNum(args, 2, ctx, 1);
    if (isError(start)) return start;
    if (start < 1) return err(ERRORS.VALUE);
    const idx = hay.indexOf(needle, start - 1);
    return idx === -1 ? err(ERRORS.VALUE) : idx + 1;
  },

  SEARCH: (args, ctx) => {
    const needle = argText(args, 0, ctx, '');
    if (isError(needle)) return needle;
    const hay = argText(args, 1, ctx, '');
    if (isError(hay)) return hay;
    const start = argNum(args, 2, ctx, 1);
    if (isError(start)) return start;
    if (start < 1) return err(ERRORS.VALUE);
    if (!needle.includes('*') && !needle.includes('?')) {
      const idx = hay.toLowerCase().indexOf(needle.toLowerCase(), start - 1);
      return idx === -1 ? err(ERRORS.VALUE) : idx + 1;
    }
    const rx = new RegExp(
      needle.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.'),
      'i'
    );
    const m = hay.slice(start - 1).match(rx);
    return m && m.index !== undefined ? m.index + start : err(ERRORS.VALUE);
  },

  VALUE: (args, ctx) => {
    const v = argValue(args, 0, ctx);
    if (isError(v)) return v;
    const n = toNumber(v);
    return n === null ? err(ERRORS.VALUE) : n;
  },

  NUMBERVALUE: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    if (isError(t)) return t;
    const decSep = args.length > 1 ? argText(args, 1, ctx, '.') : '.';
    if (isError(decSep)) return decSep;
    const groupSep = args.length > 2 ? argText(args, 2, ctx, ',') : ',';
    if (isError(groupSep)) return groupSep;
    let s = t.split(groupSep).join('').split(decSep).join('.');
    s = s.replace(/[^0-9eE+\-.]/g, '');
    const n = parseFloat(s);
    return Number.isNaN(n) ? err(ERRORS.VALUE) : n;
  },

  TEXT: (args, ctx) => {
    const v = argValue(args, 0, ctx);
    if (isError(v)) return v;
    const fmt = argText(args, 1, ctx, '');
    if (isError(fmt)) return fmt;
    return formatValue(v, fmt);
  },

  CHAR: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    if (isError(n)) return n;
    return n < 1 || n > 1114111 ? err(ERRORS.VALUE) : String.fromCodePoint(n);
  },

  CODE: (args, ctx) => {
    const t = argText(args, 0, ctx, '');
    if (isError(t)) return t;
    return t.length === 0 ? err(ERRORS.VALUE) : t.codePointAt(0)!;
  },

  EXACT: (args, ctx) => {
    const a = argText(args, 0, ctx, '');
    if (isError(a)) return a;
    const b = argText(args, 1, ctx, '');
    if (isError(b)) return b;
    return a === b;
  },

  /* ---------- Поиск и ссылки ---------- */
  VLOOKUP: (args, ctx) => {
    const needle = argValue(args, 0, ctx);
    if (isError(needle)) return needle;
    const bounds = rangeBounds(args[1]);
    if (!bounds) return err(ERRORS.VALUE);
    const colIdx = argNum(args, 2, ctx, 1);
    if (isError(colIdx)) return colIdx;
    const approx = args.length > 3 ? argBool(args, 3, ctx, true) : true;
    if (isError(approx)) return approx;

    const sheet = bounds.sheet || ctx.sheet;
    const width = bounds.c2 - bounds.c1 + 1;
    if (colIdx < 1 || colIdx > width) return err(ERRORS.REF);

    let matchRow = -1;
    if (approx) {
      for (let r = bounds.r1; r <= bounds.r2; r++) {
        const v = cellValue(ctx, sheet, bounds.c1, r);
        if (isError(v)) continue;
        if (compareValues(v, needle) <= 0) matchRow = r;
        else break;
      }
    } else {
      for (let r = bounds.r1; r <= bounds.r2; r++) {
        const v = cellValue(ctx, sheet, bounds.c1, r);
        if (isError(v)) continue;
        if (compareValues(v, needle) === 0) {
          matchRow = r;
          break;
        }
      }
    }
    if (matchRow === -1) return err(ERRORS.NA);
    const result = cellValue(ctx, sheet, bounds.c1 + colIdx - 1, matchRow);
    return isError(result) ? result : result === undefined ? null : result;
  },

  HLOOKUP: (args, ctx) => {
    const needle = argValue(args, 0, ctx);
    if (isError(needle)) return needle;
    const bounds = rangeBounds(args[1]);
    if (!bounds) return err(ERRORS.VALUE);
    const rowIdx = argNum(args, 2, ctx, 1);
    if (isError(rowIdx)) return rowIdx;
    const approx = args.length > 3 ? argBool(args, 3, ctx, true) : true;
    if (isError(approx)) return approx;

    const sheet = bounds.sheet || ctx.sheet;
    const height = bounds.r2 - bounds.r1 + 1;
    if (rowIdx < 1 || rowIdx > height) return err(ERRORS.REF);

    let matchCol = -1;
    if (approx) {
      for (let c = bounds.c1; c <= bounds.c2; c++) {
        const v = cellValue(ctx, sheet, c, bounds.r1);
        if (isError(v)) continue;
        if (compareValues(v, needle) <= 0) matchCol = c;
        else break;
      }
    } else {
      for (let c = bounds.c1; c <= bounds.c2; c++) {
        const v = cellValue(ctx, sheet, c, bounds.r1);
        if (isError(v)) continue;
        if (compareValues(v, needle) === 0) {
          matchCol = c;
          break;
        }
      }
    }
    if (matchCol === -1) return err(ERRORS.NA);
    return cellValue(ctx, sheet, matchCol, bounds.r1 + rowIdx - 1);
  },

  INDEX: (args, ctx) => {
    const bounds = rangeBounds(args[0]);
    if (!bounds) return err(ERRORS.VALUE);
    const sheet = bounds.sheet || ctx.sheet;
    const height = bounds.r2 - bounds.r1 + 1;
    const width = bounds.c2 - bounds.c1 + 1;

    const rowNum = argNum(args, 1, ctx, 0);
    if (isError(rowNum)) return rowNum;
    const colNum = args.length > 2 ? argNum(args, 2, ctx, 0) : 0;
    if (isError(colNum)) return colNum;

    // Одномерный диапазон: INDEX(A1:A10, 3)
    if (args.length === 2) {
      if (height === 1) {
        if (rowNum < 1 || rowNum > width) return err(ERRORS.REF);
        return cellValue(ctx, sheet, bounds.c1 + rowNum - 1, bounds.r1);
      }
      if (rowNum < 1 || rowNum > height) return err(ERRORS.REF);
      return cellValue(ctx, sheet, bounds.c1, bounds.r1 + rowNum - 1);
    }

    if (rowNum === 0 && colNum === 0) return err(ERRORS.VALUE);
    if (rowNum > height || colNum > width) return err(ERRORS.REF);

    if (rowNum === 0) {
      // Вся строка — возвращаем первое значение (скалярный контекст)
      return cellValue(ctx, sheet, bounds.c1 + Math.max(1, colNum) - 1, bounds.r1);
    }
    if (colNum === 0) {
      return cellValue(ctx, sheet, bounds.c1, bounds.r1 + rowNum - 1);
    }
    return cellValue(ctx, sheet, bounds.c1 + colNum - 1, bounds.r1 + rowNum - 1);
  },

  MATCH: (args, ctx) => {
    const needle = argValue(args, 0, ctx);
    if (isError(needle)) return needle;
    const type = args.length > 2 ? argNum(args, 2, ctx, 1) : 1;
    if (isError(type)) return type;

    const bounds = rangeBounds(args[1]);
    const values: (CellPrimitive | FormulaError)[] = bounds
      ? collectRange(ctx, bounds.sheet || ctx.sheet, bounds)
      : argValues(args, 1, ctx);

    if (type === 0) {
      for (let i = 0; i < values.length; i++) {
        const v = values[i];
        if (isError(v)) continue;
        if (typeof needle === 'string' && (needle.includes('*') || needle.includes('?'))) {
          const test = makeCriteria(needle);
          if (test(v)) return i + 1;
        } else if (compareValues(v, needle) === 0) {
          return i + 1;
        }
      }
      return err(ERRORS.NA);
    }

    let best = -1;
    if (type === 1) {
      for (let i = 0; i < values.length; i++) {
        const v = values[i];
        if (isError(v)) continue;
        if (compareValues(v, needle) <= 0) best = i;
        else break;
      }
    } else {
      for (let i = 0; i < values.length; i++) {
        const v = values[i];
        if (isError(v)) continue;
        if (compareValues(v, needle) >= 0) best = i;
        else break;
      }
    }
    return best === -1 ? err(ERRORS.NA) : best + 1;
  },

  XLOOKUP: (args, ctx) => {
    const needle = argValue(args, 0, ctx);
    if (isError(needle)) return needle;
    const lookupBounds = rangeBounds(args[1]);
    const returnBounds = rangeBounds(args[2]);
    if (!lookupBounds) return err(ERRORS.VALUE);

    const sheet = lookupBounds.sheet || ctx.sheet;
    const lookupVals = collectRange(ctx, sheet, lookupBounds);
    const returnVals = returnBounds ? collectRange(ctx, returnBounds.sheet || sheet, returnBounds) : [];

    for (let i = 0; i < lookupVals.length; i++) {
      const v = lookupVals[i];
      if (isError(v)) continue;
      const hit =
        typeof needle === 'string' && (needle.includes('*') || needle.includes('?'))
          ? makeCriteria(needle)(v)
          : compareValues(v, needle) === 0;
      if (hit) {
        if (returnVals.length > 0) return returnVals[i] ?? null;
        return v;
      }
    }
    // Не найдено
    if (args.length > 3) return evalNode(args[3], ctx);
    return err(ERRORS.NA);
  },

  CHOOSE: (args, ctx) => {
    const idx = argNum(args, 0, ctx, 1);
    if (isError(idx)) return idx;
    if (idx < 1 || idx >= args.length) return err(ERRORS.VALUE);
    return evalNode(args[idx], ctx);
  },

  ROW: (args, ctx) => {
    if (args.length === 0) return err(ERRORS.VALUE);
    const target = args[0];
    if (target.kind === 'ref') return target.addr.row + 1;
    if (target.kind === 'range') return Math.min(target.start.row, target.end.row) + 1;
    return err(ERRORS.VALUE);
  },

  COLUMN: (args, ctx) => {
    if (args.length === 0) return err(ERRORS.VALUE);
    const target = args[0];
    if (target.kind === 'ref') return target.addr.col + 1;
    if (target.kind === 'range') return Math.min(target.start.col, target.end.col) + 1;
    return err(ERRORS.VALUE);
  },

  ROWS: (args, ctx) => {
    if (args.length === 0) return err(ERRORS.VALUE);
    const target = args[0];
    if (target.kind === 'range') return Math.abs(target.end.row - target.start.row) + 1;
    return 1;
  },

  COLUMNS: (args, ctx) => {
    if (args.length === 0) return err(ERRORS.VALUE);
    const target = args[0];
    if (target.kind === 'range') return Math.abs(target.end.col - target.start.col) + 1;
    return 1;
  },

  /* ---------- Дата и время (серийные числа Excel) ---------- */
  TODAY: () => excelSerialFromDate(new Date()),

  NOW: () => {
    const d = new Date();
    const frac = (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
    return excelSerialFromDate(d) + frac;
  },

  DATE: (args, ctx) => {
    const y = argNum(args, 0, ctx, 1900);
    if (isError(y)) return y;
    const m = argNum(args, 1, ctx, 1);
    if (isError(m)) return m;
    const d = argNum(args, 2, ctx, 1);
    if (isError(d)) return d;
    const date = new Date(Date.UTC(y, m - 1, d));
    date.setUTCFullYear(date.getUTCFullYear() + (y < 100 ? 1900 : 0));
    return excelSerialFromDate(date);
  },

  YEAR: (args, ctx) => {
    const s = argNum(args, 0, ctx, 0);
    if (isError(s)) return s;
    return dateFromSerial(s).getUTCFullYear();
  },

  MONTH: (args, ctx) => {
    const s = argNum(args, 0, ctx, 0);
    if (isError(s)) return s;
    return dateFromSerial(s).getUTCMonth() + 1;
  },

  DAY: (args, ctx) => {
    const s = argNum(args, 0, ctx, 0);
    if (isError(s)) return s;
    return dateFromSerial(s).getUTCDate();
  },

  HOUR: (args, ctx) => {
    const s = argNum(args, 0, ctx, 0);
    if (isError(s)) return s;
    return Math.floor(((s % 1) + 1) % 1 * 24);
  },

  MINUTE: (args, ctx) => {
    const s = argNum(args, 0, ctx, 0);
    if (isError(s)) return s;
    const frac = ((s % 1) + 1) % 1;
    return Math.floor(frac * 1440) % 60;
  },

  SECOND: (args, ctx) => {
    const s = argNum(args, 0, ctx, 0);
    if (isError(s)) return s;
    const frac = ((s % 1) + 1) % 1;
    return Math.round(frac * 86400) % 60;
  },

  DAYS: (args, ctx) => {
    const end = argNum(args, 0, ctx, 0);
    if (isError(end)) return end;
    const start = argNum(args, 1, ctx, 0);
    if (isError(start)) return start;
    return Math.round(end - start);
  },

  /* ---------- Прочее ---------- */
  NA: () => err(ERRORS.NA),
  ISERROR: (args, ctx) => isError(argValue(args, 0, ctx)),
  ISERR: (args, ctx) => {
    const v = argValue(args, 0, ctx);
    return isError(v) && v.__error !== ERRORS.NA;
  },
  ISNA: (args, ctx) => {
    const v = argValue(args, 0, ctx);
    return isError(v) && v.__error === ERRORS.NA;
  },
  ISBLANK: (args, ctx) => {
    const v = argValue(args, 0, ctx);
    return v === null || v === undefined || v === '';
  },
  ISNUMBER: (args, ctx) => typeof argValue(args, 0, ctx) === 'number',
  ISTEXT: (args, ctx) => typeof argValue(args, 0, ctx) === 'string',
  ISLOGICAL: (args, ctx) => typeof argValue(args, 0, ctx) === 'boolean',
  ISEVEN: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    return isError(n) ? n : Math.floor(Math.abs(n)) % 2 === 0;
  },
  ISODD: (args, ctx) => {
    const n = argNum(args, 0, ctx, 0);
    return isError(n) ? n : Math.floor(Math.abs(n)) % 2 === 1;
  },
  N: (args, ctx) => {
    const v = argValue(args, 0, ctx);
    if (isError(v)) return v;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    return 0;
  },
};

/** Собирает значения диапазона в плоский список. */
function collectRange(
  ctx: EvalContext,
  sheet: string,
  bounds: { c1: number; r1: number; c2: number; r2: number }
): (CellPrimitive | FormulaError)[] {
  const values: (CellPrimitive | FormulaError)[] = [];
  for (let r = bounds.r1; r <= bounds.r2; r++) {
    for (let c = bounds.c1; c <= bounds.c2; c++) {
      values.push(cellValue(ctx, sheet, c, r));
    }
  }
  return values;
}

/* ---------- Даты ---------- */

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

function excelSerialFromDate(d: Date): number {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round((utc - EXCEL_EPOCH) / 86400000);
}

function dateFromSerial(serial: number): Date {
  return new Date(EXCEL_EPOCH + Math.floor(serial) * 86400000);
}

/* ---------- Форматирование TEXT ---------- */

/** Упрощённая поддержка форматов Excel для функции TEXT и отображения чисел. */
export function formatValue(v: CellPrimitive | FormulaError, fmt: string): string {
  if (isError(v)) return v.__error;
  if (v === null || v === undefined) return '';

  const f = fmt.trim();
  if (!f) return toText(v);

  // Проценты
  if (f.includes('%')) {
    const n = toNumber(v) ?? 0;
    const decimals = (f.split('.')[1] || '').replace(/[^0#]/g, '').length;
    return `${(n * 100).toFixed(decimals)}%`;
  }

  // Даты
  if (/[ymdhs]/i.test(f) && !/[#0]/.test(f)) {
    const serial = toNumber(v);
    if (serial !== null && serial > 0) {
      const d = dateFromSerial(serial);
      const pad = (n: number, len = 2) => String(n).padStart(len, '0');
      const months = [
        'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
        'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря',
      ];
      const lower = f.toLowerCase();
      if (lower.includes('mmmm')) return months[d.getUTCMonth()];
      if (lower.includes('mmm')) return months[d.getUTCMonth()].slice(0, 3);
      if (lower.includes('mm') && lower.includes('/')) {
        return `${pad(d.getUTCMonth() + 1)}/${pad(d.getUTCDate())}/${d.getUTCFullYear()}`;
      }
      if (lower.includes('г')) {
        return `${pad(d.getUTCDate())}.${pad(d.getUTCMonth() + 1)}.${d.getUTCFullYear()} г.`;
      }
      if (lower.includes('-')) {
        return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
      }
      return `${pad(d.getUTCDate())}.${pad(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`;
    }
  }

  // Денежный / числовой формат
  const num = toNumber(v);
  if (num !== null) {
    const dot = f.indexOf('.');
    const decimals = dot === -1 ? 0 : (f.slice(dot + 1).match(/[0#]/g) || []).length;
    let out = num.toFixed(decimals);

    // Разделитель тысяч: в формате это запятая (#,##0) либо пробел (# ##0)
    if (f.includes(',') || /[#0]\s+[#0]/.test(f)) {
      const [intPart, fracPart] = out.split('.');
      const sign = intPart.startsWith('-') ? '-' : '';
      const digits = sign ? intPart.slice(1) : intPart;
      const grouped = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
      out = sign + grouped + (fracPart ? `.${fracPart}` : '');
    }

    // Валюта
    const currency = f.match(/[₽$€£]/);
    if (currency) {
      return f.indexOf(currency[0]) < f.search(/[#0]/) ? `${currency[0]}${out}` : `${out} ${currency[0]}`;
    }
    return out;
  }

  return toText(v);
}

/* ============================================================
   Вычисление узлов
   ============================================================ */

export function evalNode(node: Node, ctx: EvalContext): CellPrimitive | FormulaError {
  switch (node.kind) {
    case 'number':
      return node.value;
    case 'string':
      return node.value;
    case 'bool':
      return node.value;
    case 'error':
      return err(node.value);

    case 'ref': {
      const addr = node.addr;
      const v = cellValue(ctx, addr.sheet, addr.col, addr.row);
      return v === undefined ? null : v;
    }

    case 'range': {
      // В скалярном контексте диапазон даёт значение первой ячейки
      const v = cellValue(ctx, node.start.sheet, node.start.col, node.start.row);
      return v === undefined ? null : v;
    }

    case 'name': {
      const upper = node.value.toUpperCase();
      if (upper === 'TRUE') return true;
      if (upper === 'FALSE') return false;
      // Ссылка на весь столбец/строку без номера — не поддерживается как значение
      return err(ERRORS.NAME);
    }

    case 'unary': {
      const v = evalNode(node.operand, ctx);
      if (isError(v)) return v;
      const n = toNumberStrict(v);
      if (isError(n)) return n;
      return node.op === '-' ? -n : n;
    }

    case 'percent': {
      const v = evalNode(node.operand, ctx);
      if (isError(v)) return v;
      const n = toNumberStrict(v);
      if (isError(n)) return n;
      return n / 100;
    }

    case 'binary': {
      const op = node.op;

      // Логика И/ИЛИ через операторы не поддерживается, только сравнения
      if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
        // Для диапазонов сравнение даёт поэлементный результат — в скалярном
        // контексте используем первое значение.
        const l = evalNode(node.left, ctx);
        const r = evalNode(node.right, ctx);
        if (isError(l) && isError(r)) return l;
        const cmp = compareValues(isError(l) ? null : l, isError(r) ? null : r);
        switch (op) {
          case '=':
            return cmp === 0;
          case '<>':
            return cmp !== 0;
          case '<':
            return cmp < 0;
          case '>':
            return cmp > 0;
          case '<=':
            return cmp <= 0;
          case '>=':
            return cmp >= 0;
        }
      }

      // Поэлементная арифметика над диапазонами (как формулы массива Excel):
      // =SUM(B2:B4*C2:C4) перемножает ячейки попарно.
      if (op !== '&' && (node.left.kind === 'range' || node.right.kind === 'range')) {
        const leftVals = node.left.kind === 'range' ? flattenValues(node.left, ctx) : null;
        const rightVals = node.right.kind === 'range' ? flattenValues(node.right, ctx) : null;
        const len = Math.max(leftVals?.length ?? 1, rightVals?.length ?? 1);

        // Если длины не совпадают и обе заданы — это ошибка размерности
        if (leftVals && rightVals && leftVals.length !== rightVals.length) {
          return err(ERRORS.VALUE);
        }

        const results: number[] = [];
        for (let k = 0; k < len; k++) {
          const lv = leftVals ? leftVals[k] : evalNode(node.left, ctx);
          const rv = rightVals ? rightVals[k] : evalNode(node.right, ctx);
          if (isError(lv)) return lv;
          if (isError(rv)) return rv;
          const ln = toNumberStrict(lv);
          if (isError(ln)) return ln;
          const rn = toNumberStrict(rv);
          if (isError(rn)) return rn;

          let value: number;
          switch (op) {
            case '+':
              value = ln + rn;
              break;
            case '-':
              value = ln - rn;
              break;
            case '*':
              value = ln * rn;
              break;
            case '/':
              if (rn === 0) return err(ERRORS.DIV0);
              value = ln / rn;
              break;
            case '^':
              value = Math.pow(ln, rn);
              break;
            default:
              return err(ERRORS.VALUE);
          }
          if (!Number.isFinite(value)) return err(ERRORS.NUM);
          results.push(Math.round(value * 1e10) / 1e10);
        }

        // Одиночный результат возвращаем скаляром, несколько — как массив,
        // который агрегатные функции (SUM, AVERAGE, ...) разворачивают.
        return results.length === 1 ? results[0] : (results as unknown as CellPrimitive);
      }

      const l = evalNode(node.left, ctx);
      if (isError(l)) return l;
      const r = evalNode(node.right, ctx);
      if (isError(r)) return r;

      if (op === '&') {
        return `${toText(l)}${toText(r)}`;
      }

      const ln = toNumberStrict(l);
      if (isError(ln)) return ln;
      const rn = toNumberStrict(r);
      if (isError(rn)) return rn;

      let result: number;
      switch (op) {
        case '+':
          result = ln + rn;
          break;
        case '-':
          result = ln - rn;
          break;
        case '*':
          result = ln * rn;
          break;
        case '/':
          if (rn === 0) return err(ERRORS.DIV0);
          result = ln / rn;
          break;
        case '^': {
          result = Math.pow(ln, rn);
          if (!Number.isFinite(result)) return err(ERRORS.NUM);
          break;
        }
        default:
          return err(ERRORS.VALUE);
      }
      if (!Number.isFinite(result)) return err(ERRORS.NUM);
      // Округление до 10 знаков убирает артефакты плавающей точки (0.1+0.2)
      return Math.round(result * 1e10) / 1e10;
    }

    case 'call': {
      const impl = FUNCTIONS[node.name];
      if (!impl) return err(ERRORS.NAME);
      try {
        return impl(node.args, ctx);
      } catch {
        return err(ERRORS.VALUE);
      }
    }

    default:
      return err(ERRORS.NAME);
  }
}

/* ============================================================
   Публичный API
   ============================================================ */

/** Вычисляет формулу в контексте книги. */
export function evaluateFormula(formula: string, ctx: EvalContext): CellPrimitive | FormulaError {
  let ast: Node;
  try {
    ast = parseFormula(formula);
  } catch {
    return err(ERRORS.NAME);
  }
  return evalNode(ast, ctx);
}

/** Создаёт контекст вычисления для книги. */
export function createContext(book: Workbook, sheet: string): EvalContext {
  return {
    book,
    sheet,
    cache: new Map(),
    stack: new Set(),
    budget: { steps: 0 },
  };
}

/**
 * Вычисляет значение ячейки для отображения.
 * Строки, начинающиеся с '=', считаются формулами; иначе — литералами.
 */
export function computeCellDisplay(
  book: Workbook,
  sheet: string,
  row: number,
  col: number,
  ctx?: EvalContext
): string {
  const context = ctx ?? createContext(book, sheet);
  const value = cellValue(context, sheet, col, row);
  if (isError(value)) return value.__error;
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') {
    if (Number.isInteger(value)) return String(value);
    return String(Math.round(value * 1e10) / 1e10);
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

/** Список поддерживаемых имён функций — для автодополнения. */
export const FUNCTION_NAMES = Object.keys(FUNCTIONS).sort();

/** Русские названия функций (локализованный ввод) -> канонические английские. */
export const RU_FUNCTION_ALIASES: Record<string, string> = {
  СУММ: 'SUM',
  СУММЕСЛИ: 'SUMIF',
  СУММЕСЛИМН: 'SUMIFS',
  СУММПРОИЗВ: 'SUMPRODUCT',
  ПРОИЗВЕД: 'PRODUCT',
  ABS: 'ABS',
  ОКРУГЛ: 'ROUND',
  ОКРУГЛВВЕРХ: 'ROUNDUP',
  ОКРУГЛВНИЗ: 'ROUNDDOWN',
  ЦЕЛОЕ: 'INT',
  ОТБР: 'TRUNC',
  ОСТАТ: 'MOD',
  СТЕПЕНЬ: 'POWER',
  КОРЕНЬ: 'SQRT',
  EXP: 'EXP',
  LN: 'LN',
  LOG: 'LOG',
  LOG10: 'LOG10',
  ЗНАК: 'SIGN',
  ПИ: 'PI',
  СЛЧИС: 'RAND',
  СЛУЧМЕЖДУ: 'RANDBETWEEN',
  ОКРУГЛТ: 'MROUND',
  СРЗНАЧ: 'AVERAGE',
  СРЗНАЧЕСЛИ: 'AVERAGEIF',
  СЧЁТ: 'COUNT',
  СЧЕТ: 'COUNT',
  СЧЁТЗ: 'COUNTA',
  СЧЕТЗ: 'COUNTA',
  СЧИТАТЬПУСТОТЫ: 'COUNTBLANK',
  СЧЁТЕСЛИ: 'COUNTIF',
  СЧЕТЕСЛИ: 'COUNTIF',
  СЧЁТЕСЛИМН: 'COUNTIFS',
  СЧЕТЕСЛИМН: 'COUNTIFS',
  МИН: 'MIN',
  МАКС: 'MAX',
  МЕДИАНА: 'MEDIAN',
  МОДА: 'MODE',
  СТАНДОТКЛОН: 'STDEV',
  ДИСП: 'VAR',
  НАИБОЛЬШИЙ: 'LARGE',
  НАИМЕНЬШИЙ: 'SMALL',
  РАНГ: 'RANK',
  ПЕРСЕНТИЛЬ: 'PERCENTILE',
  ЕСЛИ: 'IF',
  ЕСЛИОШИБКА: 'IFERROR',
  ЕСЛИНД: 'IFNA',
  И: 'AND',
  ИЛИ: 'OR',
  НЕ: 'NOT',
  ИСТИНА: 'TRUE',
  ЛОЖЬ: 'FALSE',
  СЦЕП: 'CONCAT',
  СЦЕПИТЬ: 'CONCATENATE',
  ОБЪЕДИНИТЬ: 'TEXTJOIN',
  ДЛСТР: 'LEN',
  ЛЕВСИМВ: 'LEFT',
  ПРАВСИМВ: 'RIGHT',
  ПСТР: 'MID',
  ПРОПИСН: 'UPPER',
  СТРОЧН: 'LOWER',
  ПРОПНАЧ: 'PROPER',
  СЖПРОБЕЛЫ: 'TRIM',
  ПЕЧСИМВ: 'CLEAN',
  ПОВТОР: 'REPT',
  ПОДСТАВИТЬ: 'SUBSTITUTE',
  ЗАМЕНИТЬ: 'REPLACE',
  НАЙТИ: 'FIND',
  ПОИСК: 'SEARCH',
  ЗНАЧЕН: 'VALUE',
  ТЕКСТ: 'TEXT',
  СИМВОЛ: 'CHAR',
  КОДСИМВ: 'CODE',
  СОВПАД: 'EXACT',
  ВПР: 'VLOOKUP',
  ГПР: 'HLOOKUP',
  ИНДЕКС: 'INDEX',
  ПОИСКПОЗ: 'MATCH',
  ПРОСМОТРX: 'XLOOKUP',
  ВЫБОР: 'CHOOSE',
  СТРОКА: 'ROW',
  СТОЛБЕЦ: 'COLUMN',
  ЧСТРОК: 'ROWS',
  ЧИСЛСТОЛБ: 'COLUMNS',
  СЕГОДНЯ: 'TODAY',
  ТДАТА: 'NOW',
  ДАТА: 'DATE',
  ГОД: 'YEAR',
  МЕСЯЦ: 'MONTH',
  ДЕНЬ: 'DAY',
  ЧАС: 'HOUR',
  МИНУТЫ: 'MINUTE',
  СЕКУНДЫ: 'SECOND',
  ДНИ: 'DAYS',
  НД: 'NA',
  ЕОШИБКА: 'ISERROR',
  ЕОШ: 'ISERR',
  ЕНД: 'ISNA',
  ЕПУСТО: 'ISBLANK',
  ЕЧИСЛО: 'ISNUMBER',
  ЕТЕКСТ: 'ISTEXT',
  ЕЛОГИЧ: 'ISLOGICAL',
  ЧЁТН: 'ISEVEN',
  ЧЕТН: 'ISEVEN',
  НЕЧЁТН: 'ISODD',
  НЕЧЕТН: 'ISODD',
};

/** Переводит локализованное имя функции в каноническое. */
export function canonicalFunctionName(name: string): string {
  const upper = name.toUpperCase();
  return RU_FUNCTION_ALIASES[upper] ?? upper;
}

/**
 * Нормализует формулу для вычисления: локализованные имена функций
 * заменяются каноническими. Разделитель аргументов ';' лексер принимает
 * наравне с ','.
 */
export function normalizeFormula(formula: string): string {
  // Имена функций (за ними следует скобка)
  return formula.replace(
    /([A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9_.]*)\s*\(/g,
    (match, name: string) => {
      const canonical = canonicalFunctionName(name);
      return FUNCTIONS[canonical] ? `${canonical}(` : match;
    }
  // Логические константы ИСТИНА/ЛОЖЬ без скобок
  ).replace(/(?<![A-Za-zА-Яа-яЁё0-9_.])(ИСТИНА|ЛОЖЬ)(?![A-Za-zА-Яа-яЁё0-9_.(])/gi, (m) =>
    canonicalFunctionName(m)
  );
}

/**
 * Расстояние Левенштейна. Имена функций короткие, поэтому квадрат
 * длины строки здесь несущественен.
 */
function editDistance(a: string, b: string): number {
  const prev = new Array(b.length + 1);
  const next = new Array(b.length + 1);
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    next[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      next[j] = Math.min(next[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    for (let j = 0; j <= b.length; j++) prev[j] = next[j];
  }
  return prev[b.length];
}

/**
 * Ближайшее известное имя функции для опечатки.
 * Порог зависит от длины: «СУМА» → SUM, а «A» ничем не заменяется.
 */
function closestFunctionName(name: string): string | null {
  const upper = name.toUpperCase();
  const limit = upper.length <= 4 ? 1 : 2;
  let best: string | null = null;
  let bestDistance = limit + 1;
  for (const candidate of FUNCTION_NAMES) {
    if (Math.abs(candidate.length - upper.length) > limit) continue;
    const distance = editDistance(upper, candidate);
    if (distance < bestDistance) {
      best = candidate;
      bestDistance = distance;
    }
  }
  return best;
}

export interface FormulaRepair {
  /** Исправленная формула, всегда начинается с «=». */
  formula: string;
  /** Что именно удалось поправить, по-русски. */
  changes: string[];
}

/**
 * Исправляет типичные ошибки формулы, не меняя её смысла:
 * русские имена функций, точку с запятой, лишние пробелы перед скобкой,
 * непарные скобки и опечатки в именах функций (СУМА → SUM).
 * Возвращает null, если формула уже корректна или правка ненадёжна.
 */
export function repairFormula(formula: string): FormulaRepair | null {
  const source = formula.trim();
  if (!source.startsWith('=')) return null;
  const changes: string[] = [];

  let body = source.slice(1).replace(/\u00A0/g, ' ');

  // Пробел между именем функции и скобкой: «SUM (A1)» → «SUM(A1)»
  if (/[A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9_.]*\s+\(/.test(body)) {
    body = body.replace(/([A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9_.]*)\s+\(/g, '$1(');
    changes.push('убран пробел перед скобкой');
  }

  // Локализованные имена и константы ИСТИНА/ЛОЖЬ
  const localized = normalizeFormula(body);
  if (localized !== body) {
    body = localized;
    changes.push('русские имена функций заменены английскими');
  }

  // Неизвестное имя перед скобкой — пробуем исправить опечатку.
  // Сначала сравниваем с русскими именами, иначе «СУМА» не дотянется до SUM:
  // каноническая форма СУММ слишком далека от опечатки.
  const aliases = Object.keys(RU_FUNCTION_ALIASES);
  body = body.replace(
    /([A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9_.]*)\(/g,
    (match, name: string) => {
      const canonical = canonicalFunctionName(name);
      if (FUNCTIONS[canonical]) return `${canonical}(`;

      const upper = name.toUpperCase();
      const limit = upper.length <= 4 ? 1 : 2;
      let bestAlias: string | null = null;
      let bestDistance = limit + 1;
      for (const alias of aliases) {
        if (Math.abs(alias.length - upper.length) > limit) continue;
        const distance = editDistance(upper, alias);
        if (distance < bestDistance) {
          bestAlias = alias;
          bestDistance = distance;
        }
      }
      if (bestAlias) {
        const resolved = RU_FUNCTION_ALIASES[bestAlias];
        changes.push(`${name} → ${resolved}`);
        return `${resolved}(`;
      }

      const guess = closestFunctionName(canonical);
      if (!guess) return match;
      changes.push(`${name} → ${guess}`);
      return `${guess}(`;
    }
  );

  // Непарные скобки вне строковых литералов
  let balance = 0;
  let inString = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      if (inString && body[i + 1] === '"') {
        i++;
        continue;
      }
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === '(') balance++;
    else if (ch === ')') balance--;
  }
  if (balance > 0) {
    body += ')'.repeat(balance);
    changes.push(balance === 1 ? 'добавлена закрывающая скобка' : `добавлено закрывающих скобок: ${balance}`);
  } else if (balance < 0) {
    let extra = -balance;
    let cut = body.length;
    while (extra > 0 && cut > 0) {
      cut--;
      if (body[cut] === ')') extra--;
      else break;
    }
    if (extra === 0) {
      body = body.slice(0, cut) + body.slice(cut - balance);
      changes.push('убраны лишние закрывающие скобки');
    }
  }

  const repaired = `=${body}`;
  if (repaired === source || changes.length === 0) return null;

  // Правка не должна сама по себе давать синтаксическую ошибку
  try {
    const ast = parseFormula(repaired);
    if (ast.kind === 'error') return null;
  } catch {
    return null;
  }
  return { formula: repaired, changes };
}

/**
 * Проходит по формуле и вызывает колбэк для каждого разделителя аргументов.
 * Запятые и точки с запятой внутри строк и внутри диапазонов не трогает:
 * "a,b" и A1:B2 остаются как есть.
 */
function mapArgumentSeparators(
  formula: string,
  replace: (separator: ';' | ',') => string
): string {
  let out = '';
  let inString = false;
  let depth = 0;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"') {
      out += ch;
      if (inString && formula[i + 1] === '"') {
        out += formula[++i];
        continue;
      }
      inString = !inString;
      continue;
    }
    if (inString) {
      out += ch;
      continue;
    }
    if (ch === '(') depth++;
    else if (ch === ')' && depth > 0) depth--;
    if (depth > 0 && (ch === ';' || ch === ',')) out += replace(ch);
    else out += ch;
  }
  return out;
}

/**
 * Формула в том виде, в каком её хранит Excel в файле .xlsx:
 * английские имена функций и запятая между аргументами.
 * В редакторе при этом можно писать по-русски и через точку с запятой.
 */
export function toExcelFormula(formula: string): string {
  const normalized = normalizeFormula(formula);
  return mapArgumentSeparators(normalized, () => ',');
}

/**
 * Формула из файла Excel для редактора: запятые-разделители аргументов
 * возвращаются к точке с запятой, к которой привык русский Excel.
 * Содержимое строк и разделитель диапазона «:» не меняются.
 */
export function fromExcelFormula(formula: string): string {
  return mapArgumentSeparators(formula, () => ';');
}
