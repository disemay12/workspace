/**
 * Проверка полного цикла «AI-формула»:
 *   запрос -> ответ модели -> разбор -> подтверждение -> запись в ячейку.
 *
 * Повторяет логику parseAIFormula и applyAIAction из ExcelEditor на чистых
 * данных, чтобы убедиться, что формула действительно попадает в ячейку,
 * а не остаётся диалогом в чате (исходная жалоба пользователя).
 */
import { normalizeFormula, computeCellDisplay, createContext, Workbook } from '../../src/lib/formulaEngine';
import { shiftFormula } from '../../src/components/editors/useFormulaMode';

const failures: string[] = [];
let passed = 0;

function check(name: string, actual: unknown, expected: unknown) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) passed++;
  else failures.push(`${name}\n     ожидалось: ${e}\n     получено:   ${a}`);
}

/* ---------- Разбор ответа модели (как в ExcelEditor) ---------- */

function parseAIFormula(answer: string, fallbackRef = 'E2'): { target: string; value: string } | null {
  const cellMatch = answer.match(/ЯЧЕЙКА\s*:\s*([A-Za-zА-Яа-яЁё0-9_$!:' ]+)/i);
  const formulaMatch = answer.match(/ФОРМУЛА\s*:\s*(.+)/i);
  if (!formulaMatch) return null;

  let value = formulaMatch[1].trim();
  value = value.replace(/^[`"']+|[`"']+$/g, '').trim();
  value = value.split(/\s+#/)[0].trim();
  if (value && !value.startsWith('=')) value = `=${value}`;

  const target = cellMatch?.[1]?.trim().replace(/\s/g, '') || fallbackRef;
  return { target, value };
}

/* ---------- Разбор адреса ---------- */

function parseTarget(target: string, defaultSheet: string) {
  let sheet = defaultSheet;
  let ref = target.trim();

  const bang = ref.lastIndexOf('!');
  if (bang !== -1) {
    sheet = ref.slice(0, bang).replace(/^'|'$/g, '');
    ref = ref.slice(bang + 1);
  }

  const parseOne = (s: string) => {
    const m = s.replace(/\$/g, '').match(/^([A-Za-z]{1,3})(\d{1,7})$/);
    if (!m) return null;
    let col = 0;
    const letters = m[1].toUpperCase();
    for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
    return { row: parseInt(m[2], 10) - 1, col: col - 1 };
  };

  const parts = ref.split(':');
  const start = parseOne(parts[0]);
  const end = parts.length > 1 ? parseOne(parts[1]) : start;
  if (!start || !end) return null;

  return {
    sheet,
    start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
    end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
  };
}

/** Применяет действие к книге — как applyAIAction. */
function applyAction(
  book: Workbook,
  sheetName: string,
  target: string,
  value: string
): { ok: boolean; sheet: string } {
  const parsed = parseTarget(target, sheetName);
  if (!parsed) return { ok: false, sheet: sheetName };

  const sheet = book[parsed.sheet] ? parsed.sheet : sheetName;
  const grid = book[sheet];
  const normalized = normalizeFormula(value);

  while (grid.length <= parsed.end.row) {
    grid.push(new Array(grid[0]?.length ?? 16).fill(''));
  }
  for (const row of grid) {
    while (row.length <= parsed.end.col) row.push('');
  }

  for (let r = parsed.start.row; r <= parsed.end.row; r++) {
    for (let c = parsed.start.col; c <= parsed.end.col; c++) {
      grid[r][c] = shiftFormula(normalized, r - parsed.start.row, c - parsed.start.col);
    }
  }
  return { ok: true, sheet };
}

/* ---------- Тестовые данные ---------- */

function makeBook(): Workbook {
  return {
    Данные: [
      ['Товар', 'Кол-во', 'Цена', 'Оплачено'],
      ['Яблоки', 10, 50, 'Да'],
      ['Груши', 5, 80, 'Нет'],
      ['Сливы', 20, 30, 'Да'],
    ],
  };
}

/* ---------- 1. Ответ модели разбирается в действие ---------- */

check(
  'простой ответ',
  parseAIFormula('ЯЧЕЙКА: E2\nФОРМУЛА: =SUM(B2:B4)'),
  { target: 'E2', value: '=SUM(B2:B4)' }
);

check(
  'русская функция',
  parseAIFormula('ЯЧЕЙКА: E2\nФОРМУЛА: =СУММ(B2:B4)'),
  { target: 'E2', value: '=СУММ(B2:B4)' }
);

check(
  'ответ с пояснением после формулы отбрасывает хвост',
  parseAIFormula('ЯЧЕЙКА: E2\nФОРМУЛА: =SUM(B2:B4)  # суммирует количество'),
  { target: 'E2', value: '=SUM(B2:B4)' }
);

check('ответ без формулы не даёт действия', parseAIFormula('Формула не нужна'), null);

/* ---------- 2. Формула попадает в указанную ячейку ---------- */

{
  const book = makeBook();
  const r = applyAction(book, 'Данные', 'E2', '=SUM(B2:B4)');
  check('действие применено', r.ok, true);

  const value = computeCellDisplay(book, 'Данные', 1, 4, createContext(book, 'Данные'));
  check('в E2 записана сумма 35', value, '35');
}

/* ---------- 3. AI-формула с условием ---------- */

{
  const book = makeBook();
  applyAction(book, 'Данные', 'E3', '=SUMIF(D2:D4;"Да";B2:B4)');
  const value = computeCellDisplay(book, 'Данные', 2, 4, createContext(book, 'Данные'));
  check('СУММЕСЛИ по условию = 30', value, '30');
}

/* ---------- 4. Заполнение диапазона со сдвигом ссылок ---------- */

{
  const book = makeBook();
  // Формула построчной суммы: B*C, протянутая на строки 2..4
  applyAction(book, 'Данные', 'E2:E4', '=B2*C2');

  const ctx = createContext(book, 'Данные');
  check('E2 = 10*50', computeCellDisplay(book, 'Данные', 1, 4, ctx), '500');
  check('E3 = 5*80', computeCellDisplay(book, 'Данные', 2, 4, ctx), '400');
  check('E4 = 20*30', computeCellDisplay(book, 'Данные', 3, 4, ctx), '600');
}

/* ---------- 5. Ячейка за пределами листа расширяется ---------- */

{
  const book = makeBook();
  const r = applyAction(book, 'Данные', 'H20', '=1+1');
  check('адрес за границами принят', r.ok, true);
  check('лист расширен', book['Данные'].length >= 20, true);
  check(
    'значение записано в H20',
    computeCellDisplay(book, 'Данные', 19, 7, createContext(book, 'Данные')),
    '2'
  );
}

/* ---------- 6. Ссылка на другой лист в ответе AI ---------- */

{
  const book = makeBook();
  book['Справочник'] = [
    ['Товар', 'Цена'],
    ['Яблоки', 50],
    ['Груши', 80],
    ['Сливы', 30],
  ];
  applyAction(book, 'Данные', 'E2', '=VLOOKUP("Груши";Справочник!A2:B4;2;FALSE)');
  check(
    'ВПР по другому листу = 80',
    computeCellDisplay(book, 'Данные', 1, 4, createContext(book, 'Данные')),
    '80'
  );
}

/* ---------- 7. Некорректный адрес не ломает таблицу ---------- */

{
  const book = makeBook();
  const before = JSON.stringify(book);
  const r = applyAction(book, 'Данные', 'не-адрес', '=SUM(A1:A2)');
  check('некорректный адрес отвергнут', r.ok, false);
  check('книга не изменена', JSON.stringify(book), before);
}

/* ---------- 8. Ошибка в формуле отображается как код ошибки ---------- */

{
  const book = makeBook();
  applyAction(book, 'Данные', 'E2', '=1/0');
  check(
    'деление на ноль даёт #DIV/0!',
    computeCellDisplay(book, 'Данные', 1, 4, createContext(book, 'Данные')),
    '#DIV/0!'
  );
}

console.log('='.repeat(72));
console.log('  ПРОВЕРКА ЦИКЛА «AI-ФОРМУЛА -> ЯЧЕЙКА»');
console.log('='.repeat(72));
console.log(`  Пройдено: ${passed}, провалов: ${failures.length}`);
if (failures.length > 0) {
  console.log('-'.repeat(72));
  for (const f of failures) console.log(`   - ${f}`);
}
console.log('='.repeat(72));
console.log(failures.length === 0 ? '  ИТОГ: ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ' : '  ИТОГ: ЕСТЬ ПРОВАЛЫ');
process.exit(failures.length === 0 ? 0 : 1);
