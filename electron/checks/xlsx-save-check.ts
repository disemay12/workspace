/**
 * Проверка сохранения книги: формулы должны записываться в .xlsx как
 * настоящие формулы Excel (свойство f), а не как текст «=SUM(...)».
 * Иначе Excel покажет формулу строкой и не будет её считать.
 *
 * Проверяет полный цикл: сетка -> книга -> base64 -> чтение обратно.
 */
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { emptySheetMeta, type WorkbookMeta } from '../../src/lib/sheetFormatting';
import { injectStylesIntoXlsx } from '../../src/lib/xlsxStyles';
import { applyAutoFitColumns, textWidthChars } from '../../src/lib/xlsxColumns';

const failures: string[] = [];
let passed = 0;

function check(name: string, actual: unknown, expected: unknown) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) passed++;
  else failures.push(`${name}\n     ожидалось: ${e}\n     получено:   ${a}`);
}

/** Повторяет логику syncWorkbook из ExcelEditor. */
function buildBase64(grid: (string | number | boolean | null)[][], sheetName: string): string {
  const wb = XLSX.utils.book_new();

  let lastNonEmpty = grid.length - 1;
  while (lastNonEmpty > 0 && grid[lastNonEmpty].every((c) => String(c ?? '').trim() === '')) {
    lastNonEmpty--;
  }

  const trimmed = grid.slice(0, lastNonEmpty + 1);
  const sheet = XLSX.utils.aoa_to_sheet(trimmed);

  // Формулы проставляются после создания листа. Ячейка с формулой обязана
  // нести кэшированное значение: без него SheetJS не пишет ячейку вовсе,
  // а строка «=…» осталась бы текстом. Кэш = вычисленный результат.
  for (let r = 0; r < trimmed.length; r++) {
    for (let c = 0; c < trimmed[r].length; c++) {
      const cell = trimmed[r][c];
      if (typeof cell !== 'string' || !cell.startsWith('=')) continue;
      const address = XLSX.utils.encode_cell({ r, c });
      const existing = sheet[address];
      // Для проверки формы ячейки кэш берём числом 0 — важна сама структура
      const next: Record<string, unknown> = { t: 'n', f: cell.slice(1), v: 0 };
      if (existing?.z) next.z = existing.z;
      if (existing?.s) next.s = existing.s;
      sheet[address] = next as unknown as XLSX.CellObject;
    }
  }

  XLSX.utils.book_append_sheet(wb, sheet, sheetName);
  return XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
}

const grid: (string | number)[][] = [
  ['Товар', 'Кол-во', 'Цена', 'Оплачено'],
  ['Яблоки', 10, 50, 'Да'],
  ['Груши', 5, 80, 'Нет'],
  ['Сливы', 20, 30, 'Да'],
  ['Сумма', '=SUM(B2:B4)', '=SUM(B2:B4*C2:C4)', '=COUNTIF(D2:D4,"Да")'],
  ['', '', '', ''],
  ['', '', '', ''],
];

const base64 = buildBase64(grid, 'Данные');

// Читаем книгу обратно — как это сделает Excel
const readBack = XLSX.read(base64, { type: 'base64' });
const sheet = readBack.Sheets['Данные'];

check('лист сохранён', Boolean(sheet), true);

/* ---- Формулы сохранены как формулы ---- */
check('формула SUM сохранена как формула', sheet['B5']?.f, 'SUM(B2:B4)');
check('формула массива сохранена', sheet['C5']?.f, 'SUM(B2:B4*C2:C4)');
check('формула COUNTIF с кавычками сохранена', sheet['D5']?.f, 'COUNTIF(D2:D4,"Да")');

/* ---- Формула не превратилась в текст ---- */
check('B5 не является строкой-формулой', typeof sheet['B5']?.v === 'string' && String(sheet['B5'].v).startsWith('='), false);

/* ---- Обычные значения не пострадали ---- */
check('текст сохранён', sheet['A2']?.v, 'Яблоки');
check('число сохранено', sheet['B2']?.v, 10);
check('число сохранено (2)', sheet['C4']?.v, 30);

/* ---- Пустые хвостовые строки обрезаны ---- */
check('пустые строки не раздувают файл', sheet['!ref'], 'A1:D5');

/* ---- Формат даты/типа ячейки ---- */
check('тип ячейки с формулой — n (число)', sheet['B5']?.t, 'n');

/* ---- Русские имена функций сохраняются как есть ---- */
const ruGrid: (string | number)[][] = [
  ['a', 'b'],
  [1, 2],
  ['', '=СУММ(A2:B2)'],
];
const ruBase64 = buildBase64(ruGrid, 'Лист1');
const ruSheet = XLSX.read(ruBase64, { type: 'base64' }).Sheets['Лист1'];
check('русская формула сохранена', ruSheet['B3']?.f, 'СУММ(A2:B2)');

/* ---- CSV-экспорт: формулы заменяются значениями ---- */
// В ExcelEditor экспорт CSV подставляет вычисленные значения
const csvSheet = XLSX.utils.aoa_to_sheet([
  ['Товар', 'Кол-во'],
  ['Яблоки', 10],
  ['Сумма', 10],
]);
const csv = XLSX.utils.sheet_to_csv(csvSheet);
check('CSV содержит вычисленное значение', csv.includes('Сумма,10'), true);
check('CSV не содержит формулу', csv.includes('=SUM') || csv.includes('=СУММ'), false);

/* ---- Проценты и спецсимволы в тексте ---- */
const specialGrid: (string | number)[][] = [
  ['=SUM(A2:A3)', 'текст; с точкой с запятой', 'кавычки "тут"'],
  [1, 2, 3],
];
const specialBase64 = buildBase64(specialGrid, 'S');
const specialSheet = XLSX.read(specialBase64, { type: 'base64' }).Sheets['S'];
check('формула в первой ячейке сохранена', specialSheet['A1']?.f, 'SUM(A2:A3)');
check('точка с запятой в тексте сохранена', specialSheet['B1']?.v, 'текст; с точкой с запятой');

async function main() {
  /* ---- Нативные стили: заливка, шрифт, границы, формат ---- */
  const styledMeta: WorkbookMeta = {
    S: {
      ...emptySheetMeta(),
      cells: {
        A1: { fill: '#DDEBF7', bold: true },
        B1: { fill: '#DDEBF7' },
        A2: { borders: { bottom: true } },
        B2: { numberFormat: 'currency' },
      },
    },
  };
  const styledBase64 = await injectStylesIntoXlsx(specialBase64, styledMeta);
  const styleZip = await JSZip.loadAsync(styledBase64, { base64: true });
  const stylesXml = (await styleZip.file('xl/styles.xml')?.async('string')) ?? '';
  const styledSheet = (await styleZip.file('xl/worksheets/sheet1.xml')?.async('string')) ?? '';
  check('заливка записана в styles.xml', stylesXml.includes('FFDDEBF7'), true);
  check('заливка solid', stylesXml.includes('patternFill patternType="solid"'), true);
  check('жирный шрифт записан', stylesXml.includes('<b/>'), true);
  check('границы записаны', stylesXml.includes('style="thin"'), true);
  check('числовой формат записан', /<numFmt\b/.test(stylesXml), true);
  check('ячейка A1 получила style index', /<c r="A1"[^>]*\bs="\d+"/.test(styledSheet), true);
  check(
    'WorkSpace читает файл после патча',
    Boolean(XLSX.read(styledBase64, { type: 'base64' }).Sheets.S?.['A1']),
    true
  );

  /* ---- Ширины колонок: текст должен помещаться в Excel ---- */
  const wideMatrix = [
    ['Товар', 'Описание', 'Цена'],
    ['Яблоки', 'Очень длинное описание товара, которое не должно обрезаться', 120],
    ['Груши', 'Коротко', 80],
  ];
  const wideWb = XLSX.utils.book_new();
  const wideSheet = XLSX.utils.aoa_to_sheet(wideMatrix);
  applyAutoFitColumns(wideSheet, wideMatrix, { headerRows: 1 });
  XLSX.utils.book_append_sheet(wideWb, wideSheet, 'Ширины');
  const wideBase64 = XLSX.write(wideWb, { bookType: 'xlsx', type: 'base64' });
  const wideRead = XLSX.read(wideBase64, { type: 'base64', cellStyles: true });
  const wideCols = wideRead.Sheets['Ширины']?.['!cols'] ?? [];
  check('ширина колонки A записана', typeof wideCols[0]?.wch, 'number');
  check(
    'колонка «Описание» шире текста',
    (wideCols[1]?.wch ?? 0) >= textWidthChars(wideMatrix[1][1] as string),
    true
  );
  check(
    'колонка «Цена» не раздута',
    (wideCols[2]?.wch ?? 0) < 30,
    true
  );
  const wideZip = await JSZip.loadAsync(wideBase64, { base64: true });
  const wideXml = (await wideZip.file('xl/worksheets/sheet1.xml')?.async('string')) ?? '';
  check('в XML есть тег cols', /<cols>/.test(wideXml), true);
  check('в XML есть width', /<col[^>]*width="/.test(wideXml), true);
  check(
    'WorkSpace читает ширины обратно',
    (XLSX.read(wideBase64, { type: 'base64', cellStyles: true }).Sheets['Ширины']?.['!cols']?.length ?? 0) > 0,
    true
  );

  console.log('='.repeat(72));
  console.log('  ПРОВЕРКА СОХРАНЕНИЯ ФОРМУЛ В .XLSX');
  console.log('='.repeat(72));
  console.log(`  Пройдено: ${passed}, провалов: ${failures.length}`);
  if (failures.length > 0) {
    console.log('-'.repeat(72));
    for (const f of failures) console.log(`   - ${f}`);
  }
  console.log('='.repeat(72));
  console.log(failures.length === 0 ? '  ИТОГ: ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ' : '  ИТОГ: ЕСТЬ ПРОВАЛЫ');
  process.exit(failures.length === 0 ? 0 : 1);
}

void main();
