/**
 * Проверка движка формул: широкий набор функций, диапазоны, ссылки на другие
 * листы, ошибки и циклические ссылки. Запускается как обычный Node-скрипт
 * (без Electron), т.к. движок не зависит от DOM.
 */
import {
  computeCellDisplay,
  createContext,
  evaluateFormula,
  fromExcelFormula,
  normalizeFormula,
  toExcelFormula,
  Workbook,
  isError,
  FUNCTION_NAMES,
} from '../../src/lib/formulaEngine';

const book: Workbook = {
  Лист1: [
    ['Товар', 'Кол-во', 'Цена', 'Оплачено'],
    ['Яблоки', 10, 50, 'Да'],
    ['Груши', 5, 80, 'Нет'],
    ['Сливы', 20, 30, 'Да'],
    ['Итого', '=SUM(B2:B4)', '=SUM(B2:B4*C2:C4)', '=COUNTIF(D2:D4;"Да")'],
  ],
  Лист2: [
    ['Ключ', 'Значение'],
    ['Яблоки', 111],
    ['Груши', 222],
  ],
};

const ctx = createContext(book, 'Лист1');

interface Case {
  formula: string;
  expected: string;
}

const cases: Case[] = [
  // Базовая арифметика и приоритеты
  { formula: '=1+2*3', expected: '7' },
  { formula: '=(1+2)*3', expected: '9' },
  { formula: '=2^3^2', expected: '512' },
  { formula: '=-5+3', expected: '-2' },
  { formula: '=10/4', expected: '2.5' },
  { formula: '=10/0', expected: '#DIV/0!' },
  { formula: '=50%', expected: '0.5' },
  { formula: '=0.1+0.2', expected: '0.3' },

  // Ссылки и диапазоны
  { formula: '=B2', expected: '10' },
  { formula: '=$B$2+$B$3', expected: '15' },
  { formula: '=SUM(B2:B4)', expected: '35' },
  { formula: '=SUM(B2:B4,C2:C4)', expected: '195' },
  { formula: '=SUM(B2:B4*C2:C4)', expected: '1500' },
  { formula: '=B2+B3*2', expected: '20' },

  // Агрегаты
  { formula: '=AVERAGE(B2:B4)', expected: '11.666667' },
  { formula: '=MIN(B2:B4)', expected: '5' },
  { formula: '=MAX(B2:B4)', expected: '20' },
  { formula: '=COUNT(B2:B4)', expected: '3' },
  { formula: '=COUNTA(A2:A4)', expected: '3' },
  { formula: '=MEDIAN(B2:B4)', expected: '10' },
  { formula: '=LARGE(B2:B4,1)', expected: '20' },
  { formula: '=SMALL(B2:B4,2)', expected: '10' },
  { formula: '=SMALL(B2:B4,1)', expected: '5' },
  { formula: '=LARGE(B2:B4,2)', expected: '10' },
  { formula: '=PRODUCT(B2:B3)', expected: '50' },
  { formula: '=ROUND(2.567,2)', expected: '2.57' },
  { formula: '=ROUNDUP(2.001,2)', expected: '2.01' },
  { formula: '=ROUNDDOWN(2.999,2)', expected: '2.99' },
  { formula: '=ABS(-7)', expected: '7' },
  { formula: '=MOD(10,3)', expected: '1' },
  { formula: '=INT(7.9)', expected: '7' },
  { formula: '=SQRT(16)', expected: '4' },
  { formula: '=SQRT(-1)', expected: '#NUM!' },
  { formula: '=POWER(2,10)', expected: '1024' },

  // Логика
  { formula: '=IF(B2>5,"много","мало")', expected: 'много' },
  { formula: '=IF(B3>100,"да","нет")', expected: 'нет' },
  { formula: '=AND(TRUE,TRUE)', expected: 'TRUE' },
  { formula: '=AND(TRUE,FALSE)', expected: 'FALSE' },
  { formula: '=OR(FALSE,FALSE)', expected: 'FALSE' },
  { formula: '=NOT(TRUE)', expected: 'FALSE' },
  { formula: '=IFERROR(10/0,"ошибка")', expected: 'ошибка' },
  { formula: '=IFERROR(5,"ошибка")', expected: '5' },
  { formula: '=ISNUMBER(B2)', expected: 'TRUE' },
  { formula: '=ISTEXT(A2)', expected: 'TRUE' },
  { formula: '=ISBLANK(Z99)', expected: 'TRUE' },

  // Текст
  { formula: '=LEN("привет")', expected: '6' },
  { formula: '=UPPER("abc")', expected: 'ABC' },
  { formula: '=LOWER("ABC")', expected: 'abc' },
  { formula: '=LEFT("привет",3)', expected: 'при' },
  { formula: '=RIGHT("привет",3)', expected: 'вет' },
  { formula: '=MID("привет",2,3)', expected: 'рив' },
  { formula: '=CONCATENATE(A2," - ",B2)', expected: 'Яблоки - 10' },
  { formula: '=A2&" - "&B2', expected: 'Яблоки - 10' },
  { formula: '=SUBSTITUTE("а-б-в","-","+")', expected: 'а+б+в' },
  { formula: '=TRIM("  много   пробелов  ")', expected: 'много пробелов' },
  { formula: '=REPT("ab",3)', expected: 'ababab' },
  { formula: '=FIND("б","яблоки")', expected: '2' },
  { formula: '=TEXT(1234.5,"# ##0.00")', expected: '1 234.50' },
  { formula: '=TEXT(0.25,"0%")', expected: '25%' },
  { formula: '=VALUE("42")', expected: '42' },

  // Условные агрегаты
  { formula: '=SUMIF(D2:D4,"Да",B2:B4)', expected: '30' },
  { formula: '=SUMIF(B2:B4,">5")', expected: '30' },
  { formula: '=SUMIFS(B2:B4,D2:D4,"Да")', expected: '30' },
  { formula: '=COUNTIF(D2:D4,"Да")', expected: '2' },
  { formula: '=COUNTIF(B2:B4,">9")', expected: '2' },
  { formula: '=COUNTIFS(D2:D4,"Да",B2:B4,">15")', expected: '1' },
  { formula: '=AVERAGEIF(D2:D4,"Да",B2:B4)', expected: '15' },
  { formula: '=SUMPRODUCT(B2:B4,C2:C4)', expected: '1500' },

  // Поиск
  { formula: '=VLOOKUP("Груши",Лист2!A2:B3,2,FALSE)', expected: '222' },
  { formula: '=VLOOKUP("Яблоки",Лист2!A2:B3,2,FALSE)', expected: '111' },
  { formula: '=VLOOKUP("Нет такого",Лист2!A2:B3,2,FALSE)', expected: '#N/A' },
  { formula: '=INDEX(B2:B4,2)', expected: '5' },
  { formula: '=MATCH("Груши",A2:A4,0)', expected: '2' },
  { formula: '=XLOOKUP("Яблоки",Лист2!A2:A3,Лист2!B2:B3)', expected: '111' },
  { formula: '=CHOOSE(2,"a","b","c")', expected: 'b' },
  { formula: '=ROW(B5)', expected: '5' },
  { formula: '=COLUMN(C1)', expected: '3' },
  { formula: '=ROWS(B2:B4)', expected: '3' },
  { formula: '=COLUMNS(B2:D2)', expected: '3' },

  // Русские имена функций
  { formula: '=СУММ(B2:B4)', expected: '35' },
  { formula: '=СРЗНАЧ(B2:B4)', expected: '11.666667' },
  { formula: '=ЕСЛИ(B2>5;"да";"нет")', expected: 'да' },
  { formula: '=СЧЁТЕСЛИ(D2:D4;"Да")', expected: '2' },
  { formula: '=ВПР("Груши";Лист2!A2:B3;2;ЛОЖЬ)', expected: '222' },
  { formula: '=ОКРУГЛ(2.567;2)', expected: '2.57' },
  { formula: '=МИН(B2:B4)+МАКС(B2:B4)', expected: '25' },

  // Ошибки
  { formula: '=НЕТТАКОЙФУНКЦИИ(1)', expected: '#NAME?' },
  { formula: '=1+"текст"', expected: '#VALUE!' },

  // Формулы массива и ссылки на другие листы
  { formula: '=SUM(B2:B4*C2:C4)', expected: '1500' },
  { formula: '=SUM((B2:B4)*(C2:C4))', expected: '1500' },
  { formula: '=Лист2!B2', expected: '111' },
  { formula: '=SUM(Лист2!B2:B3)', expected: '333' },
  { formula: '=Лист2!B2+Лист2!B3', expected: '333' },
];

let passed = 0;
const failures: string[] = [];

/** Прогоняет набор кейсов в контексте книги; провалы копит в failures. */
function checkCases(context: ReturnType<typeof createContext>, list: Case[], prefix = ''): void {
  for (const c of list) {
    const normalized = normalizeFormula(c.formula);
    let actual: string;
    try {
      const value = evaluateFormula(normalized, context);
      actual = isError(value)
        ? value.__error
        : value === null
          ? ''
          : typeof value === 'boolean'
            ? value
              ? 'TRUE'
              : 'FALSE'
            : String(value);
      if (typeof value === 'number' && !Number.isInteger(value)) {
        actual = String(Math.round(value * 1e6) / 1e6);
      }
    } catch (e) {
      actual = `ИСКЛЮЧЕНИЕ: ${(e as Error).message}`;
    }
    if (actual === c.expected) {
      passed++;
    } else {
      failures.push(`${prefix}${c.formula}\n     ожидалось: ${c.expected}\n     получено:   ${actual}`);
    }
  }
}

checkCases(createContext(book, 'Лист1'), cases);

/* ---------- Массивы, даты и функции из жалоб (#NAME?/#DIV/0!) ---------- */
const arrayBook: Workbook = {
  Данные: [[1, 2, '', 4, 5, 0, 7, 8, 9]],
  Сводка: [['x']],
};
const arrayCtx = createContext(arrayBook, 'Сводка');

const arrayCases: Case[] = [
  // SUMPRODUCT должен видеть диапазон поэлементно, а не первую ячейку
  { formula: '=SUMPRODUCT(Данные!A1:I1)', expected: '36' },
  { formula: '=SUMPRODUCT(--(Данные!A1:I1<>""))', expected: '8' },
  { formula: '=SUMPRODUCT(--ISNUMBER(Данные!A1:I1))', expected: '8' },
  { formula: '=SUMPRODUCT(Данные!A1:I1*Данные!A1:I1)', expected: '240' },
  { formula: '=SUM(--(Данные!A1:I1>0))', expected: '7' },
  {
    formula:
      '=SUMPRODUCT(--ISNUMBER(Данные!A1:I1);Данные!A1:I1)/SUMPRODUCT(--ISNUMBER(Данные!A1:I1))',
    expected: '4.5',
  },
  // DATEVALUE: раньше давал #NAME?, теперь это серийное число даты
  { formula: '=DATEVALUE("05.01.2024")', expected: '45296' },
  { formula: '=DATEVALUE("5 января 2024")', expected: '45296' },
  { formula: '=DATEVALUE("2024-01-05")', expected: '45296' },
  { formula: '=DATEVALUE("не дата")', expected: '#VALUE!' },
  { formula: '=TIME(12;0;0)', expected: '0.5' },
  { formula: '=EDATE(DATE(2024;1;31);1)', expected: '45351' },
  { formula: '=EOMONTH(DATE(2024;1;15);0)', expected: '45322' },
  { formula: '=WEEKNUM(DATE(2024;1;5))', expected: '1' },
  { formula: '=IFS(FALSE;"a";TRUE;"b")', expected: 'b' },
  // Excel не открывает файл с непарным списком аргументов ИФС — движок тоже бракует
  { formula: '=IFS(TRUE;"a";"b")', expected: '#VALUE!' },
  { formula: '=SWITCH(2;1;"один";2;"два";"иное")', expected: 'два' },
  { formula: '=SUBTOTAL(9;Данные!A1:I1)', expected: '36' },
  { formula: '=MAXIFS(Данные!A1:I1;Данные!A1:I1;">3")', expected: '9' },
  { formula: '=MINIFS(Данные!A1:I1;Данные!A1:I1;">3")', expected: '4' },
  { formula: '=AVERAGEIFS(Данные!A1:I1;Данные!A1:I1;">3")', expected: '6.6' },
  // Неизвестная функция по-прежнему #NAME? — её и показывает кнопка исправления
  { formula: '=НЕИЗВЕСТНАЯФУНКЦИЯ(1)', expected: '#NAME?' },
];

checkCases(arrayCtx, arrayCases, 'массивы: ');

/* ---------- Совместимость с Excel: запись и чтение формул ---------- */
/**
 * Функции новее Excel 2007 обязаны уезжать в .xlsx с префиксом `_xlfn.`,
 * иначе Excel не распознаёт имя и показывает #NAME?.
 */
const excelCases: [string, string][] = [
  ['=ЕСЛИМН(1>0;"да";"нет")', '=_xlfn.IFS(1>0,"да","нет")'],
  ['=МАКСЕСЛИМН(A1:A9;A1:A9;">3")', '=_xlfn.MAXIFS(A1:A9,A1:A9,">3")'],
  ['=МИНЕСЛИМН(A1:A9;A1:A9;">3")', '=_xlfn.MINIFS(A1:A9,A1:A9,">3")'],
  ['=ПЕРЕКЛЮЧ(2;1;"а";2;"б")', '=_xlfn.SWITCH(2,1,"а",2,"б")'],
  ['=ОБЪЕДИНИТЬ(", ";ИСТИНА;A1:A2)', '=_xlfn.TEXTJOIN(", ",TRUE,A1:A2)'],
  ['=ПРОСМОТРX(A1;A1:A9;A1:A9)', '=_xlfn.XLOOKUP(A1,A1:A9,A1:A9)'],
  ['=СЦЕП("a";"b")', '=_xlfn.CONCAT("a","b")'],
  ['=ДНИ(A2;A1)', '=_xlfn.DAYS(A2,A1)'],
  // Функции Excel 2007 и старше — без префикса
  ['=ДАТАЗНАЧ("05.01.2024")', '=DATEVALUE("05.01.2024")'],
  ['=ВРЕМЯ(12;0;0)', '=TIME(12,0,0)'],
  ['=ДАТАМЕС(ДАТА(2024;1;31);1)', '=EDATE(DATE(2024,1,31),1)'],
  ['=КОНМЕСЯЦА(ДАТА(2024;1;15);0)', '=EOMONTH(DATE(2024,1,15),0)'],
  ['=НОМНЕДЕЛИ(ДАТА(2024;1;5))', '=WEEKNUM(DATE(2024,1,5))'],
  ['=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;A1:A9)', '=SUBTOTAL(9,A1:A9)'],
  ['=СРЗНАЧЕСЛИМН(A1:A9;A1:A9;">3")', '=AVERAGEIFS(A1:A9,A1:A9,">3")'],
  ['=СУММ(A1:A2)', '=SUM(A1:A2)'],
];

const excelFailures: string[] = [];
for (const [source, expected] of excelCases) {
  const actual = toExcelFormula(source);
  if (actual !== expected) {
    excelFailures.push(`запись: ${source}\n     ожидалось: ${expected}\n     получено:   ${actual}`);
  }
  // Обратное чтение файла обязано вернуть формулу в диалекте редактора
  const roundTrip = fromExcelFormula(actual);
  const reference = fromExcelFormula(expected);
  if (roundTrip !== reference) {
    excelFailures.push(
      `чтение: ${actual}\n     ожидалось: ${reference}\n     получено:   ${roundTrip}`
    );
  }
}

// Движок обязан понимать формулы прямо из файла, с префиксом _xlfn.
const prefixed = evaluateFormula('=_xlfn.IFS(TRUE;1;TRUE;2)', arrayCtx);
if (isError(prefixed) || prefixed !== 1) {
  excelFailures.push(`вычисление: =_xlfn.IFS(TRUE;1;TRUE;2) => ${isError(prefixed) ? prefixed.__error : prefixed}`);
}
if (normalizeFormula('=_xlfn.XLOOKUP(A1,A1:A2,A1:A2)') !== '=XLOOKUP(A1,A1:A2,A1:A2)') {
  excelFailures.push(
    `нормализация: ${normalizeFormula('=_xlfn.XLOOKUP(A1,A1:A2,A1:A2)')}`
  );
}

/* ---------- Отображение ячеек листа ---------- */
const sheetCases: [number, number, string][] = [
  [4, 1, '35'],
  [4, 2, '1500'],
  [4, 3, '2'],
  [0, 0, 'Товар'],
];

const sheetFailures: string[] = [];
for (const [row, col, expected] of sheetCases) {
  const actual = computeCellDisplay(book, 'Лист1', row, col);
  if (actual !== expected) {
    sheetFailures.push(`Лист1!${row + 1},${col} ожидалось ${expected}, получено ${actual}`);
  }
}

/* ---------- Циклические ссылки ---------- */
const circularBook: Workbook = {
  S: [
    ['=B1', '=A1'],
  ],
};
const circular = evaluateFormula('=A1', createContext(circularBook, 'S'));
const circularOk = isError(circular) && circular.__error === '#CIRC!';

console.log('='.repeat(72));
console.log('  ПРОВЕРКА ДВИЖКА ФОРМУЛ');
console.log('='.repeat(72));
console.log(`  Функций зарегистрировано: ${FUNCTION_NAMES.length}`);
console.log(`  Проверок формул: ${cases.length}, пройдено: ${passed}`);
console.log(`  Проверок отображения ячеек: ${sheetCases.length}, провалов: ${sheetFailures.length}`);
console.log(`  Совместимость с Excel (.xlsx): ${excelCases.length}, провалов: ${excelFailures.length}`);
console.log(`  Обнаружение циклов: ${circularOk ? 'работает' : 'НЕ РАБОТАЕТ'}`);

if (failures.length > 0) {
  console.log('-'.repeat(72));
  console.log('  ПРОВАЛЫ:');
  for (const f of failures) console.log(`   - ${f}`);
}
if (sheetFailures.length > 0) {
  console.log('-'.repeat(72));
  console.log('  ПРОВАЛЫ ОТОБРАЖЕНИЯ:');
  for (const f of sheetFailures) console.log(`   - ${f}`);
}
if (excelFailures.length > 0) {
  console.log('-'.repeat(72));
  console.log('  ПРОВАЛЫ СОВМЕСТИМОСТИ С EXCEL:');
  for (const f of excelFailures) console.log(`   - ${f}`);
}

const totalFailures =
  failures.length + sheetFailures.length + excelFailures.length + (circularOk ? 0 : 1);
console.log('='.repeat(72));
console.log(totalFailures === 0 ? '  ИТОГ: ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ' : `  ИТОГ: ПРОВАЛОВ ${totalFailures}`);
process.exit(totalFailures === 0 ? 0 : 1);
