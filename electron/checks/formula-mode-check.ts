/**
 * Проверка логики формульного режима таблицы без запуска Electron:
 * автодополнение имён функций, определение незакрытой функции,
 * выбор диапазона мышью и сдвиг относительных ссылок при заполнении.
 *
 * Эти функции — чистые, поэтому проверяются в обычном Node.
 */
import {
  suggestFunctions,
  findOpenFunction,
  partialWordBefore,
  rangeToA1,
  normalizeRange,
  shiftFormula,
} from '../../src/components/editors/useFormulaMode';
import { normalizeFormula, evaluateFormula, createContext } from '../../src/lib/formulaEngine';

const failures: string[] = [];
let passed = 0;

function check(name: string, actual: unknown, expected: unknown) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    passed++;
  } else {
    failures.push(`${name}\n     ожидалось: ${e}\n     получено:   ${a}`);
  }
}

/* ---------- 1. Подсказки функций ---------- */

check(
  'латиница: "SU" → SUM первым',
  suggestFunctions('SU')[0]?.name,
  'SUM'
);

check(
  'кириллица: "СУМ" → СУММ первой',
  suggestFunctions('СУМ')[0]?.label,
  'СУММ'
);

check(
  'кириллица: "СУМ" ведёт к SUM',
  suggestFunctions('СУМ')[0]?.name,
  'SUM'
);

check('пустой ввод → нет подсказок', suggestFunctions('').length, 0);

check(
  'ссылка B2 не считается именем функции',
  partialWordBefore('=SUM(B2', 7).word,
  ''
);

check(
  'частичное имя функции распознаётся',
  partialWordBefore('=SU', 3).word,
  'SU'
);

check(
  'ЕСЛИ предлагается по кириллице',
  suggestFunctions('ЕСЛ')[0]?.name,
  'IF'
);

check(
  'ВПР предлагается по кириллице',
  suggestFunctions('ВПР')[0]?.name,
  'VLOOKUP'
);

/* ---------- 2. Открытая функция ---------- */

check(
  'находит незакрытую SUM(',
  findOpenFunction('=SUM(', 5)?.name,
  'SUM'
);

check(
  'находит вложенную активную функцию',
  findOpenFunction('=IF(A1>0,SUM(B1', 14)?.name,
  'SUM'
);

check(
  'после закрытой скобки функции нет',
  findOpenFunction('=SUM(A1:A5)', 11),
  null
);

/* ---------- 3. Диапазоны ---------- */

check(
  'диапазон A1:B3',
  rangeToA1(normalizeRange({ row: 0, col: 0 }, { row: 2, col: 1 })),
  'A1:B3'
);

check(
  'одиночная ячейка без двоеточия',
  rangeToA1(normalizeRange({ row: 4, col: 2 }, { row: 4, col: 2 })),
  'C5'
);

check(
  'обратный порядок нормализуется',
  rangeToA1(normalizeRange({ row: 5, col: 3 }, { row: 1, col: 0 })),
  'A2:D6'
);

/* ---------- 4. Вставка ссылки в формулу ---------- */
/* Имитируем то, что делает хук: "=СУММ(" + клик по B2:B10 */

function insertRef(draft: string, caret: number, ref: string): string {
  return draft.slice(0, caret) + ref + draft.slice(caret);
}

const typed = '=СУММ(';
const afterClick = insertRef(typed, typed.length, 'B2');
check('клик по ячейке дописывает ссылку', afterClick, '=СУММ(B2');

const afterDrag = insertRef('=СУММ(', 6, 'B2:B10');
check('протяжка мышью вставляет диапазон', afterDrag, '=СУММ(B2:B10');

// Готовая формула, собранная мышью, должна считаться движком.
// Клик по B2 и протяжка до B11 даёт =СУММ(B2:B11); в этих строках лежат 1..10.
const book = {
  S: Array.from({ length: 12 }, (_, r) => [null, r === 0 ? null : r]),
};
const built = normalizeFormula(insertRef('=СУММ(', 6, 'B2:B11') + ')');
const value = evaluateFormula(built, createContext(book, 'S'));
check('собранная мышью формула вычисляется', value, 55);

/* ---------- 5. Сдвиг относительных ссылок ---------- */

check(
  'относительные ссылки сдвигаются вправо',
  shiftFormula('=B2+C2', 0, 1),
  '=C2+D2'
);

check(
  'абсолютные ссылки не сдвигаются',
  shiftFormula('=$B$2+C2', 0, 1),
  '=$B$2+D2'
);

check(
  'смешанные ссылки: абсолютный столбец сохраняется',
  shiftFormula('=$B2+C$2', 1, 1),
  '=$B3+D$2'
);

check(
  'сдвиг вниз по строкам',
  shiftFormula('=SUM(B2:B4)', 1, 0),
  '=SUM(B3:B5)'
);

check(
  'ссылка на другой лист сдвигается, имя листа сохраняется',
  shiftFormula('=Лист2!B2', 1, 1),
  '=Лист2!C3'
);

check(
  'абсолютные ссылки на другом листе не сдвигаются',
  shiftFormula('=Лист2!$B$2+A1', 2, 0),
  '=Лист2!$B$2+A3'
);

check(
  'имя листа в апострофах сохраняется',
  shiftFormula("='Мой лист'!B2", 1, 1),
  "='Мой лист'!C3"
);

check(
  'текст в кавычках не сдвигается',
  shiftFormula('=CONCAT("B2","D4")', 1, 1),
  '=CONCAT("B2","D4")'
);

check(
  'ссылка рядом со строкой сдвигается',
  shiftFormula('=CONCAT("итого: ",B2)', 0, 1),
  '=CONCAT("итого: ",C2)'
);

/* ---------- 6. Разбор ответа AI ---------- */

function parseAIFormula(answer: string): { target: string; value: string } | null {
  const cellMatch = answer.match(/ЯЧЕЙКА\s*:\s*([A-Za-zА-Яа-яЁё0-9_$!:' ]+)/i);
  const formulaMatch = answer.match(/ФОРМУЛА\s*:\s*(.+)/i);
  if (!formulaMatch) return null;
  let value = formulaMatch[1].trim().replace(/^[`"']+|[`"']+$/g, '').trim();
  if (value && !value.startsWith('=')) value = `=${value}`;
  return { target: cellMatch?.[1]?.trim().replace(/\s/g, '') || 'E2', value };
}

check(
  'строгий ответ разбирается',
  parseAIFormula('ЯЧЕЙКА: B5\nФОРМУЛА: =SUM(B2:B4)'),
  { target: 'B5', value: '=SUM(B2:B4)' }
);

check(
  'формула без "=" дополняется',
  parseAIFormula('ЯЧЕЙКА: C3\nФОРМУЛА: SUM(A1:A9)'),
  { target: 'C3', value: '=SUM(A1:A9)' }
);

check(
  'формула в бэктиках очищается',
  parseAIFormula('ЯЧЕЙКА: D4\nФОРМУЛА: `=AVERAGE(B:B)`'),
  { target: 'D4', value: '=AVERAGE(B:B)' }
);

check(
  'ответ без формулы не даёт действия',
  parseAIFormula('Не могу подобрать формулу'),
  null
);

/* ---------- Итог ---------- */

console.log('='.repeat(72));
console.log('  ПРОВЕРКА ФОРМУЛЬНОГО РЕЖИМА ТАБЛИЦЫ');
console.log('='.repeat(72));
console.log(`  Пройдено: ${passed}, провалов: ${failures.length}`);
if (failures.length > 0) {
  console.log('-'.repeat(72));
  for (const f of failures) console.log(`   - ${f}`);
}
console.log('='.repeat(72));
console.log(failures.length === 0 ? '  ИТОГ: ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ' : '  ИТОГ: ЕСТЬ ПРОВАЛЫ');
process.exit(failures.length === 0 ? 0 : 1);
