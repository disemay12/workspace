/**
 * Проверка форматирования дат: коды форматов из реальных русских файлов
 * Excel, включая те, на которых встроенный SSF выбрасывает исключение.
 */
import {
  formatSerialDate,
  isDateFormat,
  formatCellValue,
  parseDateInput,
} from '../../src/lib/dateFormat';

const failures: string[] = [];
let passed = 0;

function check(name: string, actual: unknown, expected: unknown) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) passed++;
  else failures.push(`${name}\n     ожидалось: ${e}\n     получено:   ${a}`);
}

/* 15.03.2023 = серийное 45000 */
const D = 45000;

/* ---------- Основные русские форматы (SSF на них падает) ---------- */
check('ДД.ММ.ГГГГ', formatSerialDate(D, 'DD.MM.YYYY'), '15.03.2023');
check('ДД.ММ.ГГ', formatSerialDate(D, 'DD.MM.YY'), '15.03.23');
check('Д.М.ГГГГ без ведущих нулей', formatSerialDate(D, 'D.M.YYYY'), '15.3.2023');
check('строчные дд.мм.гггг', formatSerialDate(D, 'dd.mm.yyyy'), '15.03.2023');
check('префикс локали [$-419]', formatSerialDate(D, '[$-419]DD.MM.YYYY'), '15.03.2023');
check('префикс [$RU-RU]', formatSerialDate(D, '[$RU-RU]dd.mm.yyyy'), '15.03.2023');
check('дата со временем', formatSerialDate(D + 0.5, 'DD.MM.YYYY H:MM'), '15.03.2023 12:00');
check('дата с секундами', formatSerialDate(D + 0.5, 'DD.MM.YYYY HH:MM:SS'), '15.03.2023 12:00:00');

/* ---------- Латиница ---------- */
check('ISO-подобный', formatSerialDate(D, 'yyyy-mm-dd'), '2023-03-15');
check('американский', formatSerialDate(D, 'm/d/yy'), '3/15/23');
check('европейский слэш', formatSerialDate(D, 'dd/mm/yyyy'), '15/03/2023');
check('экранированные точки', formatSerialDate(D, 'dd"."mm"."yyyy'), '15.03.2023');

/* ---------- Названия месяцев ---------- */
check('месяц прописью (рус.)', formatSerialDate(D, 'Д ММММ ГГГГ'), '15 марта 2023');
check('месяц кратко (рус.)', formatSerialDate(D, 'Д МММ ГГГГ'), '15 мар 2023');
check('месяц прописью (англ.)', formatSerialDate(D, 'd mmmm yyyy', { locale: 'en' }), '15 March 2023');
check('только месяц', formatSerialDate(D, 'ММММ'), 'март');

/* ---------- Время ---------- */
check('часы и минуты', formatSerialDate(0.5, 'hh:mm'), '12:00');
check('полночь', formatSerialDate(0.0, 'hh:mm'), '00:00');
check('23:59', formatSerialDate((23 * 3600 + 59 * 60) / 86400, 'hh:mm'), '23:59');
check('12-часовой формат AM', formatSerialDate(0.25, 'h:mm AM/PM'), '6:00 AM');
check('12-часовой формат PM', formatSerialDate(0.75, 'h:mm AM/PM'), '6:00 PM');

/* ---------- День недели ---------- */
check('день недели кратко', formatSerialDate(D, 'ДДД'), 'ср');
check('день недели полностью', formatSerialDate(D, 'ДДДД'), 'среда');

/* ---------- Границы и особые случаи ---------- */
check('1 = 01.01.1900', formatSerialDate(1, 'DD.MM.YYYY'), '01.01.1900');
check('серийное 0 не дата', formatCellValue(0, 'DD.MM.YYYY'), null);
check('отрицательное не дата', formatCellValue(-5, 'DD.MM.YYYY'), null);

/* ---------- Определение формата даты ---------- */
check('DD.MM.YYYY — дата', isDateFormat('DD.MM.YYYY'), true);
check('dd/mm/yy — дата', isDateFormat('dd/mm/yy'), true);
check('0.00 — не дата', isDateFormat('0.00'), false);
check('#,##0 — не дата', isDateFormat('#,##0'), false);
check('0% — не дата', isDateFormat('0%'), false);
check('General — не дата', isDateFormat('General'), false);

/* ---------- Числовые форматы через formatCellValue ---------- */
check('денежный формат', formatCellValue(1234.5, '#,##0.00'), '1,234.50');
check('проценты', formatCellValue(0.25, '0%'), '25%');
check('без формата возвращает null', formatCellValue(42, undefined), null);

/* ---------- Разбор ввода пользователя ---------- */
check('ввод 15.03.2023', parseDateInput('15.03.2023'), D);
check('ввод 15/03/2023', parseDateInput('15/03/2023'), D);
check('ввод 2023-03-15', parseDateInput('2023-03-15'), D);
check('ввод с временем', parseDateInput('15.03.2023 12:00'), D + 0.5);
check('некорректная дата отклонена', parseDateInput('32.13.2023'), null);
check('текст отклонён', parseDateInput('привет'), null);
check('пустая строка отклонена', parseDateInput(''), null);

/* ---------- Согласованность ввода и вывода ---------- */
check(
  'ввод -> вывод',
  formatSerialDate(parseDateInput('29.02.2024')!, 'DD.MM.YYYY'),
  '29.02.2024'
);
check('29 февраля невисокосного года отклонено', parseDateInput('29.02.2023'), null);

console.log('='.repeat(72));
console.log('  ПРОВЕРКА ФОРМАТИРОВАНИЯ ДАТ');
console.log('='.repeat(72));
console.log(`  Пройдено: ${passed}, провалов: ${failures.length}`);
if (failures.length > 0) {
  console.log('-'.repeat(72));
  for (const f of failures) console.log(`   - ${f}`);
}
console.log('='.repeat(72));
console.log(failures.length === 0 ? '  ИТОГ: ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ' : '  ИТОГ: ЕСТЬ ПРОВАЛЫ');
process.exit(failures.length === 0 ? 0 : 1);
