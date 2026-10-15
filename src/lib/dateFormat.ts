/**
 * Форматирование значений ячеек по коду числового формата Excel.
 *
 * Зачем нужен собственный форматтер дат: встроенный `XLSX.SSF.format`
 * не понимает точку как разделитель частей кода и на самом распространённом
 * русском формате «ДД.ММ.ГГГГ» выбрасывает исключение «bad second format».
 * Поэтому даты и время форматируются здесь, а SSF используется только для
 * числовых и денежных форматов.
 */

import * as XLSX from 'xlsx';

/** Русские названия месяцев в родительном падеже (для «15 марта 2023 г.»). */
const MONTHS_GENITIVE = [
  'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
  'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря',
];

/** Русские названия месяцев в именительном падеже (для «Март 2023»). */
const MONTHS_NOMINATIVE = [
  'январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
  'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь',
];

/** Английские названия — для форматов с латиницей. */
const MONTHS_EN = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

export interface FormatOptions {
  /** Локаль для названий месяцев: 'ru' (по умолчанию) или 'en'. */
  locale?: 'ru' | 'en';
}

/**
 * Признак того, что код формата описывает дату или время.
 * Учитывает русские обозначения (Д, М, Г, Ч) и служебные префиксы [$…].
 */
export function isDateFormat(format: string): boolean {
  if (!format) return false;
  // Отбрасываем секцию для отрицательных/нулевых значений и префиксы локали
  const cleaned = format
    .replace(/\[[^\]]*\]/g, '')
    .replace(/"[^"]*"/g, '')
    .split(';')[0];
  return /[dmyhsдмгчDY]/i.test(cleaned);
}

/** Разбирает серийное число Excel в компоненты даты и времени. */
function serialToParts(serial: number): {
  year: number;
  month: number;
  day: number;
  hours: number;
  minutes: number;
  seconds: number;
  weekday: number;
  totalHours: number;
} {
  // Отрицательные значения Excel не отображает как даты
  const whole = Math.floor(serial);
  const fraction = serial - whole;

  // Эпоха системы 1900: серийное 1 = 01.01.1900.
  // Для серийных 1..59 отсчёт идёт от 31.12.1899. Начиная с 01.03.1900 Excel
  // добавляет мнимый 29.02.1900, поэтому для 61 и далее отсчёт — от 30.12.1899.
  // Без этого разделения все современные даты уезжали на сутки вперёд.
  const epoch = Date.UTC(1899, 11, whole < 60 ? 31 : 30);
  const date = new Date(epoch + whole * 86400000);

  // Доли суток -> время; округление до секунды убирает накопленную погрешность
  let totalSeconds = Math.round(fraction * 86400);
  if (totalSeconds >= 86400) totalSeconds = 86399;

  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  let day = date.getUTCDate();
  let month = date.getUTCMonth() + 1;
  let year = date.getUTCFullYear();

  if (whole === 60) {
    // Серийное 60 — несуществующее 29.02.1900 в Excel; показываем как 29.02.1900
    year = 1900;
    month = 2;
    day = 29;
  }

  return {
    year,
    month,
    day,
    hours,
    minutes,
    seconds,
    weekday: date.getUTCDay(),
    totalHours: Math.floor(serial * 24),
  };
}

const pad = (n: number, width = 2): string => String(Math.abs(n)).padStart(width, '0');

/**
 * Форматирует серийное число по коду формата даты.
 * Поддерживает: Д/М/Г (d/m/y), Ч/М/С (h/m/s), названия месяцев (ММММ/МММ),
 * день недели (ДДД/ДДДД), AM/PM и экранированные разделители.
 */
export function formatSerialDate(
  serial: number,
  format: string,
  options: FormatOptions = {}
): string {
  const locale = options.locale ?? 'ru';
  const parts = serialToParts(serial);

  const months = locale === 'ru' ? MONTHS_GENITIVE : MONTHS_EN;
  const monthsNom = locale === 'ru' ? MONTHS_NOMINATIVE : MONTHS_EN;
  const weekdaysRu = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
  const weekdaysShortRu = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
  const weekdaysEn = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  // Берём первую секцию (формат для положительных значений)
  let fmt = format.split(';')[0];

  // Убираем префиксы локали: [$-419], [$RU-RU]
  fmt = fmt.replace(/\[\$[^\]]*\]/g, '');

  // Определяем 12- или 24-часовой формат
  const hasAmPm = /am\/pm|a\/p/i.test(fmt);
  const is12Hour = hasAmPm;

  let hours = parts.hours;
  let ampm = '';
  if (is12Hour) {
    ampm = hours < 12 ? 'AM' : 'PM';
    // Сопоставляем регистр исходного формата
    const sample = fmt.match(/am\/pm|a\/p/i)?.[0] ?? '';
    if (sample === sample.toLowerCase()) ampm = ampm.toLowerCase();
    hours = hours % 12;
    if (hours === 0) hours = 12;
  }

  // Для английской локали названия месяцев берутся из латинского списка
  const monthNames = locale === 'ru' ? MONTHS_GENITIVE : MONTHS_EN;
  const monthNamesNom = locale === 'ru' ? MONTHS_NOMINATIVE : MONTHS_EN;

  // Токенизация: находим группы одинаковых букв формата
  const tokens: { text: string; isFormat: boolean }[] = [];
  let i = 0;
  while (i < fmt.length) {
    const ch = fmt[i];

    // Экранированный символ: \x — выводится как есть
    if (ch === '\\' && i + 1 < fmt.length) {
      tokens.push({ text: fmt[i + 1], isFormat: false });
      i += 2;
      continue;
    }

    // Текст в кавычках выводится как есть
    if (ch === '"') {
      const end = fmt.indexOf('"', i + 1);
      const text = end === -1 ? fmt.slice(i + 1) : fmt.slice(i + 1, end);
      tokens.push({ text, isFormat: false });
      i = end === -1 ? fmt.length : end + 1;
      continue;
    }

    // AM/PM и A/P — цельные обозначения до 12-часового формата
    const ampmMatch = fmt.slice(i).match(/^(AM\/PM|am\/pm|A\/P|a\/p)/);
    if (ampmMatch) {
      tokens.push({ text: ampmMatch[1], isFormat: true });
      i += ampmMatch[1].length;
      continue;
    }

    // Накопленные часы вида [ч] или [h]
    const bracketMatch = fmt.slice(i).match(/^\[([hч]+)\]/i);
    if (bracketMatch) {
      tokens.push({ text: bracketMatch[0], isFormat: true });
      i += bracketMatch[0].length;
      continue;
    }

    // Название месяца/дня или иной буквенный код
    if (/[A-Za-zА-Яа-яЁё]/.test(ch)) {
      let j = i;
      while (j < fmt.length && fmt[j] === ch) j++;
      tokens.push({ text: fmt.slice(i, j), isFormat: true });
      i = j;
      continue;
    }

    // Разделители и прочие символы
    tokens.push({ text: ch, isFormat: false });
    i++;
  }

  /**
   * Формат содержит только месяц, без дня и года («ММММ»).
   * В этом случае русское название ставится в именительный падеж: «март».
   * В составе даты («Д ММММ ГГГГ») используется родительный: «15 марта 2023».
   */
  const monthOnly = !tokens.some(
    (token) =>
      token.isFormat &&
      (/^[dд]+$/i.test(token.text) || /^[yг]+$/i.test(token.text))
  );

  const out: string[] = [];
  /** Был ли уже выведен код часов — от этого зависит смысл «m». */
  let sawHour = false;

  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index];
    if (!token.isFormat) {
      out.push(token.text);
      continue;
    }

    const t = token.text;
    const lower = t.toLowerCase();
    const upper = t.toUpperCase();
    const n = t.length;

    // Накопленные часы [ч] / [h]
    if (/^\[[hч]+\]$/i.test(t)) {
      out.push(String(parts.totalHours));
      sawHour = true;
      continue;
    }

    // Год: Г(г) или Y
    if (/^[yг]+$/i.test(t)) {
      out.push(n <= 2 ? pad(parts.year % 100) : String(parts.year));
      continue;
    }

    // AM/PM
    if (/^am\/pm$/i.test(t) || /^a\/p$/i.test(t)) {
      out.push(ampm);
      continue;
    }

    // Часы: Ч(h) или H
    if (/^[hч]+$/i.test(t)) {
      out.push(n >= 2 ? pad(hours) : String(hours));
      sawHour = true;
      continue;
    }

    // Секунды: С(s) или S
    if (/^[sс]+$/i.test(t)) {
      out.push(n >= 2 ? pad(parts.seconds) : String(parts.seconds));
      continue;
    }

    // День: Д(d) или D
    if (/^[dд]+$/i.test(t)) {
      if (n >= 4) {
        out.push(locale === 'ru' ? weekdaysRu[parts.weekday] : weekdaysEn[parts.weekday]);
      } else if (n === 3) {
        out.push(locale === 'ru' ? weekdaysShortRu[parts.weekday] : weekdaysEn[parts.weekday].slice(0, 3));
      } else if (n === 2) {
        out.push(pad(parts.day));
      } else {
        out.push(String(parts.day));
      }
      continue;
    }

    if (/^[mм]+$/i.test(t)) {
      /*
       * Различение месяца и минут — так же, как это делает Excel:
       *  - «m» сразу после кода часов или перед кодом секунд — это минуты;
       *  - в остальных случаях — месяц.
       * Кириллическая «М» в русских форматах всегда обозначает месяц.
       */
      const isCyrillic = /[мМ]/.test(t);
      const nextFormat = tokens.slice(index + 1).find((x) => x.isFormat);
      const beforeIsHour = sawHour;
      const nextIsSecond = Boolean(nextFormat && /^[sс]+$/i.test(nextFormat.text));
      const isMinutes = !isCyrillic && (beforeIsHour || nextIsSecond);

      if (isMinutes) {
        out.push(n >= 2 ? pad(parts.minutes) : String(parts.minutes));
        continue;
      }

      // Именительный падеж только для одиночного «ММММ» без дня и года.
      const names = monthOnly ? monthNamesNom : monthNames;

      if (n >= 4) {
        // В составе даты — родительный падеж: «15 марта 2023»
        out.push(names[parts.month - 1]);
      } else if (n === 3) {
        out.push(names[parts.month - 1].slice(0, 3));
      } else if (n === 2) {
        out.push(pad(parts.month));
      } else {
        out.push(String(parts.month));
      }
      continue;
    }

    // Неизвестный код — выводим как есть
    out.push(t);
  }

  return out.join('');
}

/**
 * Форматирует значение ячейки по коду формата Excel.
 * Возвращает null, если формат не является датой, — тогда вызывающий код
 * применяет числовое форматирование.
 */
export function formatCellValue(
  value: string | number | boolean | null,
  format: string | undefined,
  options: FormatOptions = {}
): string | null {
  if (value === null || value === undefined || value === '') return '';
  if (!format) return null;

  const numeric = typeof value === 'number' ? value : Number(String(value).replace(/\s/g, '').replace(',', '.'));
  if (!Number.isFinite(numeric)) return null;

  if (isDateFormat(format)) {
    // Отрицательные и нулевые серийные числа Excel датами не считает
    if (numeric <= 0) return null;
    return formatSerialDate(numeric, format, options);
  }

  // Числовые и денежные форматы — пробуем SSF, при ошибке возвращаем null
  try {
    return XLSX.SSF.format(format, numeric);
  } catch {
    return null;
  }
}

/**
 * Преобразует календарную дату в серийное число Excel.
 * Точка отсчёта — 31.12.1899, поэтому 01.01.1900 даёт серийное 1.
 */
function dateToSerial(date: Date): number {
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();

  // Серийное 60 — несуществующее 29.02.1900, которое Excel считает реальным.
  if (year === 1900 && month === 2 && day === 29) return 60;

  // Отсчёт от 31.12.1899: 01.01.1900 даёт 1, 28.02.1900 даёт 59.
  const epoch = Date.UTC(1899, 11, 31);
  const serial = Math.round((Date.UTC(year, month - 1, day) - epoch) / 86400000);

  // Начиная с 01.03.1900 Excel отсчитывает лишние сутки (мнимый високосный
  // день), поэтому даты этого дня и позже сдвинуты на +1.
  return serial >= 60 ? serial + 1 : serial;
}

/** Проверяет существование календарной даты (отсекает 31.02 и подобное). */
function makeDate(year: number, month: number, day: number): Date | null {
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return date;
}

/** Проверяет, является ли строка распознаваемой датой (для ввода пользователем). */
export function parseDateInput(input: string): number | null {
  const text = input.trim();
  if (!text) return null;

  // ДД.ММ.ГГГГ или ДД/ММ/ГГГГ либо ДД-ММ-ГГГГ, возможно со временем
  const m = text.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const day = parseInt(m[1], 10);
    const month = parseInt(m[2], 10);
    let year = parseInt(m[3], 10);
    if (year < 100) year += year < 70 ? 2000 : 1900;

    const date = makeDate(year, month, day);
    if (!date) return null;

    let serial = dateToSerial(date);

    if (m[4] !== undefined) {
      const hours = parseInt(m[4], 10);
      const minutes = parseInt(m[5], 10);
      const seconds = m[6] ? parseInt(m[6], 10) : 0;
      if (hours > 23 || minutes > 59 || seconds > 59) return null;
      serial += (hours * 3600 + minutes * 60 + seconds) / 86400;
    }
    return serial;
  }

  // ГГГГ-ММ-ДД (ISO)
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    const date = makeDate(parseInt(iso[1], 10), parseInt(iso[2], 10), parseInt(iso[3], 10));
    if (!date) return null;
    return dateToSerial(date);
  }

  return null;
}
