/**
 * Подбор ширины колонок для настоящего Excel.
 *
 * WorkSpace рисует пустые колонки шириной 104 px, а в .xlsx без `!cols`
 * Excel открывает стандартные 8,43 символа и обрезает текст. Поэтому ширины
 * пишутся всегда — по содержимому, с запасом на кириллицу и широкие символы.
 */
import type { ColInfo, WorkSheet } from 'xlsx';

export const DEFAULT_MIN_WCH = 10;
export const DEFAULT_MAX_WCH = 70;

/** Символы, которые в Calibri заметно уже цифры. */
const NARROW_CHARS = /[il.,;:'’!|`]/;
const WIDE_CHARS =
  /[\u1100-\u115f\u2e80-\u303e\u3040-\u33ff\u3400-\u4dbf\u4e00-\u9fff\ua000-\ua4cf\uac00-\ud7a3\uf900-\ufaff\ufe10-\ufe6f\uff00-\uff60\uffe0-\uffe6]/u;
const UPPERCASE = /[\p{Lu}]/u;

/** Ширина строки в «символах Excel» (MDW ≈ 7 px). */
export function textWidthChars(text: string): number {
  let width = 0;
  let line = 0;
  for (const char of Array.from(text)) {
    if (char === '\t') {
      line += 4;
      continue;
    }
    if (char === '\r') continue;
    if (char === '\n') {
      width = Math.max(width, line);
      line = 0;
      continue;
    }
    if (WIDE_CHARS.test(char)) {
      line += 2;
      continue;
    }
    if (NARROW_CHARS.test(char)) {
      line += 0.55;
      continue;
    }
    if (char === ' ') {
      line += 0.55;
      continue;
    }
    line += UPPERCASE.test(char) ? 1.1 : 1;
  }
  return Math.max(width, line);
}

/** Ширина в px из редактора → ширина в символах Excel. */
export function pxToWch(px: number): number {
  return Math.max(1, (px - 5) / 7);
}

export interface AutoFitOptions {
  /** Явно заданные пользователем ширины в px (колонка → px). */
  explicitPx?: Record<number, number>;
  /** Ячейка с включённым переносом не должна расширять колонку. */
  isWrapped?: (row: number, col: number) => boolean;
  /** Строки заголовка: жирный шрифт шире обычного. */
  headerRows?: number;
  minWch?: number;
  maxWch?: number;
  /** Запас в символах, чтобы текст не упирался в границу. */
  padding?: number;
}

/** Ширины для всех непустых колонок матрицы. */
export function autoFitColumns(
  matrix: readonly (readonly unknown[])[],
  options: AutoFitOptions = {}
): ColInfo[] {
  const {
    explicitPx = {},
    isWrapped,
    headerRows = 0,
    minWch = DEFAULT_MIN_WCH,
    maxWch = DEFAULT_MAX_WCH,
    padding = 2,
  } = options;

  let lastCol = -1;
  let lastRow = -1;
  for (let row = 0; row < matrix.length; row++) {
    const line = matrix[row] ?? [];
    for (let col = 0; col < line.length; col++) {
      if (String(line[col] ?? '').trim() !== '') {
        lastCol = Math.max(lastCol, col);
        lastRow = Math.max(lastRow, row);
      }
    }
  }
  if (lastCol < 0) return [];

  const widths: ColInfo[] = [];
  for (let col = 0; col <= lastCol; col++) {
    let contentWch = 0;
    let hasUnwrapped = false;
    for (let row = 0; row <= lastRow; row++) {
      const value = matrix[row]?.[col];
      const text = String(value ?? '');
      if (!text.trim()) continue;
      if (isWrapped?.(row, col)) continue;
      hasUnwrapped = true;
      const rowWidth =
        textWidthChars(text) + (row < headerRows ? 1 : 0) + padding;
      contentWch = Math.max(contentWch, rowWidth);
    }

    const explicit = explicitPx[col];
    const explicitWch = typeof explicit === 'number' ? pxToWch(explicit) : 0;
    // Явную ширину уважаем, но не даём тексту обрезаться. Колонки только
    // с переносом не раздуваем: им достаточно минимальной ширины.
    const target = Math.max(
      explicitWch,
      hasUnwrapped ? contentWch : minWch,
      minWch
    );
    widths.push({ wch: Math.round(Math.min(maxWch, target) * 2) / 2 });
  }
  return widths;
}

/** Пишет `!cols` на лист, сохраняя уже существующие явные ширины. */
export function applyAutoFitColumns(
  sheet: WorkSheet,
  matrix: readonly (readonly unknown[])[],
  options: AutoFitOptions = {}
): ColInfo[] {
  const existing: Record<number, number> = {};
  (sheet['!cols'] ?? []).forEach((col, index) => {
    if (!col) return;
    if (typeof col.wch === 'number') existing[index] = col.wch;
    else if (typeof col.wpx === 'number') existing[index] = pxToWch(col.wpx);
    else if (typeof col.width === 'number') existing[index] = col.width;
  });

  const explicitPx = { ...(options.explicitPx ?? {}) };
  for (const [index, wch] of Object.entries(existing)) {
    const col = Number(index);
    // Для уже записанных wch переводим обратно в px лишь приблизительно:
    // главное — не потерять ширину при повторном сохранении.
    if (explicitPx[col] === undefined) explicitPx[col] = wch * 7 + 5;
  }

  const widths = autoFitColumns(matrix, { ...options, explicitPx });
  if (widths.length > 0) sheet['!cols'] = widths;
  return widths;
}
