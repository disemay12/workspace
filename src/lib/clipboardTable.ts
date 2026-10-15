/**
 * Разбор таблицы из буфера обмена.
 *
 * Приоритет отдаётся HTML: при копировании из браузера, Word или ответа AI
 * в буфере лежит настоящая <table>, и она сохраняет структуру даже тогда,
 * когда в ячейках есть переносы строк. Если HTML нет — разбираем обычный
 * текст: TSV (как копирует Excel) или CSV.
 */

export interface ParsedTable {
  rows: string[][];
  source: 'html' | 'tsv' | 'csv' | 'text';
}

/** Текст ячейки: переносы <br> и абзацы сохраняются как переводы строк. */
function cellText(cell: HTMLTableCellElement): string {
  const clone = cell.cloneNode(true) as HTMLElement;
  clone.querySelectorAll('br').forEach((br) => br.replaceWith('\n'));
  clone.querySelectorAll('p,div,li').forEach((el) => el.append('\n'));
  const text = (clone.textContent || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{2,}/g, '\n');
  return text.replace(/^\s+|\s+$/g, '');
}

/** Разворачивает <table> в матрицу строк с учётом colspan/rowspan. */
function tableToMatrix(table: HTMLTableElement): string[][] {
  const grid: string[][] = [];
  const spans: Record<string, { remaining: number; value: string }> = {};

  const takeSpans = (r: number, startCol: number): number => {
    let c = startCol;
    for (;;) {
      const key = r + ':' + c;
      const span = spans[key];
      if (!span || span.remaining <= 0) break;
      if (!grid[r]) grid[r] = [];
      grid[r][c] = span.value;
      span.remaining -= 1;
      c += 1;
    }
    return c;
  };

  Array.from(table.rows).forEach((rowEl, r) => {
    if (!grid[r]) grid[r] = [];
    let c = takeSpans(r, 0);

    Array.from(rowEl.cells).forEach((cell) => {
      c = takeSpans(r, c);
      const value = cellText(cell);
      const colSpan = Math.max(1, cell.colSpan || 1);
      const rowSpan = Math.max(1, cell.rowSpan || 1);

      for (let k = 0; k < colSpan; k++) {
        grid[r][c] = value;
        if (rowSpan > 1) spans[r + 1 + ':' + c] = { remaining: rowSpan - 1, value };
        c += 1;
      }
      c = takeSpans(r, c);
    });
  });

  const width = grid.reduce((max, row) => Math.max(max, row.length), 0);
  return grid.map((row) => {
    const out: string[] = new Array(width).fill('');
    for (let i = 0; i < row.length; i++) out[i] = row[i] === undefined ? '' : row[i];
    return out;
  });
}

function countChar(text: string, ch: string): number {
  let n = 0;
  for (let i = 0; i < text.length; i++) if (text[i] === ch) n += 1;
  return n;
}

/** Разбор одной строки CSV с кавычками и удвоением кавычек. */
function parseDelimitedLine(line: string): string[] {
  const delim = countChar(line, ';') > countChar(line, ',') ? ';' : ',';
  const out: string[] = [];
  let cur = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        cur += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === delim) {
      out.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

/** Убирает полностью пустые строки и столбцы по краям. */
function trimEmpty(rows: string[][]): string[][] {
  let end = rows.length;
  while (end > 0 && rows[end - 1].every((c) => String(c).trim() === '')) end -= 1;
  const sliced = rows.slice(0, end);

  let width = sliced.reduce((max, r) => Math.max(max, r.length), 0);
  while (width > 0 && sliced.every((r) => String(r[width - 1] === undefined ? '' : r[width - 1]).trim() === '')) {
    width -= 1;
  }

  return sliced.map((r) => {
    const out: string[] = [];
    for (let i = 0; i < width; i++) out.push(r[i] === undefined ? '' : String(r[i]));
    return out;
  });
}

/**
 * Преобразует содержимое буфера обмена в матрицу ячеек.
 * Возвращает null, если данных для вставки нет.
 */
export function parseClipboardTable(html: string, text: string): ParsedTable | null {
  if (html && /<table[\s>]/i.test(html)) {
    try {
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const table = doc.querySelector('table');
      if (table) {
        const rows = trimEmpty(tableToMatrix(table));
        const hasData = rows.some((r) => r.some((c) => String(c).trim() !== ''));
        if (hasData) return { rows, source: 'html' };
      }
    } catch {
      /* повреждённый HTML — переходим к текстовому разбору */
    }
  }

  const plain = (text || '').replace(/\r\n?/g, '\n').replace(/\n+$/, '');
  if (!plain.trim()) return null;

  const lines = plain.split('\n');
  if (lines.some((l) => l.indexOf('\t') !== -1)) {
    return { rows: trimEmpty(lines.map((l) => l.split('\t'))), source: 'tsv' };
  }
  if (lines.some((l) => l.indexOf(',') !== -1 || l.indexOf(';') !== -1)) {
    return { rows: trimEmpty(lines.map(parseDelimitedLine)), source: 'csv' };
  }
  return { rows: trimEmpty(lines.map((l) => [l])), source: 'text' };
}
