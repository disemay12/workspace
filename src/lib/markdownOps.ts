/**
 * Структурные правки Markdown: заголовки, списки, ссылки и таблицы.
 *
 * Функции чистые: получают исходный текст и возвращают новый, поэтому
 * проверяются без редактора и не зависят от contenteditable.
 */

export interface MarkdownEditResult {
  ok: boolean;
  text: string;
  message: string;
}

function lineIndex(source: string, find: string): number {
  const lines = source.split('\n');
  if (!find.trim()) {
    return lines.findIndex((line) => line.trim().length > 0);
  }
  const needle = find.trim().toLowerCase();
  return lines.findIndex((line) => line.toLowerCase().includes(needle));
}

export function setMarkdownHeading(
  source: string,
  find: string,
  level: number
): MarkdownEditResult {
  const depth = Math.max(1, Math.min(6, Math.round(Number(level) || 1)));
  const lines = source.split('\n');
  const index = lineIndex(source, find);
  if (index < 0) {
    return { ok: false, text: source, message: `Строка «${find}» не найдена` };
  }

  const indent = /^\s*/.exec(lines[index])?.[0] ?? '';
  const body = lines[index].replace(/^\s{0,3}#{1,6}\s+/, '').trim();
  if (!body) {
    return { ok: false, text: source, message: 'Пустая строка — заголовок не создан' };
  }
  lines[index] = `${indent}${'#'.repeat(depth)} ${body}`;
  return { ok: true, text: lines.join('\n'), message: `Заголовок H${depth} задан` };
}

export function toggleMarkdownList(
  source: string,
  find: string,
  style: 'bullet' | 'numbered'
): MarkdownEditResult {
  const ordered = style === 'numbered';
  const lines = source.split('\n');
  const index = lineIndex(source, find);
  if (index < 0) {
    return { ok: false, text: source, message: `Строка «${find}» не найдена` };
  }

  const line = lines[index];
  const indent = /^\s*/.exec(line)?.[0] ?? '';
  const marker = /^\s*([-*+]|\d+[.)])\s+/.exec(line);
  const body = line.replace(/^\s*([-*+]|\d+[.)])\s+/, '').trim();
  if (!body) return { ok: false, text: source, message: 'Пустая строка — список не создан' };

  const alreadyOrdered = Boolean(marker && /^\s*\d/.test(line));
  const alreadyBullet = Boolean(marker && /^\s*[-*+]/.test(line));
  const sameKind = ordered ? alreadyOrdered : alreadyBullet;

  lines[index] = sameKind ? `${indent}${body}` : `${indent}${ordered ? '1.' : '-'} ${body}`;
  return {
    ok: true,
    text: lines.join('\n'),
    message: sameKind ? 'Пункт списка убран' : ordered ? 'Нумерованный список' : 'Маркированный список',
  };
}

export function insertMarkdownLink(
  source: string,
  find: string,
  label: string,
  url: string
): MarkdownEditResult {
  const target = url.trim();
  if (!target) return { ok: false, text: source, message: 'Пустой URL' };
  if (!/^(https?:\/\/|mailto:|#|\/|\.{1,2}\/)/i.test(target)) {
    return { ok: false, text: source, message: `Недопустимый URL «${url}»` };
  }

  const safeLabel = (label.trim() || find.trim() || target).replace(/([[\]])/g, '\\$1');
  const link = `[${safeLabel}](${target})`;
  if (!find.trim()) {
    const next = `${source.replace(/\s*$/, '')}\n\n${link}\n`;
    return { ok: true, text: next, message: 'Ссылка добавлена в конец' };
  }

  const at = source.indexOf(find);
  if (at < 0) {
    return { ok: false, text: source, message: `Текст «${find}» не найден` };
  }
  return {
    ok: true,
    text: source.slice(0, at) + link + source.slice(at + find.length),
    message: 'Ссылка вставлена',
  };
}

function tableCell(value: string): string {
  return value.replace(/\r?\n/g, ' ').replace(/\|/g, '\\|').trim();
}

export function insertMarkdownTable(
  source: string,
  headers: string[],
  rows: string[][]
): MarkdownEditResult {
  const cleanHeaders = headers.map(tableCell).filter(Boolean);
  if (cleanHeaders.length === 0) {
    return { ok: false, text: source, message: 'Не указаны заголовки таблицы' };
  }

  const width = cleanHeaders.length;
  const headerRow = `| ${cleanHeaders.join(' | ')} |`;
  const separator = `| ${cleanHeaders.map(() => '---').join(' | ')} |`;
  const bodyRows = rows.map((row) => {
    const cells = Array.from({ length: width }, (_, index) => tableCell(String(row[index] ?? '')));
    return `| ${cells.join(' | ')} |`;
  });
  const table = [headerRow, separator, ...bodyRows].join('\n');
  const next = `${source.replace(/\s*$/, '')}\n\n${table}\n`;
  return {
    ok: true,
    text: next,
    message: `Таблица ${width}×${bodyRows.length} добавлена`,
  };
}
