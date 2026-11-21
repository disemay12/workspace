/**
 * Поля колонтитулов: PAGE и NUMPAGES в .docx кэшируются, и без подстановки
 * номер страницы в шапке остаётся тем, что было в файле.
 */
import type { DocxHeaderFooterHtml } from './docx';

const PAGE_FIELD = /<span data-field="PAGE">[\s\S]*?<\/span>/gi;
const NUMPAGES_FIELD = /<span data-field="NUMPAGES">[\s\S]*?<\/span>/gi;

export function fillHeaderFooterFields(html: string, pageNumber: number, totalPages: number): string {
  if (!html) return html;
  return html
    .replace(PAGE_FIELD, String(pageNumber))
    .replace(NUMPAGES_FIELD, String(totalPages));
}

/**
 * Колонтитул для конкретного листа: первый лист может иметь свой (титульный),
 * чётные — отдельный. Word берёт «чётный» только при включённом режиме
 * Different Odd & Even Pages; без настроек sectPr используем default.
 */
export function pickHeaderFooter(
  parts: DocxHeaderFooterHtml | undefined,
  pageNumber: number
): string | undefined {
  if (!parts) return undefined;
  if (pageNumber === 1 && parts.first) return parts.first;
  return parts.default ?? parts.first;
}

/** HTML колонтитула → текст с плейсхолдерами {{page}}/{{pages}} для правки. */
export function headerFooterHtmlToText(html?: string): string {
  if (!html) return '';
  const holder = document.createElement('div');
  holder.innerHTML = html;
  for (const field of [...holder.querySelectorAll('span[data-field]')]) {
    const kind = (field.getAttribute('data-field') ?? '').toUpperCase();
    const placeholder = kind === 'NUMPAGES' ? '{{pages}}' : kind === 'PAGE' ? '{{page}}' : '';
    field.replaceWith(document.createTextNode(placeholder));
  }
  for (const br of [...holder.querySelectorAll('br')]) br.replaceWith(document.createTextNode('\n'));
  for (const block of [...holder.querySelectorAll('p,div,li')]) block.append(document.createTextNode('\n'));
  return (holder.textContent ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/\n{2,}/g, '\n')
    .trim();
}

/** Текст правки колонтитула → HTML для слоя листов (поля остаются полями). */
export function headerFooterTextToHtml(text: string): string {
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  return lines
    .map((line) => {
      const escaped = line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const withFields = escaped
        .replace(/\{\{\s*page\s*\}\}/gi, '<span data-field="PAGE">1</span>')
        .replace(/\{\{\s*pages\s*\}\}/gi, '<span data-field="NUMPAGES">1</span>');
      return `<p style="margin:0">${withFields || '&nbsp;'}</p>`;
    })
    .join('');
}
