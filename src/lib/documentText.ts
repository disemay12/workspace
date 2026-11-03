import * as XLSX from 'xlsx';
import { OpenTab } from '../types';
import { docxBase64ToHtml } from './docx';

/**
 * Содержимое активной вкладки в виде текста для модели.
 *
 * Word-документ хранится во вкладке как base64-пакет .docx, поэтому модели
 * нужен разбор в HTML и обратная сборка в Markdown: иначе в запрос уходит
 * ZIP в кодировке base64, и ассистент отвечает, что «не может прочитать
 * документ Word».
 */

/** HTML полотна Word → Markdown-подобный текст с заголовками и списками. */
export function htmlToPlainText(html: string): string {
  if (!html.trim()) return '';
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return blocksToText(doc.body).trim();
}

function inlineText(node: Node): string {
  let out = '';
  for (const child of [...node.childNodes]) {
    if (child.nodeType === Node.TEXT_NODE) {
      out += child.textContent ?? '';
      continue;
    }
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const el = child as Element;
    const tag = el.tagName.toLowerCase();
    if (tag === 'br') {
      out += '\n';
      continue;
    }
    const inner = inlineText(el);
    if (!inner) continue;
    if (tag === 'strong' || tag === 'b') out += `**${inner}**`;
    else if (tag === 'em' || tag === 'i') out += `*${inner}*`;
    else if (tag === 'u') out += `__${inner}__`;
    else if (tag === 's' || tag === 'strike' || tag === 'del') out += `~~${inner}~~`;
    else out += inner;
  }
  return out.replace(/\u00a0/g, ' ');
}

function blocksToText(parent: Node): string {
  const blocks: string[] = [];

  for (const child of [...parent.childNodes]) {
    if (child.nodeType === Node.TEXT_NODE) {
      const text = (child.textContent ?? '').trim();
      if (text) blocks.push(text);
      continue;
    }
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const el = child as Element;
    const tag = el.tagName.toLowerCase();

    if (/^h[1-6]$/.test(tag)) {
      const text = inlineText(el).trim();
      if (text) blocks.push(`${'#'.repeat(Number(tag[1]))} ${text}`);
      continue;
    }

    if (tag === 'ul' || tag === 'ol') {
      const items = [...el.children]
        .filter((li) => li.tagName.toLowerCase() === 'li')
        .map((li, index) => `${tag === 'ol' ? `${index + 1}. ` : '- '}${inlineText(li).trim()}`);
      if (items.length) blocks.push(items.join('\n'));
      continue;
    }

    if (tag === 'table') {
      const rows = [...el.querySelectorAll('tr')].map((tr) =>
        [...tr.querySelectorAll('td, th')].map((cell) =>
          inlineText(cell).trim().replace(/\|/g, '\\|')
        )
      );
      if (rows.length) {
        const width = Math.max(...rows.map((row) => row.length));
        const grid = rows.map((row) => [...row, ...Array(width - row.length).fill('')]);
        const head = `| ${grid[0].join(' | ')} |`;
        const divider = `| ${Array(width).fill('---').join(' | ')} |`;
        const body = grid.slice(1).map((row) => `| ${row.join(' | ')} |`);
        blocks.push([head, divider, ...body].join('\n'));
      }
      continue;
    }

    if (tag === 'blockquote') {
      const text = inlineText(el).trim();
      if (text) blocks.push(`> ${text}`);
      continue;
    }

    if (tag === 'hr') {
      blocks.push('---');
      continue;
    }

    if (tag === 'p' || tag === 'div' || tag === 'section' || tag === 'article') {
      const text = inlineText(el).trim();
      if (text) blocks.push(text);
      continue;
    }

    const nested = blocksToText(el).trim();
    if (nested) blocks.push(nested);
  }

  return blocks.join('\n\n');
}

/** Текст Word-документа (base64 .docx) для прикрепления к запросу. */
export async function docxText(base64: string): Promise<string> {
  if (!base64.trim()) return '';
  try {
    const { html } = await docxBase64ToHtml(base64);
    return htmlToPlainText(html);
  } catch {
    return '';
  }
}

/** Текст активной вкладки: Word, таблица или обычный текстовый файл. */
export async function tabToDocumentText(tab: OpenTab): Promise<string> {
  if (tab.type === 'word') return await docxText(tab.content);

  if (tab.type === 'excel') {
    if (tab.encoding !== 'base64') return tab.content;
    try {
      const workbook = XLSX.read(tab.content, { type: 'base64' });
      return workbook.SheetNames.map(
        (name) => `### Лист: ${name}\n${XLSX.utils.sheet_to_csv(workbook.Sheets[name])}`
      ).join('\n\n');
    } catch {
      return '';
    }
  }

  if (tab.type === 'markdown' || tab.type === 'richtext' || tab.type === 'text' || tab.type === 'code') {
    return tab.content;
  }

  return '';
}
