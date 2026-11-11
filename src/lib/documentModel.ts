/**
 * Программная модель HTML-документа Word.
 *
 * Полотно Word — contenteditable с обычным HTML, без собственной модели блоков.
 * Здесь HTML разбирается в блоки с детерминированными идентификаторами
 * (b1, b2, …), над ними выполняются операции ассистента, и HTML собирается
 * обратно. Идентификаторы не пишутся в .docx: при сохранении атрибуты
 * игнорируются, при открытии файла выдаются заново в том же порядке.
 */
import {
  extractDocumentContent,
  isAssistantReportText,
  isDocumentMetaText,
  markdownToDocumentHtml,
} from './aiDocument';
import type { PageSetup } from '../types';
import {
  customStyleCss,
  getCustomStyle,
  listCustomStyles,
  registerCustomStyle,
  styleIdFor,
  type CustomStyleDefinition,
} from './customStyles';

export const DOCUMENT_TOOL_NAMES = [
  'get_document',
  'get_outline',
  'get_block',
  'get_range',
  'search_document',
  'get_selection',
  'get_document_info',
  'replace_document',
  'replace_block',
  'insert_after',
  'insert_before',
  'append_to_block',
  'move_block',
  'duplicate_block',
  'split_block',
  'merge_blocks',
  'table_read',
  'table_insert_row',
  'table_delete_row',
  'table_insert_column',
  'table_delete_column',
  'table_set_cell',
  'table_merge_cells',
  'table_unmerge_cells',
  'table_set_column_width',
  'table_set_borders',
  'insert_image',
  'update_image',
  'delete_image',
  'insert_page_break',
  'insert_section_break',
  'set_paragraph_format',
  'set_tabs',
  'set_paragraph_style',
  'define_style',
  'set_header_footer',
  'delete_header_footer',
  'delete_section_break',
  'insert_toc',
  'refresh_toc',
  'insert_field',
  'insert_bookmark',
  'add_comment',
  'add_footnote',
  'delete_block',
  'find_replace',
  'replace_range',
  'delete_range',
  'replace_selection',
  'insert_at_cursor',
  'format',
  'page_setup',
] as const;

export type DocumentToolName = (typeof DOCUMENT_TOOL_NAMES)[number];

/** Инструменты, меняющие документ. Без прав записи они не передаются модели. */
export const DOCUMENT_WRITE_TOOLS = new Set<string>([
  'replace_document',
  'replace_block',
  'insert_after',
  'insert_before',
  'append_to_block',
  'move_block',
  'duplicate_block',
  'split_block',
  'merge_blocks',
  'table_insert_row',
  'table_delete_row',
  'table_insert_column',
  'table_delete_column',
  'table_set_cell',
  'table_merge_cells',
  'table_unmerge_cells',
  'table_set_column_width',
  'table_set_borders',
  'insert_image',
  'update_image',
  'delete_image',
  'insert_page_break',
  'insert_section_break',
  'set_paragraph_format',
  'set_tabs',
  'set_paragraph_style',
  'define_style',
  'set_header_footer',
  'delete_header_footer',
  'delete_section_break',
  'insert_toc',
  'refresh_toc',
  'insert_field',
  'insert_bookmark',
  'add_comment',
  'add_footnote',
  'delete_block',
  'find_replace',
  'replace_range',
  'delete_range',
  'replace_selection',
  'insert_at_cursor',
  'format',
  'page_setup',
  'excel_write_cells',
  'excel_add_sheet',
  'excel_rename_sheet',
  'excel_delete_sheet',
  'excel_insert_rows',
  'excel_delete_rows',
  'excel_insert_columns',
  'excel_delete_columns',
  'excel_merge_cells',
  'excel_unmerge_cells',
  'excel_format_range',
  'pptx_set_text',
  'pptx_add_slide',
  'pptx_delete_slide',
  'pptx_move_slide',
  'pptx_add_text_box',
  'md_replace',
  'md_append',
  'md_replace_document',
  'md_set_heading',
  'md_toggle_list',
  'md_insert_link',
  'md_insert_table',
  'write_document',
]);

/** Типы вкладок, для которых доступны инструменты документа. */
export const DOCUMENT_TAB_TYPES = new Set<string>(['word', 'markdown', 'richtext', 'text']);

/** Типы, для которых работает агент инструментов (Word/Markdown + Excel). */
export const AGENT_TAB_TYPES = new Set<string>([...DOCUMENT_TAB_TYPES, 'excel', 'powerpoint']);

/** Полноценные блочные операции поддерживает только HTML-полотно Word. */
export const BLOCK_TOOL_TAB_TYPES = new Set<string>(['word']);

export interface DocumentToolResult {
  ok: boolean;
  message: string;
  html?: string;
  changed?: boolean;
  structure?: string;
  /** HTML конкретного блока/диапазона (для get_block/get_range). */
  content?: string;
  /** Точные совпадения поиска. */
  matches?: Array<{
    block_id: string;
    tag: string;
    start: number;
    end: number;
    match: string;
    context: string;
  }>;
  /** Служебные данные пагинации/выделения. */
  meta?: Record<string, unknown>;
  /** Путь созданного или изменённого файла (конвертация, медиа, экспорт). */
  path?: string;
  /** Новые параметры страницы: применяются к .docx через контроллер редактора. */
  page?: PageSetup;
  selection?: {
    blockId: string;
    startOffset: number;
    endOffset: number;
    text: string;
    html: string;
  };
}

type StyleMap = Record<string, string>;

const ALLOWED_TAGS = new Set([
  'p',
  'div',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'ul',
  'ol',
  'li',
  'table',
  'thead',
  'tbody',
  'tfoot',
  'tr',
  'td',
  'th',
  'blockquote',
  'hr',
  'br',
  'strong',
  'b',
  'em',
  'i',
  'u',
  's',
  'strike',
  'del',
  'sub',
  'sup',
  'span',
  'img',
  'a',
]);

const ALLOWED_CSS = new Set([
  'font-family',
  'font-size',
  'font-weight',
  'font-style',
  'text-align',
  'text-decoration',
  'text-indent',
  'color',
  'background-color',
  'margin-left',
  'padding-left',
  'line-height',
  'width',
  'height',
  'max-width',
  'margin-top',
  'margin-bottom',
  'break-before',
  'break-after',
  'break-inside',
  'border',
  'border-top',
  'border-right',
  'border-bottom',
  'border-left',
]);

const BLOCK_TAGS = new Set([
  'p',
  'div',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'ul',
  'ol',
  'table',
  'blockquote',
  'hr',
]);

/** Создаёт контейнер с очищенным HTML: без скриптов, обработчиков и опасных стилей. */
function sanitize(raw: string): HTMLDivElement {
  const container = document.createElement('div');
  container.innerHTML = raw || '';
  sanitizeElement(container);
  normalizeBlocks(container);
  wrapStrayText(container);
  return container;
}

function sanitizeElement(root: Element): void {
  for (const child of [...root.children]) {
    const tag = child.tagName.toLowerCase();

    if (
      tag === 'script' ||
      tag === 'style' ||
      tag === 'iframe' ||
      tag === 'object' ||
      tag === 'embed' ||
      tag === 'link' ||
      tag === 'meta' ||
      tag === 'form' ||
      tag === 'input' ||
      tag === 'button'
    ) {
      child.remove();
      continue;
    }

    if (!ALLOWED_TAGS.has(tag)) {
      // Сначала чистим вложенные элементы, потом поднимаем их на уровень выше.
      sanitizeElement(child);
      unwrap(child);
      continue;
    }

    for (const attribute of [...child.attributes]) {
      const name = attribute.name.toLowerCase();
      if (
        name === 'data-bid' ||
        name === 'data-docx-id' ||
        name === 'data-docx-sig' ||
        name === 'data-docx-readonly' ||
        name === 'data-docx-kind' ||
        name === 'data-page-break' ||
        name === 'data-section-break' ||
        name === 'data-field' ||
        name === 'data-tabs' ||
        name === 'data-style' ||
        name === 'data-widow' ||
        name === 'data-table-border' ||
        name === 'data-bookmark' ||
        name === 'data-run-id' ||
        name === 'data-comment-id' ||
        name === 'data-comment-text' ||
        name === 'data-comment-author' ||
        name === 'data-footnote-id' ||
        name === 'data-footnote-text' ||
        name === 'colspan' ||
        name === 'rowspan' ||
        name === 'alt' ||
        name === 'title' ||
        (tag === 'a' && name === 'href' && isSafeLink(attribute.value)) ||
        (tag === 'img' && (name === 'width' || name === 'height'))
      ) {
        continue;
      }
      if (name === 'style') {
        child.setAttribute('style', filterStyle(attribute.value));
        continue;
      }
      if (name === 'src' && tag === 'img' && /^data:image\//i.test(attribute.value)) continue;
      child.removeAttribute(attribute.name);
    }

    sanitizeElement(child);
  }
}

function unwrap(element: Element): void {
  const parent = element.parentNode;
  if (!parent) {
    element.remove();
    return;
  }
  while (element.firstChild) parent.insertBefore(element.firstChild, element);
  element.remove();
}

/** div — тоже абзац, но в документе и в .docx используется p. */
function normalizeBlocks(container: HTMLElement): void {
  for (const div of [...container.querySelectorAll('div')]) {
    const p = document.createElement('p');
    for (const attribute of [...div.attributes]) p.setAttribute(attribute.name, attribute.value);
    while (div.firstChild) p.appendChild(div.firstChild);
    div.replaceWith(p);
  }
}

/** Текст вне блочных тегов превращается в абзац, иначе он не получит идентификатор. */
function wrapStrayText(container: HTMLElement): void {
  const nodes = [...container.childNodes];
  let buffer: Node[] = [];
  const flush = () => {
    if (buffer.length === 0) return;
    const hasContent = buffer.some(
      (node) =>
        node.nodeType === Node.ELEMENT_NODE ||
        (node.nodeType === Node.TEXT_NODE && (node.textContent ?? '').trim())
    );
    if (hasContent) {
      const anchor = buffer[0];
      const p = document.createElement('p');
      // Вставляем абзац до первого узла, пока он ещё принадлежит контейнеру.
      container.insertBefore(p, anchor);
      buffer.forEach((node) => p.appendChild(node));
    }
    buffer = [];
  };

  for (const node of nodes) {
    const isBlock =
      node.nodeType === Node.ELEMENT_NODE &&
      BLOCK_TAGS.has((node as Element).tagName.toLowerCase());
    if (isBlock) {
      flush();
    } else {
      buffer.push(node);
    }
  }
  flush();
}

function parseStyle(value: string): StyleMap {
  const entries: StyleMap = {};
  for (const part of value.split(';')) {
    const colon = part.indexOf(':');
    if (colon <= 0) continue;
    const name = part.slice(0, colon).trim().toLowerCase();
    const val = part.slice(colon + 1).trim();
    if (!name || !val) continue;
    entries[name] = val;
  }
  return entries;
}

function filterStyle(value: string): string {
  const entries = parseStyle(value);
  const safe: string[] = [];
  for (const [name, val] of Object.entries(entries)) {
    if (!ALLOWED_CSS.has(name)) continue;
    if (/url\s*\(|expression\s*\(|javascript:/i.test(val)) continue;
    safe.push(`${name}:${val}`);
  }
  return safe.join(';');
}

function mergeStyle(element: Element, updates: StyleMap): void {
  const entries = parseStyle(element.getAttribute('style') ?? '');
  for (const [name, val] of Object.entries(updates)) {
    if (val) entries[name] = val;
    else delete entries[name];
  }
  const style = Object.entries(entries)
    .map(([name, val]) => `${name}:${val}`)
    .join(';');
  if (style) element.setAttribute('style', style);
  else element.removeAttribute('style');
}

/**
 * Назначает идентификаторы блокам и элементам списков.
 *
 * Уже существующие id сохраняются: серия правок одного ответа ссылается на
 * одни и те же блоки, даже если соседние удалены или вставлены. Новые узлы
 * получают следующий свободный номер.
 */
function assignBlockIds(container: HTMLElement): void {
  let max = 0;
  for (const element of container.querySelectorAll('[data-bid]')) {
    const match = /^b(\d+)$/.exec(element.getAttribute('data-bid') ?? '');
    if (match) max = Math.max(max, Number(match[1]));
  }
  const next = () => `b${++max}`;

  for (const child of [...container.children]) {
    if (!/^b\d+$/.test(child.getAttribute('data-bid') ?? '')) {
      child.setAttribute('data-bid', next());
    }
    if (child.tagName.toLowerCase() === 'ul' || child.tagName.toLowerCase() === 'ol') {
      for (const item of [...child.children]) {
        if (item.tagName.toLowerCase() !== 'li') continue;
        if (!/^b\d+$/.test(item.getAttribute('data-bid') ?? '')) {
          item.setAttribute('data-bid', next());
        }
      }
    }
  }
}

function byId(container: HTMLElement, id: string): Element | null {
  if (!id) return null;
  return container.querySelector(`[data-bid="${cssEscape(id)}"]`);
}

function cssEscape(value: string): string {
  return value.replace(/["\\]/g, '\\$&');
}

function inlineMarkers(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) return (node.textContent ?? '').replace(/\s+/g, ' ');
  if (node.nodeType !== Node.ELEMENT_NODE) return '';

  const element = node as Element;
  const tag = element.tagName.toLowerCase();
  if (tag === 'br') return '\n';
  if (tag === 'img') return '[изображение]';

  const inner = [...element.childNodes].map(inlineMarkers).join('');
  if (tag === 'strong' || tag === 'b') return `**${inner}**`;
  if (tag === 'em' || tag === 'i') return `*${inner}*`;
  if (tag === 'u') return `__${inner}__`;
  if (tag === 's' || tag === 'strike' || tag === 'del') return `~~${inner}~~`;
  return inner;
}

function textOf(element: Element): string {
  return [...element.childNodes]
    .map(inlineMarkers)
    .join('')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function styleHint(element: Element): string {
  const entries = parseStyle(element.getAttribute('style') ?? '');
  const parts: string[] = [];
  if (entries['text-align']) parts.push(entries['text-align']);
  if (entries['font-family']) {
    parts.push(entries['font-family'].replace(/["']/g, '').split(',')[0].trim());
  }
  if (entries['font-size']) parts.push(entries['font-size']);
  if (entries['font-weight']) parts.push(`weight ${entries['font-weight']}`);
  if (entries['margin-left']) parts.push(`отступ ${entries['margin-left']}`);
  if (entries['text-indent']) parts.push(`красная строка ${entries['text-indent']}`);
  return parts.length > 0 ? ` (${parts.join(', ')})` : '';
}

function blockLabel(element: Element): string {
  const tag = element.tagName.toLowerCase();
  if (tag === 'ul') return 'ul';
  if (tag === 'ol') return 'ol';
  if (tag === 'table') {
    const rows = element.querySelectorAll('tr').length;
    const cols = element.querySelector('tr')?.children.length ?? 0;
    return `table ${rows}×${cols}`;
  }
  return tag;
}

/**
 * Компактное описание структуры для модели. Бюджет ограничен, чтобы
 * результат инструмента не съедал контекст.
 */
export function describeStructure(container: HTMLElement, budget = 12000): string {
  const lines: string[] = [];
  let used = 0;

  for (const child of [...container.children]) {
    const id = child.getAttribute('data-bid') ?? '';
    const label = blockLabel(child);

    if (label === 'ul' || label === 'ol') {
      lines.push(`${id} [${label}]${styleHint(child)}`);
      used += 24;
      for (const item of [...child.children]) {
        if (item.tagName.toLowerCase() !== 'li') continue;
        const itemId = item.getAttribute('data-bid') ?? '';
        const text = truncate(textOf(item), 300);
        lines.push(`  ${itemId} [li]${styleHint(item)} ${text}`);
        used += itemId.length + text.length + 16;
        if (used > budget) return finish(lines, true);
      }
      continue;
    }

    if (label.startsWith('table')) {
      const header = [...(child.querySelector('tr')?.children ?? [])]
        .map((cell) => textOf(cell))
        .join(' | ');
      lines.push(`${id} [${label}]${styleHint(child)} ${truncate(header, 200)}`);
      used += label.length + header.length + 24;
      if (used > budget) return finish(lines, true);
      continue;
    }

    if (label === 'hr') {
      lines.push(`${id} [hr]`);
      used += 12;
      continue;
    }

    const text = truncate(textOf(child), 600);
    lines.push(`${id} [${label}]${styleHint(child)} ${text}`);
    used += text.length + label.length + 16;
    if (used > budget) return finish(lines, true);
  }

  return finish(lines, false);
}

function finish(lines: string[], truncated: boolean): string {
  if (truncated) lines.push('… структура обрезана, вызовите get_document при необходимости');
  return lines.join('\n');
}

function truncate(value: string, max: number): string {
  return value.length <= max ? value : `${value.slice(0, max)}…`;
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char] ?? char);
}

/**
 * Идентификаторы из ответа модели не принимаем: иначе вставка копии блока
 * создаёт дубли data-bid и ломает адресацию.
 */
function cleanInputHtml(raw: string): string {
  const value = raw || '';
  if (!/<[a-z][\s\S]*>/i.test(value)) return value;
  const probe = document.createElement('div');
  probe.innerHTML = value;
  for (const element of [...probe.querySelectorAll('[data-bid]')]) {
    element.removeAttribute('data-bid');
  }
  // Модель не должна подделывать привязку к исходному OOXML: иначе её
  // фрагмент при сохранении выдаст себя за нетронутый блок Word.
  for (const element of [...probe.querySelectorAll('[data-docx-id], [data-docx-sig], [data-docx-readonly], [data-docx-kind]')]) {
    element.removeAttribute('data-docx-id');
    element.removeAttribute('data-docx-sig');
    element.removeAttribute('data-docx-readonly');
    element.removeAttribute('data-docx-kind');
  }
  return probe.innerHTML;
}

function toBlockHtml(raw: string): string {
  const value = (raw || '').trim();
  if (!value) return '';
  if (/<[a-z][\s\S]*>/i.test(value)) return value;
  // Слабая модель иногда присылает Markdown вместо HTML — разбираем его,
  // чтобы заголовки и списки не попали в документ литералами «#» и «-».
  // Заодно отбрасываются вступление чата и обёртка ```.
  const converted = markdownToDocumentHtml(extractDocumentContent(value) || value);
  return converted.trim() || `<p>${escapeHtml(value)}</p>`;
}

/** HTML без блочных обёрток — содержимое элемента списка должно остаться строчным. */
function inlineOnlyHtml(raw: string): string {
  const value = raw.trim();
  if (!value) return '';
  if (!/<[a-z][\s\S]*>/i.test(value)) return escapeHtml(value);

  const container = sanitize(value);
  const parts: string[] = [];
  for (const child of [...container.childNodes]) {
    if (child.nodeType !== Node.ELEMENT_NODE) {
      parts.push(escapeHtml(child.textContent ?? ''));
      continue;
    }
    const element = child as Element;
    const tag = element.tagName.toLowerCase();
    if (tag === 'ul' || tag === 'ol' || tag === 'table' || tag === 'blockquote') {
      parts.push(escapeHtml(textOf(element)));
    } else if (BLOCK_TAGS.has(tag)) {
      parts.push(element.innerHTML);
    } else {
      parts.push(element.outerHTML);
    }
  }
  return parts.join('').trim();
}

/* ============================================================
   Операции
   ============================================================ */

/** Вступление чата: «Понял Вас», «Вот документ…», «Ниже — …». */
const CHAT_PREAMBLE_RE =
  /^(понял|поняла|вот|ниже|готово|конечно|разумеется|держите|представляю|хорошо|с удовольствием)(?![а-яёa-z])/i;

/**
 * Убирает из готового HTML вступление чата и служебные советы по оформлению.
 *
 * Модель иногда кладёт в replace_document весь ответ целиком: абзац
 * «Понял Вас…», список «Рекомендуемые настройки страницы» и только потом
 * текст документа. В .docx эти абзацы не нужны.
 */
function isChangeReportList(element: Element): boolean {
  const items = [...element.querySelectorAll(':scope > li')];
  if (items.length === 0) return false;
  const changed = items.filter((item) =>
    /(→|->|(?:обновлен|переименован|переписан|заменен|добавлен|удален|исправлен|сокращен|настроен))/i.test(
      textOf(item)
    )
  ).length;
  return changed >= Math.ceil(items.length * 0.6);
}

function cleanChatNoise(container: HTMLElement): void {
  const children = [...container.children];
  if (children.length === 0) return;

  const removable = new Set<Element>();
  let leading = 0;
  let afterMeta = false;
  let reportContext = false;

  for (const child of children) {
    const tag = child.tagName.toLowerCase();
    const text = textOf(child).replace(/^[#>*\-\s]+/, '').trim();
    if (!text) continue;

    if (isDocumentMetaText(text)) {
      removable.add(child);
      afterMeta = true;
      continue;
    }
    if (isAssistantReportText(text)) {
      removable.add(child);
      afterMeta = false;
      reportContext = true;
      continue;
    }
    // Список полей/шрифтов идёт отдельным блоком от своего заголовка.
    if (afterMeta && (tag === 'ul' || tag === 'ol')) {
      removable.add(child);
      continue;
    }
    // Список изменений «Заголовок → …» после отчёта ассистента.
    if (
      reportContext &&
      (tag === 'ul' || tag === 'ol') &&
      isChangeReportList(child)
    ) {
      removable.add(child);
      continue;
    }
    afterMeta = false;
    reportContext = false;

    if (leading < 2 && CHAT_PREAMBLE_RE.test(text)) {
      removable.add(child);
      leading += 1;
      continue;
    }
    leading = 2;
  }

  for (const element of removable) element.remove();

  // Разделители «---» из ответа чата остаются пустыми линиями в .docx.
  for (const rule of [...container.querySelectorAll(':scope > hr')]) {
    const prev = rule.previousElementSibling;
    const next = rule.nextElementSibling;
    if (!prev || !next || prev.tagName.toLowerCase() === 'hr' || next.tagName.toLowerCase() === 'hr') {
      rule.remove();
    }
  }
}

function opGetDocument(container: HTMLElement): DocumentToolResult {
  assignBlockIds(container);
  return {
    ok: true,
    message: 'Структура документа получена',
    changed: false,
    structure: describeStructure(container, 20000),
  };
}

/** Пагинация структуры: длинный документ не обязан влезать в один ответ. */
export function describeStructurePage(
  container: HTMLElement,
  offset: number,
  limit: number
): { text: string; total: number; nextOffset: number | null } {
  const children = [...container.children];
  const total = children.length;
  const start = Math.max(0, Math.min(offset, total));
  const end = Math.max(start, Math.min(start + limit, total));
  const lines: string[] = [];

  for (const child of children.slice(start, end)) {
    const id = child.getAttribute('data-bid') ?? '';
    const label = blockLabel(child);

    if (label === 'ul' || label === 'ol') {
      lines.push(`${id} [${label}]${styleHint(child)}`);
      for (const item of [...child.children]) {
        if (item.tagName.toLowerCase() !== 'li') continue;
        const itemId = item.getAttribute('data-bid') ?? '';
        lines.push(`  ${itemId} [li]${styleHint(item)} ${truncate(textOf(item), 300)}`);
      }
      continue;
    }

    if (label.startsWith('table')) {
      const header = [...(child.querySelector('tr')?.children ?? [])]
        .map((cell) => textOf(cell))
        .join(' | ');
      lines.push(`${id} [${label}]${styleHint(child)} ${truncate(header, 200)}`);
      continue;
    }

    lines.push(`${id} [${label}]${styleHint(child)} ${truncate(textOf(child), 600)}`);
  }

  return { text: lines.join('\n'), total, nextOffset: end < total ? end : null };
}

function opGetOutline(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  assignBlockIds(container);
  const offset = Math.max(0, Math.floor(Number(args.offset ?? args.start) || 0));
  const limit = Math.max(1, Math.min(100, Math.floor(Number(args.limit ?? args.count) || 40)));
  const page = describeStructurePage(container, offset, limit);
  const shownTo = Math.min(offset + limit, page.total);
  return {
    ok: true,
    message:
      page.nextOffset === null
        ? `Блоки ${offset + 1}–${shownTo} из ${page.total}. Это весь документ.`
        : `Блоки ${offset + 1}–${shownTo} из ${page.total}. Продолжение: get_outline с offset=${page.nextOffset}.`,
    changed: false,
    structure: page.text,
    meta: { offset, total: page.total, nextOffset: page.nextOffset },
  };
}

function opGetBlock(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };

  const label = blockLabel(target);
  const text = truncate(textOf(target), 2000);
  return {
    ok: true,
    message: `Блок ${id} (${label})`,
    changed: false,
    structure: `${id} [${label}]${styleHint(target)} ${text}`,
    content: target.outerHTML,
  };
}

function opGetRange(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };

  const full = target.textContent ?? '';
  const start = clampOffset(args.start ?? args.from, 0, full.length);
  const end = clampOffset(args.end ?? args.to, start, full.length);
  const range = rangeFromOffsets(target, start, end);
  const content = range ? fragmentHtml(range) : '';
  const text = full.slice(start, end);

  return {
    ok: true,
    message: `Блок ${id}, символы ${start}–${end}: «${truncate(text, 300)}»`,
    changed: false,
    content,
    meta: { blockId: id, start, end, length: full.length },
  };
}

function opSearchDocument(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  assignBlockIds(container);
  const query = asString(args.query ?? args.find ?? args.text);
  if (!query) return { ok: false, message: 'Не указан текст для поиска' };
  const matchCase = asBool(args.match_case) || asBool(args.matchCase);
  const limit = Math.max(1, Math.min(100, Math.floor(Number(args.limit) || 20)));
  const needle = matchCase ? query : query.toLowerCase();
  const matches: NonNullable<DocumentToolResult['matches']> = [];

  for (const block of [...container.children]) {
    const text = block.textContent ?? '';
    const haystack = matchCase ? text : text.toLowerCase();
    let from = 0;
    while (matches.length < limit) {
      const index = haystack.indexOf(needle, from);
      if (index < 0) break;
      const end = index + query.length;
      const contextStart = Math.max(0, index - 40);
      const contextEnd = Math.min(text.length, end + 40);
      matches.push({
        block_id: block.getAttribute('data-bid') ?? '',
        tag: blockLabel(block),
        start: index,
        end,
        match: text.slice(index, end),
        context: text.slice(contextStart, contextEnd),
      });
      from = Math.max(index + 1, end);
    }
    if (matches.length >= limit) break;
  }

  if (matches.length === 0) return { ok: false, message: `Текст «${query}» не найден` };
  return {
    ok: true,
    message: `Найдено совпадений: ${matches.length}${matches.length >= limit ? ' (лимит)' : ''}`,
    changed: false,
    matches,
    structure: matches
      .map(
        (match) =>
          `${match.block_id} [${match.tag}] ${match.start}–${match.end}: …${match.context}…`
      )
      .join('\n'),
  };
}

function opReplaceRange(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; диапазон менять нельзя' };
  }

  const raw = asString(args.html ?? args.content ?? args.text);
  if (!raw.trim()) return { ok: false, message: 'Пустой HTML для замены диапазона' };
  const replacement = sanitizeSelectionHtml(raw);
  if (!replacement) return { ok: false, message: 'После очистки не осталось текста документа' };

  return replaceRangeInBlock(container, target, args, replacement);
}

function opDeleteRange(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; диапазон менять нельзя' };
  }
  return replaceRangeInBlock(container, target, args, '');
}

function replaceRangeInBlock(
  container: HTMLElement,
  target: Element,
  args: Record<string, unknown>,
  replacementHtml: string
): DocumentToolResult {
  const id = target.getAttribute('data-bid') ?? '';
  const full = target.textContent ?? '';
  const start = clampOffset(args.start ?? args.from, 0, full.length);
  const end = clampOffset(args.end ?? args.to, start, full.length);
  if (start === end && !replacementHtml) {
    return { ok: false, message: 'Пустой диапазон: нечего заменять или удалять' };
  }

  let range = rangeFromOffsets(target, start, end);
  if (!range) {
    // Пустой блок (например, <p><br></p>): создаём точку вставки.
    range = document.createRange();
    range.selectNodeContents(target);
    range.collapse(false);
  }

  if (replacementHtml) {
    const template = document.createElement('template');
    template.innerHTML = replacementHtml;
    range.deleteContents();
    range.insertNode(template.content);
  } else {
    range.deleteContents();
  }

  if ((target.textContent ?? '').length === 0) {
    // Word не хранит полностью пустой абзац: оставляем <br>, чтобы блок жил.
    target.innerHTML = '<br>';
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `${replacementHtml ? 'Диапазон заменён' : 'Диапазон удалён'}: ${id}, символы ${start}–${end}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

/** HTML, допустимый внутри абзаца: блочные обёртки разворачиваются в inline. */
export function sanitizeSelectionHtml(raw: string): string {
  const container = sanitize(toBlockHtml(cleanInputHtml(raw)));
  cleanChatNoise(container);
  const parts: string[] = [];
  for (const node of [...container.childNodes]) {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? '';
      if (text.trim()) parts.push(escapeHtml(text));
      continue;
    }
    if (!(node instanceof Element)) continue;
    const tag = node.tagName.toLowerCase();
    if (BLOCK_TAGS.has(tag)) parts.push(node.innerHTML);
    else parts.push(node.outerHTML);
  }
  return parts.join('<br>');
}

function clampOffset(value: unknown, min: number, max: number): number {
  const parsed = Math.floor(Number(value));
  if (!Number.isFinite(parsed)) return min;
  return Math.max(min, Math.min(parsed, max));
}

function textNodesIn(root: Node): Text[] {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  let current = walker.nextNode();
  while (current) {
    nodes.push(current as Text);
    current = walker.nextNode();
  }
  return nodes;
}

function rangeFromOffsets(target: Element, start: number, end: number): Range | null {
  const nodes = textNodesIn(target);
  if (nodes.length === 0) return null;

  const pointAt = (offset: number): { node: Text; offset: number } => {
    let position = 0;
    for (const node of nodes) {
      const length = (node.textContent ?? '').length;
      if (offset <= position + length) {
        return { node, offset: offset - position };
      }
      position += length;
    }
    const last = nodes[nodes.length - 1];
    return { node: last, offset: (last.textContent ?? '').length };
  };

  if (start === end) {
    const point = pointAt(start);
    const collapsed = document.createRange();
    collapsed.setStart(point.node, point.offset);
    collapsed.setEnd(point.node, point.offset);
    return collapsed;
  }

  let startNode: Text | null = null;
  let startOffset = 0;
  let endNode: Text | null = null;
  let endOffset = 0;
  let position = 0;

  for (const node of nodes) {
    const length = (node.textContent ?? '').length;
    if (!startNode && start < position + length) {
      startNode = node;
      startOffset = start - position;
    }
    if (!endNode && end <= position + length && end > position) {
      endNode = node;
      endOffset = end - position;
    }
    position += length;
  }

  if (!startNode) {
    startNode = nodes[nodes.length - 1];
    startOffset = (startNode.textContent ?? '').length;
  }
  if (!endNode) {
    endNode = nodes[nodes.length - 1];
    endOffset = (endNode.textContent ?? '').length;
  }

  const range = document.createRange();
  range.setStart(startNode, startOffset);
  range.setEnd(endNode, endOffset);
  return range;
}

function fragmentHtml(range: Range): string {
  const container = document.createElement('div');
  container.appendChild(range.cloneContents());
  return container.innerHTML;
}

/** HTML диапазона для моста выделения. */
export function selectionRangeHtml(range: Range): string {
  return fragmentHtml(range);
}

/** Символьное смещение узла внутри блока: используется мостом выделения. */
export function offsetWithinBlock(block: Element, node: Node, offset: number): number | null {
  if (!block.contains(node) && node !== block) return null;
  const probe = document.createRange();
  probe.setStart(block, 0);
  probe.setEnd(node, offset);
  return (probe.cloneContents().textContent ?? '').length;
}

function normalizeDocumentText(root: Element): string {
  return textOf(root).replace(/\s+/g, ' ').trim();
}

function opReplaceDocument(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const raw = asString(args.html ?? args.content ?? args.text);
  if (!raw.trim()) return { ok: false, message: 'Пустой HTML документа' };

  const next = sanitize(toBlockHtml(cleanInputHtml(raw)));
  cleanChatNoise(next);
  if (next.children.length === 0) return { ok: false, message: 'После очистки HTML документ пуст' };

  // «Заменил» тем же текстом — не замена. Это ловит копию резюме вместо
  // реальной ATS-переработки и заставляет агента повторить правку.
  if (normalizeDocumentText(container) === normalizeDocumentText(next)) {
    return {
      ok: false,
      message:
        'Содержимое документа не изменилось. Перепиши текст по существу (или используй find_replace/replace_block для точечной правки).',
    };
  }

  // Картинки модель не воспроизводит: если новый документ их не содержит,
  // возвращаем прежние в конец, чтобы переписывание не удаляло изображения.
  const oldImages = [...container.querySelectorAll('img')];
  const newImages = next.querySelectorAll('img').length;
  let restored = 0;
  if (newImages === 0 && oldImages.length > 0) {
    for (const image of oldImages) {
      const src = image.getAttribute('src') ?? '';
      if (!src) continue;
      const p = document.createElement('p');
      const img = document.createElement('img');
      img.setAttribute('src', src);
      if (image.getAttribute('alt')) img.setAttribute('alt', image.getAttribute('alt')!);
      p.appendChild(img);
      next.appendChild(p);
      restored += 1;
    }
  }

  // Read-only объекты Word (SmartArt, content controls, сложные поля) нельзя
  // потерять при полной переработке текста: возвращаем их в документ.
  const readOnlyBlocks = [...container.querySelectorAll('[data-docx-readonly]')];

  container.innerHTML = next.innerHTML;
  for (const block of readOnlyBlocks) container.appendChild(block);
  assignBlockIds(container);
  return {
    ok: true,
    message:
      (restored > 0 ? `Документ заменён, изображений сохранено: ${restored}` : 'Документ заменён') +
      (readOnlyBlocks.length > 0 ? `; объектов только для просмотра сохранено: ${readOnlyBlocks.length}` : ''),
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function isReadOnlyBlock(element: Element | null): boolean {
  return !!element?.closest('[data-docx-readonly]');
}

function opReplaceBlock(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return {
      ok: false,
      message: 'Этот блок содержит объект Word, доступный только для просмотра; его нельзя заменить текстом',
    };
  }

  const raw = asString(args.html ?? args.content ?? args.text);
  if (!raw.trim()) return { ok: false, message: 'Пустой HTML блока' };

  if (target.tagName.toLowerCase() === 'li') {
    const clean = inlineOnlyHtml(cleanInputHtml(raw));
    if (isAssistantReportText(clean)) {
      return {
        ok: false,
        message: 'Фрагмент содержит только служебный отчёт ассистента, а не текст документа',
      };
    }
    target.innerHTML = clean;
  } else {
    const next = sanitize(toBlockHtml(cleanInputHtml(raw)));
    cleanChatNoise(next);
    if (next.children.length === 0) {
      return {
        ok: false,
        message: 'Фрагмент содержит только служебный отчёт ассистента, а не текст документа',
      };
    }
    for (const node of [...next.childNodes]) target.parentNode?.insertBefore(node, target);
    target.remove();
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Блок ${id} заменён`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opInsertAfter(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };

  const raw = asString(args.html ?? args.content ?? args.text);
  if (!raw.trim()) return { ok: false, message: 'Пустой HTML для вставки' };

  const next = sanitize(toBlockHtml(cleanInputHtml(raw)));
  cleanChatNoise(next);
  if (next.children.length === 0) {
    return {
      ok: false,
      message: 'Вставка содержит только служебный отчёт ассистента, а не текст документа',
    };
  }

  const isListItem = target.tagName.toLowerCase() === 'li';
  const listParent = isListItem ? target.parentElement : null;
  const allItems = [...next.children].every((node) => node.tagName.toLowerCase() === 'li');

  if (isListItem && listParent && allItems) {
    for (const node of [...next.children]) listParent.insertBefore(node, target.nextSibling);
  } else {
    const anchor = isListItem ? listParent ?? target : target;
    for (const node of [...next.children]) anchor.parentNode?.insertBefore(node, anchor.nextSibling);
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Вставлено после ${id}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opInsertBefore(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };

  const raw = asString(args.html ?? args.content ?? args.text);
  if (!raw.trim()) return { ok: false, message: 'Пустой HTML для вставки' };

  const next = sanitize(toBlockHtml(cleanInputHtml(raw)));
  cleanChatNoise(next);
  if (next.children.length === 0) {
    return {
      ok: false,
      message: 'Вставка содержит только служебный отчёт ассистента, а не текст документа',
    };
  }

  const isListItem = target.tagName.toLowerCase() === 'li';
  const listParent = isListItem ? target.parentElement : null;
  const allItems = [...next.children].every((node) => node.tagName.toLowerCase() === 'li');

  if (isListItem && listParent && allItems) {
    for (const node of [...next.children]) listParent.insertBefore(node, target);
  } else {
    const anchor = isListItem ? listParent ?? target : target;
    for (const node of [...next.children]) anchor.parentNode?.insertBefore(node, anchor);
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Вставлено перед ${id}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opAppendToBlock(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; дописывать в него нельзя' };
  }

  const raw = asString(args.html ?? args.content ?? args.text);
  if (!raw.trim()) return { ok: false, message: 'Пустой HTML для дописывания' };

  const tag = target.tagName.toLowerCase();
  if (tag === 'ul' || tag === 'ol') {
    const next = sanitize(toBlockHtml(cleanInputHtml(raw)));
    cleanChatNoise(next);
    const items = [...next.children].filter((node) => node.tagName.toLowerCase() === 'li');
    if (items.length === 0) {
      return { ok: false, message: 'Для списка нужен HTML с <li>' };
    }
    for (const item of items) target.appendChild(item);
  } else {
    const inline = sanitizeSelectionHtml(raw);
    if (!inline) {
      return {
        ok: false,
        message: 'Фрагмент содержит только служебный отчёт ассистента, а не текст документа',
      };
    }
    const template = document.createElement('template');
    template.innerHTML = inline;
    if (target.innerHTML.trim() === '<br>') target.innerHTML = '';
    target.appendChild(template.content);
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Дописано в конец ${id}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opMoveBlock(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; перемещать его нельзя' };
  }

  const beforeId = asString(args.before_id ?? args.beforeId);
  const afterId = asString(args.after_id ?? args.afterId);
  const anchor = byId(container, beforeId || afterId);
  if (beforeId || afterId) {
    if (!anchor) return { ok: false, message: `Блок ${beforeId || afterId} не найден` };
    if (anchor === target || target.contains(anchor)) {
      return { ok: false, message: 'Нельзя переместить блок относительно самого себя' };
    }
    if (beforeId) anchor.parentNode?.insertBefore(target, anchor);
    else anchor.parentNode?.insertBefore(target, anchor.nextSibling);
  } else {
    const index = Math.floor(Number(args.index ?? args.to_index ?? args.toIndex));
    if (!Number.isFinite(index)) {
      return { ok: false, message: 'Укажите before_id, after_id или index' };
    }
    const siblings = [...container.children];
    const bounded = Math.max(0, Math.min(index, siblings.length));
    container.insertBefore(target, siblings[bounded] ?? null);
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Блок ${id} перемещён`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opDuplicateBlock(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };

  const clone = target.cloneNode(true) as Element;
  for (const node of [clone, ...clone.querySelectorAll('*')]) {
    node.removeAttribute('data-bid');
    node.removeAttribute('data-docx-id');
    node.removeAttribute('data-docx-sig');
  }
  target.parentNode?.insertBefore(clone, target.nextSibling);
  assignBlockIds(container);
  return {
    ok: true,
    message: `Блок ${id} продублирован`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opSplitBlock(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; разделять его нельзя' };
  }

  const full = target.textContent ?? '';
  const offset = clampOffset(args.offset ?? args.at ?? args.position, 0, full.length);
  if (offset <= 0 || offset >= full.length) {
    return { ok: false, message: `Смещение ${offset} должно быть внутри блока (0–${full.length})` };
  }

  const range = rangeFromOffsets(target, offset, full.length);
  if (!range) return { ok: false, message: 'Не удалось определить точку разделения' };
  const tail = range.extractContents();

  const tag = target.tagName.toLowerCase();
  const split = document.createElement(tag === 'li' ? 'li' : tag === 'td' || tag === 'th' ? tag : 'p');
  for (const attribute of [...target.attributes]) {
    if (attribute.name === 'data-bid') continue;
    split.setAttribute(attribute.name, attribute.value);
  }
  split.appendChild(tail);
  target.parentNode?.insertBefore(split, target.nextSibling);

  assignBlockIds(container);
  return {
    ok: true,
    message: `Блок ${id} разделён на символе ${offset}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opMergeBlocks(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; объединять его нельзя' };
  }

  const nextId = asString(args.next_id ?? args.nextId ?? args.with_block_id ?? args.withBlockId);
  const next = nextId
    ? byId(container, nextId)
    : (target.nextElementSibling as Element | null);
  if (!next) return { ok: false, message: 'Следующий блок не найден' };
  if (isReadOnlyBlock(next)) {
    return { ok: false, message: 'Следующий объект доступен только для просмотра' };
  }

  const needsSpace =
    (target.textContent ?? '').trim().length > 0 &&
    (next.textContent ?? '').trim().length > 0 &&
    !/\s$/.test(target.textContent ?? '');
  if (needsSpace) target.appendChild(document.createTextNode(' '));
  while (next.firstChild) target.appendChild(next.firstChild);
  next.remove();

  assignBlockIds(container);
  return {
    ok: true,
    message: `Блок ${nextId || 'следующий'} объединён с ${id}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function tableRows(table: Element): Element[] {
  return [...table.querySelectorAll(':scope > tr, :scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr')];
}

function tableTarget(container: HTMLElement, args: Record<string, unknown>): Element | DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (target.tagName.toLowerCase() !== 'table') {
    return { ok: false, message: `Блок ${id || '(без id)'} не является таблицей` };
  }
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Таблица содержит объект только для просмотра; менять её нельзя' };
  }
  return target;
}

function cellAt(table: Element, row: number, col: number): Element | null {
  const rows = tableRows(table);
  return rows[row]?.children[col] ?? null;
}

function appendTableRow(table: Element, row: Element, index: number): void {
  const rows = tableRows(table);
  const reference = rows[index] ?? null;
  const parent = reference?.parentElement ?? table;
  parent.insertBefore(row, reference ?? null);
}

function opTableRead(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const table = byId(container, id);
  if (!table || table.tagName.toLowerCase() !== 'table') {
    return { ok: false, message: `Таблица ${id || '(без id)'} не найдена` };
  }

  const rows = tableRows(table);
  const cells = rows.map((row, rowIndex) =>
    [...row.children].map((cell, colIndex) => ({
      row: rowIndex,
      col: colIndex,
      tag: cell.tagName.toLowerCase(),
      text: (cell.textContent ?? '').trim(),
      colspan: Number(cell.getAttribute('colspan') || 1),
      rowspan: Number(cell.getAttribute('rowspan') || 1),
    }))
  );
  const structure = cells
    .map(
      (row, rowIndex) =>
        `r${rowIndex}: ` +
        row.map((cell) => `c${cell.col}=\"${truncate(cell.text, 60)}\"`).join(' | ')
    )
    .join('\n');

  return {
    ok: true,
    message: `Таблица ${id}: ${rows.length} строк(и), ${rows[0]?.children.length ?? 0} столбц(ов)`,
    changed: false,
    content: table.outerHTML,
    structure,
    meta: { rows: rows.length, cols: rows[0]?.children.length ?? 0, cells },
  };
}

function opTableInsertRow(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const table = tableTarget(container, args);
  if (!(table instanceof Element)) return table;

  const rows = tableRows(table);
  if (rows.length === 0) return { ok: false, message: 'В таблице нет строк' };
  const requested = args.index ?? args.row;
  const index = requested === undefined ? rows.length : clampOffset(requested, 0, rows.length);
  const template = rows[Math.min(index, rows.length - 1)];
  const row = template.cloneNode(true) as Element;
  for (const cell of [...row.children]) {
    cell.innerHTML = '<p><br></p>';
    cell.removeAttribute('colspan');
    cell.removeAttribute('rowspan');
  }

  const values = Array.isArray(args.cells) ? args.cells : [];
  if (values.length > 0) {
    [...row.children].forEach((cell, cellIndex) => {
      const value = values[cellIndex];
      if (value !== undefined && value !== null) cell.innerHTML = `<p>${escapeHtml(String(value))}</p>`;
    });
  }

  appendTableRow(table, row, index);
  assignBlockIds(container);
  return {
    ok: true,
    message: `Добавлена строка ${index}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opTableDeleteRow(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const table = tableTarget(container, args);
  if (!(table instanceof Element)) return table;
  const rows = tableRows(table);
  const index = clampOffset(args.index ?? args.row, 0, Math.max(0, rows.length - 1));
  if (rows.length <= 1) return { ok: false, message: 'Нельзя удалить единственную строку' };
  rows[index]?.remove();
  assignBlockIds(container);
  return {
    ok: true,
    message: `Удалена строка ${index}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opTableInsertColumn(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const table = tableTarget(container, args);
  if (!(table instanceof Element)) return table;
  const rows = tableRows(table);
  if (rows.length === 0) return { ok: false, message: 'В таблице нет строк' };
  const width = rows[0].children.length;
  const requested = args.index ?? args.col;
  const index = requested === undefined ? width : clampOffset(requested, 0, width);

  for (const row of rows) {
    const cell = document.createElement('td');
    cell.innerHTML = '<p><br></p>';
    row.insertBefore(cell, row.children[index] ?? null);
  }
  assignBlockIds(container);
  return {
    ok: true,
    message: `Добавлен столбец ${index}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opTableDeleteColumn(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const table = tableTarget(container, args);
  if (!(table instanceof Element)) return table;
  const rows = tableRows(table);
  const width = rows[0]?.children.length ?? 0;
  if (width <= 1) return { ok: false, message: 'Нельзя удалить единственный столбец' };
  const index = clampOffset(args.index ?? args.col, 0, width - 1);

  for (const row of rows) row.children[index]?.remove();
  assignBlockIds(container);
  return {
    ok: true,
    message: `Удалён столбец ${index}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opTableSetCell(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const table = tableTarget(container, args);
  if (!(table instanceof Element)) return table;
  const rows = tableRows(table);
  const row = clampOffset(args.row, 0, Math.max(0, rows.length - 1));
  const col = clampOffset(args.col, 0, Math.max(0, (rows[row]?.children.length ?? 1) - 1));
  const cell = cellAt(table, row, col);
  if (!cell) return { ok: false, message: `Ячейка r${row}c${col} не найдена` };

  const raw = asString(args.html ?? args.content ?? args.text);
  if (!raw.trim()) return { ok: false, message: 'Пустое содержимое ячейки' };
  const next = sanitize(toBlockHtml(cleanInputHtml(raw)));
  cleanChatNoise(next);
  if (next.children.length === 0) {
    return { ok: false, message: 'В ячейке останется только служебный отчёт ассистента' };
  }
  cell.innerHTML = next.innerHTML;

  assignBlockIds(container);
  return {
    ok: true,
    message: `Ячейка r${row}c${col} обновлена`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opTableMergeCells(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const table = tableTarget(container, args);
  if (!(table instanceof Element)) return table;
  const rows = tableRows(table);
  const row = clampOffset(args.row, 0, Math.max(0, rows.length - 1));
  const col = clampOffset(args.col, 0, Math.max(0, (rows[row]?.children.length ?? 1) - 1));
  const rowspan = Math.max(1, Math.floor(Number(args.rowspan ?? args.rows) || 1));
  const colspan = Math.max(1, Math.floor(Number(args.colspan ?? args.cols) || 1));
  const anchor = cellAt(table, row, col);
  if (!anchor) return { ok: false, message: `Ячейка r${row}c${col} не найдена` };
  if (row + rowspan > rows.length) {
    return { ok: false, message: 'Объединение выходит за пределы таблицы по строкам' };
  }
  for (let r = 0; r < rowspan; r += 1) {
    if (col + colspan > (rows[row + r]?.children.length ?? 0)) {
      return { ok: false, message: 'Объединение выходит за пределы таблицы по столбцам' };
    }
  }

  for (let r = 0; r < rowspan; r += 1) {
    for (let c = 0; c < colspan; c += 1) {
      if (r === 0 && c === 0) continue;
      const cell = rows[row + r]?.children[col + c];
      if (!cell) continue;
      const text = (cell.textContent ?? '').trim();
      if (text) anchor.appendChild(document.createTextNode(` ${text}`));
      cell.remove();
    }
  }

  if (rowspan > 1) anchor.setAttribute('rowspan', String(rowspan));
  if (colspan > 1) anchor.setAttribute('colspan', String(colspan));

  assignBlockIds(container);
  return {
    ok: true,
    message: `Ячейка r${row}c${col} объединена (${rowspan}×${colspan})`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opInsertImage(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const src = asString(args.src ?? args.url ?? args.path);
  if (!/^data:image\//i.test(src)) {
    return {
      ok: false,
      message: 'Нужен data URL изображения. Путь к файлу workspace поддержан через инструмент приложения.',
    };
  }

  const blockId = asString(args.block_id ?? args.blockId ?? args.id);
  const anchor = blockId ? byId(container, blockId) : null;
  if (blockId && !anchor) return { ok: false, message: `Блок ${blockId} не найден` };

  const paragraph = document.createElement('p');
  const image = document.createElement('img');
  image.setAttribute('src', src);
  const alt = asString(args.alt ?? args.description);
  if (alt) image.setAttribute('alt', alt);
  const width = Number(args.width ?? args.width_px ?? args.widthPx);
  const height = Number(args.height ?? args.height_px ?? args.heightPx);
  const styles: string[] = [];
  if (Number.isFinite(width) && width > 0) styles.push(`width:${Math.round(width)}px`);
  if (Number.isFinite(height) && height > 0) styles.push(`height:${Math.round(height)}px`);
  if (styles.length) image.setAttribute('style', styles.join(';'));
  paragraph.appendChild(image);

  const align = normalizeAlign(args.align ?? args.alignment);
  if (align) paragraph.style.textAlign = align;

  if (anchor) anchor.parentNode?.insertBefore(paragraph, anchor.nextSibling);
  else container.appendChild(paragraph);

  assignBlockIds(container);
  return {
    ok: true,
    message: 'Изображение вставлено',
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opUpdateImage(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = id ? byId(container, id) : null;
  if (id && !target) return { ok: false, message: `Блок ${id} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; изображение менять нельзя' };
  }

  const image =
    (target ? target.querySelector('img') : null) ??
    [...container.querySelectorAll('img')].find((node) => node.getAttribute('src'));
  if (!image) return { ok: false, message: 'В блоке нет изображения' };

  const alt = args.alt ?? args.description;
  if (alt !== undefined) image.setAttribute('alt', asString(alt));
  const width = Number(args.width ?? args.width_px ?? args.widthPx);
  const height = Number(args.height ?? args.height_px ?? args.heightPx);
  if (Number.isFinite(width) && width > 0) image.style.width = `${Math.round(width)}px`;
  if (Number.isFinite(height) && height > 0) image.style.height = `${Math.round(height)}px`;

  const align = normalizeAlign(args.align ?? args.alignment);
  if (align) {
    const paragraph = image.closest('p') ?? image.parentElement;
    if (paragraph instanceof HTMLElement) paragraph.style.textAlign = align;
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: 'Изображение обновлено',
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opDeleteImage(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = id ? byId(container, id) : null;
  if (id && !target) return { ok: false, message: `Блок ${id} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; изображение удалить нельзя' };
  }

  const images = target
    ? [...target.querySelectorAll('img')]
    : [...container.querySelectorAll('img')].slice(0, 1);
  if (images.length === 0) return { ok: false, message: 'В блоке нет изображения' };
  for (const image of images) image.remove();

  assignBlockIds(container);
  return {
    ok: true,
    message: `Удалено изображений: ${images.length}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opInsertPageBreak(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const blockId = asString(args.block_id ?? args.blockId ?? args.id);
  const anchor = blockId ? byId(container, blockId) : null;
  if (blockId && !anchor) return { ok: false, message: `Блок ${blockId} не найден` };

  const paragraph = document.createElement('p');
  paragraph.setAttribute('data-page-break', '1');
  const br = document.createElement('br');
  br.setAttribute('data-page-break', '1');
  paragraph.appendChild(br);

  if (anchor) anchor.parentNode?.insertBefore(paragraph, anchor.nextSibling);
  else container.appendChild(paragraph);

  assignBlockIds(container);
  return {
    ok: true,
    message: 'Разрыв страницы вставлен',
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opInsertToc(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const blockId = asString(args.block_id ?? args.blockId ?? args.id);
  const anchor = blockId ? byId(container, blockId) : null;
  if (blockId && !anchor) return { ok: false, message: `Блок ${blockId} не найден` };

  const paragraph = document.createElement('p');
  paragraph.setAttribute('data-field', 'TOC');
  paragraph.textContent = 'Оглавление';
  if (anchor) anchor.parentNode?.insertBefore(paragraph, anchor.nextSibling);
  else container.prepend(paragraph);

  assignBlockIds(container);
  return {
    ok: true,
    message: 'Поле оглавления вставлено; Word обновит его при открытии',
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opInsertBookmark(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const rawName = asString(args.name ?? args.bookmark).trim();
  if (!rawName) return { ok: false, message: 'Не указано имя закладки' };
  const name = rawName.replace(/[^\p{L}\p{N}_.-]+/gu, '_');
  const blockId = asString(args.block_id ?? args.blockId ?? args.id);
  const target = blockId ? byId(container, blockId) : null;
  if (blockId && !target) return { ok: false, message: `Блок ${blockId} не найден` };

  const find = asString(args.find ?? args.text);
  if (target && find) {
    if (!wrapTextInBlock(target, find, 'span', undefined, { 'data-bookmark': name })) {
      return { ok: false, message: `Текст «${find}» не найден в блоке` };
    }
  } else {
    const span = document.createElement('span');
    span.setAttribute('data-bookmark', name);
    if (target) target.prepend(span);
    else container.appendChild(span);
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Закладка «${name}» добавлена`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opInsertField(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const kind = asString(args.kind ?? args.field ?? 'PAGE').toUpperCase();
  if (kind !== 'PAGE' && kind !== 'NUMPAGES') {
    return { ok: false, message: 'Поддерживаются поля PAGE и NUMPAGES' };
  }
  const blockId = asString(args.block_id ?? args.blockId ?? args.id);
  const anchor = blockId ? byId(container, blockId) : null;
  if (blockId && !anchor) return { ok: false, message: `Блок ${blockId} не найден` };

  const paragraph = document.createElement('p');
  const field = document.createElement('span');
  field.setAttribute('data-field', kind);
  field.textContent = '1';
  paragraph.appendChild(field);

  const at = asString(args.at).toLowerCase();
  if (!anchor || at === 'end') container.appendChild(paragraph);
  else if (at === 'before') anchor.parentNode?.insertBefore(paragraph, anchor);
  else anchor.parentNode?.insertBefore(paragraph, anchor.nextSibling);

  assignBlockIds(container);
  return {
    ok: true,
    message: `Поле ${kind} вставлено`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opRefreshToc(container: HTMLElement): DocumentToolResult {
  const toc = container.querySelector('[data-field="TOC"]');
  if (!toc) return { ok: false, message: 'Поле оглавления не найдено' };
  const headings = [...container.querySelectorAll('h1, h2, h3, h4, h5, h6')].filter(
    (heading) => !toc.contains(heading)
  );
  if (headings.length === 0) {
    return { ok: false, message: 'В документе нет заголовков для оглавления' };
  }

  const lines = headings.map((heading) => {
    const level = Math.max(1, Number(heading.tagName.slice(1)) || 1);
    return `${'  '.repeat(level - 1)}${(heading.textContent ?? '').trim()}`;
  });
  toc.textContent = lines.join('\n');
  toc.setAttribute('data-toc-updated', String(Date.now()));

  assignBlockIds(container);
  return {
    ok: true,
    message: `Оглавление обновлено: ${headings.length} заголовков`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opInsertSectionBreak(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const blockId = asString(args.block_id ?? args.blockId ?? args.id);
  const anchor = blockId ? byId(container, blockId) : null;
  if (blockId && !anchor) return { ok: false, message: `Блок ${blockId} не найден` };

  const paragraph = document.createElement('p');
  paragraph.setAttribute('data-section-break', '1');
  paragraph.appendChild(document.createElement('br'));

  if (anchor) anchor.parentNode?.insertBefore(paragraph, anchor.nextSibling);
  else container.appendChild(paragraph);

  assignBlockIds(container);
  return {
    ok: true,
    message: 'Разрыв секции вставлен',
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opDeleteSectionBreak(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = id
    ? byId(container, id)
    : container.querySelector('[data-section-break]');
  if (!target) {
    return { ok: false, message: id ? `Блок ${id} не найден` : 'Разрыв секции не найден' };
  }
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра' };
  }
  if (!target.hasAttribute('data-section-break')) {
    return { ok: false, message: 'Этот блок не является разрывом секции' };
  }

  target.remove();
  assignBlockIds(container);
  return {
    ok: true,
    message: 'Разрыв секции удалён',
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opSetParagraphFormat(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; формат менять нельзя' };
  }

  const targets =
    target.tagName.toLowerCase() === 'ul' || target.tagName.toLowerCase() === 'ol'
      ? [...target.children].filter((item) => item.tagName.toLowerCase() === 'li')
      : [target];

  const lineSpacing = normalizeSpacing(
    args.line_spacing ?? args.lineSpacing ?? args.spacing
  );
  const before = asNumber(args.space_before_pt ?? args.spaceBeforePt ?? args.before_pt);
  const after = asNumber(args.space_after_pt ?? args.spaceAfterPt ?? args.after_pt);
  const keepWithNext = args.keep_with_next ?? args.keepWithNext;
  const keepTogether = args.keep_together ?? args.keepTogether;
  const pageBreakBefore = args.page_break_before ?? args.pageBreakBefore;

  let changed = false;
  for (const element of targets) {
    const beforeStyle = element.getAttribute('style') ?? '';
    const css: StyleMap = {};
    if (lineSpacing !== null) css['line-height'] = String(lineSpacing);
    if (before !== null) css['margin-top'] = before > 0 ? `${before}pt` : '';
    if (after !== null) css['margin-bottom'] = after > 0 ? `${after}pt` : '';
    if (keepWithNext !== undefined) css['break-after'] = asBool(keepWithNext) ? 'avoid' : '';
    if (keepTogether !== undefined) css['break-inside'] = asBool(keepTogether) ? 'avoid' : '';
    if (pageBreakBefore !== undefined) {
      css['break-before'] = asBool(pageBreakBefore) ? 'page' : '';
    }
    mergeStyle(element, css);
    const widowControl = args.widow_control ?? args.widowControl;
    if (widowControl !== undefined) {
      const nextWidow = asBool(widowControl) ? 'on' : 'off';
      if (element.getAttribute('data-widow') !== nextWidow) changed = true;
      element.setAttribute('data-widow', nextWidow);
    }
    if ((element.getAttribute('style') ?? '') !== beforeStyle) changed = true;
  }

  if (!changed) return { ok: true, message: 'Формат абзаца уже соответствует запросу', changed: false };

  assignBlockIds(container);
  return {
    ok: true,
    message: 'Формат абзаца обновлён',
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opSetTabs(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; табуляции менять нельзя' };
  }

  const raw = args.positions ?? args.stops ?? args.tabs;
  const values = Array.isArray(raw)
    ? raw.map((value) => Number(value))
    : String(raw ?? '')
        .split(/[,\s;]+/)
        .map((value) => Number(value));
  const stops = [...new Set(values.filter((value) => Number.isFinite(value) && value > 0 && value <= 2000))]
    .sort((a, b) => a - b)
    .map((value) => Math.round(value));
  if (stops.length === 0) {
    return { ok: false, message: 'Укажите позиции табуляции в пунктах, например 24,48' };
  }

  const targets =
    target.tagName.toLowerCase() === 'ul' || target.tagName.toLowerCase() === 'ol'
      ? [...target.children].filter((item) => item.tagName.toLowerCase() === 'li')
      : [target];
  for (const element of targets) element.setAttribute('data-tabs', stops.join(','));

  assignBlockIds(container);
  return {
    ok: true,
    message: `Табуляции: ${stops.join(', ')} пт`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opDefineStyle(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const name = asString(args.name ?? args.style_name ?? args.styleName).trim();
  if (!name) return { ok: false, message: 'Не указано имя стиля' };

  const sizeRaw = Number(args.size_pt ?? args.sizePt);
  const lineRaw = Number(args.line_height ?? args.lineHeight);
  const beforeRaw = Number(args.space_before_pt ?? args.spaceBeforePt);
  const afterRaw = Number(args.space_after_pt ?? args.spaceAfterPt);
  const alignRaw = asString(args.align).toLowerCase();
  const definition: CustomStyleDefinition = {
    id: asString(args.style_id ?? args.styleId).trim() || styleIdFor(name),
    name,
    basedOn: asString(args.based_on ?? args.basedOn).trim() || 'Normal',
    font: asString(args.font ?? args.font_family ?? args.fontFamily).trim() || undefined,
    sizePt: Number.isFinite(sizeRaw) && sizeRaw > 0 ? sizeRaw : undefined,
    bold: args.bold === undefined ? undefined : asBool(args.bold),
    italic: args.italic === undefined ? undefined : asBool(args.italic),
    color: asString(args.color).trim() || undefined,
    align:
      alignRaw === 'left' || alignRaw === 'center' || alignRaw === 'right' || alignRaw === 'justify'
        ? alignRaw
        : undefined,
    lineHeight: Number.isFinite(lineRaw) && lineRaw > 0 ? lineRaw : undefined,
    spaceBeforePt: Number.isFinite(beforeRaw) && beforeRaw >= 0 ? beforeRaw : undefined,
    spaceAfterPt: Number.isFinite(afterRaw) && afterRaw >= 0 ? afterRaw : undefined,
  };
  if (
    !definition.font &&
    !definition.sizePt &&
    definition.bold === undefined &&
    definition.italic === undefined &&
    !definition.color &&
    !definition.align &&
    !definition.lineHeight &&
    definition.spaceBeforePt === undefined &&
    definition.spaceAfterPt === undefined
  ) {
    return { ok: false, message: 'Укажите хотя бы один параметр стиля' };
  }

  registerCustomStyle(definition);

  const blockId = asString(args.block_id ?? args.blockId ?? args.id);
  const target = blockId ? byId(container, blockId) : null;
  if (blockId && !target) return { ok: false, message: `Блок ${blockId} не найден` };
  if (target) {
    if (isReadOnlyBlock(target)) {
      return { ok: false, message: 'Объект доступен только для просмотра; стиль менять нельзя' };
    }
    const styled = target as HTMLElement;
    styled.setAttribute('data-style', definition.id);
    const css = customStyleCss(definition);
    for (const rule of css.split(';')) {
      const at = rule.indexOf(':');
      if (at < 0) continue;
      styled.style.setProperty(rule.slice(0, at).trim(), rule.slice(at + 1).trim());
    }
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: target
      ? `Стиль «${name}» создан и применён`
      : `Стиль «${name}» создан`,
    changed: Boolean(target),
    structure: describeStructure(container, 6000),
  };
}

function nextAnnotationId(container: HTMLElement, attribute: string): number {
  let max = 0;
  for (const element of container.querySelectorAll(`[${attribute}]`)) {
    const value = Number(element.getAttribute(attribute));
    if (Number.isFinite(value)) max = Math.max(max, value);
  }
  return max + 1;
}

function opAddComment(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const blockId = asString(args.block_id ?? args.blockId ?? args.id);
  const target = blockId ? byId(container, blockId) : null;
  if (blockId && !target) return { ok: false, message: `Блок ${blockId} не найден` };
  if (!target) return { ok: false, message: 'Укажите block_id для комментария' };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра' };
  }

  const text = asString(args.text ?? args.comment).trim();
  if (!text) return { ok: false, message: 'Пустой текст комментария' };
  const author = asString(args.author).trim() || 'AI';
  const id = nextAnnotationId(container, 'data-comment-id');
  const find = asString(args.find);
  const attrs: Record<string, string> = {
    'data-comment-id': String(id),
    'data-comment-text': text,
    'data-comment-author': author,
    title: `${author}: ${text}`,
  };

  if (find) {
    if (!wrapTextInBlock(target, find, 'span', undefined, attrs)) {
      return { ok: false, message: `Текст «${find}» не найден в блоке` };
    }
  } else {
    const span = document.createElement('span');
    for (const [name, value] of Object.entries(attrs)) span.setAttribute(name, value);
    target.prepend(span);
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Комментарий добавлен: ${author}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opAddFootnote(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const blockId = asString(args.block_id ?? args.blockId ?? args.id);
  const target = blockId ? byId(container, blockId) : null;
  if (blockId && !target) return { ok: false, message: `Блок ${blockId} не найден` };
  if (!target) return { ok: false, message: 'Укажите block_id для сноски' };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра' };
  }

  const text = asString(args.text ?? args.note).trim();
  if (!text) return { ok: false, message: 'Пустой текст сноски' };
  const id = nextAnnotationId(container, 'data-footnote-id');
  const find = asString(args.find);
  const attrs: Record<string, string> = {
    'data-footnote-id': String(id),
    'data-footnote-text': text,
    title: text,
  };

  if (find) {
    if (!wrapTextInBlock(target, find, 'span', undefined, attrs)) {
      return { ok: false, message: `Текст «${find}» не найден в блоке` };
    }
  } else {
    const span = document.createElement('span');
    for (const [name, value] of Object.entries(attrs)) span.setAttribute(name, value);
    target.prepend(span);
  }

  const marker = document.createElement('sup');
  marker.setAttribute('data-footnote-ref', '1');
  marker.textContent = String(id);
  const host =
    target.querySelector(`[data-footnote-id="${id}"]`) ?? target;
  host.appendChild(marker);

  assignBlockIds(container);
  return {
    ok: true,
    message: `Сноска ${id} добавлена`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opSetParagraphStyle(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return { ok: false, message: 'Объект доступен только для просмотра; стиль менять нельзя' };
  }

  const style = asString(args.style ?? args.name ?? args.style_name).trim();
  if (!style) return { ok: false, message: 'Не указано имя стиля' };

  const targets =
    target.tagName.toLowerCase() === 'ul' || target.tagName.toLowerCase() === 'ol'
      ? [...target.children].filter((item) => item.tagName.toLowerCase() === 'li')
      : [target];
  const custom =
    getCustomStyle(style) ?? listCustomStyles().find((item) => item.name === style);
  for (const element of targets) {
    element.setAttribute('data-style', custom?.id ?? style);
    if (!custom) continue;
    const css = customStyleCss(custom);
    const styled = element as HTMLElement;
    for (const rule of css.split(';')) {
      const at = rule.indexOf(':');
      if (at < 0) continue;
      styled.style.setProperty(rule.slice(0, at).trim(), rule.slice(at + 1).trim());
    }
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Стиль абзаца: ${style}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opTableUnmergeCells(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const table = tableTarget(container, args);
  if (!(table instanceof Element)) return table;
  const rows = tableRows(table);
  const row = clampOffset(args.row, 0, Math.max(0, rows.length - 1));
  const col = clampOffset(args.col, 0, Math.max(0, (rows[row]?.children.length ?? 1) - 1));
  const anchor = cellAt(table, row, col);
  if (!anchor) return { ok: false, message: `Ячейка r${row}c${col} не найдена` };

  const rowspan = Math.max(1, Number(anchor.getAttribute('rowspan') || 1));
  const colspan = Math.max(1, Number(anchor.getAttribute('colspan') || 1));
  if (rowspan === 1 && colspan === 1) {
    return { ok: false, message: 'Ячейка не объединена' };
  }

  anchor.removeAttribute('rowspan');
  anchor.removeAttribute('colspan');

  for (let r = 0; r < rowspan; r += 1) {
    const line = rows[row + r];
    if (!line) continue;
    for (let c = 0; c < colspan; c += 1) {
      if (r === 0 && c === 0) continue;
      const cell = document.createElement('td');
      cell.innerHTML = '<p><br></p>';
      if (r === 0) {
        line.insertBefore(cell, anchor.nextSibling);
      } else {
        line.insertBefore(cell, line.children[col] ?? null);
      }
    }
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Объединение снято: r${row}c${col}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opTableSetColumnWidth(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const table = tableTarget(container, args);
  if (!(table instanceof Element)) return table;
  const rows = tableRows(table);
  const col = clampOffset(args.col, 0, Math.max(0, (rows[0]?.children.length ?? 1) - 1));
  const widthPtRaw = Number(args.width_pt ?? args.widthPt);
  const widthPxRaw = Number(args.width_px ?? args.widthPx);
  const widthPt = Number.isFinite(widthPtRaw) && widthPtRaw > 0
    ? widthPtRaw
    : Number.isFinite(widthPxRaw) && widthPxRaw > 0
      ? widthPxRaw * 0.75
      : 0;
  if (!widthPt) return { ok: false, message: 'Укажите width_pt или width_px' };

  for (const row of rows) {
    const cell = row.children[col];
    if (cell instanceof HTMLElement) cell.style.width = `${Math.round(widthPt * 10) / 10}pt`;
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Ширина столбца ${col}: ${Math.round(widthPt)} пт`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opTableSetBorders(
  container: HTMLElement,
  args: Record<string, unknown>
): DocumentToolResult {
  const table = tableTarget(container, args);
  if (!(table instanceof Element)) return table;

  const rawStyle = String(args.style ?? 'single').toLowerCase();
  const style =
    rawStyle === 'none'
      ? 'none'
      : ['single', 'double', 'dashed', 'dotted'].includes(rawStyle)
        ? rawStyle
        : 'single';
  const color = String(args.color ?? '#000000').trim() || '#000000';
  const sizeRaw = Number(args.size_pt ?? args.sizePt ?? 1);
  const sizePt = Number.isFinite(sizeRaw) && sizeRaw > 0 ? sizeRaw : 1;

  table.setAttribute(
    'data-table-border',
    style === 'none' ? 'none' : `${sizePt}|${style}|${color}`
  );
  for (const cell of table.querySelectorAll('td, th')) {
    if (!(cell instanceof HTMLElement)) continue;
    if (style === 'none') {
      cell.style.borderStyle = 'none';
      cell.style.borderWidth = '0';
    } else {
      cell.style.border = `${sizePt}pt ${style === 'single' ? 'solid' : style} ${color}`;
    }
  }

  assignBlockIds(container);
  return {
    ok: true,
    message:
      style === 'none'
        ? 'Границы таблицы убраны'
        : `Границы таблицы: ${sizePt} пт, ${style}, ${color}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opDeleteBlock(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };
  if (isReadOnlyBlock(target)) {
    return {
      ok: false,
      message: 'Этот блок содержит объект Word, доступный только для просмотра; удалить его нельзя',
    };
  }

  const parent = target.parentElement;
  if (target.tagName.toLowerCase() === 'li' && parent) {
    target.remove();
    if (parent.children.length === 0) parent.remove();
  } else {
    target.remove();
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Блок ${id} удалён`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function textNodes(root: Node): Text[] {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  let current = walker.nextNode();
  while (current) {
    if ((current.textContent ?? '').length > 0) nodes.push(current as Text);
    current = walker.nextNode();
  }
  return nodes;
}

/** Блоки, внутри которых ищем текст: фраза не должна склеиваться из соседних абзацев. */
const TEXT_BLOCK_SELECTOR = 'p, li, h1, h2, h3, h4, h5, h6, td, th, blockquote, div';

/** Схлопывает пробелы для нечёткого поиска: Word/модель по-разному их ставят. */
function collapseWhitespace(text: string): { text: string; map: number[] } {
  let normalized = '';
  const map: number[] = [];
  let pendingSpace = false;
  let spaceStart = -1;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (/\s/.test(char)) {
      if (!pendingSpace) {
        pendingSpace = true;
        spaceStart = index;
      }
      continue;
    }
    if (pendingSpace && normalized.length > 0) {
      normalized += ' ';
      map.push(spaceStart);
    }
    pendingSpace = false;
    normalized += char;
    map.push(index);
  }

  return { text: normalized, map };
}

/**
 * Заменяет фразу, даже если Word разбил её на несколько runs (`<strong>`,
 * ссылки, разное форматирование). Обычный indexOf по одному текстовому узлу
 * такие фразы не находил, и модель вместо правки строки переписывала весь
 * документ.
 */
function replaceAcrossNodes(
  nodes: Text[],
  find: string,
  replacement: string,
  matchCase: boolean,
  all: boolean
): number {
  let replaced = 0;

  while (true) {
    const combined = nodes.map((node) => node.textContent ?? '').join('');
    if (!combined) break;

    const haystack = matchCase ? combined : combined.toLowerCase();
    const needle = matchCase ? find : find.toLowerCase();
    let at = haystack.indexOf(needle);
    let length = find.length;

    // Модель может прислать ту же фразу с другим числом пробелов.
    if (at === -1) {
      const normalized = collapseWhitespace(combined);
      const normalizedFind = collapseWhitespace(find).text;
      if (normalizedFind) {
        const normalizedHaystack = matchCase ? normalized.text : normalized.text.toLowerCase();
        const normalizedNeedle = matchCase ? normalizedFind : normalizedFind.toLowerCase();
        const normalizedAt = normalizedHaystack.indexOf(normalizedNeedle);
        if (normalizedAt !== -1) {
          at = normalized.map[normalizedAt];
          length =
            normalized.map[normalizedAt + normalizedNeedle.length - 1] + 1 - at;
        }
      }
    }

    if (at === -1) break;

    const starts: number[] = [];
    let cursor = 0;
    for (const node of nodes) {
      starts.push(cursor);
      cursor += (node.textContent ?? '').length;
    }

    let startIndex = 0;
    while (
      startIndex < nodes.length - 1 &&
      starts[startIndex] + (nodes[startIndex].textContent ?? '').length <= at
    ) {
      startIndex += 1;
    }

    let endIndex = startIndex;
    while (
      endIndex < nodes.length &&
      starts[endIndex] + (nodes[endIndex].textContent ?? '').length < at + length
    ) {
      endIndex += 1;
    }
    if (endIndex >= nodes.length) break;

    const startOffset = at - starts[startIndex];
    const endOffset = at + length - starts[endIndex];
    const startNode = nodes[startIndex];
    const endNode = nodes[endIndex];
    const startText = startNode.textContent ?? '';

    if (startNode === endNode) {
      startNode.textContent =
        startText.slice(0, startOffset) + replacement + startText.slice(endOffset);
    } else {
      startNode.textContent = startText.slice(0, startOffset) + replacement;
      for (let index = startIndex + 1; index < endIndex; index += 1) {
        nodes[index].textContent = '';
      }
      endNode.textContent = (endNode.textContent ?? '').slice(endOffset);
    }

    replaced += 1;
    if (!all) break;
  }

  return replaced;
}

function opFindReplace(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const find = asString(args.find ?? args.search ?? args.from).replace(/\u00a0/g, ' ');
  const replace = (asString(args.replace ?? args.replacement ?? args.to) ?? '').replace(
    /\u00a0/g,
    ' '
  );
  if (!find) return { ok: false, message: 'Не указан текст для поиска' };

  const all = asBool(args.all) || asBool(args.replace_all) || asBool(args.replaceAll);
  const matchCase = asBool(args.match_case) || asBool(args.matchCase);

  // Word хранит неразрывные пробелы как \u00a0 — модель присылает обычные.
  const groups = new Map<Element, Text[]>();
  for (const node of textNodes(container)) {
    if (node.parentElement?.closest('[data-docx-readonly]')) continue;
    const normalized = (node.textContent ?? '').replace(/\u00a0/g, ' ');
    if (normalized !== node.textContent) node.textContent = normalized;
    const block = node.parentElement?.closest(TEXT_BLOCK_SELECTOR) ?? container;
    const group = groups.get(block);
    if (group) group.push(node);
    else groups.set(block, [node]);
  }

  let replaced = 0;
  for (const nodes of groups.values()) {
    replaced += replaceAcrossNodes(nodes, find, replace, matchCase, all);
    if (!all && replaced > 0) break;
  }

  if (replaced === 0) {
    return { ok: false, message: `Текст «${find}» не найден` };
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Заменено вхождений: ${replaced}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

/** Цвет только в безопасном для inline-стиля виде. */
function normalizeColor(value: unknown): string {
  const text = asString(value).trim();
  if (/^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i.test(text)) return text;
  if (/^rgba?\(\s*[\d.%\s,]+\)$/i.test(text)) return text;
  if (/^[a-z]{3,20}$/i.test(text)) return text;
  return '';
}

/** Ссылка только с безопасной схемой. */
function isSafeLink(value: string): boolean {
  return /^(https?:\/\/|mailto:|tel:)/i.test(value.trim());
}

function normalizeLink(value: unknown): string {
  const text = asString(value).trim();
  return isSafeLink(text) ? text : '';
}

/** Стиль границы в терминах CSS/Word. */
function normalizeBorderStyle(value: unknown): string {
  const text = asString(value).trim().toLowerCase();
  if (text === 'dashed' || text === 'dotted' || text === 'double' || text === 'solid') {
    return text;
  }
  return '';
}

function normalizeHeading(value: unknown): number | null {
  if (value === undefined || value === null || value === '') return null;
  const text = String(value).trim().toLowerCase();
  if (text === 'p' || text === 'paragraph' || text === 'normal' || text === '0') return 0;
  const match = /([1-6])/.exec(text);
  return match ? Number(match[1]) : null;
}

function normalizeList(value: unknown): 'ul' | 'ol' | 'none' | null {
  if (value === undefined || value === null || value === '') return null;
  const text = String(value).trim().toLowerCase();
  if (['ul', 'bullet', 'bulleted', 'маркированный', 'bullets'].includes(text)) return 'ul';
  if (['ol', 'numbered', 'number', 'нумерованный', 'decimal'].includes(text)) return 'ol';
  if (['none', 'no', 'remove', 'нет', 'paragraph'].includes(text)) return 'none';
  return null;
}

function normalizeAlign(value: unknown): string | null {
  if (value === undefined || value === null || value === '') return null;
  const text = String(value).trim().toLowerCase();
  if (text === 'left' || text === 'start') return 'left';
  if (text === 'center' || text === 'centre' || text === 'middle') return 'center';
  if (text === 'right' || text === 'end') return 'right';
  if (text === 'justify' || text === 'both') return 'justify';
  return null;
}

function unwrapTags(root: Element, names: string[]): boolean {
  let changed = false;
  for (const name of names) {
    for (const element of [...root.querySelectorAll(name)]) {
      unwrap(element);
      changed = true;
    }
  }
  return changed;
}

function wrapWholeBlock(target: Element, tagName: string, css?: StyleMap): boolean {
  const nodes = target.tagName.toLowerCase() === 'ul' || target.tagName.toLowerCase() === 'ol'
    ? [...target.children].filter((item) => item.tagName.toLowerCase() === 'li')
    : [target];

  let changed = false;
  for (const holder of nodes) {
    const existing = [...holder.children];
    if (
      existing.length === 1 &&
      existing[0].tagName.toLowerCase() === tagName &&
      (!css || Object.entries(css).every(([key, val]) => parseStyle(existing[0].getAttribute('style') ?? '')[key] === val))
    ) {
      continue;
    }
    const wrapper = document.createElement(tagName);
    if (css) mergeStyle(wrapper, css);
    while (holder.firstChild) wrapper.appendChild(holder.firstChild);
    holder.appendChild(wrapper);
    changed = true;
  }
  return changed;
}

/** Оборачивает первое вхождение текста в блоке — «форматирование диапазона». */
function wrapTextInBlock(
  root: Element,
  find: string,
  tagName: string,
  css?: StyleMap,
  attrs?: Record<string, string>
): boolean {
  if (!find) return false;
  const nodes = textNodes(root);
  if (nodes.length === 0) return false;

  let full = '';
  for (const node of nodes) full += node.textContent ?? '';
  const at = full.toLowerCase().indexOf(find.toLowerCase());
  if (at < 0) return false;

  const range = rangeFromOffsets(root, at, at + find.length);
  if (!range) return false;
  const startBlock = (range.startContainer.parentElement ?? root).closest(TEXT_BLOCK_SELECTOR);
  const endBlock = (range.endContainer.parentElement ?? root).closest(TEXT_BLOCK_SELECTOR);
  if (startBlock && endBlock && startBlock !== endBlock) return false;

  const wrapper = document.createElement(tagName);
  if (css) mergeStyle(wrapper, css);
  if (attrs) {
    for (const [name, value] of Object.entries(attrs)) wrapper.setAttribute(name, value);
  }
  wrapper.appendChild(range.extractContents());
  range.insertNode(wrapper);
  return true;
}

function opFormat(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const find = asString(args.find ?? args.text ?? args.search);
  const target = id ? byId(container, id) : null;
  if (id && !target) return { ok: false, message: `Блок ${id} не найден` };
  if (isReadOnlyBlock(target)) {
    return {
      ok: false,
      message: 'Этот блок содержит объект Word, доступный только для просмотра; форматировать его нельзя',
    };
  }
  if (!target && !find) return { ok: false, message: 'Укажите block_id или текст для форматирования (find)' };

  const heading = normalizeHeading(args.heading ?? args.level);
  const list = normalizeList(args.list ?? args.list_type ?? args.listType);
  const align = normalizeAlign(args.align ?? args.alignment);
  const font = asString(args.font ?? args.font_family ?? args.fontFamily);
  const size = Number(args.size_pt ?? args.sizePt ?? args.font_size ?? args.fontSize);
  const indent = Number(args.indent_pt ?? args.indentPt ?? args.indent);
  const firstLine = Number(args.first_line_indent_pt ?? args.firstLineIndentPt ?? args.text_indent);

  const css: StyleMap = {};
  if (align) css['text-align'] = align;
  if (font) css['font-family'] = `"${font.replace(/["']/g, '')}"`;
  if (Number.isFinite(size) && size > 0) css['font-size'] = `${size}pt`;
  if (Number.isFinite(indent)) css['margin-left'] = indent > 0 ? `${indent}pt` : '';
  if (Number.isFinite(firstLine)) css['text-indent'] = firstLine !== 0 ? `${firstLine}pt` : '';
  const color = normalizeColor(args.color ?? args.text_color ?? args.textColor);
  const highlight = normalizeColor(
    args.highlight ?? args.background ?? args.background_color ?? args.bg
  );
  const link = normalizeLink(args.link ?? args.link_url ?? args.href ?? args.url);
  if (color) css['color'] = color;
  if (highlight) css['background-color'] = highlight;
  const fill = normalizeColor(args.fill ?? args.block_fill ?? args.blockFill);
  if (fill) css['background-color'] = fill;

  const borderRequested =
    args.border !== undefined ||
    args.border_color !== undefined ||
    args.border_width_pt !== undefined ||
    args.border_style !== undefined;
  if (borderRequested) {
    if (args.border === false) {
      css['border'] = '';
    } else {
      const borderColor = normalizeColor(args.border_color ?? args.borderColor) || '#000000';
      const borderWidth = Math.max(
        0.25,
        Math.min(12, Number(args.border_width_pt ?? args.borderWidthPt ?? 1) || 1)
      );
      const borderStyle =
        normalizeBorderStyle(args.border_style ?? args.borderStyle) || 'solid';
      css['border'] = `${borderWidth}pt ${borderStyle} ${borderColor}`;
    }
  }

  // null — параметр не передан, true/false — явное указание.
  const bold = args.bold === undefined ? null : asBool(args.bold);
  const italic = args.italic === undefined ? null : asBool(args.italic);
  const underline = args.underline === undefined ? null : asBool(args.underline);

  let changed = false;
  const applied: string[] = [];

  // Глобальный диапазон: block_id не указан, есть find — ищем по всему документу.
  if (!target && find) {
    const inlineCss: StyleMap = {};
    if (font) inlineCss['font-family'] = `"${font.replace(/["']/g, '')}"`;
    if (Number.isFinite(size) && size > 0) inlineCss['font-size'] = `${size}pt`;
    if (color) inlineCss.color = color;
    if (highlight) inlineCss['background-color'] = highlight;
    let wrapped = false;
    if (bold === true) wrapped = wrapTextInBlock(container, find, 'strong') || wrapped;
    if (italic === true) wrapped = wrapTextInBlock(container, find, 'em') || wrapped;
    if (underline === true) wrapped = wrapTextInBlock(container, find, 'u') || wrapped;
    if (link) {
      wrapped =
        wrapTextInBlock(container, find, 'a', undefined, { href: link }) || wrapped;
    }
    if (font || color || highlight || (Number.isFinite(size) && size > 0)) {
      wrapped = wrapTextInBlock(container, find, 'span', inlineCss) || wrapped;
    }
    if (!wrapped) return { ok: false, message: `Текст «${find}» не найден` };
    assignBlockIds(container);
    return {
      ok: true,
      message: 'Фрагмент отформатирован',
      changed: true,
      structure: describeStructure(container, 6000),
    };
  }

  if (!target) return { ok: false, message: 'Блок не найден' };

  const isList = target.tagName.toLowerCase() === 'ul' || target.tagName.toLowerCase() === 'ol';
  const isItem = target.tagName.toLowerCase() === 'li';
  const blockTarget = isItem ? (target.parentElement ?? target) : target;

  if (heading !== null) {
    if (isList || isItem) {
      return { ok: false, message: 'Заголовок нельзя применить к контейнеру или элементу списка' };
    }
    const tag = heading === 0 ? 'p' : `h${Math.min(6, heading)}`;
    if (target.tagName.toLowerCase() !== tag) {
      const replacement = document.createElement(tag);
      for (const attribute of [...target.attributes]) replacement.setAttribute(attribute.name, attribute.value);
      while (target.firstChild) replacement.appendChild(target.firstChild);
      target.replaceWith(replacement);
      changed = true;
      applied.push(tag.toUpperCase());
    }
  }

  if (list) {
    const current = isList ? target.tagName.toLowerCase() : isItem ? blockTarget.tagName.toLowerCase() : '';
    if (list === 'none') {
      if (isItem || isList) {
        // Превращаем весь список в абзацы: p внутри ul недопустим и теряется
        // при сохранении в .docx.
        const listElement = isItem ? blockTarget : target;
        const listId = listElement.getAttribute('data-bid');
        const fragment = document.createDocumentFragment();
        let firstParagraph: Element | null = null;
        for (const item of [...listElement.children]) {
          if (item.tagName.toLowerCase() !== 'li') continue;
          const paragraph = document.createElement('p');
          for (const attribute of [...item.attributes]) {
            paragraph.setAttribute(attribute.name, attribute.value);
          }
          while (item.firstChild) paragraph.appendChild(item.firstChild);
          fragment.appendChild(paragraph);
          if (!firstParagraph) firstParagraph = paragraph;
        }
        listElement.replaceWith(fragment);
        // Сохраняем адрес списка на первом абзаце: следующий format может
        // ссылаться на тот же block_id.
        if (firstParagraph && listId) firstParagraph.setAttribute('data-bid', listId);
        changed = true;
        applied.push('обычный текст');
      }
    } else if (current !== list) {
      if (isList) {
        const replacement = document.createElement(list);
        for (const attribute of [...target.attributes]) replacement.setAttribute(attribute.name, attribute.value);
        while (target.firstChild) replacement.appendChild(target.firstChild);
        target.replaceWith(replacement);
      } else if (isItem && blockTarget) {
        const replacement = document.createElement(list);
        for (const attribute of [...blockTarget.attributes]) replacement.setAttribute(attribute.name, attribute.value);
        const item = document.createElement('li');
        while (target.firstChild) item.appendChild(target.firstChild);
        replacement.appendChild(item);
        blockTarget.replaceWith(replacement);
      } else {
        const replacement = document.createElement(list);
        for (const attribute of [...target.attributes]) replacement.setAttribute(attribute.name, attribute.value);
        const item = document.createElement('li');
        while (target.firstChild) item.appendChild(target.firstChild);
        replacement.appendChild(item);
        target.replaceWith(replacement);
      }
      changed = true;
      applied.push(list === 'ul' ? 'маркированный список' : 'нумерованный список');
    }
  }

  // После смены типа элемента берём актуальный узел.
  let current: Element | null = id ? byId(container, id) : null;
  if (!current) {
    // При смене тега data-bid сохраняется, но узел мог быть заменён.
    current = byId(container, id) ?? container;
  }

  if (Object.keys(css).length > 0) {
    const styleTargets =
      current.tagName.toLowerCase() === 'ul' || current.tagName.toLowerCase() === 'ol'
        ? [...current.children].filter((item) => item.tagName.toLowerCase() === 'li')
        : [current];
    let styleChanged = false;
    for (const styleTarget of styleTargets) {
      const before = styleTarget.getAttribute('style') ?? '';
      mergeStyle(styleTarget, css);
      if ((styleTarget.getAttribute('style') ?? '') !== before) styleChanged = true;
    }
    if (styleChanged) {
      changed = true;
      applied.push('оформление');
    }
  }

  if (find) {
    let wrapped = false;
    if (bold === true) wrapped = wrapTextInBlock(current, find, 'strong') || wrapped;
    if (italic === true) wrapped = wrapTextInBlock(current, find, 'em') || wrapped;
    if (underline === true) wrapped = wrapTextInBlock(current, find, 'u') || wrapped;
    if (link) {
      wrapped = wrapTextInBlock(current, find, 'a', undefined, { href: link }) || wrapped;
    }
    if (color || highlight) {
      const inlineCss: StyleMap = {};
      if (color) inlineCss.color = color;
      if (highlight) inlineCss['background-color'] = highlight;
      wrapped = wrapTextInBlock(current, find, 'span', inlineCss) || wrapped;
    }
    if (wrapped) {
      changed = true;
      applied.push('выделение фрагмента');
    }
  } else {
    if (bold === false) changed = unwrapTags(current, ['strong', 'b']) || changed;
    else if (bold === true) changed = wrapWholeBlock(current, 'strong') || changed;
    if (bold === true) applied.push('жирный');
    if (bold === false) applied.push('без жирного');

    if (italic === false) changed = unwrapTags(current, ['em', 'i']) || changed;
    else if (italic === true) changed = wrapWholeBlock(current, 'em') || changed;
    if (italic === true) applied.push('курсив');

    if (underline === false) changed = unwrapTags(current, ['u']) || changed;
    else if (underline === true) changed = wrapWholeBlock(current, 'u') || changed;
    if (underline === true) applied.push('подчёркивание');
  }

  if (!changed) {
    return { ok: true, message: 'Оформление уже соответствует запросу', changed: false, structure: describeStructure(container, 6000) };
  }

  assignBlockIds(container);
  return {
    ok: true,
    message: `Блок ${id} отформатирован${applied.length ? `: ${[...new Set(applied)].join(', ')}` : ''}`,
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

/* ------------------------------------------------------- страница документа */

/** Поля делового документа в России: 3 / 1,5 / 2 / 2 см, Times New Roman 14. */
const BUSINESS_PAGE: PageSetup = {
  marginTopMm: 20,
  marginRightMm: 15,
  marginBottomMm: 20,
  marginLeftMm: 30,
  fontFamily: 'Times New Roman',
  fontSizePt: 14,
  lineSpacing: 1,
};

const BUSINESS_FIRST_LINE_PT = 35.4; // 1,25 см

function asNumber(value: unknown): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value === 'string') {
    const parsed = Number(value.replace(',', '.').replace(/[^\d.+-]/g, ''));
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function capitalize(value: string): string {
  return value.charAt(0).toUpperCase() + value.slice(1);
}

/** Длина из аргументов инструмента: сантиметры по умолчанию, ключи *_mm — миллиметры. */
function readMarginFrom(source: Record<string, unknown>, side: string): number | null {
  const candidates: Array<[string, number]> = [
    [`margin_${side}_cm`, 10],
    [`margin${capitalize(side)}Cm`, 10],
    [`${side}_cm`, 10],
    [`margin_${side}`, 10],
    [`margin${capitalize(side)}`, 10],
    [side, 10],
    [`margin_${side}_mm`, 1],
    [`${side}_mm`, 1],
  ];
  for (const [key, factor] of candidates) {
    if (source[key] === undefined) continue;
    const value = asNumber(source[key]);
    if (value !== null && value >= 0) return Math.round(value * factor * 10) / 10;
  }
  return null;
}

function readMargin(args: Record<string, unknown>, side: string): number | null {
  const nested = args.margins ?? args.page_margins ?? args.pageMargins;
  if (nested && typeof nested === 'object' && !Array.isArray(nested)) {
    const value = readMarginFrom(nested as Record<string, unknown>, side);
    if (value !== null) return value;
  }
  return readMarginFrom(args, side);
}

function normalizeSpacing(value: unknown): number | null {
  if (value === undefined || value === null || value === '') return null;
  const text = String(value).trim().toLowerCase();
  if (text === 'single' || text === 'одинарный') return 1;
  if (text === 'one_and_half' || text === '1.5' || text === 'полуторный') return 1.5;
  if (text === 'double' || text === 'двойной') return 2;
  const number = asNumber(value);
  if (number === null || number <= 0) return null;
  return Math.min(3, Math.max(1, number));
}

function isTemplateName(value: unknown): boolean {
  const text = String(value ?? '').toLowerCase();
  return /russian|business|official|gost|гост|делов|офицер|заявлен|письм/.test(text);
}

/**
 * Оформление страницы и базового текста: поля, размер бумаги, шрифт, кегль,
 * межстрочный интервал. Без этого инструмента модель могла только советовать
 * «настройте поля в Word», а не оформлять документ.
 */
function opPageSetup(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const page: PageSetup = {};
  const appliesTemplate = isTemplateName(args.template);

  if (appliesTemplate) Object.assign(page, BUSINESS_PAGE);

  const top = readMargin(args, 'top');
  const right = readMargin(args, 'right');
  const bottom = readMargin(args, 'bottom');
  const left = readMargin(args, 'left');
  if (top !== null) page.marginTopMm = top;
  if (right !== null) page.marginRightMm = right;
  if (bottom !== null) page.marginBottomMm = bottom;
  if (left !== null) page.marginLeftMm = left;

  const font = asString(args.font ?? args.font_family ?? args.fontFamily);
  if (font) page.fontFamily = font.replace(/["']/g, '').trim();

  const size = asNumber(args.size_pt ?? args.sizePt ?? args.font_size ?? args.fontSize);
  if (size !== null && size >= 6 && size <= 72) page.fontSizePt = size;

  const spacing = normalizeSpacing(args.line_spacing ?? args.lineSpacing ?? args.spacing);
  if (spacing !== null) page.lineSpacing = spacing;

  const align = normalizeAlign(args.align ?? args.alignment) ?? (appliesTemplate ? 'justify' : null);
  const rawFirstLineCm = asNumber(args.first_line_indent_cm ?? args.firstLineIndentCm);
  const rawFirstLinePt = asNumber(args.first_line_indent_pt ?? args.firstLineIndentPt);
  const indent =
    rawFirstLineCm !== null
      ? rawFirstLineCm * 28.3465
      : rawFirstLinePt !== null
        ? rawFirstLinePt
        : appliesTemplate
          ? BUSINESS_FIRST_LINE_PT
          : null;

  if (Object.keys(page).length === 0 && !align && indent === null) {
    return { ok: false, message: 'Не переданы поля, шрифт или размер страницы' };
  }

  const fontCss: StyleMap = {};
  if (page.fontFamily) fontCss['font-family'] = `"${page.fontFamily}"`;

  // Кегль и интервал — только обычному тексту: заголовки и цитаты задают
  // собственный размер, иначе документ теряет иерархию разделов.
  const textCss: StyleMap = {};
  if (page.fontSizePt) textCss['font-size'] = `${page.fontSizePt}pt`;
  if (page.lineSpacing) textCss['line-height'] = String(page.lineSpacing);

  const paragraphCss: StyleMap = {};
  if (align) paragraphCss['text-align'] = align;
  if (indent !== null) paragraphCss['text-indent'] = indent > 0 ? `${Math.round(indent)}pt` : '';

  let changed = false;

  const applyStyle = (target: Element, css: StyleMap) => {
    if (Object.keys(css).length === 0) return;
    const before = target.getAttribute('style') ?? '';
    mergeStyle(target, css);
    if ((target.getAttribute('style') ?? '') !== before) changed = true;
  };

  const bodySelector = 'p, div, li, blockquote, td, th';
  for (const element of [...container.querySelectorAll('p, div, h1, h2, h3, h4, h5, h6, li, blockquote, td, th')]) {
    applyStyle(element, fontCss);
    if (element.matches(bodySelector)) applyStyle(element, textCss);
    // Выравнивание и красная строка — только у обычных абзацев: заголовки,
    // цитаты и пункты списков имеют собственное оформление.
    if (element.tagName.toLowerCase() === 'p' && element.parentElement?.tagName.toLowerCase() !== 'td') {
      // Явное центрирование/выравнивание вправо — авторское оформление
      // (шапка, подпись); шаблон страницы его не перебивает.
      const current = (element as HTMLElement).style.textAlign;
      const authored = current === 'center' || current === 'right';
      if (authored) {
        const authoredCss = { ...paragraphCss };
        delete authoredCss['text-align'];
        delete authoredCss['text-indent'];
        applyStyle(element, authoredCss);
      } else {
        applyStyle(element, paragraphCss);
      }
    }
  }

  if (!changed && Object.keys(page).length === 0) {
    return { ok: true, message: 'Оформление уже соответствует запросу', changed: false };
  }

  const parts: string[] = [];
  if (page.marginLeftMm || page.marginRightMm || page.marginTopMm || page.marginBottomMm) {
    parts.push(
      `поля ${page.marginLeftMm ?? 20}/${page.marginRightMm ?? 15}/${page.marginTopMm ?? 20}/${page.marginBottomMm ?? 20} мм`
    );
  }
  if (page.fontFamily) parts.push(page.fontFamily);
  if (page.fontSizePt) parts.push(`${page.fontSizePt} пт`);
  if (page.lineSpacing) parts.push(`интервал ${page.lineSpacing}`);

  assignBlockIds(container);
  return {
    ok: true,
    message: `Оформление страницы обновлено${parts.length ? `: ${parts.join(', ')}` : ''}`,
    changed: true,
    page,
    html: container.innerHTML,
    structure: describeStructure(container, 6000),
  };
}

function asString(value: unknown): string {
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  return '';
}

function asBool(value: unknown): boolean {
  return value === true || value === 1 || value === 'true' || value === '1';
}

/**
 * Выполняет инструмент над HTML документа.
 * Возвращает новый HTML и краткую структуру для следующего шага модели.
 */
export function applyDocumentTool(
  html: string,
  tool: string,
  args: Record<string, unknown> = {}
): DocumentToolResult {
  const container = sanitize(html);
  assignBlockIds(container);

  let result: DocumentToolResult;
  switch (tool) {
    case 'get_document':
      return opGetDocument(container);
    case 'get_outline':
      return opGetOutline(container, args);
    case 'get_block':
      return opGetBlock(container, args);
    case 'get_range':
      return opGetRange(container, args);
    case 'search_document':
      return opSearchDocument(container, args);
    case 'replace_document':
      result = opReplaceDocument(container, args);
      break;
    case 'replace_block':
      result = opReplaceBlock(container, args);
      break;
    case 'insert_after':
      result = opInsertAfter(container, args);
      break;
    case 'insert_before':
      result = opInsertBefore(container, args);
      break;
    case 'append_to_block':
      result = opAppendToBlock(container, args);
      break;
    case 'move_block':
      result = opMoveBlock(container, args);
      break;
    case 'duplicate_block':
      result = opDuplicateBlock(container, args);
      break;
    case 'split_block':
      result = opSplitBlock(container, args);
      break;
    case 'merge_blocks':
      result = opMergeBlocks(container, args);
      break;
    case 'table_read':
      return opTableRead(container, args);
    case 'table_insert_row':
      result = opTableInsertRow(container, args);
      break;
    case 'table_delete_row':
      result = opTableDeleteRow(container, args);
      break;
    case 'table_insert_column':
      result = opTableInsertColumn(container, args);
      break;
    case 'table_delete_column':
      result = opTableDeleteColumn(container, args);
      break;
    case 'table_set_cell':
      result = opTableSetCell(container, args);
      break;
    case 'table_merge_cells':
      result = opTableMergeCells(container, args);
      break;
    case 'table_unmerge_cells':
      result = opTableUnmergeCells(container, args);
      break;
    case 'table_set_column_width':
      result = opTableSetColumnWidth(container, args);
      break;
    case 'table_set_borders':
      result = opTableSetBorders(container, args);
      break;
    case 'insert_image':
      result = opInsertImage(container, args);
      break;
    case 'update_image':
      result = opUpdateImage(container, args);
      break;
    case 'delete_image':
      result = opDeleteImage(container, args);
      break;
    case 'insert_page_break':
      result = opInsertPageBreak(container, args);
      break;
    case 'set_paragraph_format':
      result = opSetParagraphFormat(container, args);
      break;
    case 'set_tabs':
      result = opSetTabs(container, args);
      break;
    case 'set_paragraph_style':
      result = opSetParagraphStyle(container, args);
      break;
    case 'define_style':
      result = opDefineStyle(container, args);
      break;
    case 'insert_section_break':
      result = opInsertSectionBreak(container, args);
      break;
    case 'delete_section_break':
      result = opDeleteSectionBreak(container, args);
      break;
    case 'insert_toc':
      result = opInsertToc(container, args);
      break;
    case 'refresh_toc':
      result = opRefreshToc(container);
      break;
    case 'insert_field':
      result = opInsertField(container, args);
      break;
    case 'insert_bookmark':
      result = opInsertBookmark(container, args);
      break;
    case 'add_comment':
      result = opAddComment(container, args);
      break;
    case 'add_footnote':
      result = opAddFootnote(container, args);
      break;
    case 'delete_block':
      result = opDeleteBlock(container, args);
      break;
    case 'find_replace':
      result = opFindReplace(container, args);
      break;
    case 'replace_range':
      result = opReplaceRange(container, args);
      break;
    case 'delete_range':
      result = opDeleteRange(container, args);
      break;
    case 'format':
      result = opFormat(container, args);
      break;
    case 'page_setup':
      result = opPageSetup(container, args);
      break;
    default:
      return { ok: false, message: `Неизвестный инструмент документа: ${tool}` };
  }

  if (result.ok && result.changed) {
    return { ...result, html: container.innerHTML };
  }
  return result;
}
