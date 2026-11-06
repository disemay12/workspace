/**
 * Программная модель HTML-документа Word.
 *
 * Полотно Word — contenteditable с обычным HTML, без собственной модели блоков.
 * Здесь HTML разбирается в блоки с детерминированными идентификаторами
 * (b1, b2, …), над ними выполняются операции ассистента, и HTML собирается
 * обратно. Идентификаторы не пишутся в .docx: при сохранении атрибуты
 * игнорируются, при открытии файла выдаются заново в том же порядке.
 */
import { markdownToDocumentHtml } from './aiDocument';

export const DOCUMENT_TOOL_NAMES = [
  'get_document',
  'replace_document',
  'replace_block',
  'insert_after',
  'delete_block',
  'find_replace',
  'format',
] as const;

export type DocumentToolName = (typeof DOCUMENT_TOOL_NAMES)[number];

/** Инструменты, меняющие документ. Без прав записи они не передаются модели. */
export const DOCUMENT_WRITE_TOOLS = new Set<string>([
  'replace_document',
  'replace_block',
  'insert_after',
  'delete_block',
  'find_replace',
  'format',
]);

/** Типы вкладок, для которых доступны инструменты документа. */
export const DOCUMENT_TAB_TYPES = new Set<string>(['word', 'markdown', 'richtext', 'text']);

/** Полноценные блочные операции поддерживает только HTML-полотно Word. */
export const BLOCK_TOOL_TAB_TYPES = new Set<string>(['word']);

export interface DocumentToolResult {
  ok: boolean;
  message: string;
  html?: string;
  changed?: boolean;
  structure?: string;
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
      if (name === 'data-bid' || name === 'colspan' || name === 'rowspan') continue;
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
  return probe.innerHTML;
}

function toBlockHtml(raw: string): string {
  const value = (raw || '').trim();
  if (!value) return '';
  if (/<[a-z][\s\S]*>/i.test(value)) return value;
  // Слабая модель иногда присылает Markdown вместо HTML — разбираем его,
  // чтобы заголовки и списки не попали в документ литералами «#» и «-».
  const converted = markdownToDocumentHtml(value);
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

function opGetDocument(container: HTMLElement): DocumentToolResult {
  assignBlockIds(container);
  return {
    ok: true,
    message: 'Структура документа получена',
    changed: false,
    structure: describeStructure(container, 20000),
  };
}

function opReplaceDocument(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const raw = asString(args.html ?? args.content ?? args.text);
  if (!raw.trim()) return { ok: false, message: 'Пустой HTML документа' };

  const next = sanitize(toBlockHtml(cleanInputHtml(raw)));
  if (next.children.length === 0) return { ok: false, message: 'После очистки HTML документ пуст' };

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

  container.innerHTML = next.innerHTML;
  assignBlockIds(container);
  return {
    ok: true,
    message: restored > 0 ? `Документ заменён, изображений сохранено: ${restored}` : 'Документ заменён',
    changed: true,
    structure: describeStructure(container, 6000),
  };
}

function opReplaceBlock(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };

  const raw = asString(args.html ?? args.content ?? args.text);
  if (!raw.trim()) return { ok: false, message: 'Пустой HTML блока' };

  if (target.tagName.toLowerCase() === 'li') {
    target.innerHTML = inlineOnlyHtml(cleanInputHtml(raw));
  } else {
    const next = sanitize(toBlockHtml(cleanInputHtml(raw)));
    if (next.children.length === 0) {
      target.remove();
    } else {
      for (const node of [...next.childNodes]) target.parentNode?.insertBefore(node, target);
      target.remove();
    }
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
  if (next.children.length === 0) return { ok: false, message: 'Вставка пуста' };

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

function opDeleteBlock(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const target = byId(container, id);
  if (!target) return { ok: false, message: `Блок ${id || '(без id)'} не найден` };

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

function opFindReplace(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const find = asString(args.find ?? args.search ?? args.from);
  const replace = asString(args.replace ?? args.replacement ?? args.to) ?? '';
  if (!find) return { ok: false, message: 'Не указан текст для поиска' };

  const all = asBool(args.all) || asBool(args.replace_all) || asBool(args.replaceAll);
  const matchCase = asBool(args.match_case) || asBool(args.matchCase);
  const needle = matchCase ? find : find.toLowerCase();
  let replaced = 0;

  for (const node of textNodes(container)) {
    if (!all && replaced > 0) break;

    let text = node.textContent ?? '';
    let searchFrom = 0;

    while (true) {
      const haystack = matchCase ? text : text.toLowerCase();
      const at = haystack.indexOf(needle, searchFrom);
      if (at === -1) break;

      text = text.slice(0, at) + replace + text.slice(at + find.length);
      replaced += 1;
      searchFrom = at + replace.length;
      if (!all) break;
    }

    node.textContent = text;
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
function wrapTextInBlock(root: Element, find: string, tagName: string, css?: StyleMap): boolean {
  if (!find) return false;
  const needle = find.toLowerCase();

  for (const node of textNodes(root)) {
    const text = node.textContent ?? '';
    const at = text.toLowerCase().indexOf(needle);
    if (at === -1) continue;

    const before = text.slice(0, at);
    const match = text.slice(at, at + find.length);
    const after = text.slice(at + find.length);
    const wrapper = document.createElement(tagName);
    if (css) mergeStyle(wrapper, css);
    wrapper.textContent = match;

    const fragment = document.createDocumentFragment();
    if (before) fragment.appendChild(document.createTextNode(before));
    fragment.appendChild(wrapper);
    if (after) fragment.appendChild(document.createTextNode(after));
    node.replaceWith(fragment);
    return true;
  }
  return false;
}

function opFormat(container: HTMLElement, args: Record<string, unknown>): DocumentToolResult {
  const id = asString(args.block_id ?? args.blockId ?? args.id);
  const find = asString(args.find ?? args.text ?? args.search);
  const target = id ? byId(container, id) : null;
  if (id && !target) return { ok: false, message: `Блок ${id} не найден` };
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
    let wrapped = false;
    if (bold === true) wrapped = wrapTextInBlock(container, find, 'strong') || wrapped;
    if (italic === true) wrapped = wrapTextInBlock(container, find, 'em') || wrapped;
    if (underline === true) wrapped = wrapTextInBlock(container, find, 'u') || wrapped;
    if (font || (Number.isFinite(size) && size > 0)) {
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
    case 'replace_document':
      result = opReplaceDocument(container, args);
      break;
    case 'replace_block':
      result = opReplaceBlock(container, args);
      break;
    case 'insert_after':
      result = opInsertAfter(container, args);
      break;
    case 'delete_block':
      result = opDeleteBlock(container, args);
      break;
    case 'find_replace':
      result = opFindReplace(container, args);
      break;
    case 'format':
      result = opFormat(container, args);
      break;
    default:
      return { ok: false, message: `Неизвестный инструмент документа: ${tool}` };
  }

  if (result.ok && result.changed) {
    return { ...result, html: container.innerHTML };
  }
  return result;
}
