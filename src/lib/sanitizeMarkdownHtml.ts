/**
 * Экранирование HTML, попадающего в разметку из пользовательских данных.
 *
 * Модуль без зависимостей: Markdown-движок тянет KaTeX/Mermaid/CSS, а эта
 * логика нужна и проверочному набору (electron/checks), который запускается
 * вне браузерного окружения.
 */

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Безобидная разметка, которую пропускаем из исходного Markdown: только
 * сами теги, без атрибутов. Всё остальное (script/style/iframe/img/svg, a с
 * javascript:, любые обработчики) уходит в текст: приложение исполняет
 * скрипты в окне с мостом electronAPI, поэтому сырой HTML из документа или
 * ответа модели равен исполнению кода.
 */
const RAW_HTML_TAGS = new Set([
  'b', 'i', 'u', 's', 'em', 'strong', 'del', 'ins', 'mark', 'small', 'sub', 'sup',
  'kbd', 'code', 'br', 'hr', 'p', 'div', 'span', 'details', 'summary',
]);

/**
 * Тег распознаётся по имени сразу после `<`: дальше идут атрибуты, `/` или `>`.
 * Разбирать HTML регуляркой целиком нельзя: `<b/onmouseover=alert(1)>` не
 * совпадал с шаблоном «тег + пробелы + атрибуты», проходил фильтр как есть и
 * исполнялся браузером. Поэтому у разрешённого тега атрибуты срезаются всегда,
 * а всё, что тегом не является, остаётся текстом.
 */
const TAG_RE = /^<(\/?)([a-zA-Z][a-zA-Z0-9-]*)([\s/>]|$)/;

/** Пропускает разрешённые теги без атрибутов, всё остальное делает текстом. */
export function escapeRawHtml(raw: string): string {
  const source = String(raw ?? '').replace(/<!--[\s\S]*?-->/g, '');
  let out = '';
  let index = 0;
  while (index < source.length) {
    const open = source.indexOf('<', index);
    if (open === -1) {
      out += source.slice(index);
      break;
    }
    out += source.slice(index, open);
    const close = source.indexOf('>', open);
    const candidate = close === -1 ? source.slice(open) : source.slice(open, close + 1);
    const match = TAG_RE.exec(candidate);
    if (!match) {
      // `<` + не буква — по спецификации HTML это текст, а не тег.
      out += candidate;
    } else {
      const tag = match[2].toLowerCase();
      if (!RAW_HTML_TAGS.has(tag)) out += escapeHtml(candidate);
      else out += match[1] ? `</${tag}>` : `<${tag}>`;
    }
    if (close === -1) break;
    index = close + 1;
  }
  return out;
}

/** Схемы, по которым разрешено вести переход и подгружать картинку. */
const SAFE_LINK_SCHEMES = new Set(['http:', 'https:', 'mailto:']);
const SAFE_IMAGE_SCHEMES = new Set(['http:', 'https:']);

/**
 * Ссылка из документа может открываться в окне с мостом electronAPI:
 * `javascript:` — это исполнение кода, `file:`/`ms-settings:` — обработчики
 * Windows. Пропускаем только явные схемы и относительные адреса.
 */
function isSafeUrl(value: string, schemes: Set<string>, allowDataImage: boolean): boolean {
  const text = String(value ?? '').trim();
  if (!text) return false;
  // Схема ищется до первого разделителя пути: `foo/bar:baz` — не схема.
  const schemeMatch = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(text);
  if (!schemeMatch) return !text.startsWith('//');
  const scheme = `${schemeMatch[1].toLowerCase()}:`;
  if (schemes.has(scheme)) return true;
  if (allowDataImage && scheme === 'data:') return /^data:image\/(png|jpe?g|gif|webp|bmp);/i.test(text);
  return false;
}

export function isSafeLinkHref(href: string): boolean {
  return isSafeUrl(href, SAFE_LINK_SCHEMES, false);
}

export function isSafeImageSrc(src: string): boolean {
  return isSafeUrl(src, SAFE_IMAGE_SCHEMES, true);
}

/** Информационная строка код-фенса: в разметку попадает только из документа. */
export function sanitizeFenceLanguage(value: string): string {
  return String(value ?? '')
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9+#._-]/g, '');
}
