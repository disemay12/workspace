import JSZip from 'jszip';
import type { PageSetup } from '../types';
import {
  customStyleCss,
  getCustomStyle,
  listCustomStyles,
  registerCustomStyle,
  type CustomStyleDefinition,
} from './customStyles';

/**
 * Преобразование между HTML-полотном редактора Word и пакетом .docx.
 *
 * mammoth при чтении отбрасывает шрифт, кегль и цвет — здесь они поднимаются
 * из word/document.xml в inline-стили span, а при сохранении HTML собирается
 * обратно в настоящий пакет OOXML, чтобы правки не затирали файл.
 */

/** Исходный XML верхнеуровневого блока документа. */
export interface DocxBlockSource {
  html: string;
  /** Один блок HTML может соответствовать нескольким <w:p> (список). */
  xml: string;
}

export interface DocxPackageInfo {
  parts: string[];
  headers: Array<{ name: string; text: string; referenceType?: string }>;
  footers: Array<{ name: string; text: string; referenceType?: string }>;
  footnotes: string;
  comments: string;
  styleNames: string[];
  sections: number;
  page?: PageSetup;
}

interface AnnotationComment {
  id: number;
  text: string;
  author: string;
}

interface AnnotationFootnote {
  id: number;
  text: string;
}

/** Данные комментариев/сносок текущего разбираемого пакета. */
let currentComments = new Map<number, AnnotationComment>();
let currentFootnotes = new Map<number, string>();
let currentRunId = 0;

/** Служебные части пакета для инструмента get_document_info. */
export async function docxPackageInfo(base64: string): Promise<DocxPackageInfo> {
  const zip = await JSZip.loadAsync(base64ToBytes(base64));
  const parts = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  const readText = async (name: string): Promise<string> => {
    const file = zip.file(name);
    if (!file) return '';
    try {
      const xml = await file.async('string');
      const doc = parseXml(xml);
      return elementsByLocal(doc, 't')
        .map((node) => node.textContent ?? '')
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
    } catch {
      return '';
    }
  };

  const headerNames = parts.filter((name) => /^word\/header\d*\.xml$/.test(name));
  const footerNames = parts.filter((name) => /^word\/footer\d*\.xml$/.test(name));

  const documentXml = (await zip.file('word/document.xml')?.async('string')) ?? '';
  const relsXml =
    (await zip.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
  const referenceTypes = new Map<string, string>();
  for (const tag of documentXml.match(/<w:(?:header|footer)Reference\b[^>]*\/>/g) ?? []) {
    const id = /r:id="([^"]+)"/.exec(tag)?.[1] ?? '';
    const target = relationshipTarget(relsXml, id);
    if (!target) continue;
    referenceTypes.set(normalizePartTarget(target), referenceTypeOf(tag));
  }

  const headers = await Promise.all(
    headerNames.map(async (name) => ({
      name,
      text: await readText(name),
      referenceType: referenceTypes.get(name),
    }))
  );
  const footers = await Promise.all(
    footerNames.map(async (name) => ({
      name,
      text: await readText(name),
      referenceType: referenceTypes.get(name),
    }))
  );

  const stylesXml = (await zip.file('word/styles.xml')?.async('string')) ?? '';
  const styleNames: string[] = [];
  if (stylesXml) {
    const styles = parseXml(stylesXml);
    for (const style of elementsByLocal(styles, 'style')) {
      const id = style.getAttribute('w:styleId') ?? '';
      const name = attr(firstByLocal(style, 'name'), 'val') || id;
      if (name) styleNames.push(name);
    }
  }

  let page: PageSetup | undefined;
  if (documentXml) {
    const doc = parseXml(documentXml);
    const body = firstByLocal(doc, 'body');
    if (body) page = readPageSetup(body, stylesXml ? parseXml(stylesXml) : null);
  }

  return {
    parts,
    headers,
    footers,
    footnotes: await readText('word/footnotes.xml'),
    comments: await readText('word/comments.xml'),
    styleNames,
    sections: (documentXml.match(/<w:sectPr\b/g) || []).length,
    page,
  };
}
export interface DocxHtml {
  html: string;
  /** Предупреждения разбора; документ при этом всё равно открывается. */
  warnings: string[];
  /** Поля, размер страницы и базовый шрифт из sectPr/styles.xml. */
  page?: PageSetup;
  /** Исходные XML-блоки в том же порядке, что и HTML-блоки. */
  blocks?: DocxBlockSource[];
}

/** Сохранение поверх исходного пакета: части, которых нет в HTML, не теряются. */
export interface DocxWriteOptions {
  sourceBase64?: string;
  /** XML исходных блоков по data-docx-id. */
  blockXml?: Map<string, string>;
  /** Правки колонтитулов: применяются к пакету при сохранении. */
  headerFooterEdits?: HeaderFooterEdit[];
}

export interface HeaderFooterEdit {
  kind: 'header' | 'footer';
  referenceType?: 'default' | 'first' | 'even';
  text: string;
  remove?: boolean;
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const HIGHLIGHTS: Record<string, string> = {
  yellow: '#fef08a',
  green: '#bbf7d0',
  cyan: '#a5f3fc',
  magenta: '#f5d0fe',
  blue: '#bfdbfe',
  red: '#fecaca',
  darkBlue: '#1e3a8a',
  darkCyan: '#155e75',
  darkGreen: '#14532d',
  darkMagenta: '#86198f',
  darkRed: '#7f1d1d',
  darkYellow: '#a16207',
  darkGray: '#334155',
  lightGray: '#e2e8f0',
  black: '#0f172a',
  white: '#ffffff',
};

/** Читает .docx (base64) и возвращает HTML с сохранёнными шрифтом, кеглем и цветом. */
export async function docxBase64ToHtml(base64: string): Promise<DocxHtml> {
  const zip = await JSZip.loadAsync(base64ToBytes(base64));
  const docFile = zip.file('word/document.xml');
  if (!docFile) throw new Error('В пакете нет word/document.xml');

  const [docXml, relsXml, numberingXml, stylesDoc, commentsDoc, footnotesDoc] = await Promise.all([
    docFile.async('string'),
    zip.file('word/_rels/document.xml.rels')?.async('string') ?? '',
    zip.file('word/numbering.xml')?.async('string') ?? '',
    zip.file('word/styles.xml')?.async('string') ?? '',
    zip.file('word/comments.xml')?.async('string') ?? '',
    zip.file('word/footnotes.xml')?.async('string') ?? '',
  ]);

  const doc = parseXml(docXml);
  const rels = relsXml ? parseXml(relsXml) : null;
  const numbering = numberingXml ? parseXml(numberingXml) : null;
  const styles = stylesDoc ? parseXml(stylesDoc) : null;
  currentComments = parseComments(commentsDoc);
  currentFootnotes = parseFootnotes(footnotesDoc);
  currentRunId = 0;
  registerStylesFromXml(styles);

  const images = await readImages(zip, rels);
  const listKinds = readListKinds(numbering);
  const listLevelKinds = readListLevelKinds(numbering);

  const body = firstByLocal(doc, 'body');
  if (!body) return { html: '<p></p>', warnings: ['Документ пуст'], blocks: [] };

  interface PendingList {
    items: PendingListItem[];
  }

  const rendered: DocxBlockSource[] = [];
  let pendingList: PendingList | null = null;
  /** Закладки и маркеры комментариев между абзацами: переносим в следующий блок. */
  let pendingMarkers = '';

  const flushList = () => {
    if (!pendingList) return;
    const block = renderPendingList(pendingList.items);
    rendered.push({ html: block.html, xml: block.xml });
    pendingList = null;
  };

  for (const child of childElements(body)) {
    const name = localName(child);
    if (name === 'sectPr') continue;

    if (MARKER_ELEMENTS.has(name)) {
      pendingMarkers += serializeXml(child);
      continue;
    }

    if (name === 'p') {
      const info = paragraphListInfo(child);
      if (info) {
        const kind =
          listLevelKinds.get(info.numId)?.get(info.level) ??
          listKinds.get(info.numId) ??
          'ol';
        if (!pendingList) pendingList = { items: [] };
        pendingList.items.push({
          level: info.level,
          kind,
          inner: await paragraphInner(child, images),
          style: listItemStyle(child),
          xml: serializeXml(child),
        });
        continue;
      }
      flushList();
      rendered.push({
        html: renderParagraph(child, await paragraphInner(child, images)),
        xml: pendingMarkers + serializeXml(child),
      });
      pendingMarkers = '';
      continue;
    }

    if (name === 'tbl') {
      flushList();
      rendered.push({
        html: await renderTable(child, images),
        xml: pendingMarkers + serializeXml(child),
      });
      pendingMarkers = '';
      continue;
    }

    // Всё, что редактор не умеет показывать, сохраняем как read-only блок:
    // исходный XML остаётся в пакете и возвращается при сохранении.
    flushList();
    rendered.push({
      html: readOnlyPlaceholder(name),
      xml: pendingMarkers + serializeXml(child),
    });
    pendingMarkers = '';
  }
  flushList();
  if (pendingMarkers && rendered.length > 0) {
    const last = rendered[rendered.length - 1];
    last.xml += pendingMarkers;
  } else if (pendingMarkers) {
    rendered.push({ html: '<p></p>', xml: pendingMarkers });
  }

  const html = rendered.map((block) => block.html).join('') || '<p></p>';
  return { html, warnings: [], page: readPageSetup(body, styles), blocks: rendered };
}

interface PendingListItem {
  level: number;
  kind: 'ul' | 'ol';
  inner: string;
  style: string;
  xml: string;
}

/** Собирает из плоского списка абзацев вложенные HTML-списки. */
function renderPendingList(items: PendingListItem[]): { html: string; xml: string } {
  if (items.length === 0) return { html: '', xml: '' };

  let html = '';
  const stack: Array<{ level: number; kind: 'ul' | 'ol'; open: boolean }> = [];

  for (const item of items) {
    while (stack.length > 0 && item.level < stack[stack.length - 1].level) {
      const top = stack.pop()!;
      if (top.open) html += '</li>';
      html += `</${top.kind}>`;
    }

    if (stack.length === 0) {
      stack.push({ level: item.level, kind: item.kind, open: false });
      html += `<${item.kind}>`;
    } else {
      const top = stack[stack.length - 1];
      if (item.level > top.level) {
        stack.push({ level: item.level, kind: item.kind, open: false });
        html += `<${item.kind}>`;
      } else {
        if (top.open) {
          html += '</li>';
          top.open = false;
        }
        if (top.kind !== item.kind) {
          html += `</${top.kind}>`;
          stack.pop();
          stack.push({ level: item.level, kind: item.kind, open: false });
          html += `<${item.kind}>`;
        }
      }
    }

    const current = stack[stack.length - 1];
    html += `<li${item.style}>${item.inner}`;
    current.open = true;
  }

  while (stack.length > 0) {
    const top = stack.pop()!;
    if (top.open) html += '</li>';
    html += `</${top.kind}>`;
  }

  return { html, xml: items.map((item) => item.xml).join('') };
}

/** Маркеры, которые не рисуются, но нужны Word: закладки, диапазоны комментариев. */
const MARKER_ELEMENTS = new Set([
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'proofErr',
]);

/** Заглушка для объектов, которые сохраняются без редактирования. */
function readOnlyPlaceholder(name: string): string {
  const label =
    name === 'sdt'
      ? 'Элемент управления содержимым'
      : name === 'drawing' || name === 'pict'
        ? 'Графический объект'
        : name === 'object'
          ? 'Встроенный объект'
          : name === 'altChunk'
            ? 'Внешний фрагмент'
            : name === 'fldSimple'
              ? 'Поле Word'
              : name === 'ins' || name === 'del'
                ? 'Отслеживаемое изменение'
                : `Элемент Word: ${name}`;
  return `<p data-docx-readonly="1" data-docx-kind="${escapeAttr(name)}">[${escapeHtml(label)} — только просмотр]</p>`;
}

/**
 * Читает параметры страницы и базовый шрифт.
 *
 * Word всегда пишет sectPr и стиль Normal, поэтому документ открывается с
 * теми же полями и кеглем, с которыми его сохранили, а не с жёстко
 * зашитыми 2 см и Calibri 12.
 */
function readPageSetup(body: Element, styles: Document | null): PageSetup | undefined {
  const page: PageSetup = {};

  const sections = elementsByLocal(body, 'sectPr');
  const sectPr = sections[sections.length - 1] ?? null;
  if (sectPr) {
    const size = firstByLocal(sectPr, 'pgSz');
    const width = Number(attr(size, 'w'));
    const height = Number(attr(size, 'h'));
    if (width > 0) page.widthMm = twipsToMm(width);
    if (height > 0) page.heightMm = twipsToMm(height);

    const margins = firstByLocal(sectPr, 'pgMar');
    if (margins) {
      const top = Number(attr(margins, 'top'));
      const right = Number(attr(margins, 'right'));
      const bottom = Number(attr(margins, 'bottom'));
      const left = Number(attr(margins, 'left'));
      if (top > 0) page.marginTopMm = twipsToMm(top);
      if (right > 0) page.marginRightMm = twipsToMm(right);
      if (bottom > 0) page.marginBottomMm = twipsToMm(bottom);
      if (left > 0) page.marginLeftMm = twipsToMm(left);
    }
  }

  if (styles) {
    // Normal описывает реальный текст документа; docDefaults — запасной
    // вариант. Атрибуты ищем по отдельности: Word часто оставляет в Normal
    // пустой rPr/pPr и держит шрифт только в docDefaults.
    const normal =
      elementsByLocal(styles, 'style').find(
        (style) => attr(style, 'styleId').toLowerCase() === 'normal'
      ) ?? null;
    const docDefaults = firstByLocal(styles, 'docDefaults');
    const normalRun = normal ? firstByLocal(normal, 'rPr') : null;
    const defaultRun = docDefaults
      ? firstByLocal(firstByLocal(docDefaults, 'rPrDefault') ?? docDefaults, 'rPr')
      : null;
    const normalParagraph = normal ? firstByLocal(normal, 'pPr') : null;
    const defaultParagraph = docDefaults
      ? firstByLocal(firstByLocal(docDefaults, 'pPrDefault') ?? docDefaults, 'pPr')
      : null;

    const font =
      attr(firstByLocal(normalRun, 'rFonts'), 'ascii') ||
      attr(firstByLocal(defaultRun, 'rFonts'), 'ascii');
    if (font) page.fontFamily = font;

    const halfPoints =
      Number(attr(firstByLocal(normalRun, 'sz'), 'val')) ||
      Number(attr(firstByLocal(defaultRun, 'sz'), 'val'));
    if (halfPoints > 0) page.fontSizePt = halfPoints / 2;

    const line =
      Number(attr(firstByLocal(normalParagraph, 'spacing'), 'line')) ||
      Number(attr(firstByLocal(defaultParagraph, 'spacing'), 'line'));
    if (line > 0) page.lineSpacing = Math.round((line / 240) * 100) / 100;
  }

  return Object.keys(page).length > 0 ? page : undefined;
}

/** Собирает .docx из HTML полотна и возвращает его как base64. */
export async function htmlToDocxBase64(
  html: string,
  page?: PageSetup,
  options?: DocxWriteOptions
): Promise<string> {
  const container = document.createElement('div');
  container.innerHTML = html && html.trim() ? html : '<p></p>';

  if (options?.sourceBase64) {
    try {
      return await patchDocxPackage(container, page, options);
    } catch (err) {
      // Повреждённый или чужой пакет не должен блокировать сохранение:
      // собираем документ заново, но сообщаем об этом в консоль.
      console.warn('docx: сохранение поверх исходного пакета не удалось', err);
    }
  }

  return buildDocxFromScratch(container, page);
}

function collectAnnotations(container: HTMLElement): {
  comments: AnnotationComment[];
  footnotes: AnnotationFootnote[];
} {
  const comments: AnnotationComment[] = [];
  const commentIds = new Set<number>();
  for (const element of container.querySelectorAll('[data-comment-id]')) {
    const id = Number(element.getAttribute('data-comment-id'));
    if (!Number.isFinite(id) || commentIds.has(id)) continue;
    const text = element.getAttribute('data-comment-text') ?? '';
    if (!text.trim()) continue;
    commentIds.add(id);
    comments.push({
      id,
      text,
      author: element.getAttribute('data-comment-author') || 'AI',
    });
  }

  const footnotes: AnnotationFootnote[] = [];
  const footnoteIds = new Set<number>();
  for (const element of container.querySelectorAll('[data-footnote-id]')) {
    const id = Number(element.getAttribute('data-footnote-id'));
    if (!Number.isFinite(id) || footnoteIds.has(id)) continue;
    const text =
      element.getAttribute('data-footnote-text') ??
      element.getAttribute('title') ??
      '';
    if (!text.trim()) continue;
    footnoteIds.add(id);
    footnotes.push({ id, text });
  }

  return { comments, footnotes };
}

function commentXml(entry: AnnotationComment): string {
  return (
    `<w:comment w:id="${entry.id}" w:author="${escapeXml(entry.author)}" ` +
    `w:initials="${escapeXml(entry.author.slice(0, 2))}" w:date="${new Date().toISOString()}">` +
    `<w:p><w:r><w:t xml:space="preserve">${escapeXml(entry.text)}</w:t></w:r></w:p></w:comment>`
  );
}

function commentsXml(entries: AnnotationComment[]): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:comments xmlns:w="${W_NS}">${entries.map(commentXml).join('')}</w:comments>`
  );
}

function appendComments(xml: string, entries: AnnotationComment[]): string {
  if (!xml) return commentsXml(entries);
  if (!xml.includes('</w:comments>')) return xml;
  return xml.replace('</w:comments>', `${entries.map(commentXml).join('')}</w:comments>`);
}

function footnoteSeparatorsXml(): string {
  return (
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
  );
}

function footnoteXml(entry: AnnotationFootnote): string {
  return (
    `<w:footnote w:id="${entry.id}"><w:p><w:r><w:t xml:space="preserve">` +
    `${escapeXml(entry.text)}</w:t></w:r></w:p></w:footnote>`
  );
}

function footnotesXml(entries: AnnotationFootnote[]): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:footnotes xmlns:w="${W_NS}">${footnoteSeparatorsXml()}${entries.map(footnoteXml).join('')}</w:footnotes>`
  );
}

function appendFootnotes(xml: string, entries: AnnotationFootnote[]): string {
  if (!xml) return footnotesXml(entries);
  if (!xml.includes('</w:footnotes>')) return xml;
  return xml.replace('</w:footnotes>', `${entries.map(footnoteXml).join('')}</w:footnotes>`);
}

/** Новая сборка пакета: используется для только что созданных документов. */
async function buildDocxFromScratch(container: HTMLElement, page?: PageSetup): Promise<string> {
  const images = new Map<string, ImageEntry>();
  let imageCounter = 0;
  const annotations = collectAnnotations(container);

  const bodyXml = renderChildren(container, images, () => {
    imageCounter += 1;
    return 'rId' + (imageCounter + 1);
  });

  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    contentTypesXml(
      images,
      annotations.comments.length > 0,
      annotations.footnotes.length > 0
    )
  );
  zip.folder('_rels')!.file('.rels', rootRelsXml());
  const word = zip.folder('word')!;
  word.file('document.xml', documentXml(bodyXml || '<w:p/>', page));
  word.file('styles.xml', stylesXml(page, container));
  word.file('numbering.xml', numberingXml());
  word.folder('_rels')!.file(
    'document.xml.rels',
    documentRelsXml(
      images,
      annotations.comments.length > 0,
      annotations.footnotes.length > 0
    )
  );
  if (annotations.comments.length > 0) {
    word.file('comments.xml', commentsXml(annotations.comments));
  }
  if (annotations.footnotes.length > 0) {
    word.file('footnotes.xml', footnotesXml(annotations.footnotes));
  }

  for (const image of images.values()) {
    word.file(`media/${image.id}.${image.ext}`, image.bytes);
  }

  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

/* ------------------------------------------------- Колонтитулы и части */

const HEADER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const FOOTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';

/**
 * Применяет правки колонтитулов к пакету: обновляет part, rels,
 * [Content_Types].xml и ссылку в sectPr. Вызывается при сохранении.
 */
async function editHeaderFooterParts(
  zip: JSZip,
  documentXml: string,
  edits: HeaderFooterEdit[]
): Promise<string> {
  let docXml = documentXml;
  let relsXml = (await zip.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
  let contentTypes = (await zip.file('[Content_Types].xml')?.async('string')) ?? '';
  let relsChanged = false;
  let contentTypesChanged = false;

  for (const edit of edits) {
    const kind = edit.kind;
    const referenceType = edit.referenceType ?? 'default';
    const refs = headerFooterReferences(docXml, kind);
    let partName = '';
    let relationshipId = '';

    for (const ref of refs) {
      if (referenceTypeOf(ref) !== referenceType) continue;
      const id = /r:id="([^"]+)"/.exec(ref)?.[1] ?? '';
      const target = relationshipTarget(relsXml, id);
      if (target) {
        relationshipId = id;
        partName = normalizePartTarget(target);
        break;
      }
    }

    if (edit.remove) {
      if (!partName) continue;
      docXml = removeHeaderFooterReference(docXml, kind, referenceType);
      relsXml = removeRelationship(relsXml, relationshipId);
      relsChanged = true;
      zip.remove(partName);
      contentTypes = removeContentTypeOverride(contentTypes, partName);
      contentTypesChanged = true;
      continue;
    }

    if (!partName) {
      partName = nextHeaderFooterPart(zip, kind);
      relationshipId = nextRelationshipId(relsXml);
      relsXml = appendRelationship(relsXml, relationshipId, kind, partName);
      relsChanged = true;
      contentTypes = appendContentTypeOverride(contentTypes, partName, kind);
      contentTypesChanged = true;
      docXml = appendHeaderFooterReference(docXml, kind, referenceType, relationshipId);
      if (referenceType === 'first') docXml = ensureTitlePage(docXml);
      if (referenceType === 'even') await ensureEvenAndOddHeaders(zip);
      if (!/xmlns:r=/.test(docXml.slice(0, 4000))) {
        docXml = docXml.replace(/<w:document\b/, `<w:document xmlns:r="${R_NS}"`);
      }
    }

    zip.file(partName, headerFooterXml(kind, edit.text));
  }

  if (relsChanged) zip.file('word/_rels/document.xml.rels', relsXml);
  if (contentTypesChanged && contentTypes) {
    zip.file('[Content_Types].xml', contentTypes);
  }
  return docXml;
}

/** Тестовая обёртка: правка колонтитулов без изменения тела документа. */
export async function docxWithHeaderFooterEdits(
  base64: string,
  edits: HeaderFooterEdit[]
): Promise<string> {
  const zip = await JSZip.loadAsync(base64ToBytes(base64));
  const documentXml = (await zip.file('word/document.xml')?.async('string')) ?? '';
  const next = await editHeaderFooterParts(zip, documentXml, edits);
  zip.file('word/document.xml', next);
  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

function headerFooterReferences(documentXml: string, kind: 'header' | 'footer'): string[] {
  const pattern = kind === 'header' ? /<w:headerReference\b[^>]*\/>/g : /<w:footerReference\b[^>]*\/>/g;
  return documentXml.match(pattern) ?? [];
}

function referenceTypeOf(tag: string): string {
  return /w:type="([^"]+)"/.exec(tag)?.[1] ?? 'default';
}

function relationshipTarget(relsXml: string, id: string): string {
  if (!relsXml || !id) return '';
  const pattern = new RegExp(`<Relationship\\b[^>]*Id="${escapeRegExp(id)}"[^>]*/>`);
  const tag = pattern.exec(relsXml)?.[0] ?? '';
  return /Target="([^"]+)"/.exec(tag)?.[1] ?? '';
}

function normalizePartTarget(target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  return `word/${target.replace(/^\.\.\//, '')}`;
}

function nextHeaderFooterPart(zip: JSZip, kind: 'header' | 'footer'): string {
  let max = 0;
  for (const name of Object.keys(zip.files)) {
    const match = new RegExp(`^word/${kind}(\\d+)\\.xml$`).exec(name);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return `word/${kind}${max + 1}.xml`;
}

function nextRelationshipId(relsXml: string): string {
  const used = new Set([...relsXml.matchAll(/Id="([^"]+)"/g)].map((match) => match[1]));
  let index = 1;
  while (used.has(`rIdHF${index}`)) index += 1;
  return `rIdHF${index}`;
}

function appendRelationship(
  relsXml: string,
  id: string,
  kind: 'header' | 'footer',
  partName: string
): string {
  const type = kind === 'header' ? 'header' : 'footer';
  const entry =
    `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" ` +
    `Target="${partName.replace(/^word\//, '')}"/>`;
  if (!relsXml) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${REL_NS}">${entry}</Relationships>`
    );
  }
  return relsXml.replace('</Relationships>', `${entry}</Relationships>`);
}

function removeRelationship(relsXml: string, id: string): string {
  if (!relsXml || !id) return relsXml;
  return relsXml.replace(
    new RegExp(`<Relationship\\b[^>]*Id="${escapeRegExp(id)}"[^>]*/>`),
    ''
  );
}

function appendContentTypeOverride(
  xml: string,
  partName: string,
  kind: 'header' | 'footer'
): string {
  if (!xml) return xml;
  const type = kind === 'header' ? HEADER_CONTENT_TYPE : FOOTER_CONTENT_TYPE;
  return xml.replace(
    '</Types>',
    `<Override PartName="/${partName}" ContentType="${type}"/></Types>`
  );
}

function removeContentTypeOverride(xml: string, partName: string): string {
  if (!xml) return xml;
  return xml.replace(
    new RegExp(`<Override\\b[^>]*PartName="/${escapeRegExp(partName)}"[^>]*/>`),
    ''
  );
}

function appendHeaderFooterReference(
  documentXml: string,
  kind: 'header' | 'footer',
  referenceType: string,
  relationshipId: string
): string {
  const tag = `<w:${kind}Reference w:type="${referenceType}" r:id="${relationshipId}"/>`;
  const closes = [...documentXml.matchAll(/<\/w:sectPr>/g)];
  if (closes.length === 0) {
    return documentXml.replace('</w:body>', `<w:sectPr>${tag}</w:sectPr></w:body>`);
  }
  const last = closes[closes.length - 1];
  const position = last.index ?? documentXml.length;
  return `${documentXml.slice(0, position)}${tag}${documentXml.slice(position)}`;
}

function removeHeaderFooterReference(
  documentXml: string,
  kind: 'header' | 'footer',
  referenceType: string
): string {
  const pattern = kind === 'header' ? /<w:headerReference\b[^>]*\/>/g : /<w:footerReference\b[^>]*\/>/g;
  return documentXml.replace(pattern, (tag) =>
    referenceTypeOf(tag) === referenceType ? '' : tag
  );
}

function ensureTitlePage(documentXml: string): string {
  if (documentXml.includes('<w:titlePg/>')) return documentXml;
  const closes = [...documentXml.matchAll(/<\/w:sectPr>/g)];
  if (closes.length === 0) return documentXml;
  const last = closes[closes.length - 1];
  const position = last.index ?? documentXml.length;
  return `${documentXml.slice(0, position)}<w:titlePg/>${documentXml.slice(position)}`;
}

async function ensureEvenAndOddHeaders(zip: JSZip): Promise<void> {
  const file = zip.file('word/settings.xml');
  if (!file) return;
  let xml = await file.async('string');
  if (xml.includes('<w:evenAndOddHeaders/>')) return;
  xml = xml.replace(/(<w:settings\b[^>]*>)/, '$1<w:evenAndOddHeaders/>');
  zip.file('word/settings.xml', xml);
}

function headerFooterXml(kind: 'header' | 'footer', text: string): string {
  const root = kind === 'header' ? 'w:hdr' : 'w:ftr';
  const paragraphs = (text || '')
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map((line) => `<w:p>${headerFooterRuns(line)}</w:p>`)
    .join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<${root} xmlns:w="${W_NS}" xmlns:r="${R_NS}">${paragraphs || '<w:p/>'}</${root}>`
  );
}

/** Текст колонтитула: {{page}} и {{pages}} превращаются в поля Word. */
function headerFooterRuns(line: string): string {
  const parts = line.split(/(\{\{page\}\}|\{\{pages\}\})/i);
  let runs = '';
  for (const part of parts) {
    if (!part) continue;
    if (/^\{\{page\}\}$/i.test(part)) {
      runs += fieldRun('PAGE', '1');
    } else if (/^\{\{pages\}\}$/i.test(part)) {
      runs += fieldRun('NUMPAGES', '1');
    } else {
      runs += `<w:r><w:t xml:space="preserve">${escapeXml(part)}</w:t></w:r>`;
    }
  }
  return runs || '<w:r><w:t></w:t></w:r>';
}

function fieldRun(instruction: string, placeholder: string): string {
  return (
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    `<w:r><w:t>${placeholder}</w:t></w:r>` +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  );
}

/** Сохраняет документ поверх исходного пакета. */
async function patchDocxPackage(
  container: HTMLElement,
  page: PageSetup | undefined,
  options: DocxWriteOptions
): Promise<string> {
  const sourceBase64 = options.sourceBase64 as string;
  const zip = await JSZip.loadAsync(base64ToBytes(sourceBase64));
  const docFile = zip.file('word/document.xml');
  if (!docFile) throw new Error('В пакете нет word/document.xml');

  const docXml = await docFile.async('string');
  const doc = parseXml(docXml);
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error('word/document.xml не разбирается');
  }
  const body = firstByLocal(doc, 'body');
  if (!body) throw new Error('В word/document.xml нет body');
  const annotations = collectAnnotations(container);
  let relsDirty = false;

  const stylesFile = zip.file('word/styles.xml');
  const stylesText = stylesFile ? await stylesFile.async('string') : '';
  const originalPage = readPageSetup(body, stylesText ? parseXml(stylesText) : null);
  const pageChanged = page ? !samePage(originalPage, page) : false;

  // Картинки из исходного пакета: data URL → существующий rId, чтобы не
  // дублировать media при каждом сохранении.
  const images = new Map<string, ImageEntry>();
  const usedIds = new Set<string>();
  let relsXml = (await zip.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
  if (relsXml) {
    const rels = parseXml(relsXml);
    for (const rel of elementsByLocal(rels, 'Relationship')) {
      const id = rel.getAttribute('Id') ?? '';
      if (!id) continue;
      usedIds.add(id);
      const type = rel.getAttribute('Type') ?? '';
      const target = rel.getAttribute('Target') ?? '';
      if (!type.endsWith('/image') || !target) continue;
      const path = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\.\//, '')}`;
      const file = zip.file(path) ?? zip.file(decodeURIComponent(path));
      if (!file) continue;
      const bytes = new Uint8Array(await file.async('uint8array'));
      images.set(id, {
        id,
        bytes,
        contentType: mimeOf(path),
        ext: path.split('.').pop()?.toLowerCase() || 'png',
      });
    }
  }

  let imageSeq = 1;
  const nextId = () => {
    let id = '';
    do {
      id = `rIdImg${imageSeq++}`;
    } while (usedIds.has(id));
    usedIds.add(id);
    return id;
  };

  const blocks = options.blockXml ?? new Map<string, string>();
  const usedBlocks = new Set<string>();
  const bodyParts: string[] = [];
  const ids: string[] = [];
  let unchanged = true;
  let strayText = false;

  for (const node of [...container.childNodes]) {
    if (node.nodeType === Node.TEXT_NODE) {
      if ((node.textContent ?? '').trim()) {
        unchanged = false;
        strayText = true;
        bodyParts.push(paragraphXml(runsFromInline(node, {}, images, nextId), {}));
      }
      continue;
    }
    if (!(node instanceof Element)) continue;

    const element = node;
    const id = element.getAttribute('data-docx-id') ?? '';
    const sig = element.getAttribute('data-docx-sig') ?? '';
    const source = id ? blocks.get(id) : undefined;
    ids.push(id);

    if (id && source !== undefined && !usedBlocks.has(id) && sig === docxBlockSignature(element)) {
      bodyParts.push(source);
      usedBlocks.add(id);
      continue;
    }

    unchanged = false;
    bodyParts.push(renderBlockElement(element, images, nextId));
  }

  const orderUnchanged = ids.every((id, index) => id === `d${index + 1}`);
  if (
    unchanged &&
    !strayText &&
    orderUnchanged &&
    !pageChanged &&
    usedBlocks.size === blocks.size &&
    container.children.length === blocks.size
  ) {
    // Ни текста, ни порядка, ни оформления не тронули: отдаём файл байт в байт.
    return sourceBase64;
  }

  const directSectPr =
    childElements(body)
      .filter((element) => localName(element) === 'sectPr')
      .pop() ?? null;
  let sectPrXml = directSectPr ? serializeXml(directSectPr) : '';
  if (!sectPrXml) {
    sectPrXml = `<w:sectPr>${sectPrBodyXml(page)}</w:sectPr>`;
  } else if (pageChanged && page) {
    sectPrXml = applyPageToSectPr(sectPrXml, page);
  }

  const openTag = /<w:body(?:\s[^>]*)?>/.exec(docXml);
  const closeIndex = docXml.lastIndexOf('</w:body>');
  if (!openTag || closeIndex < 0) throw new Error('Не найден w:body в word/document.xml');
  const innerStart = openTag.index + openTag[0].length;
  let nextDocumentXml =
    docXml.slice(0, innerStart) + bodyParts.join('') + sectPrXml + docXml.slice(closeIndex);

  if (options.headerFooterEdits?.length) {
    nextDocumentXml = await editHeaderFooterParts(
      zip,
      nextDocumentXml,
      options.headerFooterEdits
    );
  }
  zip.file('word/document.xml', nextDocumentXml);

  if (annotations.comments.length > 0) {
    const existingComments = (await zip.file('word/comments.xml')?.async('string')) ?? '';
    const existingIds = new Set(
      [...existingComments.matchAll(/<w:comment\b[^>]*w:id="(\d+)"/g)].map((match) =>
        Number(match[1])
      )
    );
    const missing = annotations.comments.filter((entry) => !existingIds.has(entry.id));
    if (missing.length > 0) {
      zip.file('word/comments.xml', appendComments(existingComments, missing));
      if (!existingComments) {
        await ensurePartContentType(
          zip,
          '/word/comments.xml',
          'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
        );
        if (!/Type="[^"]*\/comments"/.test(relsXml)) {
          relsXml = ensureDocumentRelationship(
            relsXml,
            'rIdComments',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
            'comments.xml'
          );
          relsDirty = true;
        }
      }
    }
  }

  if (annotations.footnotes.length > 0) {
    const existingFootnotes = (await zip.file('word/footnotes.xml')?.async('string')) ?? '';
    const existingIds = new Set(
      [...existingFootnotes.matchAll(/<w:footnote\b[^>]*w:id="(\d+)"/g)].map((match) =>
        Number(match[1])
      )
    );
    const missing = annotations.footnotes.filter((entry) => !existingIds.has(entry.id));
    if (missing.length > 0) {
      zip.file('word/footnotes.xml', appendFootnotes(existingFootnotes, missing));
      if (!existingFootnotes) {
        await ensurePartContentType(
          zip,
          '/word/footnotes.xml',
          'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
        );
        if (!/Type="[^"]*\/footnotes"/.test(relsXml)) {
          relsXml = ensureDocumentRelationship(
            relsXml,
            'rIdFootnotes',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
            'footnotes.xml'
          );
          relsDirty = true;
        }
      }
    }
  }

  if (pageChanged && page && stylesText) {
    zip.file('word/styles.xml', patchStyles(stylesText, page, originalPage));
  }
  if (stylesText) {
    const withCustom = appendCustomStyles(stylesText, container);
    if (withCustom !== stylesText) zip.file('word/styles.xml', withCustom);
  }

  const fresh = [...images.values()].filter((image) => !isExistingRel(image.id, relsXml));
  if (fresh.length > 0) {
    for (const image of fresh) zip.file(`word/media/${image.id}.${image.ext}`, image.bytes);
    relsXml = appendImageRels(relsXml, fresh);
    const contentTypes = (await zip.file('[Content_Types].xml')?.async('string')) ?? '';
    zip.file('[Content_Types].xml', ensureImageContentTypes(contentTypes, fresh));
    relsDirty = true;
  }
  if (relsDirty) zip.file('word/_rels/document.xml.rels', relsXml);

  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

/* ------------------------------------------------------- привязка блоков */

/**
 * Привязывает HTML-блоки к исходному XML и возвращает карту id → XML.
 * Вызывается после установки HTML в contenteditable.
 */
export function bindDocxBlockSources(
  root: HTMLElement,
  sources?: DocxBlockSource[]
): Map<string, string> {
  const map = new Map<string, string>();
  if (!sources?.length) return map;

  [...root.children].forEach((element, index) => {
    const source = sources[index];
    if (!source) return;
    const id = `d${index + 1}`;
    element.setAttribute('data-docx-id', id);
    element.setAttribute('data-docx-sig', docxBlockSignature(element));
    map.set(id, source.xml);
  });
  return map;
}

/** Подпись блока без служебных атрибутов редактора. */
export function docxBlockSignature(element: Element): string {
  const clone = element.cloneNode(true) as Element;
  for (const node of [clone, ...clone.querySelectorAll('*')]) {
    node.removeAttribute('data-bid');
    node.removeAttribute('data-docx-id');
    node.removeAttribute('data-docx-sig');
  }
  return hash32(clone.outerHTML);
}

function hash32(text: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

function serializeXml(element: Element): string {
  return new XMLSerializer().serializeToString(element);
}

function renderBlockElement(
  element: Element,
  images: Map<string, ImageEntry>,
  nextId: () => string
): string {
  const wrapper = document.createElement('div');
  wrapper.appendChild(element.cloneNode(true));
  return renderChildren(wrapper, images, nextId);
}

/* ----------------------------------------------------- оформление страницы */

function samePage(a: PageSetup | undefined, b: PageSetup | undefined): boolean {
  if (!a || !b) return !a && !b;
  const keys: (keyof PageSetup)[] = [
    'widthMm',
    'heightMm',
    'marginTopMm',
    'marginRightMm',
    'marginBottomMm',
    'marginLeftMm',
    'fontFamily',
    'fontSizePt',
    'lineSpacing',
  ];
  return keys.every((key) => {
    const left = a[key];
    const right = b[key];
    if (typeof left === 'number' && typeof right === 'number') return Math.abs(left - right) < 0.6;
    return (left ?? '') === (right ?? '');
  });
}

function sectPrBodyXml(page?: PageSetup): string {
  const width = page?.widthMm ? mmToTwips(page.widthMm) : 11906;
  const height = page?.heightMm ? mmToTwips(page.heightMm) : 16838;
  const top = page?.marginTopMm ? mmToTwips(page.marginTopMm) : 1134;
  const right = page?.marginRightMm ? mmToTwips(page.marginRightMm) : 1134;
  const bottom = page?.marginBottomMm ? mmToTwips(page.marginBottomMm) : 1134;
  const left = page?.marginLeftMm ? mmToTwips(page.marginLeftMm) : 1134;
  return (
    `<w:pgSz w:w="${width}" w:h="${height}"/>` +
    `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}"/>`
  );
}

/** Меняет только заданные поля, не пересобирая sectPr целиком. */
function applyPageToSectPr(sectPr: string, page: PageSetup): string {
  let next = sectPr;

  if (page.widthMm || page.heightMm) {
    const attrs: string[] = [];
    if (page.widthMm) attrs.push(`w:w="${mmToTwips(page.widthMm)}"`);
    if (page.heightMm) attrs.push(`w:h="${mmToTwips(page.heightMm)}"`);
    next = upsertTag(next, 'w:pgSz', attrs, '<w:pgSz w:w="11906" w:h="16838"/>');
  }

  const margins: string[] = [];
  if (page.marginTopMm) margins.push(`w:top="${mmToTwips(page.marginTopMm)}"`);
  if (page.marginRightMm) margins.push(`w:right="${mmToTwips(page.marginRightMm)}"`);
  if (page.marginBottomMm) margins.push(`w:bottom="${mmToTwips(page.marginBottomMm)}"`);
  if (page.marginLeftMm) margins.push(`w:left="${mmToTwips(page.marginLeftMm)}"`);
  if (margins.length) {
    next = upsertTag(next, 'w:pgMar', margins, '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>');
  }

  return next;
}

/** Вставляет или обновляет одиночный тег в XML-фрагменте. */
function upsertTag(xml: string, tag: string, attrs: string[], fallback: string): string {
  const pattern = new RegExp(`<${tag}\\b[^>]*/>`);
  const apply = (current: string) => attrs.reduce((acc, attr) => setTagAttr(acc, attr), current);
  if (!pattern.test(xml)) {
    return xml.replace('</w:sectPr>', `${apply(fallback)}</w:sectPr>`);
  }
  return xml.replace(pattern, (current) => apply(current));
}

/** attr в формате w:top="1134". */
function setTagAttr(tag: string, attr: string): string {
  const [name, raw] = attr.split('=');
  const value = raw.replace(/^"|"$/g, '');
  const attrPattern = new RegExp(`\\s${name}="[^"]*"`);
  return attrPattern.test(tag)
    ? tag.replace(attrPattern, ` ${name}="${value}"`)
    : tag.replace(/\/>$/, ` ${name}="${value}"/>`);
}

/**
 * Правка styles.xml: меняются только docDefaults и Normal, остальные
 * named styles (Heading, Quote, пользовательские) остаются как в файле.
 */
function patchStyles(xml: string, page: PageSetup, original: PageSetup | undefined): string {
  const fontChanged = (page.fontFamily ?? '') !== (original?.fontFamily ?? '') && !!page.fontFamily;
  const sizeChanged =
    typeof page.fontSizePt === 'number' &&
    Math.abs((page.fontSizePt ?? 0) - (original?.fontSizePt ?? 0)) > 0.1;
  const lineChanged =
    typeof page.lineSpacing === 'number' &&
    Math.abs((page.lineSpacing ?? 0) - (original?.lineSpacing ?? 0)) > 0.01;

  if (!fontChanged && !sizeChanged && !lineChanged) return xml;

  const font = (page.fontFamily ?? '').replace(/["']/g, '');
  const half = page.fontSizePt ? Math.max(2, Math.round(page.fontSizePt * 2)) : 0;
  const line = page.lineSpacing ? Math.max(1, Math.round(240 * page.lineSpacing)) : 0;

  const fontTag = fontChanged
    ? `<w:rFonts w:ascii="${escapeXml(font)}" w:hAnsi="${escapeXml(font)}" w:cs="${escapeXml(font)}"/>`
    : '';
  const sizeTags =
    sizeChanged && half ? `<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>` : '';
  const spacingTag =
    lineChanged && line ? `<w:spacing w:line="${line}" w:lineRule="auto"/>` : '';

  const apply = (block: string) => {
    let next = block;
    if (fontTag) next = patchRunProp(next, 'rFonts', fontTag);
    if (sizeTags) {
      next = patchRunProp(next, 'sz', `<w:sz w:val="${half}"/>`);
      next = patchRunProp(next, 'szCs', `<w:szCs w:val="${half}"/>`);
    }
    if (spacingTag) next = patchParagraphSpacing(next, line);
    return next;
  };

  let next = xml;
  const defaults = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(next);
  if (defaults) {
    next = next.replace(defaults[0], apply(defaults[0]));
  } else {
    const created =
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      fontTag +
      sizeTags +
      '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' +
      spacingTag +
      '</w:pPr></w:pPrDefault></w:docDefaults>';
    next = next.replace(/<w:styles\b[^>]*>/, (open) => open + created);
  }

  const normal = /<w:style\b[^>]*w:styleId="Normal"[^>]*>[\s\S]*?<\/w:style>/.exec(next);
  if (normal) {
    next = next.replace(normal[0], apply(normal[0]));
  } else {
    const created =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr>' +
      spacingTag +
      '</w:pPr><w:rPr>' +
      fontTag +
      sizeTags +
      '</w:rPr></w:style>';
    next = next.replace('</w:styles>', `${created}</w:styles>`);
  }

  return next;
}

function patchRunProp(block: string, tag: string, replacement: string): string {
  const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(block);
  if (!rPr) return block;
  let next = rPr[0];
  const pattern = new RegExp(`<w:${tag}\\b[^>]*/>`);
  next = pattern.test(next)
    ? next.replace(pattern, replacement)
    : next.replace('</w:rPr>', `${replacement}</w:rPr>`);
  return block.replace(rPr[0], next);
}

function patchParagraphSpacing(block: string, line: number): string {
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(block);
  if (!pPr) return block;
  let next = pPr[0];
  if (/<w:spacing\b[^>]*\/>/.test(next)) {
    next = next.replace(/<w:spacing\b[^>]*\/>/, (tag) => {
      let patched = /\sw:line="[^"]*"/.test(tag)
        ? tag.replace(/\sw:line="[^"]*"/, ` w:line="${line}"`)
        : tag.replace(/\/>$/, ` w:line="${line}"/>`);
      patched = /\sw:lineRule="[^"]*"/.test(patched)
        ? patched.replace(/\sw:lineRule="[^"]*"/, ' w:lineRule="auto"')
        : patched.replace(/\/>$/, ' w:lineRule="auto"/>');
      return patched;
    });
  } else {
    next = next.replace('</w:pPr>', `<w:spacing w:line="${line}" w:lineRule="auto"/></w:pPr>`);
  }
  return block.replace(pPr[0], next);
}

/* ------------------------------------------------------------- картинки */

function isExistingRel(id: string, relsXml: string): boolean {
  return !!relsXml && new RegExp(`Id="${escapeRegExp(id)}"`).test(relsXml);
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function appendImageRels(xml: string, images: ImageEntry[]): string {
  const entries = images
    .map(
      (image) =>
        `<Relationship Id="${image.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${image.id}.${image.ext}"/>`
    )
    .join('');
  if (!xml) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${REL_NS}">${entries}</Relationships>`
    );
  }
  return xml.replace('</Relationships>', `${entries}</Relationships>`);
}

function ensureImageContentTypes(xml: string, images: ImageEntry[]): string {
  if (!xml) return '';
  let next = xml;
  for (const image of images) {
    if (new RegExp(`Extension="${escapeRegExp(image.ext)}"`, 'i').test(next)) continue;
    next = next.replace(
      '</Types>',
      `<Default Extension="${escapeAttr(image.ext)}" ContentType="${escapeAttr(image.contentType)}"/></Types>`
    );
  }
  return next;
}

async function ensurePartContentType(
  zip: JSZip,
  part: string,
  contentType: string
): Promise<void> {
  const file = zip.file('[Content_Types].xml');
  if (!file) return;
  const xml = await file.async('string');
  if (xml.includes(`PartName="${part}"`)) return;
  zip.file(
    '[Content_Types].xml',
    xml.replace(
      '</Types>',
      `<Override PartName="${part}" ContentType="${contentType}"/></Types>`
    )
  );
}

function ensureDocumentRelationship(
  xml: string,
  preferredId: string,
  type: string,
  target: string
): string {
  let id = preferredId;
  let index = 1;
  while (xml.includes(`Id="${id}"`)) {
    id = `${preferredId}${index}`;
    index += 1;
  }
  const entry = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  if (!xml) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${REL_NS}">${entry}</Relationships>`
    );
  }
  return xml.replace('</Relationships>', `${entry}</Relationships>`);
}

async function readImages(zip: JSZip, rels: Document | null): Promise<Map<string, string>> {
  const images = new Map<string, string>();
  if (!rels) return images;

  const targets = new Map<string, string>();
  for (const rel of elementsByLocal(rels, 'Relationship')) {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    const type = rel.getAttribute('Type') ?? '';
    if (id && target && type.endsWith('/image')) targets.set(id, target);
  }

  await Promise.all(
    [...targets.entries()].map(async ([id, target]) => {
      const path = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\.\//, '')}`;
      const file = zip.file(path) ?? zip.file(decodeURIComponent(path));
      if (!file) return;
      const bytes = new Uint8Array(await file.async('uint8array'));
      images.set(id, `data:${mimeOf(path)};base64,${bytesToBase64(bytes)}`);
    })
  );
  return images;
}

function readListKinds(numbering: Document | null): Map<string, 'ul' | 'ol'> {
  const kinds = new Map<string, 'ul' | 'ol'>();
  if (!numbering) return kinds;

  const abstract = new Map<string, 'ul' | 'ol'>();
  for (const abs of elementsByLocal(numbering, 'abstractNum')) {
    const id = abs.getAttribute('w:abstractNumId');
    const level = firstByLocal(abs, 'lvl');
    const fmt = level ? attr(firstByLocal(level, 'numFmt'), 'val') : '';
    if (id) abstract.set(id, fmt === 'bullet' ? 'ul' : 'ol');
  }
  for (const num of elementsByLocal(numbering, 'num')) {
    const id = num.getAttribute('w:numId');
    const absId = attr(firstByLocal(num, 'abstractNumId'), 'val');
    if (id) kinds.set(id, abstract.get(absId ?? '') ?? 'ol');
  }
  return kinds;
}

function readListLevelKinds(numbering: Document | null): Map<string, Map<number, 'ul' | 'ol'>> {
  const result = new Map<string, Map<number, 'ul' | 'ol'>>();
  if (!numbering) return result;

  const abstracts = new Map<string, Map<number, 'ul' | 'ol'>>();
  for (const abs of elementsByLocal(numbering, 'abstractNum')) {
    const id = abs.getAttribute('w:abstractNumId');
    if (!id) continue;
    const levels = new Map<number, 'ul' | 'ol'>();
    for (const level of elementsByLocal(abs, 'lvl')) {
      const ilvl = Number(level.getAttribute('w:ilvl') ?? 0);
      const format = attr(firstByLocal(level, 'numFmt'), 'val');
      levels.set(ilvl, format === 'bullet' ? 'ul' : 'ol');
    }
    abstracts.set(id, levels);
  }

  for (const num of elementsByLocal(numbering, 'num')) {
    const id = num.getAttribute('w:numId');
    const absId = attr(firstByLocal(num, 'abstractNumId'), 'val');
    if (id) result.set(id, abstracts.get(absId ?? '') ?? new Map());
  }
  return result;
}

function paragraphListInfo(p: Element): { numId: string; level: number } | null {
  const pPr = firstByLocal(p, 'pPr');
  const numPr = pPr ? firstByLocal(pPr, 'numPr') : null;
  if (!numPr) return null;
  const numId = attr(firstByLocal(numPr, 'numId'), 'val');
  if (!numId) return null;
  const level = Math.max(0, Math.min(8, Number(attr(firstByLocal(numPr, 'ilvl'), 'val')) || 0));
  return { numId, level };
}

/** Оформление элемента списка: выравнивание и отступы. */
function listItemStyle(p: Element): string {
  const pPr = firstByLocal(p, 'pPr');
  const align = pPr ? attr(firstByLocal(pPr, 'jc'), 'val') : '';
  const css = [
    align ? `text-align:${cssAlign(align)}` : '',
    paragraphIndent(pPr),
    paragraphFormatCss(pPr),
  ]
    .filter(Boolean)
    .join(';');
  return css ? ` style="${css}"` : '';
}

function parseComments(xml: string): Map<number, AnnotationComment> {
  const result = new Map<number, AnnotationComment>();
  if (!xml) return result;
  const doc = parseXml(xml);
  for (const comment of elementsByLocal(doc, 'comment')) {
    const id = Number(attr(comment, 'id'));
    if (!Number.isFinite(id)) continue;
    result.set(id, {
      id,
      text: elementsByLocal(comment, 't')
        .map((node) => node.textContent ?? '')
        .join(' '),
      author: attr(comment, 'author') || 'Автор',
    });
  }
  return result;
}

function parseFootnotes(xml: string): Map<number, string> {
  const result = new Map<number, string>();
  if (!xml) return result;
  const doc = parseXml(xml);
  for (const footnote of elementsByLocal(doc, 'footnote')) {
    const type = attr(footnote, 'type');
    if (type === 'separator' || type === 'continuationSeparator') continue;
    const id = Number(attr(footnote, 'id'));
    if (!Number.isFinite(id) || id < 0) continue;
    result.set(
      id,
      elementsByLocal(footnote, 't')
        .map((node) => node.textContent ?? '')
        .join(' ')
    );
  }
  return result;
}

function renderParagraph(p: Element, inner: string): string {
  const pPr = firstByLocal(p, 'pPr');
  const style = pPr ? attr(firstByLocal(pPr, 'pStyle'), 'val') : '';
  const align = pPr ? attr(firstByLocal(pPr, 'jc'), 'val') : '';
  const custom = style ? getCustomStyle(style) : undefined;
  const css = [
    align ? `text-align:${cssAlign(align)}` : '',
    paragraphIndent(pPr),
    paragraphFormatCss(pPr),
    custom ? customStyleCss(custom) : '',
  ]
    .filter(Boolean)
    .join(';');
  const styleAttr = css ? ` style="${css}"` : '';
  const tag = headingTag(style);
  const isToc = elementsByLocal(p, 'instrText').some((node) =>
    /\bTOC\b/i.test(node.textContent ?? '')
  );
  const isSectionBreak = !!(pPr && firstByLocal(pPr, 'sectPr'));
  const fieldAttr = isToc ? ' data-field="TOC"' : isSectionBreak ? ' data-section-break="1"' : '';
  const tabStops = pPr
    ? elementsByLocal(pPr, 'tab')
        .map((tab) => Number(attr(tab, 'pos')) / 20)
        .filter((value) => Number.isFinite(value) && value > 0)
    : [];
  const tabsAttr = tabStops.length ? ` data-tabs="${tabStops.join(',')}"` : '';
  const namedStyleAttr = style ? ` data-style="${escapeAttr(style)}"` : '';
  const widow = pPr ? firstByLocal(pPr, 'widowControl') : null;
  const widowAttr = widow ? ` data-widow="${isOn(widow) ? 'on' : 'off'}"` : '';
  return `<${tag}${fieldAttr}${tabsAttr}${namedStyleAttr}${widowAttr}${styleAttr}>${inner || '<br>'}</${tag}>`;
}

/** Интервалы, отступы до/после и запреты разрыва из pPr. */
function paragraphFormatCss(pPr: Element | null): string {
  if (!pPr) return '';
  const parts: string[] = [];
  const spacing = firstByLocal(pPr, 'spacing');
  const line = Number(attr(spacing, 'line'));
  if (line > 0) parts.push(`line-height:${Math.round((line / 240) * 100) / 100}`);
  const before = Number(attr(spacing, 'before'));
  if (before > 0) parts.push(`margin-top:${Math.round((before / 20) * 10) / 10}pt`);
  const after = Number(attr(spacing, 'after'));
  if (after > 0) parts.push(`margin-bottom:${Math.round((after / 20) * 10) / 10}pt`);

  if (firstByLocal(pPr, 'keepNext')) parts.push('break-after:avoid');
  if (firstByLocal(pPr, 'keepLines')) parts.push('break-inside:avoid');
  if (firstByLocal(pPr, 'pageBreakBefore')) parts.push('break-before:page');

  const shading = firstByLocal(pPr, 'shd');
  const fill = attr(shading, 'fill');
  if (fill && fill.toLowerCase() !== 'auto') parts.push(`background-color:#${fill}`);

  const borders = firstByLocal(pPr, 'pBdr');
  const top = borders ? firstByLocal(borders, 'top') : null;
  if (top) {
    const size = Number(attr(top, 'sz'));
    const color = attr(top, 'color') || 'auto';
    const val = attr(top, 'val') || 'single';
    const style = val === 'dashed' || val === 'dotted' || val === 'double' ? val : 'solid';
    if (size > 0) {
      parts.push(
        `border:${Math.max(0.25, size / 8)}pt ${style} #${color === 'auto' ? '000000' : color}`
      );
    }
  }
  return parts.join(';');
}

/** Отступы абзаца из OOXML: twips → pt для CSS. */
function paragraphIndent(pPr: Element | null): string {
  if (!pPr) return '';
  const ind = firstByLocal(pPr, 'ind');
  if (!ind) return '';

  const parts: string[] = [];
  const left = Number(attr(ind, 'left') || attr(ind, 'start') || 0);
  const firstLine = Number(attr(ind, 'firstLine') || 0);
  const hanging = Number(attr(ind, 'hanging') || 0);

  if (left > 0) parts.push(`margin-left:${left / 20}pt`);
  if (firstLine > 0) parts.push(`text-indent:${firstLine / 20}pt`);
  if (hanging > 0) parts.push(`text-indent:-${hanging / 20}pt`);
  return parts.join(';');
}

function headingTag(style: string): string {
  const name = style.toLowerCase().replace(/\s+/g, '');
  if (name === 'title' || name === 'heading1') return 'h1';
  if (name === 'heading2') return 'h2';
  if (name === 'heading3') return 'h3';
  if (name === 'heading4') return 'h4';
  if (name === 'quote' || name === 'intensequote') return 'blockquote';
  return 'p';
}

async function paragraphInner(p: Element, images: Map<string, string>): Promise<string> {
  let html = '';
  for (const child of childElements(p)) {
    const name = localName(child);
    if (name === 'r') html += renderRun(child, images);
    else if (name === 'hyperlink') {
      const href = '#';
      let inner = '';
      for (const run of elementsByLocal(child, 'r')) inner += renderRun(run, images);
      html += `<a href="${escapeAttr(href)}">${inner}</a>`;
    } else if (name === 'commentRangeStart') {
      const id = Number(attr(child, 'id'));
      const info = currentComments.get(id);
      const meta = info
        ? ` data-comment-text="${escapeAttr(info.text)}" data-comment-author="${escapeAttr(info.author)}" title="${escapeAttr(`${info.author}: ${info.text}`)}"`
        : '';
      html += `<span data-comment-id="${id}"${meta}>`;
    } else if (name === 'commentRangeEnd') {
      html += '</span>';
    } else if (name === 'bookmarkStart') {
      const bookmark = attr(child, 'name');
      if (bookmark) html += `<span data-bookmark="${escapeAttr(bookmark)}">`;
    } else if (name === 'bookmarkEnd') {
      html += '</span>';
    } else if (name === 'fldSimple') {
      const instruction = attr(child, 'instr');
      const href = /HYPERLINK\s+"([^"]+)"/i.exec(instruction)?.[1] ?? '';
      const fieldKind = /^\s*(PAGE|NUMPAGES)\s*$/i.exec(instruction)?.[1]?.toUpperCase() ?? '';
      let inner = '';
      for (const run of elementsByLocal(child, 'r')) inner += renderRun(run, images);
      html += href
        ? `<a href="${escapeAttr(href)}">${inner}</a>`
        : fieldKind
          ? `<span data-field="${fieldKind}">${inner || '1'}</span>`
          : inner;
    }
  }
  return html;
}

function renderRun(run: Element, images: Map<string, string>): string {
  const runId = `r${++currentRunId}`;
  const footnoteRef = firstByLocal(run, 'footnoteReference');
  if (footnoteRef) {
    const id = Number(attr(footnoteRef, 'id'));
    const text = currentFootnotes.get(id) ?? '';
    return `<sup data-footnote-ref="1" data-footnote-id="${id}" data-footnote-text="${escapeAttr(text)}" title="${escapeAttr(text)}">${id}</sup>`;
  }
  const rPr = firstByLocal(run, 'rPr');
  let text = '';
  for (const child of childElements(run)) {
    const name = localName(child);
    if (name === 't') text += escapeHtml(child.textContent ?? '');
    else if (name === 'tab') text += '&emsp;';
    else if (name === 'br') {
      text += attr(child, 'type') === 'page' ? '<br data-page-break="1">' : '<br>';
    } else if (name === 'cr') text += '<br>';
    else if (name === 'drawing' || name === 'pict') {
      const blip = firstByLocal(child, 'blip');
      const id = blip?.getAttribute('r:embed') ?? blip?.getAttributeNS(R_NS, 'embed') ?? '';
      const src = images.get(id);
      if (src) text += `<img src="${src}" alt="">`;
    }
  }
  if (!text) return '';
  return `<span data-run-id="${runId}">${wrapRun(text, rPr)}</span>`;
}

function wrapRun(text: string, rPr: Element | null): string {
  if (!rPr) return text;
  let html = text;

  const vert = attr(firstByLocal(rPr, 'vertAlign'), 'val');
  if (vert === 'superscript') html = `<sup>${html}</sup>`;
  else if (vert === 'subscript') html = `<sub>${html}</sub>`;

  if (isOn(firstByLocal(rPr, 'strike'))) html = `<s>${html}</s>`;
  if (isOn(firstByLocal(rPr, 'u'))) html = `<u>${html}</u>`;
  if (isOn(firstByLocal(rPr, 'i'))) html = `<em>${html}</em>`;
  if (isOn(firstByLocal(rPr, 'b'))) html = `<strong>${html}</strong>`;

  const styles: string[] = [];
  const fonts = firstByLocal(rPr, 'rFonts');
  const font = fonts?.getAttribute('w:ascii') || fonts?.getAttribute('w:hAnsi') || fonts?.getAttribute('w:cs');
  if (font) styles.push(`font-family:${font}`);

  const sz = attr(firstByLocal(rPr, 'sz'), 'val');
  if (sz && /^\d+$/.test(sz)) styles.push(`font-size:${Number(sz) / 2}pt`);

  const color = attr(firstByLocal(rPr, 'color'), 'val');
  if (color && color.toLowerCase() !== 'auto') styles.push(`color:#${color}`);

  const highlight = attr(firstByLocal(rPr, 'highlight'), 'val');
  if (highlight && HIGHLIGHTS[highlight]) styles.push(`background-color:${HIGHLIGHTS[highlight]}`);

  const shading = attr(firstByLocal(rPr, 'shd'), 'fill');
  if (shading && shading.toLowerCase() !== 'auto') styles.push(`background-color:#${shading}`);

  if (!styles.length) return html;
  return `<span style="${styles.join(';')}">${html}</span>`;
}

async function renderTable(tbl: Element, images: Map<string, string>): Promise<string> {
  const rows: string[] = [];
  for (const tr of elementsByLocal(tbl, 'tr')) {
    const cells: string[] = [];
    for (const tc of elementsByLocal(tr, 'tc')) {
      const paras: string[] = [];
      for (const p of elementsByLocal(tc, 'p')) paras.push(await paragraphInner(p, images));
      cells.push(`<td>${paras.join('<br>')}</td>`);
    }
    rows.push(`<tr>${cells.join('')}</tr>`);
  }
  return `<table>${rows.join('')}</table>`;
}

/* --------------------------------------------------------------- запись */

interface ImageEntry {
  id: string;
  bytes: Uint8Array;
  contentType: string;
  ext: string;
}

function renderChildren(
  parent: Node,
  images: Map<string, ImageEntry>,
  nextId: () => string
): string {
  let xml = '';
  let list: 'ul' | 'ol' | null = null;

  const closeList = () => {
    list = null;
  };

  for (const node of [...parent.childNodes]) {
    if (node.nodeType === Node.TEXT_NODE) {
      if ((node.textContent ?? '').trim()) {
        closeList();
        xml += paragraphXml(runsFromInline(node, {}, images, nextId), {});
      }
      continue;
    }
    if (!(node instanceof Element)) continue;
    const tag = node.tagName.toLowerCase();

    if (tag === 'ul' || tag === 'ol') {
      closeList();
      xml += emitListXml(node, 0, tag, images, nextId);
      continue;
    }

    if (tag === 'table') {
      closeList();
      xml += tableXml(node, images, nextId);
      continue;
    }

    if (tag === 'br') {
      closeList();
      xml += paragraphXml(
        (node as Element).hasAttribute('data-page-break')
          ? '<w:r><w:br w:type="page"/></w:r>'
          : '',
        {}
      );
      continue;
    }

    if (node.getAttribute('data-field') === 'TOC') {
      closeList();
      xml += tocParagraphXml(node);
      continue;
    }

    closeList();
    const align = readAlign(node);
    const namedStyle = node.getAttribute('data-style') || undefined;
    const style = namedStyle || blockStyle(tag);
    xml += paragraphXml(runsFromInline(node, {}, images, nextId), {
      align,
      style,
      namedStyle,
      ...readIndent(node),
      ...paragraphFormatOptions(node),
      sectionBreak: node.hasAttribute('data-section-break'),
    });
  }
  return xml;
}

/** Рекурсивно раскладывает вложенные списки в абзацы Word с ilvl. */
function emitListXml(
  list: Element,
  level: number,
  kind: 'ul' | 'ol',
  images: Map<string, ImageEntry>,
  nextId: () => string
): string {
  let xml = '';
  for (const item of [...list.children]) {
    if (item.tagName.toLowerCase() !== 'li') continue;
    let runs = '';
    for (const child of [...item.childNodes]) {
      if (child instanceof Element) {
        const tag = child.tagName.toLowerCase();
        if (tag === 'ul' || tag === 'ol') continue;
      }
      runs += runsFromInline(child, {}, images, nextId);
    }
    xml += paragraphXml(runs, {
      list: kind,
      listLevel: Math.min(level, 8),
      align: readAlign(item),
      ...readIndent(item),
      ...paragraphFormatOptions(item),
    });

    for (const nested of [...item.children]) {
      const tag = nested.tagName.toLowerCase();
      if (tag === 'ul' || tag === 'ol') {
        xml += emitListXml(
          nested,
          Math.min(level + 1, 8),
          tag,
          images,
          nextId
        );
      }
    }
  }
  return xml;
}

interface ParagraphOptions {
  align?: string;
  style?: string;
  /** Точный styleId из Word (data-style), важнее эвристики по тегу. */
  namedStyle?: string;
  list?: 'ul' | 'ol';
  /** Уровень вложенности списка, 0–8. */
  listLevel?: number;
  /** Абзац внутри таблицы: без отступа Normal, иначе ячейки раздуваются. */
  tight?: boolean;
  indentLeftPt?: number;
  firstLinePt?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  /** Заливка блока: #RRGGBB без решётки. */
  shading?: string;
  border?: { widthPt: number; style: string; color: string };
  /** Абзац завершает секцию: в pPr добавляется sectPr. */
  sectionBreak?: boolean;
  /** Позиции табуляции в пунктах. */
  tabStops?: number[];
  /** Твипы: 1 pt = 20 twips. */
  spaceBeforeTwips?: number;
  spaceAfterTwips?: number;
  /** Межстрочный интервал: 240 = одинарный. */
  lineTwips?: number;
}

function paragraphXml(runs: string, options: ParagraphOptions): string {
  const props: string[] = [];
  const pStyle = options.namedStyle || options.style;
  if (pStyle) props.push(`<w:pStyle w:val="${escapeXml(pStyle)}"/>`);
  if (options.keepNext) props.push('<w:keepNext/>');
  if (options.keepLines) props.push('<w:keepLines/>');
  if (options.pageBreakBefore) props.push('<w:pageBreakBefore/>');
  if (options.widowControl === false) props.push('<w:widowControl w:val="0"/>');
  else if (options.widowControl === true) props.push('<w:widowControl/>');
  if (options.list) {
    const level = Math.max(0, Math.min(options.listLevel ?? 0, 8));
    props.push(`<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${options.list === 'ul' ? '1' : '2'}"/></w:numPr>`);
  }
  if (options.border) {
    const size = Math.max(2, Math.round(options.border.widthPt * 8));
    const val = options.border.style;
    const color = options.border.color.replace('#', '');
    props.push(
      `<w:pBdr><w:top w:val="${val}" w:sz="${size}" w:space="1" w:color="${color}"/>` +
        `<w:left w:val="${val}" w:sz="${size}" w:space="1" w:color="${color}"/>` +
        `<w:bottom w:val="${val}" w:sz="${size}" w:space="1" w:color="${color}"/>` +
        `<w:right w:val="${val}" w:sz="${size}" w:space="1" w:color="${color}"/></w:pBdr>`
    );
  }
  if (options.shading) {
    props.push(`<w:shd w:val="clear" w:color="auto" w:fill="${options.shading}"/>`);
  }
  if (options.tabStops?.length) {
    props.push(
      '<w:tabs>' +
        options.tabStops
          .map((stop) => `<w:tab w:val="left" w:pos="${Math.round(stop * 20)}"/>`)
          .join('') +
        '</w:tabs>'
    );
  }
  if (options.tight) {
    // В ячейках таблицы абзацный отступ Normal раздувает строки.
    props.push('<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>');
  } else {
    const spacing: string[] = [];
    if (options.spaceBeforeTwips) spacing.push(`w:before="${Math.round(options.spaceBeforeTwips)}"`);
    if (options.spaceAfterTwips) spacing.push(`w:after="${Math.round(options.spaceAfterTwips)}"`);
    if (options.lineTwips) {
      spacing.push(`w:line="${Math.round(options.lineTwips)}"`, 'w:lineRule="auto"');
    }
    if (spacing.length > 0) props.push(`<w:spacing ${spacing.join(' ')}/>`);
  }
  if (options.indentLeftPt || options.firstLinePt) {
    const left = Math.round((options.indentLeftPt ?? 0) * 20);
    const firstLine = Math.round((options.firstLinePt ?? 0) * 20);
    const attrs: string[] = [];
    if (left) attrs.push(`w:left="${left}"`);
    if (firstLine > 0) attrs.push(`w:firstLine="${firstLine}"`);
    else if (firstLine < 0) attrs.push(`w:hanging="${Math.abs(firstLine)}"`);
    if (attrs.length > 0) props.push(`<w:ind ${attrs.join(' ')}/>`);
  }
  if (options.align && options.align !== 'left') {
    props.push(`<w:jc w:val="${options.align}"/>`);
  }
  if (options.sectionBreak) {
    props.push(`<w:sectPr>${sectPrBodyXml()}</w:sectPr>`);
  }
  const pPr = props.length ? `<w:pPr>${props.join('')}</w:pPr>` : '';
  return `<w:p>${pPr}${runs || '<w:r><w:t></w:t></w:r>'}</w:p>`;
}

/** Поле TOC: Word обновит список при открытии/обновлении полей. */
function tocParagraphXml(element: Element): string {
  const style = blockStyle(element.tagName.toLowerCase());
  const pPr = style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : '';
  const lines = (element.textContent ?? 'Оглавление')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);
  const cached =
    lines.length > 0
      ? lines
          .map((line) => `<w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r>`)
          .join('<w:r><w:br/></w:r>')
      : '<w:r><w:t>Оглавление</w:t></w:r>';
  return (
    `<w:p>${pPr}<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    cached +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
  );
}

/** Отступы, интервалы и запреты разрыва из inline-стилей блока. */
function paragraphFormatOptions(el: Element): ParagraphOptions {
  const css = el.getAttribute('style') ?? '';
  const lineValue = cssValue(css, 'line-height');
  let lineTwips = 0;
  if (lineValue) {
    const factor = Number(lineValue);
    if (Number.isFinite(factor) && factor > 0) {
      lineTwips = Math.round(240 * factor);
    } else {
      const pt = cssLengthToPt(lineValue);
      if (pt) lineTwips = Math.round(pt * 20);
    }
  }

  const beforePt = cssLengthToPt(cssValue(css, 'margin-top'));
  const afterPt = cssLengthToPt(cssValue(css, 'margin-bottom'));
  const fill = cssColorToHex(cssValue(css, 'background-color'));
  const tabStops = (el.getAttribute('data-tabs') ?? '')
    .split(',')
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value) && value > 0);
  const widowAttr = el.getAttribute('data-widow');
  return {
    keepNext: cssValue(css, 'break-after') === 'avoid',
    keepLines: cssValue(css, 'break-inside') === 'avoid',
    pageBreakBefore: cssValue(css, 'break-before') === 'page',
    spaceBeforeTwips: beforePt ? Math.round(beforePt * 20) : undefined,
    spaceAfterTwips: afterPt ? Math.round(afterPt * 20) : undefined,
    lineTwips: lineTwips || undefined,
    shading: fill && fill !== 'FFFFFF' ? fill : undefined,
    border: parseCssBorder(cssValue(css, 'border')),
    tabStops: tabStops.length ? tabStops : undefined,
    widowControl: widowAttr === 'on' ? true : widowAttr === 'off' ? false : undefined,
  };
}

function parseCssBorder(value: string): ParagraphOptions['border'] {
  const match = /^([\d.]+)\s*(pt|px)?\s+(solid|dashed|dotted|double)\s+(#[0-9a-f]{3,6}|[a-z]+)$/i.exec(
    value.trim()
  );
  if (!match) return undefined;
  let width = Number(match[1]);
  if (!Number.isFinite(width) || width <= 0) return undefined;
  if ((match[2] || 'pt').toLowerCase() === 'px') width = (width * 3) / 4;
  return {
    widthPt: Math.max(0.25, width),
    style: match[3].toLowerCase(),
    color: cssColorToHex(match[4]) || '000000',
  };
}

/** Отступы из inline-стилей блока: margin-left / padding-left и text-indent. */
function readIndent(el: Element): { indentLeftPt?: number; firstLinePt?: number } {
  const css = el.getAttribute('style') ?? '';
  const left = cssLengthToPt(cssValue(css, 'margin-left') || cssValue(css, 'padding-left'));
  const firstLine = cssLengthToPt(cssValue(css, 'text-indent'));
  return {
    indentLeftPt: left ?? undefined,
    firstLinePt: firstLine ?? undefined,
  };
}

/** Размер картинки в px: CSS width/height или атрибуты width/height. */
function cssLengthToPx(value: string): number | null {
  if (!value) return null;
  const match = value.match(/^([\d.]+)\s*(px|pt|cm|mm)?$/i);
  if (!match) return null;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || amount <= 0) return null;
  const unit = (match[2] || 'px').toLowerCase();
  if (unit === 'pt') return Math.round(amount * (4 / 3));
  if (unit === 'cm') return Math.round((amount / 2.54) * 96);
  if (unit === 'mm') return Math.round((amount / 25.4) * 96);
  return Math.round(amount);
}

function numericAttr(element: Element, name: string): number | null {
  const value = Number(element.getAttribute(name));
  return Number.isFinite(value) && value > 0 ? value : null;
}

function cssLengthToPt(value: string): number | null {
  if (!value) return null;
  const match = value.match(/^(-?[\d.]+)\s*(pt|px|em|rem)?$/i);
  if (!match) return null;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return null;
  const unit = (match[2] || 'pt').toLowerCase();
  if (unit === 'px') return Math.round(amount * 0.75 * 2) / 2;
  if (unit === 'em' || unit === 'rem') return Math.round(amount * 15 * 2) / 2;
  return amount;
}

interface RunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  sup?: boolean;
  sub?: boolean;
  font?: string;
  sizePt?: number;
  color?: string;
  /** Цвет фона текста в #RRGGBB: в Word это не highlight, а заливка w:shd. */
  shading?: string;
}

function runsFromInline(
  node: Node,
  style: RunStyle,
  images: Map<string, ImageEntry>,
  nextId: () => string
): string {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent ?? '';
    if (!text) return '';
    return runXml(text, style);
  }
  if (!(node instanceof Element)) return '';

  const tag = node.tagName.toLowerCase();
  if (tag === 'sup' && node.hasAttribute('data-footnote-ref')) return '';
  const commentId = Number(node.getAttribute('data-comment-id'));
  if (node.hasAttribute('data-comment-id') && Number.isFinite(commentId)) {
    let inner = '';
    for (const child of [...node.childNodes]) {
      inner += runsFromInline(child, style, images, nextId);
    }
    return (
      `<w:commentRangeStart w:id="${commentId}"/>${inner}` +
      `<w:commentRangeEnd w:id="${commentId}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>` +
      `<w:commentReference w:id="${commentId}"/></w:r>`
    );
  }
  const footnoteId = Number(node.getAttribute('data-footnote-id'));
  if (node.hasAttribute('data-footnote-id') && Number.isFinite(footnoteId)) {
    let inner = '';
    for (const child of [...node.childNodes]) {
      inner += runsFromInline(child, style, images, nextId);
    }
    return (
      `${inner}<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
      `<w:footnoteReference w:id="${footnoteId}"/></w:r>`
    );
  }
  const bookmark = node.getAttribute('data-bookmark');
  if (bookmark) {
    const id = bookmarkId(bookmark);
    let inner = '';
    for (const child of [...node.childNodes]) {
      inner += runsFromInline(child, style, images, nextId);
    }
    return `<w:bookmarkStart w:id="${id}" w:name="${escapeXml(bookmark)}"/>${inner}<w:bookmarkEnd w:id="${id}"/>`;
  }
  const field = node.getAttribute('data-field');
  if (field === 'PAGE' || field === 'NUMPAGES') {
    const text = (node.textContent ?? '').trim() || '1';
    return `<w:fldSimple w:instr=" ${field} "><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:fldSimple>`;
  }
  if (tag === 'br') {
    return node.hasAttribute('data-page-break')
      ? '<w:r><w:br w:type="page"/></w:r>'
      : '<w:r><w:br/></w:r>';
  }

  if (tag === 'img') {
    const src = node.getAttribute('src') ?? '';
    const image = takeImage(src, images, nextId);
    if (!image) return '';
    const css = node.getAttribute('style') ?? '';
    const width = cssLengthToPx(cssValue(css, 'width')) ?? numericAttr(node, 'width') ?? 576;
    const height = cssLengthToPx(cssValue(css, 'height')) ?? numericAttr(node, 'height') ?? 336;
    const alt = node.getAttribute('alt') ?? '';
    return imageRunXml(image.id, width, height, alt);
  }

  if (tag === 'a') {
    const href = node.getAttribute('href') ?? '';
    let inner = '';
    for (const child of [...node.childNodes]) {
      inner += runsFromInline(child, style, images, nextId);
    }
    if (!href || !/^(https?:\/\/|mailto:|tel:)/i.test(href)) return inner;
    return `<w:fldSimple w:instr="HYPERLINK &quot;${escapeXml(href)}&quot;">${inner}</w:fldSimple>`;
  }

  const next: RunStyle = { ...style, ...styleFromElement(node) };
  let xml = '';
  for (const child of [...node.childNodes]) xml += runsFromInline(child, next, images, nextId);
  return xml;
}

function styleFromElement(el: Element): RunStyle {
  const tag = el.tagName.toLowerCase();
  const style: RunStyle = {};
  if (tag === 'strong' || tag === 'b') style.bold = true;
  if (tag === 'em' || tag === 'i') style.italic = true;
  if (tag === 'u') style.underline = true;
  if (tag === 's' || tag === 'strike' || tag === 'del') style.strike = true;
  if (tag === 'sup') style.sup = true;
  if (tag === 'sub') style.sub = true;

  const css = el.getAttribute('style') ?? '';
  const font = cssValue(css, 'font-family');
  if (font) style.font = font.split(',')[0].replace(/["']/g, '').trim();

  const size = cssValue(css, 'font-size');
  const pt = size ? cssSizeToPt(size) : null;
  if (pt) style.sizePt = pt;

  const color = cssValue(css, 'color');
  if (color) style.color = cssColorToHex(color);

  const bg = cssValue(css, 'background-color') || cssValue(css, 'background');
  if (bg && bg !== 'transparent') {
    const fill = cssColorToHex(bg);
    if (isMeaningfulFill(fill)) style.shading = fill;
  }

  return style;
}

function runXml(text: string, style: RunStyle): string {
  // Порядок элементов в w:rPr задан схемой OOXML: rFonts, b, i, strike,
  // color, sz, szCs, u, shd, vertAlign. Word терпит нарушения, но валидаторы
  // и часть версий — нет.
  const props: string[] = [];
  if (style.font) {
    const name = escapeXml(style.font);
    props.push(`<w:rFonts w:ascii="${name}" w:hAnsi="${name}" w:cs="${name}"/>`);
  }
  if (style.bold) props.push('<w:b/>');
  if (style.italic) props.push('<w:i/>');
  if (style.strike) props.push('<w:strike/>');
  if (style.color) props.push(`<w:color w:val="${style.color}"/>`);

  // 12 pt — базовый кегль Normal: не дублируем его в каждом run, иначе
  // документ приходит в Word не тем размером, что видно в редакторе.
  const half = style.sizePt ? Math.round(style.sizePt * 2) : 0;
  if (half >= 16 && Math.abs(half - 24) > 1) {
    props.push(`<w:sz w:val="${half}"/>`, `<w:szCs w:val="${half}"/>`);
  }

  if (style.underline) props.push('<w:u w:val="single"/>');
  if (style.shading) {
    props.push(`<w:shd w:val="clear" w:color="auto" w:fill="${style.shading}"/>`);
  }
  if (style.sup) props.push('<w:vertAlign w:val="superscript"/>');
  if (style.sub) props.push('<w:vertAlign w:val="subscript"/>');

  const rPr = props.length ? `<w:rPr>${props.join('')}</w:rPr>` : '';
  const preserve = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
  return `<w:r>${rPr}<w:t${preserve}>${escapeXml(text)}</w:t></w:r>`;
}

/**
 * Идентификаторы объектов внутри документа должны быть уникальны,
 * иначе Word отказывается открывать файл («Ошибка Word при попытке
 * открытия файла») при нескольких картинках.
 */
let nextDrawingId = 1;

function imageRunXml(id: string, widthPx = 576, heightPx = 336, alt = ''): string {
  const drawingId = nextDrawingId++;
  // 1 px = 9525 EMU при 96 DPI — размеры из редактора совпадают с Word.
  const cx = Math.max(1, Math.round(widthPx * 9525));
  const cy = Math.max(1, Math.round(heightPx * 9525));
  const description = alt ? ` descr="${escapeAttr(alt)}"` : '';
  return (
    '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ` distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${drawingId}" name="Image ${drawingId}"${description}/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    // nvPicPr обязателен: без него Word считает пакет повреждённым.
    `<pic:nvPicPr><pic:cNvPr id="${drawingId}" name="Image ${drawingId}"${description}/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="${R_NS}" r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  );
}

function tableXml(table: Element, images: Map<string, ImageEntry>, nextId: () => string): string {
  const rows: string[] = [];
  for (const tr of [...table.querySelectorAll(':scope > tbody > tr, :scope > tr')]) {
    const cells: string[] = [];
    for (const tc of [...tr.querySelectorAll(':scope > td, :scope > th')]) {
      const isHeader = tc.tagName.toLowerCase() === 'th';
      const runs = runsFromInline(tc, isHeader ? { bold: true } : {}, images, nextId);
      const colspan = Math.max(1, Number(tc.getAttribute('colspan') || 1));
      const background = cssColorToHex(cssValue(tc.getAttribute('style') ?? '', 'background-color'));
      const widthPt = cssLengthToPt(cssValue(tc.getAttribute('style') ?? '', 'width'));
      const widthTwips = widthPt ? Math.round(widthPt * 20) : 0;
      // Заливка шапки и небольшие отступы: без них таблица в Word выглядит
      // голой сеткой, а не как в редакторе.
      const props =
        `<w:tcPr><w:tcW w:w="${widthTwips}" w:type="${widthTwips ? 'dxa' : 'auto'}"/>` +
        (colspan > 1 ? `<w:gridSpan w:val="${colspan}"/>` : '') +
        (background
          ? `<w:shd w:val="clear" w:color="auto" w:fill="${background}"/>`
          : isHeader
            ? '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
            : '') +
        '</w:tcPr>';
      cells.push(
        `<w:tc>${props}${paragraphXml(runs, { tight: true, align: readAlign(tc) })}</w:tc>`
      );
    }
    rows.push(`<w:tr>${cells.join('')}</w:tr>`);
  }
  const borderSpec = table.getAttribute('data-table-border') ?? '';
  const borderParts = borderSpec.split('|');
  const noBorder = borderSpec === 'none';
  const borderStyle = noBorder ? 'nil' : borderParts[1] || 'single';
  const borderSize = noBorder
    ? 0
    : Math.max(2, Math.round((Number(borderParts[0]) || 0.5) * 8));
  const borderColor = noBorder
    ? 'auto'
    : (borderParts[2] || '#A6A6A6').replace('#', '').toUpperCase();
  const edge = (name: string) =>
    `<w:${name} w:val="${borderStyle}" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;

  return (
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>' +
    '<w:tblBorders>' +
    edge('top') +
    edge('left') +
    edge('bottom') +
    edge('right') +
    edge('insideH') +
    edge('insideV') +
    '</w:tblBorders>' +
    '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
    '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>' +
    '</w:tblPr>' +
    `${rows.join('')}</w:tbl>`
  );
}

function takeImage(src: string, images: Map<string, ImageEntry>, nextId: () => string): ImageEntry | null {
  const match = src.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) return null;
  const existing = [...images.values()].find((image) => image.bytes.length > 0 && `data:${image.contentType};base64,` + bytesToBase64(image.bytes) === src);
  if (existing) return existing;
  const contentType = match[1];
  const entry: ImageEntry = {
    id: nextId(),
    bytes: base64ToBytes(match[2]),
    contentType,
    ext: extOf(contentType),
  };
  images.set(entry.id, entry);
  return entry;
}

function blockStyle(tag: string): string {
  if (tag === 'h1') return 'Heading1';
  if (tag === 'h2') return 'Heading2';
  if (tag === 'h3') return 'Heading3';
  if (tag === 'h4') return 'Heading4';
  if (tag === 'blockquote') return 'Quote';
  return '';
}

function readAlign(el: Element): string {
  const css = cssValue(el.getAttribute('style') ?? '', 'text-align');
  if (css === 'center') return 'center';
  if (css === 'right') return 'right';
  if (css === 'justify') return 'both';
  const prop = (el as HTMLElement).style?.textAlign;
  if (prop === 'center' || prop === 'right' || prop === 'justify') return prop === 'justify' ? 'both' : prop;
  return '';
}

/* ---------------------------------------------------------------- пакет */

function documentXml(body: string, page?: PageSetup): string {
  const width = page?.widthMm ? mmToTwips(page.widthMm) : 11906;
  const height = page?.heightMm ? mmToTwips(page.heightMm) : 16838;
  const top = page?.marginTopMm ? mmToTwips(page.marginTopMm) : 1134;
  const right = page?.marginRightMm ? mmToTwips(page.marginRightMm) : 1134;
  const bottom = page?.marginBottomMm ? mmToTwips(page.marginBottomMm) : 1134;
  const left = page?.marginLeftMm ? mmToTwips(page.marginLeftMm) : 1134;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}">` +
    `<w:body>${body}<w:sectPr><w:pgSz w:w="${width}" w:h="${height}"/>` +
    `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}"/></w:sectPr></w:body></w:document>`
  );
}

function stylesXml(page: PageSetup | undefined, container: HTMLElement): string {
  const font = (page?.fontFamily || 'Calibri').replace(/["']/g, '');
  const half = page?.fontSizePt ? Math.max(2, Math.round(page.fontSizePt * 2)) : 24;
  const line = page?.lineSpacing ? Math.max(1, Math.round(240 * page.lineSpacing)) : 276;
  const heading = (id: string, size: number, before: number, after: number, outline: number) =>
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id.replace(/(\d)/, ' $1')}"/>` +
    `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="${before}" w:after="${after}"/><w:outlineLvl w:val="${outline}"/></w:pPr>` +
    `<w:rPr><w:b/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:styles xmlns:w="${W_NS}">` +
    // Базовый кегль — из параметров страницы: документ совпадает с редактором.
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>` +
    `<w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="${line}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>' +
    `<w:pPr><w:spacing w:after="160" w:line="${line}" w:lineRule="auto"/></w:pPr>` +
    `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr></w:style>` +
    heading('Heading1', 44, 240, 120, 0) +
    heading('Heading2', 32, 240, 100, 1) +
    heading('Heading3', 26, 200, 80, 2) +
    heading('Heading4', 24, 160, 80, 3) +
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:ind w:left="360"/><w:spacing w:after="160"/></w:pPr>' +
    '<w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>' +
    customStylesXml(container) +
    '</w:styles>'
  );
}

function alignToXml(value: CustomStyleDefinition['align']): string {
  if (value === 'center') return 'center';
  if (value === 'right') return 'right';
  if (value === 'justify') return 'both';
  return 'left';
}

function customStyleXml(definition: CustomStyleDefinition): string {
  const pPr: string[] = [];
  if (definition.align) pPr.push(`<w:jc w:val="${alignToXml(definition.align)}"/>`);
  const spacing: string[] = [];
  if (definition.lineHeight) {
    spacing.push(`w:line="${Math.round(definition.lineHeight * 240)}" w:lineRule="auto"`);
  }
  if (definition.spaceBeforePt !== undefined) {
    spacing.push(`w:before="${Math.round(definition.spaceBeforePt * 20)}"`);
  }
  if (definition.spaceAfterPt !== undefined) {
    spacing.push(`w:after="${Math.round(definition.spaceAfterPt * 20)}"`);
  }
  if (spacing.length) pPr.push(`<w:spacing ${spacing.join(' ')}/>`);

  const rPr: string[] = [];
  if (definition.font) {
    const font = escapeXml(definition.font);
    rPr.push(`<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>`);
  }
  if (definition.bold) rPr.push('<w:b/>');
  if (definition.italic) rPr.push('<w:i/>');
  const color = (definition.color ?? '').replace('#', '');
  if (/^[0-9a-f]{3,8}$/i.test(color)) rPr.push(`<w:color w:val="${color}"/>`);
  if (definition.sizePt) {
    const half = Math.max(2, Math.round(definition.sizePt * 2));
    rPr.push(`<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`);
  }

  return (
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${escapeXml(definition.id)}">` +
    `<w:name w:val="${escapeXml(definition.name)}"/>` +
    `<w:basedOn w:val="${escapeXml(definition.basedOn || 'Normal')}"/><w:qFormat/>` +
    (pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : '') +
    (rPr.length ? `<w:rPr>${rPr.join('')}</w:rPr>` : '') +
    '</w:style>'
  );
}

function referencedCustomStyles(container: HTMLElement): CustomStyleDefinition[] {
  const ids = new Set<string>();
  for (const element of container.querySelectorAll('[data-style]')) {
    const id = element.getAttribute('data-style');
    if (id) ids.add(id);
  }
  return listCustomStyles().filter((definition) => ids.has(definition.id));
}

function customStylesXml(container: HTMLElement): string {
  return referencedCustomStyles(container).map(customStyleXml).join('');
}

function appendCustomStyles(xml: string, container: HTMLElement): string {
  let result = xml;
  for (const definition of referencedCustomStyles(container)) {
    if (result.includes(`w:styleId="${definition.id}"`)) continue;
    result = result.replace('</w:styles>', `${customStyleXml(definition)}</w:styles>`);
  }
  return result;
}

function registerStylesFromXml(styles: Document | null): void {
  if (!styles) return;
  for (const node of elementsByLocal(styles, 'style')) {
    if (attr(node, 'customStyle') !== '1') continue;
    const id = attr(node, 'styleId');
    if (!id) continue;
    const name = attr(firstByLocal(node, 'name'), 'val') || id;
    const rPr = firstByLocal(node, 'rPr');
    const pPr = firstByLocal(node, 'pPr');
    const font = rPr ? attr(firstByLocal(rPr, 'rFonts'), 'ascii') : '';
    const size = rPr ? Number(attr(firstByLocal(rPr, 'sz'), 'val')) / 2 : 0;
    const color = rPr ? attr(firstByLocal(rPr, 'color'), 'val') : '';
    const bold = rPr ? firstByLocal(rPr, 'b') : null;
    const italic = rPr ? firstByLocal(rPr, 'i') : null;
    const align = pPr ? attr(firstByLocal(pPr, 'jc'), 'val') : '';
    const spacing = pPr ? firstByLocal(pPr, 'spacing') : null;
    registerCustomStyle({
      id,
      name,
      basedOn: attr(firstByLocal(node, 'basedOn'), 'val') || 'Normal',
      font: font || undefined,
      sizePt: Number.isFinite(size) && size > 0 ? size : undefined,
      bold: bold ? isOn(bold) : undefined,
      italic: italic ? isOn(italic) : undefined,
      color: color ? `#${color}` : undefined,
      align:
        align === 'center' || align === 'right'
          ? align
          : align === 'both'
            ? 'justify'
            : align === 'left'
              ? 'left'
              : undefined,
      lineHeight: spacing
        ? Number(attr(spacing, 'line')) > 0
          ? Number(attr(spacing, 'line')) / 240
          : undefined
        : undefined,
      spaceBeforePt:
        spacing && attr(spacing, 'before') ? Number(attr(spacing, 'before')) / 20 : undefined,
      spaceAfterPt:
        spacing && attr(spacing, 'after') ? Number(attr(spacing, 'after')) / 20 : undefined,
    });
  }
}

/**
 * Описание списков: numId 1 — маркированный, 2 — нумерованный.
 * Без numbering.xml Word и собственный разбор считают любые списки
 * нумерованными, и маркеры «•» превращаются в «1. 2. 3.».
 */
function numberingXml(): string {
  // Маркер списка — обычный U+2022 в шрифте документа. С Symbol этот символ
  // не отображается: Word рисует «□» (нет глифа), хотя в редакторе точка.
  const levels = (fmt: 'bullet' | 'decimal') =>
    Array.from({ length: 9 }, (_, index) => {
      const text = fmt === 'bullet' ? '\u2022' : `%${index + 1}.`;
      return (
        `<w:lvl w:ilvl="${index}"><w:start w:val="1"/>` +
        `<w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
        `<w:pPr><w:ind w:left="${720 * (index + 1)}" w:hanging="360"/></w:pPr>` +
        (fmt === 'bullet'
          ? '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:hint="default"/></w:rPr>'
          : '') +
        '</w:lvl>'
      );
    }).join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:numbering xmlns:w="${W_NS}">` +
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>' +
    levels('bullet') +
    '</w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/>' +
    levels('decimal') +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
    '</w:numbering>'
  );
}

function contentTypesXml(
  images: Map<string, ImageEntry>,
  comments: boolean,
  footnotes: boolean
): string {
  const defaults = new Map<string, string>([
    ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
    ['xml', 'application/xml'],
  ]);
  for (const image of images.values()) defaults.set(image.ext, image.contentType);
  const body = [...defaults.entries()].map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Types xmlns="${CT_NS}">${body}` +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    (comments
      ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
      : '') +
    (footnotes
      ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      : '') +
    '</Types>'
  );
}

function rootRelsXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Relationships xmlns="${REL_NS}">` +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>'
  );
}

function documentRelsXml(
  images: Map<string, ImageEntry>,
  comments: boolean,
  footnotes: boolean
): string {
  const rels = [...images.values()]
    .map(
      (image) =>
        `<Relationship Id="${image.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${image.id}.${image.ext}"/>`
    )
    .join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Relationships xmlns="${REL_NS}">${rels}` +
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
    (comments
      ? `<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>`
      : '') +
    (footnotes
      ? `<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>`
      : '') +
    '</Relationships>'
  );
}

/* -------------------------------------------------------------- утилиты */

/** 1 мм = 56.6929 twips (1/1440 дюйма). */
const TWIPS_PER_MM = 56.692913;

function mmToTwips(mm: number): number {
  return Math.round(mm * TWIPS_PER_MM);
}

function twipsToMm(twips: number): number {
  return Math.round((twips / TWIPS_PER_MM) * 10) / 10;
}

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function localName(node: Element): string {
  return node.localName || node.nodeName.replace(/^.*:/, '');
}

function childElements(node: Element): Element[] {
  return [...node.childNodes].filter((child): child is Element => child.nodeType === Node.ELEMENT_NODE);
}

function elementsByLocal(root: ParentNode, name: string): Element[] {
  const scope = root instanceof Document ? root.documentElement : (root as Element);
  return [...scope.getElementsByTagName('*')].filter((el) => localName(el) === name);
}

function firstByLocal(root: ParentNode | null, name: string): Element | null {
  if (!root) return null;
  return elementsByLocal(root, name)[0] ?? null;
}

function attr(el: Element | null, name: string): string {
  if (!el) return '';
  return el.getAttribute(`w:${name}`) || el.getAttributeNS(W_NS, name) || el.getAttribute(name) || '';
}

function isOn(el: Element | null): boolean {
  if (!el) return false;
  const value = attr(el, 'val');
  return value !== 'false' && value !== '0' && value !== 'none';
}

function cssAlign(value: string): string {
  if (value === 'center') return 'center';
  if (value === 'right' || value === 'end') return 'right';
  if (value === 'both' || value === 'distribute') return 'justify';
  return 'left';
}

function cssValue(style: string, prop: string): string {
  const match = style.match(new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([^;]+)`, 'i'));
  return match ? match[1].trim() : '';
}

function cssSizeToPt(value: string): number | null {
  const match = value.match(/^([\d.]+)\s*(pt|px|em|rem)?$/i);
  if (!match) return null;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || amount <= 0) return null;
  const unit = (match[2] || 'pt').toLowerCase();
  if (unit === 'px') return Math.round(amount * 0.75 * 2) / 2;
  if (unit === 'em' || unit === 'rem') return Math.round(amount * 15 * 2) / 2;
  return amount;
}

function cssColorToHex(value: string): string {
  const hex = value.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (hex) {
    const raw = hex[1];
    return (raw.length === 3 ? raw.split('').map((c) => c + c).join('') : raw).toUpperCase();
  }
  const rgb = value.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (!rgb) return '';
  return [rgb[1], rgb[2], rgb[3]].map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/**
 * Почти белый фон в Word — просто шум: редактор рисует такие заливки
 * от копипасты, а в документе они дают грязные серые плашки.
 */
function isMeaningfulFill(hex: string): boolean {
  if (!/^[0-9A-F]{6}$/.test(hex)) return false;
  const [r, g, b] = hexToRgb(hex);
  return (r * 0.299 + g * 0.587 + b * 0.114) < 0.94 * 255;
}

function hexToRgb(hex: string): [number, number, number] {
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

function mimeOf(path: string): string {
  const ext = path.split('.').pop()?.toLowerCase() ?? '';
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
  if (ext === 'gif') return 'image/gif';
  if (ext === 'webp') return 'image/webp';
  if (ext === 'svg') return 'image/svg+xml';
  return 'image/png';
}

function extOf(contentType: string): string {
  if (contentType.includes('jpeg')) return 'jpeg';
  if (contentType.includes('gif')) return 'gif';
  if (contentType.includes('webp')) return 'webp';
  if (contentType.includes('svg')) return 'svg';
  return 'png';
}

function escapeHtml(text: string): string {
  return text.replace(/[&<>]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[char] ?? char);
}

function escapeAttr(text: string): string {
  return escapeHtml(text).replace(/"/g, '&quot;');
}

function bookmarkId(name: string): number {
  let hash = 7;
  for (let index = 0; index < name.length; index += 1) {
    hash = (hash * 31 + name.charCodeAt(index)) % 2147480000;
  }
  return hash;
}

function escapeXml(text: string): string {
  return text.replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char] ?? char);
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}
