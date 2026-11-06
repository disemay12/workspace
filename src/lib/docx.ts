import JSZip from 'jszip';

/**
 * Преобразование между HTML-полотном редактора Word и пакетом .docx.
 *
 * mammoth при чтении отбрасывает шрифт, кегль и цвет — здесь они поднимаются
 * из word/document.xml в inline-стили span, а при сохранении HTML собирается
 * обратно в настоящий пакет OOXML, чтобы правки не затирали файл.
 */

export interface DocxHtml {
  html: string;
  /** Предупреждения разбора; документ при этом всё равно открывается. */
  warnings: string[];
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

  const [docXml, relsXml, numberingXml] = await Promise.all([
    docFile.async('string'),
    zip.file('word/_rels/document.xml.rels')?.async('string') ?? '',
    zip.file('word/numbering.xml')?.async('string') ?? '',
  ]);

  const doc = parseXml(docXml);
  const rels = relsXml ? parseXml(relsXml) : null;
  const numbering = numberingXml ? parseXml(numberingXml) : null;

  const images = await readImages(zip, rels);
  const listKinds = readListKinds(numbering);

  const body = firstByLocal(doc, 'body');
  if (!body) return { html: '<p></p>', warnings: ['Документ пуст'] };

  const blocks: string[] = [];
  let list: { numId: string; kind: 'ul' | 'ol' } | null = null;

  const closeList = () => {
    if (list) {
      blocks.push(list.kind === 'ul' ? '</ul>' : '</ol>');
      list = null;
    }
  };

  for (const child of childElements(body)) {
    const name = localName(child);
    if (name === 'p') {
      const numId = paragraphNumbering(child);
      if (numId) {
        const kind = listKinds.get(numId) ?? 'ol';
        if (!list || list.numId !== numId) {
          closeList();
          blocks.push(kind === 'ul' ? '<ul>' : '<ol>');
          list = { numId, kind };
        }
        blocks.push('<li' + listItemStyle(child) + '>' + (await paragraphInner(child, images)) + '</li>');
      } else {
        closeList();
        blocks.push(renderParagraph(child, await paragraphInner(child, images)));
      }
    } else if (name === 'tbl') {
      closeList();
      blocks.push(await renderTable(child, images));
    }
  }
  closeList();

  const html = blocks.join('') || '<p></p>';
  return { html, warnings: [] };
}

/** Собирает .docx из HTML полотна и возвращает его как base64. */
export async function htmlToDocxBase64(html: string): Promise<string> {
  const container = document.createElement('div');
  container.innerHTML = html && html.trim() ? html : '<p></p>';

  const images = new Map<string, { id: string; bytes: Uint8Array; contentType: string; ext: string }>();
  let imageCounter = 0;

  const bodyXml = renderChildren(container, images, () => {
    imageCounter += 1;
    return 'rId' + (imageCounter + 1);
  });

  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypesXml(images));
  zip.folder('_rels')!.file('.rels', rootRelsXml());
  const word = zip.folder('word')!;
  word.file('document.xml', documentXml(bodyXml || '<w:p/>'));
  word.file('styles.xml', stylesXml());
  word.file('numbering.xml', numberingXml());
  word.folder('_rels')!.file('document.xml.rels', documentRelsXml(images));

  for (const image of images.values()) {
    word.file(`media/${image.id}.${image.ext}`, image.bytes);
  }

  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

/* ---------------------------------------------------------------- чтение */

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

function paragraphNumbering(p: Element): string | null {
  const pPr = firstByLocal(p, 'pPr');
  const numPr = pPr ? firstByLocal(pPr, 'numPr') : null;
  return numPr ? attr(firstByLocal(numPr, 'numId'), 'val') : null;
}

/** Оформление элемента списка: выравнивание и отступы. */
function listItemStyle(p: Element): string {
  const pPr = firstByLocal(p, 'pPr');
  const align = pPr ? attr(firstByLocal(pPr, 'jc'), 'val') : '';
  const css = [align ? `text-align:${cssAlign(align)}` : '', paragraphIndent(pPr)]
    .filter(Boolean)
    .join(';');
  return css ? ` style="${css}"` : '';
}

function renderParagraph(p: Element, inner: string): string {
  const pPr = firstByLocal(p, 'pPr');
  const style = pPr ? attr(firstByLocal(pPr, 'pStyle'), 'val') : '';
  const align = pPr ? attr(firstByLocal(pPr, 'jc'), 'val') : '';
  const css = [align ? `text-align:${cssAlign(align)}` : '', paragraphIndent(pPr)]
    .filter(Boolean)
    .join(';');
  const styleAttr = css ? ` style="${css}"` : '';
  const tag = headingTag(style);
  return `<${tag}${styleAttr}>${inner || '<br>'}</${tag}>`;
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
    }
  }
  return html;
}

function renderRun(run: Element, images: Map<string, string>): string {
  const rPr = firstByLocal(run, 'rPr');
  let text = '';
  for (const child of childElements(run)) {
    const name = localName(child);
    if (name === 't') text += escapeHtml(child.textContent ?? '');
    else if (name === 'tab') text += '&emsp;';
    else if (name === 'br' || name === 'cr') text += '<br>';
    else if (name === 'drawing' || name === 'pict') {
      const blip = firstByLocal(child, 'blip');
      const id = blip?.getAttribute('r:embed') ?? blip?.getAttributeNS(R_NS, 'embed') ?? '';
      const src = images.get(id);
      if (src) text += `<img src="${src}" alt="">`;
    }
  }
  if (!text) return '';
  return wrapRun(text, rPr);
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
      list = tag;
      for (const li of [...node.children]) {
        if (li.tagName.toLowerCase() !== 'li') continue;
        xml += paragraphXml(runsFromInline(li, {}, images, nextId), {
          list: tag,
          align: readAlign(li),
          ...readIndent(li),
        });
      }
      closeList();
      continue;
    }

    if (tag === 'table') {
      closeList();
      xml += tableXml(node, images, nextId);
      continue;
    }

    if (tag === 'br') {
      closeList();
      xml += paragraphXml('', {});
      continue;
    }

    closeList();
    const align = readAlign(node);
    const style = blockStyle(tag);
    xml += paragraphXml(runsFromInline(node, {}, images, nextId), {
      align,
      style,
      ...readIndent(node),
    });
  }
  return xml;
}

interface ParagraphOptions {
  align?: string;
  style?: string;
  list?: 'ul' | 'ol';
  /** Абзац внутри таблицы: без отступа Normal, иначе ячейки раздуваются. */
  tight?: boolean;
  indentLeftPt?: number;
  firstLinePt?: number;
}

function paragraphXml(runs: string, options: ParagraphOptions): string {
  const props: string[] = [];
  if (options.style) props.push(`<w:pStyle w:val="${options.style}"/>`);
  if (options.tight) {
    // В ячейках таблицы абзацный отступ Normal раздувает строки.
    props.push('<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>');
  }
  if (options.list) {
    props.push(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${options.list === 'ul' ? '1' : '2'}"/></w:numPr>`);
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
  const pPr = props.length ? `<w:pPr>${props.join('')}</w:pPr>` : '';
  return `<w:p>${pPr}${runs || '<w:r><w:t></w:t></w:r>'}</w:p>`;
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
  if (tag === 'br') return '<w:r><w:br/></w:r>';

  if (tag === 'img') {
    const src = node.getAttribute('src') ?? '';
    const image = takeImage(src, images, nextId);
    if (!image) return '';
    return imageRunXml(image.id);
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

function imageRunXml(id: string): string {
  const drawingId = nextDrawingId++;
  return (
    '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="3200400"/>' +
    `<wp:docPr id="${drawingId}" name="Image ${drawingId}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    // nvPicPr обязателен: без него Word считает пакет повреждённым.
    `<pic:nvPicPr><pic:cNvPr id="${drawingId}" name="Image ${drawingId}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="${R_NS}" r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="3200400"/></a:xfrm>' +
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
      // Заливка шапки и небольшие отступы: без них таблица в Word выглядит
      // голой сеткой, а не как в редакторе.
      const props =
        '<w:tcPr><w:tcW w:w="0" w:type="auto"/>' +
        (isHeader ? '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' : '') +
        '</w:tcPr>';
      cells.push(`<w:tc>${props}${paragraphXml(runs, { tight: true })}</w:tc>`);
    }
    rows.push(`<w:tr>${cells.join('')}</w:tr>`);
  }
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>' +
    '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>' +
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/></w:tblBorders>' +
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

function documentXml(body: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}">` +
    `<w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>'
  );
}

function stylesXml(): string {
  const heading = (id: string, size: number, before: number, after: number, outline: number) =>
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id.replace(/(\d)/, ' $1')}"/>` +
    `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="${before}" w:after="${after}"/><w:outlineLvl w:val="${outline}"/></w:pPr>` +
    `<w:rPr><w:b/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:styles xmlns:w="${W_NS}">` +
    // Базовый кегль 12 pt — так документ совпадает с тем, что показывает редактор.
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>' +
    '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' +
    heading('Heading1', 44, 240, 120, 0) +
    heading('Heading2', 32, 240, 100, 1) +
    heading('Heading3', 26, 200, 80, 2) +
    heading('Heading4', 24, 160, 80, 3) +
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:ind w:left="360"/><w:spacing w:after="160"/></w:pPr>' +
    '<w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>' +
    '</w:styles>'
  );
}

/**
 * Описание списков: numId 1 — маркированный, 2 — нумерованный.
 * Без numbering.xml Word и собственный разбор считают любые списки
 * нумерованными, и маркеры «•» превращаются в «1. 2. 3.».
 */
function numberingXml(): string {
  // Маркер списка — обычный U+2022 в шрифте документа. С Symbol этот символ
  // не отображается: Word рисует «□» (нет глифа), хотя в редакторе точка.
  const level = (fmt: string, text: string) =>
    '<w:lvl w:ilvl="0"><w:start w:val="1"/>' +
    `<w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
    '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
    (fmt === 'bullet'
      ? '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:hint="default"/></w:rPr>'
      : '') +
    '</w:lvl>';
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:numbering xmlns:w="${W_NS}">` +
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>' +
    level('bullet', '\u2022') +
    '</w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>' +
    level('decimal', '%1.') +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
    '</w:numbering>'
  );
}

function contentTypesXml(images: Map<string, ImageEntry>): string {
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

function documentRelsXml(images: Map<string, ImageEntry>): string {
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
    '</Relationships>'
  );
}

/* -------------------------------------------------------------- утилиты */

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

function firstByLocal(root: ParentNode, name: string): Element | null {
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
