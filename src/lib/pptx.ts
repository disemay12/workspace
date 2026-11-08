/**
 * Чтение и правка .pptx без пересборки пакета.
 *
 * Слайды, layout'ы, темы и media остаются исходными; меняются только
 * текстовые тела фигур и список слайдов в presentation.xml.
 */
import JSZip from 'jszip';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const SLIDE_TYPE = `${R_NS}/slide`;
const LAYOUT_TYPE = `${R_NS}/slideLayout`;
const NOTES_TYPE = `${R_NS}/notesSlide`;
const NOTES_MASTER_TYPE = `${R_NS}/notesMaster`;
const IMAGE_TYPE = `${R_NS}/image`;
const SLIDE_CT =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const NOTES_CT =
  'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';

export interface PptxTextShape {
  shapeId: string;
  name: string;
  x: number;
  y: number;
  w: number;
  h: number;
  paragraphs: string[];
  text: string;
  fontSize?: number;
  bold?: boolean;
  color?: string;
  isTitle: boolean;
}

export interface PptxSlide {
  index: number;
  part: string;
  shapes: PptxTextShape[];
  notes: string;
  background: string;
  layoutPart: string;
}

export interface PptxLayout {
  index: number;
  name: string;
  part: string;
}

export interface PptxModel {
  slides: PptxSlide[];
  width: number;
  height: number;
  layouts: PptxLayout[];
}

export interface PptxTextEdit {
  slide: number;
  shapeId: string;
  paragraphs: string[];
}

export interface PptxEditResult {
  ok: boolean;
  message: string;
}

export async function parsePptx(base64: string): Promise<PptxModel> {
  const zip = await loadZip(base64);
  const presentation = await readPresentation(zip);
  const layouts = await readLayouts(zip);
  const slides: PptxSlide[] = [];

  for (let index = 0; index < presentation.slideParts.length; index += 1) {
    const part = presentation.slideParts[index];
    const file = zip.file(part);
    if (!file) continue;
    const doc = parseXml(await file.async('string'));
    const slide = parseSlide(doc, index, part, presentation);
    slide.notes = await readNotesText(zip, part);
    slide.layoutPart = await readSlideLayoutPart(zip, part);
    slides.push(slide);
  }

  return {
    slides,
    width: presentation.width,
    height: presentation.height,
    layouts,
  };
}

export async function patchPptxText(
  base64: string,
  edits: PptxTextEdit[]
): Promise<string> {
  const zip = await loadZip(base64);
  const presentation = await readPresentation(zip);
  const bySlide = new Map<number, PptxTextEdit[]>();
  for (const edit of edits) {
    const list = bySlide.get(edit.slide) ?? [];
    list.push(edit);
    bySlide.set(edit.slide, list);
  }

  for (const [slideIndex, slideEdits] of bySlide) {
    const part = presentation.slideParts[slideIndex];
    if (!part) continue;
    const file = zip.file(part);
    if (!file) continue;
    const doc = parseXml(await file.async('string'));
    for (const edit of slideEdits) {
      const shape = findShape(doc, edit.shapeId);
      if (!shape) continue;
      const txBody = firstByLocal(shape, 'txBody');
      if (!txBody) continue;
      replaceTextBody(doc, txBody, edit.paragraphs);
    }
    zip.file(part, serializeXml(doc));
  }

  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

/** Добавляет слайд после index. title необязателен. */
export async function addPptxSlide(
  base64: string,
  index: number,
  title = ''
): Promise<string> {
  const zip = await loadZip(base64);
  const presentation = await readPresentation(zip);
  const number = nextNumber(zip, 'ppt/slides/slide');
  const part = `ppt/slides/slide${number}.xml`;
  const relsPart = `ppt/slides/_rels/slide${number}.xml.rels`;
  const layoutTarget = await findLayoutTarget(zip, presentation);

  zip.file(part, slideXml(title));
  zip.file(
    relsPart,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${REL_NS}">` +
      `<Relationship Id="rId1" Type="${LAYOUT_TYPE}" Target="${layoutTarget}"/>` +
      '</Relationships>'
  );

  let presentationXml = (await zip.file('ppt/presentation.xml')?.async('string')) ?? '';
  let relsXml =
    (await zip.file('ppt/_rels/presentation.xml.rels')?.async('string')) ?? '';
  const relId = nextRelationshipId(relsXml);
  const sldId = nextSlideId(presentationXml);

  const insertion = `<p:sldId id="${sldId}" r:id="${relId}"/>`;
  presentationXml = insertIntoSlideList(presentationXml, insertion, index);
  relsXml = appendRelationship(relsXml, relId, SLIDE_TYPE, `slides/slide${number}.xml`);
  zip.file('ppt/presentation.xml', presentationXml);
  zip.file('ppt/_rels/presentation.xml.rels', relsXml);

  const contentTypes = await addContentType(zip, `/ppt/slides/slide${number}.xml`, SLIDE_CT);
  if (contentTypes) zip.file('[Content_Types].xml', contentTypes);

  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

export async function deletePptxSlide(base64: string, index: number): Promise<string> {
  const zip = await loadZip(base64);
  const presentation = await readPresentation(zip);
  const part = presentation.slideParts[index];
  if (!part) return base64;

  let presentationXml = (await zip.file('ppt/presentation.xml')?.async('string')) ?? '';
  let relsXml =
    (await zip.file('ppt/_rels/presentation.xml.rels')?.async('string')) ?? '';

  const sldIds = [...presentationXml.matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
  const target = sldIds[index] ?? '';
  const relId = /r:id="([^"]+)"/.exec(target)?.[1] ?? '';
  if (target) presentationXml = presentationXml.replace(target, '');
  if (relId) {
    relsXml = relsXml.replace(
      new RegExp(`<Relationship\\b[^>]*Id="${escapeRegExp(relId)}"[^>]*/>`),
      ''
    );
  }

  zip.remove(part);
  zip.remove(part.replace('slides/', 'slides/_rels/') + '.rels');
  zip.file('ppt/presentation.xml', presentationXml);
  zip.file('ppt/_rels/presentation.xml.rels', relsXml);

  const contentTypes = await removeContentType(zip, `/${part}`);
  if (contentTypes) zip.file('[Content_Types].xml', contentTypes);

  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

export async function movePptxSlide(
  base64: string,
  from: number,
  to: number
): Promise<string> {
  const zip = await loadZip(base64);
  const file = zip.file('ppt/presentation.xml');
  if (!file) return base64;
  let xml = await file.async('string');
  const ids = [...xml.matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
  if (from < 0 || from >= ids.length) return base64;

  const without = ids.filter((_, index) => index !== from);
  const target = Math.max(0, Math.min(to, without.length));
  without.splice(target, 0, ids[from]);

  const listMatch = /<p:sldIdLst\b[^>]*>[\s\S]*?<\/p:sldIdLst>/.exec(xml);
  if (!listMatch) return base64;
  const list = `<p:sldIdLst>${without.join('')}</p:sldIdLst>`;
  xml = xml.replace(listMatch[0], list);
  zip.file('ppt/presentation.xml', xml);
  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

export async function addPptxTextBox(
  base64: string,
  slideIndex: number,
  text: string
): Promise<string> {
  const zip = await loadZip(base64);
  const presentation = await readPresentation(zip);
  const part = presentation.slideParts[slideIndex];
  if (!part) return base64;
  const file = zip.file(part);
  if (!file) return base64;

  const doc = parseXml(await file.async('string'));
  const spTree = firstByLocal(doc, 'spTree');
  if (!spTree) return base64;

  const id = nextShapeId(doc);
  const width = Math.round(presentation.width * 0.8);
  const height = Math.round(presentation.height * 0.2);
  const offX = Math.round(presentation.width * 0.1);
  const offY = Math.round(presentation.height * 0.35);

  const sp = doc.createElementNS(P_NS, 'p:sp');
  const nv = doc.createElementNS(P_NS, 'p:nvSpPr');
  const cNvPr = doc.createElementNS(P_NS, 'p:cNvPr');
  cNvPr.setAttribute('id', String(id));
  cNvPr.setAttribute('name', `TextBox ${id}`);
  const cNvSpPr = doc.createElementNS(P_NS, 'p:cNvSpPr');
  cNvSpPr.setAttribute('txBox', '1');
  const nvPr = doc.createElementNS(P_NS, 'p:nvPr');
  nv.append(cNvPr, cNvSpPr, nvPr);
  sp.appendChild(nv);

  const spPr = doc.createElementNS(P_NS, 'p:spPr');
  const xfrm = doc.createElementNS(A_NS, 'a:xfrm');
  const off = doc.createElementNS(A_NS, 'a:off');
  off.setAttribute('x', String(offX));
  off.setAttribute('y', String(offY));
  const ext = doc.createElementNS(A_NS, 'a:ext');
  ext.setAttribute('cx', String(width));
  ext.setAttribute('cy', String(height));
  xfrm.append(off, ext);
  const geom = doc.createElementNS(A_NS, 'a:prstGeom');
  geom.setAttribute('prst', 'rect');
  geom.appendChild(doc.createElementNS(A_NS, 'a:avLst'));
  spPr.append(xfrm, geom);
  sp.appendChild(spPr);

  const txBody = doc.createElementNS(P_NS, 'p:txBody');
  const bodyPr = doc.createElementNS(A_NS, 'a:bodyPr');
  bodyPr.setAttribute('wrap', 'square');
  txBody.append(bodyPr, doc.createElementNS(A_NS, 'a:lstStyle'));
  for (const line of (text || '').split('\n')) {
    const p = doc.createElementNS(A_NS, 'a:p');
    if (line) {
      const run = doc.createElementNS(A_NS, 'a:r');
      const rPr = doc.createElementNS(A_NS, 'a:rPr');
      rPr.setAttribute('lang', 'ru-RU');
      rPr.setAttribute('sz', '1800');
      const t = doc.createElementNS(A_NS, 'a:t');
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      t.textContent = line;
      run.append(rPr, t);
      p.appendChild(run);
    } else {
      p.appendChild(doc.createElementNS(A_NS, 'a:endParaRPr'));
    }
    txBody.appendChild(p);
  }
  sp.appendChild(txBody);

  spTree.appendChild(sp);
  zip.file(part, serializeXml(doc));
  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

export interface PptxImageOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  alt?: string;
}

/** Вставляет изображение из data URL в слайд, добавляя media и relationship. */
export async function addPptxImage(
  base64: string,
  slideIndex: number,
  dataUrl: string,
  options: PptxImageOptions = {}
): Promise<string> {
  const image = parseImageDataUrl(dataUrl);
  if (!image) throw new Error('Ожидается изображение в формате data:image/...;base64,...');

  const zip = await loadZip(base64);
  const presentation = await readPresentation(zip);
  const part = presentation.slideParts[slideIndex];
  if (!part) throw new Error('Слайд не найден');
  const file = zip.file(part);
  if (!file) throw new Error('Файл слайда не найден');

  const mediaNumber = nextMediaNumber(zip);
  const mediaPart = `ppt/media/image${mediaNumber}.${image.ext}`;
  zip.file(mediaPart, image.bytes);

  const relsPath = part.replace('slides/', 'slides/_rels/') + '.rels';
  let relsXml = (await zip.file(relsPath)?.async('string')) ?? '';
  const relId = nextRelationshipId(relsXml);
  relsXml = appendRelationship(
    relsXml,
    relId,
    IMAGE_TYPE,
    `../media/image${mediaNumber}.${image.ext}`
  );
  zip.file(relsPath, relsXml);

  const doc = parseXml(await file.async('string'));
  const spTree = firstByLocal(doc, 'spTree');
  if (!spTree) throw new Error('На слайде нет дерева фигур');

  const slideWidthPx = presentation.width / 9525;
  const slideHeightPx = presentation.height / 9525;
  let widthPx = Number(options.width) || 0;
  let heightPx = Number(options.height) || 0;
  if (!widthPx && !heightPx) {
    widthPx = Math.round(slideWidthPx * 0.6);
    heightPx = Math.round(widthPx * 0.5625);
  } else if (!heightPx) {
    heightPx = Math.round(widthPx * 0.5625);
  } else if (!widthPx) {
    widthPx = Math.round(heightPx / 0.5625);
  }
  widthPx = Math.max(24, Math.min(Math.round(widthPx), Math.round(slideWidthPx)));
  heightPx = Math.max(24, Math.min(Math.round(heightPx), Math.round(slideHeightPx)));
  const xPx = Number.isFinite(Number(options.x))
    ? Number(options.x)
    : Math.round((slideWidthPx - widthPx) / 2);
  const yPx = Number.isFinite(Number(options.y))
    ? Number(options.y)
    : Math.round((slideHeightPx - heightPx) / 2);

  const id = nextShapeId(doc);
  const pic = doc.createElementNS(P_NS, 'p:pic');
  const nvPicPr = doc.createElementNS(P_NS, 'p:nvPicPr');
  const cNvPr = doc.createElementNS(P_NS, 'p:cNvPr');
  cNvPr.setAttribute('id', String(id));
  cNvPr.setAttribute('name', `Image ${id}`);
  if (options.alt) cNvPr.setAttribute('descr', options.alt);
  const cNvPicPr = doc.createElementNS(P_NS, 'p:cNvPicPr');
  const nvPr = doc.createElementNS(P_NS, 'p:nvPr');
  nvPicPr.append(cNvPr, cNvPicPr, nvPr);

  const blipFill = doc.createElementNS(P_NS, 'p:blipFill');
  const blip = doc.createElementNS(A_NS, 'a:blip');
  blip.setAttributeNS(R_NS, 'r:embed', relId);
  const stretch = doc.createElementNS(A_NS, 'a:stretch');
  stretch.appendChild(doc.createElementNS(A_NS, 'a:fillRect'));
  blipFill.append(blip, stretch);

  const spPr = doc.createElementNS(P_NS, 'p:spPr');
  const xfrm = doc.createElementNS(A_NS, 'a:xfrm');
  const off = doc.createElementNS(A_NS, 'a:off');
  off.setAttribute('x', String(Math.round(xPx * 9525)));
  off.setAttribute('y', String(Math.round(yPx * 9525)));
  const ext = doc.createElementNS(A_NS, 'a:ext');
  ext.setAttribute('cx', String(Math.round(widthPx * 9525)));
  ext.setAttribute('cy', String(Math.round(heightPx * 9525)));
  xfrm.append(off, ext);
  const geom = doc.createElementNS(A_NS, 'a:prstGeom');
  geom.setAttribute('prst', 'rect');
  geom.appendChild(doc.createElementNS(A_NS, 'a:avLst'));
  spPr.append(xfrm, geom);

  pic.append(nvPicPr, blipFill, spPr);
  spTree.appendChild(pic);
  zip.file(part, serializeXml(doc));

  const contentTypes = await addDefaultContentType(zip, image.ext, image.mime);
  if (contentTypes) zip.file('[Content_Types].xml', contentTypes);

  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

function parseImageDataUrl(
  dataUrl: string
): { bytes: Uint8Array; mime: string; ext: string } | null {
  const match = /^data:image\/([a-z0-9.+-]+);base64,([\s\S]+)$/i.exec(dataUrl.trim());
  if (!match) return null;
  const mime = `image/${match[1].toLowerCase()}`;
  const ext =
    match[1].toLowerCase() === 'jpeg'
      ? 'jpeg'
      : match[1].toLowerCase() === 'svg+xml'
        ? 'svg'
        : match[1].toLowerCase();
  const binary = atob(match[2]);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return { bytes, mime, ext };
}

function nextMediaNumber(zip: JSZip): number {
  let max = 0;
  for (const name of Object.keys(zip.files)) {
    const match = /^ppt\/media\/image(\d+)\./i.exec(name);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return max + 1;
}

async function addDefaultContentType(
  zip: JSZip,
  ext: string,
  mime: string
): Promise<string | null> {
  const file = zip.file('[Content_Types].xml');
  if (!file) return null;
  const xml = await file.async('string');
  if (new RegExp(`Extension="${escapeRegExp(ext)}"`, 'i').test(xml)) return xml;
  return xml.replace('</Types>', `<Default Extension="${ext}" ContentType="${mime}"/></Types>`);
}

/** Обновляет заметки докладчика. Новый notesSlide создаётся только при наличии notesMaster. */
export async function setPptxNotes(
  base64: string,
  slideIndex: number,
  text: string
): Promise<string> {
  const zip = await loadZip(base64);
  const presentation = await readPresentation(zip);
  const slidePart = presentation.slideParts[slideIndex];
  if (!slidePart) return base64;
  const paragraphs = text.split('\n');
  const relsPath = slidePart.replace('slides/', 'slides/_rels/') + '.rels';
  let relsXml = (await zip.file(relsPath)?.async('string')) ?? '';
  const existing = notesTarget(relsXml);

  if (existing) {
    const part = normalizeSlideTarget(existing);
    const file = zip.file(part);
    if (!file) return base64;
    const doc = parseXml(await file.async('string'));
    const body =
      elementsByLocal(doc, 'sp').find(
        (sp) => firstByLocal(sp, 'ph')?.getAttribute('type') === 'body'
      ) ?? elementsByLocal(doc, 'sp').find((sp) => Boolean(firstByLocal(sp, 'txBody')));
    const txBody = body ? firstByLocal(body, 'txBody') : firstByLocal(doc, 'txBody');
    if (!txBody) return base64;
    replaceTextBody(doc, txBody, paragraphs);
    zip.file(part, serializeXml(doc));
    return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
  }

  // Без notesMaster новый notesSlide делает пакет невалидным для PowerPoint.
  const master = await findNotesMaster(zip);
  if (!master) {
    throw new Error('У презентации нет notesMaster: сначала добавьте заметки в PowerPoint');
  }

  const number = nextNumber(zip, 'ppt/notesSlides/notesSlide');
  const part = `ppt/notesSlides/notesSlide${number}.xml`;
  const notesRelId = nextRelationshipId(relsXml);
  relsXml = appendRelationship(
    relsXml,
    notesRelId,
    NOTES_TYPE,
    `../notesSlides/notesSlide${number}.xml`
  );
  zip.file(relsPath, relsXml);
  zip.file(part, notesSlideXml(paragraphs));
  zip.file(
    `ppt/notesSlides/_rels/notesSlide${number}.xml.rels`,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${REL_NS}">` +
      `<Relationship Id="rId1" Type="${SLIDE_TYPE}" Target="../slides/${slidePart.split('/').pop()}"/>` +
      `<Relationship Id="rId2" Type="${NOTES_MASTER_TYPE}" Target="../notesMasters/${master}"/>` +
      '</Relationships>'
  );

  const contentTypes = await addContentType(zip, `/${part}`, NOTES_CT);
  if (contentTypes) zip.file('[Content_Types].xml', contentTypes);

  const presentationRelsPath = 'ppt/_rels/presentation.xml.rels';
  let presentationRels = (await zip.file(presentationRelsPath)?.async('string')) ?? '';
  if (!/Type="[^"]*\/notesMaster"/.test(presentationRels)) {
    const relId = nextRelationshipId(presentationRels);
    presentationRels = appendRelationship(
      presentationRels,
      relId,
      NOTES_MASTER_TYPE,
      `notesMasters/${master}`
    );
    zip.file(presentationRelsPath, presentationRels);

    const presentationPath = 'ppt/presentation.xml';
    let presentationXml = (await zip.file(presentationPath)?.async('string')) ?? '';
    if (presentationXml && !/<p:notesMasterIdLst\b/.test(presentationXml)) {
      const entry = `<p:notesMasterIdLst><p:notesMasterId r:id="${relId}"/></p:notesMasterIdLst>`;
      presentationXml = /<p:sldMasterIdLst\b[^>]*>[\s\S]*?<\/p:sldMasterIdLst>/.test(
        presentationXml
      )
        ? presentationXml.replace(
            /(<p:sldMasterIdLst\b[^>]*>[\s\S]*?<\/p:sldMasterIdLst>)/,
            `$1${entry}`
          )
        : presentationXml.replace(/(<p:presentation\b[^>]*>)/, `$1${entry}`);
      zip.file(presentationPath, presentationXml);
    }
  }

  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

function notesTarget(relsXml: string): string {
  for (const rel of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    if (!/Type="[^"]*\/notesSlide"/.test(rel[0])) continue;
    const target = /Target="([^"]+)"/.exec(rel[0])?.[1];
    if (target) return target;
  }
  return '';
}

async function findNotesMaster(zip: JSZip): Promise<string> {
  const part = Object.keys(zip.files).find((name) =>
    /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(name)
  );
  return part ? part.split('/').pop() ?? '' : '';
}

function notesSlideXml(paragraphs: string[]): string {
  const body = paragraphs
    .map((line) =>
      line
        ? `<a:p><a:r><a:rPr lang="ru-RU"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`
        : '<a:p><a:endParaRPr/></a:p>'
    )
    .join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<p:notes xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
    `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>` +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
  );
}

function nextShapeId(doc: Document): number {
  let max = 1;
  for (const node of elementsByLocal(doc, 'cNvPr')) {
    const id = Number(node.getAttribute('id'));
    if (Number.isFinite(id)) max = Math.max(max, id);
  }
  return max + 1;
}

async function readLayouts(zip: JSZip): Promise<PptxLayout[]> {
  const layouts: PptxLayout[] = [];
  const parts = Object.keys(zip.files)
    .filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name))
    .sort((left, right) => localeNumber(left) - localeNumber(right));
  for (const part of parts) {
    const file = zip.file(part);
    if (!file) continue;
    const doc = parseXml(await file.async('string'));
    const name =
      firstByLocal(doc, 'cSld')?.getAttribute('name')?.trim() ||
      part.split('/').pop()?.replace(/\.xml$/, '') ||
      'Layout';
    layouts.push({ index: layouts.length + 1, name, part });
  }
  return layouts;
}

async function readSlideLayoutPart(zip: JSZip, slidePart: string): Promise<string> {
  const relsPath = slidePart.replace('slides/', 'slides/_rels/') + '.rels';
  const relsXml = (await zip.file(relsPath)?.async('string')) ?? '';
  for (const rel of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    if (!/Type="[^"]*\/slideLayout"/.test(rel[0])) continue;
    const target = /Target="([^"]+)"/.exec(rel[0])?.[1];
    if (target) return normalizeSlideTarget(target);
  }
  return '';
}

/** Меняет макет слайда, сохраняя остальные части пакета. */
export async function setPptxLayout(
  base64: string,
  slideIndex: number,
  layoutPart: string
): Promise<string> {
  const zip = await loadZip(base64);
  const presentation = await readPresentation(zip);
  const slidePart = presentation.slideParts[slideIndex];
  if (!slidePart) throw new Error('Слайд не найден');
  if (!zip.file(layoutPart)) throw new Error('Макет не найден в презентации');

  const relsPath = slidePart.replace('slides/', 'slides/_rels/') + '.rels';
  let relsXml = (await zip.file(relsPath)?.async('string')) ?? '';
  const target = `../slideLayouts/${layoutPart.split('/').pop()}`;
  const layoutRel = [...relsXml.matchAll(/<Relationship\b[^>]*\/>/g)].find((match) =>
    /Type="[^"]*\/slideLayout"/.test(match[0])
  );
  if (layoutRel) {
    relsXml = relsXml.replace(
      layoutRel[0],
      layoutRel[0].replace(/Target="[^"]*"/, `Target="${target}"`)
    );
  } else {
    const relId = nextRelationshipId(relsXml);
    relsXml = appendRelationship(relsXml, relId, LAYOUT_TYPE, target);
  }
  zip.file(relsPath, relsXml);
  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

async function readNotesText(zip: JSZip, slidePart: string): Promise<string> {
  const relsPath = slidePart.replace('slides/', 'slides/_rels/') + '.rels';
  const relsXml = (await zip.file(relsPath)?.async('string')) ?? '';
  for (const rel of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    if (!/Type="[^"]*\/notesSlide"/.test(rel[0])) continue;
    const target = /Target="([^"]+)"/.exec(rel[0])?.[1];
    if (!target) continue;
    const part = normalizeSlideTarget(target);
    const file = zip.file(part);
    if (!file) continue;
    const doc = parseXml(await file.async('string'));
    return elementsByLocal(doc, 't')
      .map((node) => node.textContent ?? '')
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();
  }
  return '';
}

interface PresentationInfo {
  slideParts: string[];
  width: number;
  height: number;
}

async function readPresentation(zip: JSZip): Promise<PresentationInfo> {
  const xml = (await zip.file('ppt/presentation.xml')?.async('string')) ?? '';
  const relsXml =
    (await zip.file('ppt/_rels/presentation.xml.rels')?.async('string')) ?? '';
  const width = Number(/<p:sldSz\b[^>]*cx="(\d+)"/.exec(xml)?.[1] ?? 12192000);
  const height = Number(/<p:sldSz\b[^>]*cy="(\d+)"/.exec(xml)?.[1] ?? 6858000);
  const targets = relationshipTargets(relsXml);
  const slideParts: string[] = [];

  for (const match of xml.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"[^>]*\/>/g)) {
    const target = targets.get(match[1]);
    if (target) slideParts.push(normalizeSlideTarget(target));
  }

  if (slideParts.length === 0) {
    const files = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort((a, b) => localeNumber(a) - localeNumber(b));
    slideParts.push(...files);
  }

  return { slideParts, width, height };
}

function parseSlide(
  doc: Document,
  index: number,
  part: string,
  presentation: PresentationInfo
): PptxSlide {
  const shapes: PptxTextShape[] = [];
  for (const sp of elementsByLocal(doc, 'sp')) {
    const body = firstByLocal(sp, 'txBody');
    if (!body) continue;
    const paragraphs = [...body.getElementsByTagNameNS(A_NS, 'p')].map((paragraph) =>
      [...paragraph.getElementsByTagNameNS(A_NS, 't')]
        .map((node) => node.textContent ?? '')
        .join('')
    );
    const text = paragraphs.join('\n').trim();
    if (!text) continue;

    const cNvPr = firstByLocal(sp, 'cNvPr');
    const shapeId = cNvPr?.getAttribute('id') ?? '';
    const name = cNvPr?.getAttribute('name') ?? '';
    const xfrm = firstByLocal(sp, 'xfrm');
    const off = xfrm ? firstByLocal(xfrm, 'off') : null;
    const ext = xfrm ? firstByLocal(xfrm, 'ext') : null;
    const x = Number(off?.getAttribute('x') ?? 0);
    const y = Number(off?.getAttribute('y') ?? 0);
    const w = Number(ext?.getAttribute('cx') ?? presentation.width / 2);
    const h = Number(ext?.getAttribute('cy') ?? presentation.height / 4);

    const firstRun = firstByLocal(body, 'rPr');
    const fontSize = Number(firstRun?.getAttribute('sz') ?? 0) / 100 || undefined;
    const ph = firstByLocal(sp, 'ph');
    const placeholder = ph?.getAttribute('type') ?? '';
    const isTitle = placeholder === 'title' || placeholder === 'ctrTitle' || y < presentation.height * 0.2;

    shapes.push({
      shapeId,
      name,
      x,
      y,
      w,
      h,
      paragraphs,
      text,
      fontSize,
      bold: firstRun?.getAttribute('b') === '1',
      color: firstRun ? fillColor(firstRun) : undefined,
      isTitle,
    });
  }

  // Фон: p:bg > p:bgPr > a:solidFill
  let background = '#ffffff';
  const bg = firstByLocal(doc, 'bg');
  if (bg) {
    const color = bg.getElementsByTagNameNS(A_NS, 'srgbClr')[0]?.getAttribute('val');
    if (color) background = `#${color}`;
  }

  return {
    index,
    part,
    shapes,
    notes: '',
    background,
    layoutPart: '',
  };
}

function fillColor(node: Element): string | undefined {
  const srgb = node.getElementsByTagNameNS(A_NS, 'srgbClr')[0]?.getAttribute('val');
  return srgb ? `#${srgb}` : undefined;
}

function findShape(doc: Document, shapeId: string): Element | null {
  return (
    elementsByLocal(doc, 'sp').find(
      (shape) => firstByLocal(shape, 'cNvPr')?.getAttribute('id') === shapeId
    ) ?? null
  );
}

function replaceTextBody(doc: Document, txBody: Element, paragraphs: string[]): void {
  const firstRunPr = firstByLocal(txBody, 'rPr')?.cloneNode(true) as Element | null;
  const pPr = firstByLocal(txBody, 'pPr')?.cloneNode(true) as Element | null;

  for (const child of [...txBody.children]) {
    if (child.localName === 'p') txBody.removeChild(child);
  }

  for (const paragraph of paragraphs) {
    const p = doc.createElementNS(A_NS, 'a:p');
    if (pPr) p.appendChild(pPr.cloneNode(true));
    if (paragraph) {
      const run = doc.createElementNS(A_NS, 'a:r');
      if (firstRunPr) run.appendChild(firstRunPr.cloneNode(true));
      const text = doc.createElementNS(A_NS, 'a:t');
      text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      text.textContent = paragraph;
      run.appendChild(text);
      p.appendChild(run);
    } else {
      p.appendChild(doc.createElementNS(A_NS, 'a:endParaRPr'));
    }
    txBody.appendChild(p);
  }
}

function slideXml(title: string): string {
  const titleShape = title
    ? '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="609600" y="365760"/><a:ext cx="10972800" cy="1371600"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ru-RU" sz="3200"/><a:t>' +
      escapeXml(title) +
      '</a:t></a:r></a:p></p:txBody></p:sp>'
    : '';
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    titleShape +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
  );
}

async function findLayoutTarget(zip: JSZip, presentation: PresentationInfo): Promise<string> {
  for (const part of presentation.slideParts) {
    const relsPath = part.replace('slides/', 'slides/_rels/') + '.rels';
    const relsXml = (await zip.file(relsPath)?.async('string')) ?? '';
    for (const rel of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
      const tag = rel[0];
      if (!/Type="[^"]*\/slideLayout"/.test(tag)) continue;
      const target = /Target="([^"]+)"/.exec(tag)?.[1];
      if (target) return target;
    }
  }
  return '../slideLayouts/slideLayout1.xml';
}

function insertIntoSlideList(xml: string, entry: string, index: number): string {
  const list = /<p:sldIdLst\b[^>]*>([\s\S]*?)<\/p:sldIdLst>/.exec(xml);
  if (!list) {
    return xml.replace(
      /<p:presentation\b[^>]*>/,
      (open) => `${open}<p:sldIdLst>${entry}</p:sldIdLst>`
    );
  }
  const ids = [...list[1].matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
  const target = Math.max(0, Math.min(index, ids.length));
  ids.splice(target, 0, entry);
  return xml.replace(list[0], `<p:sldIdLst>${ids.join('')}</p:sldIdLst>`);
}

function nextNumber(zip: JSZip, prefix: string): number {
  let max = 0;
  for (const name of Object.keys(zip.files)) {
    const match = new RegExp(`^${escapeRegExp(prefix)}(\\d+)\\.xml$`).exec(name);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return max + 1;
}

function nextSlideId(presentationXml: string): number {
  let max = 255;
  for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*id="(\d+)"/g)) {
    max = Math.max(max, Number(match[1]));
  }
  return max + 1;
}

function nextRelationshipId(relsXml: string): string {
  const used = new Set([...relsXml.matchAll(/Id="([^"]+)"/g)].map((match) => match[1]));
  let index = 1;
  while (used.has(`rId${index}`)) index += 1;
  return `rId${index}`;
}

function relationshipTargets(relsXml: string): Map<string, string> {
  const targets = new Map<string, string>();
  for (const rel of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = /Id="([^"]+)"/.exec(rel[0])?.[1];
    const target = /Target="([^"]+)"/.exec(rel[0])?.[1];
    if (id && target) targets.set(id, target);
  }
  return targets;
}

function appendRelationship(
  relsXml: string,
  id: string,
  type: string,
  target: string
): string {
  const entry = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  if (!relsXml) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${REL_NS}">${entry}</Relationships>`
    );
  }
  return relsXml.replace('</Relationships>', `${entry}</Relationships>`);
}

async function addContentType(zip: JSZip, part: string, type: string): Promise<string | null> {
  const file = zip.file('[Content_Types].xml');
  if (!file) return null;
  const xml = await file.async('string');
  if (xml.includes(`PartName="${part}"`)) return xml;
  return xml.replace('</Types>', `<Override PartName="${part}" ContentType="${type}"/></Types>`);
}

async function removeContentType(zip: JSZip, part: string): Promise<string | null> {
  const file = zip.file('[Content_Types].xml');
  if (!file) return null;
  const xml = await file.async('string');
  return xml.replace(
    new RegExp(`<Override\\b[^>]*PartName="${escapeRegExp(part)}"[^>]*/>`),
    ''
  );
}

function normalizeSlideTarget(target: string): string {
  return target.startsWith('/')
    ? target.slice(1)
    : `ppt/${target.replace(/^\.\.\//, '')}`;
}

function localeNumber(value: string): number {
  return Number(/(\d+)/.exec(value)?.[1] ?? 0);
}

async function loadZip(base64: string): Promise<JSZip> {
  return JSZip.loadAsync(base64ToBytes(base64));
}

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc);
}

function elementsByLocal(root: ParentNode, name: string): Element[] {
  const scope = root instanceof Document ? root.documentElement : (root as Element);
  return [...scope.getElementsByTagName('*')].filter((node) => node.localName === name);
}

function firstByLocal(root: ParentNode | null, name: string): Element | null {
  if (!root) return null;
  const scope = root instanceof Document ? root.documentElement : (root as Element);
  return [...scope.getElementsByTagName('*')].find((node) => node.localName === name) ?? null;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeXml(value: string): string {
  return value.replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char] ?? char);
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}
