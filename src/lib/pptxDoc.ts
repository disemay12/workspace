/**
 * Документ презентации: разбор пакета .pptx в редактируемую модель и операции
 * над частями OOXML.
 *
 * Ключевое отличие от набора отдельных функций в `pptx.ts`: пакет читается
 * один раз, XML-части держатся в памяти как строки, а операция меняет только
 * затронутые части. Сериализация выполняется лишь при сохранении, поэтому
 * правка текста на 100-слайдовой колоде не пересобирает ZIP.
 *
 * Темизация: `a:clrScheme` и `a:fontScheme` в `ppt/theme/themeN.xml` — ровно то,
 * что меняет PowerPoint при смене темы. Элементы с явным `srgbClr` остаются
 * как есть (так же ведёт себя PowerPoint).
 */
import JSZip from 'jszip';
import {
  normalizeHexColor,
  presetById,
  fontPresetById,
  type ThemeColorSet,
  type ThemeFontPair,
} from './pptxThemes';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

const REL_TYPE = {
  slide: `${R_NS}/slide`,
  slideLayout: `${R_NS}/slideLayout`,
  notesSlide: `${R_NS}/notesSlide`,
  notesMaster: `${R_NS}/notesMaster`,
  theme: `${R_NS}/theme`,
  image: `${R_NS}/image`,
  hyperlink: `${R_NS}/hyperlink`,
};

const CONTENT_TYPE = {
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  layout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  master: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  notesMaster: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
};

/* ============================================================
   Модель
   ============================================================ */

export type PptxShapeKind = 'text' | 'image' | 'table' | 'preset' | 'placeholder' | 'group';

export interface PptxTextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Размер в пунктах. */
  size?: number;
  /** Разрешённый hex (учитывает цвет схемы и его оттенок). */
  color?: string;
  /** Имя цвета схемы, если цвет взят из темы. */
  schemeColor?: string;
  family?: string;
  highlight?: string;
  hyperlink?: string;
  lang?: string;
}

export interface PptxParagraph {
  runs: PptxTextRun[];
  align?: 'l' | 'ctr' | 'r' | 'just';
  level?: number;
  bullet?: 'none' | 'bullet' | 'number';
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
}

export interface PptxTextBody {
  paragraphs: PptxParagraph[];
  verticalAnchor?: 't' | 'ctr' | 'b';
  autofit?: 'none' | 'spAutoFit' | 'normAutofit';
}

export interface PptxTableCell {
  text: PptxTextBody;
  fill?: string;
}

export interface PptxTableModel {
  rows: PptxTableCell[][];
  colWidths: number[];
  rowHeights: number[];
}

export interface PptxShapeModel {
  id: string;
  name: string;
  kind: PptxShapeKind;
  /** Геометрия в EMU. */
  x: number;
  y: number;
  w: number;
  h: number;
  /** Поворот в градусах (1/60000 долей градуса в OOXML уже пересчитаны). */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  geom?: string;
  text?: PptxTextBody;
  image?: { src: string; part?: string };
  table?: PptxTableModel;
  fill?: string | null;
  outline?: { color?: string; width?: number };
  shadow?: boolean;
  isTitle?: boolean;
  placeholderType?: string;
}

export type PptxTransitionKind = 'none' | 'fade' | 'push' | 'wipe' | 'cut';

export interface PptxTransitionModel {
  type: PptxTransitionKind;
  direction?: string;
  speed?: 'slow' | 'med' | 'fast';
  advanceOnClick?: boolean;
  advanceAfterMs?: number;
}

export interface PptxSlideModel {
  index: number;
  part: string;
  layoutPart?: string;
  layoutName?: string;
  hidden?: boolean;
  background: string;
  backgroundImage?: string;
  shapes: PptxShapeModel[];
  notes: string;
  transition?: PptxTransitionModel;
}

export interface PptxLayoutInfo {
  index: number;
  name: string;
  part: string;
}

export interface PptxThemeModel {
  part: string;
  name: string;
  colors: ThemeColorSet;
  fonts: ThemeFontPair;
}

export interface PptxDocModel {
  width: number;
  height: number;
  theme: PptxThemeModel;
  layouts: PptxLayoutInfo[];
  slides: PptxSlideModel[];
}

/** Снимок для отмены: изменённые части (null — части не было). */
interface UndoEntry {
  label: string;
  parts: Map<string, string | null>;
}

/* ============================================================
   Утилиты XML
   ============================================================ */

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc);
}

function childByLocal(parent: Element | Document | null, name: string): Element | null {
  if (!parent) return null;
  const scope = parent instanceof Document ? parent.documentElement : parent;
  if (!scope) return null;
  for (const child of [...scope.children]) {
    if (child.localName === name) return child;
  }
  return null;
}

function childrenByLocal(parent: Element | Document | null, name: string): Element[] {
  if (!parent) return [];
  const scope = parent instanceof Document ? parent.documentElement : parent;
  if (!scope) return [];
  return [...scope.children].filter((child) => child.localName === name);
}

function allByLocal(root: Element | Document | null, name: string): Element[] {
  if (!root) return [];
  const scope = root instanceof Document ? root.documentElement : root;
  if (!scope) return [];
  return [...scope.getElementsByTagName('*')].filter((node) => node.localName === name);
}

function firstByLocal(root: Element | Document | null, name: string): Element | null {
  return allByLocal(root, name)[0] ?? null;
}

function attr(element: Element | null | undefined, name: string): string {
  return element?.getAttribute(name) ?? '';
}

function numAttr(element: Element | null | undefined, name: string, fallback = 0): number {
  // Без элемента или атрибута возвращаем именно fallback: Number('') === 0
  // раньше подменял отсутствующий lumMod/lumOff нулём, и любой цвет
  // разрешался в чёрный (applyLuminance с нулевой яркостью).
  const raw = element?.getAttribute(name);
  if (raw === null || raw === undefined || raw === '') return fallback;
  const value = Number(raw);
  return Number.isFinite(value) ? value : fallback;
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

function bytesToBase64(bytes: Uint8Array): string {
  let binary = '';
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
  }
  return btoa(binary);
}

/** HSL-подобное преобразование оттенка цвета схемы (lumMod/lumOff). */
function applyLuminance(hex: string, lumMod?: number, lumOff?: number): string {
  const normalized = normalizeHexColor(hex, '000000');
  const channels = [0, 2, 4].map((offset) => parseInt(normalized.slice(offset, offset + 2), 16) / 255);
  const max = Math.max(...channels);
  const min = Math.min(...channels);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;
  const d = max - min;
  if (d !== 0) {
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === channels[0]) h = ((channels[1] - channels[2]) / d + (channels[1] < channels[2] ? 6 : 0)) / 6;
    else if (max === channels[1]) h = ((channels[2] - channels[0]) / d + 2) / 6;
    else h = ((channels[0] - channels[1]) / d + 4) / 6;
  }
  const mod = lumMod === undefined ? 1 : lumMod / 100000;
  const off = lumOff === undefined ? 0 : lumOff / 100000;
  const nextL = Math.min(1, Math.max(0, l * mod + off));
  const hue2rgb = (p: number, q: number, t: number) => {
    let value = t;
    if (value < 0) value += 1;
    if (value > 1) value -= 1;
    if (value < 1 / 6) return p + (q - p) * 6 * value;
    if (value < 1 / 2) return q;
    if (value < 2 / 3) return p + (q - p) * (2 / 3 - value) * 6;
    return p;
  };
  let rgb: number[];
  if (s === 0) {
    rgb = [nextL, nextL, nextL];
  } else {
    const q = nextL < 0.5 ? nextL * (1 + s) : nextL + s - nextL * s;
    const p = 2 * nextL - q;
    rgb = [hue2rgb(p, q, h + 1 / 3), hue2rgb(p, q, h), hue2rgb(p, q, h - 1 / 3)];
  }
  return rgb
    .map((value) => Math.round(Math.min(1, Math.max(0, value)) * 255).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

/* ============================================================
   Документ
   ============================================================ */

export class PptxDocument {
  private zip: JSZip;
  /** XML-части как строки: операции над ними синхронные. */
  private parts = new Map<string, string>();
  private binaryParts = new Map<string, Uint8Array>();
  private mediaCache = new Map<string, string>();
  private undoStack: UndoEntry[] = [];
  private redoStack: UndoEntry[] = [];
  /** Части, удалённые из пакета: при сохранении они должны исчезнуть и из ZIP. */
  private deletedParts = new Set<string>();
  private doc: PptxDocModel;
  /** Части медиа, которые ещё не прочитаны как data URL. */
  private pendingMedia: { part: string }[] = [];
  /** Фон-картинка по части слайда — чтобы обновить модель после загрузки медиа. */
  private backgroundMedia = new Map<string, string>();

  private constructor(zip: JSZip) {
    this.zip = zip;
    this.doc = emptyDoc();
  }

  /** Открывает пакет из base64: XML читаются сразу, медиа — по требованию. */
  static async open(base64: string): Promise<PptxDocument> {
    const zip = await JSZip.loadAsync(base64ToBytes(base64));
    return PptxDocument.fromZip(zip);
  }

  static async openBytes(bytes: Uint8Array): Promise<PptxDocument> {
    const zip = await JSZip.loadAsync(bytes);
    return PptxDocument.fromZip(zip);
  }

  private static async fromZip(zip: JSZip): Promise<PptxDocument> {
    const instance = new PptxDocument(zip);
    await instance.preloadXmlParts();
    instance.doc = instance.parseModel();
    return instance;
  }

  get model(): PptxDocModel {
    return this.doc;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  get undoLabel(): string {
    return this.undoStack[this.undoStack.length - 1]?.label ?? '';
  }

  /* ---------- Части пакета ---------- */

  private async preloadXmlParts(): Promise<void> {
    const names = Object.keys(this.zip.files).filter(
      (name) => !this.zip.files[name].dir && /\.(xml|rels)$/i.test(name)
    );
    for (const name of names) {
      const file = this.zip.file(name);
      if (!file) continue;
      this.parts.set(name, await file.async('string'));
    }
  }

  private slidePartsOrder(): string[] {
    const presentation = this.parts.get('ppt/presentation.xml');
    const rels = this.parts.get('ppt/_rels/presentation.xml.rels') ?? '';
    if (!presentation) return [];
    const list = /<p:sldIdLst\b[^>]*>([\s\S]*?)<\/p:sldIdLst>/.exec(presentation);
    const entries = list ? [...(list[1] ?? '').matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"[^>]*\/>/g)] : [];
    const targets = new Map<string, string>();
    for (const rel of rels.matchAll(/<Relationship\b[^>]*\/>/g)) {
      const tag = rel[0];
      const id = /Id="([^"]+)"/.exec(tag)?.[1];
      const target = /Target="([^"]+)"/.exec(tag)?.[1];
      if (id && target) targets.set(id, target.replace(/^\.?\//, ''));
    }
    return entries
      .map((entry) => targets.get(entry[1] ?? '') ?? '')
      .filter(Boolean)
      .map((target) => (target.startsWith('ppt/') ? target : 'ppt/' + target));
  }

  private layoutParts(): string[] {
    return Object.keys(this.zip.files)
      .filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name))
      .sort((a, b) => Number(a.replace(/\D/g, '')) - Number(b.replace(/\D/g, '')));
  }

  private themePart(): string {
    // Тема принадлежит мастеру: ищем по связи, иначе стандартный путь.
    for (const name of Object.keys(this.zip.files)) {
      if (!/^ppt\/slideMasters\/_rels\/slideMaster\d+\.xml\.rels$/.test(name)) continue;
      const rels = this.parts.get(name) ?? '';
      const theme = /Target="([^"]*theme[^"]*)"/.exec(rels)?.[1];
      if (theme) return theme.replace(/^\.\.\//, 'ppt/').replace(/^\.?\//, '');
    }
    return 'ppt/theme/theme1.xml';
  }

  /* ---------- Разбор ---------- */

  private parseModel(): PptxDocModel {
    const presentationXml = this.parts.get('ppt/presentation.xml') ?? '';
    const sizeMatch = /<p:sldSz\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(presentationXml);
    const width = sizeMatch ? Number(sizeMatch[1]) : 12192000;
    const height = sizeMatch ? Number(sizeMatch[2]) : 6858000;

    const theme = this.parseTheme();
    const layouts: PptxLayoutInfo[] = this.layoutParts().map((part, index) => {
      const xml = this.parts.get(part) ?? '';
      const name = /<p:cSld\b[^>]*name="([^"]*)"/.exec(xml)?.[1] ?? '';
      return { index, part, name: name || `Макет ${index + 1}` };
    });

    const slides: PptxSlideModel[] = this.slidePartsOrder().map((part, index) =>
      this.parseSlide(part, index, theme)
    );

    return { width, height, theme, layouts, slides };
  }

  private parseTheme(): PptxThemeModel {
    const part = this.themePart();
    const xml = this.parts.get(part) ?? '';
    const doc = parseXml(xml || '<a:theme xmlns:a="' + A_NS + '"/>');
    const clrScheme = firstByLocal(doc, 'clrScheme');
    const name = attr(clrScheme, 'name') || 'Office';
    const fontScheme = firstByLocal(doc, 'fontScheme');

    const readColor = (scheme: Element | null, slot: string, fallback: string): string => {
      const node = childByLocal(scheme, slot);
      if (!node) return fallback;
      const srgb = childByLocal(node, 'srgbClr');
      const sys = childByLocal(node, 'sysClr');
      const value = attr(srgb, 'val') || attr(sys, 'lastClr');
      return value ? normalizeHexColor(value, fallback) : fallback;
    };

    const defaults: ThemeColorSet = {
      dk1: '000000',
      lt1: 'FFFFFF',
      dk2: '44546A',
      lt2: 'E7E6E6',
      accent1: '4472C4',
      accent2: 'ED7D31',
      accent3: 'A5A5A5',
      accent4: 'FFC000',
      accent5: '5B9BD5',
      accent6: '70AD47',
      hlink: '0563C1',
      folHlink: '954F72',
    };

    const colors: ThemeColorSet = {
      dk1: readColor(clrScheme, 'dk1', defaults.dk1),
      lt1: readColor(clrScheme, 'lt1', defaults.lt1),
      dk2: readColor(clrScheme, 'dk2', defaults.dk2),
      lt2: readColor(clrScheme, 'lt2', defaults.lt2),
      accent1: readColor(clrScheme, 'accent1', defaults.accent1),
      accent2: readColor(clrScheme, 'accent2', defaults.accent2),
      accent3: readColor(clrScheme, 'accent3', defaults.accent3),
      accent4: readColor(clrScheme, 'accent4', defaults.accent4),
      accent5: readColor(clrScheme, 'accent5', defaults.accent5),
      accent6: readColor(clrScheme, 'accent6', defaults.accent6),
      hlink: readColor(clrScheme, 'hlink', defaults.hlink),
      folHlink: readColor(clrScheme, 'folHlink', defaults.folHlink),
    };

    const majorLatin = firstByLocal(childByLocal(fontScheme, 'majorFont'), 'latin');
    const minorLatin = firstByLocal(childByLocal(fontScheme, 'minorFont'), 'latin');
    const fonts: ThemeFontPair = {
      id: 'custom',
      label: `${attr(majorLatin, 'typeface') || 'Calibri Light'} / ${attr(minorLatin, 'typeface') || 'Calibri'}`,
      major: attr(majorLatin, 'typeface') || 'Calibri Light',
      minor: attr(minorLatin, 'typeface') || 'Calibri',
    };

    return { part, name, colors, fonts };
  }

  /** Разрешение цвета с учётом lumMod/lumOff и псевдонимов tx1/bg1. */
  private resolveColor(node: Element | null, theme: PptxThemeModel): string | undefined {
    if (!node) return undefined;
    const srgb = firstByLocal(node, 'srgbClr');
    if (srgb) {
      const mod = numAttr(firstByLocal(srgb, 'lumMod'), 'val', 100000) / 100000;
      const off = numAttr(firstByLocal(srgb, 'lumOff'), 'val', 0) / 100000;
      const base = normalizeHexColor(attr(srgb, 'val'), '000000');
      return mod === 1 && off === 0 ? base : applyLuminance(base, mod * 100000, off * 100000);
    }
    const scheme = firstByLocal(node, 'schemeClr');
    if (scheme) {
      const key = attr(scheme, 'val').toLowerCase();
      const aliases: Record<string, keyof ThemeColorSet> = {
        tx1: 'dk1',
        bg1: 'lt1',
        tx2: 'dk2',
        bg2: 'lt2',
        dk1: 'dk1',
        lt1: 'lt1',
        dk2: 'dk2',
        lt2: 'lt2',
        accent1: 'accent1',
        accent2: 'accent2',
        accent3: 'accent3',
        accent4: 'accent4',
        accent5: 'accent5',
        accent6: 'accent6',
        hlink: 'hlink',
        folHlink: 'folHlink',
      };
      const slot = aliases[key];
      if (!slot) return undefined;
      const base = theme.colors[slot];
      const mod = numAttr(firstByLocal(scheme, 'lumMod'), 'val', 100000) / 100000;
      const off = numAttr(firstByLocal(scheme, 'lumOff'), 'val', 0) / 100000;
      return mod === 1 && off === 0 ? base : applyLuminance(base, mod * 100000, off * 100000);
    }
    return undefined;
  }

  private parseTextBody(txBody: Element | null, theme: PptxThemeModel): PptxTextBody | undefined {
    if (!txBody) return undefined;
    const bodyPr = childByLocal(txBody, 'bodyPr');
    const paragraphs: PptxParagraph[] = [];

    for (const paragraph of childrenByLocal(txBody, 'p')) {
      const pPr = childByLocal(paragraph, 'pPr');
      const runs: PptxTextRun[] = [];

      for (const run of [...paragraph.children]) {
        if (run.localName === 'r') {
          const rPr = childByLocal(run, 'rPr');
          const text = firstByLocal(run, 't')?.textContent ?? '';
          const size = numAttr(rPr, 'sz') / 100;
          const color = this.resolveColor(rPr ? childByLocal(rPr, 'solidFill') : null, theme);
          const schemeNode = rPr ? firstByLocal(childByLocal(rPr, 'solidFill'), 'schemeClr') : null;
          const family = attr(childByLocal(rPr, 'latin'), 'typeface') || undefined;
          const highlight =
            this.resolveColor(rPr ? childByLocal(rPr, 'highlight') : null, theme) ?? undefined;
          runs.push({
            text,
            bold: attr(rPr, 'b') === '1' || undefined,
            italic: attr(rPr, 'i') === '1' || undefined,
            underline: attr(rPr, 'u') !== '' && attr(rPr, 'u') !== 'none' ? true : undefined,
            strike: attr(rPr, 'strike') !== '' && attr(rPr, 'strike') !== 'noStrike' ? true : undefined,
            size: size > 0 ? size : undefined,
            color,
            schemeColor: attr(schemeNode, 'val') || undefined,
            family,
            highlight,
            lang: attr(rPr, 'lang') || undefined,
          });
        } else if (run.localName === 'br') {
          runs.push({ text: '\n' });
        } else if (run.localName === 'fld') {
          const text = firstByLocal(run, 't')?.textContent ?? '';
          if (text) runs.push({ text });
        }
      }

      const bulletNode = pPr ? firstByLocal(pPr, 'buChar') : null;
      const numberNode = pPr ? firstByLocal(pPr, 'buAutoNum') : null;
      const noBullet = pPr ? firstByLocal(pPr, 'buNone') : null;
      const lnSpc = firstByLocal(pPr, 'lnSpc');
      const spcBef = firstByLocal(pPr, 'spcBef');
      const spcAft = firstByLocal(pPr, 'spcAft');

      paragraphs.push({
        runs,
        align: (attr(pPr, 'algn') as PptxParagraph['align']) || undefined,
        level: numAttr(pPr, 'lvl', 0) || undefined,
        bullet: numberNode ? 'number' : bulletNode ? 'bullet' : noBullet ? 'none' : undefined,
        lineSpacing: lnSpc ? numAttr(firstByLocal(lnSpc, 'spcPct'), 'val') / 1000 || undefined : undefined,
        spaceBefore: spcBef ? numAttr(firstByLocal(spcBef, 'spcPts'), 'val') / 100 || undefined : undefined,
        spaceAfter: spcAft ? numAttr(firstByLocal(spcAft, 'spcPts'), 'val') / 100 || undefined : undefined,
      });
    }

    const anchor = attr(bodyPr, 'anchor');
    const verticalAnchor =
      anchor === 'ctr' ? 'ctr' : anchor === 'b' ? 'b' : ('t' as PptxTextBody['verticalAnchor']);
    const autofit = childByLocal(bodyPr, 'spAutoFit')
      ? 'spAutoFit'
      : childByLocal(bodyPr, 'normAutofit')
        ? 'normAutofit'
        : 'none';

    return { paragraphs, verticalAnchor, autofit };
  }

  private parseShape(
    node: Element,
    slideRels: Map<string, string>,
    theme: PptxThemeModel
  ): PptxShapeModel | null {
    const kind: PptxShapeKind =
      node.localName === 'pic'
        ? 'image'
        : node.localName === 'graphicFrame'
          ? 'table'
          : node.localName === 'grpSp'
            ? 'group'
            : 'text';

    const nv = firstByLocal(node, 'cNvPr');
    const id = attr(nv, 'id');
    if (!id) return null;

    const spPr = firstByLocal(node, 'spPr');
    const xfrm =
      firstByLocal(spPr, 'xfrm') ??
      firstByLocal(firstByLocal(node, 'grpSpPr'), 'xfrm') ??
      firstByLocal(node, 'xfrm');
    // Для p:graphicFrame трансформ лежит в p:xfrm.
    const frameXfrm = firstByLocal(node, 'xfrm');
    const effectiveXfrm = xfrm ?? frameXfrm;
    const off = effectiveXfrm ? childByLocal(effectiveXfrm, 'off') : null;
    const ext = effectiveXfrm ? childByLocal(effectiveXfrm, 'ext') : null;

    const rotation = numAttr(effectiveXfrm, 'rot') / 60000 || undefined;
    const flipH = attr(effectiveXfrm, 'flipH') === '1' || undefined;
    const flipV = attr(effectiveXfrm, 'flipV') === '1' || undefined;

    const ph = firstByLocal(node, 'ph');
    const placeholderType = attr(ph, 'type') || (ph ? 'body' : '');

    const geom = attr(firstByLocal(spPr, 'prstGeom'), 'prst') || undefined;

    const fillNode = spPr
      ? childByLocal(spPr, 'solidFill') ?? childByLocal(spPr, 'gradFill')
      : null;
    const fill = fillNode
      ? this.resolveColor(fillNode, theme) ?? (childByLocal(spPr, 'noFill') ? null : undefined)
      : childByLocal(spPr, 'noFill')
        ? null
        : undefined;

    const ln = spPr ? childByLocal(spPr, 'ln') : null;
    const outline = ln
      ? {
          color: this.resolveColor(childByLocal(ln, 'solidFill'), theme),
          width: numAttr(ln, 'w') / 12700 || undefined,
        }
      : undefined;

    const shadow = Boolean(spPr ? childByLocal(spPr, 'effectLst') : null);

    const shape: PptxShapeModel = {
      id,
      name: attr(nv, 'name'),
      kind,
      x: numAttr(off, 'x'),
      y: numAttr(off, 'y'),
      w: numAttr(ext, 'cx'),
      h: numAttr(ext, 'cy'),
      rotation,
      flipH,
      flipV,
      geom,
      fill,
      outline,
      shadow,
      placeholderType: placeholderType || undefined,
      isTitle: placeholderType === 'title' || placeholderType === 'ctrTitle' || undefined,
    };

    const text = this.parseTextBody(firstByLocal(node, 'txBody'), theme);
    if (text && text.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text))) {
      shape.text = text;
      if (kind === 'text') shape.kind = placeholderType ? 'placeholder' : 'text';
    }

    if (kind === 'image') {
      const blip = firstByLocal(node, 'blip');
      const embed = blip ? attr(blip, 'r:embed') || attr(blip, 'embed') : '';
      const target = slideRels.get(embed) ?? '';
      if (target) {
        const part = target.startsWith('/') ? target.slice(1) : `ppt/${target.replace(/^\.\.\//, '')}`;
        const cached = this.mediaCache.get(part);
        shape.image = { part, src: cached ?? '' };
      }
    }

    if (kind === 'table') {
      const tbl = firstByLocal(node, 'tbl');
      if (tbl) {
        const colWidths = childrenByLocal(tbl, 'gridCol').map((col) => numAttr(col, 'w'));
        const rows: PptxTableCell[][] = childrenByLocal(tbl, 'tr').map((tr) =>
          childrenByLocal(tr, 'tc').map((tc) => {
            const tcPr = childByLocal(tc, 'tcPr');
            return {
              text: this.parseTextBody(firstByLocal(tc, 'txBody'), theme) ?? { paragraphs: [] },
              fill: this.resolveColor(tcPr ? childByLocal(tcPr, 'solidFill') : null, theme),
            };
          })
        );
        shape.table = {
          rows,
          colWidths,
          rowHeights: childrenByLocal(tbl, 'tr').map((tr) => numAttr(tr, 'h')),
        };
      }
    }

    return shape;
  }

  private slideRelsMap(part: string): Map<string, string> {
    const relsPart = part.replace(/slides\//, 'slides/_rels/') + '.rels';
    const xml = this.parts.get(relsPart) ?? '';
    const map = new Map<string, string>();
    for (const rel of xml.matchAll(/<Relationship\b[^>]*\/>/g)) {
      const tag = rel[0];
      const id = /Id="([^"]+)"/.exec(tag)?.[1];
      const target = /Target="([^"]+)"/.exec(tag)?.[1];
      if (id && target) map.set(id, target);
    }
    return map;
  }

  private parseTransition(doc: Document): PptxTransitionModel | undefined {
    const node = firstByLocal(childByLocal(doc, 'sld'), 'transition') ?? firstByLocal(doc, 'transition');
    if (!node) return undefined;
    const child = [...node.children][0];
    if (!child) return { type: 'none' };
    const type: PptxTransitionKind =
      child.localName === 'fade'
        ? 'fade'
        : child.localName === 'push'
          ? 'push'
          : child.localName === 'wipe'
            ? 'wipe'
            : child.localName === 'cut'
              ? 'cut'
              : 'none';
    const speed = attr(child, 'spd') as PptxTransitionModel['speed'];
    return {
      type,
      direction: attr(child, 'dir') || undefined,
      speed: speed === 'slow' || speed === 'med' || speed === 'fast' ? speed : undefined,
      advanceOnClick: attr(node, 'advClick') === '0' ? false : true,
      advanceAfterMs: numAttr(node, 'advTm') || undefined,
    };
  }

  private notesPartFor(part: string): string {
    const rels = this.slideRelsMap(part);
    for (const [, target] of rels) {
      if (!/notesSlide\d+\.xml$/.test(target)) continue;
      return target.startsWith('/') ? target.slice(1) : `ppt/${target.replace(/^\.\.\//, '')}`;
    }
    const number = Number(part.replace(/\D/g, ''));
    return `ppt/notesSlides/notesSlide${number}.xml`;
  }

  private parseNotes(part: string): string {
    const notesPart = this.notesPartFor(part);
    const xml = this.parts.get(notesPart);
    if (!xml) return '';
    const doc = parseXml(xml);
    const texts = allByLocal(doc, 't')
      .map((node) => node.textContent?.trim() ?? '')
      .filter(Boolean);
    // Пустые служебные плейсхолдеры отбрасываем, текст докладчика склеиваем.
    return texts.join('\n');
  }

  private parseSlide(part: string, index: number, theme: PptxThemeModel): PptxSlideModel {
    const xml = this.parts.get(part) ?? '';
    const doc = parseXml(xml || '<p:sld xmlns:p="' + P_NS + '"/>');
    const rels = this.slideRelsMap(part);

    // Фон: p:bg с заливкой либо ссылкой на тему.
    let background = theme.colors.lt1;
    let backgroundImage: string | undefined;
    const bg = firstByLocal(doc, 'bg');
    if (bg) {
      const solid = this.resolveColor(firstByLocal(bg, 'solidFill'), theme);
      if (solid) background = solid;
      const blip = firstByLocal(bg, 'blip');
      const embed = blip ? attr(blip, 'r:embed') || attr(blip, 'embed') : '';
      if (embed && rels.get(embed)) {
        const target = rels.get(embed) ?? '';
        const imagePart = target.startsWith('/') ? target.slice(1) : `ppt/${target.replace(/^\.\.\//, '')}`;
        backgroundImage = this.mediaCache.get(imagePart) ?? '';
        this.backgroundMedia.set(part, imagePart);
        if (!this.mediaCache.has(imagePart)) this.pendingMedia.push({ part: imagePart });
      }
      const gradient = this.resolveColor(firstByLocal(bg, 'gradFill'), theme);
      if (!solid && gradient) background = gradient;
    }

    const spTree = firstByLocal(doc, 'spTree');
    const shapes: PptxShapeModel[] = [];
    for (const node of spTree ? [...spTree.children] : []) {
      // p:nvGrpSpPr и p:grpSpPr — служебные элементы дерева, а не фигуры:
      // в nvGrpSpPr лежит cNvPr с id="1", и раньше он попадал в модель
      // фантомной фигурой без геометрии.
      if (!['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'].includes(node.localName)) continue;
      const shape = this.parseShape(node, rels, theme);
      if (shape) shapes.push(shape);
    }

    let layoutPart = '';
    for (const [, target] of rels) {
      if (/slideLayout\d+\.xml$/.test(target)) {
        layoutPart = target.startsWith('/') ? target.slice(1) : `ppt/${target.replace(/^\.\.\//, '')}`;
        break;
      }
    }

    const layoutName = this.doc?.layouts?.find((layout) => layout.part === layoutPart)?.name ?? '';
    const hidden = attr(doc.documentElement, 'show') === '0' || undefined;

    return {
      index,
      part,
      layoutPart: layoutPart || undefined,
      layoutName: layoutName || undefined,
      hidden,
      background,
      backgroundImage,
      shapes,
      notes: this.parseNotes(part),
      transition: this.parseTransition(doc),
    };
  }

  /** Догрузка изображений: вызывается после открытия, чтобы не блокировать первый кадр. */
  async loadMedia(): Promise<void> {
    const unique = [...new Set(this.pendingMedia.map((item) => item.part))];
    this.pendingMedia = [];
    for (const part of unique) await this.mediaUrl(part);
  }

  /** data URL для медиа-части (кешируется). */
  async mediaUrl(part: string): Promise<string> {
    const cached = this.mediaCache.get(part);
    if (cached !== undefined) return cached;
    let bytes = this.binaryParts.get(part);
    if (!bytes) {
      const file = this.zip.file(part);
      if (!file) return '';
      bytes = await file.async('uint8array');
      this.binaryParts.set(part, bytes);
    }
    const ext = part.split('.').pop()?.toLowerCase() ?? 'png';
    const mime =
      ext === 'jpg' || ext === 'jpeg'
        ? 'image/jpeg'
        : ext === 'gif'
          ? 'image/gif'
          : ext === 'svg'
            ? 'image/svg+xml'
            : ext === 'webp'
              ? 'image/webp'
              : ext === 'bmp'
                ? 'image/bmp'
                : 'image/png';
    const url = `data:${mime};base64,${bytesToBase64(bytes)}`;
    this.mediaCache.set(part, url);
    for (const slide of this.doc.slides) {
      for (const shape of slide.shapes) {
        if (shape.image?.part === part) shape.image.src = url;
      }
      if (this.backgroundMedia.get(slide.part) === part) slide.backgroundImage = url;
    }
    return url;
  }

  /* ---------- Снимки и запись ---------- */

  private capture(parts: string[], label: string): void {
    const snapshot = new Map<string, string | null>();
    for (const part of parts) snapshot.set(part, this.parts.get(part) ?? null);
    this.undoStack.push({ label, parts: snapshot });
    if (this.undoStack.length > 100) this.undoStack.shift();
    this.redoStack = [];
  }

  private restore(entry: UndoEntry): void {
    for (const [part, xml] of entry.parts) {
      if (xml === null) {
        this.parts.delete(part);
      } else {
        this.parts.set(part, xml);
        // Отмена удаления возвращает часть в пакет.
        this.deletedParts.delete(part);
      }
    }
    this.doc = this.parseModel();
  }

  undo(): boolean {
    const entry = this.undoStack.pop();
    if (!entry) return false;
    const current = new Map<string, string | null>();
    for (const part of entry.parts.keys()) current.set(part, this.parts.get(part) ?? null);
    this.redoStack.push({ label: entry.label, parts: current });
    this.restore(entry);
    return true;
  }

  redo(): boolean {
    const entry = this.redoStack.pop();
    if (!entry) return false;
    const current = new Map<string, string | null>();
    for (const part of entry.parts.keys()) current.set(part, this.parts.get(part) ?? null);
    this.undoStack.push({ label: entry.label, parts: current });
    this.restore(entry);
    return true;
  }

  /** Сериализация пакета: только при сохранении. */
  async toBase64(): Promise<string> {
    for (const part of this.deletedParts) this.zip.remove(part);
    for (const [part, xml] of this.parts) this.zip.file(part, xml);
    for (const [part, bytes] of this.binaryParts) this.zip.file(part, bytes);
    return this.zip.generateAsync({
      type: 'base64',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  /* ---------- Вспомогательные обращения к частям ---------- */

  private slidePart(index: number): string | null {
    return this.doc.slides[index]?.part ?? null;
  }

  private slideDoc(index: number): { doc: Document; part: string } | null {
    const part = this.slidePart(index);
    if (!part) return null;
    const xml = this.parts.get(part) ?? '';
    return { doc: parseXml(xml || '<p:sld xmlns:p="' + P_NS + '"/>'), part };
  }

  private commit(part: string, doc: Document): void {
    this.parts.set(part, serializeXml(doc));
  }

  private spTreeOf(doc: Document): Element | null {
    return firstByLocal(doc, 'spTree');
  }

  private findShapeNode(doc: Document, shapeId: string): Element | null {
    for (const node of allByLocal(doc, 'sp')) {
      if (attr(firstByLocal(node, 'cNvPr'), 'id') === shapeId) return node;
    }
    for (const node of allByLocal(doc, 'pic')) {
      if (attr(firstByLocal(node, 'cNvPr'), 'id') === shapeId) return node;
    }
    for (const node of allByLocal(doc, 'graphicFrame')) {
      if (attr(firstByLocal(node, 'cNvPr'), 'id') === shapeId) return node;
    }
    return null;
  }

  private xfrmOf(shape: Element): Element | null {
    const spPr = firstByLocal(shape, 'spPr');
    if (spPr) {
      const xfrm = childByLocal(spPr, 'xfrm');
      if (xfrm) return xfrm;
    }
    return firstByLocal(shape, 'xfrm');
  }

  private setGeometry(doc: Document, shape: Element, geometry: { x: number; y: number; w: number; h: number; rotation?: number; flipH?: boolean; flipV?: boolean }): void {
    const xfrm = this.xfrmOf(shape);
    if (!xfrm) return;
    let off = childByLocal(xfrm, 'off');
    if (!off) {
      off = doc.createElementNS(A_NS, 'a:off');
      xfrm.insertBefore(off, xfrm.firstChild);
    }
    let ext = childByLocal(xfrm, 'ext');
    if (!ext) {
      ext = doc.createElementNS(A_NS, 'a:ext');
      xfrm.appendChild(ext);
    }
    off.setAttribute('x', String(Math.round(geometry.x)));
    off.setAttribute('y', String(Math.round(geometry.y)));
    ext.setAttribute('cx', String(Math.max(1, Math.round(geometry.w))));
    ext.setAttribute('cy', String(Math.max(1, Math.round(geometry.h))));
    if (typeof geometry.rotation === 'number' && geometry.rotation !== 0) {
      xfrm.setAttribute('rot', String(Math.round(geometry.rotation * 60000)));
    } else {
      xfrm.removeAttribute('rot');
    }
    if (geometry.flipH) xfrm.setAttribute('flipH', '1');
    else xfrm.removeAttribute('flipH');
    if (geometry.flipV) xfrm.setAttribute('flipV', '1');
    else xfrm.removeAttribute('flipV');
  }

  private ensureSpPr(shape: Element, doc: Document): Element {
    const existing = firstByLocal(shape, 'spPr');
    if (existing) return existing;
    const created = doc.createElementNS(P_NS, 'p:spPr');
    shape.appendChild(created);
    return created;
  }

  private buildParagraph(doc: Document, paragraph: PptxParagraph): Element {
    const p = doc.createElementNS(A_NS, 'a:p');
    const pPr = doc.createElementNS(A_NS, 'a:pPr');
    if (paragraph.align) pPr.setAttribute('algn', paragraph.align);
    const level = paragraph.level ?? 0;
    if (level > 0) pPr.setAttribute('lvl', String(level));
    const indent = 342900;
    if (paragraph.bullet === 'bullet' || paragraph.bullet === 'number') {
      pPr.setAttribute('marL', String(indent * (level + 1)));
      pPr.setAttribute('indent', String(-indent));
      if (paragraph.bullet === 'number') {
        const num = doc.createElementNS(A_NS, 'a:buAutoNum');
        num.setAttribute('type', 'arabicPeriod');
        pPr.appendChild(num);
      } else {
        const bullet = doc.createElementNS(A_NS, 'a:buChar');
        bullet.setAttribute('char', '•');
        pPr.appendChild(bullet);
      }
    } else if (paragraph.bullet === 'none') {
      pPr.appendChild(doc.createElementNS(A_NS, 'a:buNone'));
    }
    if (paragraph.lineSpacing) {
      const lnSpc = doc.createElementNS(A_NS, 'a:lnSpc');
      const pct = doc.createElementNS(A_NS, 'a:spcPct');
      pct.setAttribute('val', String(Math.round(paragraph.lineSpacing * 1000)));
      lnSpc.appendChild(pct);
      pPr.appendChild(lnSpc);
    }
    if (paragraph.spaceBefore) {
      const spcBef = doc.createElementNS(A_NS, 'a:spcBef');
      const pts = doc.createElementNS(A_NS, 'a:spcPts');
      pts.setAttribute('val', String(Math.round(paragraph.spaceBefore * 100)));
      spcBef.appendChild(pts);
      pPr.appendChild(spcBef);
    }
    if (paragraph.spaceAfter) {
      const spcAft = doc.createElementNS(A_NS, 'a:spcAft');
      const pts = doc.createElementNS(A_NS, 'a:spcPts');
      pts.setAttribute('val', String(Math.round(paragraph.spaceAfter * 100)));
      spcAft.appendChild(pts);
      pPr.appendChild(spcAft);
    }
    if (pPr.childNodes.length > 0) p.appendChild(pPr);

    for (const run of paragraph.runs) {
      if (run.text === '\n') {
        p.appendChild(doc.createElementNS(A_NS, 'a:br'));
        continue;
      }
      const r = doc.createElementNS(A_NS, 'a:r');
      const rPr = doc.createElementNS(A_NS, 'a:rPr');
      rPr.setAttribute('lang', run.lang || 'ru-RU');
      if (run.size) rPr.setAttribute('sz', String(Math.round(run.size * 100)));
      if (run.bold) rPr.setAttribute('b', '1');
      if (run.italic) rPr.setAttribute('i', '1');
      if (run.underline) rPr.setAttribute('u', 'sng');
      if (run.strike) rPr.setAttribute('strike', 'sngStrike');
      if (run.color) {
        const fill = doc.createElementNS(A_NS, 'a:solidFill');
        if (run.schemeColor) {
          // Цвет из схемы темы сохраняем ссылкой: смена темы продолжит работать.
          const scheme = doc.createElementNS(A_NS, 'a:schemeClr');
          scheme.setAttribute('val', run.schemeColor);
          fill.appendChild(scheme);
        } else {
          const color = doc.createElementNS(A_NS, 'a:srgbClr');
          color.setAttribute('val', normalizeHexColor(run.color, '000000'));
          fill.appendChild(color);
        }
        rPr.appendChild(fill);
      }
      if (run.highlight) {
        const highlight = doc.createElementNS(A_NS, 'a:highlight');
        const color = doc.createElementNS(A_NS, 'a:srgbClr');
        color.setAttribute('val', normalizeHexColor(run.highlight, 'FFFF00'));
        highlight.appendChild(color);
        rPr.appendChild(highlight);
      }
      if (run.family) {
        const latin = doc.createElementNS(A_NS, 'a:latin');
        latin.setAttribute('typeface', run.family);
        rPr.appendChild(latin);
      }
      r.appendChild(rPr);
      const t = doc.createElementNS(A_NS, 'a:t');
      t.setAttribute('xml:space', 'preserve');
      t.textContent = run.text;
      r.appendChild(t);
      p.appendChild(r);
    }
    return p;
  }

  private writeTextBody(doc: Document, txBody: Element, body: PptxTextBody): void {
    const bodyPr = childByLocal(txBody, 'bodyPr');
    for (const child of [...txBody.children]) {
      if (child.localName === 'bodyPr') continue;
      txBody.removeChild(child);
    }
    let bodyPrNode = bodyPr;
    if (!bodyPrNode) {
      bodyPrNode = doc.createElementNS(A_NS, 'a:bodyPr');
      txBody.insertBefore(bodyPrNode, txBody.firstChild);
    }
    bodyPrNode.setAttribute('wrap', 'square');
    if (body.verticalAnchor && body.verticalAnchor !== 't') {
      bodyPrNode.setAttribute('anchor', body.verticalAnchor);
    } else {
      bodyPrNode.removeAttribute('anchor');
    }
    for (const child of [...bodyPrNode.children]) {
      if (child.localName === 'spAutoFit' || child.localName === 'normAutofit') bodyPrNode.removeChild(child);
    }
    if (body.autofit === 'spAutoFit') bodyPrNode.appendChild(doc.createElementNS(A_NS, 'a:spAutoFit'));
    else if (body.autofit === 'normAutofit') bodyPrNode.appendChild(doc.createElementNS(A_NS, 'a:normAutofit'));

    txBody.appendChild(doc.createElementNS(A_NS, 'a:lstStyle'));
    const paragraphs = body.paragraphs.length > 0 ? body.paragraphs : [{ runs: [] as PptxTextRun[] }];
    for (const paragraph of paragraphs) txBody.appendChild(this.buildParagraph(doc, paragraph));
  }

  /* ---------- Тема ---------- */

  /**
   * Применение темы: цвета и шрифты схемы. Пресет целиком либо частично
   * (только цвета / только шрифты) — как «Варианты» в PowerPoint.
   */
  applyTheme(options: { presetId?: string; colors?: Partial<ThemeColorSet>; fonts?: ThemeFontPair; label?: string }): void {
    const part = this.doc.theme.part;
    const current = this.parts.get(part) ?? '';
    this.capture([part], options.label ?? 'Тема оформления');
    const doc = parseXml(current || '<a:theme xmlns:a="' + A_NS + '"/>');

    const preset = options.presetId ? presetById(options.presetId) : null;
    const nextColors: ThemeColorSet = {
      ...this.doc.theme.colors,
      ...(preset ? preset.colors : {}),
      ...(options.colors ?? {}),
    };
    const nextFonts: ThemeFontPair = {
      ...this.doc.theme.fonts,
      ...(preset ? preset.fonts : {}),
      ...(options.fonts ?? {}),
    };

    let clrScheme = firstByLocal(doc, 'clrScheme');
    if (!clrScheme) {
      const themeElements = firstByLocal(doc, 'themeElements');
      clrScheme = doc.createElementNS(A_NS, 'a:clrScheme');
      themeElements?.insertBefore(clrScheme, themeElements.firstChild);
    }
    if (clrScheme) {
      clrScheme.setAttribute('name', preset ? preset.label : this.doc.theme.name);
      for (const [slot, value] of Object.entries(nextColors)) {
        let node = childByLocal(clrScheme, slot);
        if (!node) {
          node = doc.createElementNS(A_NS, 'a:' + slot);
          clrScheme.appendChild(node);
        }
        for (const child of [...node.children]) node.removeChild(child);
        const color = doc.createElementNS(A_NS, 'a:srgbClr');
        color.setAttribute('val', normalizeHexColor(value, '000000'));
        node.appendChild(color);
      }
    }

    const fontScheme = firstByLocal(doc, 'fontScheme');
    if (fontScheme) {
      fontScheme.setAttribute('name', nextFonts.label);
      const major = childByLocal(fontScheme, 'majorFont');
      const minor = childByLocal(fontScheme, 'minorFont');
      const setLatin = (parent: Element | null, family: string) => {
        if (!parent) return;
        let latin = childByLocal(parent, 'latin');
        if (!latin) {
          latin = doc.createElementNS(A_NS, 'a:latin');
          parent.insertBefore(latin, parent.firstChild);
        }
        latin.setAttribute('typeface', family);
      };
      setLatin(major, nextFonts.major);
      setLatin(minor, nextFonts.minor);
    }

    this.commit(part, doc);
    this.doc = this.parseModel();
  }

  /* ---------- Фон слайда ---------- */

  setSlideBackground(
    slideIndex: number,
    background: { color?: string; imageDataUrl?: string; all?: boolean }
  ): void {
    const targets = background.all ? this.doc.slides.map((_, index) => index) : [slideIndex];
    for (const index of targets) {
      const slide = this.slideDoc(index);
      if (!slide) continue;
      this.capture([slide.part], 'Фон слайда');
      const root = slide.doc.documentElement;
      if (!root) continue;
      const cSld = childByLocal(root, 'cSld');
      if (!cSld) continue;
      let bg = childByLocal(cSld, 'bg');
      if (!bg) {
        bg = slide.doc.createElementNS(P_NS, 'p:bg');
        // p:bg — первый дочерний элемент p:cSld.
        cSld.insertBefore(bg, cSld.firstChild);
      }
      for (const child of [...bg.children]) bg.removeChild(child);

      if (background.imageDataUrl) {
        const image = this.writeMedia(slide.part, background.imageDataUrl);
        const bgPr = slide.doc.createElementNS(P_NS, 'p:bgPr');
        const blipFill = slide.doc.createElementNS(A_NS, 'a:blipFill');
        const blip = slide.doc.createElementNS(A_NS, 'a:blip');
        blip.setAttributeNS(R_NS, 'r:embed', image.relId);
        const stretch = slide.doc.createElementNS(A_NS, 'a:stretch');
        stretch.appendChild(slide.doc.createElementNS(A_NS, 'a:fillRect'));
        blipFill.append(blip, stretch);
        const effectLst = slide.doc.createElementNS(A_NS, 'a:effectLst');
        bgPr.append(blipFill, effectLst);
        bg.appendChild(bgPr);
      } else {
        const bgPr = slide.doc.createElementNS(P_NS, 'p:bgPr');
        const fill = slide.doc.createElementNS(A_NS, 'a:solidFill');
        const color = slide.doc.createElementNS(A_NS, 'a:srgbClr');
        color.setAttribute('val', normalizeHexColor(background.color ?? '#FFFFFF', 'FFFFFF'));
        fill.appendChild(color);
        const effectLst = slide.doc.createElementNS(A_NS, 'a:effectLst');
        bgPr.append(fill, effectLst);
        bg.appendChild(bgPr);
      }
      this.commit(slide.part, slide.doc);
    }
    this.doc = this.parseModel();
  }

  /* ---------- Переходы ---------- */

  setSlideTransition(slideIndex: number, transition: PptxTransitionModel): void {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return;
    this.capture([slide.part], 'Переход слайда');
    const root = slide.doc.documentElement;
    if (!root) return;
    const existing = childByLocal(root, 'transition');
    if (existing) root.removeChild(existing);

    if (transition.type !== 'none') {
      const node = slide.doc.createElementNS(P_NS, 'p:transition');
      if (transition.advanceOnClick === false) node.setAttribute('advClick', '0');
      if (transition.advanceAfterMs) node.setAttribute('advTm', String(Math.round(transition.advanceAfterMs)));
      const effect = slide.doc.createElementNS(
        P_NS,
        'p:' + (transition.type === 'cut' ? 'cut' : transition.type)
      );
      if (transition.speed && transition.type !== 'cut') effect.setAttribute('spd', transition.speed);
      if (transition.direction && (transition.type === 'push' || transition.type === 'wipe')) {
        effect.setAttribute('dir', transition.direction);
      }
      node.appendChild(effect);
      // Порядок в p:sld: cSld, clrMapOvr, transition, timing.
      const anchor = childByLocal(root, 'clrMapOvr') ?? childByLocal(root, 'cSld');
      if (anchor?.nextSibling) root.insertBefore(node, anchor.nextSibling);
      else root.appendChild(node);
    }
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  /* ---------- Слайды ---------- */

  private presentationInfo(): { xml: string; rels: string; width: number; height: number } {
    const xml = this.parts.get('ppt/presentation.xml') ?? '';
    const rels = this.parts.get('ppt/_rels/presentation.xml.rels') ?? '';
    const sizeMatch = /<p:sldSz\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(xml);
    return {
      xml,
      rels,
      width: sizeMatch ? Number(sizeMatch[1]) : 12192000,
      height: sizeMatch ? Number(sizeMatch[2]) : 6858000,
    };
  }

  private nextSlideNumber(): number {
    const numbers = Object.keys(this.zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .map((name) => Number(name.replace(/\D/g, '')));
    for (const part of this.parts.keys()) {
      if (!/^ppt\/slides\/slide\d+\.xml$/.test(part)) continue;
      numbers.push(Number(part.replace(/\D/g, '')));
    }
    return (numbers.length ? Math.max(...numbers) : 0) + 1;
  }

  private nextShapeId(doc: Document): number {
    let max = 1;
    for (const node of allByLocal(doc, 'cNvPr')) {
      const id = Number(attr(node, 'id'));
      if (Number.isFinite(id) && id > max) max = id;
    }
    return max + 1;
  }

  addSlide(index: number, layoutPart?: string, title = 'Новый слайд'): number {
    const info = this.presentationInfo();
    const number = this.nextSlideNumber();
    const part = `ppt/slides/slide${number}.xml`;
    const relsPart = `ppt/slides/_rels/slide${number}.xml.rels`;
    const layout = layoutPart
      ? layoutPart.replace(/^\/?ppt\//, '')
      : this.doc.layouts[0]?.part.replace(/^\/?ppt\//, '') ?? 'slideLayouts/slideLayout1.xml';
    const layoutTarget = '../' + layout.replace(/^\/?/, '');

    this.capture(
      [part, relsPart, 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'],
      'Добавление слайда'
    );

    const titleText = escapeXml(title);
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
      `<p:cSld><p:spTree>` +
      '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Заголовок 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="' +
      Math.round(info.width * 0.08) +
      '" y="' +
      Math.round(info.height * 0.08) +
      '"/><a:ext cx="' +
      Math.round(info.width * 0.84) +
      '" cy="' +
      Math.round(info.height * 0.2) +
      '"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>' +
      '<a:p><a:r><a:rPr lang="ru-RU" sz="3200" b="1"/><a:t>' +
      titleText +
      '</a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Текст 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="' +
      Math.round(info.width * 0.08) +
      '" y="' +
      Math.round(info.height * 0.32) +
      '"/><a:ext cx="' +
      Math.round(info.width * 0.84) +
      '" cy="' +
      Math.round(info.height * 0.5) +
      '"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>' +
      '<a:p><a:r><a:rPr lang="ru-RU" sz="1800"/><a:t>Текст слайда</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
    this.parts.set(part, xml);
    this.parts.set(
      relsPart,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        `<Relationships xmlns="${REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_TYPE.slideLayout}" Target="${layoutTarget}"/>` +
        '</Relationships>'
    );

    // Регистрация части в пакете и в списке слайдов.
    const relId = this.nextRelationshipId(info.rels);
    const slideId = this.nextSlideId(info.xml);
    let presentationXml = info.xml;
    const listMatch = /<p:sldIdLst\b[^>]*>([\s\S]*?)<\/p:sldIdLst>/.exec(presentationXml);
    if (listMatch) {
      const ids = [...(listMatch[1] ?? '').matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
      ids.splice(Math.max(0, Math.min(index, ids.length)), 0, `<p:sldId id="${slideId}" r:id="${relId}"/>`);
      presentationXml = presentationXml.replace(
        listMatch[0],
        `<p:sldIdLst>${ids.join('')}</p:sldIdLst>`
      );
    }
    this.parts.set('ppt/presentation.xml', presentationXml);
    this.parts.set(
      'ppt/_rels/presentation.xml.rels',
      info.rels.replace(
        '</Relationships>',
        `<Relationship Id="${relId}" Type="${REL_TYPE.slide}" Target="slides/slide${number}.xml"/></Relationships>`
      )
    );
    this.addContentTypeOverride(`/ppt/slides/slide${number}.xml`, CONTENT_TYPE.slide);

    this.doc = this.parseModel();
    return this.doc.slides.findIndex((slide) => slide.part === part);
  }

  deleteSlide(index: number): void {
    const slide = this.doc.slides[index];
    if (!slide || this.doc.slides.length <= 1) return;
    const info = this.presentationInfo();
    const ids = [...info.xml.matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
    const target = ids[index] ?? '';
    const relId = /r:id="([^"]+)"/.exec(target)?.[1] ?? '';
    const relsPart = 'ppt/slides/_rels/' + slide.part.split('/').pop() + '.rels';
    const notesPart = this.notesPartFor(slide.part);
    const notesRelsPart = notesPart.replace(/notesSlides\//, 'notesSlides/_rels/') + '.rels';

    this.capture(
      [slide.part, relsPart, notesPart, notesRelsPart, 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'],
      'Удаление слайда'
    );

    for (const part of [slide.part, relsPart, notesPart, notesRelsPart]) {
      if (!this.parts.has(part) && !this.zip.file(part)) continue;
      this.parts.delete(part);
      this.deletedParts.add(part);
    }
    if (target) this.parts.set('ppt/presentation.xml', info.xml.replace(target, ''));
    if (relId) {
      const relPattern = new RegExp(`<Relationship\\b[^>]*Id="${relId}"[^>]*/>`);
      this.parts.set(
        'ppt/_rels/presentation.xml.rels',
        info.rels.replace(relPattern, '')
      );
    }
    this.removeContentTypeOverride(`/${slide.part}`);
    this.doc = this.parseModel();
  }

  duplicateSlide(index: number): number {
    const slide = this.doc.slides[index];
    if (!slide) return index;
    const info = this.presentationInfo();
    const number = this.nextSlideNumber();
    const part = `ppt/slides/slide${number}.xml`;
    const relsPart = 'ppt/slides/_rels/' + part.split('/').pop() + '.rels';
    const sourceRels = 'ppt/slides/_rels/' + slide.part.split('/').pop() + '.rels';

    this.capture([part, relsPart, 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'], 'Дублирование слайда');

    this.parts.set(part, this.parts.get(slide.part) ?? '');
    this.parts.set(relsPart, this.parts.get(sourceRels) ?? '');

    const relId = this.nextRelationshipId(info.rels);
    const slideId = this.nextSlideId(info.xml);
    let presentationXml = info.xml;
    const listMatch = /<p:sldIdLst\b[^>]*>([\s\S]*?)<\/p:sldIdLst>/.exec(presentationXml);
    if (listMatch) {
      const ids = [...(listMatch[1] ?? '').matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
      ids.splice(index + 1, 0, `<p:sldId id="${slideId}" r:id="${relId}"/>`);
      presentationXml = presentationXml.replace(listMatch[0], `<p:sldIdLst>${ids.join('')}</p:sldIdLst>`);
    }
    this.parts.set('ppt/presentation.xml', presentationXml);
    this.parts.set(
      'ppt/_rels/presentation.xml.rels',
      info.rels.replace(
        '</Relationships>',
        `<Relationship Id="${relId}" Type="${REL_TYPE.slide}" Target="slides/slide${number}.xml"/></Relationships>`
      )
    );
    this.addContentTypeOverride(`/ppt/slides/slide${number}.xml`, CONTENT_TYPE.slide);
    this.doc = this.parseModel();
    return index + 1;
  }

  moveSlide(from: number, to: number): number {
    const info = this.presentationInfo();
    const listMatch = /<p:sldIdLst\b[^>]*>([\s\S]*?)<\/p:sldIdLst>/.exec(info.xml);
    if (!listMatch) return from;
    const ids = [...(listMatch[1] ?? '').matchAll(/<p:sldId\b[^>]*\/>/g)].map((match) => match[0]);
    if (from < 0 || from >= ids.length) return from;
    this.capture(['ppt/presentation.xml'], 'Перестановка слайдов');
    const without = ids.filter((_, position) => position !== from);
    const target = Math.max(0, Math.min(to, without.length));
    without.splice(target, 0, ids[from]);
    this.parts.set(
      'ppt/presentation.xml',
      info.xml.replace(listMatch[0], `<p:sldIdLst>${without.join('')}</p:sldIdLst>`)
    );
    this.doc = this.parseModel();
    return target;
  }

  setSlideHidden(index: number, hidden: boolean): void {
    const slide = this.slideDoc(index);
    if (!slide) return;
    this.capture([slide.part], hidden ? 'Скрыть слайд' : 'Показать слайд');
    const root = slide.doc.documentElement;
    if (!root) return;
    if (hidden) root.setAttribute('show', '0');
    else root.removeAttribute('show');
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  setSlideLayout(index: number, layoutPart: string): void {
    const slide = this.doc.slides[index];
    if (!slide) return;
    const layout = this.doc.layouts.find((item) => item.part === layoutPart || item.name === layoutPart);
    if (!layout) return;
    const relsPart = 'ppt/slides/_rels/' + slide.part.split('/').pop() + '.rels';
    const relsXml = this.parts.get(relsPart) ?? '';
    if (!relsXml) return;
    this.capture([relsPart], 'Макет слайда');
    const layoutName = layout.part.split('/').pop() ?? '';
    const next = relsXml.replace(
      /(<Relationship\b[^>]*Type="[^"]*slideLayout"[^>]*Target=")[^"]*(")/,
      `$1../slideLayouts/${layoutName}$2`
    );
    this.parts.set(relsPart, next);
    this.doc = this.parseModel();
  }

  setSlideSize(width: number, height: number): void {
    const info = this.presentationInfo();
    this.capture(['ppt/presentation.xml'], 'Размер слайда');
    const next = info.xml.replace(
      /<p:sldSz\b[^>]*\/>/,
      `<p:sldSz cx="${Math.round(width)}" cy="${Math.round(height)}"/>`
    );
    this.parts.set('ppt/presentation.xml', next);
    this.doc = this.parseModel();
  }

  /* ---------- Текст фигур ---------- */

  setShapeText(slideIndex: number, shapeId: string, body: PptxTextBody): void {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return;
    const node = this.findShapeNode(slide.doc, shapeId);
    if (!node) return;
    this.capture([slide.part], 'Правка текста');
    let txBody = childByLocal(node, 'txBody');
    if (!txBody) {
      txBody = slide.doc.createElementNS(P_NS, 'p:txBody');
      const spPr = firstByLocal(node, 'spPr');
      if (spPr && spPr.parentNode === node) node.insertBefore(txBody, spPr.nextSibling);
      else node.appendChild(txBody);
    }
    this.writeTextBody(slide.doc, txBody, body);
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  /* ---------- Фигуры ---------- */

  setShapeGeometry(
    slideIndex: number,
    shapeId: string,
    geometry: { x: number; y: number; w: number; h: number; rotation?: number }
  ): void {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return;
    const node = this.findShapeNode(slide.doc, shapeId);
    if (!node) return;
    this.capture([slide.part], 'Перемещение фигуры');
    this.setGeometry(slide.doc, node, geometry);
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  setShapeFill(slideIndex: number, shapeId: string, color: string | null): void {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return;
    const node = this.findShapeNode(slide.doc, shapeId);
    if (!node) return;
    this.capture([slide.part], 'Заливка фигуры');
    const spPr = this.ensureSpPr(node, slide.doc);
    for (const child of [...spPr.children]) {
      if (['solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill'].includes(child.localName)) {
        spPr.removeChild(child);
      }
    }
    const insertBefore = childByLocal(spPr, 'ln') ?? childByLocal(spPr, 'effectLst');
    if (color === null) {
      const noFill = slide.doc.createElementNS(A_NS, 'a:noFill');
      if (insertBefore) spPr.insertBefore(noFill, insertBefore);
      else spPr.appendChild(noFill);
    } else {
      const fill = slide.doc.createElementNS(A_NS, 'a:solidFill');
      const clr = slide.doc.createElementNS(A_NS, 'a:srgbClr');
      clr.setAttribute('val', normalizeHexColor(color, 'FFFFFF'));
      fill.appendChild(clr);
      if (insertBefore) spPr.insertBefore(fill, insertBefore);
      else spPr.appendChild(fill);
    }
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  setShapeOutline(
    slideIndex: number,
    shapeId: string,
    outline: { color?: string; width?: number } | null
  ): void {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return;
    const node = this.findShapeNode(slide.doc, shapeId);
    if (!node) return;
    this.capture([slide.part], 'Контур фигуры');
    const spPr = this.ensureSpPr(node, slide.doc);
    const existing = childByLocal(spPr, 'ln');
    if (existing) spPr.removeChild(existing);
    if (!outline) {
      this.commit(slide.part, slide.doc);
      this.doc = this.parseModel();
      return;
    }
    const ln = slide.doc.createElementNS(A_NS, 'a:ln');
    if (outline.width) ln.setAttribute('w', String(Math.round(outline.width * 12700)));
    const fill = slide.doc.createElementNS(A_NS, 'a:solidFill');
    const clr = slide.doc.createElementNS(A_NS, 'a:srgbClr');
    clr.setAttribute('val', normalizeHexColor(outline.color ?? '#000000', '000000'));
    fill.appendChild(clr);
    ln.appendChild(fill);
    spPr.appendChild(ln);
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  setShapeShadow(slideIndex: number, shapeId: string, enabled: boolean): void {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return;
    const node = this.findShapeNode(slide.doc, shapeId);
    if (!node) return;
    this.capture([slide.part], 'Тень фигуры');
    const spPr = this.ensureSpPr(node, slide.doc);
    const existing = childByLocal(spPr, 'effectLst');
    if (existing) spPr.removeChild(existing);
    if (enabled) {
      const effect = slide.doc.createElementNS(A_NS, 'a:effectLst');
      const shadow = slide.doc.createElementNS(A_NS, 'a:outerShdw');
      shadow.setAttribute('blurRad', '40000');
      shadow.setAttribute('dist', '23000');
      shadow.setAttribute('dir', '5400000');
      shadow.setAttribute('rotWithShape', '0');
      const color = slide.doc.createElementNS(A_NS, 'a:srgbClr');
      color.setAttribute('val', '000000');
      const alpha = slide.doc.createElementNS(A_NS, 'a:alpha');
      alpha.setAttribute('val', '35000');
      color.appendChild(alpha);
      shadow.appendChild(color);
      effect.appendChild(shadow);
      spPr.appendChild(effect);
    }
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  deleteShape(slideIndex: number, shapeId: string): void {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return;
    const node = this.findShapeNode(slide.doc, shapeId);
    if (!node) return;
    this.capture([slide.part], 'Удаление фигуры');
    node.parentNode?.removeChild(node);
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  /** Копия фигуры со смещением: id и имя новые, содержимое сохраняется. */
  duplicateShape(slideIndex: number, shapeId: string): string | null {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return null;
    const node = this.findShapeNode(slide.doc, shapeId);
    const parent = node?.parentNode as Element | null;
    if (!node || !parent) return null;
    this.capture([slide.part], 'Дублирование фигуры');
    const id = this.nextShapeId(slide.doc);
    const clone = node.cloneNode(true) as Element;
    const cNvPr = firstByLocal(clone, 'cNvPr');
    if (cNvPr) {
      cNvPr.setAttribute('id', String(id));
      cNvPr.setAttribute('name', `${attr(cNvPr, 'name')} ${id}`);
    }
    const xfrm = this.xfrmOf(clone);
    const off = xfrm ? childByLocal(xfrm, 'off') : null;
    if (off) {
      off.setAttribute('x', String(numAttr(off, 'x') + 24 * 9525));
      off.setAttribute('y', String(numAttr(off, 'y') + 24 * 9525));
    }
    parent.appendChild(clone);
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
    return String(id);
  }

  /** Порядок наложения: вперёд/назад на один шаг или сразу в край. */
  reorderShape(slideIndex: number, shapeId: string, mode: 'front' | 'back' | 'forward' | 'backward'): void {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return;
    const node = this.findShapeNode(slide.doc, shapeId);
    const parent = node?.parentNode as Element | null;
    if (!node || !parent) return;
    this.capture([slide.part], 'Порядок фигур');
    const siblings = [...parent.children];
    const position = siblings.indexOf(node as Element);
    if (position < 0) return;
    const others = siblings.filter((child) => child !== node);
    const insertAt =
      mode === 'front'
        ? others.length
        : mode === 'back'
          ? 0
          : mode === 'forward'
            ? Math.min(others.length, position + 1)
            : Math.max(0, position - 1);
    parent.removeChild(node);
    if (insertAt >= others.length) parent.appendChild(node);
    else parent.insertBefore(node, others[insertAt]);
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  addShape(
    slideIndex: number,
    options: {
      geom?: string;
      x?: number;
      y?: number;
      w?: number;
      h?: number;
      fill?: string | null;
      outline?: { color?: string; width?: number } | null;
      shadow?: boolean;
      text?: string;
      name?: string;
    } = {}
  ): string | null {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return null;
    const spTree = this.spTreeOf(slide.doc);
    if (!spTree) return null;
    this.capture([slide.part], 'Добавление фигуры');

    const id = this.nextShapeId(slide.doc);
    const geom = options.geom ?? 'rect';
    const width = options.w ?? Math.round(this.doc.width * 0.3);
    const height = options.h ?? Math.round(this.doc.height * 0.15);
    const x = options.x ?? Math.round((this.doc.width - width) / 2);
    const y = options.y ?? Math.round((this.doc.height - height) / 2);

    const sp = slide.doc.createElementNS(P_NS, 'p:sp');
    const nvSpPr = slide.doc.createElementNS(P_NS, 'p:nvSpPr');
    const cNvPr = slide.doc.createElementNS(P_NS, 'p:cNvPr');
    cNvPr.setAttribute('id', String(id));
    cNvPr.setAttribute('name', options.name ?? `${geom} ${id}`);
    const cNvSpPr = slide.doc.createElementNS(P_NS, 'p:cNvSpPr');
    const nvPr = slide.doc.createElementNS(P_NS, 'p:nvPr');
    nvSpPr.append(cNvPr, cNvSpPr, nvPr);
    sp.appendChild(nvSpPr);

    const spPr = slide.doc.createElementNS(P_NS, 'p:spPr');
    const xfrm = slide.doc.createElementNS(A_NS, 'a:xfrm');
    const off = slide.doc.createElementNS(A_NS, 'a:off');
    off.setAttribute('x', String(Math.round(x)));
    off.setAttribute('y', String(Math.round(y)));
    const ext = slide.doc.createElementNS(A_NS, 'a:ext');
    ext.setAttribute('cx', String(Math.round(width)));
    ext.setAttribute('cy', String(Math.round(height)));
    xfrm.append(off, ext);
    spPr.appendChild(xfrm);

    const prstGeom = slide.doc.createElementNS(A_NS, 'a:prstGeom');
    prstGeom.setAttribute('prst', geom);
    prstGeom.appendChild(slide.doc.createElementNS(A_NS, 'a:avLst'));
    spPr.appendChild(prstGeom);

    if (options.fill === null) {
      spPr.appendChild(slide.doc.createElementNS(A_NS, 'a:noFill'));
    } else if (options.fill) {
      const fill = slide.doc.createElementNS(A_NS, 'a:solidFill');
      const clr = slide.doc.createElementNS(A_NS, 'a:srgbClr');
      clr.setAttribute('val', normalizeHexColor(options.fill, 'FFFFFF'));
      fill.appendChild(clr);
      spPr.appendChild(fill);
    }
    if (options.outline) {
      const ln = slide.doc.createElementNS(A_NS, 'a:ln');
      if (options.outline.width) ln.setAttribute('w', String(Math.round(options.outline.width * 12700)));
      const fill = slide.doc.createElementNS(A_NS, 'a:solidFill');
      const clr = slide.doc.createElementNS(A_NS, 'a:srgbClr');
      clr.setAttribute('val', normalizeHexColor(options.outline.color ?? '#000000', '000000'));
      fill.appendChild(clr);
      ln.appendChild(fill);
      spPr.appendChild(ln);
    }
    if (options.shadow) {
      const effect = slide.doc.createElementNS(A_NS, 'a:effectLst');
      const shadow = slide.doc.createElementNS(A_NS, 'a:outerShdw');
      shadow.setAttribute('blurRad', '40000');
      shadow.setAttribute('dist', '23000');
      shadow.setAttribute('dir', '5400000');
      const color = slide.doc.createElementNS(A_NS, 'a:srgbClr');
      color.setAttribute('val', '000000');
      const alpha = slide.doc.createElementNS(A_NS, 'a:alpha');
      alpha.setAttribute('val', '35000');
      color.appendChild(alpha);
      shadow.appendChild(color);
      effect.appendChild(shadow);
      spPr.appendChild(effect);
    }
    sp.appendChild(spPr);

    const txBody = slide.doc.createElementNS(P_NS, 'p:txBody');
    const bodyPr = slide.doc.createElementNS(A_NS, 'a:bodyPr');
    bodyPr.setAttribute('wrap', 'square');
    txBody.append(bodyPr, slide.doc.createElementNS(A_NS, 'a:lstStyle'));
    const paragraph = this.buildParagraph(slide.doc, {
      runs: options.text ? [{ text: options.text, size: 18 }] : [],
    });
    txBody.appendChild(paragraph);
    sp.appendChild(txBody);

    spTree.appendChild(sp);
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
    return String(id);
  }

  addTextBox(slideIndex: number, options: { x?: number; y?: number; w?: number; h?: number; text?: string } = {}): string | null {
    return this.addShape(slideIndex, {
      geom: 'rect',
      x: options.x,
      y: options.y,
      w: options.w,
      h: options.h,
      fill: null,
      outline: null,
      text: options.text ?? 'Новый текст',
      name: 'Надпись',
    });
  }

  /* ---------- Изображения ---------- */

  private nextMediaNumber(): number {
    const numbers = new Set<number>();
    for (const name of Object.keys(this.zip.files)) {
      const match = /^ppt\/media\/[^/]*?(\d+)\./.exec(name);
      if (match) numbers.add(Number(match[1]));
    }
    for (const name of this.parts.keys()) {
      const match = /^ppt\/media\/[^/]*?(\d+)\./.exec(name);
      if (match) numbers.add(Number(match[1]));
    }
    for (const name of this.binaryParts.keys()) {
      const match = /^ppt\/media\/[^/]*?(\d+)\./.exec(name);
      if (match) numbers.add(Number(match[1]));
    }
    return (numbers.size ? Math.max(...numbers) : 0) + 1;
  }

  private addContentTypeOverride(part: string, contentType: string): void {
    const xml = this.parts.get('[Content_Types].xml');
    if (!xml) return;
    const normalized = part.startsWith('/') ? part : `/${part}`;
    if (xml.includes(`PartName="${normalized}"`)) return;
    this.parts.set(
      '[Content_Types].xml',
      xml.replace(
        '</Types>',
        `<Override PartName="${normalized}" ContentType="${contentType}"/></Types>`
      )
    );
  }

  private removeContentTypeOverride(part: string): void {
    const xml = this.parts.get('[Content_Types].xml');
    if (!xml) return;
    const normalized = part.startsWith('/') ? part : `/${part}`;
    const pattern = new RegExp(`<Override\\b[^>]*PartName="${normalized.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`);
    this.parts.set('[Content_Types].xml', xml.replace(pattern, ''));
  }

  private ensureContentTypeDefault(extension: string, contentType: string): void {
    const xml = this.parts.get('[Content_Types].xml');
    if (!xml) return;
    if (new RegExp(`<Default\\b[^>]*Extension="${extension}"`, 'i').test(xml)) return;
    this.parts.set(
      '[Content_Types].xml',
      xml.replace(
        '</Types>',
        `<Default Extension="${extension}" ContentType="${contentType}"/></Types>`
      )
    );
  }

  private nextRelationshipId(relsXml: string): string {
    const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]));
    return 'rId' + ((ids.length ? Math.max(...ids) : 0) + 1);
  }

  private nextSlideId(presentationXml: string): number {
    const ids = [...presentationXml.matchAll(/<p:sldId\b[^>]*id="(\d+)"/g)].map((match) => Number(match[1]));
    return (ids.length ? Math.max(...ids) : 255) + 1;
  }

  /** Единая запись медиа: файл + связь + тип содержимого. */
  private writeMedia(slidePart: string, dataUrl: string): { relId: string; part: string } {
    const parsed = parseDataUrl(dataUrl);
    if (!parsed) throw new Error('Ожидается изображение в формате data:image/...;base64,...');
    const number = this.nextMediaNumber();
    const part = `ppt/media/image${number}.${parsed.ext}`;
    this.binaryParts.set(part, parsed.bytes);
    this.ensureContentTypeDefault(parsed.ext, parsed.mime);

    const relsPart = slidePart.replace(/slides\//, 'slides/_rels/') + '.rels';
    const relsXml = this.parts.get(relsPart) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}"/>`;
    const relId = this.nextRelationshipId(relsXml);
    const target = `../media/image${number}.${parsed.ext}`;
    this.parts.set(
      relsPart,
      relsXml.includes('</Relationships>')
        ? relsXml.replace(
            '</Relationships>',
            `<Relationship Id="${relId}" Type="${REL_TYPE.image}" Target="${target}"/></Relationships>`
          )
        : relsXml
    );
    return { relId, part };
  }

  addImage(
    slideIndex: number,
    dataUrl: string,
    options: { x?: number; y?: number; w?: number; h?: number; alt?: string } = {}
  ): string | null {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return null;
    const spTree = this.spTreeOf(slide.doc);
    if (!spTree) return null;
    this.capture([slide.part, 'ppt/slides/_rels/' + slide.part.split('/').pop() + '.rels', '[Content_Types].xml'], 'Вставка изображения');

    const media = this.writeMedia(slide.part, dataUrl);
    const id = this.nextShapeId(slide.doc);

    const slideWidthPx = this.doc.width / 9525;
    let widthPx = Number(options.w) || Math.round(slideWidthPx * 0.6);
    let heightPx = Number(options.h) || Math.round(widthPx * 0.5625);
    widthPx = Math.max(24, Math.min(widthPx, Math.round(slideWidthPx)));
    const xPx = Number.isFinite(Number(options.x)) ? Number(options.x) : Math.round((slideWidthPx - widthPx) / 2);
    const yPx = Number.isFinite(Number(options.y)) ? Number(options.y) : Math.round(this.doc.height / 9525 / 2 - heightPx / 2);

    const pic = slide.doc.createElementNS(P_NS, 'p:pic');
    const nvPicPr = slide.doc.createElementNS(P_NS, 'p:nvPicPr');
    const cNvPr = slide.doc.createElementNS(P_NS, 'p:cNvPr');
    cNvPr.setAttribute('id', String(id));
    cNvPr.setAttribute('name', options.alt || `Изображение ${id}`);
    cNvPr.setAttribute('descr', options.alt || '');
    const cNvPicPr = slide.doc.createElementNS(P_NS, 'p:cNvPicPr');
    cNvPicPr.appendChild(slide.doc.createElementNS(A_NS, 'a:picLocks'));
    const nvPr = slide.doc.createElementNS(P_NS, 'p:nvPr');
    nvPicPr.append(cNvPr, cNvPicPr, nvPr);
    pic.appendChild(nvPicPr);

    const blipFill = slide.doc.createElementNS(P_NS, 'p:blipFill');
    const blip = slide.doc.createElementNS(A_NS, 'a:blip');
    blip.setAttributeNS(R_NS, 'r:embed', media.relId);
    const stretch = slide.doc.createElementNS(A_NS, 'a:stretch');
    stretch.appendChild(slide.doc.createElementNS(A_NS, 'a:fillRect'));
    blipFill.append(blip, stretch);
    pic.appendChild(blipFill);

    const spPr = slide.doc.createElementNS(P_NS, 'p:spPr');
    const xfrm = slide.doc.createElementNS(A_NS, 'a:xfrm');
    const off = slide.doc.createElementNS(A_NS, 'a:off');
    off.setAttribute('x', String(Math.round(xPx * 9525)));
    off.setAttribute('y', String(Math.round(yPx * 9525)));
    const ext = slide.doc.createElementNS(A_NS, 'a:ext');
    ext.setAttribute('cx', String(Math.round(widthPx * 9525)));
    ext.setAttribute('cy', String(Math.round(heightPx * 9525)));
    xfrm.append(off, ext);
    const prstGeom = slide.doc.createElementNS(A_NS, 'a:prstGeom');
    prstGeom.setAttribute('prst', 'rect');
    prstGeom.appendChild(slide.doc.createElementNS(A_NS, 'a:avLst'));
    spPr.append(xfrm, prstGeom);
    pic.appendChild(spPr);

    spTree.appendChild(pic);
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
    void this.mediaUrl(media.part);
    return String(id);
  }

  /* ---------- Таблицы ---------- */

  addTable(
    slideIndex: number,
    options: { rows: number; cols: number; x?: number; y?: number; w?: number; h?: number; data?: string[][] }
  ): string | null {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return null;
    const spTree = this.spTreeOf(slide.doc);
    if (!spTree) return null;
    this.capture([slide.part], 'Вставка таблицы');

    const rows = Math.max(1, Math.min(50, options.rows));
    const cols = Math.max(1, Math.min(20, options.cols));
    const width = options.w ?? Math.round(this.doc.width * 0.7);
    const height = options.h ?? Math.round(this.doc.height * 0.4);
    const x = options.x ?? Math.round((this.doc.width - width) / 2);
    const y = options.y ?? Math.round((this.doc.height - height) / 2);
    const id = this.nextShapeId(slide.doc);
    const data = options.data ?? [];

    const frame = slide.doc.createElementNS(P_NS, 'p:graphicFrame');
    const nvFramePr = slide.doc.createElementNS(P_NS, 'p:nvGraphicFramePr');
    const cNvPr = slide.doc.createElementNS(P_NS, 'p:cNvPr');
    cNvPr.setAttribute('id', String(id));
    cNvPr.setAttribute('name', `Таблица ${id}`);
    const cNvGraphicFramePr = slide.doc.createElementNS(P_NS, 'p:cNvGraphicFramePr');
    const frameLocks = slide.doc.createElementNS(A_NS, 'a:graphicFrameLocks');
    frameLocks.setAttribute('noGrp', '1');
    cNvGraphicFramePr.appendChild(frameLocks);
    nvFramePr.append(cNvPr, cNvGraphicFramePr, slide.doc.createElementNS(P_NS, 'p:nvPr'));
    frame.appendChild(nvFramePr);

    const xfrm = slide.doc.createElementNS(P_NS, 'p:xfrm');
    const off = slide.doc.createElementNS(A_NS, 'a:off');
    off.setAttribute('x', String(Math.round(x)));
    off.setAttribute('y', String(Math.round(y)));
    const ext = slide.doc.createElementNS(A_NS, 'a:ext');
    ext.setAttribute('cx', String(Math.round(width)));
    ext.setAttribute('cy', String(Math.round(height)));
    xfrm.append(off, ext);
    frame.appendChild(xfrm);

    const graphic = slide.doc.createElementNS(A_NS, 'a:graphic');
    const graphicData = slide.doc.createElementNS(A_NS, 'a:graphicData');
    graphicData.setAttribute('uri', TABLE_URI);
    const tbl = slide.doc.createElementNS(A_NS, 'a:tbl');
    const tblPr = slide.doc.createElementNS(A_NS, 'a:tblPr');
    tblPr.setAttribute('firstRow', '1');
    tblPr.setAttribute('bandRow', '1');
    tbl.appendChild(tblPr);

    const tblGrid = slide.doc.createElementNS(A_NS, 'a:tblGrid');
    const columnWidth = Math.round(width / cols);
    for (let col = 0; col < cols; col++) {
      const gridCol = slide.doc.createElementNS(A_NS, 'a:gridCol');
      gridCol.setAttribute('w', String(columnWidth));
      tblGrid.appendChild(gridCol);
    }
    tbl.appendChild(tblGrid);

    const rowHeight = Math.round(height / rows);
    for (let row = 0; row < rows; row++) {
      const tr = slide.doc.createElementNS(A_NS, 'a:tr');
      tr.setAttribute('h', String(rowHeight));
      for (let col = 0; col < cols; col++) {
        const tc = slide.doc.createElementNS(A_NS, 'a:tc');
        const txBody = slide.doc.createElementNS(A_NS, 'a:txBody');
        const bodyPr = slide.doc.createElementNS(A_NS, 'a:bodyPr');
        txBody.append(bodyPr, slide.doc.createElementNS(A_NS, 'a:lstStyle'));
        txBody.appendChild(
          this.buildParagraph(slide.doc, {
            runs: data[row]?.[col] ? [{ text: data[row][col], size: 14 }] : [],
          })
        );
        const tcPr = slide.doc.createElementNS(A_NS, 'a:tcPr');
        tcPr.setAttribute('anchor', 'ctr');
        tc.append(txBody, tcPr);
        tr.appendChild(tc);
      }
      tbl.appendChild(tr);
    }
    graphicData.appendChild(tbl);
    graphic.appendChild(graphicData);
    frame.appendChild(graphic);

    spTree.appendChild(frame);
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
    return String(id);
  }

  setTableData(slideIndex: number, shapeId: string, data: string[][]): void {
    const slide = this.slideDoc(slideIndex);
    if (!slide) return;
    const node = this.findShapeNode(slide.doc, shapeId);
    if (!node) return;
    const tbl = firstByLocal(node, 'tbl');
    if (!tbl) return;
    this.capture([slide.part], 'Правка таблицы');
    const rows = childrenByLocal(tbl, 'tr');
    rows.forEach((tr, rowIndex) => {
      const cells = childrenByLocal(tr, 'tc');
      cells.forEach((tc, colIndex) => {
        const value = data[rowIndex]?.[colIndex];
        if (value === undefined) return;
        let txBody = childByLocal(tc, 'txBody');
        if (!txBody) {
          txBody = slide.doc.createElementNS(A_NS, 'a:txBody');
          tc.insertBefore(txBody, tc.firstChild);
        }
        const body: PptxTextBody = { paragraphs: [{ runs: value ? [{ text: value, size: 14 }] : [] }] };
        this.writeTextBody(slide.doc, txBody, body);
      });
    });
    this.commit(slide.part, slide.doc);
    this.doc = this.parseModel();
  }

  /* ---------- Заметки ---------- */

  setNotes(slideIndex: number, text: string): void {
    const slide = this.doc.slides[slideIndex];
    if (!slide) return;
    const notesPart = this.notesPartFor(slide.part);
    const existing = this.parts.get(notesPart);
    this.capture([notesPart], 'Заметки докладчика');

    if (!existing) {
      // Создаём минимальную часть заметок и связываем её со слайдом.
      const notesMasterPart = this.findNotesMasterPart();
      if (!notesMasterPart) {
        // Без мастера заметок PowerPoint откажется открывать файл — тогда только модель.
        slide.notes = text;
        return;
      }
      const xml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        `<p:notes xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
        '<p:cSld><p:spTree>' +
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>';
      this.parts.set(notesPart, xml);
      const relsPart = notesPart.replace(/notesSlides\//, 'notesSlides/_rels/') + '.rels';
      const slideRelsPart = 'ppt/slides/_rels/' + slide.part.split('/').pop() + '.rels';
      const slideRels = this.parts.get(slideRelsPart) ?? '';
      const relId = this.nextRelationshipId(slideRels);
      this.parts.set(
        slideRelsPart,
        slideRels.replace(
          '</Relationships>',
          `<Relationship Id="${relId}" Type="${REL_TYPE.notesSlide}" Target="../notesSlides/${notesPart.split('/').pop()}"/></Relationships>`
        )
      );
      this.parts.set(
        relsPart,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          `<Relationships xmlns="${REL_NS}">` +
          `<Relationship Id="rId1" Type="${REL_TYPE.notesMaster}" Target="../notesMasters/${notesMasterPart.split('/').pop()}"/></Relationships>`
      );
      this.addContentTypeOverride(`/${notesPart}`, CONTENT_TYPE.notesSlide);
    }

    const xml = this.parts.get(notesPart);
    if (xml) {
      const doc = parseXml(xml);
      const spTree = firstByLocal(doc, 'spTree');
      if (spTree) {
        // Удаляем прежнюю текстовую фигуру-плейсхолдер тела заметок.
        for (const sp of childrenByLocal(spTree, 'sp')) {
          const ph = firstByLocal(sp, 'ph');
          if (attr(ph, 'type') === 'body' || attr(ph, 'idx') === '1' || firstByLocal(sp, 'txBody')) {
            spTree.removeChild(sp);
          }
        }
        const sp = doc.createElementNS(P_NS, 'p:sp');
        const nvSpPr = doc.createElementNS(P_NS, 'p:nvSpPr');
        const cNvPr = doc.createElementNS(P_NS, 'p:cNvPr');
        cNvPr.setAttribute('id', '2');
        cNvPr.setAttribute('name', 'Текст заметок 1');
        const nvPr = doc.createElementNS(P_NS, 'p:nvPr');
        const ph = doc.createElementNS(P_NS, 'p:ph');
        ph.setAttribute('type', 'body');
        ph.setAttribute('idx', '1');
        nvPr.appendChild(ph);
        nvSpPr.append(cNvPr, doc.createElementNS(P_NS, 'p:cNvSpPr'), nvPr);
        sp.appendChild(nvSpPr);
        sp.appendChild(doc.createElementNS(P_NS, 'p:spPr'));
        const txBody = doc.createElementNS(P_NS, 'p:txBody');
        const bodyPr = doc.createElementNS(A_NS, 'a:bodyPr');
        txBody.append(bodyPr, doc.createElementNS(A_NS, 'a:lstStyle'));
        for (const line of (text || '').split('\n')) {
          txBody.appendChild(
            this.buildParagraph(doc, { runs: line ? [{ text: line, size: 12 }] : [] })
          );
        }
        sp.appendChild(txBody);
        spTree.appendChild(sp);
      }
      this.parts.set(notesPart, serializeXml(doc));
    }
    this.doc = this.parseModel();
  }

  private findNotesMasterPart(): string {
    for (const name of Object.keys(this.zip.files)) {
      if (/^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(name)) {
        return name;
      }
    }
    return '';
  }

  /* ---------- Поиск и замена ---------- */

  findReplace(query: string, replacement: string, options: { includeNotes?: boolean } = {}): { count: number; slides: number[] } {
    if (!query) return { count: 0, slides: [] };
    const touched = new Set<number>();
    let count = 0;
    for (const slide of this.doc.slides) {
      const slideDoc = this.slideDoc(slide.index);
      if (!slideDoc) continue;
      let changed = false;
      for (const textNode of allByLocal(slideDoc.doc, 't')) {
        const value = textNode.textContent ?? '';
        if (!value.includes(query)) continue;
        textNode.textContent = value.split(query).join(replacement);
        count += value.split(query).length - 1;
        changed = true;
      }
      if (changed) {
        this.capture([slideDoc.part], 'Поиск и замена');
        this.commit(slideDoc.part, slideDoc.doc);
        touched.add(slide.index);
      }
      if (options.includeNotes && slide.notes.includes(query)) {
        const notesPart = this.notesPartFor(slide.part);
        const xml = this.parts.get(notesPart);
        if (xml) {
          this.capture([notesPart], 'Поиск и замена в заметках');
          const doc = parseXml(xml);
          for (const textNode of allByLocal(doc, 't')) {
            const value = textNode.textContent ?? '';
            if (!value.includes(query)) continue;
            textNode.textContent = value.split(query).join(replacement);
            count += value.split(query).length - 1;
          }
          this.parts.set(notesPart, serializeXml(doc));
          touched.add(slide.index);
        }
      }
    }
    if (touched.size > 0) this.doc = this.parseModel();
    return { count, slides: [...touched].sort((a, b) => a - b) };
  }

  /* ---------- Совместимость со старым API ---------- */

  /**
   * Проекция на модель, которую ждут `services/pptBridge.ts` и AI-инструменты:
   * только текстовые фигуры и их абзацы.
   */
  toLegacyModel(): {
    width: number;
    height: number;
    layouts: { index: number; name: string; part: string }[];
    slides: {
      index: number;
      part: string;
      notes: string;
      background: string;
      layoutPart: string;
      shapes: {
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
      }[];
    }[];
  } {
    return {
      width: this.doc.width,
      height: this.doc.height,
      layouts: this.doc.layouts,
      slides: this.doc.slides.map((slide) => ({
        index: slide.index,
        part: slide.part,
        notes: slide.notes,
        background: `#${slide.background}`,
        layoutPart: slide.layoutPart ?? '',
        shapes: slide.shapes
          .filter((shape) => shape.text && shape.text.paragraphs.some((p) => p.runs.length > 0))
          .map((shape) => {
            const paragraphs = (shape.text?.paragraphs ?? []).map((paragraph) =>
              paragraph.runs.map((run) => run.text).join('')
            );
            const firstRun = shape.text?.paragraphs[0]?.runs[0];
            return {
              shapeId: shape.id,
              name: shape.name,
              x: shape.x,
              y: shape.y,
              w: shape.w,
              h: shape.h,
              paragraphs,
              text: paragraphs.join('\n').trim(),
              fontSize: firstRun?.size,
              bold: firstRun?.bold,
              color: firstRun?.color ? `#${firstRun.color}` : undefined,
              isTitle: Boolean(shape.isTitle),
            };
          }),
      })),
    };
  }

  /** Публичный метод предзагрузки XML — оставлен для совместимости вызовов. */
  async ready(): Promise<this> {
    return this;
  }
}

/** Пустая модель: нужна до первой загрузки частей пакета. */
function emptyDoc(): PptxDocModel {
  return {
    width: 12192000,
    height: 6858000,
    theme: {
      part: 'ppt/theme/theme1.xml',
      name: 'Office',
      colors: {
        dk1: '000000',
        lt1: 'FFFFFF',
        dk2: '44546A',
        lt2: 'E7E6E6',
        accent1: '4472C4',
        accent2: 'ED7D31',
        accent3: 'A5A5A5',
        accent4: 'FFC000',
        accent5: '5B9BD5',
        accent6: '70AD47',
        hlink: '0563C1',
        folHlink: '954F72',
      },
      fonts: { id: 'calibri', label: 'Calibri Light / Calibri', major: 'Calibri Light', minor: 'Calibri' },
    },
    layouts: [],
    slides: [],
  };
}

/** Разбор data URL изображения. */
function parseDataUrl(dataUrl: string): { mime: string; ext: string; bytes: Uint8Array } | null {
  const match = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(dataUrl || '');
  if (!match) return null;
  const mime = match[1] || 'image/png';
  const isBase64 = Boolean(match[2]);
  const payload = match[3] ?? '';
  const ext =
    mime === 'image/jpeg'
      ? 'jpeg'
      : mime === 'image/svg+xml'
        ? 'svg'
        : mime.split('/')[1]?.replace('+xml', '') || 'png';
  if (!isBase64) return null;
  try {
    return { mime, ext, bytes: base64ToBytes(payload) };
  } catch {
    return null;
  }
}

/** Открытие документа с предзагрузкой XML-частей. */
export async function openPptxDocument(base64: string): Promise<PptxDocument> {
  const doc = await PptxDocument.open(base64);
  return doc.ready();
}

export { parseDataUrl };

/** Упрощённая модель презентации для инструментов 1.1. */
export type PptxLegacyModel = ReturnType<PptxDocument['toLegacyModel']>;

