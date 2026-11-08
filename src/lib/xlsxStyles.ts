/**
 * Нативная запись стилей ячеек в xlsx.
 *
 * SheetJS CE не умеет писать заливку, шрифт, границы и числовые форматы,
 * поэтому после XLSX.write патчим xl/styles.xml и атрибуты s у ячеек.
 * WorkSpaceStyles (Custprops) остаётся для обратной совместимости, но Excel
 * теперь видит обычные стили.
 */
import JSZip from 'jszip';
import {
  numberFormatCode,
  type CellStyle,
  type WorkbookMeta,
} from './sheetFormatting';

const COLOR_INDEXED = '<bgColor indexed="64"/>';

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function argb(value: string | undefined): string {
  if (!value) return '';
  const hex = value.trim().replace(/^#/, '');
  if (/^[0-9a-f]{6}$/i.test(hex)) return `FF${hex.toUpperCase()}`;
  if (/^[0-9a-f]{8}$/i.test(hex)) return hex.toUpperCase();
  return '';
}

function alignXml(style: CellStyle): string {
  const attrs: string[] = [];
  if (style.align) attrs.push(`horizontal="${style.align}"`);
  if (style.valign) attrs.push(`vertical="${style.valign}"`);
  if (style.wrap) attrs.push('wrapText="1"');
  return attrs.length > 0 ? `<alignment ${attrs.join(' ')}/>` : '';
}

function fontXml(style: CellStyle): string {
  const parts = ['<font>'];
  if (style.bold) parts.push('<b/>');
  if (style.italic) parts.push('<i/>');
  if (style.underline) parts.push('<u/>');
  if (style.strike) parts.push('<strike/>');
  if (style.fontSize) {
    parts.push(`<sz val="${Math.max(1, Math.round(style.fontSize * 0.75 * 10) / 10)}"/>`);
  }
  const color = argb(style.color);
  if (color) parts.push(`<color rgb="${color}"/>`);
  parts.push('<name val="Calibri"/>');
  parts.push('</font>');
  return parts.join('');
}

function fillXml(style: CellStyle): string {
  const color = argb(style.fill);
  if (!color) return '';
  return (
    '<fill><patternFill patternType="solid">' +
    `<fgColor rgb="${color}"/>${COLOR_INDEXED}` +
    '</patternFill></fill>'
  );
}

function borderXml(style: CellStyle): string {
  const borders = style.borders;
  if (!borders || (!borders.top && !borders.right && !borders.bottom && !borders.left)) {
    return '';
  }
  const side = (name: string, on?: boolean) =>
    on ? `<${name} style="thin"><color rgb="FF000000"/></${name}>` : `<${name}/>`;
  return (
    '<border>' +
    side('left', borders.left) +
    side('right', borders.right) +
    side('top', borders.top) +
    side('bottom', borders.bottom) +
    '<diagonal/>' +
    '</border>'
  );
}

interface Section {
  inner: string;
  openTag: string;
  full: string;
}

function section(xml: string, tag: string): Section | null {
  const match = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)<\\/${tag}>`).exec(xml);
  if (!match) return null;
  return {
    inner: match[1],
    openTag: new RegExp(`^<${tag}\\b[^>]*>`).exec(match[0])?.[0] ?? `<${tag}>`,
    full: match[0],
  };
}

function countEntries(inner: string, tag: string): number {
  return (inner.match(new RegExp(`<${tag}\\b`, 'g')) ?? []).length;
}

function replaceCount(openTag: string, count: number): string {
  if (/\bcount="\d+"/.test(openTag)) {
    return openTag.replace(/\bcount="\d+"/, `count="${count}"`);
  }
  return openTag.replace(/>$/, ` count="${count}">`);
}

class StyleRegistry {
  private readonly fonts: string[] = [];
  private readonly fills: string[] = [];
  private readonly borders: string[] = [];
  private readonly xfs: string[] = [];
  private readonly numFmts: Array<{ id: number; code: string }> = [];

  private readonly fontIndex = new Map<string, number>();
  private readonly fillIndex = new Map<string, number>();
  private readonly borderIndex = new Map<string, number>();
  private readonly xfIndex = new Map<string, number>();
  private readonly numFmtIndex = new Map<string, number>();

  readonly baseFonts: number;
  readonly baseFills: number;
  readonly baseBorders: number;
  readonly baseXfs: number;
  private nextNumFmtId: number;

  constructor(stylesXml: string) {
    const fonts = section(stylesXml, 'fonts');
    const fills = section(stylesXml, 'fills');
    const borders = section(stylesXml, 'borders');
    const cellXfs = section(stylesXml, 'cellXfs');
    this.baseFonts = fonts ? countEntries(fonts.inner, 'font') : 0;
    this.baseFills = fills ? countEntries(fills.inner, 'fill') : 0;
    this.baseBorders = borders ? countEntries(borders.inner, 'border') : 0;
    this.baseXfs = cellXfs ? countEntries(cellXfs.inner, 'xf') : 0;

    let maxNumFmt = 163;
    for (const match of stylesXml.matchAll(/<numFmt\b[^>]*numFmtId="(\d+)"/g)) {
      maxNumFmt = Math.max(maxNumFmt, Number(match[1]));
    }
    this.nextNumFmtId = maxNumFmt + 1;
  }

  private fontId(style: CellStyle): number {
    const xml = fontXml(style);
    const hasFont =
      style.bold || style.italic || style.underline || style.strike || style.color || style.fontSize;
    if (!hasFont || !xml) return 0;
    const existing = this.fontIndex.get(xml);
    if (existing !== undefined) return existing;
    const id = this.baseFonts + this.fonts.length;
    this.fontIndex.set(xml, id);
    this.fonts.push(xml);
    return id;
  }

  private fillId(style: CellStyle): number {
    const xml = fillXml(style);
    if (!xml) return 0;
    const existing = this.fillIndex.get(xml);
    if (existing !== undefined) return existing;
    const id = this.baseFills + this.fills.length;
    this.fillIndex.set(xml, id);
    this.fills.push(xml);
    return id;
  }

  private borderId(style: CellStyle): number {
    const xml = borderXml(style);
    if (!xml) return 0;
    const existing = this.borderIndex.get(xml);
    if (existing !== undefined) return existing;
    const id = this.baseBorders + this.borders.length;
    this.borderIndex.set(xml, id);
    this.borders.push(xml);
    return id;
  }

  private numFmtId(style: CellStyle): number {
    const code = numberFormatCode(style.numberFormat);
    if (!code) return 0;
    const existing = this.numFmtIndex.get(code);
    if (existing !== undefined) return existing;
    const id = this.nextNumFmtId++;
    this.numFmtIndex.set(code, id);
    this.numFmts.push({ id, code });
    return id;
  }

  register(style: CellStyle): number {
    const font = this.fontId(style);
    const fill = this.fillId(style);
    const border = this.borderId(style);
    const numFmt = this.numFmtId(style);
    const alignment = alignXml(style);
    const key = `${numFmt}|${font}|${fill}|${border}|${alignment}`;
    const existing = this.xfIndex.get(key);
    if (existing !== undefined) return existing;

    const attrs = [
      `numFmtId="${numFmt}"`,
      `fontId="${font}"`,
      `fillId="${fill}"`,
      `borderId="${border}"`,
      'xfId="0"',
      numFmt ? 'applyNumberFormat="1"' : '',
      font ? 'applyFont="1"' : '',
      fill ? 'applyFill="1"' : '',
      border ? 'applyBorder="1"' : '',
      alignment ? 'applyAlignment="1"' : '',
    ].filter(Boolean);
    const xml = alignment
      ? `<xf ${attrs.join(' ')}>${alignment}</xf>`
      : `<xf ${attrs.join(' ')}/>`;
    const id = this.baseXfs + this.xfs.length;
    this.xfIndex.set(key, id);
    this.xfs.push(xml);
    return id;
  }

  apply(stylesXml: string): string {
    let xml = stylesXml;
    if (this.numFmts.length > 0) {
      const entries = this.numFmts
        .map(
          (item) =>
            `<numFmt numFmtId="${item.id}" formatCode="${escapeXml(item.code)}"/>`
        )
        .join('');
      if (section(xml, 'numFmts')) {
        xml = xml.replace(
          /(<numFmts\b[^>]*>)([\s\S]*?)(<\/numFmts>)/,
          (_match, open: string, inner: string, close: string) =>
            `${replaceCount(open, countEntries(inner, 'numFmt') + this.numFmts.length)}${inner}${entries}${close}`
        );
      } else {
        xml = xml.replace(
          /<fonts\b/,
          `<numFmts count="${this.numFmts.length}">${entries}</numFmts><fonts`
        );
      }
    }

    const append = (tag: string, entries: string[]) => {
      if (entries.length === 0) return;
      const current = section(xml, tag);
      if (!current) return;
      xml = xml.replace(
        current.full,
        `${replaceCount(current.openTag, countEntries(current.inner, tag === 'cellXfs' ? 'xf' : tag.slice(0, -1)) + entries.length)}${current.inner}${entries.join('')}</${tag}>`
      );
    };

    append('fonts', this.fonts);
    append('fills', this.fills);
    append('borders', this.borders);
    append('cellXfs', this.xfs);
    return xml;
  }
}

interface SheetEntry {
  name: string;
  path: string;
}

function decodeEntities(value: string): string {
  return value
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

async function readSheetMap(zip: JSZip): Promise<SheetEntry[]> {
  const workbookFile = zip.file('xl/workbook.xml');
  const relsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (!workbookFile || !relsFile) return [];
  const workbookXml = await workbookFile.async('string');
  const relsXml = await relsFile.async('string');

  const targets = new Map<string, string>();
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const tag = match[0];
    const id = /\bId="([^"]+)"/.exec(tag)?.[1];
    const target = /\bTarget="([^"]+)"/.exec(tag)?.[1];
    if (id && target) targets.set(id, target);
  }

  const result: SheetEntry[] = [];
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*\/?>/g)) {
    const tag = match[0];
    const name = /\bname="([^"]+)"/.exec(tag)?.[1];
    const rid = /\br:id="([^"]+)"/.exec(tag)?.[1];
    if (!name || !rid) continue;
    const target = targets.get(rid);
    if (!target) continue;
    const normalized = target.startsWith('/')
      ? target.slice(1)
      : `xl/${target.replace(/^\.\//, '')}`;
    result.push({ name: decodeEntities(name), path: normalized.replace(/\/\.\//g, '/') });
  }
  return result;
}

function columnIndex(letters: string): number {
  let value = 0;
  for (const char of letters.toUpperCase()) {
    value = value * 26 + (char.charCodeAt(0) - 64);
  }
  return value - 1;
}

function applyCellStyle(xml: string, address: string, styleId: number): string {
  const existing = new RegExp(`<c r="${address}"([^>]*?)(/?)>`).exec(xml);
  if (existing) {
    const attrs = existing[1].replace(/\s+s="\d+"/, '');
    return xml.replace(
      existing[0],
      `<c r="${address}" s="${styleId}"${attrs}${existing[2]}>`
    );
  }

  const rowNumber = Number(/\d+/.exec(address)?.[0] ?? 0);
  const letters = /^[A-Z]+/.exec(address)?.[0] ?? '';
  if (!rowNumber || !letters) return xml;
  const cell = `<c r="${address}" s="${styleId}"/>`;
  const rowRe = new RegExp(`<row r="${rowNumber}"([^>]*)>([\\s\\S]*?)</row>`);
  const row = rowRe.exec(xml);
  if (!row) return xml;

  const targetCol = columnIndex(letters);
  let insertion = row[2].length;
  for (const match of row[2].matchAll(/<c r="([A-Z]+)\d+"/g)) {
    if (columnIndex(match[1]) > targetCol) {
      insertion = match.index ?? insertion;
      break;
    }
  }
  const inner = row[2].slice(0, insertion) + cell + row[2].slice(insertion);
  return xml.replace(row[0], `<row r="${rowNumber}"${row[1]}>${inner}</row>`);
}

export async function injectStylesIntoXlsx(
  base64: string,
  meta: WorkbookMeta
): Promise<string> {
  const hasStyles = Object.values(meta).some(
    (sheet) => Object.keys(sheet.cells).length > 0
  );
  if (!hasStyles) return base64;

  const zip = await JSZip.loadAsync(base64, { base64: true });
  const stylesFile = zip.file('xl/styles.xml');
  if (!stylesFile) return base64;

  const stylesXml = await stylesFile.async('string');
  const registry = new StyleRegistry(stylesXml);
  const sheets = await readSheetMap(zip);

  for (const entry of sheets) {
    const sheetMeta = meta[entry.name];
    if (!sheetMeta || Object.keys(sheetMeta.cells).length === 0) continue;
    const file = zip.file(entry.path);
    if (!file) continue;
    let xml = await file.async('string');
    for (const [address, style] of Object.entries(sheetMeta.cells)) {
      xml = applyCellStyle(xml, address, registry.register(style));
    }
    zip.file(entry.path, xml);
  }

  zip.file('xl/styles.xml', registry.apply(stylesXml));
  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}
