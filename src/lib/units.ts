/**
 * Единицы OOXML ↔ CSS и геометрия листа.
 *
 * Экран, предпросмотр и печать должны совпадать, поэтому все три считаются
 * из одной точки: 1 мм = 96/25.4 CSS px (96 dpi — как у Word при масштабе
 * 100 %), twips Word (1/20 pt) и пункты переводятся без округления до целых
 * пикселей: округление каждого поля отдельно сдвигало полосу набора на
 * десятые доли пикселя и ломало совпадение переносов.
 */
import type { PageSetup } from '../types';

export const CSS_PX_PER_INCH = 96;
export const MM_PER_INCH = 25.4;
export const CSS_PX_PER_MM = CSS_PX_PER_INCH / MM_PER_INCH;
export const TWIPS_PER_MM = 1440 / MM_PER_INCH;
export const PT_PER_MM = 72 / MM_PER_INCH;

/** Промежуток между листами на экране (Word рисует его серым фоном). */
export const PAGE_GAP_PX = 24;
/** Расстояние колонтитула от края: 1,25 см — значение Word по умолчанию. */
export const HEADER_DISTANCE_MM = 12.5;
export const FOOTER_DISTANCE_MM = 12.5;

export type PaperName = 'A4' | 'A3' | 'A5' | 'A6' | 'Letter' | 'Legal' | 'Tabloid' | 'custom';

export interface PaperSize {
  widthMm: number;
  heightMm: number;
}

export const PAPER_SIZES: Record<Exclude<PaperName, 'custom'>, PaperSize> = {
  A4: { widthMm: 210, heightMm: 297 },
  A3: { widthMm: 297, heightMm: 420 },
  A5: { widthMm: 148, heightMm: 210 },
  A6: { widthMm: 105, heightMm: 148 },
  Letter: { widthMm: 215.9, heightMm: 279.4 },
  Legal: { widthMm: 215.9, heightMm: 355.6 },
  Tabloid: { widthMm: 279.4, heightMm: 431.8 },
};

/**
 * Параметры страницы нового документа: A4 и поля шаблона Word для ru
 * (3 / 1,5 / 2 / 2 см). Раньше здесь были 2 см со всех сторон — документ
 * «из ниоткуда» выглядел не как в Word.
 */
export const WORD_DEFAULT_PAGE: PageSetup = {
  widthMm: PAPER_SIZES.A4.widthMm,
  heightMm: PAPER_SIZES.A4.heightMm,
  marginTopMm: 20,
  marginRightMm: 15,
  marginBottomMm: 20,
  marginLeftMm: 30,
  headerMm: HEADER_DISTANCE_MM,
  footerMm: FOOTER_DISTANCE_MM,
  fontFamily: 'Calibri',
  fontSizePt: 11,
  lineSpacing: 1,
  lineSpacingRule: 'auto',
};

export function mmToPx(mm: number): number {
  return mm * CSS_PX_PER_MM;
}

export function pxToMm(px: number): number {
  return px / CSS_PX_PER_MM;
}

export function ptToPx(pt: number): number {
  return (pt * CSS_PX_PER_INCH) / 72;
}

export function pxToPt(px: number): number {
  return (px * 72) / CSS_PX_PER_INCH;
}

export function mmToTwips(mm: number): number {
  return Math.round(mm * TWIPS_PER_MM);
}

export function twipsToMm(twips: number): number {
  return twips / TWIPS_PER_MM;
}

export function twipsToPx(twips: number): number {
  return mmToPx(twipsToMm(twips));
}

export function pxToTwips(px: number): number {
  return mmToTwips(pxToMm(px));
}

/** Миллиметры для CSS с точностью до тысячных: «210mm», «112.53mm». */
export function mmString(mm: number): string {
  const rounded = Math.round(mm * 1000) / 1000;
  return `${rounded}mm`;
}

export function pxToMmString(px: number): string {
  return mmString(pxToMm(px));
}

/** Микроны нужны Electron для нестандартного размера печатной страницы. */
export function mmToMicrons(mm: number): number {
  return Math.round(mm * 1000);
}

export interface PageGeometry {
  paper: PaperName;
  /** false — альбомная ориентация (ширина больше высоты). */
  portrait: boolean;
  widthMm: number;
  heightMm: number;
  marginTopMm: number;
  marginRightMm: number;
  marginBottomMm: number;
  marginLeftMm: number;
  headerMm: number;
  footerMm: number;
  gutterMm: number;
  widthPx: number;
  heightPx: number;
  marginTopPx: number;
  marginRightPx: number;
  marginBottomPx: number;
  marginLeftPx: number;
  headerPx: number;
  footerPx: number;
  gutterPx: number;
  /** Верх/низ полосы набора: с учётом колонтитула, если он выше поля. */
  bodyTopPx: number;
  bodyBottomPx: number;
  contentWidthPx: number;
  contentHeightPx: number;
}

export interface PageGeometryExtras {
  /** Высота уже свёрстанного верхнего колонтитула: если он выше поля, тело уезжает вниз. */
  headerContentPx?: number;
  footerContentPx?: number;
}

function positive(value: unknown, fallback: number): number {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : fallback;
}

function detectPaper(widthMm: number, heightMm: number): PaperName {
  const short = Math.min(widthMm, heightMm);
  const long = Math.max(widthMm, heightMm);
  for (const [name, size] of Object.entries(PAPER_SIZES)) {
    const sizeShort = Math.min(size.widthMm, size.heightMm);
    const sizeLong = Math.max(size.widthMm, size.heightMm);
    if (Math.abs(short - sizeShort) < 0.6 && Math.abs(long - sizeLong) < 0.6) {
      return name as PaperName;
    }
  }
  return 'custom';
}

/** Полная геометрия листа из параметров .docx (незаданное — как в Word). */
export function resolvePageGeometry(page?: PageSetup, extras?: PageGeometryExtras): PageGeometry {
  const widthMm = positive(page?.widthMm, WORD_DEFAULT_PAGE.widthMm ?? PAPER_SIZES.A4.widthMm);
  const heightMm = positive(page?.heightMm, WORD_DEFAULT_PAGE.heightMm ?? PAPER_SIZES.A4.heightMm);
  const marginTopMm = positive(page?.marginTopMm, 20);
  const marginRightMm = positive(page?.marginRightMm, 15);
  const marginBottomMm = positive(page?.marginBottomMm, 20);
  const marginLeftMm = positive(page?.marginLeftMm, 30);
  const headerMm = positive(page?.headerMm, HEADER_DISTANCE_MM);
  const footerMm = positive(page?.footerMm, FOOTER_DISTANCE_MM);
  const gutterMm = Math.max(0, Number(page?.gutterMm) || 0);

  const widthPx = mmToPx(widthMm);
  const heightPx = mmToPx(heightMm);
  const marginTopPx = mmToPx(marginTopMm);
  const marginRightPx = mmToPx(marginRightMm);
  const marginBottomPx = mmToPx(marginBottomMm);
  const marginLeftPx = mmToPx(marginLeftMm);
  const headerPx = mmToPx(headerMm);
  const footerPx = mmToPx(footerMm);
  const gutterPx = mmToPx(gutterMm);

  // Word поднимает текст вниз, если колонтитул выше отведённого ему поля.
  const headerContentPx = Math.max(0, Number(extras?.headerContentPx) || 0);
  const footerContentPx = Math.max(0, Number(extras?.footerContentPx) || 0);
  const bodyTopPx = Math.max(marginTopPx, headerPx + headerContentPx);
  const bodyBottomPx = Math.max(marginBottomPx, footerPx + footerContentPx);

  return {
    paper: detectPaper(widthMm, heightMm),
    portrait: heightMm >= widthMm,
    widthMm,
    heightMm,
    marginTopMm,
    marginRightMm,
    marginBottomMm,
    marginLeftMm,
    headerMm,
    footerMm,
    gutterMm,
    widthPx,
    heightPx,
    marginTopPx,
    marginRightPx,
    marginBottomPx,
    marginLeftPx,
    headerPx,
    footerPx,
    gutterPx,
    bodyTopPx,
    bodyBottomPx,
    contentWidthPx: Math.max(40, widthPx - marginLeftPx - marginRightPx - gutterPx),
    contentHeightPx: Math.max(40, heightPx - bodyTopPx - bodyBottomPx),
  };
}

/** `@page` с нулевыми полями: страницу целиком рисуем сами. */
export function printPageSizeCss(geometry: PageGeometry): string {
  return `@page{size:${mmString(geometry.widthMm)} ${mmString(geometry.heightMm)};margin:0}`;
}

/**
 * Печать «потоком»: поля уходят в `@page`, поэтому постраничность остаётся
 * прежней — маркеры разрыва из листа превращаются в принудительные разрывы
 * печатного движка (высоту они на экране занимают, на бумаге нет).
 */
export function printFlowCss(geometry: PageGeometry): string {
  return [
    `@page{size:${mmString(geometry.widthMm)} ${mmString(geometry.heightMm)};margin:${mmString(
      geometry.marginTopMm
    )} ${mmString(geometry.marginRightMm)} ${mmString(geometry.marginBottomMm)} ${mmString(
      geometry.marginLeftMm
    )}}`,
    'html,body{margin:0;padding:0;background:#fff}',
    '[data-page-gap]{display:block;height:0!important;margin:0!important;padding:0!important;' +
      'border:0!important;break-before:page;page-break-before:always}',
    '.ws-print-flow{display:flex;flex-direction:column;width:100%}',
  ].join('\n');
}
