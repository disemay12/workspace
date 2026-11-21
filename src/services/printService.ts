/**
 * Печать любого документа WorkSpace одним путём.
 *
 * Порядок выбора:
 * 1. живой провайдер редактора (`printBridge`): печатается то, что видно,
 *    включая несохранённые правки;
 * 2. сборка по содержимому вкладки (`tab.content`);
 * 3. файл с диска (печать из проводника, без открытия вкладки).
 *
 * Лист Word печатается теми же страницами, что нарисованы на экране: поток
 * размечается пагинатором и режется на страницы, поэтому принтер не
 * переносит строки по своему усмотрению.
 */
import * as XLSX from 'xlsx';
import type { DocxHeaderFooterHtml } from '../lib/docx';
import { docxBase64ToHtml } from '../lib/docx';
import { fillHeaderFooterFields, pickHeaderFooter } from '../lib/docxFields';
import { markdownToDocumentHtml } from '../lib/aiDocument';
import { plainTextToHtml } from '../lib/legacyDoc';
import { paginateToPages } from '../lib/pageLayout';
import { buildPptxPrintHtml, pptxPageMicrons } from '../lib/pptxPrint';
import {
  WORD_DEFAULT_PAGE,
  mmToMicrons,
  ptToPx,
  pxToMm,
  resolvePageGeometry,
  type PageGeometry,
} from '../lib/units';
import { baseName, formatBytes, extName } from '../lib/utils';
import type { FileType, OpenTab, PageSetup } from '../types';
import { getPrintProvider, type PrintDocument } from './printBridge';
import { readFileSource } from './documentConversion';
import { fileSystemService, getFileType } from './fileSystemService';

export interface PrintResult {
  success: boolean;
  canceled?: boolean;
  error?: string;
  /** PDF в base64, если запрашивался предпросмотр. */
  pdfBase64?: string;
}

export type PrintTarget = { kind: 'tab'; tab: OpenTab } | { kind: 'file'; path: string };

/** Что уже реализовано в редакторах как живой провайдер печати. */
export { type PrintDocument } from './printBridge';

/* ------------------------------------------------------------------ общее */

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Типографика листа в печати. Дублирует `.word-sheet` из index.css:
 * печатный документ живёт в отдельном окне и должен быть самодостаточным.
 */
const WORD_TYPOGRAPHY = `
.word-sheet{display:flex;flex-direction:column;align-items:stretch;background:#fff;color:#000;
  font-family:Calibri,'Segoe UI',Arial,sans-serif;font-size:11pt;line-height:normal;text-align:left;
  tab-size:35.43px;overflow-wrap:break-word}
.word-sheet > *{flex:0 0 auto;max-width:100%}
.word-sheet p{margin:0 0 8pt}
.word-sheet h1{font-size:16pt;font-weight:700;color:#1f3864;margin:12pt 0 3pt}
.word-sheet h2{font-size:13pt;font-weight:700;color:#1f3864;margin:12pt 0 3pt}
.word-sheet h3{font-size:12pt;font-weight:700;color:#1f3864;margin:12pt 0 3pt}
.word-sheet h4,.word-sheet h5,.word-sheet h6{font-size:11pt;font-weight:700;font-style:italic;color:#1f3864;margin:10pt 0 3pt}
.word-sheet ul,.word-sheet ol{margin:0;padding-left:0;list-style-position:outside}
.word-sheet li{margin:0}
.word-sheet li:not([style]){margin-left:18pt}
.word-sheet li > ul,.word-sheet li > ol{margin-left:18pt}
.word-sheet table{border-collapse:collapse;table-layout:fixed;width:100%;margin:0}
.word-sheet td,.word-sheet th{border:none;padding:0;vertical-align:top;text-align:left}
.word-sheet th{background:transparent;font-weight:700}
.word-sheet img{max-width:100%;height:auto;margin:0}
.word-sheet blockquote{margin:0 0 8pt 18pt;padding:0;border:none;color:inherit;font-style:italic}
.word-sheet a{color:#0563c1;text-decoration:underline}
.word-sheet code,.word-sheet pre{font-family:Consolas,'Courier New',monospace;font-size:10pt}
.word-sheet pre{white-space:pre-wrap;margin:0 0 8pt}
.sheet-title{font-size:12pt;font-weight:700;margin:0 0 6pt}
`;

/** Оболочка печатного документа: один HTML на все страницы. */
function htmlShell(title: string, css: string, body: string): string {
  return (
    '<!doctype html><html lang="ru"><head><meta charset="utf-8">' +
    `<title>${escapeHtml(title)}</title><style>${css}</style></head><body>${body}</body></html>`
  );
}

/** CSS страниц Word-документа: размер листа, поля, колонтитулы. */
function pageCss(geometry: PageGeometry): string {
  const px = (value: number) => `${Math.round(value * 100) / 100}px`;
  return [
    `@page{size:${geometry.widthMm}mm ${geometry.heightMm}mm;margin:0}`,
    'html,body{margin:0;padding:0;background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}',
    `.print-page{position:relative;width:${geometry.widthMm}mm;height:${geometry.heightMm}mm;` +
      'overflow:hidden;background:#fff;break-after:page;page-break-after:always}',
    '.print-page:last-child{break-after:auto;page-break-after:auto}',
    `.print-body{position:absolute;left:${px(geometry.marginLeftPx)};top:${px(geometry.bodyTopPx)};` +
      `width:${px(geometry.contentWidthPx)};height:${px(geometry.contentHeightPx)};` +
      'overflow:hidden;display:flex;flex-direction:column}',
    `.print-header{position:absolute;left:${px(geometry.marginLeftPx)};top:${px(geometry.headerPx)};` +
      `width:${px(geometry.contentWidthPx)};font-family:Calibri,'Segoe UI',Arial,sans-serif;` +
      'font-size:10pt;line-height:1.2;overflow:hidden}',
    `.print-footer{position:absolute;left:${px(geometry.marginLeftPx)};bottom:${px(geometry.footerPx)};` +
      `width:${px(geometry.contentWidthPx)};font-family:Calibri,'Segoe UI',Arial,sans-serif;` +
      'font-size:10pt;line-height:1.2;overflow:hidden}',
  ].join('\n');
}

function pageSizeOfGeometry(geometry: PageGeometry): { width: number; height: number } {
  return { width: mmToMicrons(geometry.widthMm), height: mmToMicrons(geometry.heightMm) };
}

/* --------------------------------------------------------------- сборка */

/**
 * Готовые страницы → печатный документ: каждая страница отдельным блоком с
 * колонтитулами и номером страницы. Этой же функцией пользуется редактор
 * Word, отдавая уже разложенные листы, — тогда число страниц на экране и на
 * бумаге совпадает по построению.
 */
export function pagesToPrintDocument(doc: {
  title: string;
  pages: string[];
  geometry?: PageGeometry;
  headers?: DocxHeaderFooterHtml;
  footers?: DocxHeaderFooterHtml;
}): PrintDocument {
  const geometry = doc.geometry ?? resolvePageGeometry(WORD_DEFAULT_PAGE);
  const pages = doc.pages.length > 0 ? doc.pages : [''];
  const total = pages.length;
  const body = pages
    .map((content, index) => {
      const number = index + 1;
      const header = pickHeaderFooter(doc.headers, number);
      const footer = pickHeaderFooter(doc.footers, number);
      return (
        '<div class="print-page">' +
        (header
          ? `<div class="print-header">${fillHeaderFooterFields(header, number, total)}</div>`
          : '') +
        `<div class="print-body"><div class="word-sheet">${content || '<p></p>'}</div></div>` +
        (footer
          ? `<div class="print-footer">${fillHeaderFooterFields(footer, number, total)}</div>`
          : '') +
        '</div>'
      );
    })
    .join('');
  return {
    title: doc.title,
    html: htmlShell(doc.title, `${pageCss(geometry)}\n${WORD_TYPOGRAPHY}`, body),
    pageSizeMicrons: pageSizeOfGeometry(geometry),
  };
}

/**
 * Потоковый документ (Word, Markdown, текст): пагинатор режет поток на
 * страницы, каждая страница печати — отдельный блок.
 */
export function buildFlowDocument(doc: {
  title: string;
  bodyHtml: string;
  geometry?: PageGeometry;
  headers?: DocxHeaderFooterHtml;
  footers?: DocxHeaderFooterHtml;
}): PrintDocument {
  const geometry = doc.geometry ?? resolvePageGeometry(WORD_DEFAULT_PAGE);
  const { pages } = paginateToPages(doc.bodyHtml || '<p></p>', { geometry });
  return pagesToPrintDocument({
    title: doc.title,
    pages,
    geometry,
    headers: doc.headers,
    footers: doc.footers,
  });
}

/**
 * Таблица Excel: лист на своей странице, шапка повторяется на новых
 * страницах (пагинатор переносит `<thead>`), широкая таблица вписывается
 * масштабом, как «Вписать по ширине» в Excel.
 */
export function buildExcelDocument(doc: {
  title: string;
  sheets: Array<{ name: string; html: string; widthPx?: number }>;
}): PrintDocument {
  const geometry = resolvePageGeometry({
    ...WORD_DEFAULT_PAGE,
    widthMm: 297,
    heightMm: 210,
    marginTopMm: 12,
    marginRightMm: 10,
    marginBottomMm: 12,
    marginLeftMm: 10,
  });
  const body = doc.sheets
    .map((sheet, index) => {
      const scale =
        sheet.widthPx && sheet.widthPx > geometry.contentWidthPx
          ? Math.floor((geometry.contentWidthPx / sheet.widthPx) * 1000) / 1000
          : 1;
      const zoom = scale < 1 ? ` style="zoom:${scale}"` : '';
      return (
        `<div class="sheet-title"${index > 0 ? ' data-page-break="start"' : ''}>` +
        `Лист: ${escapeHtml(sheet.name)}</div>` +
        `<div${zoom}>${sheet.html}</div>`
      );
    })
    .join('');
  return buildFlowDocument({ title: doc.title, bodyHtml: body, geometry });
}

/** Документ из готовых страниц фиксированного размера (PDF, изображение). */
export function buildFixedPagesDocument(doc: {
  title: string;
  pages: string[];
  pageWidthPx: number;
  pageHeightPx: number;
}): PrintDocument {
  const body = doc.pages
    .map(
      (page) =>
        `<div class="print-page" style="position:relative;width:${doc.pageWidthPx}px;` +
        `height:${doc.pageHeightPx}px;overflow:hidden;break-after:page;page-break-after:always">` +
        page +
        '</div>'
    )
    .join('');
  const css = [
    `@page{size:${doc.pageWidthPx}px ${doc.pageHeightPx}px;margin:0}`,
    'html,body{margin:0;padding:0;background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}',
    '.print-page:last-child{break-after:auto;page-break-after:auto}',
    '.print-page img{display:block;width:100%;height:100%;object-fit:contain}',
  ].join('\n');
  return {
    title: doc.title,
    html: htmlShell(doc.title, css, body),
    pageSizeMicrons: {
      width: Math.round(mmToMicrons(pxToMm(doc.pageWidthPx))),
      height: Math.round(mmToMicrons(pxToMm(doc.pageHeightPx))),
    },
  };
}

/** Карточка файла: печать для форматов без собственной вёрстки. */
export function buildInfoDocument(doc: {
  title: string;
  rows: Array<[string, string]>;
  note?: string;
}): PrintDocument {
  const table = doc.rows
    .map(
      ([label, value]) =>
        '<tr>' +
        `<td style="width:38%;padding:3pt 0;border-bottom:0.5pt solid #d1d5db;font-weight:600">${escapeHtml(
          label
        )}</td>` +
        `<td style="padding:3pt 0;border-bottom:0.5pt solid #d1d5db">${escapeHtml(value)}</td>` +
        '</tr>'
    )
    .join('');
  const body =
    `<h1>${escapeHtml(doc.title)}</h1>` +
    `<p style="color:#4b5563">Содержимое этого формата нельзя разложить по страницам, ` +
    'поэтому печатается карточка файла.</p>' +
    `<table style="width:100%;border-collapse:collapse">${table}</table>` +
    (doc.note ? `<p style="margin-top:10pt;color:#6b7280">${escapeHtml(doc.note)}</p>` : '');
  return buildFlowDocument({ title: doc.title, bodyHtml: body });
}

/* ------------------------------------------------- модули по типам файлов */

async function workbookFromContent(content: string, encoding: string): Promise<XLSX.WorkBook> {
  return XLSX.read(content, {
    type: encoding === 'base64' ? 'base64' : 'string',
    cellStyles: true,
  });
}

/** Все листы книги как отдельные таблицы (без оформления ячеек). */
export function sheetsFromWorkbook(workbook: XLSX.WorkBook): Array<{ name: string; html: string }> {
  return workbook.SheetNames.map((name) => ({
    name,
    html: XLSX.utils.sheet_to_html(workbook.Sheets[name], { header: '', footer: '' }),
  }));
}

async function buildPdfDocument(title: string, base64: string): Promise<PrintDocument> {
  const { loadPdfDocument } = await import('../lib/pdfRender');
  const pdf = await loadPdfDocument(base64);
  const first = await pdf.getPage(1);
  const base = first.getViewport({ scale: 1 });
  const widthPx = ptToPx(base.width);
  const heightPx = ptToPx(base.height);
  const pages: string[] = [];
  for (let index = 0; index < pdf.numPages; index += 1) {
    const page = await pdf.getPage(index + 1);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const context = canvas.getContext('2d');
    if (!context) break;
    await page.render({ canvasContext: context, viewport, canvas }).promise;
    pages.push(`<img src="${canvas.toDataURL('image/jpeg', 0.92)}" alt="">`);
  }
  return buildFixedPagesDocument({ title, pages, pageWidthPx: widthPx, pageHeightPx: heightPx });
}

/** Изображение по площади печати: вписано, по центру, без обрезки. */
export function buildImageDocument(title: string, src: string): PrintDocument {
  const geometry = resolvePageGeometry(WORD_DEFAULT_PAGE);
  const body =
    '<div class="print-page">' +
    '<div style="position:absolute;left:0;top:0;width:100%;height:100%;display:flex;' +
    'align-items:center;justify-content:center;padding:15mm;box-sizing:border-box">' +
    `<img src="${src}" alt="" style="max-width:100%;max-height:100%;object-fit:contain">` +
    '</div></div>';
  return {
    title,
    html: htmlShell(
      title,
      `@page{size:${geometry.widthMm}mm ${geometry.heightMm}mm;margin:0}html,body{margin:0;padding:0;background:#fff}`,
      body
    ),
    pageSizeMicrons: pageSizeOfGeometry(geometry),
  };
}

/* ------------------------------------------------------------- сборка по типу */

async function buildFromWord(
  title: string,
  base64: string,
  override?: { bodyHtml?: string; page?: PageSetup; headers?: DocxHeaderFooterHtml; footers?: DocxHeaderFooterHtml }
): Promise<PrintDocument> {
  if (override?.bodyHtml !== undefined) {
    return buildFlowDocument({
      title,
      bodyHtml: override.bodyHtml,
      geometry: resolvePageGeometry(override.page),
      headers: override.headers,
      footers: override.footers,
    });
  }
  const parsed = await docxBase64ToHtml(base64);
  return buildFlowDocument({
    title,
    bodyHtml: parsed.html,
    geometry: resolvePageGeometry(parsed.page),
    headers: parsed.headers,
    footers: parsed.footers,
  });
}

async function buildFromTab(tab: OpenTab): Promise<PrintDocument> {
  const title = baseName(tab.name).replace(/\.[^.]+$/, '');
  switch (tab.type) {
    case 'word':
      return await buildFromWord(title, tab.content);
    case 'markdown':
    case 'richtext':
      return buildFlowDocument({ title, bodyHtml: markdownToDocumentHtml(tab.content) });
    case 'text':
    case 'code':
      return buildFlowDocument({ title, bodyHtml: plainTextToHtml(tab.content) });
    case 'excel': {
      const workbook = await workbookFromContent(tab.content, tab.encoding);
      return buildExcelDocument({ title, sheets: sheetsFromWorkbook(workbook) });
    }
    case 'powerpoint': {
      const { openPptxDocument } = await import('../lib/pptxDoc');
      const document = await openPptxDocument(tab.content);
      return {
        title,
        html: buildPptxPrintHtml(document.model),
        pageSizeMicrons: pptxPageMicrons(document.model),
      };
    }
    case 'pdf':
      return await buildPdfDocument(title, tab.content);
    case 'image':
      return buildImageDocument(title, tab.content);
    case 'video':
    case 'binary':
    default:
      return buildInfoDocument({
        title: tab.name,
        rows: [
          ['Файл', tab.name],
          ['Путь', tab.path],
          ['Тип', tab.type],
        ],
      });
  }
}

async function buildFromFile(filePath: string): Promise<PrintDocument> {
  const title = baseName(filePath).replace(/\.[^.]+$/, '');
  const type = getFileType(extName(filePath)) as FileType;
  const stat = await fileSystemService.statFile(filePath).catch(() => null);

  if (type === 'word' || type === 'markdown' || type === 'excel' || type === 'powerpoint' || type === 'pdf' || type === 'image') {
    const source = await readFileSource(filePath);
    if (type === 'word' && source.extension !== 'doc') {
      return await buildFromWord(title, source.content);
    }
    if (type === 'markdown') {
      return buildFlowDocument({ title, bodyHtml: markdownToDocumentHtml(source.content) });
    }
    if (type === 'excel') {
      const workbook = await workbookFromContent(source.content, source.encoding);
      return buildExcelDocument({ title, sheets: sheetsFromWorkbook(workbook) });
    }
    if (type === 'powerpoint' && source.extension === 'pptx') {
      const { openPptxDocument } = await import('../lib/pptxDoc');
      const document = await openPptxDocument(source.content);
      return {
        title,
        html: buildPptxPrintHtml(document.model),
        pageSizeMicrons: pptxPageMicrons(document.model),
      };
    }
    if (type === 'pdf') return await buildPdfDocument(title, source.content);
    if (type === 'image') {
      const src =
        source.encoding === 'base64'
          ? `data:image/${source.extension === 'svg' ? 'svg+xml' : source.extension};base64,${source.content}`
          : source.content;
      return buildImageDocument(title, src);
    }
  }

  const infoRows: Array<[string, string]> = [
    ['Файл', baseName(filePath)],
    ['Путь', filePath],
  ];
  if (stat?.size) infoRows.push(['Размер', formatBytes(stat.size)]);
  if (stat?.mtimeMs) infoRows.push(['Изменён', new Date(stat.mtimeMs).toLocaleString('ru-RU')]);
  const text =
    type === 'text' || type === 'code' || type === 'richtext'
      ? (await readFileSource(filePath)).content
      : '';
  if (text.trim()) {
    return buildFlowDocument({
      title,
      bodyHtml: `<pre style="font-family:Consolas,monospace;font-size:9pt;white-space:pre-wrap">${escapeHtml(
        text.slice(0, 200000)
      )}</pre>`,
    });
  }
  return buildInfoDocument({ title: baseName(filePath), rows: infoRows });
}

/* --------------------------------------------------------------- печать */

/** Собирает печатный документ для вкладки или файла. */
export async function buildPrintDocument(target: PrintTarget): Promise<PrintDocument> {
  if (target.kind === 'file') return await buildFromFile(target.path);
  const provider = getPrintProvider(target.tab.id);
  if (provider) {
    const built = await provider.build().catch((err) => {
      // Сбой сборки в редакторе не должен отменять печать: печатаем по
      // содержимому вкладки, как для неоткрытого файла.
      console.warn('Печать: сборщик редактора не сработал', err);
      return null;
    });
    if (built) return built;
  }
  return await buildFromTab(target.tab);
}

/**
 * Печать или предпросмотр. Предпросмотр возвращает тот же PDF, который
 * ушёл бы на принтер, и сам собранный документ — чтобы «Печать» из окна
 * предпросмотра отправила на принтер ровно показанное.
 */
export async function printTarget(
  target: PrintTarget,
  options: { preview?: boolean } = {}
): Promise<PrintResult & { document?: PrintDocument }> {
  let document: PrintDocument;
  try {
    document = await buildPrintDocument(target);
  } catch (err) {
    return { success: false, error: `Не удалось подготовить документ: ${(err as Error).message}` };
  }

  const api = window.electronAPI?.printDocument;
  if (!api) {
    // Браузерная сборка: печатаем тем же HTML через скрытый фрейм.
    return { success: false, error: 'Печать доступна только в настольном приложении' };
  }

  const result = await api({
    html: document.html,
    title: document.title,
    mode: options.preview ? 'preview' : 'print',
    pageSizeMicrons: document.pageSizeMicrons,
    landscape: document.landscape,
    scaleFactor: document.scaleFactor,
  });
  return { ...result, document };
}

/** Печать уже собранного документа (кнопка «Печать» из предпросмотра). */
export async function printBuiltDocument(document: PrintDocument): Promise<PrintResult> {
  const api = window.electronAPI?.printDocument;
  if (!api) return { success: false, error: 'Печать доступна только в настольном приложении' };
  return await api({
    html: document.html,
    title: document.title,
    mode: 'print',
    pageSizeMicrons: document.pageSizeMicrons,
    landscape: document.landscape,
    scaleFactor: document.scaleFactor,
  });
}
