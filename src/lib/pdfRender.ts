/**
 * Рендер страниц PDF в превью: pdf.js в отдельном чанке, воркер берётся из
 * собранного ассета, чтобы не зависеть от сетевых путей.
 */
import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { base64ToBytes } from './pdfOps';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

export type PdfJsDocument = Awaited<ReturnType<typeof pdfjs.getDocument>['promise']>;

export async function loadPdfDocument(base64: string): Promise<PdfJsDocument> {
  return await pdfjs.getDocument({ data: base64ToBytes(base64) }).promise;
}

/** Превью страницы как PNG data URL; максимальная ширина ограничена. */
export async function renderPageToDataUrl(
  doc: PdfJsDocument,
  pageIndex: number,
  maxWidth = 320
): Promise<string> {
  const page = await doc.getPage(pageIndex + 1);
  const base = page.getViewport({ scale: 1 });
  const scale = Math.max(0.2, Math.min(2, maxWidth / base.width));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const context = canvas.getContext('2d');
  if (!context) throw new Error('Не удалось создать холст для превью');
  await page.render({ canvasContext: context, viewport, canvas }).promise;
  return canvas.toDataURL('image/png');
}

/** Текст страницы — для поиска и будущих аннотаций по выделению. */
export async function extractPageText(doc: PdfJsDocument, pageIndex: number): Promise<string> {
  const page = await doc.getPage(pageIndex + 1);
  const content = await page.getTextContent();
  return content.items
    .map((item) => ('str' in item ? item.str : ''))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}
export interface TextMatch {
  page: number;
  /** Доли страницы 0..1 от левого верхнего угла. */
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
}

/**
 * Совпадения на странице с координатами текста: pdf.js отдаёт текст по
 * фрагментам, поэтому позицию внутри фрагмента считаем пропорционально длине.
 */
export async function findMatchesInPage(
  doc: PdfJsDocument,
  pageIndex: number,
  query: string
): Promise<TextMatch[]> {
  const needle = query.trim().toLowerCase();
  if (!needle) return [];
  const page = await doc.getPage(pageIndex + 1);
  const viewport = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();
  const matches: TextMatch[] = [];

  for (const item of content.items) {
    if (!('str' in item) || !item.str) continue;
    const haystack = item.str.toLowerCase();
    let from = haystack.indexOf(needle);
    while (from >= 0) {
      const transform = pdfjs.Util.transform(viewport.transform, item.transform as number[]);
      const fontHeight = Math.hypot(transform[2], transform[3]) || item.height || 10;
      const perChar = item.width / Math.max(1, item.str.length);
      const startX = transform[4] + perChar * from;
      const rectX = startX / viewport.width;
      const rectY = (transform[5] - fontHeight) / viewport.height;
      const rectW = (perChar * needle.length) / viewport.width;
      const rectH = fontHeight / viewport.height;
      matches.push({
        page: pageIndex,
        x: Math.min(1, Math.max(0, rectX)),
        y: Math.min(1, Math.max(0, rectY)),
        width: Math.min(1, Math.max(0.002, rectW)),
        height: Math.min(1, Math.max(0.004, rectH)),
        text: item.str.slice(from, from + needle.length),
      });
      from = haystack.indexOf(needle, from + needle.length);
    }
  }
  return matches;
}

/** Все совпадения по документу, но не больше maxPages страниц. */
export async function findAllMatches(
  doc: PdfJsDocument,
  query: string,
  maxPages = 200
): Promise<TextMatch[]> {
  const result: TextMatch[] = [];
  const limit = Math.min(doc.numPages, maxPages);
  for (let index = 0; index < limit; index += 1) {
    try {
      result.push(...(await findMatchesInPage(doc, index, query)));
    } catch {
      // страница без текстового слоя — пропускаем
    }
  }
  return result;
}