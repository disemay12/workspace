/**
 * Операции над PDF: чтение структуры страниц, поворот, удаление, перестановка
 * и извлечение. Работаем с base64, чтобы одинаково подходить и к вкладке, и к
 * файлу на диске; рендер страниц не нужен — только структура.
 */
import { PDFDocument, degrees } from 'pdf-lib';

export interface PdfPageInfo {
  /** 0-based индекс страницы в документе. */
  index: number;
  width: number;
  height: number;
  rotation: number;
}

export interface PdfStructure {
  pageCount: number;
  pages: PdfPageInfo[];
  title?: string;
}

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob((base64 || '').replace(/^data:[^,]+,/, ''));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = '';
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
  }
  return btoa(binary);
}

/** Разбор диапазона «1-3,5» в 0-based индексы (устойчив к мусору). */
export function parsePageRange(input: string, pageCount: number): number[] {
  const result = new Set<number>();
  const text = (input || '').trim();
  if (!text) return [];
  for (const part of text.split(/[,;]+/)) {
    const piece = part.trim();
    if (!piece) continue;
    const range = /^(\d+)\s*[-–]\s*(\d+)$/.exec(piece);
    if (range) {
      const from = Math.max(1, Number(range[1]));
      const to = Math.min(pageCount, Number(range[2]));
      for (let page = from; page <= to; page += 1) result.add(page - 1);
      continue;
    }
    const single = Number(piece);
    if (Number.isFinite(single) && single >= 1 && single <= pageCount) result.add(single - 1);
  }
  return [...result].sort((a, b) => a - b);
}

async function open(base64: string): Promise<PDFDocument> {
  return await PDFDocument.load(base64ToBytes(base64), { ignoreEncryption: true });
}

export async function inspectPdf(base64: string): Promise<PdfStructure> {
  const doc = await open(base64);
  const pages: PdfPageInfo[] = doc.getPages().map((page, index) => {
    const size = page.getSize();
    return {
      index,
      width: Math.round(size.width),
      height: Math.round(size.height),
      rotation: page.getRotation().angle,
    };
  });
  return { pageCount: pages.length, pages, title: doc.getTitle() };
}

/** Поворот выбранных страниц на 90/180/270 градусов. */
export async function rotatePages(base64: string, indices: number[], delta: number): Promise<string> {
  const doc = await open(base64);
  const pages = doc.getPages();
  for (const index of indices) {
    const page = pages[index];
    if (!page) continue;
    const current = page.getRotation().angle;
    page.setRotation(degrees(((current + delta) % 360 + 360) % 360));
  }
  return bytesToBase64(await doc.save());
}

/** Удаление страниц; последнюю страницу удалить нельзя. */
export async function deletePages(base64: string, indices: number[]): Promise<string> {
  const doc = await open(base64);
  const sorted = [...new Set(indices)].sort((a, b) => b - a);
  if (sorted.length >= doc.getPageCount()) {
    throw new Error('Нельзя удалить все страницы документа');
  }
  for (const index of sorted) doc.removePage(index);
  return bytesToBase64(await doc.save());
}

/** Перестановка страницы: from/to — 0-based. */
export async function movePage(base64: string, from: number, to: number): Promise<string> {
  const doc = await open(base64);
  const count = doc.getPageCount();
  if (from < 0 || from >= count) throw new Error('Нет такой страницы');
  const target = Math.max(0, Math.min(count - 1, to));
  if (target === from) return base64;
  const order = Array.from({ length: count }, (_, index) => index);
  const [moved] = order.splice(from, 1);
  order.splice(target, 0, moved);
  return await reorderPages(base64, order);
}

/** Полная перестановка по массиву 0-based индексов. */
export async function reorderPages(base64: string, order: number[]): Promise<string> {
  const source = await open(base64);
  const count = source.getPageCount();
  const unique = [...new Set(order)].filter((index) => index >= 0 && index < count);
  if (unique.length !== count) throw new Error('Порядок страниц должен содержать все страницы');
  const target = await PDFDocument.create();
  const copied = await target.copyPages(source, unique);
  for (const page of copied) target.addPage(page);
  return bytesToBase64(await target.save());
}

/** Новый документ только из выбранных страниц. */
export async function extractPages(base64: string, indices: number[]): Promise<string> {
  const source = await open(base64);
  const picked = [...new Set(indices)].sort((a, b) => a - b);
  if (picked.length === 0) throw new Error('Выберите страницы для извлечения');
  const target = await PDFDocument.create();
  const copied = await target.copyPages(source, picked);
  for (const page of copied) target.addPage(page);
  return bytesToBase64(await target.save());
}

/** Слияние нескольких PDF в один (в порядке переданных файлов). */
export async function mergePdfs(files: string[]): Promise<string> {
  if (files.length < 2) throw new Error('Нужно минимум два файла');
  const target = await PDFDocument.create();
  for (const file of files) {
    const source = await open(file);
    const copied = await target.copyPages(source, source.getPageIndices());
    for (const page of copied) target.addPage(page);
  }
  return bytesToBase64(await target.save());
}
