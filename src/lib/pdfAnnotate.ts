/**
 * Аннотации PDF: маркер и текстовые заметки, вшитые в документ.
 *
 * Координаты задаются нормализованно (0..1 от ширины/высоты страницы), чтобы
 * не зависеть от масштаба превью; при записи пересчитываются в точки страницы.
 */
import { PDFDocument, StandardFonts, degrees, rgb } from 'pdf-lib';
import type { PDFFont, PDFPage } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { bytesToBase64, loadPdfDocument } from './pdfOps';

/** pdf-lib не умеет встраивать TTF без fontkit — регистрируем на документе. */
function ensureFontkit(doc: PDFDocument): void {
  const target = doc as unknown as { registerFontkit?: (instance: unknown) => void };
  target.registerFontkit?.(fontkit);
}

export interface AnnotationPoint {
  /** 0-based индекс страницы. */
  page: number;
  /** 0..1, от левого края. */
  x: number;
  /** 0..1, от верхнего края. */
  y: number;
}

export interface HighlightOptions extends AnnotationPoint {
  /** Ширина и высота подсветки в долях страницы. */
  width?: number;
  height?: number;
  color?: { r: number; g: number; b: number };
}

export interface NoteOptions extends AnnotationPoint {
  text: string;
  size?: number;
  color?: { r: number; g: number; b: number };
  /** TTF с поддержкой кириллицы (Arial/Segoe UI); без него Latin-1 only. */
  fontBytes?: Uint8Array;
}

function clamp01(value: number): number {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.min(1, Math.max(0, num));
}

function normalizeRotation(rotation: number): number {
  return ((Math.round(rotation / 90) * 90) % 360 + 360) % 360;
}

/** Размер отображённой страницы: при 90/270 pdf.js меняет оси местами. */
function displaySize(
  page: { width: number; height: number },
  rotation: number
): { width: number; height: number } {
  const angle = normalizeRotation(rotation);
  return angle === 90 || angle === 270
    ? { width: page.height, height: page.width }
    : { width: page.width, height: page.height };
}

/**
 * Пересчёт нормализованной точки в координаты pdf-lib (начало координат снизу).
 * Поворот из page.getRotation() учитывается: без него маркер после rotatePages
 * ложился не туда, потому что pdf.js рисует страницу повёрнутой.
 */
export function toPagePoint(
  page: { width: number; height: number },
  point: { x: number; y: number },
  rotation = 0
): { x: number; y: number } {
  const x = clamp01(point.x);
  const y = clamp01(point.y);
  const angle = normalizeRotation(rotation);
  if (angle === 90) return { x: y * page.width, y: x * page.height };
  if (angle === 180) return { x: (1 - x) * page.width, y: y * page.height };
  if (angle === 270) return { x: (1 - y) * page.width, y: (1 - x) * page.height };
  return { x: x * page.width, y: (1 - y) * page.height };
}

/** Прямоугольник в долях отображённой страницы → координаты страницы. */
function annotationBox(
  size: { width: number; height: number },
  rotation: number,
  box: { x: number; y: number; width: number; height: number }
): { x: number; y: number; width: number; height: number } {
  const corners = [
    { x: box.x, y: box.y },
    { x: box.x + box.width, y: box.y },
    { x: box.x, y: box.y + box.height },
    { x: box.x + box.width, y: box.y + box.height },
  ].map((point) => toPagePoint(size, point, rotation));
  const xs = corners.map((point) => point.x);
  const ys = corners.map((point) => point.y);
  const left = Math.min(...xs);
  const bottom = Math.min(...ys);
  return {
    x: left,
    y: bottom,
    width: Math.max(...xs) - left,
    height: Math.max(...ys) - bottom,
  };
}

export const HIGHLIGHT_COLOR = { r: 1, g: 0.86, b: 0.2 };
export const NOTE_COLOR = { r: 0.95, g: 0.3, b: 0.3 };

/** Кириллица не кодируется в Helvetica — проверяем до загрузки шрифта. */
function requireNoteFont(text: string, fontBytes?: Uint8Array): void {
  if (/[^\x00-\xFF]/.test(text) && !fontBytes) {
    throw new Error('Для заметок с кириллицей нужен системный шрифт (Arial/Segoe UI)');
  }
}

async function embedNoteFont(doc: PDFDocument, fontBytes?: Uint8Array): Promise<PDFFont> {
  if (!fontBytes) return doc.embedFont(StandardFonts.Helvetica);
  ensureFontkit(doc);
  return doc.embedFont(fontBytes, { subset: true });
}

/** Маркер: полупрозрачный прямоугольник поверх содержимого страницы. */
function drawHighlight(page: PDFPage, options: HighlightOptions): void {
  const size = page.getSize();
  const box = annotationBox(size, page.getRotation().angle, {
    x: options.x,
    y: options.y,
    width: clamp01(options.width ?? 0.6),
    height: clamp01(options.height ?? 0.06),
  });
  const color = options.color ?? HIGHLIGHT_COLOR;
  page.drawRectangle({
    ...box,
    width: Math.max(4, box.width),
    height: Math.max(4, box.height),
    color: rgb(color.r, color.g, color.b),
    opacity: 0.35,
    borderWidth: 0,
  });
}

/** Текстовая заметка: подложка и строки, ужатые по ширине отображённой страницы. */
function drawNote(page: PDFPage, options: NoteOptions, text: string, font: PDFFont): void {
  const size = page.getSize();
  const rotation = page.getRotation().angle;
  const angle = normalizeRotation(rotation);
  const visible = displaySize(size, rotation);
  const fontSize = Math.max(6, Math.min(24, options.size ?? 10));
  const color = options.color ?? NOTE_COLOR;
  const maxWidth = visible.width * 0.9;
  const lines: string[] = [];
  let current = '';
  for (const word of text.split(/\s+/)) {
    const candidate = current ? `${current} ${word}` : word;
    if (font.widthOfTextAtSize(candidate, fontSize) > maxWidth && current) {
      lines.push(current);
      current = word;
    } else {
      current = candidate;
    }
  }
  if (current) lines.push(current);

  const lineHeight = fontSize * 1.35;
  const boxHeight = lines.length * lineHeight + 6;
  const box = annotationBox(size, rotation, {
    x: options.x,
    y: options.y,
    width: 0.9,
    height: Math.min(1, boxHeight / visible.height),
  });
  page.drawRectangle({
    ...box,
    color: rgb(1, 1, 1),
    opacity: 0.82,
    borderColor: rgb(color.r, color.g, color.b),
    borderWidth: 0.7,
  });
  // Подложка считается по осям отображения (annotationBox), поэтому при
  // /Rotate 90|270 строки рисуем в осях страницы с тем же поворотом: иначе
  // текст шёл поперёк рамки и вылезал за неё.
  const paint = { size: fontSize, font, color: rgb(color.r, color.g, color.b) };
  lines.forEach((line, index) => {
    if (angle === 90) {
      // Слева-вправо по отображению — это ось Y страницы, строки вниз — ось X.
      page.drawText(line, {
        ...paint,
        x: box.x + (index + 1) * lineHeight,
        y: box.y + 4,
        rotate: degrees(90),
      });
      return;
    }
    if (angle === 270) {
      page.drawText(line, {
        ...paint,
        x: box.x + box.width - (index + 1) * lineHeight,
        y: box.y + box.height - 4,
        rotate: degrees(270),
      });
      return;
    }
    page.drawText(line, {
      ...paint,
      x: box.x + 4,
      y: box.y + box.height - (index + 1) * lineHeight,
    });
  });
}

export async function addHighlight(base64: string, options: HighlightOptions): Promise<string> {
  const doc = await loadPdfDocument(base64);
  const page = doc.getPages()[options.page];
  if (!page) throw new Error('Страница не найдена');
  drawHighlight(page, options);
  return bytesToBase64(await doc.save());
}

/** Текстовая заметка с полупрозрачной подложкой. */
export async function addNote(base64: string, options: NoteOptions): Promise<string> {
  const text = (options.text || '').trim();
  if (!text) throw new Error('Текст заметки пуст');
  requireNoteFont(text, options.fontBytes);
  const doc = await loadPdfDocument(base64);
  const page = doc.getPages()[options.page];
  if (!page) throw new Error('Страница не найдена');
  drawNote(page, options, text, await embedNoteFont(doc, options.fontBytes));
  return bytesToBase64(await doc.save());
}

/** Нормализация клика по превью: пиксели → доли страницы (0..1). */
export function normalizeClick(
  rect: { left: number; top: number; width: number; height: number },
  clientX: number,
  clientY: number
): { x: number; y: number } {
  if (!rect.width || !rect.height) return { x: 0, y: 0 };
  return {
    x: clamp01((clientX - rect.left) / rect.width),
    y: clamp01((clientY - rect.top) / rect.height),
  };
}
export interface AnnotationOp {
  kind: 'marker' | 'note';
  page: number;
  x: number;
  y: number;
  text?: string;
}

/**
 * Применяет список аннотаций к исходному документу с нуля.
 * Нужно для отмены и удаления отдельной заметки: храним только операции и
 * базовый PDF, поэтому результат всегда воспроизводим.
 */
export async function applyAnnotations(
  base64: string,
  ops: AnnotationOp[],
  fontBytes?: Uint8Array
): Promise<string> {
  if (ops.length === 0) return base64;
  // Один load и один save на весь набор: раньше каждая операция переписывала файл.
  const doc = await loadPdfDocument(base64);
  let font: PDFFont | null = null;
  for (const op of ops) {
    const page = doc.getPages()[op.page];
    if (!page) throw new Error('Страница не найдена');
    if (op.kind === 'marker') {
      drawHighlight(page, { page: op.page, x: op.x, y: op.y });
      continue;
    }
    const text = (op.text || '').trim();
    if (!text) continue;
    requireNoteFont(text, fontBytes);
    if (!font) font = await embedNoteFont(doc, fontBytes);
    drawNote(page, { page: op.page, x: op.x, y: op.y, text }, text, font);
  }
  return bytesToBase64(await doc.save());
}
