/**
 * Аннотации PDF: маркер и текстовые заметки, вшитые в документ.
 *
 * Координаты задаются нормализованно (0..1 от ширины/высоты страницы), чтобы
 * не зависеть от масштаба превью; при записи пересчитываются в точки страницы.
 */
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { base64ToBytes, bytesToBase64 } from './pdfOps';

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

/** Пересчёт нормализованной точки в координаты страницы (y снизу). */
export function toPagePoint(
  page: { width: number; height: number },
  point: { x: number; y: number }
): { x: number; y: number } {
  return {
    x: clamp01(point.x) * page.width,
    y: (1 - clamp01(point.y)) * page.height,
  };
}

export const HIGHLIGHT_COLOR = { r: 1, g: 0.86, b: 0.2 };
export const NOTE_COLOR = { r: 0.95, g: 0.3, b: 0.3 };

/** Маркер: полупрозрачный прямоугольник поверх содержимого страницы. */
export async function addHighlight(base64: string, options: HighlightOptions): Promise<string> {
  const doc = await PDFDocument.load(base64ToBytes(base64), { ignoreEncryption: true });
  const page = doc.getPages()[options.page];
  if (!page) throw new Error('Страница не найдена');
  const size = page.getSize();
  const start = toPagePoint(size, options);
  const width = Math.max(4, clamp01(options.width ?? 0.6) * size.width);
  const height = Math.max(4, clamp01(options.height ?? 0.06) * size.height);
  const color = options.color ?? HIGHLIGHT_COLOR;
  page.drawRectangle({
    x: start.x,
    y: Math.min(size.height - height, Math.max(0, start.y - height)),
    width: Math.min(width, size.width - start.x),
    height,
    color: rgb(color.r, color.g, color.b),
    opacity: 0.35,
    borderWidth: 0,
  });
  return bytesToBase64(await doc.save());
}

/** Текстовая заметка с полупрозрачной подложкой. */
export async function addNote(base64: string, options: NoteOptions): Promise<string> {
  const text = (options.text || '').trim();
  if (!text) throw new Error('Текст заметки пуст');
  const needsUnicode = /[^\x00-\xFF]/.test(text);
  if (needsUnicode && !options.fontBytes) {
    throw new Error('Для заметок с кириллицей нужен системный шрифт (Arial/Segoe UI)');
  }
  const doc = await PDFDocument.load(base64ToBytes(base64), { ignoreEncryption: true });
  const page = doc.getPages()[options.page];
  if (!page) throw new Error('Страница не найдена');
  const size = page.getSize();
  if (options.fontBytes) ensureFontkit(doc);
  const font = options.fontBytes
    ? await doc.embedFont(options.fontBytes, { subset: true })
    : await doc.embedFont(StandardFonts.Helvetica);
  const fontSize = Math.max(6, Math.min(24, options.size ?? 10));
  const color = options.color ?? NOTE_COLOR;
  const maxWidth = size.width * 0.9;
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
  const start = toPagePoint(size, options);
  const boxHeight = lines.length * lineHeight + 6;
  page.drawRectangle({
    x: Math.max(2, start.x),
    y: Math.min(size.height - boxHeight, Math.max(0, start.y - boxHeight)),
    width: maxWidth,
    height: boxHeight,
    color: rgb(1, 1, 1),
    opacity: 0.82,
    borderColor: rgb(color.r, color.g, color.b),
    borderWidth: 0.7,
  });
  lines.forEach((line, index) => {
    page.drawText(line, {
      x: Math.max(2, start.x) + 4,
      y: Math.min(size.height - boxHeight, Math.max(0, start.y - boxHeight)) + boxHeight - 4 - (index + 1) * lineHeight + 4,
      size: fontSize,
      font,
      color: rgb(color.r, color.g, color.b),
    });
  });
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
  let current = base64;
  for (const op of ops) {
    if (op.kind === 'marker') {
      current = await addHighlight(current, { page: op.page, x: op.x, y: op.y });
    } else {
      current = await addNote(current, {
        page: op.page,
        x: op.x,
        y: op.y,
        text: op.text || '',
        fontBytes,
      });
    }
  }
  return current;
}