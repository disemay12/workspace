/**
 * Операции над изображениями в рендерере (canvas): поворот, отражение,
 * обрезка, масштаб, конвертация формата и сжатие до целевого размера.
 *
 * Все функции принимают и возвращают base64 без префикса data:, поэтому
 * одинаково работают и с открытой вкладкой, и с файлом на диске.
 */

export type ImageOp =
  | { op: 'rotate'; degrees: 90 | -90 | 180 | 270 }
  | { op: 'flip'; axis: 'h' | 'v' }
  | { op: 'crop'; x: number; y: number; w: number; h: number; unit?: 'px' | '%' }
  | {
      op: 'resize';
      mode?: 'fit' | 'exact' | 'max';
      width?: number;
      height?: number;
      max?: number;
      keepAspect?: boolean;
    }
  | { op: 'convert'; format: 'png' | 'jpeg' | 'webp'; quality?: number }
  | { op: 'compress'; maxKb: number; format?: 'jpeg' | 'webp' }
  | { op: 'thumbnail'; max: number };

export interface ImageBlob {
  base64: string;
  mime: string;
}

export interface ImageResult extends ImageBlob {
  width: number;
  height: number;
  bytes: number;
}

const MAX_SIDE = 8192;

export function mimeForFormat(format: string): string {
  if (format === 'jpeg' || format === 'jpg') return 'image/jpeg';
  if (format === 'webp') return 'image/webp';
  return 'image/png';
}

export function formatFromMime(mime: string): 'png' | 'jpeg' | 'webp' {
  if (/jpeg|jpg/i.test(mime)) return 'jpeg';
  if (/webp/i.test(mime)) return 'webp';
  return 'png';
}

/** Прямоугольник обрезки, приведённый к границам изображения. */
export function clampCropRect(
  rect: { x: number; y: number; w: number; h: number },
  width: number,
  height: number
): { x: number; y: number; w: number; h: number } {
  const x = Math.max(0, Math.min(Math.round(rect.x), width - 1));
  const y = Math.max(0, Math.min(Math.round(rect.y), height - 1));
  const w = Math.max(1, Math.min(Math.round(rect.w), width - x));
  const h = Math.max(1, Math.min(Math.round(rect.h), height - y));
  return { x, y, w, h };
}

/** Размер после масштабирования: fit — вписать в рамку, max — по большей стороне. */
export function resolveTargetSize(
  width: number,
  height: number,
  options: { mode?: 'fit' | 'exact' | 'max'; width?: number; height?: number; max?: number; keepAspect?: boolean }
): { width: number; height: number } {
  const mode = options.mode ?? 'fit';
  if (mode === 'max' && options.max) {
    const scale = Math.min(1, options.max / Math.max(width, height));
    return {
      width: Math.max(1, Math.round(width * scale)),
      height: Math.max(1, Math.round(height * scale)),
    };
  }
  if (mode === 'exact' && options.width && options.height && options.keepAspect === false) {
    return {
      width: Math.max(1, Math.min(MAX_SIDE, Math.round(options.width))),
      height: Math.max(1, Math.min(MAX_SIDE, Math.round(options.height))),
    };
  }
  const targetW = options.width && options.width > 0 ? options.width : undefined;
  const targetH = options.height && options.height > 0 ? options.height : undefined;
  if (!targetW && !targetH) return { width, height };
  const scale = Math.min(
    targetW ? targetW / width : Number.POSITIVE_INFINITY,
    targetH ? targetH / height : Number.POSITIVE_INFINITY
  );
  return {
    width: Math.max(1, Math.min(MAX_SIDE, Math.round(width * scale))),
    height: Math.max(1, Math.min(MAX_SIDE, Math.round(height * scale))),
  };
}

/** Следующая ступень качества при подборе размера файла. */
export function nextQualityStep(current: number): number {
  return Math.max(0.25, Math.round((current - 0.12) * 100) / 100);
}

async function loadImage(blob: ImageBlob): Promise<HTMLImageElement> {
  const url = `data:${blob.mime || 'image/png'};base64,${blob.base64}`;
  return await new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Не удалось декодировать изображение'));
    image.src = url;
  });
}

function canvasOf(width: number, height: number): HTMLCanvasElement {
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.min(MAX_SIDE, Math.round(width)));
  canvas.height = Math.max(1, Math.min(MAX_SIDE, Math.round(height)));
  return canvas;
}

function encode(canvas: HTMLCanvasElement, mime: string, quality: number): string {
  const url = canvas.toDataURL(mime, quality);
  const comma = url.indexOf(',');
  return comma >= 0 ? url.slice(comma + 1) : '';
}

/** Каркас операций без кодирования результата — для последовательных шагов. */
interface Surface {
  canvas: HTMLCanvasElement;
  mime: string;
}

function surfaceFromImage(image: HTMLImageElement, mime: string): Surface {
  const canvas = canvasOf(image.naturalWidth || image.width, image.naturalHeight || image.height);
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('Canvas 2D недоступен');
  ctx.drawImage(image, 0, 0);
  return { canvas, mime };
}

function applyRotate(surface: Surface, degrees: number): Surface {
  const { canvas } = surface;
  const swap = Math.abs(degrees) % 180 === 90;
  const next = canvasOf(swap ? canvas.height : canvas.width, swap ? canvas.width : canvas.height);
  const ctx = next.getContext('2d');
  if (!ctx) throw new Error('Canvas 2D недоступен');
  ctx.translate(next.width / 2, next.height / 2);
  ctx.rotate((degrees * Math.PI) / 180);
  ctx.drawImage(canvas, -canvas.width / 2, -canvas.height / 2);
  return { canvas: next, mime: surface.mime };
}

function applyFlip(surface: Surface, axis: 'h' | 'v'): Surface {
  const { canvas } = surface;
  const next = canvasOf(canvas.width, canvas.height);
  const ctx = next.getContext('2d');
  if (!ctx) throw new Error('Canvas 2D недоступен');
  ctx.translate(axis === 'h' ? canvas.width : 0, axis === 'v' ? canvas.height : 0);
  ctx.scale(axis === 'h' ? -1 : 1, axis === 'v' ? -1 : 1);
  ctx.drawImage(canvas, 0, 0);
  return { canvas: next, mime: surface.mime };
}

function applyCrop(
  surface: Surface,
  op: { x: number; y: number; w: number; h: number; unit?: 'px' | '%' }
): Surface {
  const { canvas } = surface;
  const unit = op.unit ?? 'px';
  const rect =
    unit === '%'
      ? {
          x: (op.x / 100) * canvas.width,
          y: (op.y / 100) * canvas.height,
          w: (op.w / 100) * canvas.width,
          h: (op.h / 100) * canvas.height,
        }
      : { x: op.x, y: op.y, w: op.w, h: op.h };
  const clamped = clampCropRect(rect, canvas.width, canvas.height);
  const next = canvasOf(clamped.w, clamped.h);
  const ctx = next.getContext('2d');
  if (!ctx) throw new Error('Canvas 2D недоступен');
  ctx.drawImage(canvas, clamped.x, clamped.y, clamped.w, clamped.h, 0, 0, clamped.w, clamped.h);
  return { canvas: next, mime: surface.mime };
}

function applyResize(
  surface: Surface,
  op: { mode?: 'fit' | 'exact' | 'max'; width?: number; height?: number; max?: number; keepAspect?: boolean }
): Surface {
  const { canvas } = surface;
  const size = resolveTargetSize(canvas.width, canvas.height, op);
  if (size.width === canvas.width && size.height === canvas.height) return surface;
  const next = canvasOf(size.width, size.height);
  const ctx = next.getContext('2d');
  if (!ctx) throw new Error('Canvas 2D недоступен');
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(canvas, 0, 0, size.width, size.height);
  return { canvas: next, mime: surface.mime };
}

function bytesOf(base64: string): number {
  return Math.max(0, Math.floor((base64.length * 3) / 4) - (base64.endsWith('==') ? 2 : base64.endsWith('=') ? 1 : 0));
}

/**
 * Применяет операции по порядку. `compress` подбирает качество, пока размер
 * не уложится в лимит (или не упрётся в нижнюю границу).
 */
export async function applyImageOps(
  source: ImageBlob,
  ops: ImageOp[],
  options: { maxKbFloor?: number } = {}
): Promise<ImageResult> {
  const image = await loadImage(source);
  let surface = surfaceFromImage(image, source.mime || 'image/png');

  for (const op of ops) {
    switch (op.op) {
      case 'rotate':
        surface = applyRotate(surface, op.degrees);
        break;
      case 'flip':
        surface = applyFlip(surface, op.axis);
        break;
      case 'crop':
        surface = applyCrop(surface, op);
        break;
      case 'resize':
        surface = applyResize(surface, op);
        break;
      case 'convert':
        surface = { canvas: surface.canvas, mime: mimeForFormat(op.format) };
        break;
      case 'thumbnail':
        surface = applyResize(surface, { mode: 'max', max: op.max });
        break;
      case 'compress': {
        const mime = mimeForFormat(op.format ?? 'jpeg');
        const floor = options.maxKbFloor ?? 64;
        const limitBytes = Math.max(floor, op.maxKb) * 1024;
        let quality = 0.92;
        let best = encode(surface.canvas, mime, quality);
        while (bytesOf(best) > limitBytes && quality > 0.3) {
          quality = nextQualityStep(quality);
          best = encode(surface.canvas, mime, quality);
        }
        surface = { canvas: surface.canvas, mime };
        return {
          base64: best,
          mime,
          width: surface.canvas.width,
          height: surface.canvas.height,
          bytes: bytesOf(best),
        };
      }
      default:
        break;
    }
  }

  const quality = 0.92;
  const base64 = encode(surface.canvas, surface.mime, quality);
  return {
    base64,
    mime: surface.mime,
    width: surface.canvas.width,
    height: surface.canvas.height,
    bytes: bytesOf(base64),
  };
}

/** Разбор операций из аргументов инструмента (JSON уже разобран моделью). */
export function normalizeImageOps(raw: unknown): ImageOp[] {
  if (!Array.isArray(raw)) return [];
  const ops: ImageOp[] = [];
  for (const item of raw) {
    if (!item || typeof item !== 'object') continue;
    const value = item as Record<string, unknown>;
    const op = String(value.op ?? value.operation ?? '').toLowerCase();
    const num = (key: string): number | undefined => {
      const rawValue = value[key];
      const parsed = Number(rawValue);
      return Number.isFinite(parsed) ? parsed : undefined;
    };
    if (op === 'rotate') {
      const degrees = num('degrees') ?? num('angle') ?? 90;
      const normalized = ((Math.round(degrees / 90) * 90) % 360 + 360) % 360;
      ops.push({ op: 'rotate', degrees: (normalized === 0 ? 0 : normalized === 270 ? -90 : normalized) as 90 | -90 | 180 | 270 });
    } else if (op === 'flip') {
      ops.push({ op: 'flip', axis: String(value.axis ?? 'h') === 'v' ? 'v' : 'h' });
    } else if (op === 'crop') {
      ops.push({
        op: 'crop',
        x: num('x') ?? 0,
        y: num('y') ?? 0,
        w: num('w') ?? num('width') ?? 1,
        h: num('h') ?? num('height') ?? 1,
        unit: String(value.unit ?? 'px') === '%' ? '%' : 'px',
      });
    } else if (op === 'resize') {
      const mode = String(value.mode ?? 'fit') as 'fit' | 'exact' | 'max';
      ops.push({
        op: 'resize',
        mode: mode === 'exact' || mode === 'max' ? mode : 'fit',
        width: num('width'),
        height: num('height'),
        max: num('max'),
        keepAspect: value.keep_aspect === false ? false : true,
      });
    } else if (op === 'convert') {
      const format = String(value.format ?? 'png').toLowerCase();
      ops.push({
        op: 'convert',
        format: format === 'jpeg' || format === 'jpg' ? 'jpeg' : format === 'webp' ? 'webp' : 'png',
        quality: num('quality'),
      });
    } else if (op === 'compress') {
      ops.push({
        op: 'compress',
        maxKb: num('max_kb') ?? num('maxKb') ?? 300,
        format: String(value.format ?? 'jpeg') === 'webp' ? 'webp' : 'jpeg',
      });
    } else if (op === 'thumbnail') {
      ops.push({ op: 'thumbnail', max: num('max') ?? 320 });
    }
  }
  return ops;
}
