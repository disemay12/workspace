/**
 * Загрузка изображения из data URL, http(s) или файла workspace.
 *
 * Возвращает data URL, чтобы инструменты документа могли вставить картинку
 * без сетевых запросов на этапе записи OOXML/PPTX.
 */
import { fileSystemService } from '../services/fileSystemService';
import { resolveWorkspacePath } from './workspacePath';

export type ImageSourceResult =
  | { ok: true; dataUrl: string }
  | { ok: false; message: string };

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
};

function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ''));
    reader.onerror = () => reject(reader.error ?? new Error('не удалось прочитать изображение'));
    reader.readAsDataURL(blob);
  });
}

export async function resolveImageSource(
  raw: string,
  workspaceRoot: string
): Promise<ImageSourceResult> {
  const source = raw.trim();
  if (!source) return { ok: false, message: 'Пустой путь к изображению' };
  if (/^data:image\//i.test(source)) return { ok: true, dataUrl: source };

  if (/^https?:\/\//i.test(source)) {
    try {
      const response = await fetch(source);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const dataUrl = await blobToDataUrl(await response.blob());
      if (!/^data:image\//i.test(dataUrl)) throw new Error('URL вернул не изображение');
      return { ok: true, dataUrl };
    } catch (error) {
      return {
        ok: false,
        message: `Не удалось загрузить изображение по URL: ${
          error instanceof Error ? error.message : String(error)
        }`,
      };
    }
  }

  const targetPath = resolveWorkspacePath(workspaceRoot, source);
  if (!targetPath) {
    return { ok: false, message: `Изображение «${source}» вне рабочей папки` };
  }
  try {
    const read = await fileSystemService.readFile(targetPath, 'base64');
    const ext = targetPath.split('.').pop()?.toLowerCase() ?? 'png';
    const mime = MIME_BY_EXT[ext] ?? 'image/png';
    return { ok: true, dataUrl: `data:${mime};base64,${read.content}` };
  } catch (error) {
    return {
      ok: false,
      message: `Не удалось прочитать изображение: ${
        error instanceof Error ? error.message : String(error)
      }`,
    };
  }
}
