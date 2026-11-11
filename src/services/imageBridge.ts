/**
 * Мост между ассистентом и просмотрщиком изображений.
 *
 * Нужен, чтобы правки шли по живому буферу редактора (с его отменой),
 * а не через перечитывание файла с диска.
 */
export interface ImageState {
  /** base64 без префикса data:. */
  base64: string;
  mime: string;
  width: number;
  height: number;
  /** Путь файла, если вкладка открыта из workspace. */
  path?: string;
}

export interface ImageEditResult {
  ok: boolean;
  message: string;
}

export interface ImageController {
  tabId: string;
  getImage: () => ImageState | null;
  /** Применяет операции из `lib/imageOps.ts` к живому буферу. */
  applyOps: (ops: unknown[]) => Promise<ImageEditResult>;
  save: () => Promise<ImageEditResult>;
}

const controllers = new Map<string, ImageController>();

export function registerImageController(controller: ImageController): () => void {
  controllers.set(controller.tabId, controller);
  return () => {
    if (controllers.get(controller.tabId) === controller) controllers.delete(controller.tabId);
  };
}

export function getImageController(tabId: string): ImageController | undefined {
  return controllers.get(tabId);
}
