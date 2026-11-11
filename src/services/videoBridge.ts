/**
 * Мост между ассистентом и видеоредактором.
 *
 * Если файл открыт во вкладке, операции идут через живое состояние редактора
 * (он сам формирует задание ffmpeg и показывает прогресс). Если вкладки нет —
 * инструменты работают с файлом напрямую через IPC главного процесса.
 */
export interface VideoSegment {
  start: number;
  end: number;
  speed: number;
}

export interface VideoState {
  filePath: string;
  duration: number;
  width: number;
  height: number;
  hasAudio: boolean;
  start: number;
  end: number;
  speed: number;
  segments: VideoSegment[];
  stripAudio: boolean;
  compress: boolean;
  exportZoom: number;
}

/** Задание обработки: те же поля, что ждёт `video:edit` в главном процессе. */
export interface VideoEditPlan {
  start?: number;
  end?: number;
  speed?: number;
  speedSegments?: VideoSegment[];
  stripAudio?: boolean;
  compress?: boolean;
  zoom?: { scale?: number; x?: number; y?: number };
  /** Пустой — сохранить рядом с исходником. */
  outputPath?: string;
}

export interface VideoEditResult {
  ok: boolean;
  message: string;
  outputPath?: string;
}

export interface VideoController {
  tabId: string;
  getState: () => VideoState | null;
  applyPlan: (plan: VideoEditPlan) => Promise<VideoEditResult>;
}

const controllers = new Map<string, VideoController>();

export function registerVideoController(controller: VideoController): () => void {
  controllers.set(controller.tabId, controller);
  return () => {
    if (controllers.get(controller.tabId) === controller) controllers.delete(controller.tabId);
  };
}

export function getVideoController(tabId: string): VideoController | undefined {
  return controllers.get(tabId);
}
