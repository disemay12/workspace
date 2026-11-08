/**
 * Мост между ассистентом и редактором презентации.
 */
import type { PptxEditResult, PptxImageOptions, PptxModel } from '../lib/pptx';

export interface PptxController {
  tabId: string;
  readModel: () => Promise<PptxModel>;
  getSelection: () => { slide: number; shapeId: string; text: string; paragraphs: string[] } | null;
  setText: (slide: number, shapeId: string, paragraphs: string[]) => Promise<PptxEditResult>;
  addSlide: (index: number, title?: string) => Promise<PptxEditResult>;
  deleteSlide: (index: number) => Promise<PptxEditResult>;
  moveSlide: (from: number, to: number) => Promise<PptxEditResult>;
  addTextBox: (slide: number, text: string) => Promise<PptxEditResult>;
  setNotes: (slide: number, text: string) => Promise<PptxEditResult>;
  addImage: (slide: number, dataUrl: string, options?: PptxImageOptions) => Promise<PptxEditResult>;
  setLayout: (slide: number, layoutPart: string) => Promise<PptxEditResult>;
}

const controllers = new Map<string, PptxController>();

export function registerPptxController(controller: PptxController): () => void {
  controllers.set(controller.tabId, controller);
  return () => {
    if (controllers.get(controller.tabId) === controller) controllers.delete(controller.tabId);
  };
}

export function getPptxController(tabId: string): PptxController | undefined {
  return controllers.get(tabId);
}
