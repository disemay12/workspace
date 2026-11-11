/**
 * Мост между ассистентом и редактором презентации.
 *
 * Интерфейс повторяет операции модели документа (`lib/pptxDoc.ts`), поэтому
 * AI-инструменты умеют всё то же, что и кнопки редактора: слайды, текст по
 * run'ам, фигуры, таблицы, темы, фон, переходы, заметки, экспорт.
 */
import type {
  PptxDocModel,
  PptxLegacyModel,
  PptxShapeModel,
  PptxSlideModel,
  PptxTextBody,
  PptxTransitionModel,
  PptxLayoutInfo,
} from '../lib/pptxDoc';
import type { ThemeColorSet, ThemeFontPair } from '../lib/pptxThemes';

export interface PptxEditResult {
  ok: boolean;
  message: string;
}

export interface PptxSelection {
  slide: number;
  shapeId: string;
  text: string;
  paragraphs: string[];
}

export interface PptxTextFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  size?: number;
  color?: string;
  font?: string;
  align?: 'l' | 'ctr' | 'r' | 'just';
  bullet?: 'none' | 'bullet' | 'number';
  level?: number;
  lineSpacing?: number;
}

export interface PptxGeometry {
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  rotation?: number;
  /** Единицы: EMU (по умолчанию) или см. */
  unit?: 'emu' | 'cm';
}

export interface PptxShapeOptions {
  geom?: string;
  geometry?: PptxGeometry;
  fill?: string | null;
  outline?: { color?: string; width?: number } | null;
  shadow?: boolean;
  text?: string;
}

export interface PptxTableOptions {
  rows: number;
  cols: number;
  data?: string[][];
  geometry?: PptxGeometry;
}

export interface PptxImageOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  alt?: string;
  /** Единицы координат: EMU (по умолчанию) или см. */
  unit?: 'emu' | 'cm';
}

export interface PptxController {
  tabId: string;
  /** Текущий base64 пакета (перед вызовом редактор фиксирует правки). */
  getContent: () => Promise<string>;
  /** Полная модель: тема, размер, слайды, фигуры, таблицы, переходы. */
  readModel: () => Promise<PptxDocModel>;
  /** Упрощённая модель для старых инструментов и AI-инструментов 1.1. */
  readLegacyModel?: () => Promise<PptxLegacyModel>;
  getSelection: () => PptxSelection | null;
  getSlide?: (index: number) => Promise<PptxSlideModel | null>;
  listLayouts?: () => Promise<PptxLayoutInfo[]>;
  listThemes?: () => Promise<{ id: string; label: string }[]>;

  /* Слайды */
  addSlide: (index: number, title?: string) => Promise<PptxEditResult>;
  deleteSlide: (index: number) => Promise<PptxEditResult>;
  moveSlide: (from: number, to: number) => Promise<PptxEditResult>;
  duplicateSlide?: (index: number) => Promise<PptxEditResult>;
  setSlideHidden?: (index: number, hidden: boolean) => Promise<PptxEditResult>;
  setLayout: (slide: number, layoutPart: string) => Promise<PptxEditResult>;
  setSlideSize?: (width: number, height: number) => Promise<PptxEditResult>;
  setNotes: (slide: number, text: string) => Promise<PptxEditResult>;
  setTransition?: (slide: number, transition: PptxTransitionModel | null) => Promise<PptxEditResult>;
  setBackground?: (
    slide: number,
    background: { color?: string; imageDataUrl?: string; all?: boolean }
  ) => Promise<PptxEditResult>;

  /* Текст */
  setText: (slide: number, shapeId: string, paragraphs: string[]) => Promise<PptxEditResult>;
  setTextBody?: (slide: number, shapeId: string, body: PptxTextBody) => Promise<PptxEditResult>;
  formatText?: (slide: number, shapeId: string, format: PptxTextFormat) => Promise<PptxEditResult>;
  findReplace?: (
    query: string,
    replacement: string,
    options?: { includeNotes?: boolean }
  ) => Promise<PptxEditResult & { count?: number }>;

  /* Фигуры */
  addTextBox: (slide: number, text: string) => Promise<PptxEditResult>;
  addShape?: (slide: number, options: PptxShapeOptions) => Promise<PptxEditResult & { shapeId?: string }>;
  setShapeGeometry?: (slide: number, shapeId: string, geometry: PptxGeometry) => Promise<PptxEditResult>;
  deleteShape?: (slide: number, shapeId: string) => Promise<PptxEditResult>;
  duplicateShape?: (slide: number, shapeId: string) => Promise<PptxEditResult & { shapeId?: string }>;
  reorderShape?: (
    slide: number,
    shapeId: string,
    mode: 'front' | 'back' | 'forward' | 'backward'
  ) => Promise<PptxEditResult>;
  setShapeStyle?: (
    slide: number,
    shapeId: string,
    style: { fill?: string | null; outline?: { color?: string; width?: number } | null; shadow?: boolean }
  ) => Promise<PptxEditResult>;

  /* Таблицы и изображения */
  addTable?: (slide: number, options: PptxTableOptions) => Promise<PptxEditResult & { shapeId?: string }>;
  setTableData?: (slide: number, shapeId: string, data: string[][]) => Promise<PptxEditResult>;
  addImage: (slide: number, dataUrl: string, options?: PptxImageOptions) => Promise<PptxEditResult>;

  /* Тема */
  applyTheme?: (options: {
    presetId?: string;
    colors?: Partial<ThemeColorSet>;
    fonts?: ThemeFontPair;
  }) => Promise<PptxEditResult>;

  /* Показ и экспорт */
  present?: (index?: number) => Promise<PptxEditResult>;
  exportPdf?: (path?: string) => Promise<PptxEditResult & { path?: string }>;
  exportSlideImages?: (folder?: string) => Promise<PptxEditResult & { files?: string[] }>;

  /** Возвращает форму фигуры по id — нужен инструментам, которые правят стиль. */
  getShape?: (slide: number, shapeId: string) => Promise<PptxShapeModel | null>;
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
