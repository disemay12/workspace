/**
 * Мост «редактор → печать».
 *
 * Печатать нужно то, что видно: у Word, Excel, PowerPoint и PDF живые модели
 * отличаются от содержимого вкладки (несохранённые правки, окно загрузки
 * строк, кэш страниц). Поэтому каждый редактор регистрирует свой сборщик
 * печатного HTML, а `printService` берёт его, если он есть.
 */

/** Готовый к печати документ: один HTML и параметры страницы. */
export interface PrintDocument {
  title: string;
  /** Полный HTML с размерами страниц и разрывами. */
  html: string;
  /** Размер страницы в микронах (как ждёт Electron); по умолчанию из CSS @page. */
  pageSizeMicrons?: { width: number; height: number };
  landscape?: boolean;
  /** Масштаб печати в процентах: вписывание широкой таблицы в лист. */
  scaleFactor?: number;
}

export interface PrintProvider {
  tabId: string;
  /** null — печатать эту вкладку нельзя, пусть используется общий путь. */
  build: () => Promise<PrintDocument | null>;
}

const providers = new Map<string, PrintProvider>();

export function registerPrintProvider(provider: PrintProvider): () => void {
  providers.set(provider.tabId, provider);
  return () => {
    if (providers.get(provider.tabId) === provider) providers.delete(provider.tabId);
  };
}

export function getPrintProvider(tabId: string): PrintProvider | undefined {
  return providers.get(tabId);
}
