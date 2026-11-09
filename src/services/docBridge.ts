/**
 * Мост между агентом и живым редактором.
 *
 * WordEditor владеет contenteditable-полотном: только он видит несохранённые
 * правки и умеет превратить HTML обратно в .docx. Агент работает через этот
 * реестр, а не через base64 вкладки, иначе каждая следующая правка читала бы
 * устаревший документ.
 */
import type { PageSetup } from '../types';
import type { DocxPackageInfo, HeaderFooterEdit } from '../lib/docx';

/** Точное выделение в живом полотне: адресуется блоком и смещениями. */
export interface DocumentSelection {
  documentId: string;
  blockId: string;
  startOffset: number;
  endOffset: number;
  text: string;
  html: string;
}

export interface DocumentEditResult {
  ok: boolean;
  message: string;
  changed?: boolean;
}

export interface DocumentController {
  tabId: string;
  /** Текущий HTML полотна (включая несохранённые правки). */
  readHtml: () => string;
  /** Сериализует живое полотно в .docx (base64) для «Сохранить как». */
  serialize?: () => Promise<string>;
  /**
   * Записывает HTML в полотно и запускает сериализацию в .docx.
   * transactionId объединяет серию правок одного ответа в один шаг отката.
   */
  writeHtml: (html: string, transactionId: string) => void;
  /** Меняет поля, размер страницы и базовый шрифт .docx. */
  applyPageSetup?: (page: PageSetup) => void;
  /** Последнее непустое выделение в полотне (переживает уход фокуса в чат). */
  getSelection?: () => DocumentSelection | null;
  /** Восстанавливает HTML без записи в историю AI (откат транзакции). */
  restoreHtml?: (html: string) => void;
  /** Заменяет выделенный фрагмент, не перестраивая документ целиком. */
  replaceSelection?: (html: string, transactionId: string) => DocumentEditResult;
  /** Вставляет фрагмент в каретку (или в конец выделения). */
  insertAtSelection?: (html: string, transactionId: string) => DocumentEditResult;
  /** Части исходного .docx: колонтитулы, сноски, комментарии, стили, секции. */
  getDocumentInfo?: () => Promise<DocxPackageInfo | null>;
  /** Меняет или удаляет верхний/нижний колонтитул в пакете. */
  setHeaderFooter?: (edit: HeaderFooterEdit) => Promise<DocumentEditResult>;
}

const controllers = new Map<string, DocumentController>();

export function registerDocumentController(controller: DocumentController): () => void {
  controllers.set(controller.tabId, controller);
  return () => {
    if (controllers.get(controller.tabId) === controller) {
      controllers.delete(controller.tabId);
    }
  };
}

export function getDocumentController(tabId: string): DocumentController | undefined {
  return controllers.get(tabId);
}
