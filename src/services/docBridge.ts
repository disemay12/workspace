/**
 * Мост между агентом и живым редактором.
 *
 * WordEditor владеет contenteditable-полотном: только он видит несохранённые
 * правки и умеет превратить HTML обратно в .docx. Агент работает через этот
 * реестр, а не через base64 вкладки, иначе каждая следующая правка читала бы
 * устаревший документ.
 */

export interface DocumentController {
  tabId: string;
  /** Текущий HTML полотна (включая несохранённые правки). */
  readHtml: () => string;
  /**
   * Записывает HTML в полотно и запускает сериализацию в .docx.
   * transactionId объединяет серию правок одного ответа в один шаг отката.
   */
  writeHtml: (html: string, transactionId: string) => void;
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
