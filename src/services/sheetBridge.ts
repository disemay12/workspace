/**
 * Мост между ассистентом и живым редактором таблицы.
 *
 * ExcelEditor держит книгу в состоянии React; агент читает и меняет её через
 * этот реестр, а не через base64 вкладки, иначе правки терялись бы до save.
 */

export interface SheetCell {
  address: string;
  value: string;
  formula: string;
  display: string;
}

export interface SheetReadResult {
  sheet: string;
  sheetNames: string[];
  range: string;
  rows: string[][];
  cells: SheetCell[];
  truncated: boolean;
}

export interface SheetEditResult {
  ok: boolean;
  message: string;
  changed?: boolean;
}

export interface SheetController {
  tabId: string;
  readSheet: (sheet?: string, range?: string, limit?: number) => SheetReadResult | null;
  getSelection: () => SheetReadResult | null;
  writeCells: (
    sheet: string | undefined,
    cells: Array<{ address: string; value: string }>,
    transactionId: string
  ) => SheetEditResult;
  addSheet: (name?: string) => SheetEditResult;
  renameSheet: (oldName: string, newName: string) => SheetEditResult;
  deleteSheet: (name: string) => SheetEditResult;
  setActiveSheet: (name: string) => SheetEditResult;
  insertRows: (sheet: string | undefined, index: number, count: number) => SheetEditResult;
  deleteRows: (sheet: string | undefined, index: number, count: number) => SheetEditResult;
  insertColumns: (sheet: string | undefined, index: number, count: number) => SheetEditResult;
  deleteColumns: (sheet: string | undefined, index: number, count: number) => SheetEditResult;
  mergeCells: (sheet: string | undefined, range: string) => SheetEditResult;
  unmergeCells: (sheet: string | undefined, range: string) => SheetEditResult;
  formatRange: (
    sheet: string | undefined,
    range: string,
    patch: {
      bold?: boolean;
      italic?: boolean;
      fill?: string;
      color?: string;
      fontSize?: number;
      align?: 'left' | 'center' | 'right';
      wrap?: boolean;
    }
  ) => SheetEditResult;
}

const controllers = new Map<string, SheetController>();

export function registerSheetController(controller: SheetController): () => void {
  controllers.set(controller.tabId, controller);
  return () => {
    if (controllers.get(controller.tabId) === controller) {
      controllers.delete(controller.tabId);
    }
  };
}

export function getSheetController(tabId: string): SheetController | undefined {
  return controllers.get(tabId);
}
