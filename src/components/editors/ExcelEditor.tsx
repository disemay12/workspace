import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { OpenTab } from '../../types';
import {
  FileSpreadsheet,
  Plus,
  Save,
  Sparkles,
  Table as TableIcon,
  Trash2,
  Download,
  Sigma,
  AlertTriangle,
  Check,
  X,
  Wand2,
} from 'lucide-react';
import { cn } from '../../lib/utils';
import { ClipboardPaste, Undo2, Redo2 } from 'lucide-react';
import { parseClipboardTable } from '../../lib/clipboardTable';
import { Button, Spinner } from '../ui';
import {
  columnName,
  computeCellDisplay,
  createContext,
  isError,
  normalizeFormula,
  repairFormula,
  toExcelFormula,
  fromExcelFormula,
  Workbook,
  Grid,
} from '../../lib/formulaEngine';
import {
  useFormulaMode,
  rangeToA1,
  normalizeRange,
  shiftFormula,
  CellCoord,
} from './useFormulaMode';

interface Props {
  tab: OpenTab;
  onChange: (newContent: string) => void;
  onSave: () => void;
  onAskAI?: (prompt: string, contextText: string) => void;
  /** Применяет действие AI к таблице (вставка формулы, заполнение и т.п.). */
  onAIApply?: (action: AIAction, result: string) => void;
}

/** Действие, которое AI предлагает выполнить над таблицей. */
export interface AIAction {
  kind: 'setCell' | 'fillRange' | 'addSheet';
  /** Адрес целевой ячейки (A1) либо диапазона (A1:B10). */
  target: string;
  /** Значение или формула (без ведущего '='). */
  value: string;
  /** Лист, к которому применяется действие. */
  sheet?: string;
  /** Пояснение для пользователя. */
  description?: string;
}

const MIN_ROWS = 60;
const MIN_COLS = 16;

/** Ссылки на аргументы функций для подсказки. */
const FUNCTION_HINTS: Record<string, string> = {
  SUM: 'СУММ(число1; [число2]; …) — сумма значений',
  AVERAGE: 'СРЗНАЧ(число1; …) — среднее арифметическое',
  IF: 'ЕСЛИ(условие; значение_если_истина; значение_если_ложь)',
  COUNTIF: 'СЧЁТЕСЛИ(диапазон; критерий) — количество по условию',
  SUMIF: 'СУММЕСЛИ(диапазон; критерий; [диапазон_суммирования])',
  SUMIFS: 'СУММЕСЛИМН(диапазон_сумм; диапазон1; условие1; …)',
  VLOOKUP: 'ВПР(искомое; таблица; номер_столбца; [точное_совпадение])',
  XLOOKUP: 'ПРОСМОТРX(искомое; просматриваемый; возвращаемый; [если_не_найдено])',
  INDEX: 'ИНДЕКС(массив; номер_строки; [номер_столбца])',
  MATCH: 'ПОИСКПОЗ(искомое; массив; [тип])',
  ROUND: 'ОКРУГЛ(число; число_разрядов)',
  TEXT: 'ТЕКСТ(значение; формат)',
  CONCATENATE: 'СЦЕПИТЬ(текст1; …)',
  IFERROR: 'ЕСЛИОШИБКА(значение; значение_при_ошибке)',
  TODAY: 'СЕГОДНЯ() — текущая дата',
  COUNT: 'СЧЁТ(значение1; …) — количество чисел',
  MAX: 'МАКС(число1; …) — максимум',
  MIN: 'МИН(число1; …) — минимум',
};

export const ExcelEditor: React.FC<Props> = ({ tab, onChange, onSave, onAskAI, onAIApply }) => {
  const [sheets, setSheets] = useState<Workbook>({});
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState('');
  const [selected, setSelected] = useState<CellCoord>({ row: 0, col: 0 });
  const [selectionAnchor, setSelectionAnchor] = useState<CellCoord>({ row: 0, col: 0 });
  const [selectionEnd, setSelectionEnd] = useState<CellCoord>({ row: 0, col: 0 });
  const [extraCells, setExtraCells] = useState<CellCoord[]>([]);
  const selectingRef = useRef(false);
  const fillRef = useRef<{ row: number; col: number } | null>(null);
  const [fillTarget, setFillTarget] = useState<CellCoord | null>(null);
  const [cellMenu, setCellMenu] = useState<{ x: number; y: number } | null>(null);
  const [editingCell, setEditingCell] = useState<CellCoord | null>(null);
  const [editorValue, setEditorValue] = useState('');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [history, setHistory] = useState<Workbook[]>([]);
  const [future, setFuture] = useState<Workbook[]>([]);

  /* ---- AI-формула ---- */
  const [showFormulaAI, setShowFormulaAI] = useState(false);
  const [repairNote, setRepairNote] = useState('');
  const [aiPrompt, setAiPrompt] = useState('');
  const [pendingAction, setPendingAction] = useState<AIAction | null>(null);
  /** Ячейка, для которой запрошена формула — ждёт ответа ассистента. */
  const pendingRef = useRef<{ ref: string } | null>(null);

  const formula = useFormulaMode();
  const gridRef = useRef<HTMLDivElement>(null);
  const gridFocusedRef = useRef(false);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const cellEditorRef = useRef<HTMLInputElement>(null);

  const isCsv = tab.extension === 'csv' || tab.extension === 'tsv';

  /* ---------------- Разбор книги ---------------- */

  useEffect(() => {
    try {
      setLoading(true);
      setError('');

      const workbook = isCsv
        ? XLSX.read(tab.content, { type: 'string' })
        : XLSX.read(tab.content, { type: 'base64' });

      const parsed: Workbook = {};
      workbook.SheetNames.forEach((name) => {
        const raw: (string | number)[][] = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
          header: 1,
          defval: '',
          blankrows: true,
        });

        // sheet_to_json отдаёт кэш формулы, а не её текст. Саму формулу
        // берём из ячейки и переводим из формата файла Excel (запятая между
        // аргументами) в формат редактора (точка с запятой).
        const cells = workbook.Sheets[name];
        for (const address of Object.keys(cells)) {
          const cell = cells[address];
          if (!cell || typeof cell.f !== 'string' || !cell.f) continue;
          const pos = XLSX.utils.decode_cell(address);
          const row = raw[pos.r] ?? (raw[pos.r] = []);
          row[pos.c] = '=' + fromExcelFormula(cell.f);
        }

        const rows = Math.max(raw.length, MIN_ROWS);
        const cols = Math.max(MIN_COLS, ...raw.map((r) => r.length), 1);

        const sheetGrid: Grid = [];
        for (let r = 0; r < rows; r++) {
          const rowData: (string | number | boolean | null)[] = [];
          for (let c = 0; c < cols; c++) {
            rowData.push(raw[r]?.[c] ?? '');
          }
          sheetGrid.push(rowData);
        }
        parsed[name] = sheetGrid;
      });

      if (workbook.SheetNames.length === 0) {
        parsed['Лист 1'] = Array.from({ length: MIN_ROWS }, () => new Array(MIN_COLS).fill(''));
        setSheetNames(['Лист 1']);
        setActiveSheet('Лист 1');
      } else {
        setSheetNames(workbook.SheetNames);
        setActiveSheet(workbook.SheetNames[0]);
      }

      setSheets(parsed);
      setHistory([]);
      setFuture([]);
      setFuture([]);
    } catch (err) {
      setError(`Не удалось открыть таблицу: ${(err as Error).message}`);
    } finally {
      setLoading(false);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  const grid = sheets[activeSheet] ?? [];

  /* ---- Виртуализация строк для больших таблиц (60 FPS при 10000+ строк) ---- */
  const [scrollTop, setScrollTop] = useState(0);
  const [containerHeight, setContainerHeight] = useState(700);

  useEffect(() => {
    const el = gridRef.current;
    if (!el) return;
    const updateSize = () => {
      if (el.clientHeight > 0) setContainerHeight(el.clientHeight);
    };
    updateSize();
    const ro = new ResizeObserver(updateSize);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const ROW_HEIGHT = 28;
  const OVERSCAN = 30;
  const totalRows = grid.length;
  const isVirtual = totalRows > 80;

  const startRow = isVirtual ? Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN) : 0;
  const endRow = isVirtual
    ? Math.min(
        totalRows,
        Math.floor(scrollTop / ROW_HEIGHT) + Math.ceil(containerHeight / ROW_HEIGHT) + OVERSCAN
      )
    : totalRows;

  const topSpacerHeight = isVirtual ? startRow * ROW_HEIGHT : 0;
  const bottomSpacerHeight = isVirtual ? Math.max(0, (totalRows - endRow) * ROW_HEIGHT) : 0;

  const visibleRows = useMemo(() => {
    if (!isVirtual) {
      return grid.map((row, r) => ({ row, r }));
    }
    return grid.slice(startRow, endRow).map((row, index) => ({
      row,
      r: startRow + index,
    }));
  }, [grid, isVirtual, startRow, endRow]);

  /* ---------------- Синхронизация с содержимым вкладки ---------------- */

  const syncWorkbook = useCallback(
    (nextSheets: Workbook, names: string[]) => {
      try {
        const wb = XLSX.utils.book_new();
        names.forEach((name) => {
          const source = nextSheets[name];
          if (!source) return;

          // Обрезаем полностью пустые строки в конце, чтобы не раздувать файл.
          let lastNonEmpty = source.length - 1;
          while (
            lastNonEmpty > 0 &&
            source[lastNonEmpty].every((c) => String(c ?? '').trim() === '')
          ) {
            lastNonEmpty--;
          }

          const trimmed = source.slice(0, lastNonEmpty + 1);
          const sheet = XLSX.utils.aoa_to_sheet(trimmed);

          // Формулы проставляются после создания листа. Ячейка с формулой
          // обязана нести кэшированное значение: без него SheetJS не пишет
          // ячейку вовсе, а строка «=…» осталась бы текстом. Кэшируем
          // вычисленный результат — ровно так делает сам Excel.
          const formulaContext = createContext(nextSheets, name);
          for (let r = 0; r < trimmed.length; r++) {
            for (let c = 0; c < trimmed[r].length; c++) {
              const cell = trimmed[r][c];
              if (typeof cell !== 'string' || !cell.startsWith('=')) continue;

              const address = XLSX.utils.encode_cell({ r, c });
              const existing = sheet[address];

              // Вычисляем значение формулы для кэша
              const computed = computeCellDisplay(nextSheets, name, r, c, formulaContext);
              const num = Number(computed.replace(/\s/g, '').replace(',', '.'));
              const isNumeric = computed !== '' && !computed.startsWith('#') && Number.isFinite(num);

              const next: Record<string, unknown> = {
                t: isNumeric ? 'n' : 's',
                // В файле .xlsx Excel хранит английские имена функций и запятую
                // между аргументами. Русский Excel покажет их локализованными.
                f: toExcelFormula(cell).slice(1),
                v: isNumeric ? num : computed,
              };
              if (existing?.z) next.z = existing.z;
              if (existing?.s) next.s = existing.s;

              sheet[address] = next as unknown as XLSX.CellObject;
            }
          }

          XLSX.utils.book_append_sheet(wb, sheet, name);
        });

        if (isCsv) {
          onChange(XLSX.utils.sheet_to_csv(wb.Sheets[activeSheet] ?? wb.Sheets[names[0]]));
        } else {
          onChange(XLSX.write(wb, { bookType: 'xlsx', type: 'base64' }));
        }
      } catch (err) {
        console.error('Ошибка синхронизации книги', err);
      }
    },
    [activeSheet, isCsv, onChange]
  );

  /* ---------------- Вычисление формул ---------------- */

  // Контекст вычисления пересоздаётся при изменении книги: он кэширует
  // значения ячеек, включая результаты формул.
  const evalContext = useMemo(
    () => createContext(sheets, activeSheet),
    [sheets, activeSheet]
  );

  const displayCell = useCallback(
    (row: number, col: number): string =>
      computeCellDisplay(sheets, activeSheet, row, col, evalContext),
    [sheets, activeSheet, evalContext]
  );

  /* ---------------- Изменение ячеек ---------------- */

  const applyChange = useCallback(
    (row: number, col: number, value: string | number) => {
      setSheets((prev) => {
        const current = prev[activeSheet];
        if (!current) return prev;
        if (String(current[row]?.[col] ?? '') === String(value)) return prev;
        setHistory((h) => [...h.slice(-49), prev]);
        setFuture([]);

        const next: Workbook = { ...prev, [activeSheet]: current.map((r) => [...r]) };
        next[activeSheet][row][col] = value;
        syncWorkbook(next, sheetNames);
        return next;
      });
    },
    [activeSheet, sheetNames, syncWorkbook]
  );

  /** Применяет пачку изменений одной операцией (для действий AI). */
  const applyBatch = useCallback(
    (mutate: (draft: Workbook) => void) => {
      setSheets((prev) => {
        const next: Workbook = {};
        for (const [name, sheetGrid] of Object.entries(prev)) {
          next[name] = sheetGrid.map((r) => [...r]);
        }
        setHistory((h) => [...h.slice(-49), prev]);
        setFuture([]);
        mutate(next);
        syncWorkbook(next, sheetNames);
        return next;
      });
    },
    [sheetNames, syncWorkbook]
  );

  /** Вставляет матрицу из буфера, начиная с выбранной ячейки. */
  const insertTable = useCallback((rows: string[][]) => {
    if (!rows.length || !rows.some((row) => row.length > 0)) return;
    const start = selected;
    applyBatch((draft) => {
      const sheet = draft[activeSheet];
      if (!sheet) return;
      const neededRows = start.row + rows.length;
      const neededCols = start.col + Math.max(0, ...rows.map((row) => row.length));
      while (sheet.length < neededRows) sheet.push(new Array(sheet[0]?.length ?? MIN_COLS).fill(''));
      for (const row of sheet) while (row.length < neededCols) row.push('');
      rows.forEach((row, r) => row.forEach((value, c) => { sheet[start.row + r][start.col + c] = value; }));
    });
    setSelected({ row: start.row, col: start.col });
  }, [activeSheet, applyBatch, selected]);

  const pasteTable = useCallback(async () => {
    try {
      let html = '';
      let text = '';
      if (window.electronAPI?.readClipboard) {
        const clip = await window.electronAPI.readClipboard();
        if (!clip.success) throw new Error(clip.error || 'Не удалось прочитать буфер обмена');
        html = clip.html;
        text = clip.text;
      } else {
        text = await navigator.clipboard.readText();
      }
      const parsed = parseClipboardTable(html, text);
      if (!parsed || !parsed.rows.length) return;
      insertTable(parsed.rows);
    } catch (err) {
      console.warn('Paste table failed', err);
    }
  }, [insertTable]);

  useEffect(() => {
    const handlePaste = (event: ClipboardEvent) => {
      if (!gridFocusedRef.current) return;
      const target = event.target as HTMLElement | null;
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable)) return;
      event.preventDefault();
      const html = event.clipboardData?.getData('text/html') || '';
      const text = event.clipboardData?.getData('text/plain') || '';
      const parsed = parseClipboardTable(html, text);
      if (parsed?.rows.length) insertTable(parsed.rows);
      else void pasteTable();
    };
    document.addEventListener('paste', handlePaste);
    return () => document.removeEventListener('paste', handlePaste);
  }, [insertTable, pasteTable]);

  const undo = useCallback(() => {
    setHistory((h) => {
      if (h.length === 0) return h;
      const previous = h[h.length - 1];
      setSheets((current) => {
        setFuture((f) => [current, ...f].slice(0, 50));
        return previous;
      });
      syncWorkbook(previous, sheetNames);
      return h.slice(0, -1);
    });
  }, [sheetNames, syncWorkbook]);

  const redo = useCallback(() => {
    setFuture((f) => {
      if (f.length === 0) return f;
      const [next, ...rest] = f;
      setSheets((current) => {
        setHistory((h) => [...h, current].slice(-50));
        return next;
      });
      syncWorkbook(next, sheetNames);
      return rest;
    });
  }, [sheetNames, syncWorkbook]);

  /* ---------------- Отмена и возврат ---------------- */

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA') && target !== formulaInputRef.current) return;
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      const key = e.key.toLowerCase();
      const code = e.code;
      const undoKey = (key === 'z' || key === 'я' || code === 'KeyZ') && !e.shiftKey;
      const redoKey = key === 'y' || key === 'н' || code === 'KeyY' || ((key === 'z' || key === 'я' || code === 'KeyZ') && e.shiftKey);
      if (undoKey) {
        e.preventDefault();
        e.stopPropagation();
        undo();
      } else if (redoKey) {
        e.preventDefault();
        e.stopPropagation();
        redo();
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [undo, redo]);

  /* ---------------- Протяжка мышью для выбора диапазона ---------------- */

  useEffect(() => {
    const onMouseUp = () => {
      selectingRef.current = false;
      if (fillRef.current && fillTarget) applyFill(fillTarget);
      fillRef.current = null;
      setFillTarget(null);
      const d = formula.dragRef.current;
      if (!d.dragging) return;
      d.dragging = false;
      formula.setDrag({ ...d });

      // Завершаем вставку ссылки на выбранный диапазон.
      // Если пользователь просто кликнул (anchor === focus), ссылка уже вставлена
      // на mousedown — повторная вставка сдвинула бы каретку и следующая ссылка
      // дописалась бы не туда («=A8+A9A10» вместо «=A8+A9+A10»).
      if (d.anchor && d.focus) {
        const dragged =
          d.anchor.row !== d.focus.row || d.anchor.col !== d.focus.col;
        if (dragged) {
          const ref = rangeToA1(normalizeRange(d.anchor, d.focus));
          formula.insertReference(ref, { replacePrevious: true });
        }
      }
    };

    document.addEventListener('mouseup', onMouseUp);
    return () => document.removeEventListener('mouseup', onMouseUp);
  }, [formula]);

  /* ---------------- Строка формул ---------------- */

  const cellRaw = grid[selected.row]?.[selected.col];
  const cellText = cellRaw === undefined || cellRaw === null ? '' : String(cellRaw);

  // Строка формул следует за выбранной ячейкой, если пользователь её не правит.
  // Условие опирается на формульный режим правки, а не на текст черновика:
  // иначе строка формул с уже введённой формулой не сбрасывалась при переходе
  // на другую ячейку и «переносила» формулу за выделением.
  useEffect(() => {
    if (!formula.isFormulaMode) {
      formula.setDraft(cellText);
      formula.setCaret(cellText.length);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected.row, selected.col, cellText, activeSheet, formula.isFormulaMode]);

  const currentRef = `${columnName(selected.col)}${selected.row + 1}`;
  const selectionRange = normalizeRange(selectionAnchor, selectionEnd);
  const selectionLabel = rangeToA1(selectionRange);

  const selectSingle = (cell: CellCoord) => {
    setSelected(cell);
    setSelectionAnchor(cell);
    setSelectionEnd(cell);
    setExtraCells([]);
  };

  const cellInSelection = (row: number, col: number) => {
    const inRange = row >= selectionRange.start.row && row <= selectionRange.end.row
      && col >= selectionRange.start.col && col <= selectionRange.end.col;
    return inRange || extraCells.some((cell) => cell.row === row && cell.col === col);
  };

  const applyFill = (target: CellCoord) => {
    const source = selectionRange;
    const rows = source.end.row - source.start.row + 1;
    const cols = source.end.col - source.start.col + 1;
    const down = target.row - source.end.row;
    const right = target.col - source.end.col;
    if (down <= 0 && right <= 0) return;
    const vertical = down >= right;
    applyBatch((draft) => {
      const sheet = draft[activeSheet];
      if (!sheet) return;
      const endRow = vertical ? source.end.row + down : source.end.row;
      const endCol = vertical ? source.end.col : source.end.col + right;
      while (sheet.length <= endRow) sheet.push(new Array(sheet[0]?.length ?? MIN_COLS).fill(''));
      for (const line of sheet) while (line.length <= endCol) line.push('');
      const copies = vertical ? Math.ceil(down / rows) : Math.ceil(right / cols);
      for (let copy = 1; copy <= copies; copy++) {
        for (let row = source.start.row; row <= source.end.row; row++) {
          for (let col = source.start.col; col <= source.end.col; col++) {
            const targetRow = vertical ? row + rows * copy : row;
            const targetCol = vertical ? col : col + cols * copy;
            if (targetRow > endRow || targetCol > endCol) continue;
            const value = sheet[row][col];
            sheet[targetRow][targetCol] = typeof value === 'string' && value.startsWith('=')
              ? shiftFormula(value, targetRow - row, targetCol - col)
              : value;
          }
        }
      }
    });
    setSelectionEnd(vertical
      ? { row: source.end.row + down, col: source.end.col }
      : { row: source.end.row, col: source.end.col + right });
  };

  const clearSelection = () => {
    applyBatch((draft) => {
      const sheet = draft[activeSheet];
      if (!sheet) return;
      for (let row = selectionRange.start.row; row <= selectionRange.end.row; row++) {
        for (let col = selectionRange.start.col; col <= selectionRange.end.col; col++) {
          if (sheet[row]) sheet[row][col] = '';
        }
      }
      for (const cell of extraCells) {
        if (sheet[cell.row]) sheet[cell.row][cell.col] = '';
      }
    });
  };

  const deleteSelectionRows = () => {
    const indexes = new Set<number>();
    for (let row = selectionRange.start.row; row <= selectionRange.end.row; row++) indexes.add(row);
    extraCells.forEach((cell) => indexes.add(cell.row));
    applyBatch((draft) => {
      const current = draft[activeSheet];
      if (!current || current.length - indexes.size < 1) return;
      draft[activeSheet] = current.filter((_, index) => !indexes.has(index));
    });
    selectSingle({ row: Math.min(selectionRange.start.row, grid.length - 2), col: selectionRange.start.col });
  };

  const deleteSelectionColumns = () => {
    const indexes = new Set<number>();
    for (let col = selectionRange.start.col; col <= selectionRange.end.col; col++) indexes.add(col);
    extraCells.forEach((cell) => indexes.add(cell.col));
    applyBatch((draft) => {
      const current = draft[activeSheet];
      if (!current || (current[0]?.length ?? 0) - indexes.size < 1) return;
      draft[activeSheet] = current.map((line) => line.filter((_, index) => !indexes.has(index)));
    });
    selectSingle({ row: selectionRange.start.row, col: Math.max(0, selectionRange.start.col - (indexes.size > 1 ? 0 : 0)) });
  };

  const copySelection = async () => {
    const lines: string[] = [];
    for (let row = selectionRange.start.row; row <= selectionRange.end.row; row++) {
      const values: string[] = [];
      for (let col = selectionRange.start.col; col <= selectionRange.end.col; col++) {
        values.push(displayCell(row, col));
      }
      lines.push(values.join('\t'));
    }
    for (const cell of extraCells) {
      const inside = cell.row >= selectionRange.start.row && cell.row <= selectionRange.end.row
        && cell.col >= selectionRange.start.col && cell.col <= selectionRange.end.col;
      if (!inside) lines.push(displayCell(cell.row, cell.col));
    }
    const text = lines.join('\n');
    if (window.electronAPI?.writeClipboard) await window.electronAPI.writeClipboard(text);
    else await navigator.clipboard.writeText(text);
  };

  /* ---------------- Навигация с клавиатуры ---------------- */

  const handleGridKeyDown = (e: React.KeyboardEvent) => {
    if (editingCell) return;
    const { row, col } = selected;
    const maxRow = grid.length - 1;
    const maxCol = (grid[0]?.length ?? 1) - 1;

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      e.preventDefault();
      void copySelection();
      return;
    }
    const move = (next: CellCoord) => {
      if (e.shiftKey) {
        setSelectionEnd(next);
        setSelected(next);
      } else {
        selectSingle(next);
      }
    };
    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        move({ row: Math.max(0, row - 1), col });
        break;
      case 'ArrowDown':
        e.preventDefault();
        move({ row: Math.min(maxRow, row + 1), col });
        break;
      case 'ArrowLeft':
        e.preventDefault();
        move({ row, col: Math.max(0, col - 1) });
        break;
      case 'ArrowRight':
        e.preventDefault();
        move({ row, col: Math.min(maxCol, col + 1) });
        break;
      case 'Tab':
        e.preventDefault();
        setSelected({ row, col: Math.min(maxCol, col + (e.shiftKey ? -1 : 1)) });
        break;
      case 'Enter':
      case 'F2':
        e.preventDefault();
        // F2 и Enter открывают редактор с текущим значением/формулой
        setEditorValue(cellText);
        setEditingCell({ row, col });
        break;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        applyChange(row, col, '');
        // Сбрасываем состояние формулы, чтобы она не вернулась при следующем фокусе
        formula.setDraft('');
        formula.setCaret(0);
        break;
      default:
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
          // Символ начинает редактирование и становится первым значением ячейки.
          // Без preventDefault то же событие доходит до появившегося поля ввода
          // и цифра записывается дважды: «1» превращалось в «11».
          e.preventDefault();
          setEditorValue(e.key);
          setEditingCell({ row, col });
          if (e.key === '=') {
            formula.setDraft('=');
            formula.setCaret(1);
            formula.beginEdit();
          }
        }
    }
  };

  /* ---------------- Операции со структурой ---------------- */

  const addRow = () => {
    applyBatch((draft) => {
      const current = draft[activeSheet];
      const cols = current[0]?.length ?? MIN_COLS;
      draft[activeSheet] = [...current, new Array(cols).fill('')];
    });
  };

  const addColumn = () => {
    applyBatch((draft) => {
      draft[activeSheet] = draft[activeSheet].map((r) => [...r, '']);
    });
  };

  const deleteRow = (rowIdx: number) => {
    applyBatch((draft) => {
      const current = draft[activeSheet];
      if (current.length <= 1) return;
      draft[activeSheet] = current.filter((_, i) => i !== rowIdx);
    });
    setSelected((s) => ({ ...s, row: Math.max(0, Math.min(s.row, grid.length - 2)) }));
  };

  const deleteColumn = (colIdx: number) => {
    applyBatch((draft) => {
      const current = draft[activeSheet];
      if ((current[0]?.length ?? 0) <= 1) return;
      draft[activeSheet] = current.map((r) => r.filter((_, i) => i !== colIdx));
    });
    setSelected((s) => ({ ...s, col: Math.max(0, Math.min(s.col, (grid[0]?.length ?? 2) - 2)) }));
  };

  const addSheet = () => {
    const existing = new Set(sheetNames);
    let n = sheetNames.length + 1;
    while (existing.has(`Лист ${n}`)) n++;
    const name = `Лист ${n}`;
    const newGrid: Grid = Array.from({ length: MIN_ROWS }, () => new Array(MIN_COLS).fill(''));

    setSheets((prev) => {
      const next = { ...prev, [name]: newGrid };
      syncWorkbook(next, [...sheetNames, name]);
      return next;
    });
    setSheetNames((prev) => [...prev, name]);
    setActiveSheet(name);
    setSelected({ row: 0, col: 0 });
  };

  const renameSheet = (oldName: string, newName: string) => {
    const trimmed = newName.trim();
    if (!trimmed || trimmed === oldName || sheetNames.includes(trimmed)) return;
    setSheets((prev) => {
      const next: Workbook = { ...prev, [trimmed]: prev[oldName] };
      delete next[oldName];
      const names = sheetNames.map((n) => (n === oldName ? trimmed : n));
      syncWorkbook(next, names);
      return next;
    });
    setSheetNames((prev) => prev.map((n) => (n === oldName ? trimmed : n)));
    setActiveSheet((cur) => (cur === oldName ? trimmed : cur));
  };

  /** Формулы, которые считаются с ошибкой, по всем листам книги. */
  const formulaErrors = useMemo(() => {
    const found: { sheet: string; ref: string; formula: string; error: string }[] = [];
    for (const name of sheetNames) {
      const sheetGrid = sheets[name];
      if (!sheetGrid) continue;
      const context = name === activeSheet ? evalContext : createContext(sheets, name);
      for (let row = 0; row < sheetGrid.length; row++) {
        const line = sheetGrid[row];
        for (let col = 0; col < line.length; col++) {
          const raw = line[col];
          if (typeof raw !== 'string' || !raw.startsWith('=')) continue;
          const shown = computeCellDisplay(sheets, name, row, col, context);
          if (shown.startsWith('#')) {
            found.push({ sheet: name, ref: `${columnName(col)}${row + 1}`, formula: raw, error: shown });
          }
        }
      }
    }
    return found;
  }, [sheets, sheetNames, activeSheet, evalContext]);

  /**
   * Исправляет всё, что чинится без догадок: русские имена функций,
   * пробел перед скобкой, непарные скобки и опечатки вроде СУМА → SUM.
   * Ошибки данных (#DIV/0!, #N/A, #REF!) не трогает — их нельзя угадать.
   */
  const repairAllErrors = useCallback(() => {
    let fixed = 0;
    const skipped: string[] = [];

    applyBatch((draft) => {
      for (const name of sheetNames) {
        const sheetGrid = draft[name];
        if (!sheetGrid) continue;
        const context = createContext(draft, name);
        for (let row = 0; row < sheetGrid.length; row++) {
          const line = sheetGrid[row];
          for (let col = 0; col < line.length; col++) {
            const raw = line[col];
            if (typeof raw !== 'string' || !raw.startsWith('=')) continue;
            const shown = computeCellDisplay(draft, name, row, col, context);
            if (!shown.startsWith('#')) continue;

            const repair = repairFormula(raw);
            if (!repair) {
              skipped.push(`${columnName(col)}${row + 1}: ${shown}`);
              continue;
            }
            line[col] = repair.formula;
            context.cache.clear();
            const after = computeCellDisplay(draft, name, row, col, context);
            if (after.startsWith('#')) {
              line[col] = raw;
              context.cache.clear();
              skipped.push(`${columnName(col)}${row + 1}: ${shown}`);
              continue;
            }
            fixed++;
          }
        }
      }
    });

    if (fixed === 0 && skipped.length === 0) {
      setRepairNote('Ошибок в формулах нет');
    } else if (skipped.length === 0) {
      setRepairNote(`Исправлено формул: ${fixed}`);
    } else {
      const rest = skipped.slice(0, 4).join(', ');
      const more = skipped.length > 4 ? ` и ещё ${skipped.length - 4}` : '';
      setRepairNote(
        fixed > 0
          ? `Исправлено: ${fixed}. Остались ошибки данных — ${rest}${more}`
          : `Автоматически исправить нельзя: ${rest}${more}`
      );
    }
  }, [applyBatch, sheetNames]);

  /* ---------------- Экспорт CSV ---------------- */

  const exportCsv = () => {
    // Экспортируем вычисленные значения, а не формулы — как «Сохранить как CSV»
    const computed = grid.map((row, r) =>
      row.map((_cell, c) => displayCell(r, c))
    );
    const csv = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(computed));
    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${tab.name.replace(/\.[^.]+$/, '')}_${activeSheet}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  /* ============================================================
     Действия AI: разбор ответа и применение
     ============================================================ */

  /**
   * Извлекает из ответа модели адрес и формулу.
   * Ожидается строгая форма, которую запрашивает buildFormulaPrompt:
   *   ЯЧЕЙКА: B5
   *   ФОРМУЛА: =СУММ(B2:B4)
   */
  const parseAIFormula = useCallback(
    (answer: string): AIAction | null => {
      const cellMatch = answer.match(/ЯЧЕЙКА\s*:\s*([A-Za-zА-Яа-яЁё0-9_$!:' ]+)/i);
      const formulaMatch = answer.match(/ФОРМУЛА\s*:\s*(.+)/i);

      if (!formulaMatch) return null;

      let value = formulaMatch[1].trim();
      // Убираем обрамляющие кавычки/бэктики и хвостовые пояснения
      value = value.replace(/^[`"']+|[`"']+$/g, '').trim();
      value = value.split(/\s+#/)[0].trim();
      if (value && !value.startsWith('=')) value = `=${value}`;

      const target = cellMatch?.[1]?.trim().replace(/\s/g, '') || currentRef;

      return {
        kind: 'setCell',
        target,
        value,
        sheet: activeSheet,
        description: `Вставить формулу в ${target}`,
      };
    },
    [activeSheet, currentRef]
  );

  /** Разбор адреса/диапазона в координаты. */
  const parseTarget = useCallback(
    (target: string): { sheet: string; range: { start: CellCoord; end: CellCoord } } | null => {
      let sheet = activeSheet;
      let ref = target.trim();

      const bang = ref.lastIndexOf('!');
      if (bang !== -1) {
        sheet = ref.slice(0, bang).replace(/^'|'$/g, '');
        ref = ref.slice(bang + 1);
      }

      const parseOne = (s: string): CellCoord | null => {
        const m = s.replace(/\$/g, '').match(/^([A-Za-z]{1,3})(\d{1,7})$/);
        if (!m) return null;
        let col = 0;
        const letters = m[1].toUpperCase();
        for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
        return { row: parseInt(m[2], 10) - 1, col: col - 1 };
      };

      const parts = ref.split(':');
      const start = parseOne(parts[0]);
      const end = parts.length > 1 ? parseOne(parts[1]) : start;
      if (!start || !end) return null;

      return { sheet, range: normalizeRange(start, end) };
    },
    [activeSheet]
  );

  /** Применяет подтверждённое действие к таблице. */
  const applyAIAction = useCallback(
    (action: AIAction) => {
      const parsed = parseTarget(action.target);
      if (!parsed) {
        onAIApply?.(action, `Не удалось разобрать адрес «${action.target}»`);
        setPendingAction(null);
        return;
      }

      const targetSheet = sheets[parsed.sheet] ? parsed.sheet : activeSheet;
      const { start, end } = parsed.range;
      const normalized = normalizeFormula(action.value);

      applyBatch((draft) => {
        const sheet = draft[targetSheet];
        if (!sheet) return;

        // Расширяем лист, если адрес выходит за текущие границы
        while (sheet.length <= end.row) {
          sheet.push(new Array(sheet[0]?.length ?? MIN_COLS).fill(''));
        }
        for (const row of sheet) {
          while (row.length <= end.col) row.push('');
        }

        if (start.row === end.row && start.col === end.col) {
          sheet[start.row][start.col] = normalized;
        } else {
          // Заполняем диапазон: формула с относительными ссылками сдвигается
          for (let r = start.row; r <= end.row; r++) {
            for (let c = start.col; c <= end.col; c++) {
              sheet[r][c] =
                start.row === end.row && start.col === end.col
                  ? normalized
                  : shiftFormula(normalized, r - start.row, c - start.col);
            }
          }
        }
      });

      setActiveSheet(targetSheet);
      setSelected({ row: start.row, col: start.col });
      onAIApply?.(
        action,
        `Формула применена к ${action.target} на листе «${targetSheet}»`
      );
      setPendingAction(null);
    },
    [activeSheet, applyBatch, onAIApply, parseTarget, sheets]
  );

  /* ---------------- Запрос формулы у AI ---------------- */

  const requestFormula = () => {
    if (!aiPrompt.trim() || !onAskAI) return;

    const headerRow = grid[0]?.slice(0, 12).map((c) => String(c ?? '')).join(' | ') ?? '';
    const sample = grid
      .slice(1, 6)
      .map((r) => r.slice(0, 12).map((c) => String(c ?? '')).join(' | '))
      .join('\n');

    // Жёсткий формат ответа — чтобы разобрать результат без догадок
    const prompt =
      `Сгенерируй формулу Excel для ячейки ${currentRef} на листе «${activeSheet}».\n` +
      `Задача: "${aiPrompt}".\n\n` +
      `Ответь СТРОГО в таком формате, без пояснений до и после:\n` +
      `ЯЧЕЙКА: ${currentRef}\n` +
      `ФОРМУЛА: =<формула>\n\n` +
      `Требования: используй английские или русские имена функций Excel; ` +
      `разделитель аргументов — точка с запятой или запятая; ` +
      `формула должна начинаться со знака «=».`;

    const context =
      `Лист: ${activeSheet}\nЯчейка: ${currentRef}\n` +
      `Заголовки: ${headerRow}\nПример данных:\n${sample}`;

    setShowFormulaAI(false);
    setAiPrompt('');
    // Запоминаем целевую ячейку: ответ придёт асинхронно, через панель AI
    pendingRef.current = { ref: currentRef };
    onAskAI(prompt, context);
  };

  /** Обрабатывает ответ модели, когда он приходит из App. */
  const handleAIAnswer = useCallback(
    (answer: string, forRef: string) => {
      const action = parseAIFormula(answer);
      if (action) {
        setPendingAction({ ...action, target: action.target || forRef });
      }
    },
    [parseAIFormula]
  );

  /**
   * Запрос к AI приходит асинхронно, поэтому ожидающая ячейка хранится в ref,
   * а не в состоянии: подписка на событие не пересоздаётся на каждый ответ.
   */
  useEffect(() => {
    const handler = (e: Event) => {
      const detail = (e as CustomEvent).detail as { tabId?: string; answer?: string } | undefined;
      if (!detail?.answer) return;
      if (detail.tabId && detail.tabId !== tab.id) return;

      const pending = pendingRef.current;
      if (!pending) return;
      pendingRef.current = null;

      handleAIAnswer(detail.answer, pending.ref);
    };

    window.addEventListener('workspace:ai-answer', handler);
    return () => window.removeEventListener('workspace:ai-answer', handler);
  }, [handleAIAnswer, tab.id]);

  /* ---------------- Статистика ---------------- */

  const stats = useMemo(() => {
    let filled = 0;
    for (const row of grid) {
      for (const cell of row) {
        if (String(cell ?? '').trim() !== '') filled++;
      }
    }
    return { filled, rows: grid.length, cols: grid[0]?.length ?? 0 };
  }, [grid]);

  /** Ячейки, содержащие формулы — для индикатора. */
  const formulaCount = useMemo(() => {
    let count = 0;
    for (const row of grid) {
      for (const cell of row) {
        if (typeof cell === 'string' && cell.startsWith('=')) count++;
      }
    }
    return count;
  }, [grid]);

  if (loading) {
    return (
      <div className="flex h-full items-center justify-center bg-ws-editor">
        <div className="flex flex-col items-center gap-2 text-ws-muted">
          <FileSpreadsheet size={28} className="animate-pulse text-emerald-500" />
          <span className="flex items-center gap-2 text-xs">
            <Spinner size={12} />
            Загрузка таблицы…
          </span>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex h-full items-center justify-center bg-ws-editor p-6">
        <div className="max-w-md rounded-lg border border-red-500/40 bg-red-500/10 p-5 text-center">
          <AlertTriangle size={22} className="mx-auto mb-2 text-ws-danger" />
          <p className="mb-1 text-sm font-semibold text-ws-danger">Не удалось открыть таблицу</p>
          <p className="text-xs text-ws-muted select-text">{error}</p>
        </div>
      </div>
    );
  }

  return (
    <div className="relative flex h-full w-full flex-col bg-ws-editor text-ws" onMouseDown={() => setCellMenu(null)}>
      {cellMenu && (
        <div
          className="fixed z-[80] min-w-52 rounded-lg border border-ws bg-ws-surface p-1 text-xs shadow-ws"
          style={{ left: cellMenu.x, top: cellMenu.y }}
          onMouseDown={(e) => e.stopPropagation()}
        >
          <button className="block w-full rounded px-3 py-1.5 text-left hover:bg-ws-hover" onClick={() => { void copySelection(); setCellMenu(null); }}>Копировать</button>
          <button className="block w-full rounded px-3 py-1.5 text-left hover:bg-ws-hover" onClick={() => { void pasteTable(); setCellMenu(null); }}>Вставить</button>
          <button className="block w-full rounded px-3 py-1.5 text-left hover:bg-ws-hover" onClick={() => { clearSelection(); setCellMenu(null); }}>Очистить содержимое</button>
          <div className="my-1 border-t border-ws" />
          <button className="block w-full rounded px-3 py-1.5 text-left text-ws-danger hover:bg-red-500/10" onClick={() => { deleteSelectionRows(); setCellMenu(null); }}>Удалить строки</button>
          <button className="block w-full rounded px-3 py-1.5 text-left text-ws-danger hover:bg-red-500/10" onClick={() => { deleteSelectionColumns(); setCellMenu(null); }}>Удалить столбцы</button>
        </div>
      )}
      {/* Панель инструментов */}
      <div className="flex shrink-0 flex-wrap items-center justify-between gap-2 border-b border-ws bg-ws-surface px-3 py-1.5 text-xs select-none">
        <div className="flex items-center gap-1.5">
          <span className="flex items-center gap-1.5 rounded border border-emerald-500/40 bg-emerald-500/10 px-2 py-0.5 font-medium text-emerald-500">
            <FileSpreadsheet size={13} />
            Таблица
          </span>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <button onClick={undo} disabled={history.length === 0} title="Отменить (Ctrl+Z)" className="rounded p-1 text-ws-muted hover:bg-ws-hover disabled:opacity-40"><Undo2 size={13} /></button>
          <button onClick={redo} disabled={future.length === 0} title="Вернуть (Ctrl+Y)" className="rounded p-1 text-ws-muted hover:bg-ws-hover disabled:opacity-40"><Redo2 size={13} /></button>
          <button
            onClick={() => void pasteTable()}
            title="Вставить таблицу из буфера (Ctrl+V)"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-ws-accent hover:text-ws-accent"
          >
            <ClipboardPaste size={11} />
            Вставить
          </button>
          <button
            onClick={addRow}
            title="Добавить строку"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-ws-accent hover:text-ws-accent"
          >
            <Plus size={11} />
            Строка
          </button>
          <button
            onClick={addColumn}
            title="Добавить столбец"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-ws-accent hover:text-ws-accent"
          >
            <Plus size={11} />
            Столбец
          </button>
          <button
            onClick={() => deleteRow(selected.row)}
            title="Удалить текущую строку"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-red-500 hover:text-ws-danger"
          >
            <Trash2 size={11} />
          </button>
          <button
            onClick={() => deleteColumn(selected.col)}
            title="Удалить текущий столбец"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-red-500 hover:text-ws-danger"
          >
            <TableIcon size={11} />
          </button>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <button
            onClick={exportCsv}
            title="Экспортировать активный лист в CSV (вычисленные значения)"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-ws-accent hover:text-ws-accent"
          >
            <Download size={11} />
            CSV
          </button>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={repairAllErrors}
            disabled={formulaErrors.length === 0}
            title={
              formulaErrors.length > 0
                ? `Исправить ошибки формул (${formulaErrors.length}): русские имена, опечатки, скобки`
                : 'Ошибок в формулах нет'
            }
            className={cn(
              'flex items-center gap-1.5 rounded border px-2.5 py-1 transition-colors',
              formulaErrors.length > 0
                ? 'border-red-500/40 bg-red-500/10 text-ws-danger hover:bg-red-500/20'
                : 'border-ws text-ws-faint'
            )}
          >
            <Wand2 size={12} />
            Исправить все ошибки
            {formulaErrors.length > 0 && (
              <span className="rounded bg-ws-danger px-1 text-[10px] font-semibold text-white">
                {formulaErrors.length}
              </span>
            )}
          </button>

          <button
            onClick={() => setShowFormulaAI(true)}
            className="flex items-center gap-1.5 rounded border border-emerald-500/40 bg-emerald-500/10 px-2.5 py-1 text-emerald-500 transition-colors hover:bg-emerald-500/20"
          >
            <Sparkles size={12} />
            AI-формула
          </button>

          <div className="h-3.5 w-px border-l border-ws" />

          <button
            onClick={onSave}
            disabled={!tab.isDirty}
            title="Сохранить (Ctrl+S)"
            className={cn(
              'flex items-center gap-1 rounded px-2.5 py-1 font-medium transition-colors',
              tab.isDirty ? 'bg-ws-accent text-white hover:opacity-90' : 'bg-ws-panel-alt text-ws-faint'
            )}
          >
            <Save size={12} />
            Сохранить
          </button>
        </div>
      </div>

      {repairNote && (
        <div className="flex shrink-0 items-center justify-between gap-3 border-b border-ws bg-ws-elevated px-3 py-1 text-[11px] text-ws-muted">
          <span className="truncate">{repairNote}</span>
          <button onClick={() => setRepairNote('')} className="shrink-0 text-ws-faint hover:text-ws" title="Скрыть">
            <X size={12} />
          </button>
        </div>
      )}

      {/* Строка формул */}
      <div className="relative flex shrink-0 items-center gap-2 border-b border-ws bg-ws-elevated px-3 py-1.5 text-xs">
        <div className="min-w-[62px] rounded border border-ws bg-ws-surface px-2 py-1 text-center font-mono font-bold text-emerald-500">
          {selectionLabel}
        </div>
        <Sigma size={13} className="shrink-0 text-ws-faint" />

        <input
          ref={formulaInputRef}
          type="text"
          value={formula.draft}
          onChange={(e) => {
            formula.setDraft(e.target.value);
            formula.setCaret(e.target.selectionStart ?? e.target.value.length);
            // Правка в строке формул: клики по ячейкам подставляют ссылки.
            formula.beginEdit();
          }}
          onFocus={() => formula.beginEdit()}
          onKeyUp={(e) => formula.setCaret((e.target as HTMLInputElement).selectionStart ?? 0)}
          onClick={(e) => formula.setCaret((e.target as HTMLInputElement).selectionStart ?? 0)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              applyChange(selected.row, selected.col, formula.draft);
              formula.endEdit();
              setSelected((s) => ({ ...s, row: Math.min(grid.length - 1, s.row + 1) }));
              (e.target as HTMLInputElement).blur();
            } else if (e.key === 'Escape') {
              e.preventDefault();
              formula.setDraft(cellText);
              formula.endEdit();
              (e.target as HTMLInputElement).blur();
            } else if (e.key === 'Tab' && formula.suggestions.length > 0) {
              e.preventDefault();
              formula.applySuggestion(formula.suggestions[0]);
            }
          }}
          placeholder="Значение или формула, например =SUM(B2:B20)"
          spellCheck={false}
          className="flex-1 rounded border border-ws bg-ws-surface px-2.5 py-1 font-mono text-ws placeholder:text-ws-faint focus:border-emerald-500 focus:outline-none"
        />

        {/* Подсказка по аргументам активной функции */}
        {formula.openFunction && FUNCTION_HINTS[normalizeFormula(`${formula.openFunction.name}(`).replace('(', '')] && (
          <div className="absolute top-full left-24 z-30 mt-0.5 max-w-[70%] rounded border border-ws bg-ws-surface px-2 py-1 font-mono text-[10px] text-ws-accent shadow-ws">
            {FUNCTION_HINTS[normalizeFormula(`${formula.openFunction.name}(`).replace('(', '')]}
          </div>
        )}

        {/* Автодополнение имён функций */}
        {formula.suggestions.length > 0 && (
          <div className="absolute top-full right-3 z-40 mt-0.5 w-72 overflow-hidden rounded-lg border border-ws bg-ws-surface shadow-ws">
            {formula.suggestions.map((s, i) => (
              <button
                key={s.name + s.label}
                onMouseDown={(e) => {
                  e.preventDefault();
                  formula.applySuggestion(s);
                  formulaInputRef.current?.focus();
                }}
                className={cn(
                  'flex w-full items-center justify-between gap-2 px-2.5 py-1.5 text-left font-mono text-[11px] transition-colors hover:bg-ws-hover',
                  i === 0 ? 'text-ws-accent' : 'text-ws-muted'
                )}
              >
                <span>{s.label}</span>
                {s.ru && s.ru !== s.label && (
                  <span className="text-[10px] text-ws-faint">{s.name}</span>
                )}
              </button>
            ))}
            <div className="border-t border-ws px-2.5 py-1 text-[10px] text-ws-faint">
              Tab — подставить
            </div>
          </div>
        )}
      </div>

      {/* Сетка */}
      <div
        ref={gridRef}
        tabIndex={0}
        onFocus={() => { gridFocusedRef.current = true; }}
        onBlur={() => { gridFocusedRef.current = false; }}
        onKeyDown={handleGridKeyDown}
        onScroll={(e) => {
          if (isVirtual) setScrollTop(e.currentTarget.scrollTop);
        }}
        className="min-h-0 flex-1 overflow-auto bg-ws-editor focus:outline-none"
      >
        <table className="border-collapse select-text">
          <thead>
            <tr className="sticky top-0 z-10 bg-ws-surface text-[11px] font-mono text-ws-faint">
              <th className="sticky left-0 z-20 h-7 w-12 border border-ws bg-ws-surface text-center">
                #
              </th>
              {(grid[0] ?? []).map((_, c) => (
                <th
                  key={c}
                  onClick={() => setSelected({ row: 0, col: c })}
                  className={cn(
                    'h-7 min-w-[104px] cursor-pointer border border-ws px-2 text-center font-medium transition-colors hover:bg-ws-hover',
                    selected.col === c && 'bg-emerald-500/15 text-emerald-500'
                  )}
                >
                  {columnName(c)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {topSpacerHeight > 0 && (
              <tr style={{ height: `${topSpacerHeight}px` }}>
                <td colSpan={(grid[0]?.length ?? 0) + 1} style={{ padding: 0, border: 'none' }} />
              </tr>
            )}
            {visibleRows.map(({ row, r }) => (
              <tr key={r}>
                <td
                  onClick={() => setSelected({ row: r, col: 0 })}
                  className={cn(
                    'sticky left-0 z-10 h-7 w-12 cursor-pointer border border-ws bg-ws-surface text-center font-mono text-[11px] transition-colors hover:bg-ws-hover',
                    selected.row === r ? 'bg-emerald-500/15 font-bold text-emerald-500' : 'text-ws-faint'
                  )}
                >
                  {r + 1}
                </td>

                {row.map((cell, c) => {
                  const isActive = selected.row === r && selected.col === c;
                  const isSelected = cellInSelection(r, c);
                  const inFill = Boolean(fillTarget && (
                    (fillTarget.row > selectionRange.end.row && r > selectionRange.end.row && r <= fillTarget.row
                      && c >= selectionRange.start.col && c <= selectionRange.end.col)
                    || (fillTarget.col > selectionRange.end.col && c > selectionRange.end.col && c <= fillTarget.col
                      && r >= selectionRange.start.row && r <= selectionRange.end.row)
                  ));
                  const isEditing = editingCell?.row === r && editingCell?.col === c;
                  const raw = cell === undefined || cell === null ? '' : String(cell);
                  const display = displayCell(r, c);
                  const isFormula = raw.startsWith('=');

                  // Подсветка диапазона, который сейчас выбирается мышью
                  const d = formula.drag;
                  let inDragRange = false;
                  if (d.anchor && d.focus) {
                    const nr = normalizeRange(d.anchor, d.focus);
                    inDragRange =
                      r >= nr.start.row && r <= nr.end.row && c >= nr.start.col && c <= nr.end.col;
                  }

                  return (
                    <td
                      key={c}
                      onMouseDown={(e) => {
                        // В формульном режиме клик по ячейке подставляет ссылку
                        if (formula.isFormulaMode && !isEditing) {
                          e.preventDefault();
                          formula.dragRef.current = {
                            anchor: { row: r, col: c },
                            focus: { row: r, col: c },
                            dragging: true,
                          };
                          formula.setDrag({ ...formula.dragRef.current });
                          formula.insertReference(`${columnName(c)}${r + 1}`);
                          return;
                        }
                        const cell = { row: r, col: c };
                        if (e.ctrlKey || e.metaKey) {
                          setSelected(cell);
                          setExtraCells((prev) => prev.some((item) => item.row === r && item.col === c)
                            ? prev.filter((item) => item.row !== r || item.col !== c)
                            : prev.concat(cell));
                        } else if (e.shiftKey) {
                          setSelectionEnd(cell);
                          setSelected(cell);
                        } else {
                          selectSingle(cell);
                          selectingRef.current = true;
                        }
                        gridRef.current?.focus();
                      }}
                      onContextMenu={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        if (!cellInSelection(r, c)) selectSingle({ row: r, col: c });
                        setCellMenu({ x: e.clientX, y: e.clientY });
                        gridRef.current?.focus();
                      }}
                      onMouseEnter={(e) => {
                        if (fillRef.current && (e.buttons & 1)) {
                          setFillTarget({ row: r, col: c });
                          return;
                        }
                        if (selectingRef.current && (e.buttons & 1) && !formula.isFormulaMode) {
                          setSelectionEnd({ row: r, col: c });
                          setSelected({ row: r, col: c });
                          return;
                        }
                        // Протяжка: обновляем конец выделения без перерисовки на каждый шаг
                        const dr = formula.dragRef.current;
                        if (formula.isFormulaMode && dr.dragging && dr.anchor && (e.buttons & 1)) {
                          dr.focus = { row: r, col: c };
                          formula.setDrag({ ...dr });
                        }
                      }}
                      onDoubleClick={() => {
                        setEditorValue(raw);
                        setEditingCell({ row: r, col: c });
                        if (raw.startsWith('=')) {
                          formula.setDraft(raw);
                          formula.setCaret(raw.length);
                        }
                      }}
                      title={isFormula ? `${raw}  →  ${display}` : display}
                      className={cn(
                        'relative h-7 max-w-[240px] min-w-[104px] cursor-cell truncate border border-ws px-2 text-xs',
                        isSelected && 'bg-emerald-500/15',
                        inFill && 'bg-emerald-500/10 outline outline-1 outline-emerald-500/50',
                        isActive && 'ring-2 ring-emerald-500 ring-inset',
                        inDragRange && 'bg-emerald-500/10',
                        !isSelected && !inDragRange && 'hover:bg-ws-hover'
                      )}
                    >
                      {isActive && !isEditing && !formula.isFormulaMode && (
                        <span
                          title="Потяните, чтобы заполнить соседние ячейки"
                          onMouseDown={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            fillRef.current = { row: r, col: c };
                            setFillTarget({ row: r, col: c });
                          }}
                          className="absolute -right-1 -bottom-1 z-20 h-2.5 w-2.5 cursor-crosshair border border-white bg-emerald-500"
                        />
                      )}
                      {isEditing ? (
                        <input
                          ref={cellEditorRef}
                          autoFocus
                          type="text"
                          defaultValue={editorValue}
                          onBlur={(e) => {
                            applyChange(r, c, e.target.value);
                            setEditingCell(null);
                            // Правка завершена: строка формул снова следует за выделением.
                            formula.endEdit();
                          }}
                          onChange={(e) => {
                            if (e.target.value.startsWith('=')) {
                              formula.setDraft(e.target.value);
                              formula.setCaret(e.target.selectionStart ?? e.target.value.length);
                              formula.beginEdit();
                            } else {
                              // Ячейка больше не содержит формулу — выходим из режима.
                              formula.endEdit();
                            }
                          }}
                          onKeyDown={(e) => {
                            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
                              // Фиксируем набранное до того, как окно выполнит сохранение.
                              applyChange(r, c, e.currentTarget.value);
                              setEditingCell(null);
                              formula.endEdit();
                              return;
                            }
                            if (e.key === 'Enter') {
                              applyChange(r, c, e.currentTarget.value);
                              setEditingCell(null);
                              formula.endEdit();
                              setSelected({ row: Math.min(grid.length - 1, r + 1), col: c });
                              gridRef.current?.focus();
                            } else if (e.key === 'Escape') {
                              setEditingCell(null);
                              formula.endEdit();
                              gridRef.current?.focus();
                            } else if (e.key === 'Tab') {
                              e.preventDefault();
                              applyChange(r, c, e.currentTarget.value);
                              setEditingCell(null);
                              formula.endEdit();
                              setSelected({ row: r, col: Math.min((grid[0]?.length ?? 1) - 1, c + 1) });
                              gridRef.current?.focus();
                            }
                          }}
                          className="h-full w-full border border-emerald-500 bg-ws-surface px-1.5 font-mono text-xs text-ws outline-none"
                        />
                      ) : (
                        <span
                          className={cn(
                            'block truncate',
                            isFormula && 'text-emerald-500',
                            isError(cell as never) && 'text-ws-danger'
                          )}
                        >
                          {display}
                        </span>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
            {bottomSpacerHeight > 0 && (
              <tr style={{ height: `${bottomSpacerHeight}px` }}>
                <td colSpan={(grid[0]?.length ?? 0) + 1} style={{ padding: 0, border: 'none' }} />
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Листы */}
      <div className="flex shrink-0 items-center gap-1 border-t border-ws bg-ws-surface px-2 py-1 text-xs select-none">
        <div className="flex max-w-[75%] items-center gap-1 overflow-x-auto">
          {sheetNames.map((name) => (
            <button
              key={name}
              onClick={() => {
                selectSingle({ row: 0, col: 0 });
                setActiveSheet(name);
              }}
              onDoubleClick={() => {
                const next = window.prompt('Новое имя листа:', name);
                if (next) renameSheet(name, next);
              }}
              title="Двойной клик — переименовать лист"
              className={cn(
                'shrink-0 rounded-t px-3 py-1 font-medium transition-colors',
                activeSheet === name
                  ? 'border-t-2 border-emerald-500 bg-ws-editor text-emerald-500'
                  : 'text-ws-muted hover:bg-ws-hover hover:text-ws'
              )}
            >
              {name}
            </button>
          ))}
        </div>

        <button
          onClick={addSheet}
          title="Добавить лист"
          className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"
        >
          <Plus size={13} />
        </button>

        <div className="ml-auto flex items-center gap-3 text-[10px] text-ws-faint">
          <span>
            {stats.rows} × {stats.cols}
          </span>
          <span>Заполнено: {stats.filled}</span>
          {formulaCount > 0 && <span className="text-emerald-500">Формул: {formulaCount}</span>}
          {tab.isDirty && <span className="text-ws-warning">● Изменено</span>}
        </div>
      </div>

      {/* Диалог AI-формулы */}
      {showFormulaAI && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/55 p-4 backdrop-blur-[2px]"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) setShowFormulaAI(false);
          }}
        >
          <div className="animate-scale-in w-full max-w-md overflow-hidden rounded-xl border border-ws bg-ws-surface shadow-ws">
            <div className="flex items-center justify-between border-b border-ws px-4 py-3">
              <div className="flex items-center gap-2 text-sm font-semibold text-emerald-500">
                <Sparkles size={15} />
                Генератор формул
              </div>
              <button onClick={() => setShowFormulaAI(false)} className="text-ws-faint hover:text-ws">
                ✕
              </button>
            </div>

            <div className="space-y-3 p-4">
              <p className="text-[11px] leading-relaxed text-ws-muted">
                Опишите задачу словами — AI подберёт формулу для ячейки{' '}
                <strong className="font-mono text-emerald-500">{currentRef}</strong> и предложит
                вставить её с подтверждением.
              </p>

              <textarea
                autoFocus
                value={aiPrompt}
                onChange={(e) => setAiPrompt(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) requestFormula();
                }}
                rows={3}
                placeholder="Например: посчитай сумму по столбцу B только там, где в столбце C указано «Оплачено»"
                className="w-full resize-none rounded-lg border border-ws bg-ws-elevated px-3 py-2 text-xs text-ws placeholder:text-ws-faint focus:border-emerald-500 focus:outline-none"
              />
            </div>

            <div className="flex justify-end gap-2 border-t border-ws px-4 py-3">
              <Button variant="ghost" onClick={() => setShowFormulaAI(false)}>
                Отмена
              </Button>
              <Button variant="primary" onClick={requestFormula} disabled={!aiPrompt.trim()}>
                <Sparkles size={12} />
                Сгенерировать
              </Button>
            </div>
          </div>
        </div>
      )}

      {/* Подтверждение действия AI */}
      {pendingAction && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/55 p-4 backdrop-blur-[2px]">
          <div className="animate-scale-in w-full max-w-lg overflow-hidden rounded-xl border border-ws bg-ws-surface shadow-ws">
            <div className="flex items-center justify-between border-b border-ws px-4 py-3">
              <div className="flex items-center gap-2 text-sm font-semibold text-emerald-500">
                <Wand2 size={15} />
                Применить действие AI?
              </div>
              <button
                onClick={() => setPendingAction(null)}
                className="text-ws-faint hover:text-ws"
              >
                ✕
              </button>
            </div>

            <div className="space-y-3 p-4">
              <p className="text-xs text-ws-muted">
                AI предлагает изменить таблицу. Проверьте формулу перед применением.
              </p>

              <div className="flex items-center gap-2 text-xs">
                <span className="text-ws-muted">Ячейка:</span>
                <span className="rounded border border-ws bg-ws-elevated px-2 py-0.5 font-mono font-bold text-emerald-500">
                  {pendingAction.target}
                </span>
                <span className="text-ws-faint">лист «{activeSheet}»</span>
              </div>

              <div className="rounded-lg border border-ws bg-ws-elevated p-3">
                <div className="mb-1 text-[10px] font-semibold tracking-wide text-ws-muted uppercase">
                  Формула
                </div>
                <code className="block font-mono text-xs break-all text-ws-success select-text">
                  {pendingAction.value}
                </code>
              </div>

              {/* Предпросмотр результата на текущих данных */}
              <div className="rounded-lg border border-ws bg-ws-editor p-3">
                <div className="mb-1 text-[10px] font-semibold tracking-wide text-ws-muted uppercase">
                  Предпросмотр на текущих данных
                </div>
                <PreviewValue sheets={sheets} sheet={activeSheet} action={pendingAction} />
              </div>
            </div>

            <div className="flex justify-end gap-2 border-t border-ws px-4 py-3">
              <Button variant="ghost" onClick={() => setPendingAction(null)}>
                Отмена
              </Button>
              <Button variant="primary" onClick={() => applyAIAction(pendingAction)}>
                <Check size={12} />
                Применить
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

/* ============================================================
   Вспомогательные компоненты и функции
   ============================================================ */

/** Показывает результат формулы до её применения. */
const PreviewValue: React.FC<{
  sheets: Workbook;
  sheet: string;
  action: AIAction;
}> = ({ sheets, sheet, action }) => {
  const result = useMemo(() => {
    const bang = action.target.lastIndexOf('!');
    const targetSheet = bang !== -1 ? action.target.slice(0, bang).replace(/^'|'$/g, '') : sheet;
    const ref = bang !== -1 ? action.target.slice(bang + 1) : action.target;

    const m = ref.replace(/\$/g, '').match(/^([A-Za-z]{1,3})(\d{1,7})$/);
    if (!m) return 'Не удалось разобрать адрес';

    let col = 0;
    const letters = m[1].toUpperCase();
    for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);

    // Временная книга с подставленной формулой
    const draft: Workbook = {};
    for (const [name, grid] of Object.entries(sheets)) {
      draft[name] = grid.map((r) => [...r]);
    }
    const sheetGrid = draft[targetSheet];
    if (!sheetGrid) return `Лист «${targetSheet}» не найден`;

    const row = parseInt(m[2], 10) - 1;
    while (sheetGrid.length <= row) sheetGrid.push(new Array(sheetGrid[0]?.length ?? 16).fill(''));
    const targetRow = sheetGrid[row];
    while (targetRow.length <= col - 1) targetRow.push('');
    targetRow[col - 1] = normalizeFormula(action.value);

    const value = computeCellDisplay(
      draft,
      targetSheet,
      row,
      col - 1,
      createContext(draft, targetSheet)
    );
    return value === '' ? '(пусто)' : value;
  }, [sheets, sheet, action]);

  return (
    <code className="block font-mono text-sm break-all text-ws select-text">
      {result}
      {result.startsWith('#') && (
        <span className="ml-2 text-[10px] text-ws-danger">
          — формула вернула ошибку, проверьте ссылки
        </span>
      )}
    </code>
  );
};

export { shiftFormula };

