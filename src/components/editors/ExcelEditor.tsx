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
  Bold,
  Italic,
  Underline,
  Strikethrough,
  AlignLeft,
  AlignCenter,
  AlignRight,
  WrapText,
  PaintBucket,
  Merge,
  SplitSquareHorizontal,
  ArrowUpAZ,
  ArrowDownAZ,
  Pin,
  PinOff,
  Search,
  Replace,
  Palette,
  BarChart3,
  MessageSquare,
  ListChecks,
  SlidersHorizontal,
  Table2,
  ExternalLink,
  RefreshCw,
  Play,
} from 'lucide-react';
import { cn } from '../../lib/utils';
import { registerSheetController } from '../../services/sheetBridge';
import type { SheetEditResult, SheetReadResult } from '../../services/sheetBridge';
import { ClipboardPaste, Undo2, Redo2 } from 'lucide-react';
import { parseClipboardTable } from '../../lib/clipboardTable';
import * as SheetFmt from '../../lib/sheetFormatting';
import { Button, Input, Modal, Spinner } from '../ui';
import { SheetChart, ChartSeries } from './SheetChart';
import { injectChartsIntoXlsx } from '../../lib/xlsxCharts';
import { injectStylesIntoXlsx } from '../../lib/xlsxStyles';
import { applyAutoFitColumns } from '../../lib/xlsxColumns';
import { detectMarkdownTables, normalizeTableValue } from '../../services/aiActions';
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
  onSave: () => void | Promise<unknown>;
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

/** Извлекает табличные данные из ответа AI: Markdown, JSON values/rows или TSV. */
function extractTableRows(text: string): string[][] {
  const table = detectMarkdownTables(text)[0];
  if (table?.columns?.length) {
    return [table.columns, ...table.rows];
  }

  const start = text.indexOf('{');
  if (start !== -1) {
    let depth = 0;
    let inString = false;
    let escaped = false;
    for (let index = start; index < text.length; index++) {
      const ch = text[index];
      if (inString) {
        if (escaped) escaped = false;
        else if (ch === '\\') escaped = true;
        else if (ch === '"') inString = false;
        continue;
      }
      if (ch === '"') inString = true;
      else if (ch === '{') depth++;
      else if (ch === '}') {
        depth--;
        if (depth === 0) {
          try {
            const parsed = JSON.parse(text.slice(start, index + 1)) as Record<string, unknown>;
            const raw = parsed.values ?? parsed.rows ?? parsed.data;
            if (Array.isArray(raw)) {
              return raw.map((row) =>
                Array.isArray(row) ? row.map((cell) => String(cell ?? '')) : [String(row ?? '')]
              );
            }
          } catch {
            // не JSON — пробуем восстановить строки ниже
          }
          break;
        }
      }
    }
    // Оборванный JSON: собираем все полностью сформированные массивы-строки
    const recovered: string[][] = [];
    for (const match of text.matchAll(/\[[^\[\]]*\]/g)) {
      try {
        const row = JSON.parse(match[0]);
        if (Array.isArray(row)) recovered.push(row.map((cell) => String(cell ?? '')));
      } catch {
        // пропускаем незавершённую строку
      }
    }
    if (recovered.length > 0) return recovered;
  }

  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (lines.some((line) => line.includes('\t'))) {
    return lines.map((line) => line.split('\t'));
  }
  if (lines.some((line) => line.includes('|'))) {
    return lines.map((line) =>
      line
        .split('|')
        .map((cell) => cell.trim())
        .filter((cell, index, list) => !(index === 0 && cell === '') && !(index === list.length - 1 && cell === ''))
    );
  }
  // Обычный текст (в том числе JSON без строк) не вставляем в таблицу:
  // иначе весь ответ оказывается в одном столбце.
  return [];
}

/** Снимок книги для отмены/возврата: значения + форматирование. */
interface WorkbookSnapshot {
  sheets: Workbook;
  meta: SheetFmt.WorkbookMeta;
}

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
  const [meta, setMeta] = useState<SheetFmt.WorkbookMeta>({});
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
  const [history, setHistory] = useState<WorkbookSnapshot[]>([]);
  const [aiDiff, setAiDiff] = useState<{
    cells: Array<{ address: string; before: string; after: string }>;
  } | null>(null);
  const [future, setFuture] = useState<WorkbookSnapshot[]>([]);
  /** Перечитывание файла после правок в настоящем Excel. */
  const [reloadToken, setReloadToken] = useState(0);
  /** VBA-проект из .xlsm: сохраняется при записи книги. */
  const vbaRawRef = useRef<ArrayBuffer | Uint8Array | null>(null);
  const [excelBusy, setExcelBusy] = useState(false);
  /** Защита от перезаписи более новой версией при асинхронном экспорте диаграмм. */
  const syncSeqRef = useRef(0);

  /** Актуальные книга и метаданные для асинхронных колбэков. */
  const sheetsRef = useRef(sheets);
  sheetsRef.current = sheets;
  const metaRef = useRef(meta);
  metaRef.current = meta;
  const sheetNamesRef = useRef(sheetNames);
  sheetNamesRef.current = sheetNames;

  /* ---- Поиск и замена ---- */
  const [findOpen, setFindOpen] = useState(false);
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [findMessage, setFindMessage] = useState('');

  /** Автофильтр активного листа: скрывает строки с невыбранными значениями. */
  const [filters, setFilters] = useState<
    Record<
      string,
      { col: number; headerRow: number; allValues: string[]; values: string[] } | null
    >
  >({});
  const [filterPanelOpen, setFilterPanelOpen] = useState(false);

  /* ---- Диаграммы, проверка данных, условное форматирование, сводные ---- */
  const [chartDialogOpen, setChartDialogOpen] = useState(false);
  const [chartType, setChartType] = useState<SheetFmt.ChartType>('bar');
  const [chartTitle, setChartTitle] = useState('Диаграмма');
  const [validationPanelOpen, setValidationPanelOpen] = useState(false);
  const [validationText, setValidationText] = useState('');
  const [conditionalPanelOpen, setConditionalPanelOpen] = useState(false);
  const [ruleOperator, setRuleOperator] = useState<SheetFmt.ConditionalOperator>('greater');
  const [ruleValue, setRuleValue] = useState('');
  const [ruleValue2, setRuleValue2] = useState('');
  const [ruleFill, setRuleFill] = useState('#fde68a');
  const [ruleColor, setRuleColor] = useState('#111827');
  const [ruleBold, setRuleBold] = useState(false);
  const [pivotOpen, setPivotOpen] = useState(false);
  const [pivotRowCol, setPivotRowCol] = useState(0);
  const [pivotValueCol, setPivotValueCol] = useState(1);
  const [pivotAgg, setPivotAgg] = useState<'sum' | 'count' | 'average' | 'min' | 'max'>('sum');

  /** Перетаскивание/изменение размера диаграммы. */
  const chartDragRef = useRef<{
    id: string;
    mode: 'move' | 'resize';
    startX: number;
    startY: number;
    x: number;
    y: number;
    width: number;
    height: number;
  } | null>(null);

  /** Уведомление в панели инструментов (сортировка, объединение и т.п.). */
  const [notice, setNotice] = useState('');
  const noticeTimerRef = useRef<number | null>(null);
  const showNotice = useCallback((message: string) => {
    setNotice(message);
    if (noticeTimerRef.current) window.clearTimeout(noticeTimerRef.current);
    noticeTimerRef.current = window.setTimeout(() => setNotice(''), 2600);
  }, []);

  /** Текущий жест изменения ширины столбца / высоты строки. */
  const resizeRef = useRef<{
    type: 'col' | 'row';
    index: number;
    start: number;
    size: number;
  } | null>(null);

  /* ---- AI-формула ---- */
  const [showFormulaAI, setShowFormulaAI] = useState(false);
  const [repairNote, setRepairNote] = useState('');
  const [aiPrompt, setAiPrompt] = useState('');
  const [pendingAction, setPendingAction] = useState<AIAction | null>(null);
  /** Ячейка, для которой запрошена формула — ждёт ответа ассистента. */
  const pendingRef = useRef<{ ref: string } | null>(null);

  const formula = useFormulaMode();
  const gridRef = useRef<HTMLDivElement>(null);
  const [zoom, setZoom] = useState(1);
  const [promptState, setPromptState] = useState<{ title: string; value: string } | null>(null);
  const promptResolve = useRef<((value: string | null) => void) | null>(null);

  const askText = useCallback(
    (title: string, initial = '') =>
      new Promise<string | null>((resolve) => {
        promptResolve.current = resolve;
        setPromptState({ title, value: initial });
      }),
    []
  );

  const finishPrompt = (value: string | null) => {
    const resolve = promptResolve.current;
    promptResolve.current = null;
    setPromptState(null);
    resolve?.(value);
  };

  useEffect(() => {
    const element = gridRef.current;
    if (!element) return;
    const onWheel = (event: WheelEvent) => {
      if (!(event.ctrlKey || event.metaKey)) return;
      event.preventDefault();
      setZoom((previous) =>
        Math.min(2, Math.max(0.5, previous - event.deltaY * 0.0015))
      );
    };
    element.addEventListener('wheel', onWheel, { passive: false });
    return () => element.removeEventListener('wheel', onWheel);
  }, [loading]);
  const gridFocusedRef = useRef(false);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const cellEditorRef = useRef<HTMLInputElement>(null);

  const isCsv = tab.extension === 'csv' || tab.extension === 'tsv';

  /* ---------------- Разбор книги ---------------- */

  useEffect(() => {
    try {
      setLoading(true);
      setError('');

      const isEmpty = !tab.content || tab.content.trim() === '';
      const workbook = isEmpty
        ? (() => {
            const wb = XLSX.utils.book_new();
            const emptyWs = XLSX.utils.aoa_to_sheet([[]]);
            XLSX.utils.book_append_sheet(wb, emptyWs, 'Лист 1');
            return wb;
          })()
        : isCsv
          ? XLSX.read(tab.content, { type: 'string' })
          : XLSX.read(tab.content, {
              type: 'base64',
              bookVBA: tab.extension === 'xlsm',
              cellStyles: true,
            });

      vbaRawRef.current = (workbook as { vbaraw?: ArrayBuffer | Uint8Array }).vbaraw ?? null;

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

      // Форматирование: ширины, высоты и объединения читаются из самого xlsx,
      // стили и закрепление — из пользовательского свойства документа.
      let savedMeta: SheetFmt.WorkbookMeta = {};
      try {
        const rawStyles = (workbook.Custprops as Record<string, unknown> | undefined)
          ?.WorkSpaceStyles;
        savedMeta = SheetFmt.decodeWorkbookMeta(rawStyles);
      } catch {
        savedMeta = {};
      }

      workbook.SheetNames.forEach((name) => {
        const source = workbook.Sheets[name];
        const sheetMeta = savedMeta[name] ?? SheetFmt.emptySheetMeta();
        source?.['!cols']?.forEach((col, index) => {
          const width = col?.wpx ?? (col?.wch ? col.wch * 7.5 : undefined);
          if (typeof width === 'number' && width > 0) sheetMeta.colWidths[index] = width;
        });
        source?.['!rows']?.forEach((row, index) => {
          const height = row?.hpx ?? (row?.hpt ? row.hpt * 1.333 : undefined);
          if (typeof height === 'number' && height > 0) sheetMeta.rowHeights[index] = height;
        });
        const nativeMerges = source?.['!merges'];
        if (nativeMerges?.length) {
          for (const merge of nativeMerges) {
            const address = SheetFmt.rangeToAddress({
              startRow: merge.s.r,
              endRow: merge.e.r,
              startCol: merge.s.c,
              endCol: merge.e.c,
            });
            if (!sheetMeta.merges.includes(address)) sheetMeta.merges.push(address);
          }
        }
        savedMeta[name] = sheetMeta;
      });

      metaRef.current = savedMeta;
      setMeta(savedMeta);

      setSheets(parsed);
      setHistory([]);
      setFuture([]);
      setAiDiff(null);
    } catch (err) {
      setError(`Не удалось открыть таблицу: ${(err as Error).message}`);
    } finally {
      setLoading(false);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id, tab.reloadKey, reloadToken]);

  const grid = sheets[activeSheet] ?? [];
  const activeMeta = useMemo(
    () => meta[activeSheet] ?? SheetFmt.emptySheetMeta(),
    [meta, activeSheet]
  );
  const mergeInfo = useMemo(
    () => SheetFmt.buildMergeInfo(activeMeta.merges),
    [activeMeta.merges]
  );
  const activeFilter = filters[activeSheet] ?? null;
  const columnCount = grid[0]?.length ?? 0;
  const colOffsets = useMemo(() => {
    const offsets = new Array<number>(columnCount + 1);
    offsets[0] = 0;
    for (let col = 0; col < columnCount; col++) {
      offsets[col + 1] =
        offsets[col] + (activeMeta.colWidths[col] ?? SheetFmt.DEFAULT_COL_WIDTH);
    }
    return offsets;
  }, [activeMeta.colWidths, columnCount]);

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

  const OVERSCAN = 30;
  const totalRows = grid.length;
  const rowOffsets = useMemo(
    () => SheetFmt.buildRowOffsets(activeMeta.rowHeights, totalRows),
    [activeMeta.rowHeights, totalRows]
  );
  // Объединённые ячейки нельзя рисовать частями: при слияниях показываем лист целиком.
  const isVirtual =
    zoom === 1 &&
    totalRows > 80 &&
    mergeInfo.anchors.size === 0 &&
    activeMeta.freezeRows === 0 &&
    !activeFilter;

  const firstVisible = SheetFmt.findRowAtOffset(rowOffsets, scrollTop);
  const startRow = isVirtual ? Math.max(0, firstVisible - OVERSCAN) : 0;
  const endRow = isVirtual
    ? Math.min(
        totalRows,
        SheetFmt.findRowAtOffset(rowOffsets, scrollTop + containerHeight) + OVERSCAN + 1
      )
    : totalRows;

  const topSpacerHeight = isVirtual ? rowOffsets[startRow] : 0;
  const bottomSpacerHeight = isVirtual ? Math.max(0, rowOffsets[totalRows] - rowOffsets[endRow]) : 0;

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
    async (nextSheets: Workbook, names: string[]): Promise<string | null> => {
      const seq = ++syncSeqRef.current;
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
              const cellStyle = metaRef.current[name]?.cells[address];
              const formatCode = SheetFmt.numberFormatCode(cellStyle?.numberFormat);
              if (formatCode && isNumeric) next.z = formatCode;

              sheet[address] = next as unknown as XLSX.CellObject;
            }
          }

          // Ширины: даже если пользователь не тянул колонки, Excel должен
          // открыть файл с текстом внутри ячеек, а не с обрезанными 8,43.
          const sheetMeta = metaRef.current[name];
          const displayGrid = trimmed.map((row, rowIndex) =>
            row.map((cell, colIndex) => {
              const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
              const style = metaRef.current[name]?.cells[address];
              if (typeof cell === 'string' && cell.startsWith('=')) {
                return SheetFmt.formatCellValue(
                  computeCellDisplay(nextSheets, name, rowIndex, colIndex, formulaContext),
                  style?.numberFormat
                );
              }
              return SheetFmt.formatCellValue(String(cell ?? ''), style?.numberFormat);
            })
          );
          applyAutoFitColumns(sheet, displayGrid, {
            explicitPx: sheetMeta?.colWidths,
            headerRows: 1,
            isWrapped: (rowIndex, colIndex) =>
              Boolean(
                metaRef.current[name]?.cells[
                  XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
                ]?.wrap
              ),
          });
          if (sheetMeta) {
            const rowIndexes = Object.keys(sheetMeta.rowHeights).map(Number);
            if (rowIndexes.length > 0) {
              const maxRow = Math.max(...rowIndexes);
              sheet['!rows'] = Array.from({ length: maxRow + 1 }, (_, index) => ({
                hpx: sheetMeta.rowHeights[index] ?? SheetFmt.DEFAULT_ROW_HEIGHT,
              }));
            }
            if (sheetMeta.merges.length > 0) {
              sheet['!merges'] = [...new Set(sheetMeta.merges)]
                .map((item) => SheetFmt.parseRange(item))
                .filter((range): range is SheetFmt.GridRange => Boolean(range))
                .map((range) => ({
                  s: { r: range.startRow, c: range.startCol },
                  e: { r: range.endRow, c: range.endCol },
                }));
            }
          }

          XLSX.utils.book_append_sheet(wb, sheet, name);
        });

        wb.Custprops = {
          ...(wb.Custprops ?? {}),
          WorkSpaceStyles: SheetFmt.encodeWorkbookMeta(metaRef.current),
        };

        let content: string;
        if (isCsv) {
          content = XLSX.utils.sheet_to_csv(wb.Sheets[activeSheet] ?? wb.Sheets[names[0]]);
        } else if (tab.extension === 'xlsm') {
          // Макросы .xlsm нельзя терять: переносим VBA-проект в новую книгу
          (wb as { vbaraw?: unknown }).vbaraw = vbaRawRef.current ?? undefined;
          content = XLSX.write(wb, { bookType: 'xlsm', type: 'base64', bookVBA: true });
        } else {
          content = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
        }

        // Нативные стили ячеек: SheetJS CE не пишет заливку/шрифт/границы.
        if (!isCsv && tab.extension !== 'xlsm') {
          try {
            content = await injectStylesIntoXlsx(content, metaRef.current);
          } catch (err) {
            console.error('Не удалось записать стили в xlsx', err);
          }
        }

        // Нативные диаграммы Excel (OOXML) — без установленного Excel
        if (!isCsv && tab.extension !== 'xlsm') {
          const chartEntries = names
            .map((name) => ({
              sheet: name,
              charts: metaRef.current[name]?.charts ?? [],
            }))
            .filter((entry) => entry.charts.length > 0);
          if (chartEntries.length > 0) {
            try {
              content = await injectChartsIntoXlsx(content, chartEntries);
            } catch (err) {
              console.error('Не удалось добавить диаграммы в xlsx', err);
            }
          }
        }

        if (seq === syncSeqRef.current) onChange(content);
        return content;
      } catch (err) {
        console.error('Ошибка синхронизации книги', err);
        return null;
      }
    },
    [activeSheet, isCsv, onChange, tab.extension]
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

  /** Границы цветовых шкал для условного форматирования. */
  const conditionalScaleBounds = useMemo(() => {
    const bounds = new Map<string, { min: number; max: number }>();
    const toNumber = (value: string): number | null => {
      const normalized = value.replace(/\s/g, '').replace(',', '.');
      if (normalized === '' || !/^-?\d+(\.\d+)?$/.test(normalized)) return null;
      const num = Number(normalized);
      return Number.isFinite(num) ? num : null;
    };
    for (const rule of activeMeta.conditional) {
      if (rule.operator !== 'scale') continue;
      const range = SheetFmt.parseRange(rule.range);
      if (!range) continue;
      let min = Number.POSITIVE_INFINITY;
      let max = Number.NEGATIVE_INFINITY;
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startCol; col <= range.endCol; col++) {
          const num = toNumber(displayCell(row, col));
          if (num === null) continue;
          if (num < min) min = num;
          if (num > max) max = num;
        }
      }
      if (Number.isFinite(min) && Number.isFinite(max)) bounds.set(rule.id, { min, max });
    }
    return bounds;
  }, [activeMeta.conditional, displayCell]);

  /* ---------------- Изменение ячеек ---------------- */

  const pushHistory = useCallback(() => {
    setHistory((h) => [
      ...h.slice(-49),
      { sheets: sheetsRef.current, meta: metaRef.current },
    ]);
    setFuture([]);
  }, []);

  /** Применяет новые метаданные форматирования и синхронизирует файл. */
  const commitMeta = useCallback(
    (nextMeta: SheetFmt.WorkbookMeta) => {
      pushHistory();
      metaRef.current = nextMeta;
      setMeta(nextMeta);
      syncWorkbook(sheetsRef.current, sheetNamesRef.current);
    },
    [pushHistory, syncWorkbook]
  );

  const applyChange = useCallback(
    (row: number, col: number, value: string | number) => {
      const current = sheetsRef.current[activeSheet];
      if (!current) return;
      if (String(current[row]?.[col] ?? '') === String(value)) return;
      pushHistory();

      const next: Workbook = { ...sheetsRef.current, [activeSheet]: current.map((r) => [...r]) };
      next[activeSheet][row][col] = value;
      syncWorkbook(next, sheetNamesRef.current);
      setSheets(next);
    },
    [activeSheet, pushHistory, syncWorkbook]
  );

  /**
   * Применяет пачку изменений одной операцией (для действий AI и структурных
   * операций). Вместе со значениями можно изменить метаданные форматирования.
   */
  const applyBatch = useCallback(
    (mutate: (draft: Workbook) => void, mutateMeta?: (draft: SheetFmt.WorkbookMeta) => void) => {
      const base = sheetsRef.current;
      const next: Workbook = {};
      for (const [name, sheetGrid] of Object.entries(base)) {
        next[name] = sheetGrid.map((r) => [...r]);
      }
      mutate(next);

      let nextMeta = metaRef.current;
      if (mutateMeta) {
        nextMeta = SheetFmt.cloneWorkbookMeta(metaRef.current);
        mutateMeta(nextMeta);
      }

      setHistory((h) => [...h.slice(-49), { sheets: base, meta: metaRef.current }]);
      setFuture([]);

      metaRef.current = nextMeta;
      setMeta(nextMeta);
      sheetsRef.current = next;
      syncWorkbook(next, sheetNamesRef.current);
      setSheets(next);
    },
    [syncWorkbook]
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

  /* ---------------- Вставка/замена из ответа AI ---------------- */
  useEffect(() => {
    const onInsert = (event: Event) => {
      const detail = (event as CustomEvent<{ tabId: string; text: string }>).detail;
      if (!detail || (detail.tabId && detail.tabId !== tab.id)) return;
      const rows = extractTableRows(detail.text);
      if (rows.length === 0) {
        showNotice('В ответе нет табличных данных');
        return;
      }
      insertTable(rows);
      showNotice(`Вставлено строк: ${rows.length}`);
    };

    const onReplace = (event: Event) => {
      const detail = (event as CustomEvent<{ tabId: string; text: string }>).detail;
      if (!detail || (detail.tabId && detail.tabId !== tab.id)) return;
      const rows = extractTableRows(detail.text);
      if (rows.length === 0) {
        showNotice('В ответе нет табличных данных');
        return;
      }
      const cols = Math.max(MIN_COLS, ...rows.map((row) => row.length));
      const height = Math.max(MIN_ROWS, rows.length);
      applyBatch((draft) => {
        draft[activeSheet] = Array.from({ length: height }, (_, row) =>
          Array.from({ length: cols }, (_, col) => rows[row]?.[col] ?? '')
        );
      });
      selectSingle({ row: 0, col: 0 });
      showNotice(`Таблица заменена: ${rows.length} строк`);
    };

    window.addEventListener('workspace:excel-insert-text', onInsert);
    window.addEventListener('workspace:excel-replace-document', onReplace);
    return () => {
      window.removeEventListener('workspace:excel-insert-text', onInsert);
      window.removeEventListener('workspace:excel-replace-document', onReplace);
    };
  }, [tab.id, insertTable, applyBatch, activeSheet, showNotice]);

  const restoreSnapshot = useCallback(
    (snapshot: WorkbookSnapshot) => {
      metaRef.current = snapshot.meta;
      setMeta(snapshot.meta);
      setSheets(snapshot.sheets);
      syncWorkbook(snapshot.sheets, sheetNamesRef.current);
    },
    [syncWorkbook]
  );

  const undo = useCallback(() => {
    if (history.length === 0) return;
    const previous = history[history.length - 1];
    setFuture((f) => [{ sheets: sheetsRef.current, meta: metaRef.current }, ...f].slice(0, 50));
    setHistory((h) => h.slice(0, -1));
    restoreSnapshot(previous);
  }, [history, restoreSnapshot]);

  const redo = useCallback(() => {
    if (future.length === 0) return;
    const [next, ...rest] = future;
    setHistory((h) => [...h, { sheets: sheetsRef.current, meta: metaRef.current }].slice(-50));
    setFuture(rest);
    restoreSnapshot(next);
  }, [future, restoreSnapshot]);

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
      } else if ((key === 'f' || key === 'а' || code === 'KeyF') && !e.shiftKey) {
        e.preventDefault();
        e.stopPropagation();
        setFindOpen(true);
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
    applyBatch(
      (draft) => {
        const current = draft[activeSheet];
        if (!current || current.length - indexes.size < 1) return;
        draft[activeSheet] = current.filter((_, index) => !indexes.has(index));
      },
      (draftMeta) => {
        [...indexes]
          .sort((a, b) => b - a)
          .forEach((index) => {
            const next = SheetFmt.removeRowFromMeta(draftMeta, activeSheet, index);
            draftMeta[activeSheet] = next[activeSheet];
          });
      }
    );
    selectSingle({ row: Math.min(selectionRange.start.row, grid.length - 2), col: selectionRange.start.col });
  };

  const deleteSelectionColumns = () => {
    const indexes = new Set<number>();
    for (let col = selectionRange.start.col; col <= selectionRange.end.col; col++) indexes.add(col);
    extraCells.forEach((cell) => indexes.add(cell.col));
    applyBatch(
      (draft) => {
        const current = draft[activeSheet];
        if (!current || (current[0]?.length ?? 0) - indexes.size < 1) return;
        draft[activeSheet] = current.map((line) => line.filter((_, index) => !indexes.has(index)));
      },
      (draftMeta) => {
        [...indexes]
          .sort((a, b) => b - a)
          .forEach((index) => {
            const next = SheetFmt.removeColumnFromMeta(draftMeta, activeSheet, index);
            draftMeta[activeSheet] = next[activeSheet];
          });
      }
    );
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

  /* ---------------- Мост в настоящий Excel ---------------- */

  const reloadFromDisk = useCallback(async (): Promise<boolean> => {
    if (!window.electronAPI?.readFile) return false;
    const res = await window.electronAPI.readFile(tab.path, isCsv ? 'utf-8' : 'base64');
    if (!res.success || typeof res.content !== 'string') {
      showNotice('Не удалось перечитать файл');
      return false;
    }
    onChange(res.content);
    setReloadToken((value) => value + 1);
    showNotice('Файл перечитан из Excel');
    return true;
  }, [tab.path, isCsv, onChange, showNotice]);

  const openInExcel = async () => {
    if (!window.electronAPI?.openInExcel) return;
    setExcelBusy(true);
    try {
      await onSave();
      const res = await window.electronAPI.openInExcel(tab.path);
      showNotice(res.success ? 'Открыто в настоящем Excel' : `Excel: ${res.error || 'ошибка'}`);
    } finally {
      setExcelBusy(false);
    }
  };

  const runExcelMacro = async () => {
    if (!window.electronAPI?.runExcelMacro) return;
    const macroName = await askText(
      'Имя VBA-макроса (например, Module1.Report).\nМакрос выполнит настоящий Excel.',
      ''
    );
    if (!macroName?.trim()) return;
    setExcelBusy(true);
    try {
      await onSave();
      const res = await window.electronAPI.runExcelMacro(tab.path, macroName.trim());
      if (res.success) {
        await reloadFromDisk();
        showNotice(`Макрос выполнен: ${macroName.trim()}`);
      } else {
        showNotice(`Макрос не выполнен: ${res.error || 'ошибка Excel'}`);
      }
    } finally {
      setExcelBusy(false);
    }
  };

  /** Экспорт диаграмм WorkSpace в книгу как нативных объектов Excel. */
  const exportChartsToExcel = async () => {
    if (!window.electronAPI?.excelExportCharts || activeMeta.charts.length === 0) return;
    setExcelBusy(true);
    try {
      await onSave();
      const res = await window.electronAPI.excelExportCharts(
        tab.path,
        activeMeta.charts.map((chart) => ({
          sheet: activeSheet,
          type: chart.type,
          range: chart.range,
          title: chart.title,
          x: chart.x,
          y: chart.y,
          width: chart.width,
          height: chart.height,
        }))
      );
      if (res.success) {
        await reloadFromDisk();
        showNotice(`Диаграммы перенесены в Excel: ${res.count ?? activeMeta.charts.length}`);
      } else {
        showNotice(`Экспорт диаграмм: ${res.error || 'ошибка Excel'}`);
      }
    } finally {
      setExcelBusy(false);
    }
  };

  /** Настоящая обновляемая сводная таблица Excel по текущему диапазону. */
  const createExcelPivot = async () => {
    if (!window.electronAPI?.excelCreatePivot) return;
    const range = SheetFmt.detectUsedRange(grid, selected.row, selected.col);
    const rowField = displayCell(range.startRow, pivotRowCol).trim();
    const valueField = displayCell(range.startRow, pivotValueCol).trim();
    if (!rowField || !valueField) {
      showNotice('Для сводной Excel нужны заголовки столбцов');
      return;
    }
    let index = 1;
    while (sheetNamesRef.current.includes(`Сводная Excel ${index}`)) index++;
    setExcelBusy(true);
    try {
      await onSave();
      const res = await window.electronAPI.excelCreatePivot(tab.path, {
        sourceSheet: activeSheet,
        range: SheetFmt.rangeToAddress(range),
        rowField,
        valueField,
        aggregation: pivotAgg,
        targetSheet: `Сводная Excel ${index}`,
      });
      if (res.success) {
        await reloadFromDisk();
        setPivotOpen(false);
        showNotice(`Сводная Excel создана: Сводная Excel ${index}`);
      } else {
        showNotice(`Сводная Excel: ${res.error || 'ошибка Excel'}`);
      }
    } finally {
      setExcelBusy(false);
    }
  };

  /* ---------------- Форматирование ячеек ---------------- */

  const selectedStyle = SheetFmt.styleAt(activeMeta, selected.row, selected.col) ?? {};

  const selectionAsRange = (): SheetFmt.GridRange => ({
    startRow: selectionRange.start.row,
    endRow: selectionRange.end.row,
    startCol: selectionRange.start.col,
    endCol: selectionRange.end.col,
  });

  const applyStyle = (patch: Partial<SheetFmt.CellStyle>) => {
    commitMeta(
      SheetFmt.applyStyleToRange(metaRef.current, activeSheet, selectionAsRange(), patch)
    );
  };

  const toggleStyle = (key: 'bold' | 'italic' | 'underline' | 'strike' | 'wrap') => {
    applyStyle({ [key]: !selectedStyle[key] } as Partial<SheetFmt.CellStyle>);
  };

  const applyBorders = (mode: 'all' | 'outer' | 'none') => {
    commitMeta(
      SheetFmt.applyBordersToRange(metaRef.current, activeSheet, selectionAsRange(), mode)
    );
  };

  const mergeSelection = () => {
    const range = selectionAsRange();
    if (range.startRow === range.endRow && range.startCol === range.endCol) {
      showNotice('Выделите несколько ячеек для объединения');
      return;
    }
    commitMeta(SheetFmt.addMerge(metaRef.current, activeSheet, range));
    showNotice('Ячейки объединены');
  };

  const unmergeSelection = () => {
    commitMeta(
      SheetFmt.removeMergesInRange(metaRef.current, activeSheet, selectionAsRange())
    );
    showNotice('Объединение снято');
  };

  const freezePanes = () => {
    commitMeta(SheetFmt.setFreeze(metaRef.current, activeSheet, selected.row, selected.col));
    showNotice(
      `Закреплено: строк ${selected.row}, столбцов ${selected.col} (выше и левее выделения)`
    );
  };

  const unfreezePanes = () => {
    commitMeta(SheetFmt.setFreeze(metaRef.current, activeSheet, 0, 0));
    showNotice('Закрепление снято');
  };

  const clearFormatting = () => {
    commitMeta(SheetFmt.clearStyleRange(metaRef.current, activeSheet, selectionAsRange()));
    showNotice('Форматирование снято');
  };

  /** Преобразует текстовые числа всего листа в числа Excel (для SUM и формул). */
  const convertTextToNumbers = () => {
    let converted = 0;
    applyBatch((draft) => {
      const sheet = draft[activeSheet];
      if (!sheet) return;
      for (let row = 0; row < sheet.length; row++) {
        for (let col = 0; col < (sheet[row]?.length ?? 0); col++) {
          const value = sheet[row][col];
          if (typeof value !== 'string') continue;
          const normalized = normalizeTableValue(value);
          if (typeof normalized === 'number') {
            sheet[row][col] = normalized;
            converted++;
          }
        }
      }
    });
    showNotice(
      converted > 0
        ? `Преобразовано чисел: ${converted}`
        : 'Текстовых чисел на листе не найдено'
    );
  };

  /** Цветовая шкала по числовым значениям выделения (статическая условная заливка). */
  const applyColorScale = () => {
    const range = selectionAsRange();
    const entries: { row: number; col: number; num: number }[] = [];
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const raw = displayCell(row, col).replace(/\s/g, '').replace(',', '.');
        if (raw === '' || !/^-?\d+(\.\d+)?$/.test(raw)) continue;
        const num = Number(raw);
        if (Number.isFinite(num)) entries.push({ row, col, num });
      }
    }
    if (entries.length < 2) {
      showNotice('Выделите минимум две ячейки с числами');
      return;
    }
    const min = Math.min(...entries.map((entry) => entry.num));
    const max = Math.max(...entries.map((entry) => entry.num));
    const next = SheetFmt.cloneWorkbookMeta(metaRef.current);
    const target = next[activeSheet] ?? SheetFmt.emptySheetMeta();
    next[activeSheet] = target;
    const stops: [number, [number, number, number]][] = [
      [0, [254, 202, 202]],
      [0.5, [254, 240, 138]],
      [1, [187, 247, 208]],
    ];
    for (const entry of entries) {
      const ratio = max === min ? 0.5 : (entry.num - min) / (max - min);
      const upperIndex = ratio <= 0.5 ? 1 : 2;
      const lower = stops[upperIndex - 1];
      const upper = stops[upperIndex];
      const local = (ratio - lower[0]) / Math.max(0.0001, upper[0] - lower[0]);
      const channels = lower[1].map((value, index) =>
        Math.round(value + (upper[1][index] - value) * local)
      );
      const address = SheetFmt.cellAddress(entry.row, entry.col);
      target.cells[address] = {
        ...target.cells[address],
        fill: `rgb(${channels[0]}, ${channels[1]}, ${channels[2]})`,
      };
    }
    commitMeta(next);
    showNotice(`Цветовая шкала: ${min} … ${max}`);
  };

  /* ---------------- Автофильтр ---------------- */

  const openFilterPanel = () => {
    if (filters[activeSheet]) {
      setFilterPanelOpen(true);
      return;
    }
    const range = SheetFmt.detectUsedRange(grid, selected.row, selected.col);
    if (range.endRow <= range.startRow) {
      showNotice('Нет данных для фильтра');
      return;
    }
    const values = new Set<string>();
    for (let row = range.startRow + 1; row <= range.endRow; row++) {
      values.add(displayCell(row, selected.col));
    }
    setFilters((prev) => ({
      ...prev,
      [activeSheet]: {
        col: selected.col,
        headerRow: range.startRow,
        allValues: [...values],
        values: [...values],
      },
    }));
    setFilterPanelOpen(true);
  };

  const toggleFilterValue = (value: string) => {
    setFilters((prev) => {
      const filter = prev[activeSheet];
      if (!filter) return prev;
      const values = filter.values.includes(value)
        ? filter.values.filter((item) => item !== value)
        : [...filter.values, value];
      return { ...prev, [activeSheet]: { ...filter, values } };
    });
  };

  const clearFilter = () => {
    setFilters((prev) => ({ ...prev, [activeSheet]: null }));
    setFilterPanelOpen(false);
    showNotice('Фильтр снят');
  };

  const rowMatchesFilter = (row: number): boolean => {
    if (!activeFilter) return true;
    if (row <= activeFilter.headerRow) return true;
    return activeFilter.values.includes(displayCell(row, activeFilter.col));
  };

  /* ---------------- Диаграммы ---------------- */

  const numericFromText = (value: string): number | null => {
    const normalized = value.replace(/\s/g, '').replace(',', '.');
    if (normalized === '' || !/^-?\d+(\.\d+)?$/.test(normalized)) return null;
    const num = Number(normalized);
    return Number.isFinite(num) ? num : null;
  };

  const buildChartData = useCallback(
    (rangeText: string): { labels: string[]; series: ChartSeries[] } => {
      const range = SheetFmt.parseRange(rangeText);
      if (!range) return { labels: [], series: [] };
      const firstValue = displayCell(range.startRow, range.startCol + 1);
      const hasHeader =
        firstValue.trim() !== '' && numericFromText(firstValue) === null;
      const dataStart = hasHeader ? range.startRow + 1 : range.startRow;
      const labels: string[] = [];
      for (let row = dataStart; row <= range.endRow; row++) {
        labels.push(displayCell(row, range.startCol) || `Строка ${row + 1}`);
      }
      const series: ChartSeries[] = [];
      for (let col = range.startCol + 1; col <= range.endCol; col++) {
        const values: (number | null)[] = [];
        for (let row = dataStart; row <= range.endRow; row++) {
          values.push(numericFromText(displayCell(row, col)));
        }
        series.push({
          name: displayCell(range.startRow, col) || columnName(col),
          values,
        });
      }
      return { labels, series };
    },
    [displayCell]
  );

  const createChart = () => {
    const range = selectionAsRange();
    if (range.startRow === range.endRow || range.startCol === range.endCol) {
      showNotice('Выделите диапазон с заголовками и данными');
      return;
    }
    const spec: SheetFmt.ChartSpec = {
      id: `chart_${Date.now()}`,
      type: chartType,
      range: SheetFmt.rangeToAddress(range),
      title: chartTitle.trim() || 'Диаграмма',
      x: 48 + activeMeta.charts.length * 24,
      y: 48 + activeMeta.charts.length * 24,
      width: 460,
      height: 280,
    };
    commitMeta(SheetFmt.addChartSpec(metaRef.current, activeSheet, spec));
    setChartDialogOpen(false);
    showNotice(`Диаграмма создана по диапазону ${spec.range}`);
  };

  const startChartMove = (
    event: React.MouseEvent,
    chart: SheetFmt.ChartSpec,
    mode: 'move' | 'resize'
  ) => {
    event.preventDefault();
    event.stopPropagation();
    pushHistory();
    chartDragRef.current = {
      id: chart.id,
      mode,
      startX: event.clientX,
      startY: event.clientY,
      x: chart.x,
      y: chart.y,
      width: chart.width,
      height: chart.height,
    };
  };

  useEffect(() => {
    const onMove = (event: MouseEvent) => {
      const drag = chartDragRef.current;
      if (!drag) return;
      const dx = event.clientX - drag.startX;
      const dy = event.clientY - drag.startY;
      const patch =
        drag.mode === 'move'
          ? { x: Math.max(0, drag.x + dx), y: Math.max(0, drag.y + dy) }
          : {
              width: Math.max(260, drag.width + dx),
              height: Math.max(180, drag.height + dy),
            };
      const next = SheetFmt.updateChartSpec(metaRef.current, activeSheet, drag.id, patch);
      metaRef.current = next;
      setMeta(next);
    };
    const onUp = () => {
      if (!chartDragRef.current) return;
      chartDragRef.current = null;
      syncWorkbook(sheetsRef.current, sheetNamesRef.current);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [activeSheet, syncWorkbook]);

  const removeChart = (id: string) => {
    commitMeta(SheetFmt.removeChartSpec(metaRef.current, activeSheet, id));
    showNotice('Диаграмма удалена');
  };

  /* ---------------- Проверка данных ---------------- */

  const applyValidation = () => {
    const values = validationText
      .split(/[;\n,]/)
      .map((value) => value.trim())
      .filter(Boolean);
    if (values.length === 0) {
      showNotice('Введите значения списка через точку с запятой');
      return;
    }
    commitMeta(
      SheetFmt.setValidationForRange(metaRef.current, activeSheet, selectionAsRange(), values)
    );
    setValidationPanelOpen(false);
    showNotice(`Проверка данных: ${values.length} значений`);
  };

  const removeValidation = () => {
    commitMeta(
      SheetFmt.clearValidationForRange(metaRef.current, activeSheet, selectionAsRange())
    );
    setValidationPanelOpen(false);
    showNotice('Проверка данных удалена');
  };

  /* ---------------- Примечания ---------------- */

  const editComment = async () => {
    const address = SheetFmt.cellAddress(selected.row, selected.col);
    const current = activeMeta.comments[address]?.text ?? '';
    const text = await askText('Примечание к ячейке ' + address, current);
    if (text === null) return;
    commitMeta(
      SheetFmt.setCommentForCell(metaRef.current, activeSheet, selected.row, selected.col, text)
    );
    setCellMenu(null);
    showNotice(text.trim() ? 'Примечание сохранено' : 'Примечание удалено');
  };

  const removeComment = () => {
    commitMeta(
      SheetFmt.setCommentForCell(metaRef.current, activeSheet, selected.row, selected.col, null)
    );
    setCellMenu(null);
    showNotice('Примечание удалено');
  };

  /* ---------------- Условное форматирование ---------------- */

  const createConditionalRule = () => {
    const needsValue =
      ruleOperator !== 'empty' && ruleOperator !== 'notEmpty' && ruleOperator !== 'scale';
    if (needsValue && !ruleValue.trim()) {
      showNotice('Укажите значение условия');
      return;
    }
    const rule: SheetFmt.ConditionalRule = {
      id: `rule_${Date.now()}`,
      range: SheetFmt.rangeToAddress(selectionAsRange()),
      operator: ruleOperator,
      value: ruleValue,
      value2: ruleValue2,
      fill: ruleOperator === 'scale' ? undefined : ruleFill,
      color: ruleOperator === 'scale' ? undefined : ruleColor,
      bold: ruleBold,
    };
    commitMeta(SheetFmt.addConditionalRule(metaRef.current, activeSheet, rule));
    showNotice('Правило условного форматирования добавлено');
  };

  const deleteConditionalRule = (id: string) => {
    commitMeta(SheetFmt.removeConditionalRule(metaRef.current, activeSheet, id));
  };

  /* ---------------- Сводная таблица ---------------- */

  const createPivot = () => {
    const range = SheetFmt.detectUsedRange(grid, selected.row, selected.col);
    if (range.endRow <= range.startRow) {
      showNotice('Нет данных для сводной таблицы');
      return;
    }
    const groups = new Map<string, number[]>();
    for (let row = range.startRow + 1; row <= range.endRow; row++) {
      const key = displayCell(row, pivotRowCol).trim() || '(пусто)';
      const value = numericFromText(displayCell(row, pivotValueCol));
      if (value === null) continue;
      const list = groups.get(key) ?? [];
      list.push(value);
      groups.set(key, list);
    }
    if (groups.size === 0) {
      showNotice('В выбранном столбце нет чисел');
      return;
    }
    const aggregate = (values: number[]): number | string => {
      switch (pivotAgg) {
        case 'count':
          return values.length;
        case 'average':
          return Math.round((values.reduce((sum, value) => sum + value, 0) / values.length) * 100) / 100;
        case 'min':
          return Math.min(...values);
        case 'max':
          return Math.max(...values);
        default:
          return Math.round(values.reduce((sum, value) => sum + value, 0) * 100) / 100;
      }
    };
    const rowHeader = displayCell(range.startRow, pivotRowCol) || columnName(pivotRowCol);
    const valueHeader = displayCell(range.startRow, pivotValueCol) || columnName(pivotValueCol);
    const aggLabels: Record<string, string> = {
      sum: 'Сумма',
      count: 'Количество',
      average: 'Среднее',
      min: 'Минимум',
      max: 'Максимум',
    };
    const matrix: Grid = [
      [rowHeader, `${aggLabels[pivotAgg]} — ${valueHeader}`],
      ...[...groups.entries()]
        .sort((a, b) => a[0].localeCompare(b[0], 'ru'))
        .map(([key, values]) => [key, aggregate(values)]),
    ];

    let index = 1;
    const existing = new Set(sheetNamesRef.current);
    while (existing.has(`Сводная ${index}`)) index++;
    const name = `Сводная ${index}`;
    const nextSheets = { ...sheetsRef.current, [name]: matrix };
    const pivotSpec: SheetFmt.PivotSpec = {
      sourceSheet: activeSheet,
      range: SheetFmt.rangeToAddress(range),
      rowCol: pivotRowCol,
      valueCol: pivotValueCol,
      aggregation: pivotAgg,
    };
    const nextMeta = {
      ...metaRef.current,
      [name]: { ...SheetFmt.emptySheetMeta(), pivotSpec },
    };

    pushHistory();
    metaRef.current = nextMeta;
    setMeta(nextMeta);
    syncWorkbook(nextSheets, [...sheetNamesRef.current, name]);
    setSheets(nextSheets);
    setSheetNames((prev) => [...prev, name]);
    setActiveSheet(name);
    setSelected({ row: 0, col: 0 });
    setPivotOpen(false);
    showNotice(`Сводная таблица: ${name}`);
  };

  /** Пересчитывает активную сводную по её источнику. */
  const refreshPivot = () => {
    const spec = activeMeta.pivotSpec;
    if (!spec) return;
    const sourceGrid = sheetsRef.current[spec.sourceSheet];
    const range = SheetFmt.parseRange(spec.range);
    if (!sourceGrid || !range) {
      showNotice('Источник сводной не найден');
      return;
    }
    const context = createContext(sheetsRef.current, spec.sourceSheet);
    const valueAt = (row: number, col: number) =>
      computeCellDisplay(sheetsRef.current, spec.sourceSheet, row, col, context);
    const groups = new Map<string, number[]>();
    for (let row = range.startRow + 1; row <= range.endRow; row++) {
      const key = valueAt(row, spec.rowCol).trim() || '(пусто)';
      const value = numericFromText(valueAt(row, spec.valueCol));
      if (value === null) continue;
      const list = groups.get(key) ?? [];
      list.push(value);
      groups.set(key, list);
    }
    if (groups.size === 0) {
      showNotice('В источнике нет чисел для сводной');
      return;
    }
    const aggregate = (values: number[]): number => {
      switch (spec.aggregation) {
        case 'count':
          return values.length;
        case 'average':
          return Math.round((values.reduce((sum, value) => sum + value, 0) / values.length) * 100) / 100;
        case 'min':
          return Math.min(...values);
        case 'max':
          return Math.max(...values);
        default:
          return Math.round(values.reduce((sum, value) => sum + value, 0) * 100) / 100;
      }
    };
    const aggLabels: Record<string, string> = {
      sum: 'Сумма',
      count: 'Количество',
      average: 'Среднее',
      min: 'Минимум',
      max: 'Максимум',
    };
    const matrix: Grid = [
      [
        valueAt(range.startRow, spec.rowCol) || columnName(spec.rowCol),
        `${aggLabels[spec.aggregation]} — ${
          valueAt(range.startRow, spec.valueCol) || columnName(spec.valueCol)
        }`,
      ],
      ...[...groups.entries()]
        .sort((a, b) => a[0].localeCompare(b[0], 'ru'))
        .map(([key, values]) => [key, aggregate(values)]),
    ];
    applyBatch((draft) => {
      draft[activeSheet] = matrix;
    });
    showNotice('Сводная таблица обновлена');
  };

  /* ---------------- Сортировка ---------------- */

  const sortByColumn = (direction: 'asc' | 'desc') => {
    if (!grid.length) return;
    const range = SheetFmt.detectUsedRange(grid, selected.row, selected.col);
    // Первая строка считается заголовком и остаётся на месте.
    const headerRow = range.startRow;
    const count = range.endRow - headerRow;
    if (count < 2) {
      showNotice('Недостаточно строк для сортировки');
      return;
    }
    const numeric = (value: string): number | null => {
      const normalized = value.replace(/\s/g, '').replace(',', '.');
      if (normalized === '' || !/^-?\d+(\.\d+)?$/.test(normalized)) return null;
      const num = Number(normalized);
      return Number.isFinite(num) ? num : null;
    };
    const order = Array.from({ length: count }, (_, index) => index);
    order.sort((a, b) => {
      const av = displayCell(headerRow + 1 + a, selected.col);
      const bv = displayCell(headerRow + 1 + b, selected.col);
      const an = numeric(av);
      const bn = numeric(bv);
      const result =
        an !== null && bn !== null
          ? an - bn
          : av.localeCompare(bv, 'ru', { numeric: true, sensitivity: 'base' });
      return direction === 'asc' ? result : -result;
    });

    if (order.every((value, index) => value === index)) {
      showNotice('Данные уже отсортированы');
      return;
    }

    const rows = grid.slice(headerRow + 1, range.endRow + 1);
    applyBatch(
      (draft) => {
        const sheet = draft[activeSheet];
        if (!sheet) return;
        for (let index = 0; index < order.length; index++) {
          sheet[headerRow + 1 + index] = [...rows[order[index]]];
        }
      },
      (draftMeta) => {
        const next = SheetFmt.reorderRowsInMeta(
          draftMeta,
          activeSheet,
          headerRow + 1,
          range.endRow,
          order
        );
        draftMeta[activeSheet] = next[activeSheet];
      }
    );
    showNotice(
      `Сортировка ${direction === 'asc' ? 'по возрастанию' : 'по убыванию'}: ${columnName(
        selected.col
      )}`
    );
  };

  /* ---------------- Поиск и замена ---------------- */

  const scrollToRow = (row: number) => {
    const el = gridRef.current;
    if (!el) return;
    const top = rowOffsets[row] ?? row * SheetFmt.DEFAULT_ROW_HEIGHT;
    if (top < el.scrollTop || top > el.scrollTop + el.clientHeight - 80) {
      el.scrollTop = Math.max(0, top - 80);
    }
  };

  const findNext = (forward = true) => {
    const query = findText.trim().toLowerCase();
    if (!query) {
      setFindMessage('Введите текст для поиска');
      return;
    }
    const rows = grid.length;
    const cols = grid[0]?.length ?? 0;
    const total = rows * cols;
    if (!total) return;
    const startIndex = selected.row * cols + selected.col;
    for (let step = 1; step <= total; step++) {
      const index = ((startIndex + (forward ? step : -step)) % total + total) % total;
      const row = Math.floor(index / cols);
      const col = index % cols;
      if (displayCell(row, col).toLowerCase().includes(query)) {
        selectSingle({ row, col });
        scrollToRow(row);
        setFindMessage(`Найдено: ${columnName(col)}${row + 1}`);
        return;
      }
    }
    setFindMessage('Совпадений нет');
  };

  const replaceCurrent = () => {
    const raw = String(grid[selected.row]?.[selected.col] ?? '');
    if (!findText) return;
    if (!raw.toLowerCase().includes(findText.toLowerCase())) {
      findNext();
      return;
    }
    const regex = new RegExp(escapeRegExp(findText), 'gi');
    applyChange(selected.row, selected.col, raw.replace(regex, replaceText));
    setFindMessage('Заменено');
    findNext();
  };

  const replaceAll = () => {
    if (!findText) return;
    const regex = new RegExp(escapeRegExp(findText), 'gi');
    let count = 0;
    applyBatch((draft) => {
      const sheet = draft[activeSheet];
      if (!sheet) return;
      for (let row = 0; row < sheet.length; row++) {
        for (let col = 0; col < (sheet[row]?.length ?? 0); col++) {
          const raw = String(sheet[row][col] ?? '');
          if (!raw || raw.startsWith('=')) continue;
          regex.lastIndex = 0;
          if (!regex.test(raw)) continue;
          regex.lastIndex = 0;
          sheet[row][col] = raw.replace(regex, replaceText);
          count++;
        }
      }
    });
    showNotice(count > 0 ? `Заменено ячеек: ${count}` : 'Совпадений нет');
  };

  /* ---------------- Изменение ширины и высоты ---------------- */

  const startColumnResize = (event: React.MouseEvent, col: number) => {
    event.preventDefault();
    event.stopPropagation();
    pushHistory();
    resizeRef.current = {
      type: 'col',
      index: col,
      start: event.clientX,
      size: activeMeta.colWidths[col] ?? SheetFmt.DEFAULT_COL_WIDTH,
    };
    document.body.style.cursor = 'col-resize';
  };

  const startRowResize = (event: React.MouseEvent, row: number) => {
    event.preventDefault();
    event.stopPropagation();
    pushHistory();
    resizeRef.current = {
      type: 'row',
      index: row,
      start: event.clientY,
      size: activeMeta.rowHeights[row] ?? SheetFmt.DEFAULT_ROW_HEIGHT,
    };
    document.body.style.cursor = 'row-resize';
  };

  useEffect(() => {
    const onMove = (event: MouseEvent) => {
      const drag = resizeRef.current;
      if (!drag) return;
      const delta = drag.type === 'col' ? event.clientX - drag.start : event.clientY - drag.start;
      const nextMeta =
        drag.type === 'col'
          ? SheetFmt.setColWidth(metaRef.current, activeSheet, drag.index, drag.size + delta)
          : SheetFmt.setRowHeight(metaRef.current, activeSheet, drag.index, drag.size + delta);
      metaRef.current = nextMeta;
      setMeta(nextMeta);
    };
    const onUp = () => {
      if (!resizeRef.current) return;
      resizeRef.current = null;
      document.body.style.cursor = '';
      syncWorkbook(sheetsRef.current, sheetNamesRef.current);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [activeSheet, syncWorkbook]);

  const escapeRegExp = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

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
    if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'x' || e.key.toLowerCase() === 'ч')) {
      e.preventDefault();
      void copySelection();
      clearSelection();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'b' || e.key.toLowerCase() === 'и')) {
      e.preventDefault();
      toggleStyle('bold');
      return;
    }
    if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'i' || e.key.toLowerCase() === 'ш')) {
      e.preventDefault();
      toggleStyle('italic');
      return;
    }
    if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'u' || e.key.toLowerCase() === 'г')) {
      e.preventDefault();
      toggleStyle('underline');
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
        if (
          extraCells.length > 0 ||
          selectionRange.start.row !== selectionRange.end.row ||
          selectionRange.start.col !== selectionRange.end.col
        ) {
          clearSelection();
        } else {
          applyChange(row, col, '');
        }
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
    applyBatch(
      (draft) => {
        const current = draft[activeSheet];
        if (current.length <= 1) return;
        draft[activeSheet] = current.filter((_, i) => i !== rowIdx);
      },
      (draftMeta) => {
        const next = SheetFmt.removeRowFromMeta(draftMeta, activeSheet, rowIdx);
        draftMeta[activeSheet] = next[activeSheet];
      }
    );
    setSelected((s) => ({ ...s, row: Math.max(0, Math.min(s.row, grid.length - 2)) }));
  };

  const deleteColumn = (colIdx: number) => {
    applyBatch(
      (draft) => {
        const current = draft[activeSheet];
        if ((current[0]?.length ?? 0) <= 1) return;
        draft[activeSheet] = current.map((r) => r.filter((_, i) => i !== colIdx));
      },
      (draftMeta) => {
        const next = SheetFmt.removeColumnFromMeta(draftMeta, activeSheet, colIdx);
        draftMeta[activeSheet] = next[activeSheet];
      }
    );
    setSelected((s) => ({ ...s, col: Math.max(0, Math.min(s.col, (grid[0]?.length ?? 2) - 2)) }));
  };

  const addSheet = (requestedName?: string) => {
    const existing = new Set(sheetNames);
    let name = requestedName?.trim() ?? '';
    if (!name || existing.has(name)) {
      let n = sheetNames.length + 1;
      while (existing.has(`Лист ${n}`)) n++;
      name = `Лист ${n}`;
    }
    const newGrid: Grid = Array.from({ length: MIN_ROWS }, () => new Array(MIN_COLS).fill(''));

    const nextSheets = { ...sheetsRef.current, [name]: newGrid };
    const nextMeta = { ...metaRef.current, [name]: SheetFmt.emptySheetMeta() };
    setHistory((h) => [...h.slice(-49), { sheets: sheetsRef.current, meta: metaRef.current }]);
    setFuture([]);
    metaRef.current = nextMeta;
    setMeta(nextMeta);
    sheetsRef.current = nextSheets;
    const nextNames = [...sheetNames, name];
    sheetNamesRef.current = nextNames;
    syncWorkbook(nextSheets, nextNames);
    setSheets(nextSheets);
    setSheetNames(nextNames);
    setActiveSheet(name);
    setSelected({ row: 0, col: 0 });
  };

  const renameSheet = (oldName: string, newName: string) => {
    const trimmed = newName.trim();
    if (!trimmed || trimmed === oldName || sheetNames.includes(trimmed)) return;
    const nextSheets: Workbook = { ...sheetsRef.current, [trimmed]: sheetsRef.current[oldName] };
    delete nextSheets[oldName];
    const names = sheetNamesRef.current.map((name) => (name === oldName ? trimmed : name));
    const nextMeta = { ...metaRef.current, [trimmed]: metaRef.current[oldName] ?? SheetFmt.emptySheetMeta() };
    delete nextMeta[oldName];

    pushHistory();
    metaRef.current = nextMeta;
    setMeta(nextMeta);
    sheetsRef.current = nextSheets;
    sheetNamesRef.current = names;
    syncWorkbook(nextSheets, names);
    setSheets(nextSheets);
    setSheetNames(names);
    setActiveSheet((cur) => (cur === oldName ? trimmed : cur));
  };

  const removeSheet = (name: string) => {
    if (sheetNamesRef.current.length <= 1) {
      showNotice('Нельзя удалить единственный лист');
      return;
    }
    const index = sheetNamesRef.current.indexOf(name);
    const names = sheetNamesRef.current.filter((item) => item !== name);
    const nextSheets = { ...sheetsRef.current };
    delete nextSheets[name];
    const nextMeta = { ...metaRef.current };
    delete nextMeta[name];

    pushHistory();
    metaRef.current = nextMeta;
    setMeta(nextMeta);
    sheetsRef.current = nextSheets;
    sheetNamesRef.current = names;
    syncWorkbook(nextSheets, names);
    setSheets(nextSheets);
    setSheetNames(names);
    setActiveSheet((current) =>
      current === name ? names[Math.max(0, index - 1)] ?? names[0] : current
    );
    setSelected({ row: 0, col: 0 });
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

  /* ---------------- Мост для AI-инструментов таблицы ---------------- */

  useEffect(() => {
    return registerSheetController({
      tabId: tab.id,
      serialize: async () =>
        (await syncWorkbook(sheetsRef.current, sheetNamesRef.current)) ?? tab.content,
      getActiveSheet: () => activeSheet,
      readSheet: (sheet, range, limit) => {
        const target = sheet && sheetsRef.current[sheet] ? sheet : activeSheet;
        const grid = sheetsRef.current[target];
        if (!grid) return null;

        let start: CellCoord = { row: 0, col: 0 };
        let end: CellCoord = { row: 0, col: 0 };
        if (range) {
          const parsed = parseTarget(`${target}!${range}`);
          if (!parsed) return null;
          start = parsed.range.start;
          end = parsed.range.end;
        } else {
          let lastRow = 0;
          let lastCol = 0;
          for (let row = 0; row < grid.length; row += 1) {
            for (let col = 0; col < (grid[row]?.length ?? 0); col += 1) {
              if (String(grid[row][col] ?? '').trim()) {
                lastRow = Math.max(lastRow, row);
                lastCol = Math.max(lastCol, col);
              }
            }
          }
          end = { row: lastRow, col: lastCol };
        }

        const maxCells = Math.max(1, Math.min(limit ?? 400, 1200));
        const rows: string[][] = [];
        const cells: SheetReadResult['cells'] = [];
        let truncated = false;

        outer: for (let row = start.row; row <= end.row; row += 1) {
          const values: string[] = [];
          for (let col = start.col; col <= end.col; col += 1) {
            if (cells.length >= maxCells) {
              truncated = true;
              break outer;
            }
            const raw = String(grid[row]?.[col] ?? '');
            const display = computeCellDisplay(sheetsRef.current, target, row, col);
            values.push(display);
            cells.push({
              address: `${columnName(col)}${row + 1}`,
              value: raw,
              formula: raw.startsWith('=') ? raw : '',
              display,
            });
          }
          rows.push(values);
        }

        return {
          sheet: target,
          sheetNames: sheetNamesRef.current,
          range: rangeToA1(normalizeRange(start, end)),
          rows,
          cells,
          truncated,
        };
      },
      getSelection: () => {
        const target = activeSheet;
        const grid = sheetsRef.current[target];
        if (!grid) return null;
        const start = selectionRange.start;
        const end = selectionRange.end;
        const rows: string[][] = [];
        const cells: SheetReadResult['cells'] = [];
        for (let row = start.row; row <= end.row; row += 1) {
          const values: string[] = [];
          for (let col = start.col; col <= end.col; col += 1) {
            const raw = String(grid[row]?.[col] ?? '');
            const display = computeCellDisplay(sheetsRef.current, target, row, col);
            values.push(display);
            cells.push({
              address: `${columnName(col)}${row + 1}`,
              value: raw,
              formula: raw.startsWith('=') ? raw : '',
              display,
            });
          }
          rows.push(values);
        }
        return {
          sheet: target,
          sheetNames: sheetNamesRef.current,
          range: rangeToA1(normalizeRange(start, end)),
          rows,
          cells,
          truncated: false,
        };
      },
      writeCells: (sheet, updates, _transactionId) => {        const target = sheet && sheetsRef.current[sheet] ? sheet : activeSheet;
        const valid = updates.filter((item) => item && item.address);
        if (valid.length === 0) return { ok: false, message: 'Нет ячеек для записи' };
        const changes = valid
          .map((item) => {
            const parsed = parseTarget(
              item.address.includes('!') ? item.address : `${target}!${item.address}`
            );
            if (!parsed) return null;
            const sheetName = sheetsRef.current[parsed.sheet] ? parsed.sheet : target;
            const grid = sheetsRef.current[sheetName];
            const { row, col } = parsed.range.start;
            const before = String(grid?.[row]?.[col] ?? '');
            const after = String(item.value ?? '');
            return before === after
              ? null
              : {
                  address: `${sheetName}!${columnName(col)}${row + 1}`,
                  before,
                  after,
                };
          })
          .filter(
            (item): item is { address: string; before: string; after: string } => Boolean(item)
          );

        let first: { value: CellCoord | null } = { value: null };
        let written = 0;
        applyBatch((draft) => {
          for (const item of valid) {
            const parsed = parseTarget(
              item.address.includes('!') ? item.address : `${target}!${item.address}`
            );
            if (!parsed) continue;
            const targetSheet = draft[parsed.sheet] ? parsed.sheet : target;
            const grid = draft[targetSheet];
            if (!grid) continue;
            const { row, col } = parsed.range.start;
            while (grid.length <= row) {
              grid.push(new Array(grid[0]?.length ?? MIN_COLS).fill(''));
            }
            for (const line of grid) {
              while (line.length <= col) line.push('');
            }
            grid[row][col] = normalizeFormula(String(item.value ?? ''));
            if (!first.value) first.value = { row, col };
            written += 1;
          }
        });

        if (written === 0) return { ok: false, message: 'Не удалось разобрать адреса ячеек' };
        setActiveSheet(target);
        if (first.value) setSelected(first.value);
        if (changes.length > 0) setAiDiff({ cells: changes });
        return { ok: true, message: `Записано ячеек: ${written}`, changed: true };
      },
      addSheet: (name) => {
        addSheet(name);
        return { ok: true, message: name ? `Лист «${name}» добавлен` : 'Лист добавлен', changed: true };
      },
      renameSheet: (oldName, newName) => {
        if (!sheetsRef.current[oldName]) {
          return { ok: false, message: `Лист «${oldName}» не найден` };
        }
        if (!newName.trim() || sheetNamesRef.current.includes(newName.trim())) {
          return { ok: false, message: `Имя «${newName}» пустое или уже занято` };
        }
        renameSheet(oldName, newName);
        return { ok: true, message: `Лист переименован в «${newName.trim()}»`, changed: true };
      },
      deleteSheet: (name) => {
        if (!sheetsRef.current[name]) return { ok: false, message: `Лист «${name}» не найден` };
        if (sheetNamesRef.current.length <= 1) {
          return { ok: false, message: 'Нельзя удалить единственный лист' };
        }
        removeSheet(name);
        return { ok: true, message: `Лист «${name}» удалён`, changed: true };
      },
      setActiveSheet: (name) => {
        if (!sheetsRef.current[name]) return { ok: false, message: `Лист «${name}» не найден` };
        setActiveSheet(name);
        return { ok: true, message: `Активный лист: «${name}»`, changed: false };
      },
      insertRows: (sheet, index, count) => {
        const target = sheet && sheetsRef.current[sheet] ? sheet : activeSheet;
        const amount = Math.max(1, Math.min(count || 1, 200));
        applyBatch((draft) => {
          const grid = draft[target];
          if (!grid) return;
          const blank = new Array(grid[0]?.length ?? MIN_COLS).fill('');
          for (let i = 0; i < amount; i += 1) {
            grid.splice(Math.max(0, Math.min(index, grid.length)), 0, [...blank]);
          }
        });
        return { ok: true, message: `Вставлено строк: ${amount}`, changed: true };
      },
      deleteRows: (sheet, index, count) => {
        const target = sheet && sheetsRef.current[sheet] ? sheet : activeSheet;
        const grid = sheetsRef.current[target];
        if (!grid) return { ok: false, message: `Лист «${target}» не найден` };
        const start = Math.max(0, Math.min(index, Math.max(0, grid.length - 1)));
        const amount = Math.max(1, Math.min(count || 1, grid.length - start));
        if (grid.length - amount < 1) return { ok: false, message: 'Нельзя удалить все строки' };
        applyBatch(
          (draft) => {
            draft[target]?.splice(start, amount);
          },
          (draftMeta) => {
            for (let i = 0; i < amount; i += 1) {
              const next = SheetFmt.removeRowFromMeta(draftMeta, target, start);
              draftMeta[target] = next[target];
            }
          }
        );
        return { ok: true, message: `Удалено строк: ${amount}`, changed: true };
      },
      insertColumns: (sheet, index, count) => {
        const target = sheet && sheetsRef.current[sheet] ? sheet : activeSheet;
        const amount = Math.max(1, Math.min(count || 1, 100));
        applyBatch((draft) => {
          const grid = draft[target];
          if (!grid) return;
          for (const row of grid) {
            for (let i = 0; i < amount; i += 1) {
              row.splice(Math.max(0, Math.min(index, row.length)), 0, '');
            }
          }
        });
        return { ok: true, message: `Вставлено столбцов: ${amount}`, changed: true };
      },
      deleteColumns: (sheet, index, count) => {
        const target = sheet && sheetsRef.current[sheet] ? sheet : activeSheet;
        const grid = sheetsRef.current[target];
        if (!grid) return { ok: false, message: `Лист «${target}» не найден` };
        const width = grid[0]?.length ?? 0;
        const start = Math.max(0, Math.min(index, Math.max(0, width - 1)));
        const amount = Math.max(1, Math.min(count || 1, width - start));
        if (width - amount < 1) return { ok: false, message: 'Нельзя удалить все столбцы' };
        applyBatch(
          (draft) => {
            const current = draft[target];
            if (!current) return;
            draft[target] = current.map((row) => row.filter((_, col) => col < start || col >= start + amount));
          },
          (draftMeta) => {
            for (let i = 0; i < amount; i += 1) {
              const next = SheetFmt.removeColumnFromMeta(draftMeta, target, start);
              draftMeta[target] = next[target];
            }
          }
        );
        return { ok: true, message: `Удалено столбцов: ${amount}`, changed: true };
      },
      mergeCells: (sheet, rangeText) => {
        const target = sheet && sheetsRef.current[sheet] ? sheet : activeSheet;
        const range = SheetFmt.parseRange(rangeText);
        if (!range) return { ok: false, message: `Не разобрать диапазон «${rangeText}»` };
        commitMeta(SheetFmt.addMerge(metaRef.current, target, range));
        return { ok: true, message: `Объединено: ${rangeText}`, changed: true };
      },
      unmergeCells: (sheet, rangeText) => {
        const target = sheet && sheetsRef.current[sheet] ? sheet : activeSheet;
        const range = SheetFmt.parseRange(rangeText);
        if (!range) return { ok: false, message: `Не разобрать диапазон «${rangeText}»` };
        commitMeta(SheetFmt.removeMergesInRange(metaRef.current, target, range));
        return { ok: true, message: `Объединение снято: ${rangeText}`, changed: true };
      },
      formatRange: (sheet, rangeText, patch) => {
        const target = sheet && sheetsRef.current[sheet] ? sheet : activeSheet;
        const range = SheetFmt.parseRange(rangeText);
        if (!range) return { ok: false, message: `Не разобрать диапазон «${rangeText}»` };
        commitMeta(SheetFmt.applyStyleToRange(metaRef.current, target, range, patch));
        return { ok: true, message: `Формат применён: ${rangeText}`, changed: true };
      },
    });
  }, [tab.id, activeSheet, addSheet, renameSheet, removeSheet, parseTarget, applyBatch, commitMeta, selectionRange, syncWorkbook]);

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
          <button className="block w-full rounded px-3 py-1.5 text-left hover:bg-ws-hover" onClick={editComment}>Примечание…</button>
          {activeMeta.comments[SheetFmt.cellAddress(selected.row, selected.col)] && (
            <button className="block w-full rounded px-3 py-1.5 text-left text-ws-danger hover:bg-red-500/10" onClick={removeComment}>Удалить примечание</button>
          )}
          <div className="my-1 border-t border-ws" />
          <button className="block w-full rounded px-3 py-1.5 text-left text-ws-danger hover:bg-red-500/10" onClick={() => { deleteSelectionRows(); setCellMenu(null); }}>Удалить строки</button>
          <button className="block w-full rounded px-3 py-1.5 text-left text-ws-danger hover:bg-red-500/10" onClick={() => { deleteSelectionColumns(); setCellMenu(null); }}>Удалить столбцы</button>
        </div>
      )}
      {aiDiff && aiDiff.cells.length > 0 && (
        <div
          data-ai-diff="1"
          className="flex shrink-0 items-center gap-2 border-b border-blue-500/30 bg-blue-500/10 px-3 py-1 text-[11px] text-ws"
        >
          <span className="font-medium text-blue-400">AI изменил ячеек: {aiDiff.cells.length}</span>
          <span className="flex-1 truncate text-ws-muted">
            {aiDiff.cells
              .slice(0, 6)
              .map((cell) =>
                cell.before
                  ? `${cell.address}: ${cell.before} → ${cell.after || '(пусто)'}`
                  : `${cell.address}: ${cell.after || '(пусто)'}`
              )
              .join(' · ')}
          </span>
          <button
            onClick={() => {
              undo();
              setAiDiff(null);
            }}
            className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 text-[10px] hover:bg-ws-hover"
          >
            Отменить
          </button>
          <button
            onClick={() => setAiDiff(null)}
            title="Скрыть"
            className="rounded px-1 text-ws-faint hover:bg-ws-hover hover:text-ws"
          >
            ×
          </button>
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
            onClick={() => void reloadFromDisk()}
            disabled={excelBusy}
            title="Перечитать файл с диска"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-ws-accent hover:text-ws-accent disabled:opacity-40"
          >
            <RefreshCw size={11} />
            Обновить
          </button>

          <button
            onClick={() => void openInExcel()}
            disabled={excelBusy || isCsv}
            title="Открыть в настоящем Excel"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-emerald-500 hover:text-emerald-400 disabled:opacity-40"
          >
            <ExternalLink size={11} />
            Excel
          </button>

          <button
            onClick={() => void runExcelMacro()}
            disabled={excelBusy || tab.extension !== 'xlsm'}
            title="Выполнить VBA-макрос в Excel"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-amber-500 hover:text-amber-400 disabled:opacity-40"
          >
            <Play size={11} />
            VBA
          </button>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

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

      {/* Панель форматирования */}
      <div className="flex shrink-0 flex-wrap items-center gap-1 overflow-x-auto border-b border-ws bg-ws-surface px-3 py-1 text-xs select-none">
        <button onClick={() => toggleStyle('bold')} title="Полужирный (Ctrl+B)" className={cn('rounded p-1 hover:bg-ws-hover', selectedStyle.bold && 'bg-ws-accent-soft text-ws-accent')}><Bold size={13} /></button>
        <button onClick={() => toggleStyle('italic')} title="Курсив (Ctrl+I)" className={cn('rounded p-1 hover:bg-ws-hover', selectedStyle.italic && 'bg-ws-accent-soft text-ws-accent')}><Italic size={13} /></button>
        <button onClick={() => toggleStyle('underline')} title="Подчёркнутый (Ctrl+U)" className={cn('rounded p-1 hover:bg-ws-hover', selectedStyle.underline && 'bg-ws-accent-soft text-ws-accent')}><Underline size={13} /></button>
        <button onClick={() => toggleStyle('strike')} title="Зачёркнутый" className={cn('rounded p-1 hover:bg-ws-hover', selectedStyle.strike && 'bg-ws-accent-soft text-ws-accent')}><Strikethrough size={13} /></button>

        <div className="mx-1 h-3.5 w-px border-l border-ws" />

        <select
          value={selectedStyle.fontSize ?? 12}
          onChange={(e) => applyStyle({ fontSize: Number(e.target.value) })}
          title="Размер шрифта"
          className="rounded border border-ws bg-ws-editor px-1 py-0.5 text-[11px] text-ws"
        >
          {[9, 10, 11, 12, 14, 16, 18, 20, 24, 28].map((size) => (
            <option key={size} value={size}>{size}</option>
          ))}
        </select>

        <label title="Цвет текста" className="flex cursor-pointer items-center gap-1 rounded border border-ws px-1.5 py-0.5">
          <Palette size={12} className="text-ws-muted" />
          <input type="color" value={selectedStyle.color ?? '#e2e8f0'} onChange={(e) => applyStyle({ color: e.target.value })} className="h-4 w-5 cursor-pointer border-0 bg-transparent p-0" />
        </label>
        <label title="Цвет заливки" className="flex cursor-pointer items-center gap-1 rounded border border-ws px-1.5 py-0.5">
          <PaintBucket size={12} className="text-ws-muted" />
          <input type="color" value={selectedStyle.fill ?? '#ffffff'} onChange={(e) => applyStyle({ fill: e.target.value })} className="h-4 w-5 cursor-pointer border-0 bg-transparent p-0" />
        </label>

        <div className="mx-1 h-3.5 w-px border-l border-ws" />

        <button onClick={() => applyStyle({ align: 'left' })} title="По левому краю" className={cn('rounded p-1 hover:bg-ws-hover', selectedStyle.align === 'left' && 'bg-ws-accent-soft text-ws-accent')}><AlignLeft size={13} /></button>
        <button onClick={() => applyStyle({ align: 'center' })} title="По центру" className={cn('rounded p-1 hover:bg-ws-hover', selectedStyle.align === 'center' && 'bg-ws-accent-soft text-ws-accent')}><AlignCenter size={13} /></button>
        <button onClick={() => applyStyle({ align: 'right' })} title="По правому краю" className={cn('rounded p-1 hover:bg-ws-hover', selectedStyle.align === 'right' && 'bg-ws-accent-soft text-ws-accent')}><AlignRight size={13} /></button>
        <button onClick={() => toggleStyle('wrap')} title="Перенос текста" className={cn('rounded p-1 hover:bg-ws-hover', selectedStyle.wrap && 'bg-ws-accent-soft text-ws-accent')}><WrapText size={13} /></button>

        <div className="mx-1 h-3.5 w-px border-l border-ws" />

        <span className="flex items-center gap-0.5 rounded border border-ws px-1 py-0.5">
          <span className="px-0.5 text-[10px] text-ws-faint">Границы</span>
          <button onClick={() => applyBorders('all')} title="Все границы" className="rounded px-1 hover:bg-ws-hover">Все</button>
          <button onClick={() => applyBorders('outer')} title="Только внешние" className="rounded px-1 hover:bg-ws-hover">Внеш.</button>
          <button onClick={() => applyBorders('none')} title="Убрать границы" className="rounded px-1 hover:bg-ws-hover">Нет</button>
        </span>

        <button onClick={mergeSelection} title="Объединить ячейки" className="flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover"><Merge size={12} />Объединить</button>
        <button onClick={unmergeSelection} title="Отменить объединение" className="rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover"><SplitSquareHorizontal size={12} /></button>

        <div className="mx-1 h-3.5 w-px border-l border-ws" />

        <select
          value={selectedStyle.numberFormat ?? 'general'}
          onChange={(e) => applyStyle({ numberFormat: e.target.value as SheetFmt.NumberFormat })}
          title="Числовой формат"
          className="rounded border border-ws bg-ws-editor px-1 py-0.5 text-[11px] text-ws"
        >
          <option value="general">Общий</option>
          <option value="number">Число</option>
          <option value="currency">Финансовый (₽)</option>
          <option value="percent">Процент</option>
          <option value="date">Дата</option>
          <option value="text">Текст</option>
        </select>

        <button onClick={() => sortByColumn('asc')} title="Сортировка по возрастанию" className="flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover"><ArrowUpAZ size={12} /></button>
        <button onClick={() => sortByColumn('desc')} title="Сортировка по убыванию" className="flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover"><ArrowDownAZ size={12} /></button>
        <button
          onClick={openFilterPanel}
          title="Фильтр по значениям выбранного столбца"
          className={cn('flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover', activeFilter && 'border-emerald-500/40 text-emerald-500')}
        >
          <TableIcon size={12} />
          Фильтр
        </button>

        <div className="mx-1 h-3.5 w-px border-l border-ws" />

        <button onClick={applyColorScale} title="Цветовая шкала по числам выделения" className="rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover">Шкала</button>
        <button onClick={clearFormatting} title="Снять форматирование выделения" className="rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover">Очистить формат</button>
        <button onClick={convertTextToNumbers} title="Преобразовать текстовые числа в числа Excel (для SUM)" className="rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover">Числа</button>

        <div className="mx-1 h-3.5 w-px border-l border-ws" />

        <button
          onClick={() => { setChartTitle(`Диаграмма ${activeMeta.charts.length + 1}`); setChartDialogOpen(true); }}
          title="Построить диаграмму по выделенному диапазону"
          className="flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover"
        >
          <BarChart3 size={12} />Диаграмма
        </button>
        {activeMeta.charts.length > 0 && (
          <span className="rounded border border-emerald-500/30 bg-emerald-500/10 px-1.5 py-0.5 text-[10px] text-emerald-500">
            Диаграммы сохраняются в xlsx
          </span>
        )}
        <button
          onClick={() => setValidationPanelOpen((value) => !value)}
          title="Проверка данных: выпадающий список"
          className={cn('flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover', validationPanelOpen && 'bg-ws-accent-soft text-ws-accent')}
        >
          <ListChecks size={12} />Список
        </button>
        <button
          onClick={() => setConditionalPanelOpen((value) => !value)}
          title="Условное форматирование"
          className={cn('flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover', conditionalPanelOpen && 'bg-ws-accent-soft text-ws-accent')}
        >
          <SlidersHorizontal size={12} />Условное
        </button>
        <button
          onClick={() => setPivotOpen(true)}
          title="Сводная таблица по выделенному диапазону"
          className="flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover"
        >
          <Table2 size={12} />Сводная
        </button>
        {activeMeta.pivotSpec && (
          <button
            onClick={refreshPivot}
            title="Пересчитать сводную по исходным данным"
            className="flex items-center gap-1 rounded border border-emerald-500/40 px-1.5 py-0.5 text-emerald-500 hover:bg-emerald-500/10"
          >
            <RefreshCw size={12} />Обновить сводную
          </button>
        )}

        <div className="mx-1 h-3.5 w-px border-l border-ws" />

        <button
          onClick={freezePanes}
          title="Закрепить строки выше и столбцы левее выделенной ячейки"
          className={cn('flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover', (activeMeta.freezeRows > 0 || activeMeta.freezeCols > 0) && 'border-emerald-500/40 text-emerald-500')}
        >
          <Pin size={12} />
          Закрепить
        </button>
        {(activeMeta.freezeRows > 0 || activeMeta.freezeCols > 0) && (
          <button onClick={unfreezePanes} title="Снять закрепление" className="rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover"><PinOff size={12} /></button>
        )}

        <button onClick={() => setFindOpen((value) => !value)} title="Найти и заменить (Ctrl+F)" className={cn('flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover', findOpen && 'bg-ws-accent-soft text-ws-accent')}><Search size={12} />Поиск</button>

        {notice && <span className="ml-2 truncate text-[11px] text-emerald-500">{notice}</span>}
      </div>

      {/* Поиск и замена */}
      {findOpen && (
        <div className="flex shrink-0 flex-wrap items-center gap-2 border-b border-ws bg-ws-elevated px-3 py-1.5 text-xs">
          <Search size={13} className="text-ws-muted" />
          <input
            autoFocus
            value={findText}
            onChange={(e) => { setFindText(e.target.value); setFindMessage(''); }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') { e.preventDefault(); findNext(!e.shiftKey); }
              if (e.key === 'Escape') setFindOpen(false);
            }}
            placeholder="Найти"
            className="w-44 rounded border border-ws bg-ws-surface px-2 py-0.5 text-xs text-ws placeholder:text-ws-faint focus:border-emerald-500 focus:outline-none"
          />
          <Replace size={13} className="text-ws-muted" />
          <input
            value={replaceText}
            onChange={(e) => setReplaceText(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); replaceCurrent(); } }}
            placeholder="Заменить на"
            className="w-44 rounded border border-ws bg-ws-surface px-2 py-0.5 text-xs text-ws placeholder:text-ws-faint focus:border-emerald-500 focus:outline-none"
          />
          <button onClick={() => findNext(true)} className="rounded border border-ws px-2 py-0.5 hover:bg-ws-hover">Найти</button>
          <button onClick={replaceCurrent} className="rounded border border-ws px-2 py-0.5 hover:bg-ws-hover">Заменить</button>
          <button onClick={replaceAll} className="rounded border border-ws px-2 py-0.5 hover:bg-ws-hover">Заменить все</button>
          {findMessage && <span className="text-[11px] text-ws-muted">{findMessage}</span>}
          <button onClick={() => setFindOpen(false)} title="Закрыть" className="ml-auto rounded p-1 text-ws-faint hover:bg-ws-hover hover:text-ws"><X size={13} /></button>
        </div>
      )}

      {/* Панель фильтра */}
      {filterPanelOpen && activeFilter && (
        <div className="flex shrink-0 flex-wrap items-center gap-2 border-b border-ws bg-ws-elevated px-3 py-1.5 text-xs">
          <TableIcon size={13} className="text-ws-muted" />
          <span className="text-ws-muted">
            Фильтр: <strong className="text-ws">{columnName(activeFilter.col)}</strong>, строка
            заголовка {activeFilter.headerRow + 1}
          </span>
          <div className="flex max-h-20 flex-wrap gap-x-3 gap-y-1 overflow-auto rounded border border-ws bg-ws-surface px-2 py-1">
            {activeFilter.allValues.map((value) => (
              <label key={value || '__empty__'} className="flex cursor-pointer items-center gap-1 text-[11px] text-ws">
                <input
                  type="checkbox"
                  checked={activeFilter.values.includes(value)}
                  onChange={() => toggleFilterValue(value)}
                />
                {value || '(пустые)'}
              </label>
            ))}
          </div>
          <button onClick={clearFilter} className="rounded border border-ws px-2 py-0.5 hover:bg-ws-hover">Сбросить</button>
          <button onClick={() => setFilterPanelOpen(false)} className="ml-auto rounded p-1 text-ws-faint hover:bg-ws-hover hover:text-ws"><X size={13} /></button>
        </div>
      )}

      {/* Проверка данных */}
      {validationPanelOpen && (
        <div className="flex shrink-0 flex-wrap items-center gap-2 border-b border-ws bg-ws-elevated px-3 py-1.5 text-xs">
          <ListChecks size={13} className="text-ws-muted" />
          <span className="text-ws-muted">
            Проверка данных для <strong className="text-ws">{selectionLabel}</strong>
          </span>
          <input
            value={validationText}
            onChange={(e) => setValidationText(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); applyValidation(); } }}
            placeholder="Значения через ; — например: Да; Нет; Не знаю"
            className="w-80 rounded border border-ws bg-ws-surface px-2 py-0.5 text-xs text-ws placeholder:text-ws-faint focus:border-emerald-500 focus:outline-none"
          />
          <button onClick={applyValidation} className="rounded border border-ws px-2 py-0.5 hover:bg-ws-hover">Применить</button>
          <button onClick={removeValidation} className="rounded border border-ws px-2 py-0.5 text-ws-danger hover:bg-red-500/10">Убрать</button>
          <button onClick={() => setValidationPanelOpen(false)} title="Закрыть" className="ml-auto rounded p-1 text-ws-faint hover:bg-ws-hover hover:text-ws"><X size={13} /></button>
        </div>
      )}

      {/* Условное форматирование */}
      {conditionalPanelOpen && (
        <div className="flex shrink-0 flex-col gap-1 border-b border-ws bg-ws-elevated px-3 py-1.5 text-xs">
          <div className="flex flex-wrap items-center gap-2">
            <SlidersHorizontal size={13} className="text-ws-muted" />
            <span className="text-ws-muted">Условие для <strong className="text-ws">{selectionLabel}</strong></span>
            <select
              value={ruleOperator}
              onChange={(e) => setRuleOperator(e.target.value as SheetFmt.ConditionalOperator)}
              className="rounded border border-ws bg-ws-surface px-1 py-0.5 text-[11px] text-ws"
            >
              <option value="greater">больше</option>
              <option value="less">меньше</option>
              <option value="equal">равно</option>
              <option value="notEqual">не равно</option>
              <option value="between">между</option>
              <option value="contains">содержит</option>
              <option value="empty">пусто</option>
              <option value="notEmpty">не пусто</option>
              <option value="scale">цветовая шкала</option>
            </select>
            {ruleOperator !== 'empty' && ruleOperator !== 'notEmpty' && ruleOperator !== 'scale' && (
              <input
                value={ruleValue}
                onChange={(e) => setRuleValue(e.target.value)}
                placeholder="Значение"
                className="w-28 rounded border border-ws bg-ws-surface px-2 py-0.5 text-xs text-ws placeholder:text-ws-faint focus:border-emerald-500 focus:outline-none"
              />
            )}
            {ruleOperator === 'between' && (
              <input
                value={ruleValue2}
                onChange={(e) => setRuleValue2(e.target.value)}
                placeholder="И"
                className="w-20 rounded border border-ws bg-ws-surface px-2 py-0.5 text-xs text-ws placeholder:text-ws-faint focus:border-emerald-500 focus:outline-none"
              />
            )}
            {ruleOperator !== 'scale' && (
              <>
                <label className="flex items-center gap-1 text-[11px] text-ws-muted">
                  Заливка
                  <input type="color" value={ruleFill} onChange={(e) => setRuleFill(e.target.value)} className="h-4 w-5 cursor-pointer border-0 bg-transparent p-0" />
                </label>
                <label className="flex items-center gap-1 text-[11px] text-ws-muted">
                  Текст
                  <input type="color" value={ruleColor} onChange={(e) => setRuleColor(e.target.value)} className="h-4 w-5 cursor-pointer border-0 bg-transparent p-0" />
                </label>
                <label className="flex items-center gap-1 text-[11px] text-ws-muted">
                  <input type="checkbox" checked={ruleBold} onChange={(e) => setRuleBold(e.target.checked)} /> Жирный
                </label>
              </>
            )}
            <button onClick={createConditionalRule} className="rounded border border-ws px-2 py-0.5 hover:bg-ws-hover">Добавить правило</button>
            <button onClick={() => setConditionalPanelOpen(false)} title="Закрыть" className="ml-auto rounded p-1 text-ws-faint hover:bg-ws-hover hover:text-ws"><X size={13} /></button>
          </div>
          {activeMeta.conditional.length > 0 && (
            <div className="flex flex-wrap items-center gap-2 border-t border-ws/50 pt-1">
              <span className="text-[10px] text-ws-faint">Правила листа:</span>
              {activeMeta.conditional.map((rule) => {
                const labels: Record<string, string> = {
                  greater: '>', less: '<', equal: '=', notEqual: '≠', between: 'между',
                  contains: 'содержит', empty: 'пусто', notEmpty: 'не пусто', scale: 'шкала',
                };
                return (
                  <span key={rule.id} className="flex items-center gap-1 rounded bg-ws-surface px-1.5 py-0.5 text-[10px] text-ws-muted">
                    {rule.range} · {labels[rule.operator]} {rule.value || ''}
                    <button onClick={() => deleteConditionalRule(rule.id)} title="Удалить правило" className="text-ws-faint hover:text-ws-danger"><X size={10} /></button>
                  </span>
                );
              })}
            </div>
          )}
        </div>
      )}

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
      <div className="relative min-h-0 flex-1">
      <div
        ref={gridRef}
        tabIndex={0}
        onFocus={() => { gridFocusedRef.current = true; }}
        onBlur={() => { gridFocusedRef.current = false; }}
        onKeyDown={handleGridKeyDown}
        onScroll={(e) => {
          if (isVirtual) setScrollTop(e.currentTarget.scrollTop);
        }}
        className="h-full overflow-auto bg-ws-editor focus:outline-none"
      >
        <table className="border-collapse select-text" style={{ tableLayout: 'fixed', zoom }}>
          <colgroup>
            <col style={{ width: 48 }} />
            {(grid[0] ?? []).map((_, c) => (
              <col key={c} style={{ width: activeMeta.colWidths[c] ?? SheetFmt.DEFAULT_COL_WIDTH }} />
            ))}
          </colgroup>
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
                    'relative h-7 cursor-pointer border border-ws px-2 text-center font-medium transition-colors hover:bg-ws-hover',
                    selected.col === c && 'bg-emerald-500/15 text-emerald-500'
                  )}
                >
                  {columnName(c)}
                  <span
                    onMouseDown={(event) => startColumnResize(event, c)}
                    title="Потяните, чтобы изменить ширину столбца"
                    className="absolute top-0 right-0 z-10 h-full w-1.5 cursor-col-resize bg-transparent hover:bg-emerald-500/40"
                  />
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
            {visibleRows.filter(({ r }) => rowMatchesFilter(r)).map(({ row, r }) => {
              const rowHeight = activeMeta.rowHeights[r] ?? SheetFmt.DEFAULT_ROW_HEIGHT;
              const rowFrozen = r < activeMeta.freezeRows;
              return (
              <tr key={r} style={{ height: rowHeight }}>
                <td
                  onClick={() => setSelected({ row: r, col: 0 })}
                  style={
                    rowFrozen
                      ? { position: 'sticky', top: 28 + rowOffsets[r], zIndex: 14 }
                      : undefined
                  }
                  className={cn(
                    'sticky left-0 z-10 w-12 cursor-pointer border border-ws bg-ws-surface text-center font-mono text-[11px] transition-colors hover:bg-ws-hover',
                    selected.row === r ? 'bg-emerald-500/15 font-bold text-emerald-500' : 'text-ws-faint'
                  )}
                >
                  {r + 1}
                  <span
                    onMouseDown={(event) => startRowResize(event, r)}
                    title="Потяните, чтобы изменить высоту строки"
                    className="absolute inset-x-0 -bottom-2 z-20 h-1.5 cursor-row-resize hover:bg-emerald-500/40"
                  />
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
                  const address = SheetFmt.cellAddress(r, c);
                  if (mergeInfo.covered.has(address)) return null;
                  const merge = mergeInfo.anchors.get(address);
                  const cellStyle = activeMeta.cells[address];
                  const shown = SheetFmt.formatCellValue(display, cellStyle?.numberFormat);
                  const comment = activeMeta.comments[address];
                  const validation = activeMeta.validations[address];
                  const conditionalCss: Record<string, string | number> = {};
                  for (const rule of activeMeta.conditional) {
                    if (!SheetFmt.ruleContainsCell(rule, r, c)) continue;
                    if (rule.operator === 'scale') {
                      const bounds = conditionalScaleBounds.get(rule.id);
                      const num = numericFromText(display);
                      if (!bounds || num === null) continue;
                      conditionalCss.backgroundColor = SheetFmt.scaleColor(
                        (num - bounds.min) / Math.max(0.0001, bounds.max - bounds.min)
                      );
                    } else if (SheetFmt.ruleMatches(rule, display)) {
                      Object.assign(conditionalCss, SheetFmt.conditionalCss(rule));
                    }
                  }
                  const frozenRow = r < activeMeta.freezeRows;
                  const frozenCol = c < activeMeta.freezeCols;
                  const stickyStyle: React.CSSProperties = {};
                  if (frozenRow) {
                    stickyStyle.position = 'sticky';
                    stickyStyle.top = 28 + rowOffsets[r];
                    stickyStyle.zIndex = 8;
                  }
                  if (frozenCol) {
                    stickyStyle.position = 'sticky';
                    stickyStyle.left = 48 + colOffsets[c];
                    stickyStyle.zIndex = frozenRow ? 9 : 8;
                  }

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
                      rowSpan={merge?.rowSpan}
                      colSpan={merge?.colSpan}
                      style={{ ...SheetFmt.styleToCss(cellStyle), ...conditionalCss, ...stickyStyle }}
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
                      title={
                        [
                          isFormula ? `${raw}  →  ${display}` : display,
                          comment ? `Примечание: ${comment.text}` : '',
                        ]
                          .filter(Boolean)
                          .join('\n')
                      }
                      className={cn(
                        'relative cursor-cell overflow-hidden border border-ws px-2 text-xs',
                        !cellStyle?.wrap && 'truncate',
                        isSelected && 'bg-emerald-500/15 outline outline-1 outline-emerald-500/40',
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
                      {comment && (
                        <span className="pointer-events-none absolute top-0 right-0 h-0 w-0 border-t-[7px] border-l-[7px] border-t-red-500 border-l-transparent" />
                      )}
                      {validation && !isEditing && (
                        <span className="pointer-events-none absolute right-0.5 bottom-0 text-[8px] text-ws-faint">▾</span>
                      )}
                      {isEditing ? (
                        validation ? (
                          <select
                            autoFocus
                            value={validation.values.includes(editorValue) ? editorValue : ''}
                            onChange={(e) => {
                              applyChange(r, c, e.target.value);
                              setEditingCell(null);
                              formula.endEdit();
                              gridRef.current?.focus();
                            }}
                            onBlur={() => setEditingCell(null)}
                            className="h-full w-full border border-emerald-500 bg-ws-surface px-1 text-xs text-ws outline-none"
                          >
                            <option value="">—</option>
                            {validation.values.map((value) => (
                              <option key={value} value={value}>{value}</option>
                            ))}
                          </select>
                        ) : (
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
                        )
                      ) : (
                        <span
                          className={cn(
                            'block',
                            !cellStyle?.wrap && 'truncate',
                            isFormula && 'text-emerald-500',
                            isError(cell as never) && 'text-ws-danger'
                          )}
                        >
                          {shown}
                        </span>
                      )}
                    </td>
                  );
                })}
              </tr>
              );
            })}
            {bottomSpacerHeight > 0 && (
              <tr style={{ height: `${bottomSpacerHeight}px` }}>
                <td colSpan={(grid[0]?.length ?? 0) + 1} style={{ padding: 0, border: 'none' }} />
              </tr>
            )}
          </tbody>
        </table>
      </div>

        {/* Диаграммы листа */}
        {activeMeta.charts.map((chart) => {
          const data = buildChartData(chart.range);
          return (
            <div
              key={chart.id}
              className="absolute z-30 flex flex-col overflow-hidden rounded-lg border border-ws bg-ws-surface shadow-xl"
              style={{ left: chart.x, top: chart.y, width: chart.width, height: chart.height }}
            >
              <div
                onMouseDown={(event) => startChartMove(event, chart, 'move')}
                className="flex cursor-move items-center justify-between gap-2 border-b border-ws bg-ws-elevated px-2 py-1 text-[11px] select-none"
              >
                <span className="truncate font-medium text-ws">{chart.title}</span>
                <button
                  onMouseDown={(event) => event.stopPropagation()}
                  onClick={() => removeChart(chart.id)}
                  title="Удалить диаграмму"
                  className="rounded p-0.5 text-ws-faint hover:bg-ws-hover hover:text-ws-danger"
                >
                  <X size={12} />
                </button>
              </div>
              <div className="min-h-0 flex-1 p-1 text-ws">
                <SheetChart spec={chart} labels={data.labels} series={data.series} />
              </div>
              <span
                onMouseDown={(event) => startChartMove(event, chart, 'resize')}
                title="Изменить размер"
                className="absolute right-0 bottom-0 h-3 w-3 cursor-nwse-resize bg-emerald-500/60"
              />
            </div>
          );
        })}
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
              onDoubleClick={async () => {
                const next = await askText('Новое имя листа:', name);
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
          onClick={() => addSheet()}
          title="Добавить лист"
          className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"
        >
          <Plus size={13} />
        </button>
        {sheetNames.length > 1 && (
          <button
            onClick={() => removeSheet(activeSheet)}
            title="Удалить активный лист"
            className="rounded p-1 text-ws-faint transition-colors hover:bg-red-500/10 hover:text-ws-danger"
          >
            <Trash2 size={13} />
          </button>
        )}

        <div className="ml-auto flex items-center gap-3 text-[10px] text-ws-faint">
          <span>
            {stats.rows} × {stats.cols}
          </span>
          <span>Заполнено: {stats.filled}</span>
          <span>Масштаб: {Math.round(zoom * 100)}%</span>
          {formulaCount > 0 && <span className="text-emerald-500">Формул: {formulaCount}</span>}
          {tab.isDirty && <span className="text-ws-warning">● Изменено</span>}
        </div>
      </div>

      {/* Диалог диаграммы */}
      {chartDialogOpen && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/55 p-4 backdrop-blur-[2px]"
          onMouseDown={(e) => { if (e.target === e.currentTarget) setChartDialogOpen(false); }}
        >
          <div className="animate-scale-in w-full max-w-md overflow-hidden rounded-xl border border-ws bg-ws-surface shadow-ws">
            <div className="flex items-center justify-between border-b border-ws px-4 py-3">
              <div className="flex items-center gap-2 text-sm font-semibold text-emerald-500">
                <BarChart3 size={15} />
                Диаграмма
              </div>
              <button onClick={() => setChartDialogOpen(false)} className="text-ws-faint hover:text-ws"><X size={14} /></button>
            </div>
            <div className="space-y-3 p-4 text-xs">
              <p className="text-[11px] text-ws-muted">
                Диапазон: <strong className="font-mono text-ws">{selectionLabel}</strong>. Первая
                строка — заголовки, первый столбец — подписи категорий.
              </p>
              <label className="block space-y-1">
                <span className="text-ws-muted">Тип</span>
                <select
                  value={chartType}
                  onChange={(e) => setChartType(e.target.value as SheetFmt.ChartType)}
                  className="w-full rounded border border-ws bg-ws-editor px-2 py-1 text-xs text-ws"
                >
                  <option value="bar">Столбчатая</option>
                  <option value="line">Линейная</option>
                  <option value="pie">Круговая</option>
                </select>
              </label>
              <label className="block space-y-1">
                <span className="text-ws-muted">Название</span>
                <input
                  value={chartTitle}
                  onChange={(e) => setChartTitle(e.target.value)}
                  className="w-full rounded border border-ws bg-ws-editor px-2 py-1 text-xs text-ws focus:border-emerald-500 focus:outline-none"
                />
              </label>
            </div>
            <div className="flex justify-end gap-2 border-t border-ws px-4 py-3">
              <Button variant="secondary" onClick={() => setChartDialogOpen(false)}>Отмена</Button>
              <Button variant="primary" onClick={createChart}>Создать</Button>
            </div>
          </div>
        </div>
      )}

      {/* Диалог сводной таблицы */}
      {pivotOpen && (() => {
        const pivotRange = SheetFmt.detectUsedRange(grid, selected.row, selected.col);
        const headerOptions: { index: number; label: string }[] = [];
        for (let col = pivotRange.startCol; col <= pivotRange.endCol; col++) {
          headerOptions.push({
            index: col,
            label: `${columnName(col)} — ${displayCell(pivotRange.startRow, col) || 'без названия'}`,
          });
        }
        return (
          <div
            className="fixed inset-0 z-50 flex items-center justify-center bg-black/55 p-4 backdrop-blur-[2px]"
            onMouseDown={(e) => { if (e.target === e.currentTarget) setPivotOpen(false); }}
          >
            <div className="animate-scale-in w-full max-w-md overflow-hidden rounded-xl border border-ws bg-ws-surface shadow-ws">
              <div className="flex items-center justify-between border-b border-ws px-4 py-3">
                <div className="flex items-center gap-2 text-sm font-semibold text-emerald-500">
                  <Table2 size={15} />
                  Сводная таблица
                </div>
                <button onClick={() => setPivotOpen(false)} className="text-ws-faint hover:text-ws"><X size={14} /></button>
              </div>
              <div className="space-y-3 p-4 text-xs">
                <p className="text-[11px] leading-relaxed text-ws-muted">
                  Источник: <strong className="font-mono text-ws">{SheetFmt.rangeToAddress(pivotRange)}</strong>.
                  Результат будет создан на новом листе.
                </p>
                <label className="block space-y-1">
                  <span className="text-ws-muted">Строки (поле группировки)</span>
                  <select
                    value={pivotRowCol}
                    onChange={(e) => setPivotRowCol(Number(e.target.value))}
                    className="w-full rounded border border-ws bg-ws-editor px-2 py-1 text-xs text-ws"
                  >
                    {headerOptions.map((option) => (
                      <option key={option.index} value={option.index}>{option.label}</option>
                    ))}
                  </select>
                </label>
                <label className="block space-y-1">
                  <span className="text-ws-muted">Значения</span>
                  <select
                    value={pivotValueCol}
                    onChange={(e) => setPivotValueCol(Number(e.target.value))}
                    className="w-full rounded border border-ws bg-ws-editor px-2 py-1 text-xs text-ws"
                  >
                    {headerOptions.map((option) => (
                      <option key={option.index} value={option.index}>{option.label}</option>
                    ))}
                  </select>
                </label>
                <label className="block space-y-1">
                  <span className="text-ws-muted">Агрегация</span>
                  <select
                    value={pivotAgg}
                    onChange={(e) => setPivotAgg(e.target.value as typeof pivotAgg)}
                    className="w-full rounded border border-ws bg-ws-editor px-2 py-1 text-xs text-ws"
                  >
                    <option value="sum">Сумма</option>
                    <option value="count">Количество</option>
                    <option value="average">Среднее</option>
                    <option value="min">Минимум</option>
                    <option value="max">Максимум</option>
                  </select>
                </label>
              </div>
              <div className="flex justify-end gap-2 border-t border-ws px-4 py-3">
                <Button variant="secondary" onClick={() => setPivotOpen(false)}>Отмена</Button>
                <Button variant="primary" onClick={createPivot}>Создать лист</Button>
              </div>
            </div>
          </div>
        );
      })()}

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

      {promptState && (
        <Modal
          open
          onClose={() => finishPrompt(null)}
          title={promptState.title}
          width="max-w-md"
          footer={
            <>
              <Button variant="ghost" onClick={() => finishPrompt(null)}>
                Отмена
              </Button>
              <Button variant="primary" onClick={() => finishPrompt(promptState.value)}>
                OK
              </Button>
            </>
          }
        >
          <Input
            autoFocus
            value={promptState.value}
            onChange={(event) =>
              setPromptState((state) => (state ? { ...state, value: event.target.value } : state))
            }
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                finishPrompt(promptState.value);
              }
            }}
          />
        </Modal>
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

