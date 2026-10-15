/**
 * Хук «формульного режима» таблицы: автодополнение имён функций и выбор
 * диапазона мышью при вводе формулы (как в Excel).
 *
 * Когда пользователь начинает ввод с '=' или редактирует формулу, клик по
 * другой ячейке не перемещает выделение, а подставляет её адрес в формулу —
 * «=СУММ(» + выделение мышью B2:B10 даёт «=СУММ(B2:B10))».
 */
import { useCallback, useMemo, useRef, useState } from 'react';
import { columnName, FUNCTION_NAMES, RU_FUNCTION_ALIASES } from '../../lib/formulaEngine';

/** Точка вставки ссылки внутри формулы. */
export interface FormulaAnchor {
  /** Позиция курсора в тексте формулы. */
  caret: number;
  /** Длина уже вставленной ссылки, которую надо заменить при протяжке. */
  replaceLen: number;
}

export interface CellCoord {
  row: number;
  col: number;
}

/** Диапазон ячеек (включительно). */
export interface CellRange {
  start: CellCoord;
  end: CellCoord;
}

export function rangeToA1(range: CellRange): string {
  const startRef = `${columnName(range.start.col)}${range.start.row + 1}`;
  const endRef = `${columnName(range.end.col)}${range.end.row + 1}`;
  return startRef === endRef ? startRef : `${startRef}:${endRef}`;
}

export function normalizeRange(a: CellCoord, b: CellCoord): CellRange {
  return {
    start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
}

/**
 * Сдвигает относительные ссылки формулы при копировании в другую ячейку.
 * Абсолютные части ($B$2) и ссылки на другие листы (Лист2!B2) не меняются,
 * текст в кавычках игнорируется.
 */
export function shiftFormula(formula: string, rowDelta: number, colDelta: number): string {
  if (!formula.startsWith('=')) return formula;
  if (rowDelta === 0 && colDelta === 0) return formula;

  /** Позиции символов внутри строковых литералов — их менять нельзя. */
  const inString = (text: string, index: number): boolean => {
    let count = 0;
    for (let i = 0; i < index && i < text.length; i++) {
      if (text[i] === '"') count++;
    }
    return count % 2 === 1;
  };

  /**
   * Проход 1: ссылки с именем листа ('Лист 2'!B2 или Лист2!B2). Имя листа
   * сохраняется, сдвигается только адрес. Результат временно заменяется
   * плейсхолдером, чтобы проход 2 не сдвинул ту же ссылку повторно.
   */
  const saved: string[] = [];
  const PH = (i: number) => `\u0001${i}\u0001`;

  let result = formula.replace(
    /('[^']*'|[A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_.]*)!(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g,
    (
      match: string,
      sheetName: string,
      absCol: string,
      letters: string,
      absRow: string,
      digits: string,
      offset: number,
      full: string
    ) => {
      if (inString(full, offset)) return match;
      const shifted = shiftRef(letters, digits, absCol, absRow, rowDelta, colDelta);
      saved.push(`${sheetName}!${shifted}`);
      return PH(saved.length - 1);
    }
  );

  // Проход 2: локальные ссылки (квалифицированные уже заменены плейсхолдерами)
  result = result.replace(
    /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g,
    (match: string, absCol: string, letters: string, absRow: string, digits: string, offset: number, full: string) => {
      if (inString(full, offset)) return match;
      return shiftRef(letters, digits, absCol, absRow, rowDelta, colDelta);
    }
  );

  // Возвращаем сохранённые ссылки на место
  return result.replace(/\u0001(\d+)\u0001/g, (_m, i: string) => saved[Number(i)] ?? '');
}

/** Сдвигает одну ссылку с учётом абсолютных частей. */
function shiftRef(
  letters: string,
  digits: string,
  absCol: string,
  absRow: string,
  rowDelta: number,
  colDelta: number
): string {
  let col = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i++) col = col * 26 + (upper.charCodeAt(i) - 64);
  col -= 1;

  let row = parseInt(digits, 10) - 1;

  if (!absCol) col += colDelta;
  if (!absRow) row += rowDelta;

  if (col < 0 || row < 0) return '#REF!';
  return `${absCol}${columnName(col)}${absRow}${row + 1}`;
}

/** Определяет, находится ли каретка внутри строкового литерала. */
function insideString(text: string, index: number): boolean {
  let inString = false;
  for (let i = 0; i < index && i < text.length; i++) {
    if (text[i] === '"') {
      if (inString && text[i + 1] === '"') {
        i++;
        continue;
      }
      inString = !inString;
    }
  }
  return inString;
}

/**
 * Находит незакрытую функцию перед кареткой — для подсказки аргументов.
 * Возвращает имя функции и позицию её открывающей скобки.
 */
export function findOpenFunction(
  text: string,
  caret: number
): { name: string; parenIndex: number } | null {
  let depth = 0;
  for (let i = caret - 1; i >= 0; i--) {
    const ch = text[i];
    if (insideString(text, i)) continue;
    if (ch === ')') depth++;
    else if (ch === '(') {
      if (depth === 0) {
        // Имя перед скобкой
        const before = text.slice(0, i);
        const m = before.match(/([A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9_.]*)\s*$/);
        return m ? { name: m[1], parenIndex: i } : null;
      }
      depth--;
    }
  }
  return null;
}

/** Извлекает частичное слово перед кареткой — для автодополнения имён. */
export function partialWordBefore(text: string, caret: number): { word: string; from: number } {
  const before = text.slice(0, caret);
  // Слово не должно содержать ссылку на ячейку (буквы + цифры)
  const m = before.match(/([A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9_.]*)$/);
  if (!m) return { word: '', from: caret };

  const word = m[1];
  // Если слово целиком похоже на адрес ячейки (B2, AA10) — это не функция
  if (/^\$?[A-Za-z]{1,3}\$?\d+$/.test(word)) return { word: '', from: caret };

  return { word, from: caret - word.length };
}

export interface FunctionSuggestion {
  /** Каноническое имя, которое будет вставлено. */
  name: string;
  /** Отображаемое имя (может быть русским). */
  label: string;
  /** Русское название, если есть. */
  ru?: string;
}

/**
 * Подбирает функции по введённому фрагменту (английские и русские имена).
 * Сначала идут точные и короткие совпадения: «SU» → SUM, а не SUBSTITUTE.
 */
export function suggestFunctions(partial: string, limit = 8): FunctionSuggestion[] {
  const q = partial.trim().toUpperCase();
  if (!q) return [];

  // Обратный индекс: каноническое имя -> русское название
  const ruByCanonical = new Map<string, string>();
  for (const [ru, canonical] of Object.entries(RU_FUNCTION_ALIASES)) {
    if (!ruByCanonical.has(canonical)) ruByCanonical.set(canonical, ru);
  }

  interface Ranked extends FunctionSuggestion {
    score: number;
    length: number;
  }

  const ranked: Ranked[] = [];

  for (const name of FUNCTION_NAMES) {
    const ru = ruByCanonical.get(name);
    if (name.startsWith(q)) {
      ranked.push({ name, label: name, ru, score: 0, length: name.length });
    }
    if (ru && ru.startsWith(q)) {
      // Русское имя — тоже полноценное начало совпадения
      ranked.push({ name, label: ru, ru, score: 0, length: ru.length });
    }
  }

  // Если в начале ничего не нашлось — ищем вхождение внутри имени
  if (ranked.length === 0 && q.length >= 2) {
    for (const name of FUNCTION_NAMES) {
      if (name.includes(q)) {
        ranked.push({
          name,
          label: name,
          ru: ruByCanonical.get(name),
          score: 1,
          length: name.length,
        });
      }
    }
  }

  // Ранжирование: качество совпадения, затем точное равенство, затем краткость
  ranked.sort((a, b) => {
    if (a.score !== b.score) return a.score - b.score;
    const aExact = a.label.toUpperCase() === q ? 0 : 1;
    const bExact = b.label.toUpperCase() === q ? 0 : 1;
    if (aExact !== bExact) return aExact - bExact;
    // Латинские имена при равной длине идут первыми — это канонический ввод
    if (a.length !== b.length) return a.length - b.length;
    return a.name.localeCompare(b.name);
  });

  // Убираем дубли (одна функция может совпасть и по EN, и по RU)
  const out: FunctionSuggestion[] = [];
  const seen = new Set<string>();
  for (const r of ranked) {
    const key = `${r.name}|${r.label}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ name: r.name, label: r.label, ru: r.ru });
    if (out.length >= limit) break;
  }

  return out;
}

/**
 * Состояние выбора диапазона мышью. Живёт вне React-рендера в ref,
 * потому что обновляется на каждом mousemove.
 */
export interface SelectionDragState {
  /** Якорь выделения — первая ячейка. */
  anchor: CellCoord | null;
  /** Текущий конец выделения. */
  focus: CellCoord | null;
  /** Активна ли протяжка мышью. */
  dragging: boolean;
}

/** Хук управления формульным режимом. */
export function useFormulaMode() {
  const [draft, setDraft] = useState('');
  const [caret, setCaret] = useState(0);
  const [anchor, setAnchor] = useState<FormulaAnchor | null>(null);
  /**
   * Идёт ли прямо сейчас ввод формулы. В отличие от проверки draft на '=',
   * этот флаг сбрасывается, когда пользователь зафиксировал формулу или ушёл
   * в другую ячейку. Без него строка формул оставалась в режиме правки и
   * «переезжала» вместе с выделением на следующую ячейку.
   */
  const [editing, setEditing] = useState(false);
  const [drag, setDrag] = useState<SelectionDragState>({
    anchor: null,
    focus: null,
    dragging: false,
  });
  const dragRef = useRef<SelectionDragState>({ anchor: null, focus: null, dragging: false });
  /**
   * Последняя вставленная ссылка. Держим её в ref, а не только в состоянии:
   * mousedown и mouseup приходят в одном тике, и обработчик mouseup видел бы
   * устаревшее значение из замыкания — ссылка вставлялась бы дважды
   * («=SUM(A2A2» вместо «=SUM(A2»).
   */
  const anchorRef = useRef<FormulaAnchor | null>(null);

  /** Идёт ли сейчас ввод формулы (для перехвата кликов по ячейкам). */
  const isFormulaMode = editing && draft.startsWith('=');

  const partial = useMemo(() => partialWordBefore(draft, caret), [draft, caret]);
  const suggestions = useMemo(
    () => (isFormulaMode ? suggestFunctions(partial.word) : []),
    [isFormulaMode, partial.word]
  );
  const openFunction = useMemo(
    () => (isFormulaMode ? findOpenFunction(draft, caret) : null),
    [isFormulaMode, draft, caret]
  );

  /** Синхронизирует anchor в состоянии и в ref. */
  const updateAnchor = useCallback((next: FormulaAnchor | null) => {
    anchorRef.current = next;
    setAnchor(next);
  }, []);

  /**
   * Вставляет ссылку в формулу в позиции каретки.
   * Если ссылка уже была вставлена и её тянут дальше — заменяет её.
   */
  const insertReference = useCallback(
    (ref: string, options?: { replacePrevious?: boolean }) => {
      setDraft((prev) => {
        const current = anchorRef.current;
        // Замена — только для протяжки: правим ранее вставленную ссылку на месте.
        // Для новой ссылки берём ТЕКУЩУЮ каретку (её обновляют ввод и клики), а не
        // старую позицию из anchorRef: иначе «=A8+» + клик A9 давал «=A9A8+».
        const replacing = Boolean(options?.replacePrevious && current);
        const rawPos = replacing ? current!.caret : caret;
        // Каретка не может выходить за пределы текста (страховка от рассинхрона).
        const pos = Math.max(0, Math.min(rawPos, prev.length));
        const replaceLen = replacing ? current!.replaceLen : 0;
        const next = prev.slice(0, pos) + ref + prev.slice(pos + replaceLen);
        const nextCaret = pos + ref.length;
        setCaret(nextCaret);
        updateAnchor({ caret: pos, replaceLen: ref.length });
        return next;
      });
    },
    [caret, updateAnchor]
  );

  /** Заменяет частично введённое имя функции на выбранное. */
  const applySuggestion = useCallback(
    (suggestion: FunctionSuggestion) => {
      setDraft((prev) => {
        const from = partial.from;
        const to = caret;
        const needsParen = !prev.slice(to).startsWith('(');
        const inserted = `${suggestion.name}${needsParen ? '(' : ''}`;
        const next = prev.slice(0, from) + inserted + prev.slice(to);
        const nextCaret = from + inserted.length;
        setCaret(nextCaret);
        updateAnchor(null);
        return next;
      });
    },
    [partial.from, caret, updateAnchor]
  );

  const reset = useCallback(() => {
    setDraft('');
    setCaret(0);
    updateAnchor(null);
    setEditing(false);
    dragRef.current = { anchor: null, focus: null, dragging: false };
    setDrag({ anchor: null, focus: null, dragging: false });
  }, [updateAnchor]);

  /** Начало правки формулы: клики по ячейкам начинают подставлять ссылки. */
  const beginEdit = useCallback(() => {
    setEditing(true);
  }, []);

  /** Конец правки формулы: строка формул снова следует за выделением. */
  const endEdit = useCallback(() => {
    setEditing(false);
    updateAnchor(null);
    dragRef.current = { anchor: null, focus: null, dragging: false };
    setDrag({ anchor: null, focus: null, dragging: false });
  }, [updateAnchor]);

  return {
    draft,
    setDraft,
    caret,
    setCaret,
    anchor,
    // Наружу отдаём синхронизированный сеттер: прямая запись в состояние
    // разошлась бы с anchorRef, и вставка ссылок заменила бы не тот фрагмент.
    setAnchor: updateAnchor,
    editing,
    beginEdit,
    endEdit,
    isFormulaMode,
    suggestions,
    openFunction,
    partial,
    insertReference,
    applySuggestion,
    reset,
    drag,
    setDrag,
    dragRef,
  };
}
