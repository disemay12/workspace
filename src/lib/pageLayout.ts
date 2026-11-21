/**
 * Постраничная разметка полотна Word: лист за листом, как в Word.
 *
 * Содержимое остаётся ОДНИМ потоком (contenteditable не разрезается), а в
 * точку разрыва вставляется блок-маркер `[data-page-gap]` нулевой ширины с
 * высотой «нижнее поле + промежуток + верхнее поле»: текст до него заканчивает
 * лист, текст после продолжается ровно с верха полосы набора следующего.
 * Фон листов рисуется отдельным слоем по вычисленным координатам.
 *
 * Такой маркер не меняет абзацы, блоки и идентификаторы (data-bid,
 * data-docx-id, смещения выделения), поэтому AI, сохранение, оглавление и
 * поиск продолжают работать: перед сериализацией маркеры вырезаются
 * (`stripPagination`).
 *
 * Таблицы режутся по строкам — между tr DOM-маркер не вставить: остаток строк
 * переносится в клон `<table>` с `[data-page-cont]`, шапка (`thead`)
 * повторяется, а `stripPagination` склеивает клоны обратно.
 */
import { PAGE_GAP_PX, type PageGeometry } from './units';

export const PAGE_GAP_ATTR = 'data-page-gap';
export const TABLE_CONT_ATTR = 'data-page-cont';
export const GAP_CLASS = 'ws-page-gap';

/** Лист в координатах полотна (без масштаба). */
export interface PageBox {
  index: number;
  topPx: number;
  heightPx: number;
  /** Полоса набора: где начинается и заканчивается текст листа. */
  contentTopPx: number;
  contentBottomPx: number;
}

export interface PaginateOptions {
  geometry: PageGeometry;
  /** Промежуток между листами; по умолчанию как у Word. */
  gapPx?: number;
}

export interface PaginateResult {
  pages: PageBox[];
  /** Принудительные разрывы (w:pageBreakBefore, разрыв в тексте). */
  forcedBreaks: number;
  warnings: string[];
  ms: number;
  /** Высота полотна вместе с промежутками. */
  totalHeightPx: number;
  /** Диагностика поставленных разрывов: цель и фактическая координата. */
  breaks: Array<{ target: number; actual: number; height: number }>;
}

interface FlowPosition {
  node: Node;
  offset: number;
}

interface LineInfo {
  /** Позиция начала строки. */
  start: FlowPosition;
  /** Нижняя граница строки в координатах полотна. */
  bottom: number;
}

const EPSILON = 0.6;
const MAX_PAGES = 4000;

/* ------------------------------------------------------------------ замеры */

/**
 * Масштаб полотна: CSS zoom меняет координаты getBoundingClientRect.
 * offsetWidth округлён до целого пикселя, поэтому по нему масштаб считать
 * нельзя — на десяти листах ошибка накапливалась в несколько пикселей.
 */
function layoutScale(root: HTMLElement): number {
  const probe = document.createElement('div');
  probe.style.cssText =
    'position:absolute;left:-10000px;top:0;width:100px;height:0;visibility:hidden;pointer-events:none;';
  root.append(probe);
  const width = probe.getBoundingClientRect().width;
  probe.remove();
  if (!(width > 0)) return 1;
  const scale = width / 100;
  return Number.isFinite(scale) && scale > 0.1 ? scale : 1;
}

interface Meter {
  /** Координата полотна по координате окна. */
  at(viewportY: number): number;
  top(element: Element): number;
  bottom(element: Element): number;
  refresh(): void;
}

function createMeter(root: HTMLElement): Meter {
  let origin = 0;
  const scale = layoutScale(root);
  const refresh = () => {
    origin = root.getBoundingClientRect().top;
  };
  refresh();
  return {
    at: (viewportY) => (viewportY - origin) / scale,
    top: (element) => (element.getBoundingClientRect().top - origin) / scale,
    bottom: (element) => (element.getBoundingClientRect().bottom - origin) / scale,
    refresh,
  };
}

/* ------------------------------------------------------------------ разметка */

/** Убирает маркеры разрывов и склеивает разрезанные таблицы. */
export function unpaginateDom(root: HTMLElement): void {
  // Сначала маркеры: между исходной таблицей и её клоном стоит именно маркер,
  // и при склейке он оказывался «предыдущим соседом» вместо таблицы.
  for (const gap of [...root.querySelectorAll(`[${PAGE_GAP_ATTR}]`)]) gap.remove();
  for (const table of [...root.querySelectorAll<HTMLTableElement>(`table[${TABLE_CONT_ATTR}]`)]) {
    mergeContinuationTable(table);
  }
}

/** Предыдущая таблица в том же контейнере (между ними могли быть маркеры). */
function previousTable(table: HTMLTableElement): HTMLTableElement | null {
  let node: Element | null = table.previousElementSibling;
  while (node) {
    if (node.tagName.toLowerCase() === 'table') return node as HTMLTableElement;
    node = node.previousElementSibling;
  }
  return null;
}

function mergeContinuationTable(table: HTMLTableElement): void {
  const previous = previousTable(table);
  if (!previous) {
    table.removeAttribute(TABLE_CONT_ATTR);
    return;
  }
  const target = lastSection(previous);
  for (const section of [...table.querySelectorAll('tbody')]) {
    for (const row of [...section.children]) target.append(row);
    section.remove();
  }
  table.remove();
}

function lastSection(table: HTMLTableElement): HTMLElement {
  const bodies = [...table.querySelectorAll<HTMLElement>(':scope > tbody')];
  const body = bodies[bodies.length - 1];
  if (body) return body;
  const created = document.createElement('tbody');
  for (const child of [...table.children]) {
    if (child.tagName.toLowerCase() === 'tr') created.append(child);
  }
  table.append(created);
  return created;
}

/** Пустой служебный элемент разрыва в сериализованном HTML. */
const GAP_TAG_RE = /<(span|div)\b[^>]*\sdata-page-gap="1"[^>]*>\s*<\/\1>/gi;

/** HTML без служебных элементов разметки: для сохранения, AI и истории. */
export function stripPagination(html: string): string {
  if (!html.includes(PAGE_GAP_ATTR) && !html.includes(TABLE_CONT_ATTR)) return html;
  // Маркеры вырезаются по строке: разбор HTML с блочным элементом внутри <p>
  // закрывает абзац, и документ делился на два (видно по лишнему </p>).
  const withoutGaps = html.replace(GAP_TAG_RE, '');
  if (!withoutGaps.includes(TABLE_CONT_ATTR)) return withoutGaps;
  const holder = document.createElement('div');
  holder.innerHTML = withoutGaps;
  unpaginateDom(holder);
  return holder.innerHTML;
}

/* ---------------------------------------------------------------- пагинация */

/**
 * Раскладывает поток по листам. Идемпотентно: сначала снимается прошлая
 * разметка, поэтому повторный вызов на изменённом тексте безопасен.
 */
export function paginate(root: HTMLElement, options: PaginateOptions): PaginateResult {
  const started = performance.now();
  const warnings: string[] = [];
  unpaginateDom(root);

  const { geometry } = options;
  const gapPx = options.gapPx ?? PAGE_GAP_PX;
  const meter = createMeter(root);
  const pages: PageBox[] = [];
  let forcedBreaks = 0;
  const breaks: Array<{ target: number; actual: number; height: number }> = [];

  let pageTop = 0;
  let element: HTMLElement | null = firstFlowChild(root);
  /** Продолжение разрезанного блока: с какого текста начинать. */
  let resume: { block: HTMLElement; from: FlowPosition } | null = null;
  /** Блок, который уже переносили целиком: защита от бесконечного цикла. */
  let moved: HTMLElement | null = null;
  /**
   * Уже отработанные принудительные разрывы. Без этого признака маркер
   * `w:br type="page"` находился снова и снова: маркер разрыва остаётся в
   * абзаце, и каждая следующая страница ставила ещё один (4000 листов).
   */
  const handledBreaks = new Set<Element>();

  const contentStart = () => pageTop + geometry.bodyTopPx;
  const contentEnd = () => contentStart() + geometry.contentHeightPx;
  const nextTop = () => pageTop + geometry.heightPx + gapPx;
  const nextContentStart = () => nextTop() + geometry.bodyTopPx;

  /**
   * Ставит маркер разрыва и переводит страницу.
   * attach вставляет маркер, measure сообщает фактическую координату
   * продолжения — по ней высота маркера доводится до точной.
   */
  const breakHere = (
    attach: (gap: HTMLElement) => void,
    measure: (gap: HTMLElement) => number
  ): HTMLElement => {
    const gap = makeGap();
    attach(gap);
    const target = nextContentStart();
    meter.refresh();
    let height = Math.max(0, target - meter.top(gap));
    gap.style.height = `${height}px`;
    let actual = Number.NaN;
    for (let attempt = 0; attempt < 4; attempt += 1) {
      actual = measure(gap);
      if (!Number.isFinite(actual)) break;
      const delta = target - actual;
      if (Math.abs(delta) < 0.35) break;
      height = Math.max(0, height + delta);
      gap.style.height = `${height}px`;
      meter.refresh();
    }
    breaks.push({ target, actual, height });
    pageTop = nextTop();
    return gap;
  };

  let guard = 0;
  while (guard++ < MAX_PAGES) {
    // Верх листа фиксируется до разрывов: pageTop меняется при вставке маркера.
    const pageBaseTop = pageTop;
    const limit = contentEnd();
    const start = contentStart();
    let broke = false;

    /* 1. Продолжение разрезанного блока. */
    if (resume) {
      const block: HTMLElement = resume.block;
      const lines = collectLines(block, meter, resume.from);
      const cut = pickLineCut(lines, limit);
      if (cut) {
        const gap = breakHere(
          (element) => insertGapInside(element, cut.start),
          (element) => meter.bottom(element)
        );
        // Продолжение — первый текст после маркера: он же первая строка листа.
        const next = textAfter(gap, block);
        resume = next ? { block, from: next } : null;
        if (!resume) element = nextFlowChild(block, root);
        broke = true;
      } else {
        resume = null;
        element = nextFlowChild(block, root);
      }
    }

    /* 2. Блоки по порядку. */
    while (!broke && element) {
      const block = element;

      // Разрыв «с новой страницы» из pPr (w:pageBreakBefore).
      if (
        block.getAttribute('data-page-break') === 'start' &&
        !handledBreaks.has(block) &&
        meter.top(block) > start + 2
      ) {
        breakHere(
          (gap) => block.parentElement?.insertBefore(gap, block),
          () => meter.top(block)
        );
        handledBreaks.add(block);
        forcedBreaks += 1;
        broke = true;
        break;
      }

      // Разрыв внутри блока (w:br type="page"): маркер встаёт перед br.
      const hardBreak = findPageBreak(block, handledBreaks);
      if (hardBreak) {
        breakHere(
          (gap) => hardBreak.parentElement?.insertBefore(gap, hardBreak),
          (gap) => meter.bottom(gap)
        );
        handledBreaks.add(hardBreak);
        forcedBreaks += 1;
        const resumeFrom = textAfter(gapInAfterBreak(hardBreak), block);
        resume = resumeFrom ? { block, from: resumeFrom } : null;
        broke = true;
        break;
      }

      const top = meter.top(block);
      const bottom = meter.bottom(block);

      if (bottom <= limit + EPSILON) {
        element = nextFlowChild(block, root);
        continue;
      }

      if (top >= limit - EPSILON) {
        breakHere(
          (gap) => block.parentElement?.insertBefore(gap, block),
          () => meter.top(block)
        );
        moved = block;
        broke = true;
        break;
      }

      // Блок длиннее листа: перенос целиком не поможет.
      const oversized = bottom - top > geometry.contentHeightPx + EPSILON;

      if (isTable(block)) {
        const rowIndex = fitRows(block, meter, limit);
        if (rowIndex > 0) {
          const continuation = continuationTable(block, rowIndex);
          if (continuation) {
            breakHere(
              (gap) => continuation.insertAdjacentElement('beforebegin', gap),
              () => meter.top(continuation.rows[0] ?? continuation)
            );
            resumedFromContinuation(continuation);
            broke = true;
            break;
          }
        }
        if (!oversized) {
          breakHere(
            (gap) => block.parentElement?.insertBefore(gap, block),
            () => meter.top(block)
          );
          moved = block;
          broke = true;
          break;
        }
        warnings.push('таблица выше листа — выходит за поля');
        element = nextFlowChild(block, root);
        continue;
      }

      const lines = collectLines(block, meter);
      let cut = pickLineCut(lines, limit);
      if (!cut && (oversized || top <= start + 2) && lines.length > 1) {
        // Ни одна строка не влезла, но блок выше листа — режем по первой строке.
        cut = { start: lines[0].start, index: 0 };
        warnings.push('абзац выше листа — разрезан по строке');
      }

      if (cut) {
        const gap = breakHere(
          (element) => insertGapInside(element, cut!.start),
          (element) => meter.bottom(element)
        );
        const next = textAfter(gap, block);
        resume = next ? { block, from: next } : null;
        if (!resume) element = nextFlowChild(block, root);
        broke = true;
        break;
      }

      if (moved === block) {
        // Второй раз переносить тот же блок нельзя — принимаем выход за поля.
        warnings.push('блок не помещается на лист');
        element = nextFlowChild(block, root);
        continue;
      }
      breakHere(
        (gap) => block.parentElement?.insertBefore(gap, block),
        () => meter.top(block)
      );
      moved = block;
      broke = true;
      break;
    }

    pages.push({
      index: pages.length,
      topPx: pageBaseTop,
      heightPx: geometry.heightPx,
      contentTopPx: pageBaseTop + geometry.bodyTopPx,
      contentBottomPx: pageBaseTop + geometry.bodyTopPx + geometry.contentHeightPx,
    });

    if (!broke) break;
  }

  return {
    pages,
    forcedBreaks,
    warnings,
    breaks,
    ms: Math.round(performance.now() - started),
    totalHeightPx: root.offsetHeight,
  };

  /** После разреза таблицы продолжаем с её клона. */
  function resumedFromContinuation(continuation: HTMLTableElement): void {
    resume = null;
    element = continuation;
  }
}

/* -------------------------------------------------------------- вспомогательное */

function isTable(element: Element): element is HTMLTableElement {
  return element.tagName.toLowerCase() === 'table';
}

function firstFlowChild(root: HTMLElement): HTMLElement | null {
  for (const child of [...root.children]) {
    if (child instanceof HTMLElement && !child.hasAttribute(PAGE_GAP_ATTR)) return child;
  }
  return null;
}

function nextFlowChild(element: Element | null, root: HTMLElement): HTMLElement | null {
  if (!element) return null;
  void root;
  let node: Element | null = element.nextElementSibling;
  while (node) {
    if (node instanceof HTMLElement && !node.hasAttribute(PAGE_GAP_ATTR)) return node;
    node = node.nextElementSibling;
  }
  return null;
}

function makeGap(): HTMLElement {
  // Именно span: блокирующий элемент внутри <p> при разборе HTML закрывает
  // абзац. span с display:block в потоке ведёт себя так же, но структуру
  // документа не ломает.
  const gap = document.createElement('span');
  gap.className = GAP_CLASS;
  gap.setAttribute(PAGE_GAP_ATTR, '1');
  gap.setAttribute('contenteditable', 'false');
  gap.style.cssText = 'display:block;height:0;margin:0;padding:0;border:0;';
  return gap;
}

/** Вставка маркера в текст: разрывает текстовый узел и ставит блок между частями. */
function insertGapInside(gap: HTMLElement, position: FlowPosition): void {
  const { node, offset } = position;
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node as Text;
    const parent = text.parentElement;
    if (!parent) return;
    if (offset <= 0) parent.insertBefore(gap, text);
    else if (offset >= text.length) parent.insertBefore(gap, text.nextSibling);
    else {
      const tail = text.splitText(offset);
      tail.parentElement?.insertBefore(gap, tail);
    }
    return;
  }
  if (node.nodeType === Node.ELEMENT_NODE) {
    (node as Element).parentElement?.insertBefore(gap, node);
  }
}

/** Маркер, который уже стоит перед br с принудительным разрывом. */
function gapInAfterBreak(br: HTMLElement): HTMLElement | null {
  const sibling = br.previousSibling;
  if (sibling instanceof HTMLElement && sibling.hasAttribute(PAGE_GAP_ATTR)) return sibling;
  return null;
}

/** Первый текстовый узел блока после маркера. */
function textAfter(gap: HTMLElement | null, block: HTMLElement): FlowPosition | null {
  const nodes = textNodes(block);
  const first = gap
    ? nodes.find((text) => gap.compareDocumentPosition(text) & Node.DOCUMENT_POSITION_FOLLOWING)
    : nodes[0];
  return first ? { node: first, offset: 0 } : null;
}

/** Элемент w:br type="page", который ещё не обработан. */
function findPageBreak(block: HTMLElement, handled: Set<Element>): HTMLElement | null {
  const breaks = block.querySelectorAll<HTMLElement>('br[data-page-break]');
  for (const item of breaks) {
    if (handled.has(item)) continue;
    // Маркер разрыва остаётся в абзаце: если перед ним уже стоит маркер
    // листа, второй раз ломать не нужно (иначе бесконечная разметка).
    const previous = item.previousSibling;
    if (previous instanceof Element && previous.hasAttribute(PAGE_GAP_ATTR)) continue;
    return item;
  }
  return null;
}

/** Первая строка таблицы, которая не помещается в полосу. */
function fitRows(table: HTMLTableElement, meter: Meter, limit: number): number {
  const rows = [...table.rows];
  for (let index = 0; index < rows.length; index += 1) {
    if (meter.bottom(rows[index]) <= limit + EPSILON) continue;
    return index;
  }
  return -1;
}

/**
 * Разрезает таблицу: строки начиная с rowIndex уходят в клон-продолжение.
 * Шапка повторяется (Word так делает на новой странице).
 */
function continuationTable(table: HTMLTableElement, rowIndex: number): HTMLTableElement | null {
  const rows = [...table.rows];
  if (rowIndex <= 0 || rowIndex >= rows.length) return null;
  const clone = table.cloneNode(false) as HTMLTableElement;
  clone.setAttribute(TABLE_CONT_ATTR, '1');
  const colgroup = table.querySelector(':scope > colgroup');
  if (colgroup) clone.append(colgroup.cloneNode(true));
  const thead = table.querySelector(':scope > thead');
  if (thead) clone.append(thead.cloneNode(true));
  const tbody = document.createElement('tbody');
  for (const row of rows.slice(rowIndex)) tbody.append(row);
  clone.append(tbody);
  table.insertAdjacentElement('afterend', clone);
  return clone;
}

/* ------------------------------------------------------------------- строки */

/** Текстовые узлы блока в порядке документа, без маркеров. */
function textNodes(block: HTMLElement, from?: FlowPosition): Text[] {
  const result: Text[] = [];
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parent = (node as Text).parentElement;
      if (parent?.closest(`[${PAGE_GAP_ATTR}]`)) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });
  let started = !from;
  let node = walker.nextNode();
  while (node) {
    const text = node as Text;
    if (!started && from) {
      if (text === from.node) started = true;
      else {
        node = walker.nextNode();
        continue;
      }
    }
    if ((text.textContent ?? '').length > 0) result.push(text);
    node = walker.nextNode();
  }
  return result;
}

/** Нижняя граница строки, на которую попадает смещение (координата полотна). */
function lineBottomAt(text: Text, offset: number, meter: Meter): number | null {
  if (text.length === 0) return null;
  const range = document.createRange();
  range.setStart(text, Math.max(0, Math.min(text.length - 1, offset)));
  range.setEnd(text, Math.min(text.length, Math.max(1, offset + 1)));
  const rect = range.getBoundingClientRect();
  if (!rect.height && !rect.width) return null;
  return meter.at(rect.bottom);
}

/**
 * Строки блока с позициями их начала.
 * Обычно строка — это весь текстовый узел; многострочный узел делится
 * двоичным поиском по смещению (границы строк — ступенчатая функция).
 */
function collectLines(block: HTMLElement, meter: Meter, from?: FlowPosition): LineInfo[] {
  meter.refresh();
  const lines: LineInfo[] = [];
  for (const text of textNodes(block, from)) {
    const range = document.createRange();
    range.selectNodeContents(text);
    const rects = [...range.getClientRects()].filter((rect) => rect.height > 0);
    if (rects.length === 0) continue;
    const bottoms = [...new Set(rects.map((rect) => Math.round(meter.at(rect.bottom) * 10) / 10))].sort(
      (left, right) => left - right
    );
    if (bottoms.length <= 1) {
      lines.push({ start: { node: text, offset: 0 }, bottom: bottoms[0] ?? meter.at(rects[0].bottom) });
      continue;
    }
    let lineStart = 0;
    for (let index = 1; index < bottoms.length; index += 1) {
      const targetBottom = bottoms[index];
      let low = lineStart;
      let high = text.length;
      while (low < high) {
        const middle = Math.floor((low + high) / 2);
        const bottom = lineBottomAt(text, middle, meter);
        if (bottom === null || bottom <= targetBottom - 0.5) low = middle + 1;
        else high = middle;
      }
      const next = Math.min(text.length, low);
      lines.push({ start: { node: text, offset: lineStart }, bottom: bottoms[index - 1] });
      lineStart = next;
    }
    lines.push({ start: { node: text, offset: lineStart }, bottom: bottoms[bottoms.length - 1] });
  }
  return lines;
}

/**
 * Строка, после которой ставится разрыв.
 * Возвращает null, если блок лучше перенести целиком (в том числе из-за
 * висячих строк: Word держит минимум две строки с каждой стороны).
 */
function pickLineCut(
  lines: LineInfo[],
  limit: number
): { start: FlowPosition; index: number } | null {
  if (lines.length < 2) return null;
  let fit = -1;
  for (let index = 0; index < lines.length; index += 1) {
    if (lines[index].bottom <= limit + EPSILON) fit = index;
    else break;
  }
  if (fit < 0) return null;
  if (fit >= lines.length - 1) return null;
  if (fit === 0 && lines.length > 2) return null;
  if (lines.length - fit === 2) {
    // Одна строка ушла бы на следующий лист — переносим разрыв на строку выше.
    const previous = lines[fit - 1];
    if (previous) return { start: previous.start, index: fit - 1 };
    return null;
  }
  return { start: lines[fit + 1].start, index: fit + 1 };
}

/* ------------------------------------------------------------------- печать */

/**
 * Разметка HTML в скрытом контейнере + нарезка потока на страницы.
 * Нужно для печати файла, который не открыт в редакторе, и для печати
 * Word-документа: каждая страница печати — отдельный блок нужного размера,
 * поэтому принтер не переносит строки по своему усмотрению.
 */
export function paginateToPages(
  html: string,
  options: PaginateOptions & { className?: string }
): { pages: string[]; result: PaginateResult } {
  const geometry = options.geometry;
  const host = document.createElement('div');
  host.style.cssText = [
    'position:fixed',
    'left:-100000px',
    'top:0',
    `width:${geometry.contentWidthPx}px`,
    'visibility:hidden',
    'pointer-events:none',
    'z-index:-1',
  ].join(';');
  const flow = document.createElement('div');
  flow.className = options.className ?? 'word-sheet';
  flow.style.width = `${geometry.contentWidthPx}px`;
  flow.innerHTML = html;
  host.append(flow);
  document.body.append(host);
  try {
    const result = paginate(flow, options);
    return { pages: splitPaginatedFlow(flow), result };
  } finally {
    host.remove();
  }
}

/** Служебная проверка: элемент без содержимого (маркер разрыва). */
function isGapNode(node: Node): boolean {
  return node instanceof Element && node.hasAttribute(PAGE_GAP_ATTR);
}

function serialized(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent ?? '';
  if (node.nodeType !== Node.ELEMENT_NODE) return '';
  return (node as Element).outerHTML;
}

/**
 * Копия контейнера с частью содержимого.
 * У продолжения снимаются отступы сверху и красная строка — как у строк,
 * продолжившихся с предыдущего листа.
 */
function wrapFragment(element: Element, inner: string, index: number, previousItems: number): string {
  const clone = element.cloneNode(false) as Element;
  clone.innerHTML = inner;
  if (index > 0) {
    const style = ((clone as HTMLElement).getAttribute('style') ?? '').trim();
    (clone as HTMLElement).setAttribute(
      'style',
      `${style}${style && !style.endsWith(';') ? ';' : ''}margin-top:0;text-indent:0`
    );
    if (clone.tagName.toLowerCase() === 'ol') {
      const declared = Number((element as HTMLElement).getAttribute('start') ?? 1) || 1;
      clone.setAttribute('start', String(declared + previousItems));
    }
  }
  void index;
  return clone.outerHTML;
}

/**
 * Режет поток по маркерам разрывов на страницы.
 * Маркер может стоять внутри абзаца или списка — контейнеры клонируются, для
 * продолжения поправляются отступы и нумерация.
 */
export function splitPaginatedFlow(root: HTMLElement): string[] {
  const pages: string[] = [];
  let current: string[] = [];
  const flush = () => {
    pages.push(current.join(''));
    current = [];
  };

  const emit = (node: Node) => {
    if (isGapNode(node)) {
      flush();
      return;
    }
    if (node.nodeType === Node.ELEMENT_NODE) {
      const element = node as Element;
      if (element.querySelector(`[${PAGE_GAP_ATTR}]`)) {
        const parts = splitElement(element);
        parts.forEach((part, index) => {
          if (index > 0) flush();
          current.push(part);
        });
        return;
      }
    }
    current.push(serialized(node));
  };

  for (const child of [...root.childNodes]) emit(child);
  flush();
  return pages.length > 0 ? pages : [''];
}

/** Фрагменты элемента между маркерами разрывов. */
function splitElement(element: Element): string[] {
  const children = [...element.childNodes];
  const fragments: string[] = [];
  let buffer = '';

  const flush = () => {
    fragments.push(buffer);
    buffer = '';
  };

  for (const child of children) {
    if (isGapNode(child)) {
      flush();
      continue;
    }
    if (child.nodeType === Node.ELEMENT_NODE && (child as Element).querySelector(`[${PAGE_GAP_ATTR}]`)) {
      splitElement(child as Element).forEach((part, index) => {
        if (index > 0) flush();
        buffer += part;
      });
      continue;
    }
    buffer += serialized(child);
  }
  flush();

  const result: string[] = [];
  let previousItems = 0;
  fragments.forEach((inner, index) => {
    result.push(wrapFragment(element, inner, index, previousItems));
    previousItems += (inner.match(/<li\b/gi) ?? []).length;
  });
  return result;
}


/** Индекс листа, на котором находится узел. */
export function pageOfNode(node: Node, pages: PageBox[], root: HTMLElement): number {
  const element = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
  if (!element) return 0;
  const scale = layoutScale(root);
  const origin = root.getBoundingClientRect().top;
  const offset = (element.getBoundingClientRect().top - origin) / scale;
  for (let index = pages.length - 1; index >= 0; index -= 1) {
    if (offset >= pages[index].topPx) return index;
  }
  return 0;
}
