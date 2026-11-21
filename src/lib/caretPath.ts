/**
 * Каретка через пересборку разметки листов.
 *
 * После каждой правки постраничная разметка снимается и ставится заново:
 * маркеры разрывов вырезаются, текст может быть разбит на новые текстовые
 * узлы. Обычное выделение браузера при этом съезжает, поэтому позиция
 * каретки запоминается как «блок + смещение в символах» и возвращается после
 * разметки. Служебные маркеры при подсчёте пропускаются — иначе смещение
 * уезжало бы на каждый разрыв.
 */
import { PAGE_GAP_ATTR } from './pageLayout';

const GAP_SELECTOR = `[${PAGE_GAP_ATTR}]`;

export interface CaretSnapshot {
  /** Индекс прямого потомка полотна (без маркеров разрыва). */
  childIndex: number;
  /** Смещение в символах от начала блока. */
  textOffset: number;
}

/** Прямые дети полотна без служебных маркеров. */
export function flowChildren(root: HTMLElement): HTMLElement[] {
  return [...root.children].filter(
    (child): child is HTMLElement =>
      child instanceof HTMLElement && !child.hasAttribute(PAGE_GAP_ATTR)
  );
}

function textWalker(node: Node): TreeWalker {
  return document.createTreeWalker(node, NodeFilter.SHOW_TEXT, {
    acceptNode(candidate) {
      const parent = (candidate as Text).parentElement;
      if (parent?.closest(GAP_SELECTOR)) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });
}

/** Смещение в символах от начала блока до позиции (node, offset). */
function offsetWithin(block: Element, node: Node, offset: number): number {
  const range = document.createRange();
  try {
    range.setStart(block, 0);
    range.setEnd(node, offset);
    return range.toString().length;
  } catch {
    return 0;
  }
}

export function captureCaret(root: HTMLElement): CaretSnapshot | null {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return null;
  const range = selection.getRangeAt(0);
  return snapshotPoint(root, range.startContainer, range.startOffset);
}

/** Выделение целиком: начало и конец, чтобы выделенный текст переживал разметку. */
export interface SelectionSnapshot {
  start: CaretSnapshot;
  end: CaretSnapshot | null;
  collapsed: boolean;
}

function snapshotPoint(
  root: HTMLElement,
  container: Node,
  offset: number
): CaretSnapshot | null {
  if (!root.contains(container)) return null;
  const element =
    container.nodeType === Node.ELEMENT_NODE
      ? (container as Element)
      : container.parentElement;
  if (!element) return null;
  const children = flowChildren(root);
  const block = children.find((child) => child.contains(element));
  if (!block) return null;
  return {
    childIndex: children.indexOf(block),
    textOffset: offsetWithin(block, container, offset),
  };
}

export function captureSelection(root: HTMLElement): SelectionSnapshot | null {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return null;
  const range = selection.getRangeAt(0);
  const start = snapshotPoint(root, range.startContainer, range.startOffset);
  if (!start) return null;
  const end = range.collapsed
    ? null
    : snapshotPoint(root, range.endContainer, range.endOffset);
  return { start, end, collapsed: range.collapsed || !end };
}

/** Точка (node, offset) по номеру блока и смещению в символах. */
function pointAt(root: HTMLElement, snapshot: CaretSnapshot): { block: Element; node: Text | null; offset: number } | null {
  const children = flowChildren(root);
  const block = children[snapshot.childIndex] ?? children[children.length - 1];
  if (!block) return null;
  const walker = textWalker(block);
  let remaining = snapshot.textOffset;
  let node = walker.nextNode();
  while (node) {
    const length = (node.textContent ?? '').length;
    if (remaining <= length) return { block, node: node as Text, offset: remaining };
    remaining -= length;
    node = walker.nextNode();
  }
  return { block, node: null, offset: 0 };
}

/** Возвращает выделение (или каретку) в запомненное место. */
export function restoreSelection(root: HTMLElement, snapshot: SelectionSnapshot | null): void {
  if (!snapshot) return;
  const selection = window.getSelection();
  if (!selection) return;
  const start = pointAt(root, snapshot.start);
  if (!start) return;
  const end = snapshot.end ? pointAt(root, snapshot.end) : null;
  const range = document.createRange();
  const place = (target: { block: Element; node: Text | null; offset: number }, isEnd: boolean) => {
    if (target.node) {
      const offset = Math.max(0, Math.min(target.node.length, target.offset));
      if (isEnd) range.setEnd(target.node, offset);
      else range.setStart(target.node, offset);
      return;
    }
    // Блок без текста (картинка) или смещение за концом.
    if (isEnd) range.setEnd(target.block, target.block.childNodes.length);
    else range.setStart(target.block, 0);
  };
  try {
    place(start, false);
    if (end && !snapshot.collapsed) place(end, true);
    else range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  } catch {
    // Узлы могли переехать: выделение просто не восстанавливаем.
  }
}

/** Возвращает каретку в запомненное место; дочитывает до конца блока при переполнении. */
export function restoreCaret(root: HTMLElement, snapshot: CaretSnapshot | null): void {
  if (!snapshot) return;
  const children = flowChildren(root);
  const block = children[snapshot.childIndex] ?? children[children.length - 1];
  if (!block) return;
  const walker = textWalker(block);
  let remaining = snapshot.textOffset;
  let node = walker.nextNode();
  const target = (text: Text, offset: number) => {
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.setStart(text, Math.max(0, Math.min(text.length, offset)));
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  };
  while (node) {
    const length = (node.textContent ?? '').length;
    if (remaining <= length) {
      target(node as Text, remaining);
      return;
    }
    remaining -= length;
    node = walker.nextNode();
  }
  // Блок без текста (картинка) или смещение за концом: ставим каретку в конец.
  const selection = window.getSelection();
  if (!selection) return;
  const range = document.createRange();
  range.selectNodeContents(block);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}
