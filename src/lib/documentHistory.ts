/**
 * История AI-транзакций документа.
 *
 * Каждый ответ ассистента — одна транзакция: HTML до и после, вычисленный diff
 * и метка времени. История живёт в редакторе, не в чате, поэтому переживает
 * переключение вкладок и позволяет откатить/повторить правку независимо от
 * браузерного undo.
 */

export type BlockChangeKind = 'added' | 'removed' | 'changed';

export interface BlockChange {
  id: string;
  kind: BlockChangeKind;
  before?: string;
  after?: string;
}

export interface DocumentDiff {
  changed: number;
  added: number;
  removed: number;
  blocks: BlockChange[];
}

export interface DocumentTransaction {
  id: string;
  label: string;
  timestamp: number;
  before: string;
  after: string;
  diff: DocumentDiff;
}

export interface TransactionInput {
  id: string;
  label: string;
  timestamp: number;
  before: string;
  after: string;
}

const MAX_TRANSACTIONS = 50;
const SNIPPET_LENGTH = 120;

/** Сравнивает верхнеуровневые блоки двух HTML-снимков. */
export function diffDocumentHtml(beforeHtml: string, afterHtml: string): DocumentDiff {
  const before = blocks(beforeHtml);
  const after = blocks(afterHtml);
  const keys = new Set([...before.keys(), ...after.keys()]);
  const changes: BlockChange[] = [];

  for (const key of keys) {
    const left = before.get(key);
    const right = after.get(key);
    if (left && !right) {
      changes.push({ id: key, kind: 'removed', before: left.text });
      continue;
    }
    if (!left && right) {
      changes.push({ id: key, kind: 'added', after: right.text });
      continue;
    }
    if (left && right && left.hash !== right.hash) {
      changes.push({ id: key, kind: 'changed', before: left.text, after: right.text });
    }
  }

  return {
    changed: changes.filter((change) => change.kind === 'changed').length,
    added: changes.filter((change) => change.kind === 'added').length,
    removed: changes.filter((change) => change.kind === 'removed').length,
    blocks: changes,
  };
}

interface BlockSnapshot {
  hash: string;
  text: string;
}

function blocks(html: string): Map<string, BlockSnapshot> {
  const container = document.createElement('div');
  container.innerHTML = html || '';
  const result = new Map<string, BlockSnapshot>();
  const seen = new Map<string, number>();

  [...container.children].forEach((element, index) => {
    const rawKey =
      element.getAttribute('data-docx-id') ||
      element.getAttribute('data-bid') ||
      `i${index}`;
    const count = seen.get(rawKey) ?? 0;
    seen.set(rawKey, count + 1);
    const key = count === 0 ? rawKey : `${rawKey}#${count}`;
    result.set(key, signature(element));
  });

  return result;
}

/** Служебные id не должны выглядеть как изменение содержимого. */
function signature(element: Element): BlockSnapshot {
  const clone = element.cloneNode(true) as Element;
  for (const node of [clone, ...clone.querySelectorAll('*')]) {
    node.removeAttribute('data-bid');
    node.removeAttribute('data-docx-sig');
  }
  const text = (clone.textContent ?? '').replace(/\s+/g, ' ').trim();
  const html = clone.outerHTML;
  return {
    hash: `${html.length}:${hash(html)}`,
    text: text.length > SNIPPET_LENGTH ? `${text.slice(0, SNIPPET_LENGTH)}…` : text,
  };
}

function hash(value: string): string {
  let result = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    result ^= value.charCodeAt(index);
    result = Math.imul(result, 0x01000193);
  }
  return (result >>> 0).toString(16);
}

export class DocumentHistory {
  private transactions: DocumentTransaction[] = [];
  private cursor = -1;

  constructor(private readonly limit = MAX_TRANSACTIONS) {}

  get records(): readonly DocumentTransaction[] {
    return this.transactions;
  }

  get index(): number {
    return this.cursor;
  }

  get current(): DocumentTransaction | null {
    return this.cursor >= 0 ? this.transactions[this.cursor] ?? null : null;
  }

  canUndo(): boolean {
    return this.cursor >= 0;
  }

  canRedo(): boolean {
    return this.cursor + 1 < this.transactions.length;
  }

  commit(input: TransactionInput): DocumentTransaction {
    // Новая правка после отката обрезает «будущее»: redo больше не актуален.
    if (this.cursor + 1 < this.transactions.length) {
      this.transactions = this.transactions.slice(0, this.cursor + 1);
    }

    const transaction: DocumentTransaction = {
      ...input,
      diff: diffDocumentHtml(input.before, input.after),
    };
    this.transactions.push(transaction);
    if (this.transactions.length > this.limit) this.transactions.shift();
    this.cursor = this.transactions.length - 1;
    return transaction;
  }

  undo(): DocumentTransaction | null {
    if (!this.canUndo()) return null;
    const transaction = this.transactions[this.cursor];
    this.cursor -= 1;
    return transaction;
  }

  redo(): DocumentTransaction | null {
    if (!this.canRedo()) return null;
    this.cursor += 1;
    return this.transactions[this.cursor];
  }

  /** Перейти к состоянию после транзакции index (для списка истории). */
  goto(index: number): DocumentTransaction | null {
    if (index < 0 || index >= this.transactions.length) return null;
    this.cursor = index;
    return this.transactions[index];
  }

  clear(): void {
    this.transactions = [];
    this.cursor = -1;
  }
}
