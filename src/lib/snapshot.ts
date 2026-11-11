/**
 * Идентификаторы снимков автосохранения.
 *
 * Чистые функции без Electron: используются и в сервисе, и в проверках.
 */
export function snapshotId(filePath: string): string {
  const normalized = (filePath || '').replace(/\//g, '\\').toLowerCase();
  let hash = 2166136261;
  for (let index = 0; index < normalized.length; index += 1) {
    hash ^= normalized.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return `snap-${(hash >>> 0).toString(36)}`;
}

/** Разрешённые имена: только строчные буквы, цифры, дефис и подчёркивание. */
export function safeSnapshotName(value: string): string {
  return (value || '').replace(/[^a-z0-9_-]/gi, '').slice(0, 64);
}

/** Снимок нужен, только если в нём есть что восстанавливать. */
export function shouldSnapshot(input: {
  isDirty: boolean;
  content?: string;
  maxBytes?: number;
}): boolean {
  if (!input.isDirty) return false;
  const content = input.content ?? '';
  if (!content) return false;
  const limit = input.maxBytes ?? 8 * 1024 * 1024;
  return content.length <= limit;
}
