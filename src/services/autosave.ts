/**
 * Автосохранение черновиков: снимок содержимого вкладки в профиль пользователя.
 *
 * Снимки живут в `%APPDATA%\WorkSpace\autosave` и удаляются сразу после
 * успешного сохранения файла или восстановления. Это страховка от падения,
 * а не замена сохранению.
 */
import type { OpenTab } from '../types';
import { shouldSnapshot, snapshotId } from '../lib/snapshot';

/** Снимок не делаем для очень больших буферов и не держим дольше недели. */
const MAX_BYTES = 8 * 1024 * 1024;
const MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;

export interface AutosaveSnapshot {
  id: string;
  path: string;
  name: string;
  type: string;
  content: string;
  savedAt: number;
}

export async function writeSnapshot(tab: OpenTab): Promise<void> {
  const api = window.electronAPI;
  if (!api?.autosaveWrite) return;
  if (!shouldSnapshot({ isDirty: tab.isDirty, content: tab.content, maxBytes: MAX_BYTES })) return;
  const snapshot: AutosaveSnapshot = {
    id: snapshotId(tab.path),
    path: tab.path,
    name: tab.name,
    type: tab.type,
    content: tab.content,
    savedAt: Date.now(),
  };
  try {
    await api.autosaveWrite(snapshot.id, snapshot);
  } catch {
    // запись снимка не должна ломать редактирование
  }
}

export async function listSnapshots(): Promise<AutosaveSnapshot[]> {
  const api = window.electronAPI;
  if (!api?.autosaveList) return [];
  let result: { success?: boolean; items?: unknown[] } | undefined;
  try {
    result = await api.autosaveList();
  } catch {
    // IPC может быть недоступен (веб-режим, старый preload) — не повод падать
    return [];
  }
  if (!result?.success || !Array.isArray(result.items)) return [];
  const now = Date.now();
  return (result.items as AutosaveSnapshot[]).filter(
    (item) =>
      item &&
      typeof item.path === 'string' &&
      item.path.length > 0 &&
      typeof item.content === 'string' &&
      item.content.length > 0 &&
      now - (Number(item.savedAt) || 0) < MAX_AGE_MS
  );
}

export async function dropSnapshot(id: string): Promise<void> {
  await window.electronAPI?.autosaveRemove?.(id);
}

/** Снимает снимки для документов, которые уже сохранены. */
export async function pruneSnapshots(savedPaths: string[]): Promise<void> {
  if (savedPaths.length === 0) return;
  const wanted = new Set(savedPaths.map((path) => snapshotId(path)));
  for (const snapshot of await listSnapshots()) {
    if (wanted.has(snapshot.id)) await dropSnapshot(snapshot.id);
  }
}
