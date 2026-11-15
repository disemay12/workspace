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

/** Путь в одном виде: снимок на диске и вкладка должны совпасть по файлу. */
function normalizePath(value: string): string {
  return (value || '').replace(/\//g, '\\').toLowerCase();
}

function samePath(a: string, b: string): boolean {
  return normalizePath(a) === normalizePath(b);
}

/**
 * Закрытые вкладки. Снимок ставится в очередь снимком состояния, а запись идёт
 * через несколько `await`: вкладка, закрытая между проверкой и записью, иначе
 * вернула бы при следующем запуске черновик, который пользователь выбросил.
 */
const closedPaths = new Map<string, number>();
const CLOSED_PATH_TTL_MS = 30_000;

/** Вкладку закрыли с отказом от сохранения: её снимок писать нельзя. */
export function noteClosedPath(filePath: string): void {
  const key = normalizePath(filePath);
  if (!key) return;
  const now = Date.now();
  for (const [path, at] of closedPaths) {
    if (now - at >= CLOSED_PATH_TTL_MS) closedPaths.delete(path);
  }
  closedPaths.set(key, now);
}

/** Файл открыли снова: запрет на снимок снимается. */
export function noteOpenedPath(filePath: string): void {
  closedPaths.delete(normalizePath(filePath));
}

function isClosedRecently(filePath: string): boolean {
  const key = normalizePath(filePath);
  const at = closedPaths.get(key);
  if (at === undefined) return false;
  if (Date.now() - at >= CLOSED_PATH_TTL_MS) {
    closedPaths.delete(key);
    return false;
  }
  return true;
}

const FNV64_OFFSET = 0xcbf29ce484222325n;
const FNV64_PRIME = 0x100000001b3n;
const FNV64_MASK = 0xffffffffffffffffn;

/**
 * Второй хеш пути. FNV-1a в snapshot.ts 32-битный: у двух разных файлов он
 * может совпасть, и снимок одного затрёт снимок другого. Здесь 64 бита.
 */
function pathSuffix(filePath: string): string {
  const normalized = (filePath || '').replace(/\//g, '\\').toLowerCase();
  let hash = FNV64_OFFSET;
  for (let index = 0; index < normalized.length; index += 1) {
    hash ^= BigInt(normalized.charCodeAt(index));
    hash = (hash * FNV64_PRIME) & FNV64_MASK;
  }
  // Только [a-z0-9_-]: main вычищает остальные символы из имени файла.
  return hash.toString(36);
}

/**
 * Владельцы имён снимков. Без кэша каждая запись читала через IPC содержимое
 * всех снимков (до 8 МБ каждый) только ради проверки коллизии.
 */
const snapshotOwners = new Map<string, string>();
let snapshotOwnersLoaded = false;
let snapshotOwnersLoading: Promise<void> | null = null;

/** Список читаем один раз; force — перечитать после обнаруженной коллизии. */
async function loadSnapshotOwners(force = false): Promise<void> {
  if (snapshotOwnersLoaded && !force) return;
  if (snapshotOwnersLoading) return snapshotOwnersLoading;
  const list = window.electronAPI?.autosaveListMeta ?? window.electronAPI?.autosaveList;
  if (!list) return;
  const loading = (async () => {
    try {
      const result = await list();
      if (!result?.success || !Array.isArray(result.items)) return;
      snapshotOwners.clear();
      for (const item of result.items as AutosaveSnapshot[]) {
        if (item && typeof item.id === 'string' && typeof item.path === 'string') {
          snapshotOwners.set(item.id, item.path);
        }
      }
      snapshotOwnersLoaded = true;
    } catch {
      // IPC может быть недоступен (веб-режим, старый preload) — попробуем позже.
    }
  })();
  snapshotOwnersLoading = loading;
  try {
    await loading;
  } finally {
    if (snapshotOwnersLoading === loading) snapshotOwnersLoading = null;
  }
}

/** Метаданные снимков: содержимое для проверок владельца и срока не нужно. */
async function readSnapshotMeta(): Promise<AutosaveSnapshot[]> {
  const api = window.electronAPI;
  if (!api) return [];
  let result: { success?: boolean; items?: unknown[] } | undefined;
  try {
    result = api.autosaveListMeta
      ? await api.autosaveListMeta()
      : await api.autosaveList?.();
  } catch {
    // IPC может быть недоступен (веб-режим, старый preload) — не повод падать
    return [];
  }
  if (!result?.success || !Array.isArray(result.items)) return [];
  return (result.items as AutosaveSnapshot[]).filter(
    (item) => item && typeof item.path === 'string' && item.path.length > 0
  );
}

/** Сырые снимки с диска, включая содержимое: нужны только для восстановления. */
async function readSnapshotItems(): Promise<AutosaveSnapshot[]> {
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
  return (result.items as AutosaveSnapshot[]).filter(
    (item) =>
      item &&
      typeof item.path === 'string' &&
      item.path.length > 0 &&
      typeof item.content === 'string' &&
      item.content.length > 0
  );
}

/** Имя снимка для файла: своё переиспользуем, чужое (коллизия хеша) — нет. */
async function snapshotNameFor(filePath: string): Promise<string> {
  await loadSnapshotOwners();
  const base = snapshotId(filePath);
  const owner = snapshotOwners.get(base);
  if (!owner || samePath(owner, filePath)) return base;
  const suffixed = `${base}_${pathSuffix(filePath)}`;
  // Суффикс уже наш: перечитывать список незачем.
  const suffixedOwner = snapshotOwners.get(suffixed);
  if (suffixedOwner && samePath(suffixedOwner, filePath)) return suffixed;
  // Только при коллизии перечитываем список: имя могло освободиться.
  await loadSnapshotOwners(true);
  const actual = snapshotOwners.get(base);
  if (!actual || samePath(actual, filePath)) return base;
  return suffixed;
}

/** Возвращает true, только если снимок действительно записан на диск. */
export async function writeSnapshot(tab: OpenTab): Promise<boolean> {
  const api = window.electronAPI;
  if (!api?.autosaveWrite) return false;
  if (isClosedRecently(tab.path)) return false;
  if (!shouldSnapshot({ isDirty: tab.isDirty, content: tab.content, maxBytes: MAX_BYTES })) {
    return false;
  }
  try {
    const id = await snapshotNameFor(tab.path);
    // Вкладку могли закрыть, пока выбиралось имя снимка.
    if (isClosedRecently(tab.path)) return false;
    const snapshot: AutosaveSnapshot = {
      id,
      path: tab.path,
      name: tab.name,
      type: tab.type,
      content: tab.content,
      savedAt: Date.now(),
    };
    const result = await api.autosaveWrite(id, snapshot);
    // Главный процесс на отказе не бросает, а возвращает success: false —
    // без проверки снимок терялся молча.
    if (result?.success === false) {
      void Promise.resolve(
        api.diagnosticsLog?.(
          'warn',
          `autosave: снимок «${tab.name}» не записан: ${result.error ?? 'причина неизвестна'}`
        )
      ).catch(() => undefined);
      return false;
    }
    snapshotOwners.set(id, tab.path);
    return true;
  } catch {
    // запись снимка не должна ломать редактирование
    return false;
  }
}

export async function listSnapshots(): Promise<AutosaveSnapshot[]> {
  const now = Date.now();
  return (await readSnapshotItems()).filter(
    (item) => now - (Number(item.savedAt) || 0) < MAX_AGE_MS
  );
}

export async function dropSnapshot(id: string): Promise<void> {
  const api = window.electronAPI;
  if (!api?.autosaveRemove) return;
  try {
    const result = await api.autosaveRemove(id);
    // Имя освободилось: иначе следующий файл получит суффикс без нужды.
    if (result?.success) snapshotOwners.delete(id);
  } catch {
    // Снимок — страховка от падения: сбой его удаления не должен валить
    // сохранение файла, которое уже прошло.
  }
}

/**
 * Убирает снимки конкретного файла: имя могло получить суффикс при коллизии.
 * olderThan — начало записи файла: снимки, сделанные позже, содержат правки,
 * которых в записанном файле нет, и стирать их нельзя.
 */
export async function dropSnapshotsForPath(filePath: string, olderThan?: number): Promise<void> {
  for (const snapshot of await readSnapshotMeta()) {
    if (!samePath(snapshot.path, filePath)) continue;
    if (olderThan !== undefined && Number(snapshot.savedAt) > olderThan) continue;
    await dropSnapshot(snapshot.id);
  }
}

/** Снимает снимки сохранённых документов и просроченные: иначе каталог растёт. */
export async function pruneSnapshots(savedPaths: string[]): Promise<void> {
  const now = Date.now();
  for (const snapshot of await readSnapshotMeta()) {
    const expired = now - (Number(snapshot.savedAt) || 0) >= MAX_AGE_MS;
    // Сверяем путь, а не только id: id мог получить суффикс из-за коллизии.
    const saved = savedPaths.some((path) => samePath(path, snapshot.path));
    if (expired || saved) await dropSnapshot(snapshot.id);
  }
}
