/**
 * Наблюдение за изменениями в рабочей папке (File System Watcher).
 *
 * Использует нативный fs.watch с рекурсивным режимом на Windows.
 * Фильтрует служебные каталоги (node_modules, .git, dist) и временные файлы,
 * объединяет частые события через debounce и шлёт в рендерер событие
 * `fs:workspaceChanged`.
 */
import * as fs from 'fs';
import * as path from 'path';
import { BrowserWindow } from 'electron';

let currentWatcher: fs.FSWatcher | null = null;
let debounceTimer: NodeJS.Timeout | null = null;
let currentWatchedDir = '';

const IGNORED_PATTERNS = [
  /[\\/]\.git([\\/]|$)/i,
  /[\\/]node_modules([\\/]|$)/i,
  /[\\/]dist(-electron)?([\\/]|$)/i,
  /[\\/]\.tmp/i,
  /\.ws-tmp$/i,
  /[\\/]release([\\/]|$)/i,
  /[\\/]\.vscode([\\/]|$)/i,
  /[\\/]\.idea([\\/]|$)/i,
];

export function setWatchRoot(dirPath: string, windowGetter: () => BrowserWindow | null): void {
  if (currentWatcher) {
    try {
      currentWatcher.close();
    } catch {
      // Игнорируем ошибку закрытия предыдущего watcher
    }
    currentWatcher = null;
    currentWatchedDir = '';
  }

  if (!dirPath || !fs.existsSync(dirPath)) return;

  try {
    currentWatchedDir = dirPath;
    currentWatcher = fs.watch(dirPath, { recursive: true }, (eventType, filename) => {
      if (!filename) return;

      const fullPath = path.isAbsolute(filename) ? filename : path.join(dirPath, filename);
      if (IGNORED_PATTERNS.some((p) => p.test(fullPath))) return;

      if (debounceTimer) clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        const win = windowGetter();
        if (win && !win.isDestroyed()) {
          win.webContents.send('fs:workspaceChanged', {
            eventType,
            filename,
            fullPath,
          });
        }
      }, 250);
    });
  } catch (err) {
    console.warn('[Watcher] Не удалось запустить отслеживание папки:', dirPath, (err as Error).message);
  }
}

export function stopWatch(): void {
  if (currentWatcher) {
    try {
      currentWatcher.close();
    } catch {
      // ignore
    }
    currentWatcher = null;
    currentWatchedDir = '';
  }
}
