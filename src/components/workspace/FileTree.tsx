import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  File,
  FileCode,
  FileSpreadsheet,
  FileText,
  FileType2,
  Folder,
  FolderOpen,
  Presentation,
  Image as ImageIcon,
  ChevronRight,
  ChevronDown,
  Loader2,
  Files,
  RefreshCw,
  FolderPlus,
  FilePlus2,
  Clapperboard,
  ExternalLink,
  Trash2,
  Pencil,
  Copy,
  CloudUpload,
  Share2,
  X,
  FolderOpen as RevealIcon,
  Star,
  FolderInput,
  FileSearch,
  SlidersHorizontal,
} from 'lucide-react';
import { FileItem, FileType } from '../../types';
import { cn, formatBytes } from '../../lib/utils';
import {
  ExplorerHistory,
  applyRenamePattern,
  canDrop,
  childItems,
  flattenVisible,
  formatDate,
  matchesFilter as matchesExplorerFilter,
} from '../../lib/explorerOps';
import { readPreferences, writePreferences } from '../../state/preferences';
import { fileSystemService, imageDataUrl } from '../../services/fileSystemService';
import { ContextMenu, ContextMenuItem } from './ContextMenu';
import { SHARE_SERVICES } from '../../services/shareService';

const THUMB_CACHE = new Map<string, string>();

/** Миниатюра изображения: файл читается один раз и кешируется как data-URL. */
const FileThumb: React.FC<{ item: FileItem }> = ({ item }) => {
  const [src, setSrc] = useState(() => THUMB_CACHE.get(item.path) ?? '');
  useEffect(() => {
    if (src || item.size > 8 * 1024 * 1024) return;
    let alive = true;
    void (async () => {
      try {
        const bytes = new Uint8Array(await fileSystemService.readBinary(item.path));
        let binary = '';
        for (let index = 0; index < bytes.length; index += 0x8000) {
          binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
        }
        const url = imageDataUrl(btoa(binary), item.extension);
        THUMB_CACHE.set(item.path, url);
        if (alive) setSrc(url);
      } catch {
        // превью не критично: останется значок типа
      }
    })();
    return () => {
      alive = false;
    };
  }, [item.path, item.extension, item.size, src]);
  return src ? (
    <img src={src} alt="" loading="lazy" className="h-4 w-4 shrink-0 rounded-sm object-cover" />
  ) : (
    fileIcon(item.type, item.name)
  );
};

interface Props {
  rootPath: string;
  tree: { [dirPath: string]: FileItem[] };
  expandedDirs: Set<string>;
  loadingDirs: Set<string>;
  activePath: string;
  dirtyPaths: Set<string>;
  onToggleDir: (dirPath: string) => void;
  onOpenFile: (filePath: string, pane?: 'primary' | 'secondary') => void;
  onPreviewFile?: (filePath: string) => void;
  onRefresh: () => void;
  onNewFile: (parentDir: string) => void;
  onNewFolder: (parentDir: string) => void;
  onOpenFolder: () => void;
  onOpenGoogleDriveSettings?: () => void;
  favorites?: { path: string; name: string }[];
  onOpenFavorite?: (dirPath: string) => void;
  onToggleFavorite?: (dirPath: string) => void;
  /** Inline-переименование: возвращает новый путь или null, если не удалось. */
  onRenamePath?: (filePath: string, newName: string) => Promise<string | null>;
  /** Перемещение/копирование выделения в папку (drag&drop и Ctrl+V). */
  onMovePaths?: (paths: string[], targetDir: string, copy: boolean) => Promise<void>;
  /** Файл, открытый из контекстного меню Windows: подсветить и прокрутить. */
  revealTarget?: { path: string; nonce: number } | null;
  onItemAction: (
    action: string,
    item: FileItem,
    serviceId?: string,
    selection?: FileItem[]
  ) => void;
}

/** Иконка и цвет по типу файла — визуальная навигация в дереве. */
function fileIcon(type: FileType | undefined, name: string) {
  switch (type) {
    case 'excel':
      return <FileSpreadsheet size={14} className="shrink-0 text-emerald-500" />;
    case 'word':
      return <FileText size={14} className="shrink-0 text-blue-500" />;
    case 'powerpoint':
      return <Presentation size={14} className="shrink-0 text-orange-500" />;
    case 'markdown':
    case 'richtext':
      return <FileType2 size={14} className="shrink-0 text-purple-400" />;
    case 'code':
      return <FileCode size={14} className="shrink-0 text-amber-500" />;
    case 'image':
      return <ImageIcon size={14} className="shrink-0 text-pink-400" />;
    case 'pdf':
      return <FileText size={14} className="shrink-0 text-red-500" />;
    case 'video':
      return <Clapperboard size={14} className="shrink-0 text-sky-400" />;
    case 'text':
      return <File size={14} className="shrink-0 text-ws-faint" />;
    default:
      return <File size={14} className="shrink-0 text-ws-faint" />;
  }
}

export const FileTree: React.FC<Props> = ({
  rootPath,
  tree,
  expandedDirs,
  loadingDirs,
  activePath,
  dirtyPaths,
  onToggleDir,
  onOpenFile,
  onPreviewFile,
  onRefresh,
  onNewFile,
  onNewFolder,
  onOpenFolder,
  onOpenGoogleDriveSettings,
  favorites = [],
  onOpenFavorite,
  onToggleFavorite,
  onItemAction,
  revealTarget,
  onRenamePath,
  onMovePaths,
}) => {
  const [filter, setFilter] = useState('');
  const [menu, setMenu] = useState<{ x: number; y: number; item: FileItem } | null>(null);
  const [selectedPaths, setSelectedPaths] = useState<string[]>([]);
  const [focusedPath, setFocusedPath] = useState('');
  const [renaming, setRenaming] = useState<{ path: string; value: string } | null>(null);
  const [clipboard, setClipboard] = useState<{ mode: 'cut' | 'copy'; paths: string[] } | null>(null);
  const [dropTarget, setDropTarget] = useState('');
  const [pathEdit, setPathEdit] = useState<{ open: boolean; value: string }>({ open: false, value: '' });
  const [bulk, setBulk] = useState<{ open: boolean; pattern: string; paths: string[] }>({
    open: false,
    pattern: '{name}',
    paths: [],
  });
  const [explorerPrefs, setExplorerPrefs] = useState(() => readPreferences().explorer);
  const historyRef = useRef(new ExplorerHistory(50));
  const [historySize, setHistorySize] = useState(0);
  const anchorRef = useRef('');
  /** Строки дерева по нормализованному пути — для прокрутки к открытому файлу. */
  const rowRefs = useRef<Map<string, HTMLDivElement>>(new Map());

  const normalizePath = (value: string) => value.toLowerCase().replace(/\//g, '\\');

  useEffect(() => {
    if (!revealTarget?.path) return;
    const element = rowRefs.current.get(normalizePath(revealTarget.path));
    if (!element) {
      // Мешает фильтр — снимаем его, чтобы строка существовала.
      if (filter) setFilter('');
      return;
    }
    element.scrollIntoView({ block: 'center' });
    setSelectedPaths([revealTarget.path]);
    element.classList.add('ws-reveal-flash');
    const timer = window.setTimeout(() => element.classList.remove('ws-reveal-flash'), 1600);
    return () => window.clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [revealTarget, filter]);

  /** Рекурсивная фильтрация дерева по подстроке имени. */
  const matchesFilter = (item: FileItem): boolean => {
    if (!filter.trim()) return true;
    if (matchesExplorerFilter(item, filter)) return true;
    if (item.isDirectory) {
      const children = tree[item.path] || [];
      return children.some(matchesFilter);
    }
    return false;
  };

  const viewOptions = useMemo(
    () => ({
      sort: {
        by: explorerPrefs.sortBy,
        dir: explorerPrefs.dir,
        foldersFirst: explorerPrefs.foldersFirst,
      },
      showHidden: explorerPrefs.showHidden,
      filter,
    }),
    [explorerPrefs, filter]
  );

  const updateExplorerPrefs = (patch: Partial<typeof explorerPrefs>) => {
    const next = { ...explorerPrefs, ...patch };
    setExplorerPrefs(next);
    writePreferences({ ...readPreferences(), explorer: next });
  };

  const rootItems = useMemo(() => tree[rootPath] || [], [tree, rootPath]);

  /** Плоский список видимых строк: выделение диапазона, фокус, виртуализация. */
  const visibleFlat = useMemo(
    () => (rootPath ? flattenVisible(tree, expandedDirs, rootPath, viewOptions) : []),
    [tree, rootPath, expandedDirs, viewOptions]
  );

  const focusIndex = () => visibleFlat.findIndex((item) => item.path === focusedPath);

  const moveFocus = (delta: number) => {
    if (visibleFlat.length === 0) return;
    const index = focusIndex();
    const next = index < 0 ? 0 : Math.max(0, Math.min(visibleFlat.length - 1, index + delta));
    const item = visibleFlat[next];
    if (!item) return;
    setFocusedPath(item.path);
    setSelectedPaths([item.path]);
    rowRefs.current.get(normalizePath(item.path))?.scrollIntoView({ block: 'nearest' });
  };

  const startRename = (item: FileItem) => setRenaming({ path: item.path, value: item.name });

  const commitRename = async () => {
    if (!renaming) return;
    const item = visibleFlat.find((entry) => entry.path === renaming.path);
    const value = renaming.value.trim();
    setRenaming(null);
    if (!item || !value || value === item.name || !onRenamePath) return;
    try {
      const target = await onRenamePath(item.path, value);
      if (target) {
        historyRef.current.push({ kind: 'rename', from: item.path, to: target });
        setHistorySize(historyRef.current.size);
      }
    } catch {
      // сообщение об ошибке показывает App
    }
  };

  const runPaste = async () => {
    if (!clipboard || clipboard.paths.length === 0 || !onMovePaths) return;
    const targetItem = visibleFlat.find((item) => item.path === focusedPath);
    const dir = targetItem?.isDirectory
      ? targetItem.path
      : (targetItem?.path ?? rootPath).replace(/[\\/][^\\/]+$/, '') || rootPath;
    const copy = clipboard.mode === 'copy';
    await onMovePaths(clipboard.paths, dir, copy);
    if (clipboard.mode === 'cut') setClipboard(null);
  };

  const undoLast = async () => {
    const op = historyRef.current.pop();
    setHistorySize(historyRef.current.size);
    if (!op || !onMovePaths) return;
    const inverse = ExplorerHistory.invert(op);
    if (inverse.moveFrom && inverse.moveTo) {
      const targetDir = inverse.moveTo.replace(/[\\/][^\\/]+$/, '');
      await onMovePaths([inverse.moveFrom], targetDir, false);
    }
  };

  const handleTreeKeys = (event: React.KeyboardEvent<HTMLDivElement>) => {
    if (renaming) return;
    const mod = event.ctrlKey || event.metaKey;
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      moveFocus(1);
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      moveFocus(-1);
    } else if (event.key === 'Enter') {
      event.preventDefault();
      const item = visibleFlat.find((entry) => entry.path === focusedPath);
      if (!item) return;
      if (item.isDirectory) onToggleDir(item.path);
      else onOpenFile(item.path);
    } else if (event.key === 'F2') {
      event.preventDefault();
      const item = visibleFlat.find((entry) => entry.path === focusedPath);
      if (item) startRename(item);
    } else if (event.key === 'Delete') {
      event.preventDefault();
      const item = visibleFlat.find((entry) => entry.path === focusedPath);
      if (item) onItemAction('delete', item, undefined, selectedItems);
    } else if (mod && event.key.toLowerCase() === 'a') {
      event.preventDefault();
      setSelectedPaths(visibleFlat.map((item) => item.path));
    } else if (mod && event.key.toLowerCase() === 'c') {
      event.preventDefault();
      if (selectedPaths.length) setClipboard({ mode: 'copy', paths: selectedPaths });
    } else if (mod && event.key.toLowerCase() === 'x') {
      event.preventDefault();
      if (selectedPaths.length) setClipboard({ mode: 'cut', paths: selectedPaths });
    } else if (mod && event.key.toLowerCase() === 'v') {
      event.preventDefault();
      void runPaste();
    } else if (mod && event.key.toLowerCase() === 'z') {
      event.preventDefault();
      void undoLast();
    } else if (mod && event.key.toLowerCase() === 'l') {
      event.preventDefault();
      setPathEdit({ open: true, value: rootPath });
    } else if (event.key === 'Escape') {
      setSelectedPaths([]);
      setRenaming(null);
    }
  };

  const handleDropOn = async (target: FileItem | null) => {
    setDropTarget('');
    const paths = selectedPaths.length > 0 ? selectedPaths : [];
    if (!onMovePaths || paths.length === 0) return;
    const dir = target?.isDirectory ? target.path : rootPath;
    const sources = paths.map((path) => ({
      path,
      name: path.split(/[\\/]/).pop() ?? path,
      isDirectory: tree[path.replace(/[\\/][^\\/]+$/, '')]?.some(
        (item) => item.path === path && item.isDirectory
      ) ?? false,
    }));
    if (sources.some((source) => !canDrop(source, { path: dir, isDirectory: true }).ok)) return;
    await onMovePaths(paths, dir, false);
  };

  const selectedItems = useMemo(
    () => visibleFlat.filter((item) => selectedPaths.includes(item.path)),
    [visibleFlat, selectedPaths]
  );

  const selectItem = (item: FileItem, event: React.MouseEvent) => {
    const extend = event.shiftKey;
    const toggle = event.ctrlKey || event.metaKey;
    if (extend && anchorRef.current) {
      const from = visibleFlat.findIndex((it) => it.path === anchorRef.current);
      const to = visibleFlat.findIndex((it) => it.path === item.path);
      if (from >= 0 && to >= 0) {
        const [a, b] = from < to ? [from, to] : [to, from];
        const range = visibleFlat.slice(a, b + 1).map((it) => it.path);
        setSelectedPaths(toggle ? Array.from(new Set(selectedPaths.concat(range))) : range);
        return;
      }
    }
    anchorRef.current = item.path;
    if (toggle) {
      setSelectedPaths((prev) =>
        prev.includes(item.path) ? prev.filter((p) => p !== item.path) : prev.concat(item.path)
      );
      return;
    }
    setSelectedPaths([item.path]);
  };

  if (!rootPath) {
    return (
      <div className="flex flex-1 flex-col items-center justify-center gap-3 p-6 text-center">
        <Files size={30} className="text-ws-faint" />
        <p className="text-xs text-ws-muted">Рабочая папка не выбрана</p>
        <button
          onClick={onOpenFolder}
          className="rounded bg-ws-accent px-3 py-1.5 text-xs font-medium text-white transition-opacity hover:opacity-90"
        >
          Открыть папку
        </button>
      </div>
    );
  }

  const startBulkRename = (paths: string[]) => {
    if (paths.length === 0) return;
    setBulk({ open: true, pattern: '{name}', paths });
  };

  const applyBulkRename = async () => {
    const paths = bulk.paths;
    const pattern = bulk.pattern;
    setBulk({ open: false, pattern: '{name}', paths: [] });
    if (!onRenamePath || paths.length === 0) return;
    const names = paths.map((path) => path.split(/[\\/]/).pop() ?? path);
    const pairs = applyRenamePattern(names, pattern);
    let index = 0;
    for (const path of paths) {
      const pair = pairs[index];
      index += 1;
      if (!pair || !pair.to || pair.to === pair.from) continue;
      await onRenamePath(path, pair.to);
    }
  };

  const renderItems = (dirPath: string, depth: number): React.ReactNode => {
    const items = tree[dirPath];
    if (!items) return null;

    const visible = childItems(tree, dirPath, viewOptions);
    if (visible.length === 0) {
      return (
        <div className="px-2 py-1 text-[11px] text-ws-faint" style={{ paddingLeft: depth * 12 + 16 }}>
          {filter ? 'Нет совпадений' : 'Папка пуста'}
        </div>
      );
    }

    return visible.map((item) => {
      const isExpanded = expandedDirs.has(item.path);
      const isLoading = loadingDirs.has(item.path);
      const isActive = item.path === activePath;
      const isDirty = dirtyPaths.has(item.path);
      const isSelected = selectedPaths.includes(item.path);

      return (
        <div
          key={item.path}
          // Большие папки: строки вне вьюпорта не раскладываются и не рисуются.
          style={{ contentVisibility: 'auto', containIntrinsicSize: '24px' }}
        >
          <div
            ref={(element) => {
              const key = normalizePath(item.path);
              if (element) rowRefs.current.set(key, element);
              else rowRefs.current.delete(key);
            }}
            draggable
            onDragStart={(event) => {
              const paths = selectedPaths.includes(item.path) ? selectedPaths : [item.path];
              event.dataTransfer.setData('application/x-workspace-paths', JSON.stringify(paths));
              event.dataTransfer.setData('text/plain', paths.join('\n'));
              event.dataTransfer.effectAllowed = 'copyMove';
            }}
            onDragOver={(event) => {
              if (!item.isDirectory) return;
              event.preventDefault();
              event.dataTransfer.dropEffect = event.ctrlKey ? 'copy' : 'move';
              if (dropTarget !== item.path) setDropTarget(item.path);
            }}
            onDragLeave={() => {
              if (dropTarget === item.path) setDropTarget('');
            }}
            onDrop={(event) => {
              if (!item.isDirectory) return;
              event.preventDefault();
              event.stopPropagation();
              void handleDropOn(item);
            }}
            onClick={(e) => {
              setFocusedPath(item.path);
              if (e.shiftKey || e.ctrlKey || e.metaKey) {
                selectItem(item, e);
                return;
              }
              selectItem(item, e);
              if (item.isDirectory) onToggleDir(item.path);
              else onPreviewFile?.(item.path);
            }}
            onDoubleClick={(e) => {
              if (item.isDirectory) return;
              e.preventDefault();
              onOpenFile(item.path);
            }}
            onContextMenu={(e) => {
              e.preventDefault();
              e.stopPropagation();
              if (!selectedPaths.includes(item.path)) {
                anchorRef.current = item.path;
                setSelectedPaths([item.path]);
              }
              setMenu({ x: e.clientX, y: e.clientY, item });
            }}
            title={`${item.name}${item.size ? ` — ${formatBytes(item.size)}` : ''}`}
            style={{ paddingLeft: depth * 12 + 6 }}
            className={cn(
              'group flex cursor-pointer items-center gap-1.5 rounded py-[3px] pr-2 text-xs transition-colors',
              isSelected
                ? 'bg-ws-accent-soft text-ws ring-1 ring-inset ring-ws-accent/40'
                : isActive
                  ? 'bg-ws-hover text-ws-accent'
                  : 'text-ws-muted hover:bg-ws-hover hover:text-ws'
            )}
          >
            {item.isDirectory ? (
              <>
                {isLoading ? (
                  <Loader2 size={12} className="shrink-0 animate-spin text-ws-faint" />
                ) : isExpanded ? (
                  <ChevronDown size={12} className="shrink-0 text-ws-faint" />
                ) : (
                  <ChevronRight size={12} className="shrink-0 text-ws-faint" />
                )}
                {isExpanded ? (
                  <FolderOpen size={14} className="shrink-0 text-ws-accent" />
                ) : (
                  <Folder size={14} className="shrink-0 text-ws-accent" />
                )}
              </>
            ) : (
              <>
                <span className="w-3 shrink-0" />
                {item.type === 'image' && explorerPrefs.showThumbs ? (
                  <FileThumb item={item} />
                ) : (
                  fileIcon(item.type, item.name)
                )}
              </>
            )}

            {renaming?.path === item.path ? (
              <input
                autoFocus
                value={renaming.value}
                onChange={(e) => setRenaming({ path: item.path, value: e.target.value })}
                onBlur={() => void commitRename()}
                onKeyDown={(e) => {
                  e.stopPropagation();
                  if (e.key === 'Enter') void commitRename();
                  else if (e.key === 'Escape') setRenaming(null);
                }}
                className="min-w-0 flex-1 rounded border border-ws-accent bg-ws-elevated px-1 py-0.5 text-xs text-ws focus:outline-none"
              />
            ) : (
              <span
                className={cn('truncate', clipboard?.mode === 'cut' && clipboard.paths.includes(item.path) && 'opacity-50')}
                onDoubleClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  startRename(item);
                }}
              >
                {item.name}
              </span>
            )}
            {explorerPrefs.showMeta && !item.isDirectory && (
              <span className="ml-auto shrink-0 truncate pl-2 text-[10px] text-ws-faint">
                {formatBytes(item.size)}
                {item.mtime ? ` · ${formatDate(item.mtime)}` : ''}
              </span>
            )}
            {isDirty && <span className="ml-auto shrink-0 text-[14px] leading-none text-ws-warning">•</span>}
          </div>

          {item.isDirectory && isExpanded && renderItems(item.path, depth + 1)}
        </div>
      );
    });
  };

  return (
    <div className="flex flex-1 flex-col overflow-hidden">
      {/* Заголовок панели */}
      <div className="flex items-center justify-between border-b border-ws px-3 py-2">
        <span className="text-[11px] font-semibold tracking-wider text-ws-muted uppercase">Проводник</span>
        <div className="flex items-center gap-0.5">
          <button onClick={onOpenFolder} title="Открыть папку" className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"><FolderInput size={13} /></button>
          <button onClick={() => onToggleFavorite?.(rootPath)} title="Добавить текущую папку в избранное" className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"><Star size={13} className={favorites.some((f) => f.path === rootPath) ? 'fill-amber-400 text-amber-400' : ''} /></button>
          <button onClick={onOpenGoogleDriveSettings} title="Настроить Google Drive" className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"><CloudUpload size={13} /></button>
          <button onClick={() => onItemAction('upload-drive', { name: rootPath.split(/[\\/]/).pop() || rootPath, path: rootPath, isDirectory: true, size: 0, mtime: null, extension: '' })} title="Загрузить рабочую папку в Google Drive" className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"><Share2 size={13} /></button>
          <button
            onClick={() => onNewFile(rootPath)}
            title="Новый файл в корне"
            className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"
          >
            <FilePlus2 size={13} />
          </button>
          <button
            onClick={() => onNewFolder(rootPath)}
            title="Новая папка в корне"
            className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"
          >
            <FolderPlus size={13} />
          </button>
          <button
            onClick={onRefresh}
            title="Обновить дерево файлов"
            className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"
          >
            <RefreshCw size={13} />
          </button>
        </div>
      </div>

      {favorites.length > 0 && (
        <div className="border-b border-ws px-2 py-1.5">
          <div className="mb-1 px-1 text-[10px] font-semibold tracking-wider text-ws-faint uppercase">Избранное</div>
          <div className="max-h-28 space-y-0.5 overflow-y-auto">
            {favorites.map((fav) => (
              <div key={fav.path} className={cn('group flex items-center rounded pr-1', fav.path === rootPath ? 'bg-ws-accent-soft text-ws-accent' : 'text-ws-muted hover:bg-ws-hover hover:text-ws')}>
                <button onClick={() => onOpenFavorite?.(fav.path)} title={fav.path} className="flex min-w-0 flex-1 items-center gap-1.5 px-1.5 py-1 text-left text-[11px]">
                  <Star size={11} className="shrink-0 fill-amber-400 text-amber-400" />
                  <span className="truncate">{fav.name}</span>
                </button>
                <button onClick={() => onToggleFavorite?.(fav.path)} title="Убрать из избранного" className="rounded p-1 text-ws-faint opacity-0 transition-opacity group-hover:opacity-100 hover:text-ws-danger">
                  <X size={11} />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Фильтр по имени */}
      <div className="border-b border-ws px-2 py-1.5">
        <input
          type="text"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
          placeholder="Фильтр файлов..."
          className="w-full rounded border border-ws bg-ws-elevated px-2 py-1 text-[11px] text-ws placeholder:text-ws-faint focus:border-ws-accent focus:outline-none"
        />
      </div>

      {/* Массовое переименование по шаблону */}
      {bulk.open && (
        <div className="border-b border-ws bg-ws-elevated px-2 py-1.5 text-[11px]">
          <div className="mb-1 text-ws-muted">Шаблон: {'{name}'} — имя, {'{n}'} — номер, {'{ext}'} — расширение</div>
          <input
            autoFocus
            value={bulk.pattern}
            onChange={(e) => setBulk({ ...bulk, pattern: e.target.value })}
            onKeyDown={(e) => {
              e.stopPropagation();
              if (e.key === 'Enter') void applyBulkRename();
              else if (e.key === 'Escape') setBulk({ open: false, pattern: '{name}', paths: [] });
            }}
            className="w-full rounded border border-ws-accent bg-ws-surface px-2 py-1 font-mono text-[11px] text-ws focus:outline-none"
          />
          <div className="mt-1 max-h-16 overflow-y-auto text-[10px] text-ws-faint">
            {applyRenamePattern(
              bulk.paths.map((path) => path.split(/[\\/]/).pop() ?? path),
              bulk.pattern
            )
              .slice(0, 5)
              .map((pair) => (
                <div key={pair.from} className="truncate">
                  {pair.from} → {pair.to}
                </div>
              ))}
          </div>
          <div className="mt-1 flex justify-end gap-2">
            <button
              onClick={() => setBulk({ open: false, pattern: '{name}', paths: [] })}
              className="rounded px-2 py-0.5 text-ws-muted hover:bg-ws-hover"
            >
              Отмена
            </button>
            <button onClick={() => void applyBulkRename()} className="rounded bg-ws-accent px-2 py-0.5 text-white">
              Переименовать ({bulk.paths.length})
            </button>
          </div>
        </div>
      )}

      {/* Навигация по пути (Ctrl+L) */}
      {pathEdit.open && (
        <div className="border-b border-ws bg-ws-elevated px-2 py-1">
          <input
            autoFocus
            value={pathEdit.value}
            onChange={(e) => setPathEdit({ open: true, value: e.target.value })}
            onKeyDown={(e) => {
              e.stopPropagation();
              if (e.key === 'Enter') {
                const target = pathEdit.value.trim();
                setPathEdit({ open: false, value: '' });
                if (target) onOpenFavorite?.(target);
              } else if (e.key === 'Escape') {
                setPathEdit({ open: false, value: '' });
              }
            }}
            placeholder="Путь к папке и Enter"
            className="w-full rounded border border-ws-accent bg-ws-surface px-2 py-1 font-mono text-[11px] text-ws focus:outline-none"
          />
        </div>
      )}

      {/* Панель вида: путь, сортировка, скрытые, отмена */}
      <div className="flex items-center gap-1 border-b border-ws px-2 py-1 text-[10px] text-ws-faint">
        <button
          onClick={() => {
            const parent = rootPath.replace(/[\\/][^\\/]+$/, '');
            if (parent && parent !== rootPath) onOpenFavorite?.(parent);
          }}
          title="На уровень выше"
          className="rounded px-1 py-0.5 hover:bg-ws-hover hover:text-ws"
        >
          ↑
        </button>
        <span className="min-w-0 flex-1 truncate" title={rootPath}>
          {rootPath.split(/[\\/]/).pop() || rootPath}
        </span>
        <button
          onClick={() => setPathEdit({ open: true, value: rootPath })}
          title="Изменить путь (Ctrl+L)"
          className="rounded px-1 py-0.5 hover:bg-ws-hover hover:text-ws"
        >
          …
        </button>
        <button
          onClick={() => updateExplorerPrefs({ dir: explorerPrefs.dir === 'asc' ? 'desc' : 'asc' })}
          title="Направление сортировки"
          className="rounded px-1 py-0.5 hover:bg-ws-hover hover:text-ws"
        >
          {explorerPrefs.dir === 'asc' ? 'A↓' : 'A↑'}
        </button>
        <select
          value={explorerPrefs.sortBy}
          onChange={(e) => updateExplorerPrefs({ sortBy: e.target.value as typeof explorerPrefs.sortBy })}
          title="Сортировка"
          className="rounded border border-ws bg-ws-elevated px-1 py-0.5 text-[10px] text-ws-muted"
        >
          <option value="name">имя</option>
          <option value="mtime">дата</option>
          <option value="size">размер</option>
          <option value="type">тип</option>
        </select>
        <button
          onClick={() => updateExplorerPrefs({ showHidden: !explorerPrefs.showHidden })}
          title={explorerPrefs.showHidden ? 'Скрыть служебные файлы' : 'Показать скрытые файлы'}
          className={cn(
            'rounded px-1 py-0.5 hover:bg-ws-hover hover:text-ws',
            explorerPrefs.showHidden && 'text-ws-accent'
          )}
        >
          {explorerPrefs.showHidden ? '◉' : '◎'}
        </button>
        <button
          onClick={() => updateExplorerPrefs({ showThumbs: !explorerPrefs.showThumbs })}
          title={explorerPrefs.showThumbs ? 'Скрыть миниатюры' : 'Показать миниатюры изображений'}
          className={cn(
            'rounded px-1 py-0.5 hover:bg-ws-hover hover:text-ws',
            explorerPrefs.showThumbs && 'text-ws-accent'
          )}
        >
          {explorerPrefs.showThumbs ? '▣' : '▢'}
        </button>
        <button
          onClick={() => void undoLast()}
          disabled={historySize === 0}
          title="Отменить перемещение или переименование (Ctrl+Z)"
          className="rounded px-1 py-0.5 hover:bg-ws-hover hover:text-ws disabled:opacity-30"
        >
          ↺
        </button>
      </div>

      {/* Дерево */}
      <div
        tabIndex={0}
        onKeyDown={handleTreeKeys}
        onDragOver={(event) => {
          if (!event.dataTransfer.types.includes('application/x-workspace-paths')) return;
          event.preventDefault();
          if (dropTarget !== rootPath) setDropTarget(rootPath);
        }}
        onDragLeave={() => setDropTarget('')}
        onDrop={(event) => {
          if (event.dataTransfer.types.includes('application/x-workspace-paths')) {
            event.preventDefault();
            void handleDropOn(null);
          }
        }}
        className={cn(
          'flex-1 overflow-y-auto py-1 outline-none',
          dropTarget === rootPath && 'ring-1 ring-inset ring-ws-accent/50'
        )}
      >
        {renderItems(rootPath, 0)}
      </div>
      {menu && (() => {
        const image = menu.item.type === 'image';
        const chosen = selectedItems.some((it) => it.path === menu.item.path)
          ? selectedItems
          : [menu.item];
        const images = chosen.filter((it) => it.type === 'image' && !it.isDirectory);
        const items: ContextMenuItem[] = [
          { id: 'open', label: menu.item.isDirectory ? 'Развернуть / свернуть' : 'Открыть для редактирования', icon: <ExternalLink size={13} /> },
          ...(!menu.item.isDirectory ? [
            { id: 'preview', label: 'Предпросмотр', icon: <FileSearch size={13} /> },
            { id: 'open-new', label: 'Открыть в новой вкладке', icon: <Files size={13} /> },
            { id: 'open-secondary', label: 'Открыть во второй области', icon: <ExternalLink size={13} /> },
            { id: 'open-external', label: 'Открыть в системе', icon: <ExternalLink size={13} /> },
          ] : [
            { id: 'new-file', label: 'Новый файл', icon: <FilePlus2 size={13} /> },
            { id: 'new-folder', label: 'Новая папка', icon: <FolderPlus size={13} /> },
            { id: 'favorite', label: favorites.some((f) => f.path === menu.item.path) ? 'Убрать из избранного' : 'В избранное', icon: <Star size={13} /> },
          ]),
          { id: 'rename', label: 'Переименовать', icon: <Pencil size={13} />, separatorBefore: true },
          { id: 'copy-path', label: 'Копировать путь', icon: <Copy size={13} /> },
          { id: 'search-folder', label: 'Искать в этой папке', icon: <FileSearch size={13} /> },
          ...(chosen.length > 1
            ? [{ id: 'bulk-rename', label: 'Переименовать по шаблону (' + chosen.length + ')', icon: <Pencil size={13} /> }]
            : []),
          { id: 'cut', label: 'Вырезать', icon: <Copy size={13} /> },
          { id: 'copy', label: 'Копировать', icon: <Copy size={13} /> },
          ...(clipboard ? [{ id: 'paste', label: 'Вставить в эту папку', icon: <Copy size={13} /> }] : []),
          { id: 'reveal', label: 'Показать в проводнике', icon: <RevealIcon size={13} />, separatorBefore: true },
          ...(image && images.length < 2 ? SHARE_SERVICES.map((s) => ({ id: 'share:' + s.id, label: 'Экспорт: ' + s.label, icon: <Share2 size={13} /> })) : []),
          ...(images.length > 0 ? [
            { id: 'image-process', label: 'Сжать для веба (' + images.length + ')', icon: <SlidersHorizontal size={13} />, separatorBefore: true },
            { id: 'image-process:thumb', label: 'Миниатюры (' + images.length + ')', icon: <SlidersHorizontal size={13} /> },
          ] : []),
          ...(images.length > 1 ? [
            { id: 'album:imgbb', label: 'ImgBB (' + images.length + ')', icon: <Share2 size={13} />, separatorBefore: true },
            { id: 'album:freeimage', label: 'Freeimage.host (' + images.length + ')', icon: <Share2 size={13} /> },
            { id: 'album:yapx', label: 'Япикс, альбом (' + images.length + ')', icon: <Share2 size={13} /> },
            { id: 'album:telegraph', label: 'Страница Telegraph (' + images.length + ')', icon: <Share2 size={13} /> },
            { id: 'album:catbox', label: 'Catbox без регистрации (' + images.length + ')', icon: <Share2 size={13} /> },
            { id: 'album:imgur', label: 'Альбом Imgur (' + images.length + ')', icon: <Share2 size={13} /> },
            { id: 'album:gdrive', label: 'Папка Google Drive (' + images.length + ')', icon: <CloudUpload size={13} /> },
          ] : []),
          { id: 'upload-drive', label: menu.item.isDirectory ? 'Загрузить папку в Google Drive' : 'Загрузить в Google Drive', icon: <CloudUpload size={13} />, separatorBefore: true },
          { id: 'delete', label: 'Удалить', icon: <Trash2 size={13} />, danger: true, separatorBefore: true },
        ];
        return <ContextMenu x={menu.x} y={menu.y} items={items} onClose={() => setMenu(null)} onSelect={(id) => {
          if (id.startsWith('share:')) onItemAction('share', menu.item, id.slice(6), chosen);
          else if (id.startsWith('album:')) onItemAction('album', menu.item, id.slice(6), images);
          else if (id === 'bulk-rename') startBulkRename(chosen.map((it) => it.path));
          else if (id === 'cut') setClipboard({ mode: 'cut', paths: chosen.map((it) => it.path) });
          else if (id === 'copy') setClipboard({ mode: 'copy', paths: chosen.map((it) => it.path) });
          else if (id === 'paste') {
            setFocusedPath(menu.item.isDirectory ? menu.item.path : focusedPath);
            void runPaste();
          }
          else if (id.startsWith('image-process')) {
            onItemAction('image-process', menu.item, id.includes(':') ? 'thumb' : 'web', images);
          }
          else onItemAction(id, menu.item, undefined, chosen);
        }} />;
      })()}
    </div>
  );
};
