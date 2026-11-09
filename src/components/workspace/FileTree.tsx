import React, { useMemo, useRef, useState } from 'react';
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
} from 'lucide-react';
import { FileItem, FileType } from '../../types';
import { cn, formatBytes } from '../../lib/utils';
import { ContextMenu, ContextMenuItem } from './ContextMenu';
import { SHARE_SERVICES } from '../../services/shareService';

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
}) => {
  const [filter, setFilter] = useState('');
  const [menu, setMenu] = useState<{ x: number; y: number; item: FileItem } | null>(null);
  const [selectedPaths, setSelectedPaths] = useState<string[]>([]);
  const anchorRef = useRef('');

  /** Рекурсивная фильтрация дерева по подстроке имени. */
  const matchesFilter = (item: FileItem): boolean => {
    if (!filter.trim()) return true;
    const q = filter.toLowerCase();
    if (item.name.toLowerCase().includes(q)) return true;
    if (item.isDirectory) {
      const children = tree[item.path] || [];
      return children.some(matchesFilter);
    }
    return false;
  };

  const rootItems = useMemo(() => tree[rootPath] || [], [tree, rootPath]);

  /** Плоский список видимых строк — нужен для выделения диапазона Shift-кликом. */
  const visibleFlat = useMemo(() => {
    const out: FileItem[] = [];
    const walk = (dir: string) => {
      const items = (tree[dir] || []).filter(matchesFilter);
      for (const item of items) {
        out.push(item);
        if (item.isDirectory && expandedDirs.has(item.path)) walk(item.path);
      }
    };
    if (rootPath) walk(rootPath);
    return out;
    // matchesFilter зависит от filter и tree, они уже в зависимостях через замыкание
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tree, rootPath, expandedDirs, filter]);

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

  const renderItems = (dirPath: string, depth: number): React.ReactNode => {
    const items = tree[dirPath];
    if (!items) return null;

    const visible = items.filter(matchesFilter);
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
        <div key={item.path}>
          <div
            draggable={!item.isDirectory}
            onDragStart={(event) => {
              if (item.isDirectory) return;
              event.dataTransfer.setData('application/x-workspace-path', item.path);
              event.dataTransfer.setData('text/plain', item.path);
              event.dataTransfer.effectAllowed = 'copy';
            }}
            onClick={(e) => {
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
                {fileIcon(item.type, item.name)}
              </>
            )}

            <span className="truncate">{item.name}</span>
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

      {/* Дерево */}
      <div className="flex-1 overflow-y-auto py-1">{renderItems(rootPath, 0)}</div>
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
          { id: 'reveal', label: 'Показать в проводнике', icon: <RevealIcon size={13} />, separatorBefore: true },
          ...(image && images.length < 2 ? SHARE_SERVICES.map((s) => ({ id: 'share:' + s.id, label: 'Экспорт: ' + s.label, icon: <Share2 size={13} /> })) : []),
          ...(images.length > 1 ? [
            { id: 'album:imgbb', label: 'ImgBB (' + images.length + ')', icon: <Share2 size={13} />, separatorBefore: true },
            { id: 'album:freeimage', label: 'Freeimage.host (' + images.length + ')', icon: <Share2 size={13} /> },
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
          else onItemAction(id, menu.item, undefined, chosen);
        }} />;
      })()}
    </div>
  );
};
