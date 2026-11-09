import React, { useEffect, useState } from 'react';
import {
  Minus,
  Square,
  Copy,
  X,
  FolderOpen,
  FilePlus2,
  Save,
  Download,
  Moon,
  Sun,
  Columns2,
  PanelLeft,
  PanelRight,
  Settings,
  Sparkles,
  FileSearch,
  Terminal,
  RefreshCw,
} from 'lucide-react';
import { cn } from '../../lib/utils';

interface Props {
  workspacePath: string;
  onOpenFolder: () => void;
  onNewFile: () => void;
  onSave: () => void;
  onSaveAs: () => void;
  onToggleSidebar: () => void;
  onToggleAIPanel: () => void;
  onToggleTerminal?: () => void;
  onToggleSplit: () => void;
  onOpenSettings: () => void;
  onCheckUpdates: () => void;
  onQuickOpen: () => void;
  sidebarOpen: boolean;
  aiPanelOpen: boolean;
  terminalOpen?: boolean;
  splitEnabled: boolean;
  isDark: boolean;
  onToggleTheme: () => void;
  isElectron: boolean;
}

export const TitleBar: React.FC<Props> = ({
  workspacePath,
  onOpenFolder,
  onNewFile,
  onSave,
  onSaveAs,
  onToggleSidebar,
  onToggleAIPanel,
  onToggleTerminal,
  onToggleSplit,
  onOpenSettings,
  onCheckUpdates,
  onQuickOpen,
  sidebarOpen,
  aiPanelOpen,
  terminalOpen = false,
  splitEnabled,
  isDark,
  onToggleTheme,
  isElectron,
}) => {
  const [isMaximized, setIsMaximized] = useState(false);

  useEffect(() => {
    if (!isElectron) return;
    void window.electronAPI?.isWindowMaximized().then(setIsMaximized);
    return window.electronAPI?.onWindowMaximizeChange?.(setIsMaximized);
  }, [isElectron]);

  const folderLabel = workspacePath
    ? workspacePath.split(/[\\/]/).filter(Boolean).pop() || workspacePath
    : 'Папка не выбрана';

  return (
    <div
      className={cn(
        'flex h-10 shrink-0 items-center justify-between border-b border-ws bg-ws-titlebar px-2 select-none',
        isElectron && 'titlebar-drag'
      )}
    >
      {/* Левая часть: логотип, рабочая папка, быстрые действия */}
      <div className="titlebar-no-drag flex items-center gap-1">
        <div className="flex items-center gap-2 pr-2 pl-1">
          <div className="flex h-5 w-5 items-center justify-center rounded bg-ws-accent text-[11px] font-bold text-white">
            W
          </div>
          <span className="text-xs font-semibold tracking-wide text-ws">WorkSpace</span>
        </div>

        <div className="mx-1 h-4 w-px bg-ws border-l border-ws" />

        <button
          onClick={onOpenFolder}
          title="Открыть рабочую папку"
          className="flex max-w-[240px] items-center gap-1.5 rounded px-2 py-1 text-xs text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
        >
          <FolderOpen size={13} className="shrink-0 text-ws-accent" />
          <span className="truncate">{folderLabel}</span>
        </button>

        <div className="mx-1 h-4 w-px border-l border-ws" />

        <button
          onClick={onNewFile}
          title="Новый файл"
          className="rounded p-1.5 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
        >
          <FilePlus2 size={14} />
        </button>
        <button
          onClick={onSave}
          title="Сохранить активный файл (Ctrl+S)"
          className="rounded p-1.5 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
        >
          <Save size={14} />
        </button>
        <button
          onClick={onSaveAs}
          title="Сохранить как… (Ctrl+Shift+S)"
          className="rounded p-1.5 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
        >
          <Download size={14} />
        </button>
        <button
          onClick={onQuickOpen}
          title="Быстрый переход к файлу (Ctrl+P)"
          className="rounded p-1.5 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
        >
          <FileSearch size={14} />
        </button>
      </div>

      {/* Правая часть: переключатели панелей и управление окном */}
      <div className="titlebar-no-drag flex items-center gap-0.5">
        <button
          onClick={onToggleSidebar}
          title="Проводник (Ctrl+B)"
          className={cn(
            'rounded p-1.5 transition-colors hover:bg-ws-hover',
            sidebarOpen ? 'text-ws-accent' : 'text-ws-faint'
          )}
        >
          <PanelLeft size={14} />
        </button>
        <button
          onClick={onToggleSplit}
          title="Разделить редактор на две области"
          className={cn(
            'rounded p-1.5 transition-colors hover:bg-ws-hover',
            splitEnabled ? 'text-ws-accent' : 'text-ws-faint'
          )}
        >
          <Columns2 size={14} />
        </button>
        <button
          onClick={onToggleAIPanel}
          title="AI-ассистент"
          className={cn(
            'rounded p-1.5 transition-colors hover:bg-ws-hover',
            aiPanelOpen ? 'text-ws-accent' : 'text-ws-faint'
          )}
        >
          <Sparkles size={14} />
        </button>
        {onToggleTerminal && (
          <button
            onClick={onToggleTerminal}
            title="Встроенный терминал (Ctrl+`)"
            className={cn(
              'rounded p-1.5 transition-colors hover:bg-ws-hover',
              terminalOpen ? 'text-emerald-500' : 'text-ws-faint'
            )}
          >
            <Terminal size={14} />
          </button>
        )}

        <div className="mx-1 h-4 w-px border-l border-ws" />

        <button
          onClick={onToggleTheme}
          title={isDark ? 'Светлая тема' : 'Тёмная тема'}
          className="rounded p-1.5 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
        >
          {isDark ? <Sun size={14} /> : <Moon size={14} />}
        </button>
        <button
          onClick={onCheckUpdates}
          title="Обновления WorkSpace"
          className="rounded p-1.5 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
        >
          <RefreshCw size={14} />
        </button>
        <button
          onClick={onOpenSettings}
          title="Настройки"
          className="rounded p-1.5 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
        >
          <Settings size={14} />
        </button>

        {isElectron && (
          <>
            <div className="mx-1 h-4 w-px border-l border-ws" />
            <button
              onClick={() => window.electronAPI?.minimizeWindow()}
              title="Свернуть"
              className="rounded p-1.5 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
            >
              <Minus size={14} />
            </button>
            <button
              onClick={() => window.electronAPI?.maximizeWindow()}
              title={isMaximized ? 'Восстановить' : 'Развернуть'}
              className="rounded p-1.5 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws"
            >
              {isMaximized ? <Copy size={12} /> : <Square size={12} />}
            </button>
            <button
              onClick={() => window.electronAPI?.closeWindow()}
              title="Закрыть"
              className="rounded p-1.5 text-ws-muted transition-colors hover:bg-red-600 hover:text-white"
            >
              <X size={14} />
            </button>
          </>
        )}
      </div>
    </div>
  );
};
