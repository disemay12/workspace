import React, { useEffect, useRef, useState, useCallback } from 'react';
import { Terminal } from '@xterm/xterm';
import { FitAddon } from '@xterm/addon-fit';
import '@xterm/xterm/css/xterm.css';
import {
  Terminal as TerminalIcon,
  Trash2,
  RotateCw,
  X,
  Maximize2,
  Minimize2,
} from 'lucide-react';
import { cn } from '../../lib/utils';
import { IconButton } from '../ui';

interface Props {
  cwd?: string;
  isDark?: boolean;
  onClose: () => void;
  height?: number;
  onHeightChange?: (height: number) => void;
  /** Команда, отправленная ИИ: выполняется, как только сессия готова. */
  command?: { id: string; command: string } | null;
  onCommandConsumed?: () => void;
}

export const TerminalPanel: React.FC<Props> = ({
  cwd,
  isDark = true,
  onClose,
  height = 240,
  onHeightChange,
  command,
  onCommandConsumed,
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const terminalRef = useRef<Terminal | null>(null);
  const fitAddonRef = useRef<FitAddon | null>(null);
  const sentCommandRef = useRef('');
  const sessionIdRef = useRef<string>('term_default');
  const [isMaximized, setIsMaximized] = useState(false);
  const [resizing, setResizing] = useState(false);
  const [ready, setReady] = useState(false);
  const startYRef = useRef(0);
  const startHeightRef = useRef(height);

  /* ---------------- Инициализация сессии терминала ---------------- */
  const initTerminal = useCallback(async () => {
    if (!containerRef.current) return;
    setReady(false);

    // Очищаем предыдущий терминал если был
    if (terminalRef.current) {
      terminalRef.current.dispose();
      terminalRef.current = null;
    }
    containerRef.current.innerHTML = '';

    const term = new Terminal({
      cursorBlink: true,
      fontSize: 12,
      fontFamily: 'Consolas, "Courier New", monospace',
      theme: isDark
        ? {
            background: '#0b0f17',
            foreground: '#e2e8f0',
            cursor: '#10b981',
            selectionBackground: 'rgba(16, 185, 129, 0.3)',
            black: '#1e293b',
            red: '#ef4444',
            green: '#10b981',
            yellow: '#f59e0b',
            blue: '#3b82f6',
            magenta: '#a855f7',
            cyan: '#06b6d4',
            white: '#f8fafc',
          }
        : {
            background: '#f8fafc',
            foreground: '#0f172a',
            cursor: '#10b981',
            selectionBackground: 'rgba(16, 185, 129, 0.3)',
            black: '#0f172a',
            red: '#dc2626',
            green: '#16a34a',
            yellow: '#d97706',
            blue: '#2563eb',
            magenta: '#9333ea',
            cyan: '#0891b2',
            white: '#ffffff',
          },
      convertEol: true,
    });

    const fitAddon = new FitAddon();
    term.loadAddon(fitAddon);

    term.open(containerRef.current);
    try {
      fitAddon.fit();
    } catch {
      // ignore
    }

    terminalRef.current = term;
    fitAddonRef.current = fitAddon;

    const id = sessionIdRef.current;
    if (window.electronAPI?.terminalCreate) {
      await window.electronAPI.terminalCreate({ id, cwd });
    }

    setReady(true);

    term.onData((data) => {
      window.electronAPI?.terminalInput?.({ id, data });
    });

    term.focus();
  }, [cwd, isDark]);

  useEffect(() => {
    void initTerminal();

    const unsubData = window.electronAPI?.onTerminalData?.((data) => {
      if (data.id === sessionIdRef.current && terminalRef.current) {
        terminalRef.current.write(data.data);
      }
    });

    const unsubExit = window.electronAPI?.onTerminalExit?.((data) => {
      if (data.id === sessionIdRef.current && terminalRef.current) {
        terminalRef.current.write(`\r\n[Процесс завершился с кодом ${data.code}]\r\n`);
      }
    });

    const handleResize = () => {
      try {
        fitAddonRef.current?.fit();
      } catch {
        // ignore
      }
    };
    window.addEventListener('resize', handleResize);

    return () => {
      unsubData?.();
      unsubExit?.();
      window.removeEventListener('resize', handleResize);
      if (terminalRef.current) {
        terminalRef.current.dispose();
        terminalRef.current = null;
      }
    };
  }, [initTerminal]);

  useEffect(() => {
    if (!command || !ready) return;
    // StrictMode запускает эффект дважды: команда должна уйти ровно один раз
    if (sentCommandRef.current === command.id) return;
    sentCommandRef.current = command.id;
    window.electronAPI?.terminalInput?.({ id: sessionIdRef.current, data: command.command + '\r' });
    onCommandConsumed?.();
  }, [command, ready, onCommandConsumed]);

  /* ---------------- Изменение размера тягой за край ---------------- */
  const onMouseDownResize = (e: React.MouseEvent) => {
    e.preventDefault();
    setResizing(true);
    startYRef.current = e.clientY;
    startHeightRef.current = height;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const delta = startYRef.current - moveEvent.clientY;
      const next = Math.max(120, Math.min(window.innerHeight - 150, startHeightRef.current + delta));
      onHeightChange?.(next);
      try {
        fitAddonRef.current?.fit();
      } catch {
        // ignore
      }
    };

    const onMouseUp = () => {
      setResizing(false);
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      try {
        fitAddonRef.current?.fit();
      } catch {
        // ignore
      }
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  const handleClear = () => {
    terminalRef.current?.clear();
  };

  const handleRestart = () => {
    void initTerminal();
  };

  const panelHeight = isMaximized ? 'calc(100vh - 120px)' : `${height}px`;

  return (
    <div
      style={{ height: panelHeight }}
      className={cn(
        'relative flex flex-col border-t border-ws bg-ws-surface transition-all select-none',
        resizing && 'cursor-ns-resize'
      )}
    >
      {/* Ресайз-хэндлер */}
      {!isMaximized && (
        <div
          onMouseDown={onMouseDownResize}
          className="group absolute -top-1.5 left-0 right-0 z-20 h-3 cursor-ns-resize"
          title="Потяните для изменения высоты терминала"
        >
          <div className="mx-auto h-1 w-12 rounded-full bg-ws-border transition-colors group-hover:bg-ws-accent" />
        </div>
      )}

      {/* Шапка терминала */}
      <div className="flex shrink-0 items-center justify-between border-b border-ws bg-ws-elevated px-3 py-1.5 text-xs">
        <div className="flex items-center gap-2">
          <TerminalIcon size={14} className="text-emerald-500" />
          <span className="font-medium text-ws">Терминал</span>
          <span className="rounded bg-ws-accent-soft px-1.5 py-0.2 text-[10px] font-mono text-ws-accent">
            PowerShell
          </span>
          {cwd && (
            <span className="max-w-[320px] truncate font-mono text-[11px] text-ws-faint" title={cwd}>
              {cwd}
            </span>
          )}
        </div>

        <div className="flex items-center gap-1">
          <IconButton title="Очистить экран" onClick={handleClear}>
            <Trash2 size={13} />
          </IconButton>
          <IconButton title="Перезапустить терминал" onClick={handleRestart}>
            <RotateCw size={13} />
          </IconButton>
          <IconButton
            title={isMaximized ? 'Свернуть' : 'Развернуть'}
            onClick={() => {
              setIsMaximized((v) => !v);
              setTimeout(() => fitAddonRef.current?.fit(), 100);
            }}
          >
            {isMaximized ? <Minimize2 size={13} /> : <Maximize2 size={13} />}
          </IconButton>
          <div className="mx-1 h-3.5 w-px border-l border-ws" />
          <IconButton title="Закрыть терминал (Ctrl+`)" onClick={onClose}>
            <X size={13} />
          </IconButton>
        </div>
      </div>

      {/* Окно xterm */}
      <div
        ref={containerRef}
        className="flex-1 overflow-hidden p-2 font-mono text-xs select-text"
        onClick={() => terminalRef.current?.focus()}
      />
    </div>
  );
};
