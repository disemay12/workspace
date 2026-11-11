import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  FilePlus2,
  FolderPlus,
  X,
  Files,
  Sparkles,
  Maximize2,
  Save,
  Pencil,
  Download,
  Eye,
} from 'lucide-react';

import { OpenTab, ChatMessage, AIProvider, FileItem, ChatAttachment, FileType } from './types';
import { fileSystemService, getFileType } from './services/fileSystemService';
import { getSheetController } from './services/sheetBridge';
import { getPptxController } from './services/pptBridge';
import { fontPresetById } from './lib/pptxThemes';
import { resolvePasteName } from './lib/explorerOps';
import { applyImageOps, normalizeImageOps } from './lib/imageOps';
import { markdownToPptxBase64 } from './lib/pptxBuilder';
import { getImageController } from './services/imageBridge';
import { getVideoController } from './services/videoBridge';
import { useWorkspace } from './state/workspace';
import { useAISettings, useChatSessions } from './state/settings';
import { useTheme } from './hooks/useTheme';
import { applyAppearance } from './lib/appearance';
import { layoutFingerprint, normalizeLayout, MAX_RESTORED_TABS } from './lib/layout';
import { playChimeFor } from './services/chime';
import { listSnapshots, pruneSnapshots, writeSnapshot } from './services/autosave';
import { copyDiagnosticsReport, installDiagnosticsHooks } from './services/diagnostics';
import type { NotificationKind } from './lib/notifyRules';
import { useLocalStorage } from './hooks/useLocalStorage';
import { cn, baseName, dirName, extName, joinPath, uid } from './lib/utils';
import { docxText, pptxText, tabToDocumentText, htmlToPlainText } from './lib/documentText';
import { extractLegacyDocText } from './lib/legacyDoc';
import { docxBase64ToHtml, htmlToDocxBase64 } from './lib/docx';
import { applyAutoFitColumns } from './lib/xlsxColumns';
import { resolveWorkspacePath as resolvePathInsideWorkspace } from './lib/workspacePath';
import { resolveImageSource } from './lib/imageSource';
import {
  getDocumentSession,
  kindForTab,
  listDocumentSessions,
  registerDocumentSession,
  removeDocumentSession,
  updateDocumentSession,
} from './lib/documentSession';
import { extractDocumentContent, markdownToDocumentHtml } from './lib/aiDocument';
import {
  AGENT_TAB_TYPES,
  applyDocumentTool,
  BLOCK_TOOL_TAB_TYPES,
  DOCUMENT_TAB_TYPES,
  DOCUMENT_WRITE_TOOLS,
  type DocumentToolResult,
} from './lib/documentModel';
import { getDocumentController } from './services/docBridge';
import { buildDocumentTools } from './services/aiTools';
import { DOCUMENT_TEMPLATES } from './lib/templates';
import {
  insertMarkdownLink,
  insertMarkdownTable,
  setMarkdownHeading,
  toggleMarkdownList,
} from './lib/markdownOps';

import { TitleBar } from './components/workspace/TitleBar';
import { FileTree } from './components/workspace/FileTree';
import { SaveAsDialog } from './components/workspace/SaveAsDialog';
import { QuickOpenModal } from './components/workspace/QuickOpenModal';
import { GlobalSearchPanel } from './components/workspace/GlobalSearchPanel';
import { SettingsModal } from './components/settings/SettingsModal';
import { AIChatPanel } from './components/ai/AIChatPanel';
import { InlineCopilot, InlineEditTarget } from './components/ai/InlineCopilot';
import { Button, Input, Modal, Spinner, Toast } from './components/ui';

import { MonacoEditorWrapper } from './components/editors/MonacoEditorWrapper';
import { RichTextEditor } from './components/editors/RichTextEditor';
import type { AIAction } from './components/editors/ExcelEditor';

// Тяжёлые редакторы грузятся по требованию: старт не ждёт Monaco/SheetJS/презентаций.
const ExcelEditor = React.lazy(() =>
  import('./components/editors/ExcelEditor').then((module) => ({ default: module.ExcelEditor }))
);
const WordEditor = React.lazy(() =>
  import('./components/editors/WordEditor').then((module) => ({ default: module.WordEditor }))
);
const PowerPointEditor = React.lazy(() =>
  import('./components/editors/PowerPointEditor').then((module) => ({ default: module.PowerPointEditor }))
);
const ImageViewer = React.lazy(() =>
  import('./components/editors/ImageViewer').then((module) => ({ default: module.ImageViewer }))
);
const VideoEditor = React.lazy(() =>
  import('./components/editors/VideoEditor').then((module) => ({ default: module.VideoEditor }))
);
const PdfEditor = React.lazy(() =>
  import('./components/editors/PdfEditor').then((module) => ({ default: module.PdfEditor }))
);
const BinaryViewer = React.lazy(() =>
  import('./components/editors/BinaryViewer').then((module) => ({ default: module.BinaryViewer }))
);





import { TerminalPanel } from './components/terminal/TerminalPanel';
import { copyToClipboard, getImgurClientId, uploadFilePath, uploadAlbum, uploadMany, gdriveUpload, shareViaSite } from './services/shareService';
import { readPreferences } from './state/preferences';
import {
  AIActionPayload,
  AppActionResult,
  SettingsTab,
  detectMarkdownTables,
  generateExcelBase64,
  generateCsvContent,
  generateDocxBase64,
  normalizeTableValue,
  sanitizeFileName,
} from './services/aiActions';
import {
  type ConversionFormat,
  defaultTargetName,
  guessFormat,
  normalizeConversionFormat,
} from './services/documentConversion';
import type { UpdateCheckResult } from './types';

/** Форматы, в которые приложение умеет записывать правки ИИ. */
const EDITABLE_TYPES = new Set<FileType>(['word', 'excel', 'markdown', 'richtext', 'text', 'code']);

/** A1 → координаты (без листа). */
function parseA1Cell(value: string): { row: number; col: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(value.trim());
  if (!match) return null;
  let col = 0;
  const letters = match[1].toUpperCase();
  for (const letter of letters) col = col * 26 + (letter.charCodeAt(0) - 64);
  return { row: Number(match[2]) - 1, col: col - 1 };
}

function columnLabel(index: number): string {
  let value = index + 1;
  let label = '';
  while (value > 0) {
    const rest = (value - 1) % 26;
    label = String.fromCharCode(65 + rest) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

const App: React.FC = () => {
  const ws = useWorkspace();
  const ai = useAISettings();
  const chat = useChatSessions();
  const { isDark, toggleTheme } = useTheme();

  // Оформление из настроек: тема, акцент, плотность, масштаб, меньше анимации.
  useEffect(() => {
    const media = window.matchMedia?.('(prefers-color-scheme: dark)');
    const apply = () => {
      applyAppearance(readPreferences().appearance, media?.matches ?? true);
    };
    apply();
    media?.addEventListener?.('change', apply);
    return () => media?.removeEventListener?.('change', apply);
  }, []);

  // Бюджет старта: время до первого кадра уходит в диагностический отчёт.
  useEffect(() => {
    const started = performance.now();
    const frame = window.requestAnimationFrame(() => {
      window.requestAnimationFrame(() => {
        const ms = Math.round(performance.now());
        void Promise.resolve(
          window.electronAPI?.diagnosticsLog?.(
            'info',
            `startup: первый кадр через ${ms} мс (effect стартовал на ${Math.round(started)} мс)`
          )
        ).catch(() => undefined);
      });
    });
    return () => window.cancelAnimationFrame(frame);
  }, []);
  // Диагностика: ошибки окна складываются в буфер главного процесса.
  useEffect(() => installDiagnosticsHooks(), []);
  // Автосохранение: снимок несохранённых вкладок раз в 2 секунды после правки.
  useEffect(() => {
    const dirty = ws.tabs.filter((tab) => tab.isDirty && tab.content);
    if (dirty.length === 0) return;
    const timer = window.setTimeout(() => {
      for (const tab of dirty) void writeSnapshot(tab);
      void pruneSnapshots(
        ws.tabs.filter((tab) => !tab.isDirty).map((tab) => tab.path)
      );
    }, 2000);
    return () => window.clearTimeout(timer);
  }, [ws.tabs]);

  // Восстановление после сбоя: поднимаем снимки в вкладки и сразу их убираем.
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      const snapshots = await listSnapshots();
      if (cancelled || snapshots.length === 0) return;
      let restored = 0;
      for (const snapshot of snapshots) {
        try {
          const tabId = await ws.openFile(snapshot.path);
          if (!tabId) continue;
          ws.updateTabContent(tabId, snapshot.content);
          // Снимок удаляем только после сохранения файла: иначе закрытие окна
          // без Ctrl+S потеряет восстановленное содержимое.
          restored += 1;
        } catch {
          // файл мог исчезнуть — снимок не восстанавливаем
        }
      }
      if (!cancelled && restored > 0) {
        ws.flash(`Восстановлены несохранённые документы: ${restored}`);
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  /** Уведомление: тост + тихий сигнал по правилам из настроек. */
  const notify = useCallback(
    (message: string, kind: NotificationKind = 'save') => {
      playChimeFor(kind);
      ws.flash(message);
    },
    [ws]
  );

  // Нечитаемый ключ (перенос профиля/portable на другой компьютер) — сказать
  // об этом прямо, иначе приложение выглядит «сломанным».
  useEffect(() => {
    if (ai.keyProblem) ws.flash(ai.keyProblem);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ai.keyProblem]);

  const [sidebarOpen, setSidebarOpen] = useLocalStorage('ws_sidebar_open', true);
  const [aiPanelOpen, setAiPanelOpen] = useLocalStorage('ws_ai_panel_open', true);
  /** Права ассистента на запись в открытый документ: по умолчанию выданы. */
  const [aiDocWrite, setAiDocWrite] = useLocalStorage('ws_ai_doc_write', true);
  /** Актуальные значения для агентного цикла: колбэк живёт дольше одного рендера. */
  const wsRef = useRef(ws);
  wsRef.current = ws;
  const aiDocWriteRef = useRef(aiDocWrite);
  aiDocWriteRef.current = aiDocWrite;
  const [aiExtraWrite, setAiExtraWrite] = useLocalStorage('ws_ai_extra_write', false);
  const aiExtraWriteRef = useRef(aiExtraWrite);
  aiExtraWriteRef.current = aiExtraWrite;
  /** Снимок HTML на первую правку ответа: при ошибке весь ответ откатывается. */
  const docTxnRef = useRef<{ id: string; html: string } | null>(null);

  /** Реестр сессий открытых документов. */
  useEffect(() => {
    const openIds = new Set(ws.tabs.map((tab) => tab.id));
    for (const tab of ws.tabs) {
      registerDocumentSession({
        id: tab.id,
        kind: kindForTab(tab.type),
        path: tab.path,
        name: tab.name,
        writeAccess: aiDocWrite,
        tools: buildDocumentTools(aiDocWrite, tab.type).map((tool) => tool.function.name),
      });
    }
    for (const session of listDocumentSessions()) {
      if (!openIds.has(session.id)) removeDocumentSession(session.id);
    }
  }, [ws.tabs, aiDocWrite]);
  const [terminalOpen, setTerminalOpen] = useLocalStorage('ws_terminal_open', false);
  const [terminalHeight, setTerminalHeight] = useLocalStorage('ws_terminal_height', 240);
  const [quickOpenOpen, setQuickOpenOpen] = useState(false);
  const [commandOpen, setCommandOpen] = useState(false);
  const [commandQuery, setCommandQuery] = useState('');
  const [templatesOpen, setTemplatesOpen] = useState(false);
  const [searchPanelOpen, setSearchPanelOpen] = useState(false);

  // Раскладка: панели и открытые вкладки общие для окон, живут в workspace-session.json.
  const layoutApplyingRef = useRef(false);
  const layoutSavedRef = useRef('');
  useEffect(() => {
    const api = window.electronAPI;
    const applyLayout = (layout?: {
      aiPanelOpen?: boolean;
      terminalOpen?: boolean;
      searchPanelOpen?: boolean;
      openPaths?: string[];
    }) => {
      if (!layout) return;
      const normalized = normalizeLayout(layout);
      layoutApplyingRef.current = true;
      setAiPanelOpen(normalized.aiPanelOpen);
      setTerminalOpen(normalized.terminalOpen);
      setSearchPanelOpen(normalized.searchPanelOpen);
      window.setTimeout(() => {
        layoutApplyingRef.current = false;
      }, 0);
    };

    let disposed = false;
    void (async () => {
      try {
        const session = (await api?.sessionGet?.()) as
          | { layout?: Parameters<typeof applyLayout>[0] }
          | undefined;
        if (disposed) return;
        applyLayout(session?.layout);
        for (const path of normalizeLayout(session?.layout).openPaths) {
          if (path) void ws.openFile(path);
        }
      } catch {
        // сессия недоступна — стартуем как обычно
      }
    })();

    const off = api?.onSessionChanged?.((session) => {
      const layout = (session as { layout?: Parameters<typeof applyLayout>[0] } | undefined)?.layout;
      applyLayout(layout);
    });
    return () => {
      disposed = true;
      off?.();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Сохранение раскладки: только при реальных изменениях, с дебаунсом.
  useEffect(() => {
    if (layoutApplyingRef.current) return;
    const layoutState = normalizeLayout({
      aiPanelOpen,
      terminalOpen,
      searchPanelOpen,
      openPaths: ws.tabs.map((tab) => tab.path).filter(Boolean).slice(0, MAX_RESTORED_TABS),
      activePath: ws.activeTab?.path ?? '',
    });
    const fingerprint = layoutFingerprint(layoutState);
    if (fingerprint === layoutSavedRef.current) return;
    const timer = window.setTimeout(() => {
      layoutSavedRef.current = fingerprint;
      void window.electronAPI?.sessionPatch?.({ layout: layoutState });
    }, 600);
    return () => window.clearTimeout(timer);
  }, [aiPanelOpen, terminalOpen, searchPanelOpen, ws.tabs, ws.activeTab]);

  /** Папка, в которой открыт поиск (пункт «Искать в этой папке»). */
  const [searchScope, setSearchScope] = useState('');
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [settingsTab, setSettingsTab] = useState<SettingsTab>('provider');
  const [saveAsOpen, setSaveAsOpen] = useState(false);
  const [saveAsTabId, setSaveAsTabId] = useState('');
  const [saveAsBusy, setSaveAsBusy] = useState(false);
  const [updateInfo, setUpdateInfo] = useState<UpdateCheckResult | null>(null);
  const [updateDismissed, setUpdateDismissed] = useState(false);
  /** Команда, отправленная AI в терминал: выполняется после монтирования панели. */
  const [terminalCommand, setTerminalCommand] = useState<{ id: string; command: string } | null>(null);
  const [inlineTarget, setInlineTarget] = useState<InlineEditTarget | null>(null);
  const [pendingPrompt, setPendingPrompt] = useState<{
    prompt: string;
    context: string;
    token: number;
  } | null>(null);
  const [newFileDialog, setNewFileDialog] = useState<{
    open: boolean;
    parentDir: string;
    kind: 'file' | 'folder';
    name: string;
  }>({ open: false, parentDir: '', kind: 'file', name: '' });
  const [renameDialog, setRenameDialog] = useState<{
    open: boolean;
    path: string;
    name: string;
    busy: boolean;
  }>({ open: false, path: '', name: '', busy: false });
  const [unsavedDialog, setUnsavedDialog] = useState<{ open: boolean; tabId: string }>({
    open: false,
    tabId: '',
  });
  /** Подтверждение выхода из приложения с несохранёнными файлами. */
  const [exitDialogOpen, setExitDialogOpen] = useState(false);
  const [exitBusy, setExitBusy] = useState(false);
  const [bootError, setBootError] = useState('');

  const isElectron = fileSystemService.isElectron();

  // Запрещаем браузеру открывать перетащенные файлы вместо приложения:
  // обработчик чата получает их, а окно не должно уходить на file://.
  useEffect(() => {
    const preventFileNavigation = (event: DragEvent) => {
      if (Array.from(event.dataTransfer?.types || []).includes('Files')) {
        event.preventDefault();
      }
    };
    window.addEventListener('dragover', preventFileNavigation);
    window.addEventListener('drop', preventFileNavigation);
    return () => {
      window.removeEventListener('dragover', preventFileNavigation);
      window.removeEventListener('drop', preventFileNavigation);
    };
  }, []);

  /* ============================================================
     Подтверждение выхода из приложения
     Без этого закрытие окна теряло несохранённые изменения.
     ============================================================ */

  // Актуальный список несохранённых файлов для главного процесса
  const dirtyTabsRef = useRef(ws.dirtyTabs);
  dirtyTabsRef.current = ws.dirtyTabs;

  const requestExit = useCallback(() => {
    if (dirtyTabsRef.current.length === 0) {
      // Нечего сохранять — выходим без диалога
      window.electronAPI?.confirmClose(true);
    } else {
      setExitDialogOpen(true);
    }
  }, []);

  const requestExitRef = useRef(requestExit);
  requestExitRef.current = requestExit;

  useEffect(() => {
    if (!isElectron) return;
    // Главный процесс спрашивает подтверждение вместо немедленного закрытия
    const unsubscribe = window.electronAPI?.onCloseRequested?.(() => {
      requestExitRef.current();
    });
    return unsubscribe;
  }, [isElectron, requestExit]);

  const doExit = useCallback(() => {
    setExitDialogOpen(false);
    window.electronAPI?.confirmClose(true);
  }, []);

  const saveAllAndExit = useCallback(async () => {
    setExitBusy(true);
    try {
      // saveAll возвращает список неудачных файлов: решение о выходе
      // принимается по фактическому результату записи, а не по состоянию UI.
      const { saved, failed } = await ws.saveAll();
      if (failed.length === 0) {
        ws.flash(saved > 0 ? `Сохранено файлов: ${saved}` : 'Все файлы уже сохранены');
        doExit();
      } else {
        ws.flash(`Не удалось сохранить: ${failed.join(', ')} — выход отменён`);
      }
    } finally {
      setExitBusy(false);
    }
  }, [ws, doExit]);

  /* ============================================================
     Инициализация рабочего пространства
     ============================================================ */
  useEffect(() => {
    if (!isElectron) {
      setBootError(
        'WorkSpace запущен вне Electron. Файловые операции и офисные редакторы доступны только в настольном приложении (npm run dev).'
      );
      return;
    }
    void ws.initWorkspace().catch((err) => setBootError((err as Error).message));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isElectron]);

  /* ---- Тихая проверка обновлений при запуске ---- */
  useEffect(() => {
    if (!isElectron || !window.electronAPI?.checkForUpdates) return;
    const check = window.electronAPI.checkForUpdates;
    let cancelled = false;
    const timer = window.setTimeout(async () => {
      try {
        const info = await check();
        if (!cancelled && info?.success && info.available) setUpdateInfo(info);
      } catch {
        // Нет сети или GitHub недоступен — приложение работает как обычно.
      }
    }, 4000);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [isElectron]);

  /* ============================================================
     Горячие клавиши (Этап 6.3)
     ============================================================ */
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;

      // Ctrl+Shift+P — командная палитра
      if ((e.key === 'P' || e.key === 'p') && e.shiftKey) {
        e.preventDefault();
        setCommandOpen(true);
        setCommandQuery('');
        return;
      }

      // Ctrl+P — быстрый переход к файлу
      if (e.key === 'p' && !e.shiftKey) {
        e.preventDefault();
        setQuickOpenOpen(true);
        return;
      }

      // Ctrl+Shift+F — поиск по содержимому
      if (e.key === 'F' && e.shiftKey) {
        e.preventDefault();
        setSearchPanelOpen(true);
        setSidebarOpen(true);
        return;
      }

      // Ctrl+S — сохранить активную вкладку
      if (e.key === 's' && !e.shiftKey && !e.altKey) {
        e.preventDefault();
        if (ws.activeTabId) void ws.saveTab(ws.activeTabId);
        return;
      }

      // Ctrl+Shift+S — сохранить как (с конвертацией формата)
      if ((e.key === 's' || e.key === 'S') && e.shiftKey && !e.altKey) {
        e.preventDefault();
        if (ws.activeTabId) {
          setSaveAsTabId(ws.activeTabId);
          setSaveAsOpen(true);
        }
        return;
      }

      // Ctrl+Alt+S — сохранить все
      if (e.key === 's' && e.altKey) {
        e.preventDefault();
        void ws.saveAll().then(({ saved, failed }) => {
          if (failed.length === 0) {
            ws.flash(saved > 0 ? `Сохранено файлов: ${saved}` : 'Нет изменений для сохранения');
          } else {
            ws.flash(`Не удалось сохранить: ${failed.join(', ')}`);
          }
        });
        return;
      }

      // Ctrl+B — переключить боковую панель
      if (e.key === 'b' && !e.shiftKey) {
        e.preventDefault();
        setSidebarOpen((v) => !v);
        return;
      }

      // Ctrl+W — закрыть вкладку (с проверкой несохранённых изменений)
      if (e.key === 'w' && !e.shiftKey) {
        e.preventDefault();
        if (ws.activeTabId) requestCloseTab(ws.activeTabId);
        return;
      }

      // Ctrl+Tab — следующая вкладка
      if (e.key === 'Tab') {
        e.preventDefault();
        if (ws.tabs.length > 1) {
          const idx = ws.tabs.findIndex((t) => t.id === ws.activeTabId);
          const next = ws.tabs[(idx + 1) % ws.tabs.length];
          ws.setActiveTabId(next.id);
        }
        return;
      }

      // Ctrl+K — инлайн-копилот
      if (e.key === 'k' && !e.shiftKey) {
        e.preventDefault();
        openInlineCopilot();
        return;
      }

      // Ctrl+N — новый файл
      if (e.key === 'n' && !e.shiftKey) {
        e.preventDefault();
        if (ws.rootPath) {
          setNewFileDialog({ open: true, parentDir: ws.rootPath, kind: 'file', name: '' });
        }
        return;
      }

      // Ctrl+` (или Ctrl+~) — открыть/закрыть терминал
      if (e.key === '`' || e.key === '~' || e.key === 'ё' || e.key === 'Ё') {
        e.preventDefault();
        setTerminalOpen((v) => !v);
        return;
      }
    };

    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ws.activeTabId, ws.tabs, ws.rootPath, ws.saveTab, ws.saveAll]);

  /* ============================================================
     Открытие файлов и папок из проводника Windows (контекстное меню)
     ============================================================ */
  useEffect(() => {
    if (!window.electronAPI?.onOpenExternalPath) return;

    const handleExternalTarget = async (data: { path: string; isDirectory: boolean }) => {
      if (!data?.path) return;
      try {
        if (data.isDirectory) {
          await ws.openWorkspace(data.path);
          ws.flash(`Открыта папка: ${baseName(data.path)}`);
          return;
        }

        const parentDir = dirName(data.path);
        const revealEnabled = readPreferences().explorerRevealOnOpen;
        if (!revealEnabled) {
          if (!ws.rootPath) await ws.openWorkspace(parentDir);
          await ws.openFile(data.path);
          ws.flash(`Открыт файл: ${baseName(data.path)}`);
          return;
        }
        // Файл внутри текущего дерева — только раскрываем и подсвечиваем.
        const insideCurrent = ws.rootPath ? await ws.revealPath(data.path) : false;
        if (!insideCurrent) {
          // Иначе корнем становится папка файла: так в проводнике видно его
          // окружение, а не дерево посторонней папки.
          await ws.openWorkspace(parentDir);
          await ws.revealPath(data.path);
        }
        await ws.openFile(data.path);
        ws.flash(`Открыт файл: ${baseName(data.path)}`);
      } catch (err) {
        ws.flash(`Ошибка открытия: ${(err as Error).message}`);
      }
    };

    // Проверяем аргумент при холодном запуске
    void window.electronAPI
      .getInitialExternalPath?.()
      .then((init) => {
        if (init) void handleExternalTarget(init);
      })
      .catch(() => {
        // Проверочные среды могут не регистрировать этот IPC-канал
      });

    // Слушаем открытие при уже работающем приложении (second-instance)
    const unsub = window.electronAPI.onOpenExternalPath((data) => {
      void handleExternalTarget(data);
    });

    return unsub;
  }, [ws]);

  /* ============================================================
     Запрос контекста у активного редактора (для Ctrl+K)
     ============================================================ */
  const editorContextGetter = useRef<(() => { text: string; from: number; to: number }) | null>(null);

  const registerEditorContext = useCallback(
    (getter: (() => { text: string; from: number; to: number }) | null) => {
      editorContextGetter.current = getter;
    },
    []
  );

  const openInlineCopilot = useCallback(() => {
    const tab = ws.activeTab;
    if (!tab) {
      ws.flash('Откройте файл, чтобы использовать AI-правку');
      return;
    }

    const ctx = editorContextGetter.current?.();
    const selection = ctx?.text ?? '';
    const hasSelection = selection.trim().length > 0;

    // Окно появляется в центре верхней части рабочей области
    setInlineTarget({
      tabId: tab.id,
      original: hasSelection ? selection : tab.content,
      from: hasSelection ? ctx!.from : 0,
      to: hasSelection ? ctx!.to : tab.content.length,
      fileName: tab.name,
      anchor: { top: 96, left: Math.max(24, window.innerWidth / 2 - 280) },
      hasSelection,
    });
  }, [ws]);

  const applyInlineEdit = useCallback(
    (target: InlineEditTarget, newText: string) => {
      const tab = ws.tabs.find((t) => t.id === target.tabId);
      if (!tab) return;
      const next = tab.content.slice(0, target.from) + newText + tab.content.slice(target.to);
      ws.updateTabContent(target.tabId, next);
      ws.flash('Изменения применены');
      // Применяем к живому редактору через зарегистрированный аппlier
      editorApplier.current?.({ from: target.from, to: target.to, text: newText });
    },
    [ws]
  );

  const editorApplier = useRef<((edit: { from: number; to: number; text: string }) => void) | null>(
    null
  );

  /* ============================================================
     Работа с вкладками и запросами к AI
     ============================================================ */
  const requestCloseTab = useCallback(
    (tabId: string) => {
      const tab = ws.tabs.find((t) => t.id === tabId);
      if (tab?.isDirty) {
        setUnsavedDialog({ open: true, tabId });
      } else {
        ws.closeTab(tabId);
      }
    },
    [ws]
  );

  /** Режим Excel-формулы: AI-ответ передаётся обратно в редактор таблицы. */
  const [formulaRequestTabId, setFormulaRequestTabId] = useState<string>('');

  /** Запрос к AI из любого редактора — открывает панель и отправляет промт. */
  const askAI = useCallback(
    (prompt: string, contextText: string) => {
      setAiPanelOpen(true);
      // Запоминаем вкладку, инициировавшую запрос, чтобы вернуть ответ адресно
      setFormulaRequestTabId(ws.activeTabId);
      setPendingPrompt({ prompt, context: contextText, token: Date.now() });
    },
    [setAiPanelOpen, ws.activeTabId]
  );

  /**
   * Ответ AI готов. Редактор таблицы слушает это событие и предлагает
   * применить формулу с подтверждением.
   */
  const handleAIAnswer = useCallback(
    (answer: string) => {
      window.dispatchEvent(
        new CustomEvent('workspace:ai-answer', {
          detail: { tabId: formulaRequestTabId, answer },
        })
      );
    },
    [formulaRequestTabId]
  );

  const handleAIApply = useCallback(
    (action: AIAction, result: string) => {
      ws.flash(result);
    },
    [ws]
  );

  const createFromTemplate = useCallback(
    async (template: (typeof DOCUMENT_TEMPLATES)[number]) => {
      try {
        const base = template.name.replace(/[\\/:*?"<>|]/g, ' ').trim() || 'Документ';
        const extension = template.extension;
        const used = new Set((ws.allFiles ?? []).map((item) => item.name));
        let name = `${base}.${extension}`;
        let index = 2;
        while (used.has(name)) name = `${base} (${index++}).${extension}`;

        if (extension === 'docx') {
          const base64 = await htmlToDocxBase64(template.html ?? '<p></p>');
          await ws.createFile(ws.rootPath, name, base64, 'base64');
        } else {
          await ws.createFile(ws.rootPath, name, template.markdown ?? '', 'utf-8');
        }
        setTemplatesOpen(false);
        ws.flash(`Создан «${name}»`);
      } catch (error) {
        ws.flash(`Не удалось создать документ: ${(error as Error).message}`);
      }
    },
    [ws]
  );

  const insertIntoEditor = useCallback(
    (text: string) => {
      const tab = ws.activeTab;
      if (!tab) {
        ws.flash('Нет активного документа для вставки');
        return;
      }

      if (!EDITABLE_TYPES.has(tab.type)) {
        ws.flash(`Формат ${tab.extension.toUpperCase() || tab.type} доступен только для просмотра`);
        return;
      }

      if (tab.type === 'excel') {
        window.dispatchEvent(
          new CustomEvent('workspace:excel-insert-text', {
            detail: { tabId: tab.id, text },
          })
        );
        ws.flash('Данные вставлены в таблицу');
        return;
      }

      // В документ попадает только сам текст: отчёт модели («Понял…»,
      // «Готово», советы по оформлению, ```) вставлять нельзя.
      const body =
        tab.type === 'word' || tab.type === 'markdown' || tab.type === 'richtext'
          ? extractDocumentContent(text) || text
          : text;

      if (tab.type === 'word') {
        window.dispatchEvent(
          new CustomEvent('workspace:insert-text', {
            detail: { tabId: tab.id, text: body },
          })
        );
        ws.flash('Текст вставлен в документ Word');
        return;
      }

      if (tab.type === 'markdown' || tab.type === 'richtext') {
        window.dispatchEvent(
          new CustomEvent('workspace:insert-text', {
            detail: { tabId: tab.id, text: body },
          })
        );
        const next = tab.content ? `${tab.content}\n\n${body}` : body;
        ws.updateTabContent(tab.id, next);
        ws.flash('Текст вставлен в документ');
        return;
      }

      const applier = editorApplier.current;
      if (applier) {
        applier({ from: -1, to: -1, text });
      } else {
        ws.updateTabContent(tab.id, `${tab.content}\n\n${text}`);
      }
      ws.flash('Текст вставлен в документ');
    },
    [ws]
  );

  const replaceDocument = useCallback(
    async (text: string) => {
      const tab = ws.activeTab;
      if (!tab) return;

      // Модель может прислать весь ответ чата: вступление, советы по
      // оформлению и обёртку ```. В документ попадает только его текст.
      const body = extractDocumentContent(text) || text;

      if (!EDITABLE_TYPES.has(tab.type)) {
        ws.flash(
          `Формат ${tab.extension.toUpperCase() || tab.type} доступен только для просмотра — правку некуда записать`
        );
        return;
      }

      if (tab.type === 'excel') {
        window.dispatchEvent(
          new CustomEvent('workspace:excel-replace-document', {
            detail: { tabId: tab.id, text: body },
          })
        );
        ws.flash('Таблица заменена данными ответа');
        return;
      }

      if (tab.type === 'word') {
        // Живой редактор сам пересоберёт пакет: сохранит поля страницы и
        // вернёт картинки, которые модель не воспроизводит. Заголовок из
        // имени файла при замене не подставляем — это не часть документа.
        if (getDocumentController(tab.id)) {
          window.dispatchEvent(
            new CustomEvent('workspace:replace-document', {
              detail: { tabId: tab.id, text: body },
            })
          );
          ws.flash('Документ Word заменён ответом AI');
          return;
        }
        try {
          const base64 = await htmlToDocxBase64(markdownToDocumentHtml(body));
          ws.updateTabContent(tab.id, base64);
          ws.flash('Документ Word заменён ответом AI');
        } catch (err) {
          ws.flash(`Ошибка замены документа: ${(err as Error).message}`);
        }
        return;
      }

      if (tab.type === 'markdown' || tab.type === 'richtext') {
        ws.updateTabContent(tab.id, body);
        window.dispatchEvent(
          new CustomEvent('workspace:replace-document', {
            detail: { tabId: tab.id, text: body },
          })
        );
        ws.flash('Документ заменён ответом AI');
        return;
      }

      ws.updateTabContent(tab.id, body);
      editorApplier.current?.({ from: 0, to: tab.content.length, text: body });
      ws.flash('Документ заменён ответом AI');
    },
    [ws]
  );

  /* ============================================================
     Действия AI — полное управление приложением
     ============================================================ */

  const workspaceDir = useCallback(
    async (): Promise<string> =>
      ws.rootPath || (await fileSystemService.getCurrentWorkspace()) || '.',
    [ws.rootPath]
  );

  /** Относительные пути в действиях AI считаются от корня рабочей области. */
  const resolveWorkspacePath = useCallback(
    async (rawPath?: string): Promise<string> => {
      const value = (rawPath || '').trim().replace(/^["']|["']$/g, '');
      const root = await workspaceDir();
      if (!value) return root;
      if (/^[a-zA-Z]:[\\/]/.test(value) || value.startsWith('\\\\')) return value;
      return joinPath(root, value.replace(/^[\\/]+/, ''));
    },
    [workspaceDir]
  );

  /** Поиск открытой таблицы по имени/пути — приоритетнее склейки с корнем. */
  const resolveTableTarget = useCallback(
    async (rawTarget: string): Promise<string> => {
      const normalized = rawTarget.replace(/\\/g, '/').toLowerCase();
      const base = normalized.split('/').pop() || normalized;
      const open = ws.tabs.find(
        (tab) =>
          tab.path.replace(/\\/g, '/').toLowerCase() === normalized ||
          tab.name.toLowerCase() === base
      );
      if (open) return open.path;
      const inTree = ws.allFiles.find(
        (file) =>
          file.name.toLowerCase() === base ||
          file.path.replace(/\\/g, '/').toLowerCase().endsWith(`/${normalized}`)
      );
      if (inTree) return inTree.path;
      return resolveWorkspacePath(rawTarget);
    },
    [ws.tabs, ws.allFiles, resolveWorkspacePath]
  );

  /** Поиск открытой вкладки по пути, имени или индексу — для select_tab/close_tab. */
  const findOpenTab = useCallback(
    (target: { path?: string; name?: string; index?: number }) => {
      const norm = (p: string) => p.toLowerCase().replace(/\//g, '\\');
      if (target.path) {
        const wanted = norm(target.path);
        const found = ws.tabs.find(
          (t) => norm(t.path) === wanted || norm(t.path).endsWith('\\' + wanted)
        );
        if (found) return found;
      }
      if (target.name) {
        const wanted = target.name.toLowerCase().trim();
        const exact = ws.tabs.find((t) => t.name.toLowerCase() === wanted);
        if (exact) return exact;
        const partial = ws.tabs.find((t) => t.name.toLowerCase().includes(wanted));
        if (partial) return partial;
      }
      if (typeof target.index === 'number' && ws.tabs[target.index]) {
        return ws.tabs[target.index];
      }
      return ws.activeTab;
    },
    [ws.tabs, ws.activeTab]
  );

  /**
   * Живое содержимое вкладки для экспорта: Word/Excel/PowerPoint держат
   * актуальные правки в своих редакторах, а не в tab.content.
   */
  const resolveLiveTabContent = useCallback(async (tab: OpenTab): Promise<string> => {
    if (tab.type === 'word') {
      const controller = getDocumentController(tab.id);
      if (controller?.serialize) {
        try {
          return await controller.serialize();
        } catch {
          return tab.content;
        }
      }
    }
    if (tab.type === 'excel') {
      const controller = getSheetController(tab.id);
      if (controller?.serialize) {
        try {
          return await controller.serialize();
        } catch {
          return tab.content;
        }
      }
    }
    if (tab.type === 'powerpoint') {
      const controller = getPptxController(tab.id);
      if (controller?.getContent) {
        try {
          return await controller.getContent();
        } catch {
          return tab.content;
        }
      }
    }
    return tab.content;
  }, []);

  const openSaveAs = useCallback(
    (tabId?: string) => {
      const id = tabId || ws.activeTabId;
      if (!id) return;
      setSaveAsTabId(id);
      setSaveAsOpen(true);
    },
    [ws.activeTabId]
  );

  /** Сохраняет открытую вкладку в новый файл; использует живой редактор. */
  const exportTabAs = useCallback(
    async (
      tab: OpenTab,
      targetPath: string,
      format: ConversionFormat
    ): Promise<AppActionResult> => {
      const content = await resolveLiveTabContent(tab);
      const activeSheet =
        tab.type === 'excel' ? getSheetController(tab.id)?.getActiveSheet?.() : undefined;
      const ok = await wsRef.current.saveTabAs(
        tab.id,
        targetPath,
        format,
        content,
        activeSheet
      );
      return ok
        ? { ok: true, message: `Сохранено: ${baseName(targetPath)}`, path: targetPath }
        : { ok: false, message: `Не удалось сохранить: ${baseName(targetPath)}` };
    },
    [resolveLiveTabContent]
  );

  /** Колбэк диалога «Сохранить как». */
  const handleSaveAs = useCallback(
    async (targetPath: string, format: ConversionFormat): Promise<boolean> => {
      const tab = wsRef.current.tabs.find((item) => item.id === saveAsTabId);
      if (!tab) return false;
      setSaveAsBusy(true);
      try {
        const result = await exportTabAs(tab, targetPath, format);
        return result.ok;
      } finally {
        setSaveAsBusy(false);
      }
    },
    [saveAsTabId, exportTabAs]
  );

  /**
   * Исполнитель инструментов документа для агентного цикла.
   * Читает живое полотно через docBridge, пишет через него же, чтобы правки
   * сразу были видны в редакторе, а не только в base64 вкладки.
   */
  const runDocumentTool = useCallback(
    async (
      name: string,
      args: Record<string, unknown> = {},
      transactionId = `doc_${Date.now()}`,
      tabId?: string
    ): Promise<DocumentToolResult> => {
      const current = wsRef.current;
      const tab = tabId
        ? current.tabs.find((item) => item.id === tabId)
        : current.activeTab;

      if (name === 'convert_file') {
        const rawSource = String(args.source ?? args.path ?? args.file ?? '').trim();
        if (!rawSource) {
          return { ok: false, message: 'Укажите исходный файл для конвертации' };
        }
        const sourcePath = await resolveWorkspacePath(rawSource);
        const rawTarget = String(args.target ?? args.output ?? args.to ?? '').trim();
        const rawFormat = String(args.format ?? args.as ?? '').trim();
        const format =
          normalizeConversionFormat(rawFormat) || (rawTarget ? guessFormat(rawTarget) : null);
        if (!format) {
          return { ok: false, message: 'Укажите формат: docx, xlsx, csv, md, txt, html или pdf' };
        }
        const ext = extName(sourcePath) || 'файл';
        const targetPath = rawTarget
          ? await resolveWorkspacePath(rawTarget)
          : await fileSystemService.uniquePath(
              dirName(sourcePath),
              defaultTargetName(baseName(sourcePath), format)
            );
        const result = await current.convertFile(sourcePath, targetPath, format);
        return {
          ok: result.ok,
          message:
            result.message ||
            (result.ok
              ? `Конвертировано: ${baseName(sourcePath)} → ${baseName(targetPath)}`
              : `Не удалось конвертировать .${ext}`),
          changed: false,
          meta: { path: targetPath, format },
        };
      }

      if (name === 'save_as') {
        if (!tab) return { ok: false, message: 'Нет открытого документа для сохранения' };
        const rawFormat = String(args.format ?? args.as ?? '').trim();
        const rawTarget = String(args.path ?? args.target ?? '').trim();
        const format =
          normalizeConversionFormat(rawFormat) || (rawTarget ? guessFormat(rawTarget) : null);
        if (!format) {
          return {
            ok: false,
            message:
              'Укажите формат сохранения: xlsx, csv, tsv, docx, md, txt, html или pdf',
          };
        }
        let targetPath = '';
        if (rawTarget) {
          const resolved = await resolveWorkspacePath(rawTarget);
          const targetExt = extName(resolved);
          // path может быть папкой: тогда имя берём из name/исходного файла.
          targetPath =
            targetExt && normalizeConversionFormat(targetExt)
              ? resolved
              : joinPath(
                  resolved,
                  String(args.name ?? '').trim() || defaultTargetName(tab.name, format)
                );
        } else {
          targetPath = joinPath(
            dirName(tab.path),
            String(args.name ?? '').trim() || defaultTargetName(tab.name, format)
          );
        }
        if (extName(targetPath) !== format) {
          targetPath = targetPath.replace(/\.[^./\\]+$/, '') + `.${format}`;
        }
        if (!rawTarget) {
          targetPath = await fileSystemService.uniquePath(
            dirName(targetPath),
            baseName(targetPath)
          );
        }
        const result = await exportTabAs(tab, targetPath, format);
        return {
          ok: result.ok,
          message: result.message,
          changed: false,
          meta: { path: targetPath, format },
        };
      }

      if (!tab || !AGENT_TAB_TYPES.has(tab.type)) {
        return { ok: false, message: 'Документ закрыт или недоступен для правки' };
      }

      updateDocumentSession(tab.id, {
        capabilities: { read: true, write: aiDocWriteRef.current },
      });

      if (DOCUMENT_WRITE_TOOLS.has(name) && !aiDocWriteRef.current) {
        return {
          ok: false,
          message:
            'Нет прав записи в открытый документ. Нажмите «Предоставить права на запись» в панели AI.',
        };
      }

      if (name.startsWith('excel_')) {
        if (tab.type !== 'excel') {
          return { ok: false, message: 'Инструменты Excel доступны только для открытой таблицы' };
        }
        const sheet = getSheetController(tab.id);
        if (!sheet) return { ok: false, message: 'Редактор таблицы ещё не готов' };

        if (name === 'excel_read') {
          const read = sheet.readSheet(
            args.sheet ? String(args.sheet) : undefined,
            args.range ? String(args.range) : undefined,
            Number(args.limit) || undefined
          );
          if (!read) return { ok: false, message: 'Лист или диапазон не найдены' };
          const startRow = read.cells[0] ? parseA1Cell(read.cells[0].address)?.row ?? 0 : 0;
          const table = read.rows
            .map((row, index) => `${startRow + index + 1}\t${row.join('\t')}`)
            .join('\n');
          const header = `Лист «${read.sheet}», диапазон ${read.range}${
            read.truncated ? ' (обрезано)' : ''
          }. Листы: ${read.sheetNames.join(', ')}`;
          return {
            ok: true,
            message: header,
            changed: false,
            content: table,
            structure: `${header}\n${table}`,
            meta: { sheet: read.sheet, sheetNames: read.sheetNames, cells: read.cells },
          };
        }

        if (name === 'excel_get_selection') {
          const read = sheet.getSelection();
          if (!read) return { ok: false, message: 'Нет активного выделения' };
          const table = read.rows.map((row, index) => `${index + 1}\t${row.join('\t')}`).join('\n');
          return {
            ok: true,
            message: `Выделено: лист «${read.sheet}», диапазон ${read.range}`,
            changed: false,
            content: table,
            structure: table,
            meta: { sheet: read.sheet, range: read.range, cells: read.cells },
          };
        }

        if (name === 'excel_write_cells') {
          const cells: Array<{ address: string; value: string }> = [];
          if (Array.isArray(args.cells)) {
            for (const item of args.cells) {
              if (!item || typeof item !== 'object') continue;
              const address = String((item as Record<string, unknown>).address ?? '');
              const value = (item as Record<string, unknown>).value;
              if (address) cells.push({ address, value: value == null ? '' : String(value) });
            }
          } else if (typeof args.range === 'string' && Array.isArray(args.values)) {
            const [startRef, endRef] = args.range.split(':');
            const start = parseA1Cell(startRef);
            const end = parseA1Cell(endRef || startRef);
            if (start && end) {
              const matrix = args.values as unknown[];
              for (let row = start.row; row <= end.row; row += 1) {
                const line = Array.isArray(matrix[row - start.row])
                  ? (matrix[row - start.row] as unknown[])
                  : [];
                for (let col = start.col; col <= end.col; col += 1) {
                  cells.push({
                    address: `${columnLabel(col)}${row + 1}`,
                    value: line[col - start.col] == null ? '' : String(line[col - start.col]),
                  });
                }
              }
            }
          }
          if (cells.length === 0) {
            return { ok: false, message: 'Передайте cells или range + values' };
          }
          const result = sheet.writeCells(
            args.sheet ? String(args.sheet) : undefined,
            cells,
            transactionId
          );
          return result;
        }

        if (name === 'excel_add_sheet') {
          return sheet.addSheet(args.name ? String(args.name) : undefined);
        }
        if (name === 'excel_rename_sheet') {
          return sheet.renameSheet(
            String(args.old_name ?? args.oldName ?? ''),
            String(args.new_name ?? args.newName ?? '')
          );
        }
        if (name === 'excel_delete_sheet') {
          return sheet.deleteSheet(String(args.name ?? ''));
        }
        if (name === 'excel_set_active_sheet') {
          return sheet.setActiveSheet(String(args.name ?? ''));
        }
        if (name === 'excel_insert_rows') {
          return sheet.insertRows(
            args.sheet ? String(args.sheet) : undefined,
            Math.max(0, (Number(args.index) || 1) - 1),
            Math.max(1, Number(args.count) || 1)
          );
        }
        if (name === 'excel_delete_rows') {
          return sheet.deleteRows(
            args.sheet ? String(args.sheet) : undefined,
            Math.max(0, (Number(args.index) || 1) - 1),
            Math.max(1, Number(args.count) || 1)
          );
        }
        if (name === 'excel_insert_columns') {
          return sheet.insertColumns(
            args.sheet ? String(args.sheet) : undefined,
            Math.max(0, (Number(args.index) || 1) - 1),
            Math.max(1, Number(args.count) || 1)
          );
        }
        if (name === 'excel_delete_columns') {
          return sheet.deleteColumns(
            args.sheet ? String(args.sheet) : undefined,
            Math.max(0, (Number(args.index) || 1) - 1),
            Math.max(1, Number(args.count) || 1)
          );
        }
        if (name === 'excel_merge_cells') {
          return sheet.mergeCells(
            args.sheet ? String(args.sheet) : undefined,
            String(args.range ?? '')
          );
        }
        if (name === 'excel_unmerge_cells') {
          return sheet.unmergeCells(
            args.sheet ? String(args.sheet) : undefined,
            String(args.range ?? '')
          );
        }
        if (name === 'excel_format_range') {
          const patch: Parameters<typeof sheet.formatRange>[2] = {};
          if (args.bold !== undefined) patch.bold = Boolean(args.bold);
          if (args.italic !== undefined) patch.italic = Boolean(args.italic);
          if (args.fill) patch.fill = String(args.fill);
          if (args.color) patch.color = String(args.color);
          if (Number(args.font_size)) patch.fontSize = Number(args.font_size);
          const align = String(args.align ?? '');
          if (align === 'left' || align === 'center' || align === 'right') patch.align = align;
          if (args.wrap !== undefined) patch.wrap = Boolean(args.wrap);
          return sheet.formatRange(
            args.sheet ? String(args.sheet) : undefined,
            String(args.range ?? ''),
            patch
          );
        }

        return { ok: false, message: `Неизвестный инструмент Excel: ${name}` };
      }

      if (name.startsWith('pptx_')) {
        if (tab.type !== 'powerpoint') {
          return { ok: false, message: 'Инструменты PowerPoint доступны только для презентации' };
        }
        const ppt = getPptxController(tab.id);
        if (!ppt) return { ok: false, message: 'Редактор презентации ещё не готов' };

        if (name === 'pptx_read') {
          const model = await ppt.readLegacyModel?.();
          if (!model) return { ok: false, message: 'Модель презентации недоступна' };
          const lines = model.slides.map((slide) => {
            const shapes = slide.shapes
              .map(
                (shape) =>
                  `  shape_id=${shape.shapeId}${shape.isTitle ? ' [title]' : ''} «${shape.text.replace(/\s+/g, ' ').slice(0, 160)}»`
              )
              .join('\n');
            const notes = slide.notes ? `\n  заметки: ${slide.notes}` : '';
            return `Слайд ${slide.index + 1} (${slide.shapes.length} фигур):\n${shapes || '  (текста нет)'}${notes}`;
          });
          return {
            ok: true,
            message: `Слайдов: ${model.slides.length}`,
            changed: false,
            structure: `Макеты: ${
              model.layouts.map((layout) => `${layout.index}. ${layout.name}`).join(', ') ||
              '(нет)'
            }\n\n${lines.join('\n\n')}`,
            meta: {
              width: model.width,
              height: model.height,
              layouts: model.layouts,
              slides: model.slides.map((slide) => ({
                index: slide.index,
                shapes: slide.shapes.map((shape) => ({
                  shapeId: shape.shapeId,
                  name: shape.name,
                  text: shape.text,
                  isTitle: shape.isTitle,
                })),
              })),
            },
          };
        }

        if (name === 'pptx_get_selection') {
          const selection = ppt.getSelection();
          if (!selection) return { ok: false, message: 'Нет выделенной текстовой фигуры' };
          return {
            ok: true,
            message: `Выделена фигура ${selection.shapeId} на слайде ${selection.slide + 1}`,
            changed: false,
            content: selection.text,
            selection: {
              blockId: selection.shapeId,
              startOffset: 0,
              endOffset: selection.text.length,
              text: selection.text,
              html: '',
            },
            meta: { slide: selection.slide, shapeId: selection.shapeId, paragraphs: selection.paragraphs },
          };
        }

        if (name === 'pptx_set_text') {          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const shapeId = String(args.shape_id ?? args.shapeId ?? '');
          const text = String(args.text ?? args.content ?? '');
          if (!shapeId) return { ok: false, message: 'Не указан shape_id' };
          return ppt.setText(slide, shapeId, text.split('\n'));
        }

        if (name === 'pptx_add_slide') {
          const model = await ppt.readModel();
          const requested = Number(args.index ?? args.after);
          const index = Number.isFinite(requested)
            ? Math.max(0, Math.min(requested, model.slides.length))
            : model.slides.length;
          return ppt.addSlide(index, String(args.title ?? 'Новый слайд'));
        }

        if (name === 'pptx_delete_slide') {
          const index = Math.max(1, Number(args.index) || 1) - 1;
          return ppt.deleteSlide(index);
        }

        if (name === 'pptx_move_slide') {
          const from = Math.max(1, Number(args.from) || 1) - 1;
          const to = Math.max(1, Number(args.to) || 1) - 1;
          return ppt.moveSlide(from, to);
        }

        if (name === 'pptx_add_text_box') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const text = String(args.text ?? args.content ?? '');
          if (!text.trim()) return { ok: false, message: 'Пустой текст фигуры' };
          return ppt.addTextBox(slide, text);
        }

        if (name === 'pptx_insert_image') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const src = String(args.src ?? args.image ?? '').trim();
          if (!src) return { ok: false, message: 'Не указан источник изображения' };
          const root =
            current.rootPath || (await fileSystemService.getCurrentWorkspace()) || '';
          const resolved = await resolveImageSource(src, root);
          if (!resolved.ok) return { ok: false, message: resolved.message };
          return ppt.addImage(slide, resolved.dataUrl, {
            width: args.width !== undefined ? Number(args.width) : undefined,
            height: args.height !== undefined ? Number(args.height) : undefined,
            x: args.x !== undefined ? Number(args.x) : undefined,
            y: args.y !== undefined ? Number(args.y) : undefined,
            alt: typeof args.alt === 'string' ? args.alt : undefined,
          });
        }

        if (name === 'pptx_set_layout') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const requested = String(args.layout ?? args.name ?? args.part ?? '').trim();
          if (!requested) return { ok: false, message: 'Не указан макет' };
          const model = await ppt.readLegacyModel?.();
          if (!model) return { ok: false, message: 'Макеты недоступны' };
          const numeric = /^\d+$/.test(requested) ? Number(requested) : 0;
          const layout = numeric
            ? model.layouts[numeric - 1]
            : model.layouts.find(
                (item) =>
                  item.name.toLowerCase() === requested.toLowerCase() ||
                  item.part.toLowerCase() === requested.toLowerCase()
              );
          if (!layout) {
            return {
              ok: false,
              message: `Макет «${requested}» не найден. Доступны: ${model.layouts
                .map((item) => item.name)
                .join(', ')}`,
            };
          }
          return ppt.setLayout(slide, layout.part);
        }

        if (name === 'pptx_set_notes') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const text = String(args.text ?? args.content ?? '');
          return ppt.setNotes(slide, text);
        }

        /* ---------- Полное управление презентацией ---------- */

        if (name === 'pptx_get_slide') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const detail = await ppt.getSlide?.(slide);
          if (!detail) return { ok: false, message: `Слайд ${slide + 1} не найден` };
          const lines = detail.shapes.map(
            (shape) =>
              `  shape_id=${shape.id} ${shape.kind}${/заголовок/i.test(shape.name) ? ' [title]' : ''}` +
              ` (${Math.round(shape.x / 36000)},${Math.round(shape.y / 36000)} см · ${Math.round(
                shape.w / 36000
              )}×${Math.round(shape.h / 36000)} см)` +
              (shape.text
                ? ` «${shape.text.paragraphs
                    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
                    .join(' / ')
                    .slice(0, 160)}»`
                : '')
          );
          return {
            ok: true,
            message: `Слайд ${slide + 1}: ${detail.shapes.length} фигур`,
            changed: false,
            structure: [
              `Фон: #${detail.background}; переход: ${detail.transition?.type ?? 'none'}; скрыт: ${
                detail.hidden ? 'да' : 'нет'
              }`,
              ...lines,
              detail.notes ? `Заметки: ${detail.notes}` : '',
            ]
              .filter(Boolean)
              .join('\n'),
            meta: { index: slide, shapes: detail.shapes.length },
          };
        }

        if (name === 'pptx_duplicate_slide') {
          const index = Math.max(1, Number(args.index ?? args.slide) || 1) - 1;
          return ppt.duplicateSlide
            ? await ppt.duplicateSlide(index)
            : { ok: false, message: 'Дублирование не поддержано' };
        }

        if (name === 'pptx_hide_slide') {
          const index = Math.max(1, Number(args.index ?? args.slide) || 1) - 1;
          const hidden = args.hidden !== false && args.hidden !== 'false';
          return ppt.setSlideHidden
            ? await ppt.setSlideHidden(index, hidden)
            : { ok: false, message: 'Скрытие слайдов не поддержано' };
        }

        if (name === 'pptx_set_slide_size') {
          const preset = String(args.preset ?? '').toLowerCase();
          const sizes: Record<string, [number, number]> = {
            '16:9': [12192000, 6858000],
            '4:3': [9144000, 6858000],
            '16:10': [12192000, 7620000],
          };
          const width = Number(args.width_cm)
            ? Math.round(Number(args.width_cm) * 360000)
            : sizes[preset]?.[0];
          const height = Number(args.height_cm)
            ? Math.round(Number(args.height_cm) * 360000)
            : sizes[preset]?.[1];
          if (!width || !height) {
            return { ok: false, message: 'Укажите preset (16:9, 4:3, 16:10) или width_cm и height_cm' };
          }
          return ppt.setSlideSize
            ? await ppt.setSlideSize(width, height)
            : { ok: false, message: 'Размер слайда не поддержан' };
        }

        if (name === 'pptx_set_transition') {
          const index = Math.max(1, Number(args.index ?? args.slide) || 1) - 1;
          const type = String(args.type ?? 'none');
          const allowed = ['none', 'fade', 'push', 'wipe', 'cut'];
          if (!allowed.includes(type)) {
            return { ok: false, message: `Переход «${type}» неизвестен. Доступны: ${allowed.join(', ')}` };
          }
          return ppt.setTransition
            ? await ppt.setTransition(index, {
                type: type as 'none' | 'fade' | 'push' | 'wipe' | 'cut',
                speed: (args.speed as 'slow' | 'med' | 'fast') ?? 'med',
              })
            : { ok: false, message: 'Переходы не поддержаны' };
        }

        if (name === 'pptx_set_background') {
          const index = Math.max(1, Number(args.index ?? args.slide) || 1) - 1;
          const color = String(args.color ?? '').replace('#', '');
          const all = args.all === true || args.all === 'true';
          if (!color && !args.image) return { ok: false, message: 'Укажите color или image' };
          let imageDataUrl: string | undefined;
          if (args.image) {
            const root = current.rootPath || (await fileSystemService.getCurrentWorkspace()) || '';
            const resolved = await resolveImageSource(String(args.image), root);
            if (!resolved.ok) return { ok: false, message: resolved.message };
            imageDataUrl = resolved.dataUrl;
          }
          return ppt.setBackground
            ? await ppt.setBackground(index, { color: color || undefined, imageDataUrl, all })
            : { ok: false, message: 'Фон слайда не поддержан' };
        }

        if (name === 'pptx_format_text') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const shapeId = String(args.shape_id ?? args.shapeId ?? '');
          if (!shapeId) return { ok: false, message: 'Не указан shape_id' };
          const format: Record<string, unknown> = {};
          for (const key of ['bold', 'italic', 'underline', 'strike'] as const) {
            if (args[key] !== undefined) format[key] = args[key] === true || args[key] === 'true';
          }
          if (args.size !== undefined) format.size = Number(args.size);
          if (args.color !== undefined) format.color = String(args.color).replace('#', '');
          if (args.font !== undefined) format.font = String(args.font);
          if (args.align !== undefined) format.align = String(args.align);
          if (args.bullet !== undefined) format.bullet = String(args.bullet);
          if (args.level !== undefined) format.level = Number(args.level);
          if (args.line_spacing !== undefined) format.lineSpacing = Number(args.line_spacing);
          return ppt.formatText
            ? await ppt.formatText(slide, shapeId, format as Parameters<NonNullable<typeof ppt.formatText>>[2])
            : { ok: false, message: 'Форматирование текста не поддержано' };
        }

        if (name === 'pptx_find_replace') {
          const query = String(args.find ?? args.query ?? '');
          const replacement = String(args.replace ?? args.replacement ?? '');
          if (!query) return { ok: false, message: 'Не указан текст для поиска' };
          return ppt.findReplace
            ? await ppt.findReplace(query, replacement, {
                includeNotes: args.include_notes === true || args.include_notes === 'true',
              })
            : { ok: false, message: 'Поиск-замена не поддержан' };
        }

        if (name === 'pptx_add_shape') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const unit: 'cm' | 'emu' = args.unit === 'cm' ? 'cm' : 'emu';
          const geometry = {
            x: args.x !== undefined ? Number(args.x) : undefined,
            y: args.y !== undefined ? Number(args.y) : undefined,
            w: args.w !== undefined ? Number(args.w) : undefined,
            h: args.h !== undefined ? Number(args.h) : undefined,
            unit,
          };
          return ppt.addShape
            ? await ppt.addShape(slide, {
                geom: args.geom !== undefined ? String(args.geom) : undefined,
                geometry,
                fill: args.fill !== undefined ? String(args.fill).replace('#', '') : undefined,
                outline:
                  args.outline_color !== undefined
                    ? { color: String(args.outline_color).replace('#', ''), width: Number(args.outline_width) || 1 }
                    : undefined,
                shadow: args.shadow === true || args.shadow === 'true',
                text: args.text !== undefined ? String(args.text) : undefined,
              })
            : { ok: false, message: 'Вставка фигур не поддержана' };
        }

        if (name === 'pptx_set_shape_geometry') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const shapeId = String(args.shape_id ?? args.shapeId ?? '');
          if (!shapeId) return { ok: false, message: 'Не указан shape_id' };
          if (!ppt.setShapeGeometry) return { ok: false, message: 'Геометрия не поддержана' };
          return await ppt.setShapeGeometry(slide, shapeId, {
            x: args.x !== undefined ? Number(args.x) : undefined,
            y: args.y !== undefined ? Number(args.y) : undefined,
            w: args.w !== undefined ? Number(args.w) : undefined,
            h: args.h !== undefined ? Number(args.h) : undefined,
            rotation: args.rotation !== undefined ? Number(args.rotation) : undefined,
            unit: args.unit === 'cm' ? 'cm' : 'emu',
          });
        }

        if (name === 'pptx_delete_shape') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const shapeId = String(args.shape_id ?? args.shapeId ?? '');
          if (!shapeId) return { ok: false, message: 'Не указан shape_id' };
          return ppt.deleteShape
            ? await ppt.deleteShape(slide, shapeId)
            : { ok: false, message: 'Удаление фигур не поддержано' };
        }

        if (name === 'pptx_duplicate_shape') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const shapeId = String(args.shape_id ?? args.shapeId ?? '');
          if (!shapeId) return { ok: false, message: 'Не указан shape_id' };
          return ppt.duplicateShape
            ? await ppt.duplicateShape(slide, shapeId)
            : { ok: false, message: 'Дублирование фигур не поддержано' };
        }

        if (name === 'pptx_reorder_shape') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const shapeId = String(args.shape_id ?? args.shapeId ?? '');
          const mode = String(args.mode ?? 'front');
          const allowed = ['front', 'back', 'forward', 'backward'] as const;
          if (!shapeId) return { ok: false, message: 'Не указан shape_id' };
          if (!allowed.includes(mode as (typeof allowed)[number])) {
            return { ok: false, message: `Порядок «${mode}» неизвестен. Доступны: ${allowed.join(', ')}` };
          }
          return ppt.reorderShape
            ? await ppt.reorderShape(slide, shapeId, mode as (typeof allowed)[number])
            : { ok: false, message: 'Порядок наложения не поддержан' };
        }

        if (name === 'pptx_set_shape_style') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const shapeId = String(args.shape_id ?? args.shapeId ?? '');
          if (!shapeId) return { ok: false, message: 'Не указан shape_id' };
          return ppt.setShapeStyle
            ? await ppt.setShapeStyle(slide, shapeId, {
                fill:
                  args.fill === undefined
                    ? undefined
                    : args.fill === null || args.fill === 'none'
                      ? null
                      : String(args.fill).replace('#', ''),
                outline:
                  args.outline === undefined
                    ? undefined
                    : args.outline === null || args.outline === 'none'
                      ? null
                      : {
                          color: args.outline_color ? String(args.outline_color).replace('#', '') : undefined,
                          width: Number(args.outline_width) || 1,
                        },
                shadow: args.shadow === undefined ? undefined : args.shadow === true || args.shadow === 'true',
              })
            : { ok: false, message: 'Стиль фигуры не поддержан' };
        }

        if (name === 'pptx_add_table') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const rows = Math.max(1, Number(args.rows) || (Array.isArray(args.data) ? (args.data as string[][]).length : 2));
          const cols = Math.max(
            1,
            Number(args.cols) ||
              (Array.isArray(args.data) ? Math.max(...(args.data as string[][]).map((row) => row.length)) : 2)
          );
          if (!ppt.addTable) return { ok: false, message: 'Вставка таблиц не поддержана' };
          return await ppt.addTable(slide, {
            rows,
            cols,
            data: Array.isArray(args.data) ? (args.data as string[][]) : undefined,
            geometry: {
              x: args.x !== undefined ? Number(args.x) : undefined,
              y: args.y !== undefined ? Number(args.y) : undefined,
              w: args.w !== undefined ? Number(args.w) : undefined,
              h: args.h !== undefined ? Number(args.h) : undefined,
              unit: args.unit === 'cm' ? 'cm' : 'emu',
            },
          });
        }

        if (name === 'pptx_set_table_data') {
          const slide = Math.max(1, Number(args.slide) || 1) - 1;
          const shapeId = String(args.shape_id ?? args.shapeId ?? '');
          if (!shapeId) return { ok: false, message: 'Не указан shape_id' };
          if (!Array.isArray(args.data)) return { ok: false, message: 'Не переданы данные таблицы' };
          return ppt.setTableData
            ? await ppt.setTableData(slide, shapeId, args.data as string[][])
            : { ok: false, message: 'Правка таблиц не поддержана' };
        }

        if (name === 'pptx_apply_theme') {
          if (!ppt.applyTheme) return { ok: false, message: 'Темы не поддержаны' };
          const presetId = args.preset !== undefined ? String(args.preset) : undefined;
          return await ppt.applyTheme({
            presetId,
            colors:
              args.colors && typeof args.colors === 'object'
                ? (args.colors as Record<string, string>)
                : undefined,
            fonts: args.font_pair ? fontPresetById(String(args.font_pair)) : undefined,
          });
        }

        if (name === 'pptx_present') {
          if (!ppt.present) return { ok: false, message: 'Показ слайдов не поддержан' };
          const index = Number(args.slide);
          return await ppt.present(Number.isFinite(index) ? Math.max(0, index - 1) : undefined);
        }

        if (name === 'pptx_export_pdf') {
          if (!ppt.exportPdf) return { ok: false, message: 'Экспорт PDF не поддержан' };
          return await ppt.exportPdf(args.path ? String(args.path) : undefined);
        }

        if (name === 'pptx_export_images') {
          if (!ppt.exportSlideImages) return { ok: false, message: 'Экспорт PNG не поддержан' };
          const folder = args.dir ? await resolveWorkspacePath(String(args.dir)) : undefined;
          return await ppt.exportSlideImages(folder);
        }

        return { ok: false, message: `Неизвестный инструмент PowerPoint: ${name}` };
      }

      if (name === 'create_presentation') {
        const markdown = String(args.markdown ?? args.text ?? args.content ?? '');
        if (!markdown.trim()) return { ok: false, message: 'Пустая разметка презентации' };
        const dir = await resolveWorkspacePath(args.path ? String(args.path) : '');
        const fileName = sanitizeFileName(String(args.name ?? 'Презентация.pptx'), '.pptx');
        const base64 = await markdownToPptxBase64(markdown, {
          title: args.title ? String(args.title) : baseName(fileName).replace(/\.pptx$/i, ''),
          themePresetId: args.theme ? String(args.theme) : undefined,
        });
        const created = await ws.createFile(dir, fileName, base64, 'base64');
        if (!created) return { ok: false, message: `Не удалось создать ${fileName}` };
        await ws.openFile(created);
        return {
          ok: true,
          message: `Презентация создана: ${fileName} (слайдов по заголовкам)`,
          path: created,
        };
      }

      if (name.startsWith('image_')) {
        const root = current.rootPath || (await fileSystemService.getCurrentWorkspace()) || '';
        const imageTab = tab.type === 'image' ? tab : null;
        const bridge = imageTab ? getImageController(imageTab.id) : undefined;

        const readBase64 = async (
          filePath: string
        ): Promise<{ ok: true; base64: string; mime: string } | { ok: false; message: string }> => {
          const resolved = await resolveImageSource(filePath, root);
          if (!resolved.ok) return { ok: false, message: resolved.message };
          const match = /^data:([^;]+);base64,(.*)$/s.exec(resolved.dataUrl);
          if (!match) return { ok: false, message: 'Источник не является изображением' };
          return { ok: true, base64: match[2], mime: match[1] };
        };

        if (name === 'image_info') {
          const target = args.path ? await resolveWorkspacePath(String(args.path)) : '';
          const live = bridge?.getImage() ?? null;
          if (!target && !live) return { ok: false, message: 'Укажите path изображения' };
          const source = target ? await readBase64(target) : { ok: true as const, base64: live!.base64, mime: live!.mime };
          if (!source.ok) return { ok: false, message: source.message };
          const info = await applyImageOps({ base64: source.base64, mime: source.mime }, []);
          return {
            ok: true,
            message: `${info.width}×${info.height}, ${info.mime}, ${Math.round(info.bytes / 1024)} КБ`,
            changed: false,
            meta: { width: info.width, height: info.height, mime: info.mime, bytes: info.bytes },
          };
        }

        if (name === 'image_transform') {
          const ops = normalizeImageOps(args.ops);
          if (ops.length === 0) return { ok: false, message: 'Не переданы операции (ops)' };
          const target = args.path ? await resolveWorkspacePath(String(args.path)) : '';
          if (!target && bridge) {
            const result = await bridge.applyOps(ops);
            return { ok: result.ok, message: result.message, changed: result.ok };
          }
          if (!target) return { ok: false, message: 'Укажите path изображения' };
          const source = await readBase64(target);
          if (!source.ok) return { ok: false, message: source.message };
          const info = await applyImageOps({ base64: source.base64, mime: source.mime }, ops);
          const mode = String(args.output ?? 'copy');
          const ext = info.mime === 'image/jpeg' ? '.jpg' : info.mime === 'image/webp' ? '.webp' : '.png';
          const suffix = String(args.suffix ?? '-edit');
          let outPath: string;
          if (mode === 'overwrite') {
            outPath = target;
          } else if (mode === 'copy') {
            outPath = target.replace(/\.[^./\\]+$/, '') + suffix + ext;
          } else {
            outPath = await resolveWorkspacePath(mode);
          }
          await fileSystemService.writeFile(outPath, info.base64, 'base64');
          if (mode !== 'overwrite') await ws.openFile(outPath);
          return {
            ok: true,
            message: `${info.width}×${info.height}, ${Math.round(info.bytes / 1024)} КБ → ${baseName(outPath)}`,
            path: outPath,
            changed: true,
          };
        }

        if (name === 'image_batch') {
          const ops = normalizeImageOps(args.ops);
          if (ops.length === 0) return { ok: false, message: 'Не переданы операции (ops)' };
          const rawPaths = Array.isArray(args.paths) ? (args.paths as unknown[]).map(String) : [];
          if (rawPaths.length === 0) {
            return { ok: false, message: 'Передайте paths: обработка папок — через выделение в дереве файлов' };
          }
          const overwrite = args.overwrite === true || args.overwrite === 'true';
          const suffix = String(args.suffix ?? '-edit');
          const done: string[] = [];
          const failed: string[] = [];
          for (const relative of rawPaths) {
            try {
              const target = await resolveWorkspacePath(relative);
              const source = await readBase64(target);
              if (!source.ok) throw new Error(source.message);
              const info = await applyImageOps({ base64: source.base64, mime: source.mime }, ops);
              const ext = info.mime === 'image/jpeg' ? '.jpg' : info.mime === 'image/webp' ? '.webp' : '.png';
              const outPath = overwrite ? target : target.replace(/\.[^./\\]+$/, '') + suffix + ext;
              await fileSystemService.writeFile(outPath, info.base64, 'base64');
              done.push(baseName(outPath));
            } catch (err) {
              failed.push(`${relative}: ${(err as Error).message}`);
            }
          }
          return {
            ok: failed.length === 0,
            message:
              `Обработано ${done.length} из ${rawPaths.length}` + (failed.length ? `; ошибки: ${failed.join('; ')}` : ''),
            changed: done.length > 0,
            content: done.join('\n'),
          };
        }

        if (name === 'image_insert') {
          const path = String(args.path ?? '');
          if (!path) return { ok: false, message: 'Не указан path' };
          const target = await resolveWorkspacePath(path);
          const source = await readBase64(target);
          if (!source.ok) return { ok: false, message: source.message };
          const dataUrl = `data:${source.mime};base64,${source.base64}`;
          if (String(args.target ?? 'word') === 'slide') {
            const ppt = getPptxController(current.activeTab?.id ?? '');
            if (!ppt) return { ok: false, message: 'Откройте презентацию, чтобы вставить изображение на слайд' };
            const slide = Math.max(1, Number(args.slide) || 1) - 1;
            return await ppt.addImage(slide, dataUrl);
          }
          return await runDocumentTool('insert_image', { src: dataUrl }, transactionId, tab.id);
        }

        if (name === 'image_upload') {
          const host = String(args.host ?? 'imgbb').toLowerCase();
          if (!['imgbb', 'freeimage', 'yapx'].includes(host)) {
            return { ok: false, message: `Хостинг ${host} не поддержан инструментом. Доступны: imgbb, freeimage, yapx` };
          }
          const rawPaths = Array.isArray(args.paths) ? (args.paths as unknown[]).map(String) : [];
          if (rawPaths.length === 0) return { ok: false, message: 'Не указаны файлы (paths)' };
          const files: { path: string; name: string }[] = [];
          for (const relative of rawPaths) {
            const absolute = await resolveWorkspacePath(relative);
            files.push({ path: absolute, name: baseName(absolute) });
          }
          const result = await shareViaSite(host as 'imgbb' | 'freeimage' | 'yapx', files, {});
          return {
            ok: result.success,
            message: result.success
              ? `Загружено ${result.count} файл(ов)${result.albums?.length ? `: ${result.albums.join(', ')}` : ''}`
              : result.error || 'Не удалось загрузить изображения',
            path: result.albums?.[0],
            changed: false,
          };
        }

        if (name === 'image_open') {
          const path = String(args.path ?? '');
          if (!path) return { ok: false, message: 'Не указан path' };
          const target = await resolveWorkspacePath(path);
          const tabId = await ws.openFile(target);
          return tabId
            ? { ok: true, message: `Открыто изображение: ${baseName(target)}`, path: target }
            : { ok: false, message: 'Не удалось открыть изображение' };
        }
      }

      if (name.startsWith('video_')) {
        const api = window.electronAPI;
        const root = current.rootPath || (await fileSystemService.getCurrentWorkspace()) || '';
        const videoTab = tab.type === 'video' ? tab : null;
        const videoBridge = videoTab ? getVideoController(videoTab.id) : undefined;
        const filePath = args.path ? await resolveWorkspacePath(String(args.path)) : videoBridge?.getState()?.filePath;

        if (name === 'video_info') {
          if (!filePath) return { ok: false, message: 'Укажите path видео или откройте его во вкладке' };
          const info = await api?.videoProbe?.(filePath);
          if (!info?.success) return { ok: false, message: info?.error || 'Не удалось прочитать видео' };
          return {
            ok: true,
            message: `${Math.round(info.duration ?? 0)} с, ${info.width ?? 0}×${info.height ?? 0}, звук: ${
              info.hasAudio ? 'есть' : 'нет'
            }`,
            changed: false,
            meta: info,
          };
        }

        if (name === 'video_operate') {
          if (!filePath) return { ok: false, message: 'Укажите path видео или откройте его во вкладке' };
          const plan = {
            start: args.start !== undefined ? Number(args.start) : undefined,
            end: args.end !== undefined ? Number(args.end) : undefined,
            speed: args.speed !== undefined ? Number(args.speed) : undefined,
            speedSegments: Array.isArray(args.speed_segments)
              ? (args.speed_segments as { start: number; end: number; speed: number }[])
              : undefined,
            stripAudio: args.strip_audio === true || args.strip_audio === 'true',
            compress: args.compress === true || args.compress === 'true',
            zoom:
              args.zoom_scale !== undefined || args.zoom_x !== undefined || args.zoom_y !== undefined
                ? {
                    scale: args.zoom_scale !== undefined ? Number(args.zoom_scale) : undefined,
                    x: args.zoom_x !== undefined ? Number(args.zoom_x) : undefined,
                    y: args.zoom_y !== undefined ? Number(args.zoom_y) : undefined,
                  }
                : undefined,
            outputPath: args.output ? String(args.output) : undefined,
          };
          if (videoBridge) {
            const result = await videoBridge.applyPlan(plan);
            return { ok: result.ok, message: result.message, path: result.outputPath, changed: result.ok };
          }
          const info = await api?.videoProbe?.(filePath);
          if (!info?.success) return { ok: false, message: info?.error || 'Не удалось прочитать видео' };
          const result = await api?.videoEdit?.({
            filePath,
            duration: info.duration,
            hasAudio: info.hasAudio,
            ...plan,
          });
          if (!result?.success) return { ok: false, message: result?.error || 'Не удалось обработать видео' };
          if (result.outputPath) await ws.openFile(result.outputPath);
          return {
            ok: true,
            message: `Готово: ${result.outputPath ? baseName(result.outputPath) : 'файл сохранён'}`,
            path: result.outputPath,
            changed: true,
          };
        }

        if (name === 'video_extract_audio') {
          if (!filePath) return { ok: false, message: 'Укажите path видео' };
          const result = await api?.videoExtractAudio?.(filePath, {
            format: (args.format as 'mp3' | 'aac' | 'copy') ?? 'mp3',
            outputPath: args.output ? await resolveWorkspacePath(String(args.output)) : undefined,
          });
          if (!result?.success) return { ok: false, message: result?.error || 'Не удалось извлечь звук' };
          return { ok: true, message: `Звук извлечён: ${baseName(result.outputPath ?? '')}`, path: result.outputPath, changed: true };
        }

        if (name === 'video_frames') {
          if (!filePath) return { ok: false, message: 'Укажите path видео' };
          const result = await api?.videoFrames?.(filePath, {
            count: args.count !== undefined ? Number(args.count) : undefined,
            at: Array.isArray(args.at) ? (args.at as unknown[]).map(Number) : undefined,
            maxWidth: args.max_width !== undefined ? Number(args.max_width) : undefined,
            dir: args.dir ? await resolveWorkspacePath(String(args.dir)) : undefined,
          });
          if (!result?.success) return { ok: false, message: result?.error || 'Не удалось сделать кадры' };
          return {
            ok: true,
            message: `Кадров: ${result.files?.length ?? 0}`,
            path: result.files?.[0],
            content: (result.files ?? []).join('\n'),
            changed: true,
          };
        }

        if (name === 'video_convert') {
          if (!filePath) return { ok: false, message: 'Укажите path видео' };
          const result = await api?.videoConvert?.(filePath, {
            format: (args.format as 'mp4' | 'webm' | 'gif') ?? 'mp4',
            quality: (args.quality as 'low' | 'medium' | 'high') ?? 'medium',
            outputPath: args.output ? await resolveWorkspacePath(String(args.output)) : undefined,
          });
          if (!result?.success) return { ok: false, message: result?.error || 'Не удалось конвертировать видео' };
          if (result.outputPath) await ws.openFile(result.outputPath);
          return { ok: true, message: `Готово: ${baseName(result.outputPath ?? '')}`, path: result.outputPath, changed: true };
        }

        if (name === 'video_gif') {
          if (!filePath) return { ok: false, message: 'Укажите path видео' };
          const result = await api?.videoGif?.(filePath, {
            start: args.start !== undefined ? Number(args.start) : 0,
            duration: args.duration !== undefined ? Number(args.duration) : 3,
            fps: args.fps !== undefined ? Number(args.fps) : undefined,
            width: args.width !== undefined ? Number(args.width) : undefined,
            outputPath: args.output ? await resolveWorkspacePath(String(args.output)) : undefined,
          });
          if (!result?.success) return { ok: false, message: result?.error || 'Не удалось сделать GIF' };
          if (result.outputPath) await ws.openFile(result.outputPath);
          return { ok: true, message: `GIF готов: ${baseName(result.outputPath ?? '')}`, path: result.outputPath, changed: true };
        }

        if (name === 'video_publish') {
          if (!filePath) return { ok: false, message: 'Укажите path видео' };
          const title = String(args.title ?? baseName(filePath));
          const result = await api?.youtubeUpload?.(filePath, {
            title,
            description: String(args.description ?? ''),
            privacyStatus: ((args.privacy as 'private' | 'unlisted' | 'public') ?? 'unlisted') as
              | 'private'
              | 'unlisted'
              | 'public',
            categoryId: args.category !== undefined ? String(args.category) : undefined,
            playlistId: args.playlist !== undefined ? String(args.playlist) : undefined,
            madeForKids: false,
          });
          if (!result?.success) return { ok: false, message: result?.error || 'Не удалось опубликовать видео' };
          return {
            ok: true,
            message: `Опубликовано: ${(result as { url?: string }).url ?? title}`,
            path: (result as { url?: string }).url,
            changed: false,
          };
        }
      }

      if (name.startsWith('md_')) {
        if (!['markdown', 'richtext', 'text'].includes(tab.type)) {
          return { ok: false, message: 'Markdown-инструменты доступны только для текстовых документов' };
        }
        const content = String(tab.content ?? '');
        if (name === 'md_read') {
          return {
            ok: true,
            message: `Файл «${tab.name}», символов: ${content.length}`,
            changed: false,
            content,
            structure: content.slice(0, 12000),
          };
        }
        if (name === 'md_replace') {
          const find = String(args.find ?? '');
          const replace = String(args.replace ?? '');
          if (!find) return { ok: false, message: 'Не указан текст для поиска' };
          const all = args.all === true || args.all === 'true';
          const next = all
            ? content.split(find).join(replace)
            : content.replace(find, replace);
          if (next === content) {
            return { ok: false, message: `Текст «${find}» не найден` };
          }
          current.updateTabContent(tab.id, next);
          return { ok: true, message: all ? 'Все вхождения заменены' : 'Вхождение заменено', changed: true };
        }
        if (name === 'md_append') {
          const text = String(args.text ?? args.content ?? '');
          if (!text.trim()) return { ok: false, message: 'Пустой текст для добавления' };
          const next = `${content.replace(/\s*$/, '')}\n\n${text}`;
          current.updateTabContent(tab.id, next);
          return { ok: true, message: 'Текст добавлен в конец', changed: true };
        }
        if (name === 'md_replace_document') {
          const text = String(args.text ?? args.content ?? '');
          if (!text.trim()) return { ok: false, message: 'Пустой текст документа' };
          current.updateTabContent(tab.id, text);
          return { ok: true, message: 'Документ заменён', changed: true };
        }
        if (name === 'md_set_heading') {
          const result = setMarkdownHeading(
            content,
            String(args.find ?? ''),
            Number(args.level ?? 2)
          );
          if (!result.ok) return { ok: false, message: result.message };
          current.updateTabContent(tab.id, result.text);
          return { ok: true, message: result.message, changed: true };
        }
        if (name === 'md_toggle_list') {
          const result = toggleMarkdownList(
            content,
            String(args.find ?? ''),
            args.style === 'numbered' ? 'numbered' : 'bullet'
          );
          if (!result.ok) return { ok: false, message: result.message };
          current.updateTabContent(tab.id, result.text);
          return { ok: true, message: result.message, changed: true };
        }
        if (name === 'md_insert_link') {
          const result = insertMarkdownLink(
            content,
            String(args.find ?? ''),
            String(args.text ?? args.label ?? ''),
            String(args.url ?? '')
          );
          if (!result.ok) return { ok: false, message: result.message };
          current.updateTabContent(tab.id, result.text);
          return { ok: true, message: result.message, changed: true };
        }
        if (name === 'md_insert_table') {
          const headers = Array.isArray(args.headers)
            ? args.headers.map((value) => String(value))
            : [];
          const rows = Array.isArray(args.rows)
            ? args.rows.map((row) => (Array.isArray(row) ? row.map((value) => String(value)) : []))
            : [];
          const result = insertMarkdownTable(content, headers, rows);
          if (!result.ok) return { ok: false, message: result.message };
          current.updateTabContent(tab.id, result.text);
          return { ok: true, message: result.message, changed: true };
        }
        return { ok: false, message: `Неизвестный Markdown-инструмент: ${name}` };
      }

      if (
        !BLOCK_TOOL_TAB_TYPES.has(tab.type) &&
        !['get_document', 'get_document_info', 'replace_document', 'find_replace', 'search_document', 'read_document'].includes(name)
      ) {
        return {
          ok: false,
          message: `Инструмент ${name} недоступен для «${tab.name}»: для текста доступны только get_document, replace_document и find_replace.`,
        };
      }

      try {
        const controller = getDocumentController(tab.id);
        if (tab.type === 'word' && controller) {
          if (!docTxnRef.current || docTxnRef.current.id !== transactionId) {
            docTxnRef.current = { id: transactionId, html: controller.readHtml() };
          }
        }

        // Колонтитулы меняются в OOXML-пакете, а не в HTML-полотне.
        if (name === 'set_header_footer' || name === 'delete_header_footer') {
          const kindRaw = String(args.kind ?? args.part ?? 'header').toLowerCase();
          if (kindRaw !== 'header' && kindRaw !== 'footer') {
            return { ok: false, message: 'kind должен быть "header" или "footer"' };
          }
          const typeRaw = String(args.type ?? args.reference_type ?? args.referenceType ?? 'default').toLowerCase();
          const referenceType =
            typeRaw === 'first' || typeRaw === 'even' ? typeRaw : 'default';
          const remove =
            name === 'delete_header_footer' || args.remove === true || args.remove === 'true';
          const text = String(args.text ?? args.content ?? '');

          if (!remove && !text.trim()) {
            return { ok: false, message: 'Пустой текст колонтитула' };
          }
          if (!controller?.setHeaderFooter) {
            return {
              ok: false,
              message: 'Колонтитулы доступны только в документе Word',
            };
          }

          const edit = await controller.setHeaderFooter({
            kind: kindRaw,
            referenceType,
            text,
            remove,
          });
          return edit;
        }

        // Запись в другой файл — только после явного разрешения в панели AI.
        if (name === 'write_document') {
          if (!aiExtraWriteRef.current) {
            return {
              ok: false,
              message:
                'Запись в дополнительные файлы запрещена. Включите «Доп. файлы: запись» в панели AI.',
            };
          }
          const requested = String(args.path ?? args.file ?? '');
          const root =
            current.rootPath || (await fileSystemService.getCurrentWorkspace()) || '';
          const targetPath = resolvePathInsideWorkspace(root, requested);
          if (!targetPath) {
            return { ok: false, message: `Путь «${requested}» вне рабочей папки` };
          }

          const ext = targetPath.split('.').pop()?.toLowerCase() ?? '';
          let payload = '';
          let encoding = 'utf-8';
          if (ext === 'docx') {
            const html = String(
              args.html ??
                markdownToDocumentHtml(String(args.text ?? args.content ?? ''))
            );
            payload = await htmlToDocxBase64(html);
            encoding = 'base64';
          } else if (['md', 'markdown', 'txt', 'text'].includes(ext)) {
            payload = String(args.text ?? args.content ?? '');
          } else if (ext === 'xlsx' || ext === 'xlsm') {
            const columns = Array.isArray(args.columns)
              ? args.columns.map((value) => String(value))
              : [];
            const rows = Array.isArray(args.rows)
              ? args.rows.map((row) =>
                  Array.isArray(row)
                    ? row.map((value) =>
                        typeof value === 'number' || typeof value === 'boolean'
                          ? value
                          : String(value ?? '')
                      )
                    : []
                )
              : [];
            if (columns.length === 0 && rows.length === 0) {
              return { ok: false, message: 'Для .xlsx нужны columns или rows' };
            }
            payload = generateExcelBase64(
              columns,
              rows,
              String(args.sheet ?? 'Лист 1')
            );
            encoding = 'base64';
          } else {
            return {
              ok: false,
              message: `Запись в «${ext || 'без расширения'}» пока не поддерживается: доступны .docx, .xlsx, .md и .txt`,
            };
          }

          try {
            await fileSystemService.writeFile(targetPath, payload, encoding);
          } catch (error) {
            return {
              ok: false,
              message: `Не удалось записать файл: ${
                error instanceof Error ? error.message : String(error)
              }`,
            };
          }

          const openTab = current.tabs.find((item) => item.path === targetPath);
          if (openTab) current.updateTabContent(openTab.id, payload);
          else await current.refreshTree();
          return { ok: true, message: `Файл «${requested}» обновлён`, changed: true };
        }

        // Служебные части Word: колонтитулы, сноски, комментарии, стили.
        if (name === 'get_document_info') {
          const info = await controller?.getDocumentInfo?.();
          if (!info) {
            return {
              ok: false,
              message: 'Информация о пакете доступна только для документа Word',
            };
          }
          const lines = [
            `Документ: ${tab.name} (documentId: ${tab.id})`,
            `Частей в пакете: ${info.parts.length}`,
            `Секций: ${info.sections}`,
            info.styleNames.length
              ? `Стили: ${info.styleNames.slice(0, 40).join(', ')}`
              : 'Именованных стилей нет',
            info.headers.length
              ? `Верхние колонтитулы: ${info.headers.map((part) => `${part.name} — ${part.text || '(пусто)'}`).join(' | ')}`
              : 'Верхних колонтитулов нет',
            info.footers.length
              ? `Нижние колонтитулы: ${info.footers.map((part) => `${part.name} — ${part.text || '(пусто)'}`).join(' | ')}`
              : 'Нижних колонтитулов нет',
            info.footnotes ? `Сноски: ${info.footnotes}` : 'Сносок нет',
            info.comments ? `Комментарии: ${info.comments}` : 'Комментариев нет',
          ];
          return {
            ok: true,
            message: 'Информация о документе получена',
            changed: false,
            structure: lines.join('\n'),
            meta: {
              documentId: tab.id,
              sections: info.sections,
              headers: info.headers,
              footers: info.footers,
              footnotes: info.footnotes,
              comments: info.comments,
              styleNames: info.styleNames,
              parts: info.parts,
            },
          };
        }

        // Другой файл workspace — только чтение. Запись в него не входит в
        // инструменты: доп. документы по умолчанию read-only.
        if (name === 'read_document') {
          const requested = String(args.path ?? args.file ?? args.name ?? '');
          const root =
            current.rootPath || (await fileSystemService.getCurrentWorkspace()) || '';
          const targetPath = resolvePathInsideWorkspace(root, requested);
          if (!targetPath) {
            return {
              ok: false,
              message: `Путь «${requested}» вне рабочей папки или не указан. Доступны только файлы workspace.`,
            };
          }

          const ext = targetPath.split('.').pop()?.toLowerCase() ?? '';
          const type = getFileType(ext);
          const binary = ['word', 'excel', 'powerpoint'].includes(type);
          let read: { content: string; isBinary: boolean };
          try {
            read = await fileSystemService.readFile(targetPath, binary ? 'base64' : 'utf-8');
          } catch (error) {
            return {
              ok: false,
              message: `Не удалось прочитать «${requested}»: ${error instanceof Error ? error.message : String(error)}`,
            };
          }

          let text = '';
          if (ext === 'docx') text = await docxText(read.content);
          else if (ext === 'pptx') text = await pptxText(read.content);
          else if (ext === 'doc') text = await extractLegacyDocText(read.content);
          else if (type === 'excel') {
            text = await tabToDocumentText({
              type: 'excel',
              content: read.content,
              encoding: 'base64',
            } as OpenTab);
          } else if (['markdown', 'richtext', 'text', 'code'].includes(type)) {
            text = String(read.content);
          } else {
            return {
              ok: false,
              message: `Формат «${ext || 'без расширения'}» не поддерживается для чтения: доступны Word, Excel, PowerPoint, Markdown и текст.`,
            };
          }

          if (!text.trim()) return { ok: false, message: `Файл «${requested}» пуст или не разобран` };
          const limit = 30000;
          const truncated = text.length > limit;
          return {
            ok: true,
            message: `Прочитан файл «${requested}»${truncated ? ` (первые ${limit} символов)` : ''}. Только чтение.`,
            changed: false,
            content: truncated ? `${text.slice(0, limit)}\n…` : text,
            meta: { path: targetPath, truncated, length: text.length },
          };
        }

        // Картинка из файла workspace превращается в data URL: документ
        // хранит изображения внутри .docx, а не ссылками на диск.
        if (name === 'insert_image' && typeof args.src === 'string' && args.src.trim()) {
          const root =
            current.rootPath || (await fileSystemService.getCurrentWorkspace()) || '';
          const resolved = await resolveImageSource(args.src, root);
          if (!resolved.ok) return { ok: false, message: resolved.message };
          args = { ...args, src: resolved.dataUrl };
        }

        // Выделение живёт только в живом полотне: читаем и правим его без
        // пересборки HTML всего документа.
        if (
          name === 'get_selection' ||
          name === 'replace_selection' ||
          name === 'insert_at_cursor'
        ) {
          const selection = controller?.getSelection?.() ?? null;
          if (name === 'get_selection') {
            if (!selection || !selection.text) {
              return {
                ok: false,
                message: 'Нет выделенного текста в документе: попросите пользователя выделить фрагмент',
              };
            }
            return {
              ok: true,
              message: `Выделение: блок ${selection.blockId}, символы ${selection.startOffset}–${selection.endOffset}`,
              changed: false,
              content: selection.text,
              selection,
            };
          }

          const raw = String(args.html ?? args.content ?? args.text ?? '');
          if (!raw.trim()) return { ok: false, message: 'Пустой фрагмент для правки выделения' };

          const edit =
            name === 'replace_selection'
              ? controller?.replaceSelection?.(raw, transactionId)
              : controller?.insertAtSelection?.(raw, transactionId);

          if (!edit) {
            return {
              ok: false,
              message: 'Редактор не поддерживает работу с выделением: откройте документ Word',
            };
          }
          return { ok: edit.ok, message: edit.message, changed: edit.changed };
        }

        let html = '';

        if (tab.type === 'word') {
          html = controller?.readHtml() || (await docxBase64ToHtml(tab.content)).html;
        } else {
          html = markdownToDocumentHtml(tab.content);
        }

        const result = applyDocumentTool(html, name, args);
        if (!result.ok) {
          if (
            tab.type === 'word' &&
            controller?.restoreHtml &&
            docTxnRef.current?.id === transactionId
          ) {
            controller.restoreHtml(docTxnRef.current.html);
          }
          return result;
        }

        // Поля, размер страницы и базовый шрифт живут в .docx, а не в HTML:
        // передаём их живому редактору, он пересоберёт пакет.
        if (result.page && tab.type === 'word') {
          getDocumentController(tab.id)?.applyPageSetup?.(result.page);
        }

        if (result.changed && result.html !== undefined) {
          if (tab.type === 'word') {
            if (controller) {
              controller.writeHtml(result.html, transactionId);
            } else {
              const base64 = await htmlToDocxBase64(result.html);
              current.updateTabContent(tab.id, base64);
            }
          } else {
            // Markdown/текст хранят содержимое строкой — HTML разбираем обратно.
            current.updateTabContent(tab.id, htmlToPlainText(result.html));
          }
        }

        return {
          ...result,
          structure: result.structure ? `Файл: ${tab.name}\n${result.structure}` : result.structure,
        };
      } catch (error) {
        return {
          ok: false,
          message: `Не удалось применить ${name}: ${(error as Error).message}`,
        };
      }
    },
    [resolveWorkspacePath, exportTabAs]
  );

  const handleAIRunAction = useCallback(
    async (action: AIActionPayload): Promise<AppActionResult> => {
      const fail = (message: string): AppActionResult => ({ ok: false, message });

      try {
        switch (action.action) {
          case 'import_table': {
            const sourcePath = await resolveTableTarget(action.source);
            const ext = extName(sourcePath).toLowerCase();
            const targetPath = action.target
              ? await resolveTableTarget(action.target)
              : sourcePath.replace(/\.[^.\\/]+$/, '') + '.xlsx';

            if (ext === 'xlsx' || ext === 'xls' || ext === 'xlsm') {
              const res = await fileSystemService.readFile(sourcePath, 'base64');
              try {
                await fileSystemService.writeFile(targetPath, res.content, 'base64');
              } catch (err) {
                return fail(
                  `Не удалось записать ${baseName(targetPath)}: закройте файл в Excel и повторите. ${(err as Error).message}`
                );
              }
              await ws.refreshTree();
              const openTab = ws.tabs.find((item) => item.path === targetPath);
              if (openTab) {
                ws.patchTab(openTab.id, {
                  content: res.content,
                  initialContent: res.content,
                  isDirty: false,
                  reloadKey: Date.now(),
                });
              } else {
                await ws.openFile(targetPath);
              }
              return {
                ok: true,
                message: `Таблица скопирована: ${baseName(sourcePath)} → ${baseName(targetPath)}`,
                path: targetPath,
              };
            }

            if (ext !== 'md' && ext !== 'markdown' && ext !== 'txt' && ext !== 'csv' && ext !== 'tsv') {
              return fail(`Импорт из .${ext || 'файла'} не поддерживается`);
            }

            const read = await fileSystemService.readFile(sourcePath, 'utf-8');
            let columns: string[] = [];
            let rows: (string | number)[][] = [];
            const table = detectMarkdownTables(read.content)
              .sort((a, b) => b.rows.length - a.rows.length)[0];
            if (table) {
              columns = table.columns;
              rows = table.rows.map((row) => row.map((cell) => normalizeTableValue(cell) as string | number));
            } else {
              const parsed = XLSX.read(read.content, { type: 'string' });
              const sheet = parsed.Sheets[parsed.SheetNames[0]];
              const matrix = XLSX.utils.sheet_to_json(sheet, {
                header: 1,
                defval: '',
                blankrows: false,
              }) as (string | number)[][];
              if (matrix.length > 0) {
                columns = matrix[0].map((cell) => String(cell ?? ''));
                rows = matrix
                  .slice(1)
                  .map((row) =>
                    row.map(
                      (cell) => normalizeTableValue(cell as string | number) as string | number
                    )
                  );
              }
            }
            if (columns.length === 0 || rows.length === 0) {
              return fail(`В файле ${baseName(sourcePath)} не найдена таблица`);
            }

            const sheetName = action.sheetName || 'Лист 1';
            let base64 = generateExcelBase64(columns, rows, sheetName);
            try {
              const existing = await fileSystemService.readFile(targetPath, 'base64');
              const workbook = XLSX.read(existing.content, { type: 'base64' });
              const sheet = XLSX.utils.aoa_to_sheet([columns, ...rows]);
              applyAutoFitColumns(sheet, [columns, ...rows], { headerRows: 1 });
              workbook.Sheets[sheetName] = sheet;
              if (!workbook.SheetNames.includes(sheetName)) workbook.SheetNames.push(sheetName);
              base64 = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
            } catch {
              // целевого файла ещё нет — оставляем новую книгу
            }
            try {
              await fileSystemService.writeFile(targetPath, base64, 'base64');
            } catch (err) {
              return fail(
                `Не удалось записать ${baseName(targetPath)}: закройте файл в Excel и повторите. ${(err as Error).message}`
              );
            }
            await ws.refreshTree();
            const openTab = ws.tabs.find((item) => item.path === targetPath);
            if (openTab) {
              ws.patchTab(openTab.id, {
                content: base64,
                initialContent: base64,
                isDirty: false,
                reloadKey: Date.now(),
              });
            } else {
              await ws.openFile(targetPath);
            }
            return {
              ok: true,
              message: `Импортировано ${rows.length} строк из ${baseName(
                sourcePath
              )} в ${baseName(targetPath)}`,
              path: targetPath,
            };
          }

          case 'create_table': {
            const isCsv = action.format === 'csv' || /\.csv$/i.test(action.name || '');
            const name = sanitizeFileName(action.name || 'Таблица', isCsv ? '.csv' : '.xlsx');
            const dir = await resolveWorkspacePath(action.path);
            const content = isCsv
              ? generateCsvContent(action.columns || [], action.rows || [])
              : generateExcelBase64(
                  action.columns || [],
                  action.rows || [],
                  action.sheetName || 'Лист 1'
                );
            const created = await ws.createFile(dir, name, content, isCsv ? 'utf-8-bom' : 'base64');
            return created
              ? { ok: true, message: `Создана таблица: ${name}`, path: created }
              : fail(`Не удалось создать таблицу: ${name}`);
          }

          case 'edit_table': {
            const rawTarget = action.path || action.name;
            const targetPath = rawTarget
              ? await resolveTableTarget(rawTarget)
              : ws.activeTab?.path;
            if (!targetPath) return fail('Не указана таблица для изменения');
            const ext = extName(targetPath).toLowerCase();
            const tableRows = (action.rows ?? []).map((row) =>
              row.map((cell) => normalizeTableValue(cell))
            );
            if (ext === 'csv' || ext === 'tsv') {
              const read = await fileSystemService
                .readFile(targetPath, 'utf-8')
                .catch(() => null);
              const workbook = read
                ? XLSX.read(read.content, { type: 'string' })
                : XLSX.utils.book_new();
              if (workbook.SheetNames.length === 0) {
                XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[]]), 'Лист 1');
              }
              const sheetName = workbook.SheetNames[0];
              const existing = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
                header: 1,
                defval: '',
                blankrows: false,
              }) as XLSX.CellObject[][];
              const base =
                action.mode === 'append'
                  ? [...existing, ...tableRows]
                  : [...(action.columns?.length ? [action.columns] : []), ...tableRows];
              const sheet = XLSX.utils.aoa_to_sheet(base as unknown[][]);
              const csv = XLSX.utils.sheet_to_csv(sheet);
              await fileSystemService.writeFile(targetPath, csv, 'utf-8-bom');
              const openTab = ws.tabs.find((item) => item.path === targetPath);
              if (openTab) ws.patchTab(openTab.id, { content: csv, initialContent: csv, isDirty: false, reloadKey: Date.now() });
              else await ws.openFile(targetPath);
              return { ok: true, message: `Таблица обновлена: ${baseName(targetPath)}`, path: targetPath };
            }

            const read = await fileSystemService
              .readFile(targetPath, 'base64')
              .catch(() => null);
            const workbook = read
              ? XLSX.read(read.content, { type: 'base64', bookVBA: ext === 'xlsm' })
              : XLSX.utils.book_new();
            if (workbook.SheetNames.length === 0) {
              XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[]]), 'Лист 1');
            }
            const sheetName =
              action.sheetName && workbook.SheetNames.includes(action.sheetName)
                ? action.sheetName
                : workbook.SheetNames[0];
            if (!sheetName) return fail('В книге нет листов');
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
              header: 1,
              defval: '',
              blankrows: false,
            }) as XLSX.CellObject[][];
            const nonEmpty = rows.filter((row) =>
              row.some((cell) => String(cell ?? '').trim() !== '')
            );
            const base =
              action.mode === 'append'
                ? [...nonEmpty, ...tableRows]
                : [...(action.columns?.length ? [action.columns] : []), ...tableRows];
            workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(base as unknown[][]);
            applyAutoFitColumns(workbook.Sheets[sheetName], base as unknown[][], {
              headerRows: action.mode === 'append' ? 0 : 1,
            });
            const content = XLSX.write(workbook, {
              bookType: ext === 'xlsm' ? 'xlsm' : 'xlsx',
              type: 'base64',
              bookVBA: ext === 'xlsm',
            });
            await fileSystemService.writeFile(targetPath, content, 'base64');
            const openTab = ws.tabs.find((item) => item.path === targetPath);
            if (openTab) {
              ws.patchTab(openTab.id, {
                content,
                initialContent: content,
                isDirty: false,
                reloadKey: Date.now(),
              });
            } else {
              await ws.openFile(targetPath);
            }
            return {
              ok: true,
              message: `${
                action.mode === 'append' ? 'Добавлено строк' : 'Заменено'
              }: ${action.rows?.length ?? 0} — ${baseName(targetPath)}`,
              path: targetPath,
            };
          }

          case 'create_document': {            const name = sanitizeFileName(
              action.name || (action.title ? `${action.title}.docx` : 'Документ'),
              '.docx'
            );
            const dir = await resolveWorkspacePath(action.path);
            const title = action.title || name.replace(/\.docx?$/i, '');
            const base64 = await generateDocxBase64(action.content || '', title);
            const created = await ws.createFile(dir, name, base64, 'base64');
            return created
              ? { ok: true, message: `Создан документ Word: ${name}`, path: created }
              : fail(`Не удалось создать документ: ${name}`);
          }

          case 'create_file': {
            const name = sanitizeFileName(action.name || 'Файл', '.txt');
            const dir = await resolveWorkspacePath(action.path);
            const created = await ws.createFile(
              dir,
              name,
              action.content || '',
              /\.(csv|tsv)$/i.test(name) ? 'utf-8-bom' : 'utf-8'
            );
            return created
              ? { ok: true, message: `Создан файл: ${name}`, path: created }
              : fail(`Не удалось создать файл: ${name}`);
          }

          case 'open_file': {
            if (!action.path) return fail('Не указан путь к файлу');
            const target = await resolveWorkspacePath(action.path);
            const tabId = await ws.openFile(target);
            return tabId
              ? { ok: true, message: `Открыт файл: ${baseName(target)}`, path: target }
              : fail(`Файл не открыт: ${baseName(target)}`);
          }

          case 'open_folder': {
            if (!action.path) {
              const picked = await ws.chooseWorkspaceFolder();
              return picked
                ? { ok: true, message: `Открыта папка: ${picked}`, path: picked }
                : fail('Выбор папки отменён');
            }
            const target = await resolveWorkspacePath(action.path);
            await ws.openWorkspace(target);
            return { ok: true, message: `Открыта папка: ${baseName(target)}`, path: target };
          }

          case 'create_folder': {
            const target = await resolveWorkspacePath(action.path || action.name);
            await fileSystemService.createDirectory(target);
            await ws.refreshTree();
            return { ok: true, message: `Создана папка: ${baseName(target)}`, path: target };
          }

          case 'rename_path': {
            if (!action.path) return fail('Не указан путь для переименования');
            const target = await resolveWorkspacePath(action.path);
            const newName = (action.newName || action.newPath || '')
              .replace(/[\\/]+$/, '')
              .split(/[\\/]/)
              .pop();
            if (!newName) return fail('Не указано новое имя');
            const renamed = await ws.renamePath(target, newName);
            return renamed
              ? { ok: true, message: `Переименовано в: ${newName}` }
              : fail(`Не удалось переименовать: ${baseName(target)}`);
          }

          case 'delete_path': {
            if (!action.path) return fail('Не указан путь для удаления');
            const target = await resolveWorkspacePath(action.path);
            const deleted = await ws.deletePath(target);
            return deleted
              ? { ok: true, message: `Удалено: ${baseName(target)}` }
              : fail(`Не удалось удалить: ${baseName(target)}`);
          }

          case 'insert_into_editor': {
            if (!ws.activeTab) return fail('Нет активного документа для вставки');
            insertIntoEditor(action.text);
            return { ok: true, message: `Текст вставлен в «${ws.activeTab.name}»` };
          }

          case 'replace_document': {
            if (!ws.activeTab) return fail('Нет активного документа для замены');
            await replaceDocument(action.text);
            return { ok: true, message: `Документ «${ws.activeTab.name}» заменён` };
          }

          case 'document_apply': {
            if (!wsRef.current.activeTab) return fail('Нет активного документа');
            const transactionId = `action_${Date.now()}`;
            let applied = 0;

            for (const operation of action.operations) {
              const { tool, ...args } = operation;
              if (!tool) continue;
              const result = await runDocumentTool(tool, args, transactionId);
              if (!result.ok) {
                return fail(`Операция «${tool}»: ${result.message}`);
              }
              applied += 1;
            }

            return {
              ok: true,
              message: `Выполнено операций с документом: ${applied}`,
            };
          }

          case 'open_settings': {
            setSettingsTab((action.tab as SettingsTab) || 'provider');
            setSettingsOpen(true);
            return { ok: true, message: 'Настройки открыты' };
          }

          case 'toggle_panel': {
            const state = action.state || 'toggle';
            const next = (current: boolean) =>
              state === 'open' ? true : state === 'close' ? false : !current;

            switch (action.panel) {
              case 'sidebar':
                setSidebarOpen(next(sidebarOpen));
                break;
              case 'ai':
                setAiPanelOpen(next(aiPanelOpen));
                break;
              case 'terminal':
                setTerminalOpen(next(terminalOpen));
                break;
              case 'search':
                setSidebarOpen(true);
                setSearchPanelOpen(next(searchPanelOpen));
                break;
              default:
                return fail(`Неизвестная панель: ${String(action.panel)}`);
            }
            const labels: Record<string, string> = {
              sidebar: 'Проводник',
              ai: 'Панель AI',
              terminal: 'Терминал',
              search: 'Поиск',
            };
            return { ok: true, message: `${labels[action.panel]}: ${state === 'close' ? 'закрыто' : 'готово'}` };
          }

          case 'select_tab': {
            const tab = findOpenTab(action);
            if (!tab) return fail('Вкладка не найдена');
            ws.setActiveTabId(tab.id);
            return { ok: true, message: `Активная вкладка: ${tab.name}`, path: tab.path };
          }

          case 'close_tab': {
            const tab = findOpenTab(action);
            if (!tab) return fail('Вкладка не найдена');
            if (tab.isDirty) {
              ws.setActiveTabId(tab.id);
              setUnsavedDialog({ open: true, tabId: tab.id });
              return fail(`«${tab.name}» содержит несохранённые изменения — подтвердите закрытие`);
            }
            ws.closeTab(tab.id);
            return { ok: true, message: `Вкладка закрыта: ${tab.name}` };
          }

          case 'save': {
            if (action.all) {
              const { saved, failed } = await ws.saveAll();
              return failed.length === 0
                ? { ok: true, message: `Сохранено файлов: ${saved}` }
                : fail(`Не сохранено: ${failed.join(', ')}`);
            }
            if (!ws.activeTabId) return fail('Нет активной вкладки');
            const saved = await ws.saveTab(ws.activeTabId);
            return saved
              ? { ok: true, message: `Сохранено: ${ws.activeTab?.name || 'файл'}` }
              : fail('Не удалось сохранить файл');
          }

          case 'save_as': {
            const result = await runDocumentTool('save_as', {
              format: action.format,
              path: action.path,
              target: action.target,
              name: action.name,
            });
            return {
              ok: result.ok,
              message: result.message,
              path:
                result.meta && typeof result.meta.path === 'string'
                  ? result.meta.path
                  : undefined,
            };
          }

          case 'convert_file': {
            const result = await runDocumentTool('convert_file', {
              source: action.source,
              path: action.source,
              target: action.target,
              format: action.format,
            });
            return {
              ok: result.ok,
              message: result.message,
              path:
                result.meta && typeof result.meta.path === 'string'
                  ? result.meta.path
                  : undefined,
            };
          }

          case 'create_presentation': {
            const markdown = action.markdown ?? '';
            if (!markdown.trim()) return fail('Пустая разметка презентации');
            const fileName = sanitizeFileName(action.name || 'Презентация.pptx', '.pptx');
            const dir = await resolveWorkspacePath(action.path || '');
            const base64 = await markdownToPptxBase64(markdown, {
              title: action.title || baseName(fileName).replace(/\.pptx$/i, ''),
              themePresetId: action.theme,
            });
            const created = await ws.createFile(dir, fileName, base64, 'base64');
            if (!created) return fail(`Не удалось создать ${fileName}`);
            await ws.openFile(created);
            return { ok: true, message: `Создана презентация: ${fileName}`, path: created };
          }

          case 'image_edit': {
            const result = await runDocumentTool('image_transform', {
              path: action.path,
              ops: action.ops,
              output: action.output,
              suffix: action.suffix,
            });
            return result.ok ? { ok: true, message: result.message, path: result.path } : fail(result.message);
          }

          case 'video_edit': {
            const toolName =
              action.operation === 'extract_audio'
                ? 'video_extract_audio'
                : action.operation === 'frames'
                  ? 'video_frames'
                  : action.operation === 'convert'
                    ? 'video_convert'
                    : action.operation === 'gif'
                      ? 'video_gif'
                      : 'video_operate';
            const params = (action.params ?? {}) as Record<string, unknown>;
            const result = await runDocumentTool(toolName, {
              ...params,
              path: action.path ?? (typeof params.path === 'string' ? params.path : undefined),
            });
            return result.ok ? { ok: true, message: result.message, path: result.path } : fail(result.message);
          }

          case 'notify': {
            ws.flash(action.message || 'Готово');
            return { ok: true, message: action.message || 'Готово' };
          }

          case 'run_command': {
            if (!action.command?.trim()) return fail('Не указана команда');
            const cwd = action.cwd ? await resolveWorkspacePath(action.cwd) : '';
            const command = cwd
              ? `Set-Location -LiteralPath '${cwd.replace(/'/g, "''")}'; ${action.command}`
              : action.command;
            setTerminalOpen(true);
            setTerminalCommand({ id: uid('cmd'), command });
            return { ok: true, message: `Команда отправлена в терминал: ${action.command}` };
          }

          default:
            return fail(`Действие не поддерживается: ${(action as { action: string }).action}`);
        }
      } catch (err) {
        const message = (err as Error).message;
        ws.flash(`Ошибка действия AI: ${message}`);
        return fail(message);
      }
    },
    [
      ws,
      resolveWorkspacePath,
      findOpenTab,
      insertIntoEditor,
      replaceDocument,
      runDocumentTool,
      sidebarOpen,
      aiPanelOpen,
      terminalOpen,
      searchPanelOpen,
      setSidebarOpen,
      setAiPanelOpen,
      setTerminalOpen,
    ]
  );

  /** Читает файл по абсолютному пути для вложения в чат. */
  const readAttachmentByPath = useCallback(
    async (filePath: string): Promise<ChatAttachment | null> => {
      const ext = extName(filePath).toLowerCase();
      try {
        if (ext === 'xlsx' || ext === 'xls' || ext === 'xlsm') {
          const res = await fileSystemService.readFile(filePath, 'base64');
          const workbook = XLSX.read(res.content, { type: 'base64' });
          const text = workbook.SheetNames.map(
            (name) => `### Лист: ${name}\n${XLSX.utils.sheet_to_csv(workbook.Sheets[name])}`
          ).join('\n\n');
          return {
            name: baseName(filePath),
            path: filePath,
            content: text.slice(0, 40000),
            truncated: text.length > 40000,
          };
        }
        if (ext === 'pptx') {
          const res = await fileSystemService.readFile(filePath, 'base64');
          const text = await pptxText(res.content);
          if (!text.trim()) return null;
          return {
            name: baseName(filePath),
            path: filePath,
            content: text.slice(0, 40000),
            truncated: text.length > 40000,
          };
        }
        if (ext === 'md' || ext === 'markdown') {
          const res = await fileSystemService.readFile(filePath, 'utf-8');
          return {
            name: baseName(filePath),
            path: filePath,
            content: res.content.slice(0, 40000),
            truncated: res.content.length > 40000,
          };
        }
        if (ext === 'docx' || ext === 'doc') {
          const res = await fileSystemService.readFile(filePath, 'base64');
          const text =
            ext === 'doc' ? await extractLegacyDocText(res.content) : await docxText(res.content);
          if (!text.trim()) return null;
          return {
            name: baseName(filePath),
            path: filePath,
            content: text.slice(0, 40000),
            truncated: text.length > 40000,
          };
        }
        if (
          ext === 'csv' ||
          ext === 'tsv' ||
          ext === 'txt' ||
          ext === 'log' ||
          ext === 'json' ||
          ext === 'xml' ||
          ext === 'yaml' ||
          ext === 'yml' ||
          ext === 'ini' ||
          ext === 'cfg' ||
          ext === 'js' ||
          ext === 'ts' ||
          ext === 'tsx' ||
          ext === 'jsx' ||
          ext === 'py' ||
          ext === 'css' ||
          ext === 'html'
        ) {
          const res = await fileSystemService.readFile(filePath, 'utf-8');
          return {
            name: baseName(filePath),
            path: filePath,
            content: res.content.slice(0, 40000),
            truncated: res.content.length > 40000,
          };
        }
        return null;
      } catch {
        return null;
      }
    },
    []
  );

  /** Файлы, перетащенные в чат из Проводника Windows. */
  const readDroppedChatFiles = useCallback(
    async (paths: string[]): Promise<ChatAttachment[]> => {
      const attachments: ChatAttachment[] = [];
      for (const filePath of paths.slice(0, 5)) {
        const attachment = await readAttachmentByPath(filePath);
        if (attachment) attachments.push(attachment);
      }
      return attachments;
    },
    [readAttachmentByPath]
  );

  /** Файл из сообщения чата: читаем содержимое, чтобы модель могла с ним работать. */
  const resolveChatAttachment = useCallback(
    async (
      rawName: string
    ): Promise<{ name: string; path: string; content: string } | null> => {
      const cleaned = rawName
        .trim()
        .replace(/^["'«]+|["'»]+$/g, '')
        .replace(/\\/g, '/');
      const fileName = cleaned.split('/').pop() || cleaned;
      const root = ws.rootPath || (await fileSystemService.getCurrentWorkspace()) || '';
      const normalized = cleaned.toLowerCase();
      const byPath =
        ws.allFiles.find(
          (file) =>
            file.name.toLowerCase() === fileName.toLowerCase() ||
            file.path.replace(/\\/g, '/').toLowerCase().endsWith(`/${normalized}`)
        ) ||
        ws.allFiles.find((file) =>
          file.path.replace(/\\/g, '/').toLowerCase().endsWith(normalized)
        );
      let target = byPath?.path || (root ? joinPath(root, cleaned) : '');
      if (!target) return null;
      return readAttachmentByPath(target);
    },
    [ws.rootPath, ws.allFiles, readAttachmentByPath]
  );

  /**
   * Текст активного документа для автоматического прикрепления к запросу.
   * Word-вкладка хранит base64 .docx — модели нужен разобранный текст.
   */
  const resolveActiveDocument = useCallback(async (): Promise<ChatAttachment | null> => {
    const tab = ws.activeTab;
    if (!tab) return null;
    const text = (await tabToDocumentText(tab)).trim();
    if (!text) return null;
    return {
      name: tab.name,
      path: tab.path,
      content: text.slice(0, 40000),
      truncated: text.length > 40000,
    };
  }, [ws.activeTab]);

  /** Снимок состояния приложения для системного промта AI. */
  const aiAppContext = useMemo(() => {
    const lines: string[] = [
      `Рабочая папка: ${ws.rootPath || 'не выбрана'}`,
      `Активная вкладка: ${ws.activeTab ? `${ws.activeTab.name} (${ws.activeTab.type})` : 'нет'}`,
    ];
    if (ws.tabs.length > 0) {
      lines.push(
        'Открытые вкладки:\n' +
          ws.tabs
            .map((tab) => `- ${tab.name}${tab.isDirty ? ' *не сохранён*' : ''}`)
            .join('\n')
      );
    }
    lines.push(
      `Панели: проводник ${sidebarOpen ? 'открыт' : 'закрыт'}, AI ${aiPanelOpen ? 'открыта' : 'закрыта'}, терминал ${terminalOpen ? 'открыт' : 'закрыт'}`
    );
    lines.push(`Модель: ${ai.settings.model || 'не выбрана'}`);
    lines.push(`Тема: ${isDark ? 'тёмная' : 'светлая'}`);
    if (ws.activeTab && DOCUMENT_TAB_TYPES.has(ws.activeTab.type)) {
      lines.push(
        `Открытый документ: ${ws.activeTab.name} · права записи: ${aiDocWrite ? 'есть' : 'нет'}`
      );
    } else {
      lines.push(`Права записи AI: ${aiDocWrite ? 'есть' : 'нет'}`);
    }
    return lines.join('\n');
  }, [
    ws.rootPath,
    ws.activeTab,
    ws.tabs,
    sidebarOpen,
    aiPanelOpen,
    terminalOpen,
    ai.settings.model,
    aiDocWrite,
    isDark,
  ]);

  /** Папка для диалога создания: по файлу берём его каталог. */
  const sanitizeParentDir = useCallback(
    (dir: string): string => dir || ws.rootPath || '',
    [ws.rootPath]
  );

  const handleExplorerRename = useCallback(
    async (filePath: string, newName: string): Promise<string | null> => {
      const dir = filePath.replace(/[\\/][^\\/]+$/, '');
      const target = `${dir}\\${newName}`;
      try {
        await fileSystemService.renamePath(filePath, target);
        await ws.loadDirectory(dir);
        return target;
      } catch (err) {
        ws.flash('Переименование не удалось: ' + (err as Error).message);
        return null;
      }
    },
    [ws]
  );

  const handleExplorerMove = useCallback(
    async (paths: string[], targetDir: string, copy: boolean): Promise<void> => {
      let done = 0;
      const failed: string[] = [];
      for (const source of paths) {
        const name = source.split(/[\\/]/).pop() ?? source;
        const existing = (ws.tree[targetDir] ?? []).map((item) => item.name);
        const unique = resolvePasteName(name, existing, copy ? 'copy' : 'move');
        const target = `${targetDir}\\${unique}`;
        try {
          if (copy) await fileSystemService.copyPath(source, target);
          else await fileSystemService.renamePath(source, target);
          done += 1;
        } catch (err) {
          failed.push(`${name}: ${(err as Error).message}`);
        }
      }
      const sourceDir = paths[0]?.replace(/[\\/][^\\/]+$/, '');
      await ws.loadDirectory(targetDir);
      if (sourceDir && sourceDir !== targetDir) await ws.loadDirectory(sourceDir);
      ws.flash(
        failed.length === 0
          ? `${copy ? 'Скопировано' : 'Перемещено'}: ${done}`
          : `${copy ? 'Скопировано' : 'Перемещено'} ${done} из ${paths.length}; ${failed.slice(0, 2).join('; ')}`
      );
    },
    [ws]
  );

  const handleFileAction = useCallback(async (action: string, item: FileItem, serviceId?: string, selection?: FileItem[]) => {
    if (action === 'open') {
      if (item.isDirectory) void ws.toggleDirectory(item.path);
      else void ws.openFile(item.path);
    } else if (action === 'preview') {
      if (!item.isDirectory) void ws.openFile(item.path, 'primary', false, true);
    } else if (action === 'new-file') {
      setNewFileDialog({
        open: true,
        parentDir: sanitizeParentDir(
          item.isDirectory ? item.path : item.path.replace(/[\\/][^\\/]+$/, '')
        ),
        kind: 'file',
        name: '',
      });
    } else if (action === 'new-folder') {
      setNewFileDialog({
        open: true,
        parentDir: sanitizeParentDir(
          item.isDirectory ? item.path : item.path.replace(/[\\/][^\\/]+$/, '')
        ),
        kind: 'folder',
        name: '',
      });
    } else if (action === 'favorite') {
      await ws.toggleFavorite(item.path);
    } else if (action === 'open-secondary') {
      void ws.openFile(item.path, 'secondary');
    } else if (action === 'open-new') {
      void ws.openFile(item.path, 'primary', true);
    } else if (action === 'open-external') {
      if (
        /\.(exe|msi|com|scr|bat|cmd|ps1|vbs|js|jar)$/i.test(item.path) &&
        !window.confirm(`Запустить файл?\n\n${item.path}`)
      ) {
        return;
      }
      try {
        await fileSystemService.openExternally(item.path);
        ws.flash(`Открыто в системе: ${item.name}`);
      } catch (err) {
        ws.flash(`Не удалось открыть: ${(err as Error).message}`);
      }
    } else if (action === 'rename') {
      setRenameDialog({ open: true, path: item.path, name: item.name, busy: false });
    } else if (action === 'delete') {
      if (window.confirm('Удалить «' + item.name + '»' + (item.isDirectory ? ' и всё его содержимое' : '') + '?')) await ws.deletePath(item.path);
    } else if (action === 'copy-path') {
      const ok = await copyToClipboard(item.path);
      if (ok) ws.flash('Путь скопирован'); else ws.flash('Не удалось скопировать путь');
    } else if (action === 'reveal') {
      await fileSystemService.showInExplorer(item.path);
    } else if (action === 'image-process') {
      const images = (selection || []).filter((it) => it.type === 'image' && !it.isDirectory);
      if (images.length === 0) {
        ws.flash('Выделите изображения для пакетной обработки');
        return;
      }
      const preset = serviceId === 'thumb' ? 'thumb' : 'web';
      const ops: Parameters<typeof applyImageOps>[1] =
        preset === 'thumb'
          ? [{ op: 'thumbnail', max: 512 }, { op: 'compress', maxKb: 120 }]
          : [{ op: 'resize', mode: 'max', max: 1920 }, { op: 'compress', maxKb: 400 }];
      const toBase64 = (bytes: Uint8Array): string => {
        let binary = '';
        const chunk = 0x8000;
        for (let i = 0; i < bytes.length; i += chunk) {
          binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
        }
        return btoa(binary);
      };
      ws.flash(`Обрабатываю ${images.length} изображений (${preset === 'thumb' ? 'миниатюры' : 'для веба'})…`);
      let done = 0;
      const failed: string[] = [];
      for (const image of images) {
        try {
          const bytes = new Uint8Array(await fileSystemService.readBinary(image.path));
          const mime =
            /\.png$/i.test(image.name)
              ? 'image/png'
              : /\.webp$/i.test(image.name)
                ? 'image/webp'
                : 'image/jpeg';
          const result = await applyImageOps({ base64: toBase64(bytes), mime }, ops);
          const ext = result.mime === 'image/png' ? '.png' : result.mime === 'image/webp' ? '.webp' : '.jpg';
          const outPath = image.path.replace(/\.[^./\\]+$/, '') + (preset === 'thumb' ? '-thumb' : '-web') + ext;
          await fileSystemService.writeFile(outPath, result.base64, 'base64');
          done += 1;
        } catch (err) {
          failed.push(`${image.name}: ${(err as Error).message}`);
        }
      }
      ws.flash(
        failed.length === 0
          ? `Готово: ${done} изображений`
          : `Обработано ${done} из ${images.length}; ошибки: ${failed.slice(0, 2).join('; ')}`
      );
      void ws.loadDirectory(images[0].path.replace(/[\\/][^\\/]+$/, ''));
    } else if (action === 'search-folder') {
      setSearchScope(item.isDirectory ? item.path : item.path.replace(/[\\/][^\\/]+$/, ''));
      setSearchPanelOpen(true);
    } else if (action === 'album' && serviceId) {
      const images = (selection || []).filter((it) => it.type === 'image' && !it.isDirectory);
      if (images.length < 2) {
        ws.flash('Выделите несколько фото: Ctrl — по одному, Shift — диапазон');
        return;
      }
      const title = 'WorkSpace ' + new Date().toLocaleString('ru-RU');
      if (serviceId === 'imgbb' || serviceId === 'freeimage' || serviceId === 'yapx') {
        const authorized = readPreferences().yapxAccountMode;
        const isSite = serviceId === 'yapx';
        ws.flash(
          isSite
            ? `Загружаю ${images.length} фото на Япикс…`
            : `Открываю ${serviceId === 'imgbb' ? 'ImgBB' : 'Freeimage.host'}…`
        );
        const result = await shareViaSite(
          serviceId,
          images.map((it) => ({ path: it.path, name: it.name })),
          { authorized }
        );
        const name =
          serviceId === 'imgbb' ? 'ImgBB' : serviceId === 'freeimage' ? 'Freeimage.host' : 'Япикс';
        if (!isSite) {
          ws.flash(result.success ? name + ' открыт, фото добавлены в форму' : 'Не удалось открыть ' + name);
        } else if (result.albums.length > 0) {
          await copyToClipboard(result.albums.join('\n'));
          const skipped = result.skipped.length ? `, пропущено ${result.skipped.length}` : '';
          ws.flash(
            `Япикс: ${result.count}/${images.length} фото в ${result.albums.length} альбом(ах)${skipped}, ссылки скопированы`
          );
        } else if (result.count > 0) {
          ws.flash(`Япикс: ${result.count}/${images.length} фото добавлены в форму загрузки`);
        } else {
          ws.flash(
            `Япикс не принял фото: ${result.error || result.skipped.map((item) => item.reason).join('; ')}`
          );
        }
      } else if (serviceId === 'telegraph') {
        ws.flash('Собираю страницу Telegraph: ' + images.length + ' фото…');
        const result = await uploadAlbum('telegraph', title, images.map((it) => ({ path: it.path, name: it.name })));
        if (result.url) await copyToClipboard(result.url);
        ws.flash(result.uploaded > 0
          ? 'Telegraph: ' + result.uploaded + '/' + images.length + (result.url ? ', ссылка скопирована' : '')
          : 'Telegraph не собрал страницу: ' + (result.error || result.failed.join('; ')));
      } else if (serviceId === 'catbox') {
        ws.flash('Загружаю ' + images.length + ' фото на Catbox…');
        const result = await uploadMany('catbox', images.map((it) => ({ path: it.path, name: it.name })));
        if (result.links.length) await copyToClipboard(result.links.join('\n'));
        ws.flash(result.uploaded > 0
          ? 'Catbox: ' + result.uploaded + '/' + images.length + ', ссылки скопированы'
          : 'Catbox не принял фото: ' + (result.error || result.failed.join('; ')));
      } else if (serviceId === 'imgur') {
        if (!getImgurClientId()) {
          setSettingsTab('integrations');
          setSettingsOpen(true);
          ws.flash('Укажите Imgur Client-ID на вкладке «Сервисы» в настройках');
          return;
        }
        ws.flash('Создаю альбом Imgur: ' + images.length + ' фото…');
        const result = await uploadAlbum('imgur', title, images.map((it) => ({ path: it.path, name: it.name })));
        if (result.url) await copyToClipboard(result.url);
        if (result.uploaded > 0) {
          ws.flash('Альбом Imgur (' + result.uploaded + '/' + images.length + ')' + (result.url ? ': ссылка скопирована' : ''));
        } else {
          ws.flash('Альбом не создан: ' + (result.error || result.failed.join('; ')));
        }
      } else if (serviceId === 'gdrive') {
        ws.flash('Загружаю ' + images.length + ' фото в Google Drive…');
        const result = await gdriveUpload(images.map((it) => it.path), title);
        if (result.link) await copyToClipboard(result.link);
        if (result.success || result.uploaded > 0) {
          ws.flash('Папка Google Drive: ' + result.uploaded + ' фото' + (result.link ? ', ссылка скопирована' : ''));
        } else {
          ws.flash('Загрузка не выполнена: ' + (result.error || result.failed.join('; ')));
        }
      }
    } else if (action === 'upload-drive') {
      const result = await gdriveUpload([item.path]);
      if (result.success) {
        if (result.link) await copyToClipboard(result.link);
        ws.flash('Загружено в Google Drive: ' + result.uploaded + ' файл(ов)');
      } else ws.flash('Загрузка в Google Drive не выполнена: ' + (result.error || result.failed.join('; ')));
    } else if (action === 'share' && serviceId) {
      if (serviceId === 'imgbb' || serviceId === 'freeimage' || serviceId === 'yapx') {
        const result = await shareViaSite(
          serviceId,
          [{ path: item.path, name: item.name }],
          { authorized: readPreferences().yapxAccountMode }
        );
        const name =
          serviceId === 'imgbb' ? 'ImgBB' : serviceId === 'freeimage' ? 'Freeimage.host' : 'Япикс';
        if (!result.success) {
          ws.flash(`Не удалось открыть ${name}: ${result.error || 'ошибка'}`);
          return;
        }
        if (result.albums.length > 0) {
          await copyToClipboard(result.albums.join('\n'));
          ws.flash(`${name}: альбом создан, ссылка скопирована`);
        } else {
          ws.flash(name + ' открыт, фото добавлено в форму');
        }
        return;
      }
      if (serviceId === 'imgur' && !getImgurClientId()) {
        setSettingsTab('integrations');
        setSettingsOpen(true);
        ws.flash('Укажите Imgur Client-ID на вкладке «Сервисы»');
        return;
      }
      const result = await uploadFilePath(serviceId, item.path);
      if (result.success && result.url) {
        await copyToClipboard(result.url);
        ws.flash('Ссылка скопирована: ' + result.url);
      } else ws.flash('Экспорт не выполнен: ' + (result.error || 'ошибка'));
    }
  }, [ws, sanitizeParentDir]);

  /* ============================================================
     Рендер редактора по типу файла
     ============================================================ */
  const renderEditor = (tab: OpenTab) => {
    const common = {
      tab,
      onChange: (content: string) => ws.updateTabContent(tab.id, content),
      onSave: () => ws.saveTab(tab.id),
      onAskAI: askAI,
      onRequestInlineEdit: openInlineCopilot,
      onViewModeChange: (viewMode: OpenTab['viewMode']) =>
        ws.patchTab(tab.id, { viewMode }),
    };

    let editor: React.ReactNode;
    switch (tab.type) {
      case 'markdown':
      case 'richtext':
        editor = <RichTextEditor {...common} />;
        break;
      case 'excel':
        editor = <ExcelEditor {...common} onAIApply={handleAIApply} />;
        break;
      case 'word':
        editor = <WordEditor {...common} />;
        break;
      case 'powerpoint':
        editor = <PowerPointEditor {...common} />;
        break;
      case 'pdf':
        editor = <PdfEditor {...common} onNotify={notify} />;
        break;
      case 'video':
        editor = (
          <VideoEditor
            tab={tab}
            onNotify={notify}
            onOpenResult={(filePath) => void ws.openFile(filePath)}
            onCloseTab={() => ws.closeTab(tab.id)}
          />
        );
        break;
      case 'binary':
        editor = <BinaryViewer tab={tab} onNotify={notify} />;
        break;
      case 'image':
        editor = (
          <ImageViewer
            tab={tab}
            onChange={(content) => ws.updateTabContent(tab.id, content)}
            onSave={() => void ws.saveTab(tab.id)}
            onNotify={notify}
          />
        );
        break;
      default:
        editor = (
          <MonacoEditorWrapper
            {...common}
            onRegisterContext={registerEditorContext}
            onRegisterApplier={(fn) => {
              editorApplier.current = fn;
            }}
          />
        );
    }

    // Предпросмотр: редактор рисует реальный документ, но взаимодействие
    // заблокировано. Двойной клик по файлу или кнопка «Редактировать»
    // снимает блокировку с этой же вкладки.
    const locked =
      Boolean(tab.isPreview) &&
      ['word', 'excel', 'powerpoint', 'markdown', 'richtext', 'text', 'code'].includes(
        tab.type
      );
    return (
      <div className="h-full" inert={locked || undefined}>
        <React.Suspense
          fallback={
            <div className="flex h-full items-center justify-center gap-2 text-xs text-ws-muted">
              <span className="h-3 w-3 animate-spin rounded-full border-2 border-ws-faint border-t-transparent" />
              Загрузка редактора…
            </div>
          }
        >
          {editor}
        </React.Suspense>
      </div>
    );
  };

  const dirtyPaths = useMemo(
    () => new Set(ws.dirtyTabs.map((t) => t.path)),
    [ws.dirtyTabs]
  );

  const hasContent = ws.tabs.length > 0;

  return (
    <div className="flex h-full w-full flex-col overflow-hidden bg-ws-bg text-ws">
      <TitleBar
        workspacePath={ws.rootPath}
        onOpenFolder={() => void ws.chooseWorkspaceFolder()}
        onNewFile={() =>
          ws.rootPath &&
          setNewFileDialog({ open: true, parentDir: ws.rootPath, kind: 'file', name: '' })
        }
        onSave={() => ws.activeTabId && void ws.saveTab(ws.activeTabId)}
        onSaveAs={() => openSaveAs()}
        onToggleSidebar={() => setSidebarOpen((v) => !v)}
        onToggleAIPanel={() => setAiPanelOpen((v) => !v)}
        onToggleSplit={() => ws.setSplitEnabled(!ws.splitEnabled)}
        onOpenSettings={() => setSettingsOpen(true)}
        onCheckUpdates={() => {
          setSettingsTab('system');
          setSettingsOpen(true);
        }}
        onQuickOpen={() => setQuickOpenOpen(true)}
        sidebarOpen={sidebarOpen}
        aiPanelOpen={aiPanelOpen}
        terminalOpen={terminalOpen}
        onToggleTerminal={() => setTerminalOpen((v) => !v)}
        splitEnabled={ws.splitEnabled}
        isDark={isDark}
        onToggleTheme={toggleTheme}
        isElectron={isElectron}
      />

      {updateInfo?.available && !updateDismissed && (
        <div className="flex items-center gap-2 border-b border-sky-500/40 bg-sky-500/10 px-3 py-1.5 text-[11px] text-ws">
          <Download size={13} className="shrink-0 text-sky-400" />
          <span className="min-w-0 flex-1 truncate">
            Доступна новая версия WorkSpace {updateInfo.latestVersion}
            {updateInfo.currentVersion ? ` (установлена ${updateInfo.currentVersion})` : ''}.
            Ключи и настройки сохранятся.
          </span>
          <button
            onClick={() => {
              setSettingsTab('system');
              setSettingsOpen(true);
              setUpdateDismissed(true);
            }}
            className="shrink-0 rounded bg-sky-600 px-2 py-0.5 text-[10px] font-medium text-white hover:bg-sky-500"
          >
            Обновить
          </button>
          <button
            onClick={() => setUpdateDismissed(true)}
            title="Скрыть"
            className="shrink-0 rounded p-0.5 text-ws-faint hover:bg-ws-hover hover:text-ws"
          >
            <X size={12} />
          </button>
        </div>
      )}

      {bootError && (
        <div className="flex items-center gap-2 border-b border-amber-500/40 bg-amber-500/10 px-3 py-2 text-[11px] text-ws-warning">
          <AlertTriangle size={13} className="shrink-0" />
          <span>{bootError}</span>
        </div>
      )}

      {/* Основная область */}
      <div className="flex min-h-0 flex-1">
        {/* Боковая панель: проводник или поиск */}
        {sidebarOpen && (
          <div className="flex w-72 shrink-0 flex-col border-r border-ws bg-ws-panel">
            <div className="flex items-center gap-1 border-b border-ws px-2 py-1.5">
              <button
                onClick={() => setSearchPanelOpen(false)}
                className={cn(
                  'flex items-center gap-1.5 rounded px-2 py-1 text-[11px] font-medium transition-colors',
                  !searchPanelOpen
                    ? 'bg-ws-accent-soft text-ws-accent'
                    : 'text-ws-muted hover:bg-ws-hover'
                )}
              >
                <Files size={12} />
                Файлы
              </button>
              <button
                onClick={() => {
                  setSearchPanelOpen(true);
                  setSidebarOpen(true);
                }}
                className={cn(
                  'flex items-center gap-1.5 rounded px-2 py-1 text-[11px] font-medium transition-colors',
                  searchPanelOpen
                    ? 'bg-ws-accent-soft text-ws-accent'
                    : 'text-ws-muted hover:bg-ws-hover'
                )}
              >
                <Sparkles size={12} />
                Поиск
              </button>
            </div>

            {searchPanelOpen ? (
              <GlobalSearchPanel
                workspacePath={searchScope || ws.rootPath}
                onOpenFile={(p) => void ws.openFile(p)}
                onClose={() => {
                  setSearchPanelOpen(false);
                  setSearchScope('');
                }}
              />
            ) : (
              <FileTree
                rootPath={ws.rootPath}
                tree={ws.tree}
                expandedDirs={ws.expandedDirs}
                loadingDirs={ws.loadingDirs}
                activePath={ws.activeTab?.path ?? ''}
                dirtyPaths={dirtyPaths}
                onRenamePath={handleExplorerRename}
                onMovePaths={handleExplorerMove}
                onToggleDir={(p) => void ws.toggleDirectory(p)}
                onOpenFile={(p) => void ws.openFile(p)}
                onPreviewFile={(p) => void ws.openFile(p, 'primary', false, true)}
                onRefresh={() => void ws.refreshTree()}
                onNewFile={(parentDir) =>
                  setNewFileDialog({
                    open: true,
                    parentDir: sanitizeParentDir(parentDir),
                    kind: 'file',
                    name: '',
                  })
                }
                onNewFolder={(parentDir) =>
                  setNewFileDialog({
                    open: true,
                    parentDir: sanitizeParentDir(parentDir),
                    kind: 'folder',
                    name: '',
                  })
                }
                onOpenFolder={() => void ws.chooseWorkspaceFolder()}
                favorites={ws.favorites}
                onOpenFavorite={(dirPath) => void ws.openWorkspace(dirPath)}
                onToggleFavorite={(dirPath) => void ws.toggleFavorite(dirPath)}
                onOpenGoogleDriveSettings={() => { setSettingsTab('integrations'); setSettingsOpen(true); }}
                onItemAction={(action, item, serviceId, selection) => void handleFileAction(action, item, serviceId, selection)}
                revealTarget={ws.revealTarget}
              />
            )}
          </div>
        )}

        {/* Рабочая область: вкладки + редакторы */}
        <div className="flex min-w-0 flex-1 flex-col bg-ws-editor">
          {/* Панель вкладок */}
          <div className="flex h-9 shrink-0 items-stretch overflow-x-auto border-b border-ws bg-ws-panel">
            {ws.tabs.map((tab) => {
              const isActive = tab.id === ws.activeTabId;
              return (
                <div
                  key={tab.id}
                  onClick={() => ws.setActiveTabId(tab.id)}
                  onAuxClick={(e) => {
                    if (e.button === 1) requestCloseTab(tab.id);
                  }}
                  title={tab.path}
                  className={cn(
                    'group flex min-w-[128px] max-w-[220px] cursor-pointer items-center gap-2 border-r border-ws px-3 text-xs transition-colors',
                    isActive
                      ? 'bg-ws-editor text-ws'
                      : 'text-ws-muted hover:bg-ws-hover hover:text-ws'
                  )}
                >
                  {isActive && <span className="absolute-none" />}
                  {tab.isPreview && <Eye size={11} className="shrink-0 text-sky-400" />}
                  <span className="truncate">{tab.name}</span>
                  {tab.isDirty && (
                    <span className="text-[14px] leading-none text-ws-warning">•</span>
                  )}
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      requestCloseTab(tab.id);
                    }}
                    className={cn(
                      'ml-auto rounded p-0.5 transition-opacity hover:bg-ws-hover',
                      isActive ? 'opacity-70' : 'opacity-0 group-hover:opacity-70'
                    )}
                    title="Закрыть вкладку (Ctrl+W)"
                  >
                    <X size={12} />
                  </button>
                </div>
              );
            })}

            <div className="flex flex-1 items-center justify-end gap-1 px-2">
              {ws.dirtyTabs.length > 0 && (
                <span className="mr-1 text-[10px] text-ws-warning">
                  {ws.dirtyTabs.length} не сохранено
                </span>
              )}
              <button
                onClick={() => setQuickOpenOpen(true)}
                title="Открыть файл (Ctrl+P)"
                className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"
              >
                <FilePlus2 size={13} />
              </button>
            </div>
          </div>

          {/* Предпросмотр: один клик по файлу — временная вкладка без правки */}
          {ws.activeTab?.isPreview && (
            <div
              data-preview-banner="1"
              className="flex h-8 shrink-0 items-center gap-2 border-b border-sky-500/30 bg-sky-500/10 px-3 text-[11px] text-ws-muted"
            >
              <Eye size={13} className="shrink-0 text-sky-400" />
              <span className="min-w-0 flex-1 truncate">
                Предпросмотр: <span className="text-ws">{ws.activeTab.name}</span> — только
                чтение
              </span>
              <button
                onClick={() => void ws.openFile(ws.activeTab!.path)}
                className="shrink-0 rounded border border-ws bg-ws-surface px-2 py-0.5 text-[10px] text-ws hover:bg-ws-hover"
              >
                Редактировать
              </button>
            </div>
          )}

          {/* Редакторы (обычный или разделённый режим) */}
          <div className="min-h-0 flex-1">
            {!hasContent ? (
              <EmptyWorkspace
                hasRoot={Boolean(ws.rootPath)}
                onOpenFolder={() => void ws.chooseWorkspaceFolder()}
                onOpenFile={() => setQuickOpenOpen(true)}
                onNewFile={() =>
                  ws.rootPath &&
                  setNewFileDialog({ open: true, parentDir: ws.rootPath, kind: 'file', name: '' })
                }
                isElectron={isElectron}
              />
            ) : ws.splitEnabled && ws.secondaryTab ? (
              <SplitView
                primary={ws.activeTab}
                secondary={ws.secondaryTab}
                ratio={ws.splitRatio}
                onRatioChange={ws.setSplitRatio}
                renderEditor={renderEditor}
                onFocusPane={(id) => ws.setActiveTabId(id)}
                secondaryId={ws.secondaryTabId}
              />
            ) : ws.activeTab ? (
              renderEditor(ws.activeTab)
            ) : (
              <div className="flex h-full items-center justify-center text-xs text-ws-faint">
                Выберите вкладку
              </div>
            )}
          </div>

          {/* Встроенный терминал */}
          {terminalOpen && (
            <TerminalPanel
              cwd={ws.rootPath}
              isDark={isDark}
              height={terminalHeight}
              onHeightChange={setTerminalHeight}
              onClose={() => setTerminalOpen(false)}
              command={terminalCommand}
              onCommandConsumed={() => setTerminalCommand(null)}
            />
          )}

          {/* Строка состояния */}
          <StatusBar
            message={ws.statusMessage}
            activeTab={ws.activeTab}
            model={ai.settings.model}
            isConfigured={ai.isConfigured}
            onOpenSettings={() => setSettingsOpen(true)}
          />
        </div>

        {/* Панель AI */}
        {aiPanelOpen && (
          <div className="w-96 shrink-0">
            <AIChatPanel
              settings={ai.settings}
              isConfigured={ai.isConfigured}
              availableModels={ai.availableModels}
              onUpdateModel={(model) => ai.updateSettings({ model })}
              activeTab={ws.activeTab}
              sessions={chat.sessions}
              activeSessionId={chat.activeId}
              onSelectSession={chat.setActiveId}
              onCreateSession={chat.createSession}
              onUpdateMessages={chat.updateSessionMessages}
              onDeleteSession={chat.deleteSession}
              onClearAll={chat.clearAll}
              onOpenSettings={() => setSettingsOpen(true)}
              onInsertIntoEditor={insertIntoEditor}
              onApplyToDocument={replaceDocument}
              pendingPrompt={pendingPrompt}
              onPendingPromptConsumed={() => setPendingPrompt(null)}
              onAnswerComplete={handleAIAnswer}
              onRunAction={handleAIRunAction}
              onDocumentTool={runDocumentTool}
              appContext={aiAppContext}
              onResolveAttachment={resolveChatAttachment}
              onReadDroppedFiles={readDroppedChatFiles}
              onResolveActiveDocument={resolveActiveDocument}
              docWriteAccess={aiDocWrite}
              onToggleDocWriteAccess={setAiDocWrite}
              extraWriteAccess={aiExtraWrite}
              onToggleExtraWriteAccess={setAiExtraWrite}
            />
          </div>
        )}
      </div>

      {/* ---- Оверлеи ---- */}
      {templatesOpen && (
        <div
          className="fixed inset-0 z-50 flex items-start justify-center bg-black/40 pt-24"
          onClick={() => setTemplatesOpen(false)}
        >
          <div
            data-template-palette="1"
            className="w-full max-w-xl overflow-hidden rounded-lg border border-ws bg-ws-surface p-2 shadow-xl"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="flex items-center justify-between px-2 py-1">
              <span className="text-xs font-medium text-ws">Шаблоны документов</span>
              <button
                onClick={() => setTemplatesOpen(false)}
                className="rounded px-1 text-ws-faint hover:bg-ws-hover hover:text-ws"
              >
                ×
              </button>
            </div>
            <div className="max-h-96 overflow-y-auto">
              {DOCUMENT_TEMPLATES.map((template) => (
                <button
                  key={template.id}
                  data-template-item={template.id}
                  onClick={() => void createFromTemplate(template)}
                  className="flex w-full flex-col gap-0.5 rounded px-2 py-1.5 text-left hover:bg-ws-hover"
                >
                  <span className="flex items-center gap-2 text-xs text-ws">
                    {template.name}
                    <span className="rounded border border-ws px-1 text-[9px] text-ws-faint">
                      {template.extension}
                    </span>
                  </span>
                  <span className="text-[10px] text-ws-faint">{template.description}</span>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {commandOpen && (
        <div
          className="fixed inset-0 z-50 flex items-start justify-center bg-black/40 pt-24"
          onClick={() => setCommandOpen(false)}
        >
          <div
            data-command-palette="1"
            className="w-full max-w-lg overflow-hidden rounded-lg border border-ws bg-ws-surface shadow-xl"
            onClick={(event) => event.stopPropagation()}
          >
            <input
              data-command-input="1"
              autoFocus
              value={commandQuery}
              onChange={(event) => setCommandQuery(event.target.value)}
              placeholder="Команда…"
              className="w-full border-b border-ws bg-ws-elevated px-3 py-2 text-sm text-ws placeholder:text-ws-faint focus:outline-none"
            />
            <div className="max-h-80 overflow-y-auto p-1">
              {[
                {
                  label: 'Сохранить активный файл',
                  hint: 'Ctrl+S',
                  run: () => {
                    const tab = ws.activeTab;
                    if (tab) void ws.saveTab(tab.id);
                  },
                },
                {
                  label: 'Сохранить все файлы',
                  hint: 'Ctrl+Alt+S',
                  run: () => void ws.saveAll(),
                },
                {
                  label: 'Сохранить как… (сменить формат)',
                  hint: 'Ctrl+Shift+S',
                  run: () => openSaveAs(),
                },
                {
                  label: 'Быстрый переход к файлу',
                  hint: 'Ctrl+P',
                  run: () => setQuickOpenOpen(true),
                },
                {
                  label: 'Поиск по workspace',
                  hint: 'Ctrl+Shift+F',
                  run: () => {
                    setSearchPanelOpen(true);
                    setSidebarOpen(true);
                  },
                },
                {
                  label: 'Найти и заменить в Word',
                  hint: '',
                  run: () => window.dispatchEvent(new CustomEvent('workspace:word-find')),
                },
                {
                  label: 'История правок AI',
                  hint: '',
                  run: () => window.dispatchEvent(new CustomEvent('workspace:word-history')),
                },
                {
                  label: 'Версии документа',
                  hint: '',
                  run: () => window.dispatchEvent(new CustomEvent('workspace:word-versions')),
                },
                {
                  label: 'Показать/скрыть AI-панель',
                  hint: '',
                  run: () => setAiPanelOpen((value) => !value),
                },
                {
                  label: 'Показать/скрыть терминал',
                  hint: '',
                  run: () => setTerminalOpen((value) => !value),
                },
                {
                  label: 'Создать из шаблона',
                  hint: '',
                  run: () => setTemplatesOpen(true),
                },
                {
                  label: 'Настройки',
                  hint: '',
                  run: () => setSettingsOpen(true),
                },
                {
                  label: 'Проверить обновление',
                  hint: '',
                  run: () => {
                    setSettingsTab('system');
                    setSettingsOpen(true);
                  },
                },
                {
                  label: 'Диагностика: скопировать отчёт',
                  hint: 'без секретов',
                  run: () => {
                    void copyDiagnosticsReport()
                      .then((size) => notify(`Отчёт диагностики скопирован (${size} символов)`, 'export'))
                      .catch((error: Error) => notify(`Отчёт не собран: ${error.message}`, 'error'));
                  },
                },
                {
                  label: 'Тема: светлая / тёмная',
                  hint: 'Ctrl+Shift+P',
                  run: () => toggleTheme(),
                },
                {
                  label: 'Проверить уведомление',
                  hint: 'тост и сигнал',
                  run: () => notify('Проверка уведомления', 'update'),
                },
                {
                  label: 'Новое окно',
                  hint: 'второй монитор',
                  run: () => {
                    void window.electronAPI
                      ?.windowNew?.()
                      .then((result) =>
                        result?.success ? undefined : notify('Окно не открыто: ' + (result?.error ?? ''), 'error')
                      );
                  },
                },
              ]
                .filter((item) =>
                  item.label.toLowerCase().includes(commandQuery.trim().toLowerCase())
                )
                .map((item) => (
                  <button
                    key={item.label}
                    data-command-item="1"
                    onClick={() => {
                      setCommandOpen(false);
                      setCommandQuery('');
                      item.run();
                    }}
                    className="flex w-full items-center justify-between rounded px-2 py-1.5 text-left text-xs text-ws hover:bg-ws-hover"
                  >
                    <span>{item.label}</span>
                    {item.hint && <span className="text-[10px] text-ws-faint">{item.hint}</span>}
                  </button>
                ))}
            </div>
          </div>
        </div>
      )}

      <QuickOpenModal
        isOpen={quickOpenOpen}
        onClose={() => setQuickOpenOpen(false)}
        workspacePath={ws.rootPath}
        onOpenFile={(p) => void ws.openFile(p)}
      />

      <SaveAsDialog
        open={saveAsOpen}
        tab={ws.tabs.find((tab) => tab.id === saveAsTabId) ?? null}
        busy={saveAsBusy}
        onClose={() => setSaveAsOpen(false)}
        onSave={handleSaveAs}
        onNotify={notify}
      />

      <SettingsModal
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        settings={ai.settings}
        onChangeProvider={(p) => void ai.changeProvider(p)}
        onUpdateSettings={ai.updateSettings}
        onSaveApiKey={ai.saveApiKey}
        availableModels={ai.availableModels}
        onFetchModels={ai.fetchModels}
        modelsCache={ai.modelsCache}
        initialTab={settingsTab}
        onNotify={notify}
        onBeforeUpdate={async () => {
          const { failed } = await ws.saveAll();
          if (failed.length > 0) {
            ws.flash(`Не удалось сохранить: ${failed.join(', ')} — обновление отменено`);
            return false;
          }
          return true;
        }}
      />

      <InlineCopilot
        target={inlineTarget}
        settings={ai.settings}
        isConfigured={ai.isConfigured}
        onClose={() => setInlineTarget(null)}
        onApply={applyInlineEdit}
        onOpenSettings={() => setSettingsOpen(true)}
      />

      {/* Диалог создания файла/папки */}
      <Modal
        open={newFileDialog.open}
        onClose={() => setNewFileDialog((d) => ({ ...d, open: false }))}
        title={newFileDialog.kind === 'file' ? 'Новый файл' : 'Новая папка'}
        icon={
          newFileDialog.kind === 'file' ? (
            <FilePlus2 size={15} className="text-ws-accent" />
          ) : (
            <FolderPlus size={15} className="text-ws-accent" />
          )
        }
        footer={
          <>
            <Button
              variant="ghost"
              onClick={() => setNewFileDialog((d) => ({ ...d, open: false }))}
            >
              Отмена
            </Button>
            <Button
              variant="primary"
              disabled={!newFileDialog.name.trim()}
              onClick={async () => {
                const { parentDir, kind, name } = newFileDialog;
                const created =
                  kind === 'file'
                    ? await ws.createFile(parentDir, name.trim(), '')
                    : await ws.createFolder(parentDir, name.trim());
                // Диалог закрываем только при успехе: иначе пользователь видит
                // тост с причиной и может исправить имя.
                if (created) setNewFileDialog((d) => ({ ...d, open: false }));
              }}
            >
              Создать
            </Button>
          </>
        }
      >
        <div className="space-y-3">
          <p className="text-[11px] text-ws-muted">
            Папка: <span className="font-mono text-ws">{newFileDialog.parentDir}</span>
          </p>
          <Input
            autoFocus
            value={newFileDialog.name}
            onChange={(e) => setNewFileDialog((d) => ({ ...d, name: e.target.value }))}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && newFileDialog.name.trim()) {
                const { parentDir, kind, name } = newFileDialog;
                void (async () => {
                  const created =
                    kind === 'file'
                      ? await ws.createFile(parentDir, name.trim(), '')
                      : await ws.createFolder(parentDir, name.trim());
                  if (created) setNewFileDialog((d) => ({ ...d, open: false }));
                })();
              }
            }}
            placeholder={
              newFileDialog.kind === 'file' ? 'например, Заметки.md' : 'например, Документы'
            }
          />
          {newFileDialog.kind === 'file' && (
            <div className="flex flex-wrap gap-1">
              {['.md', '.txt', '.json', '.csv', '.xlsx', '.docx', '.pptx', '.py', '.sql'].map(
                (ext) => (
                  <button
                    key={ext}
                    onClick={() =>
                      setNewFileDialog((d) => {
                        const stem = d.name.replace(/\.[^.]+$/, '') || 'Новый файл';
                        return { ...d, name: `${stem}${ext}` };
                      })
                    }
                    className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 font-mono text-[10px] text-ws-muted hover:border-ws-accent hover:text-ws-accent"
                  >
                    {ext}
                  </button>
                )
              )}
            </div>
          )}
        </div>
      </Modal>

      {/* Переименование файла или папки */}
      <Modal
        open={renameDialog.open}
        onClose={() => setRenameDialog((d) => ({ ...d, open: false }))}
        title="Переименовать"
        icon={<Pencil size={15} className="text-ws-accent" />}
        width="max-w-md"
        footer={
          <>
            <Button
              variant="ghost"
              onClick={() => setRenameDialog((d) => ({ ...d, open: false }))}
            >
              Отмена
            </Button>
            <Button
              variant="primary"
              disabled={!renameDialog.name.trim() || renameDialog.busy}
              onClick={async () => {
                const next = renameDialog.name.trim();
                if (!next) return;
                setRenameDialog((d) => ({ ...d, busy: true }));
                const ok = await ws.renamePath(renameDialog.path, next);
                setRenameDialog((d) => ({ ...d, busy: false, open: !ok }));
              }}
            >
              Переименовать
            </Button>
          </>
        }
      >
        <div className="space-y-3">
          <p className="text-[11px] text-ws-muted">
            Текущее имя: <span className="font-mono text-ws">{baseName(renameDialog.path)}</span>
          </p>
          <Input
            autoFocus
            value={renameDialog.name}
            onChange={(e) => setRenameDialog((d) => ({ ...d, name: e.target.value }))}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && renameDialog.name.trim() && !renameDialog.busy) {
                e.preventDefault();
                const next = renameDialog.name.trim();
                void (async () => {
                  setRenameDialog((d) => ({ ...d, busy: true }));
                  const ok = await ws.renamePath(renameDialog.path, next);
                  setRenameDialog((d) => ({ ...d, busy: false, open: !ok }));
                })();
              }
            }}
            placeholder="Новое имя"
          />
        </div>
      </Modal>

      <Modal
        open={unsavedDialog.open}
        onClose={() => setUnsavedDialog({ open: false, tabId: '' })}
        title="Несохранённые изменения"
        icon={<AlertTriangle size={15} className="text-ws-warning" />}
        width="max-w-md"
        footer={
          <>
            <Button variant="ghost" onClick={() => setUnsavedDialog({ open: false, tabId: '' })}>
              Отмена
            </Button>
            <Button
              variant="danger"
              onClick={() => {
                ws.closeTab(unsavedDialog.tabId);
                setUnsavedDialog({ open: false, tabId: '' });
              }}
            >
              Не сохранять
            </Button>
            <Button
              variant="primary"
              onClick={async () => {
                const id = unsavedDialog.tabId;
                setUnsavedDialog({ open: false, tabId: '' });
                const ok = await ws.saveTab(id);
                if (ok) ws.closeTab(id);
              }}
            >
              Сохранить и закрыть
            </Button>
          </>
        }
      >
        <p className="text-xs leading-relaxed text-ws-muted">
          Файл{' '}
          <strong className="text-ws">
            {ws.tabs.find((t) => t.id === unsavedDialog.tabId)?.name}
          </strong>{' '}
          содержит несохранённые изменения. Сохранить их перед закрытием?
        </p>
      </Modal>

      {/* Диалог выхода из приложения с несохранёнными файлами */}
      <Modal
        open={exitDialogOpen}
        onClose={() => {
          // Закрытие диалога = отмена выхода, приложение продолжает работу
          setExitDialogOpen(false);
          window.electronAPI?.confirmClose(false);
        }}
        title="Несохранённые изменения"
        icon={<AlertTriangle size={15} className="text-ws-warning" />}
        width="max-w-lg"
        footer={
          <>
            <Button
              variant="ghost"
              onClick={() => {
                setExitDialogOpen(false);
                window.electronAPI?.confirmClose(false);
              }}
              disabled={exitBusy}
            >
              Отмена
            </Button>
            <Button variant="danger" onClick={doExit} disabled={exitBusy}>
              Выйти без сохранения
            </Button>
            <Button variant="primary" onClick={() => void saveAllAndExit()} disabled={exitBusy}>
              {exitBusy ? <Spinner size={12} /> : <Save size={12} />}
              Сохранить всё и выйти
            </Button>
          </>
        }
      >
        <div className="space-y-3">
          <p className="text-xs leading-relaxed text-ws-muted">
            Следующие файлы содержат несохранённые изменения. Если выйти сейчас, изменения будут
            потеряны.
          </p>

          <div className="max-h-56 space-y-1 overflow-y-auto rounded-lg border border-ws bg-ws-elevated p-2">
            {ws.dirtyTabs.map((t) => (
              <div
                key={t.id}
                className="flex items-center gap-2 rounded px-2 py-1 text-[11px] hover:bg-ws-hover"
              >
                <span className="h-1.5 w-1.5 shrink-0 rounded-full bg-ws-warning" />
                <span className="truncate font-medium text-ws">{t.name}</span>
                <span className="ml-auto shrink-0 truncate font-mono text-[10px] text-ws-faint">
                  {t.path}
                </span>
              </div>
            ))}
          </div>

          <p className="text-[11px] text-ws-faint">
            «Сохранить всё и выйти» запишет {ws.dirtyTabs.length}{' '}
            {ws.dirtyTabs.length === 1 ? 'файл' : 'файлов'} на диск перед закрытием.
          </p>
        </div>
      </Modal>

      {/* Уведомления */}
      {ws.statusMessage && <Toast message={ws.statusMessage} />}
    </div>
  );
};

/* ============================================================
   Разделённый экран (Этап 6.1)
   ============================================================ */

const SplitView: React.FC<{
  primary: OpenTab | null;
  secondary: OpenTab;
  ratio: number;
  onRatioChange: (r: number) => void;
  renderEditor: (tab: OpenTab) => React.ReactNode;
  onFocusPane: (id: string) => void;
  secondaryId: string;
}> = ({ primary, secondary, ratio, onRatioChange, renderEditor, onFocusPane, secondaryId }) => {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const dragging = React.useRef(false);

  const handleMouseDown = () => {
    dragging.current = true;
    const onMove = (e: MouseEvent) => {
      if (!dragging.current || !containerRef.current) return;
      const rect = containerRef.current.getBoundingClientRect();
      const pct = ((e.clientX - rect.left) / rect.width) * 100;
      onRatioChange(Math.min(85, Math.max(15, pct)));
    };
    const onUp = () => {
      dragging.current = false;
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.body.style.cursor = '';
    };
    document.body.style.cursor = 'col-resize';
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  };

  return (
    <div ref={containerRef} className="flex h-full w-full">
      <div style={{ width: `${ratio}%` }} className="h-full min-w-0 overflow-hidden">
        {primary ? (
          <div onMouseDown={() => onFocusPane(primary.id)} className="h-full">
            {renderEditor(primary)}
          </div>
        ) : (
          <div className="flex h-full items-center justify-center text-xs text-ws-faint">
            Первая область пуста
          </div>
        )}
      </div>

      <div
        onMouseDown={handleMouseDown}
        className="w-1 shrink-0 cursor-col-resize bg-ws-panel transition-colors hover:bg-ws-accent"
        title="Потяните, чтобы изменить ширину областей"
      />

      <div style={{ width: `${100 - ratio}%` }} className="h-full min-w-0 overflow-hidden">
        <div onMouseDown={() => onFocusPane(secondaryId)} className="h-full">
          {renderEditor(secondary)}
        </div>
      </div>
    </div>
  );
};

/* ============================================================
   Пустое состояние рабочей области
   ============================================================ */

const EmptyWorkspace: React.FC<{
  hasRoot: boolean;
  onOpenFolder: () => void;
  onOpenFile: () => void;
  onNewFile: () => void;
  isElectron: boolean;
}> = ({ hasRoot, onOpenFolder, onOpenFile, onNewFile, isElectron }) => {
  const shortcuts = [
    { keys: 'Ctrl+P', label: 'Быстрый переход к файлу' },
    { keys: 'Ctrl+N', label: 'Новый файл' },
    { keys: 'Ctrl+S', label: 'Сохранить' },
    { keys: 'Ctrl+Shift+S', label: 'Сохранить как (сменить формат)' },
    { keys: 'Ctrl+Alt+S', label: 'Сохранить всё' },
    { keys: 'Ctrl+K', label: 'AI-правка выделенного текста' },
    { keys: 'Ctrl+B', label: 'Скрыть/показать проводник' },
    { keys: 'Ctrl+Shift+F', label: 'Поиск по всем файлам' },
  ];

  return (
    <div className="flex h-full flex-col items-center justify-center gap-5 p-8">
      <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-ws-accent-soft">
        <Sparkles size={26} className="text-ws-accent" />
      </div>

      <div className="text-center">
        <h2 className="text-lg font-semibold text-ws">WorkSpace</h2>
        <p className="mt-1 max-w-md text-xs leading-relaxed text-ws-muted">
          Единое рабочее пространство для документов, таблиц, презентаций, кода и заметок —
          с персональным AI-ассистентом.
        </p>
      </div>

      <div className="flex items-center gap-2">
        <Button variant="primary" onClick={onOpenFolder}>
          Открыть папку
        </Button>
        <Button variant="secondary" onClick={onNewFile} disabled={!hasRoot}>
          Новый файл
        </Button>
        <Button variant="secondary" onClick={onOpenFile} disabled={!hasRoot}>
          Найти файл
        </Button>
      </div>

      <div className="mt-2 w-full max-w-sm space-y-1.5 rounded-lg border border-ws bg-ws-surface p-3">
        <p className="mb-2 text-[10px] font-semibold tracking-wider text-ws-muted uppercase">
          Горячие клавиши
        </p>
        {shortcuts.map((s) => (
          <div key={s.keys} className="flex items-center justify-between text-[11px]">
            <span className="text-ws-muted">{s.label}</span>
            <kbd className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 font-mono text-[10px] text-ws-faint">
              {s.keys}
            </kbd>
          </div>
        ))}
      </div>

      {!isElectron && (
        <p className="max-w-md text-center text-[11px] text-ws-warning">
          Приложение открыто вне Electron — файловые операции недоступны. Запустите{' '}
          <code className="font-mono">npm run dev</code> для полноценной работы.
        </p>
      )}
    </div>
  );
};

/* ============================================================
   Строка состояния
   ============================================================ */

const StatusBar: React.FC<{
  message: string;
  activeTab: OpenTab | null;
  model: string;
  isConfigured: boolean;
  onOpenSettings: () => void;
}> = ({ message, activeTab, model, isConfigured, onOpenSettings }) => (
  <div className="flex h-6 shrink-0 items-center justify-between border-t border-ws bg-ws-panel px-3 text-[10px] text-ws-faint">
    <div className="flex items-center gap-3">
      {message ? (
        <span className="text-ws-accent">{message}</span>
      ) : activeTab ? (
        <>
          <span className="truncate">{activeTab.path}</span>
          <span className="uppercase">{activeTab.encoding}</span>
        </>
      ) : (
        <span>Готово</span>
      )}
    </div>

    <button
      onClick={onOpenSettings}
      className="flex items-center gap-1.5 rounded px-1.5 py-0.5 transition-colors hover:bg-ws-hover"
      title="Настройки AI"
    >
      <span
        className={cn(
          'h-1.5 w-1.5 rounded-full',
          isConfigured ? 'bg-emerald-500' : 'bg-amber-500'
        )}
      />
      <span>{isConfigured ? model : 'AI не настроен'}</span>
    </button>
  </div>
);

export default App;
