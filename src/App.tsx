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
} from 'lucide-react';

import { OpenTab, ChatMessage, AIProvider, FileItem, ChatAttachment } from './types';
import { fileSystemService, getFileType } from './services/fileSystemService';
import { useWorkspace } from './state/workspace';
import { useAISettings, useChatSessions } from './state/settings';
import { useTheme } from './hooks/useTheme';
import { useLocalStorage } from './hooks/useLocalStorage';
import { cn, baseName, extName, joinPath, uid } from './lib/utils';

import { TitleBar } from './components/workspace/TitleBar';
import { FileTree } from './components/workspace/FileTree';
import { QuickOpenModal } from './components/workspace/QuickOpenModal';
import { GlobalSearchPanel } from './components/workspace/GlobalSearchPanel';
import { SettingsModal } from './components/settings/SettingsModal';
import { AIChatPanel } from './components/ai/AIChatPanel';
import { InlineCopilot, InlineEditTarget } from './components/ai/InlineCopilot';
import { Button, Input, Modal, Spinner, Toast } from './components/ui';

import { MonacoEditorWrapper } from './components/editors/MonacoEditorWrapper';
import { RichTextEditor } from './components/editors/RichTextEditor';
import { ExcelEditor, AIAction } from './components/editors/ExcelEditor';
import { WordEditor } from './components/editors/WordEditor';
import { PowerPointViewer } from './components/editors/PowerPointViewer';
import { ImageViewer } from './components/editors/ImageViewer';
import { BinaryViewer } from './components/editors/BinaryViewer';
import { VideoEditor } from './components/editors/VideoEditor';
import { TerminalPanel } from './components/terminal/TerminalPanel';
import { copyToClipboard, getImgurClientId, uploadFilePath, uploadAlbum, uploadMany, gdriveUpload } from './services/shareService';
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

const App: React.FC = () => {
  const ws = useWorkspace();
  const ai = useAISettings();
  const chat = useChatSessions();
  const { isDark, toggleTheme } = useTheme();

  const [sidebarOpen, setSidebarOpen] = useLocalStorage('ws_sidebar_open', true);
  const [aiPanelOpen, setAiPanelOpen] = useLocalStorage('ws_ai_panel_open', true);
  const [terminalOpen, setTerminalOpen] = useLocalStorage('ws_terminal_open', false);
  const [terminalHeight, setTerminalHeight] = useLocalStorage('ws_terminal_height', 240);
  const [quickOpenOpen, setQuickOpenOpen] = useState(false);
  const [searchPanelOpen, setSearchPanelOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [settingsTab, setSettingsTab] = useState<SettingsTab>('provider');
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

  /* ============================================================
     Горячие клавиши (Этап 6.3)
     ============================================================ */
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;

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
      if (e.key === 's' && !e.shiftKey) {
        e.preventDefault();
        if (ws.activeTabId) void ws.saveTab(ws.activeTabId);
        return;
      }

      // Ctrl+Shift+S — сохранить всё
      if (e.key === 'S' && e.shiftKey) {
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
        } else {
          const parentDir = data.path.slice(
            0,
            Math.max(data.path.lastIndexOf('\\'), data.path.lastIndexOf('/'))
          );
          if (!ws.rootPath) {
            await ws.openWorkspace(parentDir);
          }
          await ws.openFile(data.path);
          ws.flash(`Открыт файл: ${baseName(data.path)}`);
        }
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

  const insertIntoEditor = useCallback(
    (text: string) => {
      const tab = ws.activeTab;
      if (!tab) {
        ws.flash('Нет активного документа для вставки');
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

      if (tab.type === 'word') {
        window.dispatchEvent(
          new CustomEvent('workspace:insert-text', {
            detail: { tabId: tab.id, text },
          })
        );
        ws.flash('Текст вставлен в документ Word');
        return;
      }

      if (tab.type === 'markdown' || tab.type === 'richtext') {
        window.dispatchEvent(
          new CustomEvent('workspace:insert-text', {
            detail: { tabId: tab.id, text },
          })
        );
        const next = tab.content ? `${tab.content}\n\n${text}` : text;
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

      if (tab.type === 'excel') {
        window.dispatchEvent(
          new CustomEvent('workspace:excel-replace-document', {
            detail: { tabId: tab.id, text },
          })
        );
        ws.flash('Таблица заменена данными ответа');
        return;
      }

      if (tab.type === 'word') {
        try {
          const docTitle = tab.name.replace(/\.docx?$/i, '');
          const base64 = await generateDocxBase64(text, docTitle);
          ws.updateTabContent(tab.id, base64);
          window.dispatchEvent(
            new CustomEvent('workspace:replace-document', {
              detail: { tabId: tab.id, text, base64 },
            })
          );
          ws.flash('Документ Word заменён ответом AI');
        } catch (err) {
          ws.flash(`Ошибка замены документа: ${(err as Error).message}`);
        }
        return;
      }

      if (tab.type === 'markdown' || tab.type === 'richtext') {
        ws.updateTabContent(tab.id, text);
        window.dispatchEvent(
          new CustomEvent('workspace:replace-document', {
            detail: { tabId: tab.id, text },
          })
        );
        ws.flash('Документ заменён ответом AI');
        return;
      }

      ws.updateTabContent(tab.id, text);
      editorApplier.current?.({ from: 0, to: tab.content.length, text });
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
              workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet([columns, ...rows]);
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
            const created = await ws.createFile(dir, name, content, isCsv ? 'utf-8' : 'base64');
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
              await fileSystemService.writeFile(targetPath, csv, 'utf-8');
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
            const created = await ws.createFile(dir, name, action.content || '', 'utf-8');
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
        if (
          ext === 'csv' ||
          ext === 'tsv' ||
          ext === 'md' ||
          ext === 'markdown' ||
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
    return lines.join('\n');
  }, [
    ws.rootPath,
    ws.activeTab,
    ws.tabs,
    sidebarOpen,
    aiPanelOpen,
    terminalOpen,
    ai.settings.model,
    isDark,
  ]);

  const handleFileAction = useCallback(async (action: string, item: FileItem, serviceId?: string, selection?: FileItem[]) => {
    if (action === 'open') {
      if (item.isDirectory) void ws.toggleDirectory(item.path);
      else void ws.openFile(item.path);
    } else if (action === 'new-file') {
      setNewFileDialog({ open: true, parentDir: item.path, kind: 'file', name: '' });
    } else if (action === 'new-folder') {
      setNewFileDialog({ open: true, parentDir: item.path, kind: 'folder', name: '' });
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
      const next = window.prompt('Новое имя:', item.name);
      if (next && next.trim() && next.trim() !== item.name) await ws.renamePath(item.path, next.trim());
    } else if (action === 'delete') {
      if (window.confirm('Удалить «' + item.name + '»' + (item.isDirectory ? ' и всё его содержимое' : '') + '?')) await ws.deletePath(item.path);
    } else if (action === 'copy-path') {
      const ok = await copyToClipboard(item.path);
      if (ok) ws.flash('Путь скопирован'); else ws.flash('Не удалось скопировать путь');
    } else if (action === 'reveal') {
      await fileSystemService.showInExplorer(item.path);
    } else if (action === 'album' && serviceId) {
      const images = (selection || []).filter((it) => it.type === 'image' && !it.isDirectory);
      if (images.length < 2) {
        ws.flash('Выделите несколько фото: Ctrl — по одному, Shift — диапазон');
        return;
      }
      const title = 'WorkSpace ' + new Date().toLocaleString('ru-RU');
      if (serviceId === 'imgbb' || serviceId === 'freeimage') {
        const open = serviceId === 'imgbb' ? window.electronAPI?.openImgbb : window.electronAPI?.openFreeimage;
        const result = await open?.(images.map((it) => it.path));
        const name = serviceId === 'imgbb' ? 'ImgBB' : 'Freeimage.host';
        ws.flash(result?.success ? name + ' открыт, фото добавлены в форму' : 'Не удалось открыть ' + name);
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
      if (serviceId === 'imgbb' || serviceId === 'freeimage') {
        const open = serviceId === 'imgbb' ? window.electronAPI?.openImgbb : window.electronAPI?.openFreeimage;
        const result = await open?.([item.path]);
        const name = serviceId === 'imgbb' ? 'ImgBB' : 'Freeimage.host';
        ws.flash(result?.success ? name + ' открыт, фото добавлено в форму' : 'Не удалось открыть ' + name);
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
  }, [ws]);

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

    switch (tab.type) {
      case 'markdown':
      case 'richtext':
        return <RichTextEditor {...common} />;
      case 'excel':
        return <ExcelEditor {...common} onAIApply={handleAIApply} />;
      case 'word':
        return <WordEditor {...common} />;
      case 'powerpoint':
        return <PowerPointViewer tab={tab} onAskAI={askAI} />;
      case 'pdf':
        return (
          <embed
            src={'media://local/?path=' + encodeURIComponent(tab.path)}
            type="application/pdf"
            className="h-full w-full"
          />
        );
      case 'video':
        return (
          <VideoEditor
            tab={tab}
            onNotify={ws.flash}
            onOpenResult={(filePath) => void ws.openFile(filePath)}
          />
        );
      case 'binary':
        return <BinaryViewer tab={tab} onNotify={ws.flash} />;
      case 'image':
        return (
          <ImageViewer
            tab={tab}
            onChange={(content) => ws.updateTabContent(tab.id, content)}
            onSave={() => void ws.saveTab(tab.id)}
            onNotify={ws.flash}
          />
        );
      default:
        return (
          <MonacoEditorWrapper
            {...common}
            onRegisterContext={registerEditorContext}
            onRegisterApplier={(fn) => {
              editorApplier.current = fn;
            }}
          />
        );
    }
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
        onToggleSidebar={() => setSidebarOpen((v) => !v)}
        onToggleAIPanel={() => setAiPanelOpen((v) => !v)}
        onToggleSplit={() => ws.setSplitEnabled(!ws.splitEnabled)}
        onOpenSettings={() => setSettingsOpen(true)}
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
                workspacePath={ws.rootPath}
                onOpenFile={(p) => void ws.openFile(p)}
                onClose={() => setSearchPanelOpen(false)}
              />
            ) : (
              <FileTree
                rootPath={ws.rootPath}
                tree={ws.tree}
                expandedDirs={ws.expandedDirs}
                loadingDirs={ws.loadingDirs}
                activePath={ws.activeTab?.path ?? ''}
                dirtyPaths={dirtyPaths}
                onToggleDir={(p) => void ws.toggleDirectory(p)}
                onOpenFile={(p) => void ws.openFile(p)}
                onRefresh={() => void ws.refreshTree()}
                onNewFile={(parentDir) =>
                  setNewFileDialog({ open: true, parentDir, kind: 'file', name: '' })
                }
                onNewFolder={(parentDir) =>
                  setNewFileDialog({ open: true, parentDir, kind: 'folder', name: '' })
                }
                onOpenFolder={() => void ws.chooseWorkspaceFolder()}
                favorites={ws.favorites}
                onOpenFavorite={(dirPath) => void ws.openWorkspace(dirPath)}
                onToggleFavorite={(dirPath) => void ws.toggleFavorite(dirPath)}
                onOpenGoogleDriveSettings={() => { setSettingsTab('integrations'); setSettingsOpen(true); }}
                onItemAction={(action, item, serviceId, selection) => void handleFileAction(action, item, serviceId, selection)}
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
              appContext={aiAppContext}
              onResolveAttachment={resolveChatAttachment}
              onReadDroppedFiles={readDroppedChatFiles}
            />
          </div>
        )}
      </div>

      {/* ---- Оверлеи ---- */}
      <QuickOpenModal
        isOpen={quickOpenOpen}
        onClose={() => setQuickOpenOpen(false)}
        workspacePath={ws.rootPath}
        onOpenFile={(p) => void ws.openFile(p)}
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
        onNotify={ws.flash}
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
                setNewFileDialog((d) => ({ ...d, open: false }));
                if (kind === 'file') {
                  await ws.createFile(parentDir, name.trim(), '');
                } else {
                  await ws.createFolder(parentDir, name.trim());
                }
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
                setNewFileDialog((d) => ({ ...d, open: false }));
                if (kind === 'file') void ws.createFile(parentDir, name.trim(), '');
                else void ws.createFolder(parentDir, name.trim());
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

      {/* Диалог несохранённых изменений */}
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
