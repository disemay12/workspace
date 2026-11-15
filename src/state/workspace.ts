import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { FileItem, OpenTab, FileType } from '../types';
import { fileSystemService, getFileType } from '../services/fileSystemService';
import { extName, baseName, dirName, joinPath, uid } from '../lib/utils';
import { generateDocxBase64, generateExcelBase64 } from '../services/aiActions';
import { htmlToDocxBase64 } from '../lib/docx';
import { BLANK_PPTX_BASE64 } from '../lib/pptxTemplate';
import { docxPathFor, extractLegacyDocText, plainTextToHtml } from '../lib/legacyDoc';
import { dropSnapshotsForPath, noteClosedPath, noteOpenedPath } from '../services/autosave';
import {
  type ConversionFormat,
  convertDocument,
  readFileSource,
} from '../services/documentConversion';

/** Типы файлов, содержимое которых читается как base64 (бинарные офисные форматы). */
const BINARY_TYPES: FileType[] = ['word', 'powerpoint', 'image'];
/** Excel читается как base64 (xlsx/xls), CSV/TSV — как текст. */
const isBinaryExcel = (ext: string) => ext === 'xlsx' || ext === 'xls';

/** Путь в одном виде: регистр и разделители в Windows различаются между IPC и UI. */
function normalizePath(value: string): string {
  return (value || '').replace(/\//g, '\\').toLowerCase().replace(/\\+$/, '');
}

function samePath(a: string, b: string): boolean {
  return normalizePath(a) === normalizePath(b);
}

/**
 * Очередь сохранений на вкладку. Два перекрывающихся сохранения завершались в
 * произвольном порядке, и старое содержимое могло лечь на диск поверх нового.
 */
const saveQueues = new Map<string, Promise<unknown>>();

export interface TreeState {
  [dirPath: string]: FileItem[];
}

export function useWorkspace() {
  const [rootPath, setRootPath] = useState<string>('');
  const [tree, setTree] = useState<TreeState>({});
  const [expandedDirs, setExpandedDirs] = useState<Set<string>>(new Set());
  const [loadingDirs, setLoadingDirs] = useState<Set<string>>(new Set());
  const [tabs, setTabs] = useState<OpenTab[]>([]);
  const [activeTabId, setActiveTabId] = useState<string>('');
  const [secondaryTabId, setSecondaryTabId] = useState<string>('');
  const [splitEnabled, setSplitEnabled] = useState<boolean>(false);
  const [splitRatio, setSplitRatio] = useState<number>(50);
  const [statusMessage, setStatusMessage] = useState<string>('');
  const [favorites, setFavorites] = useState<{ path: string; name: string }[]>([]);

  // Ref для доступа к актуальным вкладкам внутри асинхронных колбэков
  const tabsRef = useRef<OpenTab[]>([]);
  // Идентификаторы панелей для тех же колбэков: состояние к моменту вызова отстаёт.
  const activeTabIdRef = useRef('');
  const secondaryTabIdRef = useRef('');
  // Синхронизация после коммита: ref не должен указывать на отброшенный рендер.
  useLayoutEffect(() => {
    tabsRef.current = tabs;
    activeTabIdRef.current = activeTabId;
    secondaryTabIdRef.current = secondaryTabId;
  }, [tabs, activeTabId, secondaryTabId]);

  /**
   * Вкладки, которые пользователь действительно правил руками.
   *
   * Редакторы вызывают `onChange` и программно — Word/Excel ресериализуют
   * документ сразу после открытия, — поэтому снятие предпросмотра привязано к
   * вводу, а не к факту изменения содержимого: иначе вкладка, открытая одиночным
   * кликом, становилась постоянной без участия пользователя.
   */
  const userEditedTabsRef = useRef<Set<string>>(new Set());
  useEffect(() => {
    if (typeof document === 'undefined') return;
    const markEdited = (event: Event) => {
      if (!event.isTrusted) return;
      const target = event.target as HTMLElement | null;
      if (!target || typeof target.closest !== 'function') return;
      if (!target.closest('[contenteditable="true"], .monaco-editor')) return;
      const tabId = activeTabIdRef.current;
      if (tabId) userEditedTabsRef.current.add(tabId);
    };
    const events = ['input', 'paste', 'cut', 'drop', 'keydown'];
    for (const name of events) document.addEventListener(name, markEdited, true);
    return () => {
      for (const name of events) document.removeEventListener(name, markEdited, true);
    };
  }, []);

  // Активная вкладка обязана существовать: два закрытия в одном тике иначе
  // оставляли activeTabId на уже закрытой вкладке.
  useEffect(() => {
    if (!activeTabId) return;
    if (tabs.some((tab) => tab.id === activeTabId)) return;
    setActiveTabId(tabs[0]?.id ?? '');
  }, [tabs, activeTabId]);
  /** Открытия файлов в процессе: защита от двойного клика и гонок. */
  const pendingOpenRef = useRef<Map<string, Promise<string | null>>>(new Map());
  const previewOpenSeqRef = useRef(0);

  const flashTimerRef = useRef<number | null>(null);

  const flash = useCallback((message: string) => {
    setStatusMessage(message);
    // Один таймер: иначе на каждое сообщение остаётся висящий setTimeout.
    if (flashTimerRef.current !== null) window.clearTimeout(flashTimerRef.current);
    flashTimerRef.current = window.setTimeout(() => {
      flashTimerRef.current = null;
      setStatusMessage((prev) => (prev === message ? '' : prev));
    }, 3200);
  }, []);

  useEffect(() => {
    return () => {
      if (flashTimerRef.current !== null) window.clearTimeout(flashTimerRef.current);
    };
  }, []);

  // Вторичная панель не может показывать ту же вкладку, что основная:
  // редактор получил бы два независимых состояния одного файла.
  useEffect(() => {
    if (secondaryTabId && secondaryTabId === activeTabId) setSecondaryTabId('');
  }, [activeTabId, secondaryTabId]);

  /* ---------------- Файловое дерево ---------------- */

  const loadDirectory = useCallback(async (dirPath: string): Promise<FileItem[]> => {
    setLoadingDirs((prev) => new Set(prev).add(dirPath));
    try {
      const items = await fileSystemService.listDirectory(dirPath);
      setTree((prev) => ({ ...prev, [dirPath]: items }));
      return items;
    } catch (err) {
      flash(`Не удалось прочитать папку: ${(err as Error).message}`);
      return [];
    } finally {
      setLoadingDirs((prev) => {
        const next = new Set(prev);
        next.delete(dirPath);
        return next;
      });
    }
  }, [flash]);

  const openWorkspace = useCallback(
    async (dirPath: string) => {
      setRootPath(dirPath);
      setTree({});
      setExpandedDirs(new Set([dirPath]));
      // Открытые документы не закрываем: переключение избранной папки
      // не должно терять вкладки и несохранённые правки.
      await fileSystemService.setWorkspaceRoot(dirPath);
      await loadDirectory(dirPath);
    },
    [loadDirectory]
  );

  const refreshFavorites = useCallback(async () => {
    setFavorites(await fileSystemService.getFavorites());
  }, []);

  const toggleFavorite = useCallback(async (dirPath: string) => {
    const next = await fileSystemService.toggleFavorite(dirPath);
    setFavorites(next);
    const added = next.some((item) => samePath(item.path, dirPath));
    flash(added ? 'Добавлено в избранное' : 'Убрано из избранного');
  }, [flash]);

  const initWorkspace = useCallback(async () => {
    const dir = await fileSystemService.getCurrentWorkspace();
    await refreshFavorites();
    if (dir) await openWorkspace(dir);
  }, [openWorkspace, refreshFavorites]);

  const chooseWorkspaceFolder = useCallback(async () => {
    const dir = await fileSystemService.openFolderDialog();
    if (dir) await openWorkspace(dir);
    return dir;
  }, [openWorkspace]);

  const toggleDirectory = useCallback(
    async (dirPath: string) => {
      const isExpanded = expandedDirs.has(dirPath);
      if (isExpanded) {
        setExpandedDirs((prev) => {
          const next = new Set(prev);
          next.delete(dirPath);
          return next;
        });
      } else {
        setExpandedDirs((prev) => new Set(prev).add(dirPath));
        if (!tree[dirPath]) {
          await loadDirectory(dirPath);
        }
      }
    },
    [expandedDirs, tree, loadDirectory]
  );

  /** Полный рекурсивный сброс кэша дерева и его перезагрузка. */
  const refreshTree = useCallback(async () => {
    if (!rootPath) return;
    const previouslyExpanded = Array.from(expandedDirs);
    setTree({});
    await loadDirectory(rootPath);
    for (const dir of previouslyExpanded) {
      if (dir !== rootPath) await loadDirectory(dir);
    }
  }, [rootPath, expandedDirs, loadDirectory]);

  /* ---------------- Показ файла в дереве ---------------- */

  /** Файл, который нужно подсветить в проводнике (nonce — повторный показ). */
  const [revealTarget, setRevealTarget] = useState<{ path: string; nonce: number } | null>(null);

  /**
   * Раскрывает дерево до файла и просит панель подсветить строку.
   * Возвращает false, если файл вне текущего корня — тогда вызывающий меняет корень.
   */
  const revealPath = useCallback(
    async (filePath: string): Promise<boolean> => {
      if (!rootPath) return false;
      const root = normalizePath(rootPath);
      const target = normalizePath(filePath);
      if (target !== root && !target.startsWith(`${root}\\`)) return false;

      const base = rootPath.replace(/[\\/]+$/, '');
      const relative = filePath.slice(base.length).replace(/^[\\/]+/, '');
      const segments = relative.split(/[\\/]/).slice(0, -1);
      let cursor = base;
      for (const segment of segments) {
        if (!segment) continue;
        cursor = `${cursor}\\${segment}`;
        await loadDirectory(cursor);
        setExpandedDirs((prev) => new Set(prev).add(cursor));
      }
      setRevealTarget({ path: filePath, nonce: Date.now() });
      return true;
    },
    [loadDirectory, rootPath]
  );

  /* ---------------- Вкладки ---------------- */

  /** Определяет способ чтения файла и нормализует содержимое вкладки. */
  const readTabContent = useCallback(
    async (filePath: string, type: FileType, ext: string) => {
      // .doc — это OLE2-контейнер, а редактор умеет только .docx.
      // Текст извлекается в главном процессе и упаковывается в OOXML.
      if (type === 'word' && ext === 'doc') {
        const legacy = await fileSystemService.readFile(filePath, 'base64');
        const text = await extractLegacyDocText(legacy.content);
        const content = await htmlToDocxBase64(plainTextToHtml(text));
        return { content, encoding: 'base64' as const, isBinary: true };
      }
      const isBinary = BINARY_TYPES.includes(type) || isBinaryExcel(ext);
      if (isBinary) {
        const res = await fileSystemService.readFile(filePath, 'base64');
        if (type === 'powerpoint' && !res.content.trim()) {
          await fileSystemService.writeFile(filePath, BLANK_PPTX_BASE64, 'base64');
          return { content: BLANK_PPTX_BASE64, encoding: 'base64' as const, isBinary: true };
        }
        return { content: res.content, encoding: 'base64' as const, isBinary: true };
      }
      const res = await fileSystemService.readFile(filePath, 'utf-8');
      // CSV/TSV сохраняем как UTF-8 с BOM: иначе русский Excel открывает
      // файл в ANSI и показывает «кракозябры».
      const csvLike = ext === 'csv' || ext === 'tsv';
      return {
        content: res.content,
        encoding: csvLike ? 'utf-8-bom' : 'utf-8',
        isBinary: false,
      };
    },
    []
  );

  /* ---------------- Отслеживание изменений на диске ---------------- */

  // Подписка на события ФС одна на всё время жизни хука: expandedDirs меняет
  // идентичность на каждый expand/collapse, а переподписка теряет события,
  // пришедшие между рендерами.
  const watcherRef = useRef({ rootPath, expandedDirs, loadDirectory, readTabContent, flash });
  watcherRef.current = { rootPath, expandedDirs, loadDirectory, readTabContent, flash };

  useEffect(() => {
    if (!window.electronAPI?.onWorkspaceChanged) return;
    let timer: number | null = null;
    const pendingDirs = new Set<string>();

    const flushTree = async () => {
      timer = null;
      const dirs = Array.from(pendingDirs);
      pendingDirs.clear();
      for (const dir of dirs) await watcherRef.current.loadDirectory(dir);
    };

    const unsub = window.electronAPI.onWorkspaceChanged(async ({ fullPath }) => {
      const current = watcherRef.current;

      // 1. Перечитываем только затронутые каталоги и с дебаунсом: событий ФС
      // много (каждое сохранение), полное дерево на каждое — лишний трафик.
      if (current.rootPath) {
        if (fullPath) {
          const parent = dirName(fullPath);
          if (current.expandedDirs.has(fullPath)) pendingDirs.add(fullPath);
          if (parent !== fullPath && current.expandedDirs.has(parent)) pendingDirs.add(parent);
          if (parent === current.rootPath) pendingDirs.add(current.rootPath);
        } else {
          pendingDirs.add(current.rootPath);
        }
        if (pendingDirs.size > 0) {
          if (timer) window.clearTimeout(timer);
          timer = window.setTimeout(() => void flushTree(), 250);
        }
      }

      // 2. Проверяем, не открыт ли изменённый файл в одной из вкладок
      if (!fullPath) return;
      const openTab = tabsRef.current.find((t) => samePath(t.path, fullPath));
      if (!openTab) return;
      if (!openTab.isDirty) {
        // Файл не редактировался пользователем — автообновляем
        try {
          const { content } = await current.readTabContent(
            openTab.path,
            openTab.type,
            openTab.extension
          );
          setTabs((prev) =>
            prev.map((t) => {
              if (t.id !== openTab.id) return t;
              // Пока читали файл, вкладку могли начать править: прочитанное устарело.
              if (t.isDirty || t.content !== openTab.content) return t;
              return { ...t, content, initialContent: content, isDirty: false };
            })
          );
        } catch {
          // Файл могли удалить или переименовать — вкладку не трогаем.
        }
      } else {
        current.flash(`Внимание: открытый файл "${openTab.name}" изменён на диске`);
      }
    });

    return () => {
      if (timer) window.clearTimeout(timer);
      unsub();
    };
  }, []);

  const openFile = useCallback(
    async (
      filePath: string,
      targetPane: 'primary' | 'secondary' = 'primary',
      forceNewTab = false,
      preview = false
    ) => {
      const existing = tabsRef.current.find((t) => samePath(t.path, filePath));
      // Панели не показывают одну вкладку: уже открытая в основной остаётся в ней.
      const openInPane = (id: string) => {
        if (targetPane === 'secondary') {
          if (id !== activeTabIdRef.current) setSecondaryTabId(id);
          setSplitEnabled(true);
          return;
        }
        if (id === secondaryTabIdRef.current) setSecondaryTabId('');
        setActiveTabId(id);
      };
      if (existing && !forceNewTab) {
        if (!preview && existing.isPreview) {
          // Двойной клик по файлу из предпросмотра превращает ту же вкладку
          // в обычную редактируемую.
          setTabs((prev) =>
            prev.map((t) => (t.id === existing.id ? { ...t, isPreview: false } : t))
          );
        }
        if (preview) {
          const staleIds = new Set(
            tabsRef.current
              .filter((t) => t.isPreview && !t.isDirty && t.id !== existing.id)
              .map((t) => t.id)
          );
          if (staleIds.size > 0) {
            setTabs((prev) => prev.filter((t) => !staleIds.has(t.id)));
            setActiveTabId((cur) => (staleIds.has(cur) ? existing.id : cur));
            setSecondaryTabId((cur) => (staleIds.has(cur) ? '' : cur));
          }
        }
        openInPane(existing.id);
        return existing.id;
      }

      const key = normalizePath(filePath);
      const pending = pendingOpenRef.current.get(key);
      if (pending) {
        const id = await pending;
        if (id) {
          if (!preview) {
            setTabs((prev) =>
              prev.map((t) => (t.id === id ? { ...t, isPreview: false } : t))
            );
          }
          openInPane(id);
        }
        return id;
      }

      const previewToken = preview ? ++previewOpenSeqRef.current : 0;
      const openPromise = (async (): Promise<string | null> => {
        // Один клик по другому файлу закрывает предыдущий предпросмотр,
        // чтобы дерево не плодило десятки временных вкладок.
        if (preview && !forceNewTab) {
          const staleIds = new Set(
            tabsRef.current
              .filter((t) => t.isPreview && !t.isDirty && !samePath(t.path, filePath))
              .map((t) => t.id)
          );
          if (staleIds.size > 0) {
            setTabs((prev) => prev.filter((t) => !staleIds.has(t.id)));
            setActiveTabId((cur) => (staleIds.has(cur) ? '' : cur));
            setSecondaryTabId((cur) => (staleIds.has(cur) ? '' : cur));
          }
        }

        const ext = extName(filePath);
        const type = getFileType(ext);

        if (type === 'video' || type === 'pdf' || type === 'binary') {
          if (preview && previewToken !== previewOpenSeqRef.current) return null;
          const tab: OpenTab = {
            id: uid('tab'),
            path: filePath,
            name: baseName(filePath),
            extension: ext,
            type,
            content: '',
            initialContent: '',
            isDirty: false,
            encoding: 'path',
            isPreview: preview,
          };
          setTabs((prev) => [...prev, tab]);
          // Файл открыли заново: запрет на снимок, оставленный закрытием
          // вкладки, больше не действует.
          noteOpenedPath(tab.path);
          openInPane(tab.id);
          return tab.id;
        }

        try {
          const { content, encoding } = await readTabContent(filePath, type, ext);
          // Пока читали файл, пользователь кликнул другой — этот предпросмотр не нужен.
          if (preview && previewToken !== previewOpenSeqRef.current) return null;
          const tab: OpenTab = {
            id: uid('tab'),
            path: filePath,
            name: baseName(filePath),
            extension: ext,
            type,
            content,
            initialContent: content,
            isDirty: false,
            encoding,
            isPreview: preview,
            viewMode: type === 'markdown' || type === 'richtext' ? 'visual' : undefined,
          };

          setTabs((prev) => [...prev, tab]);
          noteOpenedPath(tab.path);
          openInPane(tab.id);
          return tab.id;
        } catch (err) {
          flash(`Ошибка открытия файла: ${(err as Error).message}`);
          return null;
        }
      })();

      pendingOpenRef.current.set(key, openPromise);
      try {
        return await openPromise;
      } finally {
        if (pendingOpenRef.current.get(key) === openPromise) {
          pendingOpenRef.current.delete(key);
        }
      }
    },
    [flash, readTabContent]
  );

  const closeTab = useCallback((tabId: string) => {
    const tabsBefore = tabsRef.current;
    const idx = tabsBefore.findIndex((t) => t.id === tabId);
    if (idx === -1) return;
    const closing = tabsBefore[idx];
    const rest = tabsBefore.filter((t) => t.id !== tabId);
    const nextActiveId =
      activeTabIdRef.current === tabId ? rest[Math.min(idx, rest.length - 1)]?.id ?? '' : activeTabIdRef.current;

    // Удаление — только функциональным апдейтером: `next`, собранный из ref,
    // отстаёт на коммит, из-за чего второй closeTab в одном тике воскрешал
    // закрытую вкладку и терял уже поставленные в очередь апдейтеры.
    setTabs((prev) => (prev.some((t) => t.id === tabId) ? prev.filter((t) => t.id !== tabId) : prev));
    setActiveTabId(nextActiveId);
    setSecondaryTabId((currentSecondary) => {
      if (currentSecondary !== tabId) {
        return currentSecondary === nextActiveId ? '' : currentSecondary;
      }
      const fallback = rest[0]?.id ?? '';
      return fallback === nextActiveId ? '' : fallback;
    });
    // Черновик закрытой вкладки иначе «восстановится» при следующем запуске.
    // Запрет ставится до удаления снимков: запись снимка могла идти параллельно.
    if (closing.isDirty) {
      noteClosedPath(closing.path);
      void dropSnapshotsForPath(closing.path);
    }
  }, []);

  /** Редактирование содержимого вкладки с пересчётом флага несохранённых изменений. */
  const updateTabContent = useCallback((tabId: string, content: string) => {
    setTabs((prev) =>
      prev.map((t) => {
        if (t.id !== tabId) return t;
        // PDF и прочие бинарные вкладки открываются в режиме «path»: содержимое
        // читает сам редактор. Как только редактор отдал данные, вкладка
        // становится base64 — иначе сохранение считает файл доступным только для чтения.
        const switchToBase64 = t.encoding === 'path' && Boolean(content) && (t.type === 'pdf' || t.type === 'binary');
        // Первый настоящий ввод снимает предпросмотр: иначе saveTab отказывает,
        // а клик по дереву теряет черновик. Программное изменение содержимого
        // (ресериализация при открытии) предпросмотр не снимает.
        const clearPreview =
          t.isPreview && content !== t.initialContent && userEditedTabsRef.current.has(t.id);
        return {
          ...t,
          content,
          encoding: switchToBase64 ? ('base64' as const) : t.encoding,
          isPreview: clearPreview ? false : t.isPreview,
          isDirty: content !== t.initialContent,
        };
      })
    );
  }, []);

  const patchTab = useCallback((tabId: string, patch: Partial<OpenTab>) => {
    setTabs((prev) => prev.map((t) => (t.id === tabId ? { ...t, ...patch } : t)));
  }, []);

  /**
   * Запись вкладки на диск.
   * contentOverride — живое содержимое редактора: tab.content обновляется через
   * setTabs и к моменту вызова отстаёт на один рендер (иначе на диск уходит
   * предыдущая версия, а UI рапортует успех).
   */
  const runSave = useCallback(
    async (tabId: string, contentOverride?: string): Promise<boolean> => {
      for (let attempt = 0; attempt < 2; attempt += 1) {
        // Вкладку перечитываем внутри задачи: между вызовом и записью её могли
        // переименовать, и путь из замыкания уже указывает в никуда.
        const tab = tabsRef.current.find((t) => t.id === tabId);
        if (!tab) return false;
        if (tab.isPreview) {
          flash('Это предпросмотр. Откройте файл двойным кликом или кнопкой «Редактировать»');
          return false;
        }
        if (tab.encoding === 'path') {
          flash('Этот файл открыт только для просмотра');
          return false;
        }
        const written = contentOverride ?? tab.content;
        try {
          // .doc перезаписать нельзя: внутри OLE2, а мы держим OOXML.
          // Сохраняем рядом .docx и переключаем вкладку на него.
          if (tab.extension === 'doc') {
            const target = docxPathFor(tab.path);
            await fileSystemService.writeFile(target, written, 'base64');
            setTabs((prev) =>
              prev.map((t) =>
                t.id === tabId
                  ? {
                      ...t,
                      path: target,
                      name: baseName(target),
                      extension: 'docx',
                      isDirty: t.content !== written,
                      initialContent: written,
                    }
                  : t
              )
            );
            // Снимок старого пути после записи не нужен: иначе следующий запуск
            // «восстановит» черновик поверх свежего файла.
            await dropSnapshotsForPath(tab.path);
            await refreshTree();
            flash(`Сохранено как ${baseName(target)} — исходный .doc не изменён`);
            return true;
          }
          // Начало записи: снимки, сделанные позже, содержат правки, которых
          // в записанном файле нет, и должны остаться.
          const writeStartedAt = Date.now();
          await fileSystemService.writeFile(tab.path, written, tab.encoding);
          const renamed = tabsRef.current.find((t) => t.id === tabId);
          // Путь сменился во время записи: пишем заново по актуальному.
          if (renamed && !samePath(renamed.path, tab.path)) continue;
          setTabs((prev) =>
            prev.map((t) =>
              t.id === tabId
                ? {
                    ...t,
                    // Правки, набранные во время записи, остаются несохранёнными.
                    isDirty: t.content !== written,
                    initialContent: written,
                  }
                : t
            )
          );
          await dropSnapshotsForPath(tab.path, writeStartedAt);
          flash(`Сохранено: ${tab.name}`);
          return true;
        } catch (err) {
          flash(`Ошибка сохранения: ${(err as Error).message}`);
          return false;
        }
      }
      flash('Сохранение отменено: файл переименован во время записи');
      return false;
    },
    [flash, refreshTree]
  );

  const saveTab = useCallback(
    async (tabId: string, contentOverride?: string): Promise<boolean> => {
      const previous = saveQueues.get(tabId) ?? Promise.resolve();
      const task = previous.then(
        () => runSave(tabId, contentOverride),
        () => runSave(tabId, contentOverride)
      );
      saveQueues.set(tabId, task);
      try {
        return await task;
      } finally {
        // Очередь не держим: иначе карта растёт на каждую закрытую вкладку.
        if (saveQueues.get(tabId) === task) saveQueues.delete(tabId);
      }
    },
    [runSave]
  );

  /**
   * Сохранение под новым именем/в другом формате.
   * contentOverride передаёт живое содержимое редактора (Word/Excel/PPT),
   * если оно ещё не успело попасть в tab.content.
   */
  const saveTabAs = useCallback(
    async (
      tabId: string,
      targetPath: string,
      format: ConversionFormat,
      contentOverride?: string,
      activeSheet?: string
    ): Promise<boolean> => {
      const tab = tabsRef.current.find((t) => t.id === tabId);
      if (!tab) return false;
      // Имя могло быть введено вручную с другим расширением: формат важнее
      // суффикса, иначе Excel откроет .csv с содержимым xlsx.
      const expectedExt = format === 'md' ? 'md' : format;
      if (extName(targetPath) !== expectedExt) {
        targetPath = targetPath.replace(/\.[^./\\]+$/, '') + `.${expectedExt}`;
      }
      try {
        const result = await convertDocument(
          {
            path: tab.path,
            name: tab.name,
            extension: tab.extension,
            type: tab.type,
            content: contentOverride ?? tab.content,
            encoding: tab.encoding,
            activeSheet,
          },
          format
        );

        if (result.kind === 'pdf') {
          const exported = await window.electronAPI?.exportPdf(
            result.html,
            baseName(targetPath),
            targetPath
          );
          if (!exported?.success) {
            throw new Error(exported?.error || 'Не удалось создать PDF');
          }
        } else {
          await fileSystemService.writeFile(targetPath, result.content, result.encoding);
        }

        await refreshTree();
        const targetExt = extName(targetPath);
        const targetName = baseName(targetPath);
        const existing = tabsRef.current.find((t) => samePath(t.path, targetPath));
        const nextContent =
          result.kind === 'data' ? result.content : existing?.content ?? '';
        if (existing) {
          patchTab(existing.id, {
            path: targetPath,
            name: targetName,
            extension: targetExt,
            type: getFileType(targetExt),
            content: nextContent,
            initialContent: nextContent,
            isDirty: false,
            reloadKey: Date.now(),
            isPreview: false,
            encoding:
              result.kind === 'data' ? result.encoding : existing.encoding,
          });
          setActiveTabId(existing.id);
        } else {
          await openFile(targetPath);
        }
        flash(
          `Сохранено как ${targetName}${result.note ? ` — ${result.note}` : ''}`
        );
        return true;
      } catch (err) {
        flash(`Ошибка сохранения как: ${(err as Error).message}`);
        return false;
      }
    },
    [flash, refreshTree, patchTab, openFile]
  );

  /** Конвертация произвольного файла рабочей папки (для AI-агента). */
  const convertFile = useCallback(
    async (
      sourcePath: string,
      targetPath: string,
      format: ConversionFormat
    ): Promise<{ ok: boolean; message: string; path?: string }> => {
      try {
        const source = await readFileSource(sourcePath);
        const result = await convertDocument(source, format);
        if (result.kind === 'pdf') {
          const exported = await window.electronAPI?.exportPdf(
            result.html,
            baseName(targetPath),
            targetPath
          );
          if (!exported?.success) {
            return {
              ok: false,
              message: exported?.error || 'Не удалось создать PDF',
            };
          }
        } else {
          await fileSystemService.writeFile(targetPath, result.content, result.encoding);
        }
        await refreshTree();
        await openFile(targetPath);
        return {
          ok: true,
          message: `Конвертировано: ${baseName(sourcePath)} → ${baseName(targetPath)}${
            result.note ? ` (${result.note})` : ''
          }`,
          path: targetPath,
        };
      } catch (err) {
        return {
          ok: false,
          message: `Не удалось конвертировать ${baseName(sourcePath)}: ${(err as Error).message}`,
        };
      }
    },
    [refreshTree, openFile]
  );

  /**
   * Сохраняет все изменённые вкладки.
   * Возвращает число сохранённых файлов и имена тех, что сохранить не удалось,
   * — вызывающий код (например, подтверждение выхода) должен знать о сбоях.
   */
  const saveAll = useCallback(async () => {
    const dirty = tabsRef.current.filter((t) => t.isDirty);
    const failed: string[] = [];
    let saved = 0;

    for (const tab of dirty) {
      const ok = await saveTab(tab.id);
      if (ok) saved++;
      else failed.push(tab.name);
    }

    return { saved, failed };
  }, [saveTab]);

  /* ---------------- Файловые операции ---------------- */

  const createFile = useCallback(
    async (
      parentDir: string,
      name: string,
      content = '',
      encoding?: string
    ): Promise<string | null> => {
      const targetDir = parentDir || rootPath || (await fileSystemService.getCurrentWorkspace()) || '.';
      const target = joinPath(targetDir, name);
      try {
        // Пустой .docx/.xlsx — недействительный ZIP: Word и Excel такие файлы
        // не открывают, а встроенные редакторы падают на разборе пакета.
        // Сразу пишем корректный минимальный пакет.
        const ext = extName(name).toLowerCase();
        let payload = content;
        let payloadEncoding = encoding;
        if (!content) {
          if (ext === 'docx') {
            payload = await generateDocxBase64('');
            payloadEncoding = 'base64';
          } else if (ext === 'xlsx' || ext === 'xlsm') {
            payload = generateExcelBase64([], [], 'Лист 1');
            payloadEncoding = 'base64';
          } else if (ext === 'pptx') {
            payload = BLANK_PPTX_BASE64;
            payloadEncoding = 'base64';
          }
        }

        // Антивирус или git могут на мгновение блокировать каталог — повторяем
        // запись, иначе пользователь видит EPERM при полностью свободной папке.
        let created = false;
        let lastError: unknown = null;
        for (let attempt = 0; attempt < 4 && !created; attempt++) {
          try {
            await fileSystemService.createFile(target, payload, payloadEncoding);
            created = true;
          } catch (err) {
            lastError = err;
            const message = (err as Error).message || '';
            if (!/EPERM|EBUSY|EACCES/i.test(message)) throw err;
            await new Promise((resolve) => setTimeout(resolve, 120 + attempt * 180));
          }
        }
        if (!created) throw lastError ?? new Error('не удалось создать файл');
        if (tree[targetDir]) await loadDirectory(targetDir);
        else await refreshTree();
        const tabId = await openFile(target);
        if (!tabId) return null;
        flash(`Создан файл: ${name}`);
        return target;
      } catch (err) {
        const message = (err as Error).message || '';
        flash(
          /EPERM|EACCES/i.test(message)
            ? `Нет прав на запись в «${targetDir.split(/[\\/]/).pop() || targetDir}». Выберите другую папку`
            : `Не удалось создать файл: ${message}`
        );
        return null;
      }
    },
    [rootPath, tree, loadDirectory, refreshTree, openFile, flash]
  );

  const createFolder = useCallback(
    async (parentDir: string, name: string): Promise<string | null> => {
      const target = joinPath(parentDir, name);
      try {
        await fileSystemService.createDirectory(target);
        if (tree[parentDir]) await loadDirectory(parentDir);
        else await refreshTree();
        setExpandedDirs((prev) => new Set(prev).add(parentDir));
        flash(`Создана папка: ${name}`);
        return target;
      } catch (err) {
        flash(`Не удалось создать папку: ${(err as Error).message}`);
        return null;
      }
    },
    [tree, loadDirectory, refreshTree, flash]
  );

  const renamePath = useCallback(
    async (oldPath: string, newName: string): Promise<boolean> => {
      // dirName не режет последний символ у пути без разделителя (slice(0, -1)).
      const parent = dirName(oldPath);
      const newPath = joinPath(parent, newName);
      try {
        await fileSystemService.renamePath(oldPath, newPath);
        // Обновляем открытые вкладки, указывающие на переименованный путь
        setTabs((prev) =>
          prev.map((t) =>
            samePath(t.path, oldPath)
              ? { ...t, path: newPath, name: newName, extension: extName(newName) }
              : t
          )
        );
        await refreshTree();
        flash(`Переименовано в: ${newName}`);
        return true;
      } catch (err) {
        flash(`Ошибка переименования: ${(err as Error).message}`);
        return false;
      }
    },
    [refreshTree, flash]
  );

  const deletePath = useCallback(
    async (targetPath: string): Promise<boolean> => {
      try {
        await fileSystemService.deletePath(targetPath);
        // Закрываем вкладки удалённого файла или файлов внутри удалённой папки.
        // Выбор активной — снаружи апдейтера: он должен оставаться чистым.
        const prev = tabsRef.current;
        const target = normalizePath(targetPath);
        const remaining = prev.filter((t) => {
          const tabPath = normalizePath(t.path);
          return tabPath !== target && !tabPath.startsWith(`${target}\\`);
        });
        if (remaining.length !== prev.length) {
          const removed = prev.filter((t) => !remaining.includes(t));
          setTabs(remaining);
          setActiveTabId((cur) =>
            remaining.some((t) => t.id === cur) ? cur : remaining[0]?.id || ''
          );
          setSecondaryTabId((cur) => (remaining.some((t) => t.id === cur) ? cur : ''));
          // Черновик удалённого файла не должен всплыть при следующем запуске.
          for (const tab of removed) if (tab.isDirty) void dropSnapshotsForPath(tab.path);
        }
        await refreshTree();
        flash(`Удалено: ${baseName(targetPath)}`);
        return true;
      } catch (err) {
        flash(`Ошибка удаления: ${(err as Error).message}`);
        return false;
      }
    },
    [refreshTree, flash]
  );

  /* ---------------- Производные значения ---------------- */

  const activeTab = useMemo(
    () => tabs.find((t) => t.id === activeTabId) || null,
    [tabs, activeTabId]
  );

  const secondaryTab = useMemo(
    () => tabs.find((t) => t.id === secondaryTabId) || null,
    [tabs, secondaryTabId]
  );

  const dirtyTabs = useMemo(() => tabs.filter((t) => t.isDirty), [tabs]);

  /** Плоский список файлов рабочей папки для быстрого перехода (Ctrl+P). */
  const allFiles = useMemo(() => {
    const result: FileItem[] = [];
    for (const items of Object.values(tree)) {
      for (const item of items) {
        if (!item.isDirectory) result.push(item);
      }
    }
    return result;
  }, [tree]);

  return {
    // рабочая папка
    rootPath,
    tree,
    expandedDirs,
    loadingDirs,
    openWorkspace,
    initWorkspace,
    chooseWorkspaceFolder,
    favorites,
    toggleFavorite,
    toggleDirectory,
    loadDirectory,
    refreshTree,
    revealPath,
    revealTarget,
    // вкладки
    tabs,
    activeTabId,
    activeTab,
    secondaryTabId,
    secondaryTab,
    dirtyTabs,
    openFile,
    closeTab,
    updateTabContent,
    patchTab,
    saveTab,
    saveTabAs,
    convertFile,
    saveAll,
    setActiveTabId,
    setSecondaryTabId,
    // раскладка
    splitEnabled,
    setSplitEnabled,
    splitRatio,
    setSplitRatio,
    // файловые операции
    createFile,
    createFolder,
    renamePath,
    deletePath,
    // прочее
    allFiles,
    statusMessage,
    flash,
  };
}

export type WorkspaceApi = ReturnType<typeof useWorkspace>;
