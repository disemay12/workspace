import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { FileItem, OpenTab, FileType } from '../types';
import { fileSystemService, getFileType } from '../services/fileSystemService';
import { extName, baseName, joinPath, uid } from '../lib/utils';
import { generateDocxBase64, generateExcelBase64 } from '../services/aiActions';
import { htmlToDocxBase64 } from '../lib/docx';
import { docxPathFor, extractLegacyDocText, plainTextToHtml } from '../lib/legacyDoc';

/** Типы файлов, содержимое которых читается как base64 (бинарные офисные форматы). */
const BINARY_TYPES: FileType[] = ['word', 'powerpoint', 'image'];
/** Excel читается как base64 (xlsx/xls), CSV/TSV — как текст. */
const isBinaryExcel = (ext: string) => ext === 'xlsx' || ext === 'xls';

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
  tabsRef.current = tabs;

  const flash = useCallback((message: string) => {
    setStatusMessage(message);
    window.setTimeout(() => {
      setStatusMessage((prev) => (prev === message ? '' : prev));
    }, 3200);
  }, []);

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
    const added = next.some((item) => item.path === dirPath);
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
        return { content: res.content, encoding: 'base64' as const, isBinary: true };
      }
      const res = await fileSystemService.readFile(filePath, 'utf-8');
      return { content: res.content, encoding: 'utf-8' as const, isBinary: false };
    },
    []
  );

  /* ---------------- Отслеживание изменений на диске ---------------- */
  useEffect(() => {
    if (!window.electronAPI?.onWorkspaceChanged) return;
    const unsub = window.electronAPI.onWorkspaceChanged(async ({ fullPath }) => {
      // 1. Обновляем файловое дерево для открытых каталогов
      if (rootPath) {
        const expanded = Array.from(expandedDirs);
        for (const dir of expanded) {
          await loadDirectory(dir);
        }
      }

      // 2. Проверяем, не открыт ли изменённый файл в одной из вкладок
      if (fullPath) {
        const normalizedChanged = fullPath.toLowerCase().replace(/\\/g, '/');
        const openTab = tabsRef.current.find(
          (t) => t.path.toLowerCase().replace(/\\/g, '/') === normalizedChanged
        );
        if (openTab) {
          if (!openTab.isDirty) {
            // Файл не редактировался пользователем — автообновляем
            try {
              const { content } = await readTabContent(openTab.path, openTab.type, openTab.extension);
              setTabs((prev) =>
                prev.map((t) =>
                  t.id === openTab.id ? { ...t, content, initialContent: content, isDirty: false } : t
                )
              );
            } catch {
              // Игнорируем ошибку чтения (например, если файл удалён)
            }
          } else {
            flash(`Внимание: открытый файл "${openTab.name}" изменён на диске`);
          }
        }
      }
    });
    return unsub;
  }, [rootPath, expandedDirs, loadDirectory, readTabContent, flash]);

  const openFile = useCallback(
    async (
      filePath: string,
      targetPane: 'primary' | 'secondary' = 'primary',
      forceNewTab = false
    ) => {
      const existing = tabsRef.current.find((t) => t.path === filePath);
      if (existing && !forceNewTab) {
        if (targetPane === 'secondary') {
          setSecondaryTabId(existing.id);
          setSplitEnabled(true);
        } else {
          setActiveTabId(existing.id);
        }
        return existing.id;
      }

      const ext = extName(filePath);
      const type = getFileType(ext);

      if (type === 'video' || type === 'pdf' || type === 'binary') {
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
        };
        setTabs((prev) => [...prev, tab]);
        if (targetPane === 'secondary') {
          setSecondaryTabId(tab.id);
          setSplitEnabled(true);
        } else {
          setActiveTabId(tab.id);
        }
        return tab.id;
      }

      try {
        const { content, encoding } = await readTabContent(filePath, type, ext);
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
          viewMode: type === 'markdown' || type === 'richtext' ? 'visual' : undefined,
        };

        setTabs((prev) => [...prev, tab]);
        if (targetPane === 'secondary') {
          setSecondaryTabId(tab.id);
          setSplitEnabled(true);
        } else {
          setActiveTabId(tab.id);
        }
        return tab.id;
      } catch (err) {
        flash(`Ошибка открытия файла: ${(err as Error).message}`);
        return null;
      }
    },
    [flash, readTabContent]
  );

  const closeTab = useCallback((tabId: string) => {
    setTabs((prev) => {
      const idx = prev.findIndex((t) => t.id === tabId);
      if (idx === -1) return prev;
      const next = prev.filter((t) => t.id !== tabId);

      setActiveTabId((currentActive) => {
        if (currentActive !== tabId) return currentActive;
        const fallback = next[Math.min(idx, next.length - 1)];
        return fallback ? fallback.id : '';
      });
      setSecondaryTabId((currentSecondary) => {
        if (currentSecondary !== tabId) return currentSecondary;
        const fallback = next.find((t) => t.id !== tabId);
        return fallback ? fallback.id : '';
      });
      return next;
    });
  }, []);

  /** Редактирование содержимого вкладки с пересчётом флага несохранённых изменений. */
  const updateTabContent = useCallback((tabId: string, content: string) => {
    setTabs((prev) =>
      prev.map((t) =>
        t.id === tabId ? { ...t, content, isDirty: content !== t.initialContent } : t
      )
    );
  }, []);

  const patchTab = useCallback((tabId: string, patch: Partial<OpenTab>) => {
    setTabs((prev) => prev.map((t) => (t.id === tabId ? { ...t, ...patch } : t)));
  }, []);

  const saveTab = useCallback(
    async (tabId: string): Promise<boolean> => {
      const tab = tabsRef.current.find((t) => t.id === tabId);
      if (!tab) return false;
      if (tab.encoding === 'path') {
        flash('Этот файл открыт только для просмотра');
        return false;
      }
      try {
        // .doc перезаписать нельзя: внутри OLE2, а мы держим OOXML.
        // Сохраняем рядом .docx и переключаем вкладку на него.
        if (tab.extension === 'doc') {
          const target = docxPathFor(tab.path);
          await fileSystemService.writeFile(target, tab.content, 'base64');
          setTabs((prev) =>
            prev.map((t) =>
              t.id === tabId
                ? {
                    ...t,
                    path: target,
                    name: baseName(target),
                    extension: 'docx',
                    isDirty: false,
                    initialContent: t.content,
                  }
                : t
            )
          );
          await refreshTree();
          flash(`Сохранено как ${baseName(target)} — исходный .doc не изменён`);
          return true;
        }
        await fileSystemService.writeFile(tab.path, tab.content, tab.encoding);
        setTabs((prev) =>
          prev.map((t) => (t.id === tabId ? { ...t, isDirty: false, initialContent: t.content } : t))
        );
        flash(`Сохранено: ${tab.name}`);
        return true;
      } catch (err) {
        flash(`Ошибка сохранения: ${(err as Error).message}`);
        return false;
      }
    },
    [flash, refreshTree]
  );

  /** Сохранение файла под новым именем/расширением (Экспорт как…). */
  const saveTabAs = useCallback(
    async (tabId: string, suggestedName: string): Promise<boolean> => {
      const tab = tabsRef.current.find((t) => t.id === tabId);
      if (!tab) return false;
      const target = await fileSystemService.saveFileDialog(suggestedName);
      if (!target) return false;
      try {
        await fileSystemService.writeFile(target, tab.content, tab.encoding);
        flash(`Экспортировано: ${baseName(target)}`);
        return true;
      } catch (err) {
        flash(`Ошибка экспорта: ${(err as Error).message}`);
        return false;
      }
    },
    [flash]
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
      const parent = oldPath.slice(0, Math.max(oldPath.lastIndexOf('\\'), oldPath.lastIndexOf('/')));
      const newPath = joinPath(parent, newName);
      try {
        await fileSystemService.renamePath(oldPath, newPath);
        // Обновляем открытые вкладки, указывающие на переименованный путь
        setTabs((prev) =>
          prev.map((t) =>
            t.path === oldPath
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
        // Закрываем вкладки удалённого файла или файлов внутри удалённой папки
        setTabs((prev) => {
          const remaining = prev.filter(
            (t) => t.path !== targetPath && !t.path.startsWith(`${targetPath}\\`) && !t.path.startsWith(`${targetPath}/`)
          );
          if (remaining.length !== prev.length) {
            setActiveTabId((cur) =>
              remaining.some((t) => t.id === cur) ? cur : remaining[0]?.id || ''
            );
          }
          return remaining;
        });
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
