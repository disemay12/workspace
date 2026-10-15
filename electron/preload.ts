import { contextBridge, ipcRenderer } from 'electron';

export interface FileItem {
  name: string;
  path: string;
  isDirectory: boolean;
  size: number;
  mtime: string | null;
  birthtime?: string | null;
  extension: string;
}

export interface SearchResult {
  path: string;
  name: string;
  isDirectory: boolean;
  line?: number;
  matchSnippet?: string;
  totalMatches?: number;
}

const electronAPI = {
  /* ---------- Управление окном ---------- */
  minimizeWindow: (): Promise<void> => ipcRenderer.invoke('window:minimize'),
  maximizeWindow: (): Promise<boolean> => ipcRenderer.invoke('window:maximize'),
  isWindowMaximized: (): Promise<boolean> => ipcRenderer.invoke('window:isMaximized'),
  closeWindow: (): Promise<void> => ipcRenderer.invoke('window:close'),

  /**
   * Ответ на запрос подтверждения закрытия окна.
   * proceed = true — закрыть приложение, false — остаться.
   */
  confirmClose: (proceed: boolean): Promise<{ closed: boolean }> =>
    ipcRenderer.invoke('window:confirmClose', proceed),

  /**
   * Подписка на запрос подтверждения закрытия. Приходит при попытке
   * закрыть окно (системная кнопка или кнопка в заголовке), если главный
   * процесс перехватил закрытие. Возвращает функцию отписки.
   */
  onCloseRequested: (callback: () => void): (() => void) => {
    const listener = () => callback();
    ipcRenderer.on('window:closeRequested', listener);
    return () => ipcRenderer.removeListener('window:closeRequested', listener);
  },

  /** Подписка на изменение состояния окна. Возвращает функцию отписки. */
  onWindowMaximizeChange: (callback: (isMaximized: boolean) => void): (() => void) => {
    const listener = (_event: unknown, isMaximized: boolean) => callback(isMaximized);
    ipcRenderer.on('window:maximizeChanged', listener);
    return () => ipcRenderer.removeListener('window:maximizeChanged', listener);
  },

  /* ---------- Рабочая папка и диалоги ---------- */
  getCurrentWorkspace: (): Promise<string> => ipcRenderer.invoke('fs:getCurrentWorkspace'),
  setWorkspaceRoot: (dirPath: string): Promise<{ success: boolean }> =>
    ipcRenderer.invoke('fs:setWorkspaceRoot', dirPath),
  getFavorites: (): Promise<{ path: string; name: string }[]> => ipcRenderer.invoke('fs:getFavorites'),
  toggleFavorite: (
    dirPath: string
  ): Promise<{ success: boolean; favorites: { path: string; name: string }[] }> =>
    ipcRenderer.invoke('fs:toggleFavorite', dirPath),
  openFolderDialog: (): Promise<string | null> => ipcRenderer.invoke('fs:openFolderDialog'),
  openFileDialog: (filters?: { name: string; extensions: string[] }[]): Promise<string | null> =>
    ipcRenderer.invoke('fs:openFileDialog', filters),
  saveFileDialog: (
    defaultName: string,
    filters?: { name: string; extensions: string[] }[]
  ): Promise<string | null> => ipcRenderer.invoke('fs:saveFileDialog', defaultName, filters),
  showInExplorer: (filePath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:showInExplorer', filePath),

  /* ---------- Файловая система ---------- */
  listDirectory: (dirPath: string): Promise<{ success: boolean; items: FileItem[]; error?: string }> =>
    ipcRenderer.invoke('fs:listDirectory', dirPath),
  readFile: (
    filePath: string,
    encoding = 'utf-8'
  ): Promise<{ success: boolean; content?: string; isBinary?: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:readFile', filePath, encoding),
  writeFile: (
    filePath: string,
    content: string,
    encoding = 'utf-8'
  ): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:writeFile', filePath, content, encoding),
  createFile: (
    filePath: string,
    initialContent = ''
  ): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:createFile', filePath, initialContent),
  createDirectory: (dirPath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:createDirectory', dirPath),
  deletePath: (targetPath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:deletePath', targetPath),
  renamePath: (oldPath: string, newPath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:renamePath', oldPath, newPath),

  /* ---------- Прокси запросов к AI ---------- */
  /**
   * Запуск стримингового запроса к провайдеру в главном процессе.
   * Обходит CORS, которому подчиняется браузерный fetch.
   */
  aiStream: (req: {
    requestId: string;
    url: string;
    method: string;
    headers: Record<string, string>;
    body: string;
  }): Promise<{ started: boolean }> => ipcRenderer.invoke('ai:stream', req),

  /** Прерывание активного запроса. */
  aiAbort: (requestId: string): Promise<{ aborted: boolean }> =>
    ipcRenderer.invoke('ai:abort', requestId),

  /** Проверка доступности endpoint. */
  aiProbe: (url: string): Promise<{ reachable: boolean; status?: number; error?: string }> =>
    ipcRenderer.invoke('ai:probe', url),

  /** Запрос списка моделей от API провайдера без ограничений CORS. */
  aiListModels: (req: {
    provider: string;
    baseUrl?: string;
    apiKey?: string;
  }): Promise<{ success: boolean; models: string[]; error?: string }> =>
    ipcRenderer.invoke('ai:listModels', req),

  /** Подписки на события потока. Возвращают функцию отписки. */
  onAiChunk: (cb: (data: { requestId: string; chunk: string }) => void): (() => void) => {
    const listener = (_e: unknown, data: { requestId: string; chunk: string }) => cb(data);
    ipcRenderer.on('ai:chunk', listener);
    return () => ipcRenderer.removeListener('ai:chunk', listener);
  },
  onAiEnd: (cb: (data: { requestId: string }) => void): (() => void) => {
    const listener = (_e: unknown, data: { requestId: string }) => cb(data);
    ipcRenderer.on('ai:end', listener);
    return () => ipcRenderer.removeListener('ai:end', listener);
  },
  onAiError: (
    cb: (data: { requestId: string; message: string; status?: number; isNetworkError?: boolean }) => void
  ): (() => void) => {
    const listener = (
      _e: unknown,
      data: { requestId: string; message: string; status?: number; isNetworkError?: boolean }
    ) => cb(data);
    ipcRenderer.on('ai:error', listener);
    return () => ipcRenderer.removeListener('ai:error', listener);
  },

  /* ---------- Локальные секреты (secrets.local.ps1) ---------- */
  /** Метаданные файла секретов — без самого ключа. */
  describeSecrets: (): Promise<{
    found: boolean;
    filePath: string;
    baseUrl?: string;
    textModel?: string;
    imageModel?: string;
    hasKey: boolean;
    hasOtherSecrets: boolean;
    error?: string;
  }> => ipcRenderer.invoke('secrets:describe'),

  /** Импорт API-ключа из локального файла — только по действию пользователя. */
  importSecretsKey: (): Promise<{
    success: boolean;
    apiKey?: string;
    baseUrl?: string;
    model?: string;
    error?: string;
  }> => ipcRenderer.invoke('secrets:importKey'),

  /* ---------- Поиск ---------- */
  searchWorkspace: (
    rootDir: string,
    query: string,
    searchContent = false,
    options?: { caseSensitive?: boolean; useRegex?: boolean; includeGlob?: string }
  ): Promise<{ success: boolean; results: SearchResult[]; error?: string }> =>
    ipcRenderer.invoke('fs:searchWorkspace', rootDir, query, searchContent, options),
  /* ---------- Буфер обмена ---------- */
  /** Читает текст и HTML из системного буфера обмена. */
  readClipboard: (): Promise<{ success: boolean; text: string; html: string; error?: string }> =>
    ipcRenderer.invoke('clipboard:read'),

  /** Записывает текст в системный буфер обмена. */
  writeClipboard: (text: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('clipboard:write', text),

  /* ---------- Экспорт в сервисы обмена ---------- */
  /**
   * Загружает файл (base64) в публичный сервис обмена.
   * Запрос идёт из главного процесса — обходит CORS файлообменников.
   */
  uploadToService: (req: {
    service: string;
    fileName: string;
    mime: string;
    base64: string;
    clientId?: string;
  }): Promise<{ success: boolean; url?: string; error?: string }> =>
    ipcRenderer.invoke('net:upload', req),

  /** Несколько изображений в один альбом Imgur. */
  uploadAlbum: (req: {
    service: string;
    title?: string;
    clientId?: string;
    images: { fileName: string; mime: string; base64: string; title?: string }[];
  }): Promise<{
    success: boolean;
    url?: string;
    uploaded: number;
    failed: string[];
    links: string[];
    error?: string;
  }> => ipcRenderer.invoke('net:uploadAlbum', req),

  /* ---------- Google Drive (OAuth 2.0) ---------- */
  gdriveGetConfig: (): Promise<{ clientId: string; clientSecret: string }> =>
    ipcRenderer.invoke('gdrive:getConfig'),
  gdriveSetConfig: (cfg: {
    clientId: string;
    clientSecret: string;
  }): Promise<{ success: boolean }> => ipcRenderer.invoke('gdrive:setConfig', cfg),
  gdriveStatus: (): Promise<{ configured: boolean; connected: boolean; email?: string }> =>
    ipcRenderer.invoke('gdrive:status'),
  gdriveConnect: (): Promise<{ success: boolean; email?: string; error?: string }> =>
    ipcRenderer.invoke('gdrive:connect'),
  gdriveDisconnect: (): Promise<{ success: boolean }> => ipcRenderer.invoke('gdrive:disconnect'),
  gdriveUpload: (
    paths: string[],
    folderName?: string
  ): Promise<{
    success: boolean;
    uploaded: number;
    failed: string[];
    link?: string;
    error?: string;
  }> => ipcRenderer.invoke('gdrive:upload', paths, folderName),

  videoCapabilities: (): Promise<{ ffmpeg: string | null; ffprobe: string | null }> =>
    ipcRenderer.invoke('video:capabilities'),
  videoProbe: (filePath: string): Promise<{
    success: boolean;
    duration?: number;
    width?: number;
    height?: number;
    hasAudio?: boolean;
    videoCodec?: string;
    audioCodec?: string;
    size?: number;
    mime?: string;
    error?: string;
  }> => ipcRenderer.invoke('video:probe', filePath),
  videoEdit: (req: {
    filePath: string;
    start?: number;
    end?: number;
    stripAudio?: boolean;
    compress?: boolean;
    outputPath?: string;
  }): Promise<{ success: boolean; outputPath?: string; error?: string }> =>
    ipcRenderer.invoke('video:edit', req),
  youtubeCategories: (): Promise<{ success: boolean; categories: { id: string; title: string }[]; error?: string }> =>
    ipcRenderer.invoke('youtube:categories'),
  youtubePlaylists: (): Promise<{ success: boolean; playlists: { id: string; title: string }[]; error?: string }> =>
    ipcRenderer.invoke('youtube:playlists'),
  youtubeUpload: (
    filePath: string,
    meta: {
      title: string;
      description?: string;
      tags?: string[];
      categoryId?: string;
      privacyStatus: 'public' | 'unlisted' | 'private';
      madeForKids: boolean;
      embeddable?: boolean;
      license?: 'youtube' | 'creativeCommon';
      defaultLanguage?: string;
      publishAt?: string;
      playlistId?: string;
    }
  ): Promise<{ success: boolean; id?: string; url?: string; error?: string }> =>
    ipcRenderer.invoke('youtube:upload', filePath, meta),
  openImgbb: (paths: string[]): Promise<{ success: boolean; count?: number; error?: string }> =>
    ipcRenderer.invoke('imagehost:open', 'imgbb', paths),
  openFreeimage: (paths: string[]): Promise<{ success: boolean; count?: number; error?: string }> =>
    ipcRenderer.invoke('imagehost:open', 'freeimage', paths),
  onYoutubeProgress: (cb: (data: { filePath: string; sent: number; total: number }) => void): (() => void) => {
    const listener = (_e: unknown, data: { filePath: string; sent: number; total: number }) => cb(data);
    ipcRenderer.on('youtube:progress', listener);
    return () => ipcRenderer.removeListener('youtube:progress', listener);
  },

  /** Постоянные настройки AI: переживают перезапуск и пересборку. */
  loadAiSettings: (): Promise<{
    settings: {
      provider?: string;
      baseUrl?: string;
      model?: string;
      temperature?: number;
      maxTokens?: number;
      systemPrompt?: string;
      activePresetId?: string;
      modelsCache?: Record<string, string[]>;
    };
    keys: Record<string, string>;
  }> => ipcRenderer.invoke('aiSettings:load'),
  saveAiSettings: (patch: {
    provider?: string;
    baseUrl?: string;
    model?: string;
    temperature?: number;
    maxTokens?: number;
    systemPrompt?: string;
    activePresetId?: string;
    modelsCache?: Record<string, string[]>;
    apiKey?: string;
  }): Promise<{ success: boolean }> => ipcRenderer.invoke('aiSettings:save', patch),

  /* ---------- Отслеживание изменений на диске (Watcher) ---------- */
  onWorkspaceChanged: (
    cb: (data: { eventType: string; filename: string; fullPath: string }) => void
  ): (() => void) => {
    const listener = (
      _e: unknown,
      data: { eventType: string; filename: string; fullPath: string }
    ) => cb(data);
    ipcRenderer.on('fs:workspaceChanged', listener);
    return () => ipcRenderer.removeListener('fs:workspaceChanged', listener);
  },

  /* ---------- Встроенный терминал ---------- */
  terminalCreate: (req?: {
    id?: string;
    cwd?: string;
    shell?: string;
  }): Promise<{ success: boolean; id?: string; error?: string }> =>
    ipcRenderer.invoke('terminal:create', req),

  terminalInput: (req: { id: string; data: string }): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('terminal:input', req),

  terminalKill: (id: string): Promise<{ success: boolean }> =>
    ipcRenderer.invoke('terminal:kill', id),

  onTerminalData: (cb: (data: { id: string; data: string }) => void): (() => void) => {
    const listener = (_e: unknown, data: { id: string; data: string }) => cb(data);
    ipcRenderer.on('terminal:data', listener);
    return () => ipcRenderer.removeListener('terminal:data', listener);
  },

  onTerminalExit: (cb: (data: { id: string; code: number }) => void): (() => void) => {
    const listener = (_e: unknown, data: { id: string; code: number }) => cb(data);
    ipcRenderer.on('terminal:exit', listener);
    return () => ipcRenderer.removeListener('terminal:exit', listener);
  },
};

contextBridge.exposeInMainWorld('electronAPI', electronAPI);
