import { contextBridge, ipcRenderer, webUtils } from 'electron';

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
  /** Закрепить/открепить окно поверх остальных приложений. */
  setAlwaysOnTop: (value: boolean): Promise<{ success: boolean; alwaysOnTop: boolean }> =>
    ipcRenderer.invoke('window:setAlwaysOnTop', value),
  getAlwaysOnTop: (): Promise<boolean> => ipcRenderer.invoke('window:getAlwaysOnTop'),
  /** Открыть таблицу в настоящем Excel (точные формулы, сводные, диаграммы, VBA). */
  openInExcel: (filePath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('excel:openFile', filePath),
  /** Выполнить VBA-макрос в Excel и сохранить книгу. */
  runExcelMacro: (
    filePath: string,
    macroName: string
  ): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('excel:runMacro', filePath, macroName),
  /** Перенести диаграммы WorkSpace в xlsx как нативные объекты Excel. */
  excelExportCharts: (
    filePath: string,
    charts: {
      sheet: string;
      type: 'bar' | 'line' | 'pie';
      range: string;
      title: string;
      x: number;
      y: number;
      width: number;
      height: number;
    }[]
  ): Promise<{ success: boolean; count?: number; error?: string }> =>
    ipcRenderer.invoke('excel:exportCharts', filePath, charts),
  /** Создать настоящую обновляемую сводную таблицу Excel. */
  excelCreatePivot: (
    filePath: string,
    spec: {
      sourceSheet: string;
      range: string;
      rowField: string;
      valueField: string;
      aggregation: 'sum' | 'count' | 'average' | 'min' | 'max';
      targetSheet: string;
    }
  ): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('excel:createPivot', filePath, spec),
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
    filters?: { name: string; extensions: string[] }[],
    defaultDir?: string
  ): Promise<string | null> => ipcRenderer.invoke('fs:saveFileDialog', defaultName, filters, defaultDir),
  pathExists: (targetPath: string): Promise<boolean> =>
    ipcRenderer.invoke('fs:pathExists', targetPath),
  exportPdf: (
    html: string,
    suggestedName?: string,
    targetPath?: string,
    pageSize?: { width: number; height: number }
  ): Promise<{ success: boolean; path?: string; canceled?: boolean; error?: string }> =>
    ipcRenderer.invoke('pdf:export', html, suggestedName, targetPath, pageSize),
  exportSlideImages: (
    htmlPages: string[],
    outDir: string,
    width?: number,
    height?: number
  ): Promise<{ success: boolean; files?: string[]; dir?: string; error?: string }> =>
    ipcRenderer.invoke('ppt:exportImages', htmlPages, outDir, width, height),
  uniquePath: (dirPath: string, fileName: string): Promise<string> =>
    ipcRenderer.invoke('fs:uniquePath', dirPath, fileName),
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
  readPreview: (
    filePath: string,
    maxBytes?: number
  ): Promise<{ success: boolean; base64?: string; size?: number; error?: string }> =>
    ipcRenderer.invoke('fs:readPreview', filePath, maxBytes),
  /** Размер и время изменения без чтения содержимого. */
  statFile: (
    filePath: string
  ): Promise<{
    success: boolean;
    size?: number;
    mtimeMs?: number;
    isDirectory?: boolean;
    error?: string;
  }> => ipcRenderer.invoke('fs:statFile', filePath),
  /** Бинарное чтение: большие таблицы не гоняются через base64. */
  readBinary: (
    filePath: string
  ): Promise<{
    success: boolean;
    bytes?: ArrayBuffer;
    size?: number;
    tooLarge?: boolean;
    error?: string;
  }> => ipcRenderer.invoke('fs:readBinary', filePath),
  /** Кусок файла — потоковый разбор CSV. */
  readBinarySlice: (
    filePath: string,
    offset = 0,
    length = 0
  ): Promise<{
    success: boolean;
    bytes?: ArrayBuffer;
    size?: number;
    offset?: number;
    total?: number;
    eof?: boolean;
    error?: string;
  }> => ipcRenderer.invoke('fs:readBinarySlice', filePath, offset, length),
  writeBinary: (
    filePath: string,
    bytes: ArrayBuffer | Uint8Array
  ): Promise<{ success: boolean; size?: number; error?: string }> =>
    ipcRenderer.invoke('fs:writeBinary', filePath, bytes),
  guessEncoding: (
    filePath: string
  ): Promise<{ success: boolean; encoding?: string; hasBom?: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:guessEncoding', filePath),
  /** Текст устаревшего .doc: разбор OLE2 в главном процессе. */
  extractDocText: (payload: { path?: string; base64?: string }): Promise<{
    success: boolean;
    text?: string;
    error?: string;
  }> => ipcRenderer.invoke('doc:extractText', payload),
  openPath: (filePath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('system:openPath', filePath),
  /** Реальный путь файла, перетащенного из Проводника. */
  getPathForFile: (file: File): string => webUtils.getPathForFile(file),
  writeFile: (
    filePath: string,
    content: string,
    encoding = 'utf-8'
  ): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:writeFile', filePath, content, encoding),
  createFile: (
    filePath: string,
    initialContent = '',
    encoding?: string
  ): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:createFile', filePath, initialContent, encoding),
  createDirectory: (dirPath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:createDirectory', dirPath),
  deletePath: (targetPath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:deletePath', targetPath),
  /** В корзину (обратимо), в отличие от deletePath. */
  trashPath: (targetPath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:trashPath', targetPath),
  /** Выбрать папку (без проверки права записи, в отличие от диалога сохранения). */
  pickDirectory: (defaultDir?: string): Promise<string | null> =>
    ipcRenderer.invoke('fs:pickDirectory', defaultDir),
  renamePath: (oldPath: string, newPath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:renamePath', oldPath, newPath),
  copyPath: (source: string, target: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('fs:copyPath', source, target),
  autosaveWrite: (name: string, payload: unknown): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('autosave:write', name, payload),
  autosaveList: (): Promise<{ success: boolean; items?: unknown[]; error?: string }> =>
    ipcRenderer.invoke('autosave:list'),
  autosaveRead: (name: string): Promise<{ success: boolean; snapshot?: unknown; error?: string }> =>
    ipcRenderer.invoke('autosave:read', name),
  autosaveRemove: (name: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('autosave:remove', name),
  diagnosticsLog: (level: string, message: string): Promise<{ success: boolean }> =>
    ipcRenderer.invoke('diagnostics:log', level, message),
  diagnosticsReport: (): Promise<{ success: boolean; report?: unknown; error?: string }> =>
    ipcRenderer.invoke('diagnostics:report'),
  windowNew: (): Promise<{ success: boolean; error?: string }> => ipcRenderer.invoke('window:new'),
  sessionGet: (): Promise<unknown> => ipcRenderer.invoke('session:get'),
  sessionPatch: (patch: unknown): Promise<unknown> => ipcRenderer.invoke('session:patch', patch),
  onSessionChanged: (cb: (session: unknown) => void): (() => void) => {
    const listener = (_event: unknown, session: unknown) => cb(session);
    ipcRenderer.on('session:changed', listener);
    return () => ipcRenderer.removeListener('session:changed', listener);
  },

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
    speed?: number;
    speedSegments?: { start: number; end: number; speed: number }[];
    duration?: number;
    hasAudio?: boolean;
    zoom?: { scale?: number; x?: number; y?: number };
  }): Promise<{ success: boolean; outputPath?: string; error?: string }> =>
    ipcRenderer.invoke('video:edit', req),
  /** Прогресс кодирования: проценты по `-progress` от ffmpeg. */
  onVideoProgress: (cb: (data: { filePath: string; percent: number }) => void): (() => void) => {
    const listener = (_e: unknown, data: { filePath: string; percent: number }) => cb(data);
    ipcRenderer.on('video:progress', listener);
    return () => ipcRenderer.removeListener('video:progress', listener);
  },
  /** Отмена активного кодирования файла. */
  videoCancel: (filePath: string): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('video:cancel', filePath),
  videoExtractAudio: (
    filePath: string,
    req?: { format?: 'mp3' | 'aac' | 'copy'; outputPath?: string }
  ): Promise<{ success: boolean; outputPath?: string; error?: string }> =>
    ipcRenderer.invoke('video:extractAudio', filePath, req),
  videoFrames: (
    filePath: string,
    req?: { count?: number; at?: number[]; maxWidth?: number; dir?: string }
  ): Promise<{ success: boolean; files?: string[]; error?: string }> =>
    ipcRenderer.invoke('video:frames', filePath, req),
  videoConvert: (
    filePath: string,
    req?: { format?: 'mp4' | 'webm' | 'gif'; quality?: 'low' | 'medium' | 'high'; outputPath?: string }
  ): Promise<{ success: boolean; outputPath?: string; error?: string }> =>
    ipcRenderer.invoke('video:convert', filePath, req),
  videoGif: (
    filePath: string,
    req?: { start?: number; duration?: number; fps?: number; width?: number; outputPath?: string }
  ): Promise<{ success: boolean; outputPath?: string; error?: string }> =>
    ipcRenderer.invoke('video:gif', filePath, req),
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
  /* ---------- Google: Диск и YouTube по отдельности ---------- */
  googleConnect: (
    service: 'drive' | 'youtube'
  ): Promise<{ success: boolean; email?: string; scopes?: string[]; error?: string }> =>
    ipcRenderer.invoke('google:connect', service),
  googleStatus: (): Promise<{
    configured: boolean;
    drive: { connected: boolean; email?: string; scopes?: string };
    youtube: { connected: boolean; email?: string; scopes?: string };
    email?: string;
  }> => ipcRenderer.invoke('google:status'),
  googleDisconnect: (service: 'drive' | 'youtube'): Promise<{ success: boolean }> =>
    ipcRenderer.invoke('google:disconnect', service),
  googleDiagnose: (): Promise<{
    configured: boolean;
    tokenValid: boolean;
    driveApi: 'ok' | 'no-access' | 'error' | 'not-connected';
    youtubeApi: 'ok' | 'no-access' | 'error' | 'not-connected';
    quotaBytes?: number;
    channelTitle?: string;
    error?: string;
  }> => ipcRenderer.invoke('google:diagnose'),
  onGdriveProgress: (cb: (data: { filePath: string; sent: number; total: number }) => void): (() => void) => {
    const listener = (_e: unknown, data: { filePath: string; sent: number; total: number }) =>
      cb(data);
    ipcRenderer.on('gdrive:progress', listener);
    return () => ipcRenderer.removeListener('gdrive:progress', listener);
  },
  openImgbb: (paths: string[]): Promise<{ success: boolean; count?: number; error?: string }> =>
    ipcRenderer.invoke('imagehost:open', 'imgbb', paths),
  openFreeimage: (paths: string[]): Promise<{ success: boolean; count?: number; error?: string }> =>
    ipcRenderer.invoke('imagehost:open', 'freeimage', paths),
  /** Япикс: загрузка в альбом, большие наборы — пачками. */
  openYapx: (
    paths: string[],
    options?: { authorized?: boolean; perAlbum?: number }
  ): Promise<{
    success: boolean;
    count?: number;
    albums?: string[];
    skipped?: { path: string; reason: string }[];
    error?: string;
  }> => ipcRenderer.invoke('imagehost:open', 'yapx', paths, options ?? {}),
  /** Лимиты и форматы сайтов-хостингов: их читает интерфейс для валидации. */
  imageHostLimits: (): Promise<
    Record<
      string,
      {
        id: string;
        title: string;
        formats: string[];
        maxFileMb: number;
        perAlbum: number;
        perAlbumAuthorized?: number;
        perHour?: number;
        resultPattern?: string;
      }
    >
  > => ipcRenderer.invoke('imagehost:hosts'),
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

  /* ---------- Обновления с GitHub ---------- */
  getAppVersion: (): Promise<string> => ipcRenderer.invoke('update:getVersion'),
  checkForUpdates: (): Promise<{
    success: boolean;
    supported?: boolean;
    currentVersion: string;
    latestVersion?: string;
    available?: boolean;
    assetName?: string;
    assetUrl?: string;
    assetDigest?: string;
    size?: number;
    publishedAt?: string;
    releaseNotes?: string;
    htmlUrl?: string;
    error?: string;
  }> => ipcRenderer.invoke('update:check'),
  downloadUpdate: (payload: {
    version: string;
    assetUrl: string;
    assetName: string;
    size?: number;
    digest?: string;
  }): Promise<{
    success: boolean;
    zipPath?: string;
    extractDir?: string;
    version?: string;
    error?: string;
  }> => ipcRenderer.invoke('update:download', payload),
  installUpdate: (payload: {
    zipPath: string;
    extractDir: string;
    version: string;
  }): Promise<{ success: boolean; autoRestart?: boolean; deferred?: boolean; error?: string }> =>
    ipcRenderer.invoke('update:install', payload),
  onUpdateProgress: (
    cb: (data: {
      phase: 'download' | 'extract' | 'install';
      received: number;
      total: number;
      percent: number;
    }) => void
  ): (() => void) => {
    const listener = (
      _e: unknown,
      data: {
        phase: 'download' | 'extract' | 'install';
        received: number;
        total: number;
        percent: number;
      }
    ) => cb(data);
    ipcRenderer.on('update:progress', listener);
    return () => ipcRenderer.removeListener('update:progress', listener);
  },

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

  /* ---------- Интеграция с Проводником Windows ---------- */
  getContextMenuStatus: (): Promise<{
    installed: boolean;
    exePath: string;
    labels?: Record<string, string>;
    command?: string;
    icon?: string;
    iconMatches?: boolean;
    error?: string;
  }> => ipcRenderer.invoke('system:getContextMenuStatus'),
  installContextMenu: (): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('system:installContextMenu'),
  removeContextMenu: (): Promise<{ success: boolean; error?: string }> =>
    ipcRenderer.invoke('system:removeContextMenu'),
  createDesktopShortcut: (): Promise<{ success: boolean; shortcutPath?: string; error?: string }> =>
    ipcRenderer.invoke('system:createDesktopShortcut'),
  getDesktopShortcutStatus: (): Promise<{ installed: boolean; shortcutPath: string }> =>
    ipcRenderer.invoke('system:getDesktopShortcutStatus'),
  getInitialExternalPath: (): Promise<{ path: string; isDirectory: boolean } | null> =>
    ipcRenderer.invoke('system:getInitialExternalPath'),
  onOpenExternalPath: (
    cb: (data: { path: string; isDirectory: boolean }) => void
  ): (() => void) => {
    const listener = (_e: unknown, data: { path: string; isDirectory: boolean }) => cb(data);
    ipcRenderer.on('app:openExternalPath', listener);
    return () => ipcRenderer.removeListener('app:openExternalPath', listener);
  },
};

contextBridge.exposeInMainWorld('electronAPI', electronAPI);
