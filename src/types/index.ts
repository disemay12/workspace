export type FileType =
  | 'text'        // txt, log, ini, cfg, env, bat, sh
  | 'code'        // js, ts, py, html, css, json, yaml, xml, sql
  | 'markdown'    // md, markdown
  | 'richtext'    // note, rtf
  | 'excel'       // xlsx, xls, csv, tsv
  | 'word'        // docx, doc
  | 'powerpoint'  // pptx, ppt
  | 'image'       // png, jpg, jpeg, gif, svg, webp, ico
  | 'pdf'         // pdf
  | 'video'       // mp4, webm, mkv, mov и другие контейнеры
  | 'binary';     // прочие

export interface FileItem {
  name: string;
  path: string;
  isDirectory: boolean;
  size: number;
  mtime: string | null;
  birthtime?: string | null;
  extension: string;
  type?: FileType;
  children?: FileItem[];
  isExpanded?: boolean;
}

export interface OpenTab {
  id: string;
  path: string;
  name: string;
  extension: string;
  type: FileType;
  /** Текстовое содержимое либо base64 для бинарных офисных форматов. */
  content: string;
  /** Снимок содержимого на момент последнего сохранения — база для проверки dirty. */
  initialContent: string;
  isDirty: boolean;
  /** 'utf-8' | 'windows-1251' | ... | 'base64' */
  encoding: string;
  viewMode?: 'visual' | 'source' | 'split';
  activeSheet?: string;
  activeSlide?: number;
}

/* ============================================================
   AI
   ============================================================ */

export type AIProvider =
  | 'openai'
  | 'anthropic'
  | 'gemini'
  | 'deepseek'
  | 'openrouter'
  | 'ollama'
  | 'starimg'
  | 'custom';

export interface ProviderPreset {
  id: AIProvider;
  label: string;
  baseUrl: string;
  models: string[];
  keyHint: string;
  docsUrl?: string;
  requiresKey: boolean;
}

export interface AISettings {
  provider: AIProvider;
  /** Ключ активного провайдера. В localStorage не пишется — хранится в keyVault. */
  apiKey: string;
  baseUrl: string;
  model: string;
  temperature: number;
  maxTokens: number;
  systemPrompt: string;
  activePresetId?: string;
  modelsCache?: Record<string, string[]>;
}

export interface ChatAttachment {
  name: string;
  path: string;
  content: string;
  /** Обрезано ли содержимое при прикреплении. */
  truncated?: boolean;
}

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: number;
  attachedFile?: ChatAttachment;
  /** Ошибка генерации для отображения в ленте. */
  error?: string;
  /** Признак незавершённого стрима. */
  streaming?: boolean;
}

export interface ChatSession {
  id: string;
  title: string;
  messages: ChatMessage[];
  createdAt: number;
  updatedAt: number;
}

/* ============================================================
   Поиск
   ============================================================ */

export interface SearchResult {
  path: string;
  name: string;
  isDirectory: boolean;
  line?: number;
  matchSnippet?: string;
  /** Все совпадения в файле (для панели поиска). */
  totalMatches?: number;
}

export interface SearchOptions {
  query: string;
  searchContent: boolean;
  caseSensitive: boolean;
  useRegex: boolean;
  includeGlob: string;
}

/* ============================================================
   Инлайн-копилот (Ctrl+K)
   ============================================================ */

export interface InlineEditRequest {
  tabId: string;
  /** Исходный фрагмент, который надо изменить. */
  original: string;
  /** Инструкция пользователя. */
  instruction: string;
  /** Диапазон выделения в документе. */
  range?: { from: number; to: number };
}

export type InlineEditStage = 'prompt' | 'loading' | 'diff' | 'error';

/* ============================================================
   Результат анализа документа
   ============================================================ */

export interface DocumentStats {
  lines: number;
  chars: number;
  words: number;
  encoding: string;
}

declare global {
  interface Window {
    electronAPI?: {
      minimizeWindow: () => Promise<void>;
      maximizeWindow: () => Promise<boolean>;
      isWindowMaximized: () => Promise<boolean>;
      closeWindow: () => Promise<void>;
      /** Ответ на запрос подтверждения закрытия: true — закрыть приложение. */
      confirmClose: (proceed: boolean) => Promise<{ closed: boolean }>;
      /** Подписка на запрос подтверждения закрытия окна. */
      onCloseRequested?: (callback: () => void) => () => void;
      onWindowMaximizeChange?: (callback: (isMaximized: boolean) => void) => () => void;

      getCurrentWorkspace: () => Promise<string>;
      openFolderDialog: () => Promise<string | null>;
      openFileDialog: (filters?: { name: string; extensions: string[] }[]) => Promise<string | null>;
      saveFileDialog: (defaultName: string, filters?: { name: string; extensions: string[] }[]) => Promise<string | null>;
      showInExplorer: (filePath: string) => Promise<{ success: boolean; error?: string }>;
      setWorkspaceRoot: (dirPath: string) => Promise<{ success: boolean }>;
      getFavorites: () => Promise<{ path: string; name: string }[]>;
      toggleFavorite: (dirPath: string) => Promise<{ success: boolean; favorites: { path: string; name: string }[] }>;

      listDirectory: (dirPath: string) => Promise<{ success: boolean; items: FileItem[]; error?: string }>;
      readFile: (filePath: string, encoding?: string) => Promise<{ success: boolean; content?: string; isBinary?: boolean; error?: string }>;
      writeFile: (filePath: string, content: string, encoding?: string) => Promise<{ success: boolean; error?: string }>;
      createFile: (filePath: string, initialContent?: string) => Promise<{ success: boolean; error?: string }>;
      createDirectory: (dirPath: string) => Promise<{ success: boolean; error?: string }>;
      deletePath: (targetPath: string) => Promise<{ success: boolean; error?: string }>;
      renamePath: (oldPath: string, newPath: string) => Promise<{ success: boolean; error?: string }>;

      searchWorkspace: (
        rootDir: string,
        query: string,
        searchContent?: boolean,
        options?: { caseSensitive?: boolean; useRegex?: boolean; includeGlob?: string }
      ) => Promise<{ success: boolean; results: SearchResult[]; error?: string }>;

      readClipboard: () => Promise<{ success: boolean; text: string; html: string; error?: string }>;
      writeClipboard: (text: string) => Promise<{ success: boolean; error?: string }>;
      uploadToService: (req: { service: string; fileName: string; mime: string; base64: string; clientId?: string }) => Promise<{ success: boolean; url?: string; error?: string }>;
      uploadAlbum: (req: {
        service: string;
        title?: string;
        clientId?: string;
        images: { fileName: string; mime: string; base64: string; title?: string }[];
      }) => Promise<{ success: boolean; url?: string; uploaded: number; failed: string[]; links: string[]; error?: string }>;
      gdriveGetConfig: () => Promise<{ clientId: string; clientSecret: string }>;
      gdriveSetConfig: (cfg: { clientId: string; clientSecret: string }) => Promise<{ success: boolean }>;
      gdriveStatus: () => Promise<{ configured: boolean; connected: boolean; email?: string }>;
      gdriveConnect: () => Promise<{ success: boolean; email?: string; error?: string }>;
      gdriveDisconnect: () => Promise<{ success: boolean }>;
      gdriveUpload: (paths: string[], folderName?: string) => Promise<{ success: boolean; uploaded: number; failed: string[]; link?: string; error?: string }>;
      videoCapabilities: () => Promise<{ ffmpeg: string | null; ffprobe: string | null }>;
      videoProbe: (filePath: string) => Promise<{ success: boolean; duration?: number; width?: number; height?: number; hasAudio?: boolean; videoCodec?: string; audioCodec?: string; size?: number; mime?: string; error?: string }>;
      videoEdit: (req: { filePath: string; start?: number; end?: number; stripAudio?: boolean; compress?: boolean; outputPath?: string }) => Promise<{ success: boolean; outputPath?: string; error?: string }>;
      youtubeCategories: () => Promise<{ success: boolean; categories: { id: string; title: string }[]; error?: string }>;
      youtubePlaylists: () => Promise<{ success: boolean; playlists: { id: string; title: string }[]; error?: string }>;
      youtubeUpload: (filePath: string, meta: { title: string; description?: string; tags?: string[]; categoryId?: string; privacyStatus: 'public' | 'unlisted' | 'private'; madeForKids: boolean; embeddable?: boolean; license?: 'youtube' | 'creativeCommon'; defaultLanguage?: string; publishAt?: string; playlistId?: string }) => Promise<{ success: boolean; id?: string; url?: string; error?: string }>;
      openImgbb: (paths: string[]) => Promise<{ success: boolean; count?: number; error?: string }>;
      openFreeimage: (paths: string[]) => Promise<{ success: boolean; count?: number; error?: string }>;
      onYoutubeProgress: (cb: (data: { filePath: string; sent: number; total: number }) => void) => () => void;

      /** Постоянные настройки AI из профиля пользователя. */
      loadAiSettings?: () => Promise<{
        settings: Partial<AISettings>;
        keys: Record<string, string>;
      }>;
      saveAiSettings?: (patch: Partial<AISettings> & { apiKey?: string }) => Promise<{ success: boolean }>;

      /** Метаданные файла secrets.local.ps1 — без самого ключа. */
      describeSecrets: () => Promise<{
        found: boolean;
        filePath: string;
        baseUrl?: string;
        textModel?: string;
        imageModel?: string;
        hasKey: boolean;
        hasOtherSecrets: boolean;
        error?: string;
      }>;
      /** Импорт ключа из локального файла секретов. */
      importSecretsKey: () => Promise<{
        success: boolean;
        apiKey?: string;
        baseUrl?: string;
        model?: string;
        error?: string;
      }>;

      /* ---------- Прокси запросов к AI ---------- */
      /**
       * Стриминговый запрос к провайдеру в главном процессе.
       * Обходит CORS, которому подчиняется браузерный fetch.
       */
      aiStream: (req: {
        requestId: string;
        url: string;
        method: string;
        headers: Record<string, string>;
        body: string;
      }) => Promise<{ started: boolean }>;
      aiAbort: (requestId: string) => Promise<{ aborted: boolean }>;
      aiProbe: (
        url: string
      ) => Promise<{ reachable: boolean; status?: number; error?: string }>;
      aiListModels?: (req: {
        provider: string;
        baseUrl?: string;
        apiKey?: string;
      }) => Promise<{ success: boolean; models: string[]; error?: string }>;
      onAiChunk: (cb: (data: { requestId: string; chunk: string }) => void) => () => void;
      onAiEnd: (cb: (data: { requestId: string }) => void) => () => void;
      onAiError: (
        cb: (data: {
          requestId: string;
          message: string;
          status?: number;
          isNetworkError?: boolean;
        }) => void
      ) => () => void;

      /* ---------- Отслеживание изменений на диске ---------- */
      onWorkspaceChanged?: (
        cb: (data: { eventType: string; filename: string; fullPath: string }) => void
      ) => () => void;

      /* ---------- Встроенный терминал ---------- */
      terminalCreate?: (req?: {
        id?: string;
        cwd?: string;
        shell?: string;
      }) => Promise<{ success: boolean; id?: string; error?: string }>;
      terminalInput?: (req: { id: string; data: string }) => Promise<{ success: boolean; error?: string }>;
      terminalKill?: (id: string) => Promise<{ success: boolean }>;
      onTerminalData?: (cb: (data: { id: string; data: string }) => void) => () => void;
      onTerminalExit?: (cb: (data: { id: string; code: number }) => void) => () => void;
    };
  }
}
