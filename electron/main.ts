import { app, BrowserWindow, ipcMain, shell, dialog, clipboard, protocol } from 'electron';
import { execFileSync } from 'child_process';
import * as path from 'path';
import * as fs from 'fs';
import * as iconv from 'iconv-lite';
import { secretsImporter } from './secrets';
import { registerAiProxy } from './aiProxy';
import { registerShare } from './share';
import { registerGDrive } from './gdrive';
import { registerVideo } from './video';
import { registerYoutube } from './youtube';
import { registerHostedImageUpload } from './imageHost';
import { registerAiSettings } from './aiSettings';
import { setWatchRoot } from './watcher';
import { registerTerminal } from './terminal';

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'media',
    privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true, bypassCSP: true },
  },
]);

let mainWindow: BrowserWindow | null = null;

/**
 * Разрешение на закрытие окна. Устанавливается, когда пользователь
 * подтвердил выход (с сохранением или без) либо когда сохранений нет.
 */
let allowClose = false;

/** Признак того, что диалог подтверждения уже отправлен в рендерер. */
let closeRequestPending = false;

/** Текущая корневая папка рабочего пространства (выбирается пользователем). */
let workspaceRoot: string = '';

/**
 * Режим разработки определяется наличием URL dev-сервера, а не только
 * признаком упаковки. Иначе запуск собранного приложения распакованным
 * Electron (`npx electron dist-electron/main.js`, как в start-prod.bat)
 * пытался бы открыть localhost:5173 вместо файлов из dist/.
 */
const devServerUrl = process.env.VITE_DEV_SERVER_URL || '';
const isDev = !app.isPackaged && Boolean(devServerUrl);

/**
 * Chromium не может поднять sandbox, когда в Windows отключён UAC
 * (EnableLUA = 0): процесс получает полный токен администратора и падает
 * с STATUS_BREAKPOINT ещё до создания окна. В собранном приложении это
 * проверяется здесь, потому что рядом нет скрипта запуска, который добавил
 * бы --no-sandbox. При включённом UAC sandbox остаётся рабочим.
 */
function isChromiumSandboxUsable(): boolean {
  if (process.platform !== 'win32') return true;
  if (process.env.ELECTRON_DISABLE_SANDBOX) return false;
  try {
    const out = execFileSync(
      'reg',
      ['query', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System', '/v', 'EnableLUA'],
      { encoding: 'utf8', windowsHide: true }
    );
    const match = /EnableLUA\s+REG_DWORD\s+0x([0-9a-fA-F]+)/.exec(out);
    if (match) return parseInt(match[1], 16) !== 0;
  } catch {
    // Политику прочитать не удалось — не вмешиваемся в запуск.
  }
  return true;
}

if (app.isPackaged && !isChromiumSandboxUsable() && !process.argv.includes('--no-sandbox')) {
  const { spawn } = require('child_process') as typeof import('child_process');
  const child = spawn(process.execPath, ['--no-sandbox', ...process.argv.slice(1)], {
    detached: true,
    stdio: 'ignore',
    windowsHide: true,
  });
  child.unref();
  app.exit(0);
}

function createWindow() {
  // Новое окно начинает с чистого состояния перехвата закрытия
  allowClose = false;
  closeRequestPending = false;

  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1000,
    minHeight: 640,
    title: 'WorkSpace',
    backgroundColor: '#0b0f17',
    frame: false,
    titleBarStyle: 'hidden',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
      spellcheck: false,
    },
  });

  // Показываем окно только после готовности отрисовки — убирает белую вспышку
  mainWindow.once('ready-to-show', () => mainWindow?.show());

  if (isDev) {
    mainWindow.loadURL(devServerUrl);
  } else {
    const indexPath = path.join(__dirname, '..', 'dist', 'index.html');
    if (fs.existsSync(indexPath)) {
      mainWindow.loadFile(indexPath);
    } else {
      // Сборка отсутствует — показываем понятную подсказку вместо пустого окна
      mainWindow.loadURL(
        'data:text/html;charset=utf-8,' +
          encodeURIComponent(
            '<body style="background:#0b0f17;color:#e2e8f0;font-family:Segoe UI,sans-serif;' +
              'display:flex;align-items:center;justify-content:center;height:100vh;margin:0">' +
              '<div style="text-align:center;max-width:520px">' +
              '<h2 style="margin:0 0 12px">Сборка не найдена</h2>' +
              '<p style="color:#94a3b8;line-height:1.6">Файл <code>dist/index.html</code> отсутствует.<br>' +
              'Выполните <b>npm run build</b> или запустите <b>build.bat</b>, затем повторите запуск.</p>' +
              '</div></body>'
          )
      );
    }
  }

  // Сообщаем рендереру о смене состояния окна для корректной иконки maximize/restore
  const notifyMaximizeState = () => {
    mainWindow?.webContents.send('window:maximizeChanged', mainWindow.isMaximized());
  };
  mainWindow.on('maximize', notifyMaximizeState);
  mainWindow.on('unmaximize', notifyMaximizeState);

  /**
   * Закрытие окна перехватывается, чтобы рендерер успел спросить
   * пользователя о несохранённых файлах. Без этого изменения терялись.
   */
  mainWindow.on('close', (event) => {
    if (allowClose) return;

    // Рендерер не загружен — закрываем без диалога, иначе окно не закроется
    if (!mainWindow || mainWindow.webContents.isDestroyed()) {
      allowClose = true;
      return;
    }

    event.preventDefault();

    if (closeRequestPending) return; // диалог уже показан
    closeRequestPending = true;

    mainWindow.webContents.send('window:closeRequested');
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  registerVideo();
  registerTerminal(() => mainWindow);
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

/* ============================================================
   Управление окном
   ============================================================ */

ipcMain.handle('window:minimize', () => {
  mainWindow?.minimize();
});

ipcMain.handle('window:maximize', () => {
  if (mainWindow?.isMaximized()) {
    mainWindow.unmaximize();
    return false;
  }
  mainWindow?.maximize();
  return true;
});

ipcMain.handle('window:isMaximized', () => mainWindow?.isMaximized() ?? false);

ipcMain.handle('window:close', () => {
  // Запрос на закрытие от кнопки в заголовке — идёт через тот же диалог,
  // что и системная кнопка окна, чтобы не терять несохранённые изменения.
  mainWindow?.close();
});

/**
 * Ответ рендерера на запрос подтверждения закрытия.
 * proceed = true — закрыть окно, false — отменить выход.
 */
ipcMain.handle('window:confirmClose', (_event, proceed: boolean) => {
  closeRequestPending = false;
  if (!proceed) {
    allowClose = false;
    return { closed: false };
  }
  allowClose = true;
  mainWindow?.close();
  return { closed: true };
});

/* ============================================================
   Прокси запросов к AI-провайдерам
   Выполняются в главном процессе, т.к. браузерный fetch
   блокируется политикой CORS у части OpenAI-совместимых шлюзов.
   ============================================================ */
registerAiProxy();
registerAiSettings();

/* ============================================================
   Локальные секреты (secrets.local.ps1)
   ============================================================ */

/** Метаданные файла секретов — без самого ключа. */
ipcMain.handle('secrets:describe', () => {
  const info = secretsImporter.read(false);
  return {
    found: info.found,
    filePath: info.filePath,
    baseUrl: info.baseUrl,
    textModel: info.textModel,
    imageModel: info.imageModel,
    hasKey: Boolean(info.apiKey),
    hasOtherSecrets: info.hasOtherSecrets,
    error: info.error,
  };
});

/** Импорт ключа — вызывается только по действию пользователя. */
ipcMain.handle('secrets:importKey', () => {
  const info = secretsImporter.read(true);
  if (!info.found || !info.apiKey) {
    return { success: false, error: info.error || 'API-ключ не найден в secrets.local.ps1' };
  }
  return {
    success: true,
    apiKey: info.apiKey,
    baseUrl: info.baseUrl,
    model: info.textModel,
  };
});

/* ============================================================
   Рабочая папка и диалоги
   ============================================================ */

function sessionPath(): string {
  return path.join(app.getPath('userData'), 'workspace-session.json');
}

interface WorkspaceSession {
  lastPath?: string;
  favorites?: { path: string; name: string }[];
}

function readSession(): WorkspaceSession {
  try {
    const parsed = JSON.parse(fs.readFileSync(sessionPath(), 'utf-8')) as WorkspaceSession;
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

function writeSession(patch: Partial<WorkspaceSession>): WorkspaceSession {
  const next = { ...readSession(), ...patch };
  try {
    fs.writeFileSync(sessionPath(), JSON.stringify(next, null, 2), 'utf-8');
  } catch {
    // Каталог данных может быть недоступен — не срываем открытие папки
  }
  return next;
}

ipcMain.handle('fs:getCurrentWorkspace', () => {
  if (!workspaceRoot) {
    const saved = readSession().lastPath;
    workspaceRoot = saved && fs.existsSync(saved) ? saved : app.getPath('documents');
  }
  if (workspaceRoot) setWatchRoot(workspaceRoot, () => mainWindow);
  return workspaceRoot;
});

ipcMain.handle('fs:setWorkspaceRoot', (_event, dirPath: string) => {
  workspaceRoot = dirPath;
  if (dirPath) {
    writeSession({ lastPath: dirPath });
    setWatchRoot(dirPath, () => mainWindow);
  }
  return { success: true };
});

ipcMain.handle('fs:getFavorites', () => {
  return (readSession().favorites || []).filter((item) => item.path && fs.existsSync(item.path));
});

ipcMain.handle('fs:toggleFavorite', (_event, dirPath: string) => {
  const current = readSession().favorites || [];
  const exists = current.some((item) => item.path === dirPath);
  const favorites = exists
    ? current.filter((item) => item.path !== dirPath)
    : current.concat({ path: dirPath, name: path.basename(dirPath) || dirPath });
  writeSession({ favorites });
  return { success: true, favorites: favorites.filter((item) => fs.existsSync(item.path)) };
});

ipcMain.handle('fs:openFolderDialog', async () => {
  if (!mainWindow) return null;
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory', 'createDirectory'],
    title: 'Выберите рабочую папку WorkSpace',
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  workspaceRoot = result.filePaths[0];
  return workspaceRoot;
});

ipcMain.handle(
  'fs:openFileDialog',
  async (_event, filters?: { name: string; extensions: string[] }[]) => {
    if (!mainWindow) return null;
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openFile'],
      filters: filters || [
        { name: 'Все файлы', extensions: ['*'] },
        { name: 'Офисные документы', extensions: ['xlsx', 'xls', 'docx', 'pptx', 'csv'] },
        { name: 'Текстовые файлы', extensions: ['txt', 'md', 'log', 'ini', 'json', 'yaml', 'xml'] },
      ],
    });
    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0];
  }
);

ipcMain.handle(
  'fs:saveFileDialog',
  async (_event, defaultName: string, filters?: { name: string; extensions: string[] }[]) => {
    if (!mainWindow) return null;
    const result = await dialog.showSaveDialog(mainWindow, {
      defaultPath: defaultName,
      filters: filters || [{ name: 'Все файлы', extensions: ['*'] }],
    });
    if (result.canceled || !result.filePath) return null;
    return result.filePath;
  }
);

ipcMain.handle('fs:showInExplorer', (_event, filePath: string) => {
  try {
    shell.showItemInFolder(filePath);
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

/* ============================================================
   Файловые операции
   ============================================================ */

ipcMain.handle('fs:listDirectory', async (_event, dirPath: string) => {
  try {
    const entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
    const items = await Promise.all(
      entries.map(async (entry) => {
        const fullPath = path.join(dirPath, entry.name);
        let size = 0;
        let mtimeMs = 0;
        let birthMs = 0;
        try {
          const stat = await fs.promises.stat(fullPath);
          size = stat.size;
          mtimeMs = stat.mtimeMs;
          birthMs = stat.birthtimeMs;
        } catch {
          // Системные файлы могут быть недоступны для stat
        }
        return {
          name: entry.name,
          path: fullPath,
          isDirectory: entry.isDirectory(),
          size,
          mtime: mtimeMs ? new Date(mtimeMs).toISOString() : null,
          birthtime: birthMs ? new Date(birthMs).toISOString() : null,
          extension: entry.isDirectory() ? '' : path.extname(entry.name).toLowerCase().replace('.', ''),
        };
      })
    );

    // Свежие файлы сверху: сначала дата изменения, при равенстве — дата создания.
    items.sort((a, b) => {
      const modified = Date.parse(b.mtime || '') - Date.parse(a.mtime || '');
      if (modified) return modified;
      const created = Date.parse(b.birthtime || '') - Date.parse(a.birthtime || '');
      if (created) return created;
      return a.name.localeCompare(b.name, 'ru', { numeric: true, sensitivity: 'base' });
    });
    return { success: true, items };
  } catch (err) {
    return { success: false, error: (err as Error).message, items: [] };
  }
});

ipcMain.handle('fs:readFile', async (_event, filePath: string, encoding = 'utf-8') => {
  try {
    const buffer = await fs.promises.readFile(filePath);
    if (encoding === 'base64') {
      return { success: true, content: buffer.toString('base64'), isBinary: true };
    }
    if (encoding && encoding.toLowerCase() !== 'utf-8' && iconv.encodingExists(encoding)) {
      return { success: true, content: iconv.decode(buffer, encoding), isBinary: false };
    }
    // Определяем BOM UTF-16 — иначе файл прочитается как «кракозябры»
    if (buffer.length >= 2) {
      const b0 = buffer[0];
      const b1 = buffer[1];
      if (b0 === 0xff && b1 === 0xfe) {
        return { success: true, content: iconv.decode(buffer, 'utf-16le'), isBinary: false };
      }
      if (b0 === 0xfe && b1 === 0xff) {
        return { success: true, content: iconv.decode(buffer, 'utf-16be'), isBinary: false };
      }
    }
    return { success: true, content: buffer.toString('utf-8'), isBinary: false };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

ipcMain.handle(
  'fs:writeFile',
  async (_event, filePath: string, content: string, encoding = 'utf-8') => {
    try {
      // Запись через временный файл — защита от повреждения при сбое
      const tmpPath = `${filePath}.ws-tmp`;
      if (encoding === 'base64') {
        await fs.promises.writeFile(tmpPath, Buffer.from(content, 'base64'));
      } else if (encoding && encoding.toLowerCase() !== 'utf-8' && iconv.encodingExists(encoding)) {
        await fs.promises.writeFile(tmpPath, iconv.encode(content, encoding));
      } else {
        await fs.promises.writeFile(tmpPath, content, 'utf-8');
      }
      // Пока документ открыт в редакторе, Windows не даёт переименовать
      // временный файл поверх него (EPERM). Копирование заменяет содержимое
      // открытого файла, поэтому Ctrl+S действительно записывает правки.
      try {
        await fs.promises.rename(tmpPath, filePath);
      } catch (err) {
        const code = (err as NodeJS.ErrnoException).code;
        if (code !== 'EPERM' && code !== 'EEXIST' && code !== 'EBUSY') throw err;
        await fs.promises.copyFile(tmpPath, filePath);
        await fs.promises.unlink(tmpPath).catch(() => undefined);
      }
      return { success: true };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  }
);

ipcMain.handle('fs:createFile', async (_event, filePath: string, initialContent = '') => {
  try {
    await fs.promises.writeFile(filePath, initialContent, { flag: 'wx' });
    return { success: true };
  } catch (err) {
    const e = err as NodeJS.ErrnoException;
    if (e.code === 'EEXIST') {
      return { success: false, error: 'Файл с таким именем уже существует' };
    }
    return { success: false, error: e.message };
  }
});

ipcMain.handle('fs:createDirectory', async (_event, dirPath: string) => {
  try {
    await fs.promises.mkdir(dirPath, { recursive: true });
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

ipcMain.handle('fs:deletePath', async (_event, targetPath: string) => {
  try {
    const stat = await fs.promises.stat(targetPath);
    if (stat.isDirectory()) {
      await fs.promises.rm(targetPath, { recursive: true, force: true });
    } else {
      await fs.promises.unlink(targetPath);
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

ipcMain.handle('fs:renamePath', async (_event, oldPath: string, newPath: string) => {
  try {
    if (fs.existsSync(newPath)) {
      return { success: false, error: 'Объект с таким именем уже существует' };
    }
    await fs.promises.rename(oldPath, newPath);
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

/* ============================================================
   Поиск по рабочему пространству
   ============================================================ */

const IGNORED_DIRS = new Set([
  'node_modules',
  '.git',
  '.svn',
  '.hg',
  'dist',
  'dist-electron',
  'build',
  'out',
  '.next',
  '.cache',
  '.vscode',
  '.idea',
  '__pycache__',
  '.venv',
  'venv',
  'target',
]);

/** Расширения, содержимое которых имеет смысл искать как текст. */
const TEXT_EXTENSIONS = new Set([
  '.txt', '.md', '.markdown', '.log', '.json', '.yaml', '.yml', '.xml', '.ini', '.cfg', '.conf',
  '.env', '.properties', '.js', '.jsx', '.mjs', '.cjs', '.ts', '.tsx', '.html', '.htm', '.css',
  '.scss', '.less', '.py', '.sql', '.sh', '.bash', '.bat', '.cmd', '.ps1', '.csv', '.tsv', '.toml',
  '.c', '.cpp', '.h', '.hpp', '.cs', '.java', '.go', '.rs', '.php', '.rb', '.lua',
]);

const MAX_CONTENT_SEARCH_BYTES = 4 * 1024 * 1024;
const MAX_RESULTS = 300;

interface RawSearchResult {
  path: string;
  name: string;
  isDirectory: boolean;
  line?: number;
  matchSnippet?: string;
  totalMatches?: number;
}

/** Простейшая проверка glob-шаблона (`*.ts`, `src/*`, `*.{js,ts}`). */
function matchesGlob(fileName: string, filePath: string, glob: string): boolean {
  if (!glob || glob === '*') return true;
  const patterns = glob.includes('{')
    ? glob
        .replace(/^\{|\}$/g, '')
        .split(',')
        .map((p) => p.trim())
    : [glob];

  return patterns.some((pattern) => {
    const target = pattern.includes('/') || pattern.includes('\\') ? filePath.replace(/\\/g, '/') : fileName;
    const escaped = pattern
      .replace(/\\/g, '/')
      .replace(/[.+^${}()|[\]\\]/g, '\\$&')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.');
    return new RegExp(`^${escaped}$`, 'i').test(target) || new RegExp(escaped, 'i').test(target);
  });
}

async function walkSearch(
  dir: string,
  matcher: (text: string) => { index: number } | null,
  searchContent: boolean,
  includeGlob: string,
  results: RawSearchResult[]
): Promise<void> {
  if (results.length >= MAX_RESULTS) return;

  let entries: fs.Dirent[] = [];
  try {
    entries = await fs.promises.readdir(dir, { withFileTypes: true });
  } catch {
    return; // Нет прав доступа — пропускаем каталог
  }

  for (const entry of entries) {
    if (results.length >= MAX_RESULTS) break;
    const fullPath = path.join(dir, entry.name);

    if (entry.isDirectory()) {
      if (IGNORED_DIRS.has(entry.name) || entry.name.startsWith('.')) continue;
      await walkSearch(fullPath, matcher, searchContent, includeGlob, results);
      continue;
    }

    if (!matchesGlob(entry.name, fullPath, includeGlob)) continue;

    // Совпадение по имени файла
    const nameMatch = matcher(entry.name);
    if (nameMatch) {
      results.push({ path: fullPath, name: entry.name, isDirectory: false });
      continue;
    }

    if (!searchContent) continue;

    const ext = path.extname(entry.name).toLowerCase();
    if (!TEXT_EXTENSIONS.has(ext)) continue;

    try {
      const stat = await fs.promises.stat(fullPath);
      if (stat.size > MAX_CONTENT_SEARCH_BYTES) continue;

      const text = await fs.promises.readFile(fullPath, 'utf-8');
      const lines = text.split(/\r?\n/);
      let firstMatch: { line: number; snippet: string } | null = null;
      let totalMatches = 0;

      for (let i = 0; i < lines.length; i++) {
        if (matcher(lines[i])) {
          totalMatches++;
          if (!firstMatch) {
            firstMatch = { line: i + 1, snippet: lines[i].trim().slice(0, 160) };
          }
          if (totalMatches >= 50) break; // Достаточно для превью
        }
      }

      if (firstMatch) {
        results.push({
          path: fullPath,
          name: entry.name,
          isDirectory: false,
          line: firstMatch.line,
          matchSnippet: firstMatch.snippet,
          totalMatches,
        });
      }
    } catch {
      // Файл недоступен или не является UTF-8 текстом — пропускаем
    }
  }
}

ipcMain.handle(
  'fs:searchWorkspace',
  async (
    _event,
    rootDir: string,
    query: string,
    searchContent = false,
    options?: { caseSensitive?: boolean; useRegex?: boolean; includeGlob?: string }
  ) => {
    try {
      if (!query || !query.trim()) return { success: true, results: [] };
      if (!rootDir || !fs.existsSync(rootDir)) {
        return { success: false, error: 'Рабочая папка недоступна', results: [] };
      }

      const caseSensitive = Boolean(options?.caseSensitive);
      const useRegex = Boolean(options?.useRegex);
      const includeGlob = options?.includeGlob ?? '*';

      let regex: RegExp;
      try {
        regex = useRegex
          ? new RegExp(query, caseSensitive ? '' : 'i')
          : new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), caseSensitive ? '' : 'i');
      } catch (err) {
        return { success: false, error: `Некорректное регулярное выражение: ${(err as Error).message}`, results: [] };
      }

      const matcher = (text: string) => {
        regex.lastIndex = 0;
        const m = regex.exec(text);
        return m ? { index: m.index } : null;
      };

      const results: RawSearchResult[] = [];
      await walkSearch(rootDir, matcher, searchContent, includeGlob, results);
      return { success: true, results };
    } catch (err) {
      return { success: false, error: (err as Error).message, results: [] };
    }
  }
);

/* ============================================================
   Буфер обмена
   Чтение выполняется в главном процессе: рендерер получает HTML
   и текст без запроса прав у Chromium, а вставка таблиц работает
   даже когда фокус не в редактируемом поле.
   ============================================================ */

ipcMain.handle('clipboard:read', async () => {
  try {
    const items = await clipboard.read();
    let html = '';
    for (const item of items) {
      if (item.types.includes('text/html')) {
        const htmlBlob = await item.getType('text/html');
        html = await htmlBlob.text();
        break;
      }
    }
    return { success: true, text: await clipboard.readText(), html };
  } catch (err) {
    return { success: false, text: '', html: '', error: (err as Error).message };
  }
});

ipcMain.handle('clipboard:write', async (_event, text: string) => {
  try {
    await clipboard.writeText(String(text ?? ''));
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

/* ============================================================
   Экспорт в сервисы обмена и Google Drive
   ============================================================ */

registerShare();
registerGDrive();
registerYoutube();
registerHostedImageUpload();
