import { app, BrowserWindow, ipcMain, shell, dialog, clipboard, protocol } from 'electron';
import { execFile, execFileSync } from 'child_process';
import * as path from 'path';
import * as fs from 'fs';
import * as iconv from 'iconv-lite';
import { secretsImporter } from './secrets';
import { registerAiProxy } from './aiProxy';
import { registerShare } from './share';
import { registerGDrive } from './gdrive';
import { registerVideo } from './video';
import { registerYoutube } from './youtube';
import { registerImageHosts } from './imageHost';
import { registerAiSettings } from './aiSettings';
import { docBufferToText } from './docText';
import { setWatchRoot } from './watcher';
import { registerTerminal } from './terminal';
import { registerContextMenuIpc } from './contextMenu';
import { registerUpdater, UPDATE_EXIT_CODE } from './updater';

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'media',
    privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true, bypassCSP: true },
  },
]);

/**
 * Извлекает путь к файлу или папке из аргументов командной строки.
 */
function extractTargetArg(args: string[]): { path: string; isDirectory: boolean } | null {
  for (let i = 1; i < args.length; i++) {
    const arg = args[i].trim();
    if (!arg || arg.startsWith('-')) continue;
    const lower = arg.toLowerCase();
    if (
      lower.endsWith('.js') ||
      lower.endsWith('.ts') ||
      lower.endsWith('electron.exe') ||
      lower.endsWith('workspace.exe')
    ) {
      continue;
    }
    try {
      if (fs.existsSync(arg)) {
        const fullPath = path.resolve(arg);
        const stat = fs.statSync(fullPath);
        return { path: fullPath, isDirectory: stat.isDirectory() };
      }
    } catch {}
  }
  return null;
}

let pendingStartupPath: { path: string; isDirectory: boolean } | null = extractTargetArg(process.argv);

// Единый экземпляр приложения: при повторном запуске (например, из контекстного меню)
// фокус передаётся в уже работающее окно, и открывается запрошенный файл/папка.
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
} else {
  app.on('second-instance', (_event, commandLine) => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();

      const target = extractTargetArg(commandLine);
      if (target) {
        mainWindow.webContents.send('app:openExternalPath', target);
      }
    }
  });
}

let mainWindow: BrowserWindow | null = null;

ipcMain.handle('system:getInitialExternalPath', () => {
  const p = pendingStartupPath;
  pendingStartupPath = null;
  return p;
});

/**
 * Разрешение на закрытие окна. Устанавливается, когда пользователь
 * подтвердил выход (с сохранением или без) либо когда сохранений нет.
 */
let allowClose = false;

/** Признак того, что диалог подтверждения уже отправлен в рендерер. */
let closeRequestPending = false;

/**
 * Код выхода для загрузчика: он применяет скачанное обновление и запускает
 * приложение заново. Обычный выход оставляет здесь null.
 */
let updateExitCode: number | null = null;

/** Текущая корневая папка рабочего пространства (выбирается пользователем). */
let workspaceRoot: string = '';

/**
 * Процесс с низким integrity level (Windows выдаёт его при запуске из
 * некоторых каталогов, в т.ч. при выключенном UAC) не может писать в Videos,
 * Documents и Desktop: ядро отвечает EACCES. Проверяем один раз.
 */
let lowIntegrityLevel: boolean | null = null;
function isLowIntegrity(): boolean {
  if (lowIntegrityLevel !== null) return lowIntegrityLevel;
  try {
    const out = execFileSync('whoami', ['/groups'], { encoding: 'utf8', windowsHide: true });
    lowIntegrityLevel = /S-1-16-4096/.test(out);
  } catch {
    lowIntegrityLevel = false;
  }
  return lowIntegrityLevel;
}

/** Реальная запись важнее метки токена: бывает, что low integrity не мешает. */
function canWriteDocuments(): boolean {
  try {
    const probe = path.join(app.getPath('documents'), `.workspace-write-${process.pid}.tmp`);
    fs.writeFileSync(probe, 'ok', 'utf8');
    fs.unlinkSync(probe);
    return true;
  } catch {
    return false;
  }
}

const LOW_INTEGRITY_HINT =
  'Windows ограничила права этого процесса (низкий уровень целостности) — так бывает при запуске из отдельных папок. ' +
  'Запустите WorkSpace из обычного каталога, например C:\\WorkSpace\\WorkSpace.exe.';

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

  const appIcon = path.join(__dirname, '..', 'dist', 'icon.png');

  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1000,
    minHeight: 640,
    title: 'WorkSpace',
    backgroundColor: '#0b0f17',
    icon: fs.existsSync(appIcon) ? appIcon : undefined,
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

  // Восстанавливаем закрепление окна поверх остальных
  if (readSession().alwaysOnTop) mainWindow.setAlwaysOnTop(true);

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

  // Передаём файл или папку, переданные в аргументах запуска
  mainWindow.webContents.on('did-finish-load', () => {
    if (pendingStartupPath) {
      mainWindow?.webContents.send('app:openExternalPath', pendingStartupPath);
      pendingStartupPath = null;
    }
  });

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
  registerContextMenuIpc();
  registerVideo();
  registerTerminal(() => mainWindow);
  registerUpdater({
    quitForUpdate: () => {
      // Загрузчик WorkSpace.exe ждёт завершения с этим кодом, копирует
      // подготовленное обновление и запускает приложение заново. Свои
      // процессы пережить выход не могут: Chromium держит их в job-объекте
      // с KILL_ON_JOB_CLOSE.
      allowClose = true;
      updateExitCode = UPDATE_EXIT_CODE;
      app.quit();
    },
  });
  // Файл секретов создаётся сразу: в релизе его нет, пользователь должен
  // увидеть готовый шаблон в профиле, а не «файл не найден».
  secretsImporter.ensure();
  createWindow();

  // Молчаливая потеря данных хуже предупреждения: при низком уровне
  // целостности запись в Videos/Documents/Desktop невозможна вообще.
  if (isLowIntegrity() && !canWriteDocuments()) {
    const options = {
      type: 'warning' as const,
      title: 'WorkSpace: ограниченные права',
      message: 'Windows ограничила права приложения',
      detail:
        LOW_INTEGRITY_HINT +
        '\n\nПока приложение запущено из этого каталога, файлы не сохраняются в Videos, Documents и Desktop.',
      buttons: ['Понятно'],
    };
    if (mainWindow) void dialog.showMessageBox(mainWindow, options);
    else void dialog.showMessageBox(options);
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

/**
 * Код выхода загрузчику возвращаем через process.exit: app.quit() завершает
 * процесс с нулём, и обновление не применилось бы.
 */
app.on('will-quit', () => {
  if (updateExitCode !== null) process.exit(updateExitCode);
});

/* ============================================================
   Управление окном
   ============================================================ */

/** Второе окно: отдельный рендерер для второго монитора, состояние своё. */
ipcMain.handle('window:new', async () => {
  try {
    const win = new BrowserWindow({
      width: 1280,
      height: 820,
      minWidth: 960,
      minHeight: 600,
      backgroundColor: '#0b0f17',
      show: false,
      webPreferences: {
        preload: path.join(__dirname, 'preload.js'),
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: false,
      },
    });
    win.once('ready-to-show', () => win.show());
    const devUrl = process.env.VITE_DEV_SERVER_URL;
    if (devUrl) await win.loadURL(devUrl);
    else await win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});
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

/** Закрепление окна поверх остальных — настройка сохраняется между запусками. */
ipcMain.handle('window:setAlwaysOnTop', (_event, value: boolean) => {
  const alwaysOnTop = Boolean(value);
  mainWindow?.setAlwaysOnTop(alwaysOnTop);
  writeSession({ alwaysOnTop });
  return { success: true, alwaysOnTop };
});

ipcMain.handle('window:getAlwaysOnTop', () => readSession().alwaysOnTop ?? false);

/* ============================================================
   Мост в настоящий Excel: точные формулы, сводные, диаграммы и VBA.
   Собственный движок WorkSpace не может гарантировать 100% совместимость,
   поэтому тяжёлые операции делегируются установленному Excel через COM.
   ============================================================ */

function runPowerShell(script: string): Promise<{ ok: boolean; output?: string; error?: string }> {
  return new Promise((resolve) => {
    const encoded = Buffer.from(script, 'utf16le').toString('base64');
    execFile(
      'powershell',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded],
      { windowsHide: true, timeout: 180000 },
      (error, stdout, stderr) => {
        if (error) {
          resolve({ ok: false, error: (stderr || error.message).trim().slice(0, 1200) });
          return;
        }
        resolve({ ok: true, output: String(stdout).trim() });
      }
    );
  });
}

/** Открывает файл в приложении, связанном с .xlsx/.xlsm (обычно Excel). */
ipcMain.handle('excel:openFile', async (_event, filePath: string) => {
  if (!filePath) return { success: false, error: 'Не указан путь к файлу' };
  if (!fs.existsSync(filePath)) return { success: false, error: 'Файл не найден' };
  const error = await shell.openPath(filePath);
  return error ? { success: false, error } : { success: true };
});

/** Выполняет VBA-макрос в настоящем Excel и сохраняет книгу. */
ipcMain.handle('excel:runMacro', async (_event, filePath: string, macroName: string) => {
  if (!filePath || !macroName?.trim()) {
    return { success: false, error: 'Укажите файл и имя макроса' };
  }
  if (!fs.existsSync(filePath)) return { success: false, error: 'Файл не найден' };

  const safePath = filePath.replace(/'/g, "''");
  const safeMacro = macroName.trim().replace(/'/g, "''");
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$excel = New-Object -ComObject Excel.Application",
    '$excel.Visible = $false',
    '$excel.DisplayAlerts = $false',
    'try {',
    `  $workbook = $excel.Workbooks.Open('${safePath}')`,
    `  $excel.Run('${safeMacro}')`,
    '  $workbook.Save()',
    '  $workbook.Close($false)',
    '} finally {',
    '  $excel.Quit()',
    '  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null',
    '}',
  ].join('\r\n');

  const result = await runPowerShell(script);
  return result.ok
    ? { success: true }
    : { success: false, error: result.error || 'Excel не выполнил макрос' };
});

interface ExcelChartSpec {
  sheet: string;
  type: 'bar' | 'line' | 'pie';
  range: string;
  title: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Переносит диаграммы WorkSpace в книгу как нативные объекты Excel. */
ipcMain.handle('excel:exportCharts', async (_event, filePath: string, charts: ExcelChartSpec[]) => {
  if (!filePath || !Array.isArray(charts) || charts.length === 0) {
    return { success: false, error: 'Нет диаграмм для экспорта' };
  }
  if (!fs.existsSync(filePath)) return { success: false, error: 'Файл не найден' };

  const payload = Buffer.from(JSON.stringify(charts), 'utf8').toString('base64');
  const safePath = filePath.replace(/'/g, "''");
  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${payload}'))`,
    '$charts = $json | ConvertFrom-Json',
    '$excel = New-Object -ComObject Excel.Application',
    '$excel.Visible = $false',
    '$excel.DisplayAlerts = $false',
    'try {',
    `  $workbook = $excel.Workbooks.Open('${safePath}')`,
    '  foreach ($chartSpec in $charts) {',
    '    $sheet = $workbook.Worksheets.Item($chartSpec.sheet)',
    '    $shape = $sheet.ChartObjects().Add(',
    '      [double]$chartSpec.x * 0.75, [double]$chartSpec.y * 0.75,',
    '      [double]$chartSpec.width * 0.75, [double]$chartSpec.height * 0.75)',
    '    $chart = $shape.Chart',
    '    $chart.SetSourceData($sheet.Range($chartSpec.range))',
    '    switch ($chartSpec.type) {',
    "      'bar' { $chart.ChartType = 51 }",
    "      'line' { $chart.ChartType = 4 }",
    "      'pie' { $chart.ChartType = 5 }",
    '    }',
    '    $chart.HasTitle = $true',
    '    $chart.ChartTitle.Text = $chartSpec.title',
    '  }',
    '  $workbook.Save()',
    '  $workbook.Close($false)',
    '} finally {',
    '  $excel.Quit()',
    '  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null',
    '}',
  ].join('\r\n');

  const result = await runPowerShell(script);
  return result.ok
    ? { success: true, count: charts.length }
    : { success: false, error: result.error || 'Excel не создал диаграммы' };
});

interface ExcelPivotSpec {
  sourceSheet: string;
  range: string;
  rowField: string;
  valueField: string;
  aggregation: 'sum' | 'count' | 'average' | 'min' | 'max';
  targetSheet: string;
}

/** Создаёт в книге настоящую обновляемую сводную таблицу Excel. */
ipcMain.handle('excel:createPivot', async (_event, filePath: string, spec: ExcelPivotSpec) => {
  if (!filePath || !spec?.sourceSheet || !spec?.range) {
    return { success: false, error: 'Не указан источник сводной' };
  }
  if (!fs.existsSync(filePath)) return { success: false, error: 'Файл не найден' };

  const payload = Buffer.from(JSON.stringify(spec), 'utf8').toString('base64');
  const safePath = filePath.replace(/'/g, "''");
  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${payload}'))`,
    '$spec = $json | ConvertFrom-Json',
    '$functions = @{ sum = -4157; count = -4112; average = -4106; min = -4139; max = -4136 }',
    '$excel = New-Object -ComObject Excel.Application',
    '$excel.Visible = $false',
    '$excel.DisplayAlerts = $false',
    'try {',
    `  $workbook = $excel.Workbooks.Open('${safePath}')`,
    '  $source = $workbook.Worksheets.Item($spec.sourceSheet)',
    '  $cache = $workbook.PivotCaches().Create(1, $source.Range($spec.range))',
    '  $target = $workbook.Worksheets.Add()',
    '  $target.Name = $spec.targetSheet',
    '  $pivot = $cache.CreatePivotTable($target.Range("A1"), "WorkSpacePivot_" + (Get-Date -Format "HHmmssfff"))',
    '  $pivot.PivotFields($spec.rowField).Orientation = 1',
    '  $caption = "$($spec.aggregation) — $($spec.valueField)"',
    '  $pivot.AddDataField($pivot.PivotFields($spec.valueField), $caption, $functions[$spec.aggregation]) | Out-Null',
    '  $workbook.Save()',
    '  $workbook.Close($false)',
    '} finally {',
    '  $excel.Quit()',
    '  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null',
    '}',
  ].join('\r\n');

  const result = await runPowerShell(script);
  return result.ok
    ? { success: true }
    : { success: false, error: result.error || 'Excel не создал сводную таблицу' };
});

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
  /** Окно закреплено поверх остальных приложений. */
  alwaysOnTop?: boolean;
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

ipcMain.handle('session:get', () => readSession());

/** Сохраняет раскладку и рассылает её остальным окнам. */
ipcMain.handle('session:patch', (_event, patch: Partial<WorkspaceSession>) => {
  const next = writeSession(patch);
  for (const window of BrowserWindow.getAllWindows()) {
    if (!window.isDestroyed()) window.webContents.send('session:changed', next);
  }
  return next;
});
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
  async (
    _event,
    defaultName: string,
    filters?: { name: string; extensions: string[] }[],
    defaultDir?: string
  ) => {
    if (!mainWindow) return null;
    // Голое имя файла Windows трактует как «текущая папка процесса» и диалог
    // открывается в случайном месте (часто в Videos) — передаём полный путь.
    const dir = defaultDir && fs.existsSync(defaultDir) ? defaultDir : undefined;
    const result = await dialog.showSaveDialog(mainWindow, {
      defaultPath: dir ? path.join(dir, defaultName) : defaultName,
      filters: filters || [{ name: 'Все файлы', extensions: ['*'] }],
    });
    if (result.canceled || !result.filePath) return null;
    return result.filePath;
  }
);

ipcMain.handle('fs:pathExists', async (_event, targetPath: string) => {
  try {
    await fs.promises.access(targetPath);
    return true;
  } catch {
    return false;
  }
});

/** Свободный путь в папке: name.mp4 → name (2).mp4 → name (3).mp4 … */
ipcMain.handle('fs:uniquePath', async (_event, dirPath: string, fileName: string) => {
  const ext = path.extname(fileName);
  const base = path.basename(fileName, ext);
  let candidate = path.join(dirPath, fileName);
  for (let i = 2; fs.existsSync(candidate) && i < 1000; i += 1) {
    candidate = path.join(dirPath, `${base} (${i})${ext}`);
  }
  return candidate;
});

ipcMain.handle('fs:showInExplorer', async (_event, filePath: string) => {
  try {
    const stat = await fs.promises.stat(filePath);
    // Окно WorkSpace может быть закреплено поверх всех: Explorer откроется,
    // но останется за ним. Снимаем закрепление до возврата фокуса.
    const wasAlwaysOnTop = mainWindow?.isAlwaysOnTop() ?? false;
    if (wasAlwaysOnTop) mainWindow?.setAlwaysOnTop(false);
    if (wasAlwaysOnTop && mainWindow) {
      const restore = () => {
        mainWindow?.setAlwaysOnTop(true);
        mainWindow?.removeListener('focus', restore);
      };
      mainWindow.once('focus', restore);
    }

    if (stat.isDirectory()) {
      const error = await shell.openPath(filePath);
      return error ? { success: false, error } : { success: true };
    }

    // Штатный API Electron: открывает папку и выделяет файл.
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

/**
 * Текст устаревшего .doc (Word 97-2003). Разбор идёт в главном процессе:
 * формат требует Buffer и разбора OLE2-контейнера.
 */
ipcMain.handle(
  'doc:extractText',
  async (_event, payload: { path?: string; base64?: string } = {}) => {
    try {
      const buffer = payload.base64
        ? Buffer.from(payload.base64, 'base64')
        : await fs.promises.readFile(payload.path ?? '');
      return { success: true, text: docBufferToText(buffer) };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  }
);

ipcMain.handle('fs:readFile', async (_event, filePath: string, encoding = 'utf-8') => {  try {
    const buffer = await fs.promises.readFile(filePath);
    if (encoding === 'base64') {
      return { success: true, content: buffer.toString('base64'), isBinary: true };
    }
    if (encoding && encoding.toLowerCase() !== 'utf-8' && encoding.toLowerCase() !== 'utf-8-bom' && iconv.encodingExists(encoding)) {
      return { success: true, content: iconv.decode(buffer, encoding).replace(/^\uFEFF/, ''), isBinary: false };
    }
    // Определяем BOM UTF-16 — иначе файл прочитается как «кракозябры»
    if (buffer.length >= 2) {
      const b0 = buffer[0];
      const b1 = buffer[1];
      if (b0 === 0xff && b1 === 0xfe) {
        return { success: true, content: iconv.decode(buffer, 'utf-16le').replace(/^\uFEFF/, ''), isBinary: false };
      }
      if (b0 === 0xfe && b1 === 0xff) {
        return { success: true, content: iconv.decode(buffer, 'utf-16be').replace(/^\uFEFF/, ''), isBinary: false };
      }
    }
    // UTF-8 BOM Excel ставит сам при экспорте CSV. В содержимом редактора
    // он не нужен: иначе первая ячейка начинается с невидимого символа.
    return { success: true, content: buffer.toString('utf-8').replace(/^\uFEFF/, ''), isBinary: false };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

/**
 * Читает первые байты файла для бинарного просмотрщика.
 * Полный файл не загружается: EXE и архивы могут быть сотнями мегабайт.
 */
ipcMain.handle(
  'fs:readPreview',
  async (_event, filePath: string, maxBytes = 65536) => {
    try {
      const stat = await fs.promises.stat(filePath);
      const size = Math.min(Math.max(1024, maxBytes), stat.size);
      const handle = await fs.promises.open(filePath, 'r');
      try {
        const buffer = Buffer.alloc(size);
        const { bytesRead } = await handle.read(buffer, 0, size, 0);
        return {
          success: true,
          base64: buffer.subarray(0, bytesRead).toString('base64'),
          size: stat.size,
        };
      } finally {
        await handle.close();
      }
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  }
);

/** Открывает любой файл средствами Windows (EXE запускается, PDF открывается в просмотрщике). */
ipcMain.handle('system:openPath', async (_event, filePath: string) => {
  if (!filePath) return { success: false, error: 'Не указан путь к файлу' };
  if (!fs.existsSync(filePath)) return { success: false, error: 'Файл не найден' };
  const error = await shell.openPath(filePath);
  return error ? { success: false, error } : { success: true };
});

/**
 * Атомарная запись: временный файл рядом с целевым и переименование.
 * Если переименование не удалось (антивирус, занятый файл), пробуем копирование
 * и повторную прямую запись — но только для текстовых файлов: недописанный ZIP
 * Excel/Word считают повреждённым.
 */
async function writeFileAtomic(
  filePath: string,
  payload: Buffer | string,
  binaryFile: boolean
): Promise<void> {
  const writeWithRetry = async (target: string): Promise<void> => {
    let lastError: unknown;
    for (let attempt = 0; attempt < 4; attempt++) {
      try {
        await fs.promises.writeFile(target, payload);
        return;
      } catch (err) {
        lastError = err;
        const code = (err as NodeJS.ErrnoException).code;
        if (code !== 'EPERM' && code !== 'EBUSY' && code !== 'EACCES') throw err;
        await new Promise((resolve) => setTimeout(resolve, 200 + attempt * 200));
      }
    }
    throw lastError;
  };

  const lockedError = (code?: string) =>
    new Error(
      code === 'EACCES' && isLowIntegrity()
        ? 'Не удалось сохранить файл: ' + LOW_INTEGRITY_HINT
        : code === 'EACCES'
          ? 'Нет прав на запись в эту папку. Выберите другую папку или запустите WorkSpace от администратора.'
          : 'Файл открыт в другой программе (Excel/Word). Закройте его и повторите сохранение.'
    );

  // Уникальное имя: два параллельных сохранения не должны открывать
  // один и тот же .ws-tmp (иначе Windows возвращает EPERM).
  const tmpPath = `${filePath}.${process.pid}.${Date.now()}.ws-tmp`;
  try {
    await fs.promises.writeFile(tmpPath, payload);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'EPERM' || code === 'EBUSY' || code === 'EACCES') {
      // Временный файл создать нельзя (антивирус/блокировка).
      if (binaryFile) throw lockedError(code);
      await writeWithRetry(filePath);
      return;
    }
    throw err;
  }

  try {
    await fs.promises.rename(tmpPath, filePath);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code !== 'EPERM' && code !== 'EEXIST' && code !== 'EBUSY' && code !== 'EACCES') throw err;
    try {
      await fs.promises.copyFile(tmpPath, filePath);
    } catch (copyErr) {
      const copyCode = (copyErr as NodeJS.ErrnoException).code;
      if (copyCode !== 'EPERM' && copyCode !== 'EBUSY' && copyCode !== 'EACCES') {
        throw copyErr;
      }
      if (binaryFile) throw lockedError(copyCode);
      await writeWithRetry(filePath);
    }
    await fs.promises.unlink(tmpPath).catch(() => undefined);
  }
}

ipcMain.handle(
  'fs:writeFile',
  async (_event, filePath: string, content: string, encoding = 'utf-8') => {
    try {
      const payload =
        encoding === 'base64'
          ? Buffer.from(content, 'base64')
          : encoding === 'utf-8-bom'
            ? Buffer.from(`\uFEFF${content}`, 'utf8')
            : encoding && encoding.toLowerCase() !== 'utf-8' && iconv.encodingExists(encoding)
              ? iconv.encode(content, encoding)
              : content;

      // Бинарные офисные файлы нельзя писать поверх открытого файла:
      // Excel/Word видят недописанный ZIP и сообщают «недопустимый формат».
      const binaryFile = /\.(xlsx|xlsm|xlsb|docx|doc|pptx|ppt|zip|7z|pdf|png|jpe?g)$/i.test(
        filePath
      );
      await writeFileAtomic(filePath, payload, binaryFile);
      return { success: true };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  }
);

/** Предельный размер файла для бинарного чтения в рендерер (воркер разбора таблиц). */
const MAX_BINARY_BYTES = 512 * 1024 * 1024;

/** Приводит то, что пришло через IPC, к Buffer. */
function toBuffer(data: unknown): Buffer {
  if (Buffer.isBuffer(data)) return data;
  if (data instanceof Uint8Array) return Buffer.from(data.buffer, data.byteOffset, data.byteLength);
  if (data instanceof ArrayBuffer) return Buffer.from(new Uint8Array(data));
  if (Array.isArray(data)) return Buffer.from(data as number[]);
  throw new Error('Ожидались байты файла');
}

/** ArrayBuffer ровно по границам буфера: лишние байты пула не клонируются. */
function toArrayBuffer(buffer: Buffer | Uint8Array): ArrayBuffer {
  return buffer.buffer.slice(
    buffer.byteOffset,
    buffer.byteOffset + buffer.byteLength
  ) as ArrayBuffer;
}

ipcMain.handle('fs:statFile', async (_event, filePath: string) => {
  try {
    const stat = await fs.promises.stat(filePath);
    return {
      success: true,
      size: stat.size,
      mtimeMs: stat.mtimeMs,
      isDirectory: stat.isDirectory(),
    };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

/**
 * Бинарное чтение файла для больших офисных документов и воркера разбора
 * таблиц: base64 для файла в 100 МБ — это лишние +33 % и копия в JS-строку.
 */
ipcMain.handle('fs:readBinary', async (_event, filePath: string) => {
  try {
    const stat = await fs.promises.stat(filePath);
    if (stat.size > MAX_BINARY_BYTES) {
      return {
        success: false,
        tooLarge: true,
        size: stat.size,
        error:
          'Файл больше ' +
          Math.round(MAX_BINARY_BYTES / (1024 * 1024)) +
          ' МБ — откройте его в системном приложении.',
      };
    }
    const buffer = await fs.promises.readFile(filePath);
    return { success: true, bytes: toArrayBuffer(buffer), size: buffer.byteLength };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

/** Кусок файла: CSV разбирается потоково, без чтения всего файла в память рендерера. */
ipcMain.handle(
  'fs:readBinarySlice',
  async (_event, filePath: string, offset = 0, length = 0) => {
    try {
      const stat = await fs.promises.stat(filePath);
      const start = Math.max(0, Math.min(Number(offset) || 0, stat.size));
      const size = Math.max(0, Math.min(Number(length) || 0, stat.size - start));
      const handle = await fs.promises.open(filePath, 'r');
      try {
        const buffer = Buffer.alloc(size);
        const { bytesRead } = await handle.read(buffer, 0, size, start);
        const slice = buffer.subarray(0, bytesRead);
        return {
          success: true,
          bytes: toArrayBuffer(slice),
          size: bytesRead,
          offset: start,
          total: stat.size,
          eof: start + bytesRead >= stat.size,
        };
      } finally {
        await handle.close().catch(() => undefined);
      }
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  }
);

/** Запись байтов с той же атомарностью, что и текстовая запись. */
ipcMain.handle('fs:writeBinary', async (_event, filePath: string, data: unknown) => {
  try {
    const buffer = toBuffer(data);
    await writeFileAtomic(filePath, buffer, true);
    return { success: true, size: buffer.byteLength };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

/**
 * Определение кодировки по первым 64 КБ: BOM, UTF-16 без BOM,
 * валидность UTF-8, иначе windows-1251 (русская ANSI-ветка этого приложения).
 */
ipcMain.handle('fs:guessEncoding', async (_event, filePath: string) => {
  try {
    const handle = await fs.promises.open(filePath, 'r');
    let sample: Buffer;
    try {
      const buffer = Buffer.alloc(65536);
      const { bytesRead } = await handle.read(buffer, 0, buffer.length, 0);
      sample = buffer.subarray(0, bytesRead);
    } finally {
      await handle.close().catch(() => undefined);
    }

    if (sample.length >= 3 && sample[0] === 0xef && sample[1] === 0xbb && sample[2] === 0xbf) {
      return { success: true, encoding: 'utf-8', hasBom: true };
    }
    if (sample.length >= 2 && sample[0] === 0xff && sample[1] === 0xfe) {
      return { success: true, encoding: 'utf-16le', hasBom: true };
    }
    if (sample.length >= 2 && sample[0] === 0xfe && sample[1] === 0xff) {
      return { success: true, encoding: 'utf-16be', hasBom: true };
    }

    // UTF-16 без BOM: половина байтов нулевые.
    if (sample.length >= 4) {
      let evenZeros = 0;
      let oddZeros = 0;
      for (let i = 0; i < Math.min(sample.length, 4096); i++) {
        if (sample[i] !== 0) continue;
        if (i % 2 === 0) evenZeros++;
        else oddZeros++;
      }
      const checked = Math.floor(Math.min(sample.length, 4096) / 2);
      if (checked > 0 && oddZeros / checked > 0.35) {
        return { success: true, encoding: 'utf-16le', hasBom: false };
      }
      if (checked > 0 && evenZeros / checked > 0.35) {
        return { success: true, encoding: 'utf-16be', hasBom: false };
      }
    }

    let validUtf8 = true;
    try {
      new TextDecoder('utf-8', { fatal: true }).decode(sample);
    } catch {
      validUtf8 = false;
    }
    if (validUtf8) return { success: true, encoding: 'utf-8', hasBom: false };

    // Старшие байты в русских текстах — кириллица cp1251, иначе западная ANSI.
    let highBytes = 0;
    for (const byte of sample) if (byte >= 0xc0) highBytes++;
    const encoding = highBytes / Math.max(1, sample.length) > 0.02 ? 'windows-1251' : 'windows-1252';
    return { success: true, encoding, hasBom: false };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

/** PNG по слайдам: каждая страница htmlPages рендерится в окно размера слайда. */
ipcMain.handle(
  'ppt:exportImages',
  async (_event, htmlPages: string[], outDir: string, width = 1280, height = 720) => {
    if (!Array.isArray(htmlPages) || htmlPages.length === 0) {
      return { success: false, error: 'Нет слайдов для экспорта' };
    }
    const dir =
      outDir ||
      path.join(app.getPath('downloads'), `WorkSpace-slides-${new Date().toISOString().slice(0, 10)}`);
    let win: BrowserWindow | null = null;
    const files: string[] = [];
    try {
      fs.mkdirSync(dir, { recursive: true });
      win = new BrowserWindow({
        show: false,
        width: Math.max(320, Math.min(3840, Math.round(width))),
        height: Math.max(240, Math.min(2160, Math.round(height))),
      });
      for (let index = 0; index < htmlPages.length; index += 1) {
        await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(htmlPages[index])}`);
        await new Promise((resolve) => setTimeout(resolve, 120));
        const image = await win.webContents.capturePage();
        const file = path.join(dir, `slide-${String(index + 1).padStart(2, '0')}.png`);
        fs.writeFileSync(file, image.toPNG());
        files.push(file);
      }
      return { success: true, files, dir };
    } catch (err) {
      return { success: false, files, error: (err as Error).message };
    } finally {
      if (win) win.destroy();
    }
  }
);

ipcMain.handle(
  'pdf:export',
  async (
    _event,
    html: string,
    suggestedName = 'document.pdf',
    targetPath?: string,
    pageSize?: { width: number; height: number }
  ) => {
    let win: BrowserWindow | null = null;
    try {
      win = new BrowserWindow({ show: false });
      await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
      // Размер страницы в микронах: презентация печатается своим форматом,
      // иначе слайд 16:9 попадёт на A4 с полями.
      const hasCustomSize = Boolean(pageSize && pageSize.width > 0 && pageSize.height > 0);
      const pdf = await win.webContents.printToPDF({
        printBackground: true,
        ...(hasCustomSize && pageSize
          ? {
              pageSize: { width: Math.round(pageSize.width), height: Math.round(pageSize.height) },
              margins: { top: 0, bottom: 0, left: 0, right: 0 },
              preferCSSPageSize: true,
            }
          : { pageSize: 'A4' as const }),
      });
      let output = targetPath || '';
      if (!output) {
        const result = await dialog.showSaveDialog({
          title: 'Экспорт в PDF',
          defaultPath: suggestedName,
          filters: [{ name: 'PDF', extensions: ['pdf'] }],
        });
        if (result.canceled || !result.filePath) return { success: false, canceled: true };
        output = result.filePath;
      }
      await fs.promises.writeFile(output, pdf);
      return { success: true, path: output };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    } finally {
      win?.destroy();
    }
  }
);

ipcMain.handle(
  'fs:createFile',
  async (_event, filePath: string, initialContent = '', encoding?: string) => {
    try {
      const isBase64 =
        encoding === 'base64' ||
        (typeof initialContent === 'string' &&
          initialContent.length > 0 &&
          /\.(xlsx|xls|docx|doc|pptx|png|jpg|ico)$/i.test(filePath) &&
          /^[A-Za-z0-9+/=\r\n]+$/.test(initialContent));

      const payload = isBase64
        ? Buffer.from(initialContent.replace(/\s+/g, ''), 'base64')
        : encoding === 'utf-8-bom'
          ? Buffer.from(`\uFEFF${initialContent}`, 'utf8')
          : initialContent;

      await fs.promises.mkdir(path.dirname(filePath), { recursive: true });
      try {
        await fs.promises.writeFile(filePath, payload, { flag: 'wx' });
      } catch (err) {
        const e = err as NodeJS.ErrnoException;
        // Антивирус/git на мгновение держат каталог: одна повторная попытка
        // без эксклюзивного флага и с паузой снимает EPERM/EACCES.
        if (e.code === 'EPERM' || e.code === 'EACCES' || e.code === 'EBUSY') {
          await new Promise((resolve) => setTimeout(resolve, 150));
          await fs.promises.writeFile(filePath, payload);
        } else {
          throw err;
        }
      }
      return { success: true };
    } catch (err) {
      const e = err as NodeJS.ErrnoException;
      if (e.code === 'EEXIST') {
        return { success: false, error: 'Файл с таким именем уже существует' };
      }
      return { success: false, error: e.message };
    }
  }
);

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

ipcMain.handle('fs:trashPath', async (_event, targetPath: string) => {
  try {
    await shell.trashItem(targetPath);
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

/** Выбор папки (в отличие от диалога сохранения не проверяет права записи). */
ipcMain.handle('fs:pickDirectory', async (_event, defaultDir?: string) => {
  if (!mainWindow) return null;
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory', 'createDirectory'],
    title: 'Куда сохранить копию',
    defaultPath: defaultDir && fs.existsSync(defaultDir) ? defaultDir : undefined,
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
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
   Автосохранение: снимки несохранённых документов
   ============================================================ */

const autosaveDir = () => path.join(app.getPath('userData'), 'autosave');

function autosaveFile(name: string): string {
  const safe = String(name || '').replace(/[^a-z0-9_-]/gi, '').slice(0, 64);
  if (!safe) throw new Error('Некорректное имя снимка');
  return path.join(autosaveDir(), `${safe}.json`);
}

/* ============================================================
   Диагностика: кольцевой буфер ошибок и отчёт о системе
   ============================================================ */

const diagLog: { at: number; level: string; message: string }[] = [];
const DIAG_LIMIT = 200;

function pushDiag(level: string, message: string): void {
  diagLog.push({ at: Date.now(), level, message: String(message).slice(0, 2000) });
  if (diagLog.length > DIAG_LIMIT) diagLog.shift();
}

ipcMain.handle('diagnostics:log', (_event, level: string, message: string) => {
  pushDiag(level || 'error', message);
  return { success: true };
});

ipcMain.handle('diagnostics:report', async () => {
  try {
    const autosaveDirPath = autosaveDir();
    let snapshots = 0;
    try {
      snapshots = (await fs.promises.readdir(autosaveDirPath)).filter((name) => name.endsWith('.json')).length;
    } catch {
      snapshots = 0;
    }
    return {
      success: true,
      report: {
        app: {
          version: app.getVersion(),
          name: app.getName(),
          packaged: app.isPackaged,
          locale: app.getLocale(),
        },
        system: {
          platform: process.platform,
          arch: process.arch,
          release: process.getSystemVersion?.() ?? '',
          versions: {
            electron: process.versions.electron,
            chrome: process.versions.chrome,
            node: process.versions.node,
          },
          paths: { userData: app.getPath('userData'), home: app.getPath('home'), appData: app.getPath('appData') },
          memoryMb: Math.round(process.memoryUsage().rss / (1024 * 1024)),
        },
        autosave: { dir: autosaveDirPath, snapshots },
        logs: [...diagLog],
      },
    };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});
ipcMain.handle('autosave:write', async (_event, name: string, payload: unknown) => {
  try {
    fs.mkdirSync(autosaveDir(), { recursive: true });
    const target = autosaveFile(name);
    const temp = `${target}.tmp`;
    await fs.promises.writeFile(temp, JSON.stringify(payload), 'utf8');
    await fs.promises.rename(temp, target);
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

ipcMain.handle('autosave:list', async () => {
  try {
    fs.mkdirSync(autosaveDir(), { recursive: true });
    const files = await fs.promises.readdir(autosaveDir());
    const items: unknown[] = [];
    for (const file of files) {
      if (!file.endsWith('.json')) continue;
      try {
        const raw = await fs.promises.readFile(path.join(autosaveDir(), file), 'utf8');
        items.push(JSON.parse(raw));
      } catch {
        // битый снимок пропускаем
      }
    }
    return { success: true, items };
  } catch (err) {
    return { success: false, items: [], error: (err as Error).message };
  }
});

ipcMain.handle('autosave:read', async (_event, name: string) => {
  try {
    const raw = await fs.promises.readFile(autosaveFile(name), 'utf8');
    return { success: true, snapshot: JSON.parse(raw) };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});

ipcMain.handle('autosave:remove', async (_event, name: string) => {
  try {
    await fs.promises.rm(autosaveFile(name), { force: true });
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
});
ipcMain.handle('fs:copyPath', async (_event, source: string, target: string) => {
  try {
    if (fs.existsSync(target)) {
      return { success: false, error: 'Объект с таким именем уже существует' };
    }
    await fs.promises.cp(source, target, { recursive: true, errorOnExist: true, force: false });
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
registerImageHosts();
