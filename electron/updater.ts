/**
 * Обновление portable-сборки WorkSpace с GitHub Releases.
 *
 * Установка заменяет только папку программы (launcher + app). Профиль
 * пользователя %APPDATA%\WorkSpace с API-ключами, сессией и настройками
 * не трогается; перед обновлением он дополнительно копируется в backup.
 */
import { app, BrowserWindow, ipcMain, net } from 'electron';
import { execFile } from 'child_process';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';

const REPO = 'disemay12/workspace';
const USER_AGENT = 'WorkSpace-Updater';

export interface UpdateInfo {
  success: boolean;
  supported: boolean;
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
}

interface ReleaseAsset {
  name: string;
  browser_download_url: string;
  size?: number;
  digest?: string;
}

/** -1 / 0 / 1 без зависимостей. */
export function compareVersions(a: string, b: string): number {
  const parse = (value: string) =>
    value
      .replace(/^v/i, '')
      .split(/[.+-]/)
      .map((part) => Number.parseInt(part, 10) || 0);
  const left = parse(a);
  const right = parse(b);
  const length = Math.max(left.length, right.length);
  for (let index = 0; index < length; index++) {
    const diff = (left[index] ?? 0) - (right[index] ?? 0);
    if (diff !== 0) return diff > 0 ? 1 : -1;
  }
  return 0;
}

function isAllowedUrl(value: string): boolean {
  try {
    const url = new URL(value);
    if (url.protocol !== 'https:') return false;
    return (
      url.hostname === 'github.com' ||
      url.hostname === 'api.github.com' ||
      url.hostname.endsWith('.githubusercontent.com')
    );
  } catch {
    return false;
  }
}

function pickAsset(assets: ReleaseAsset[]): ReleaseAsset | null {
  return (
    assets.find((asset) => /^WorkSpace-.*-win-x64\.zip$/i.test(asset.name)) ??
    assets.find((asset) => /\.zip$/i.test(asset.name)) ??
    null
  );
}

async function checkViaApi(): Promise<UpdateInfo> {
  const response = await net.fetch(`https://api.github.com/repos/${REPO}/releases/latest`, {
    headers: {
      Accept: 'application/vnd.github+json',
      'User-Agent': USER_AGENT,
      'X-GitHub-Api-Version': '2022-11-28',
    },
  });
  if (!response.ok) throw new Error(`GitHub API: HTTP ${response.status}`);
  const data = (await response.json()) as {
    tag_name?: string;
    name?: string;
    body?: string;
    html_url?: string;
    published_at?: string;
    assets?: ReleaseAsset[];
  };
  const latestVersion = String(data.tag_name || '').replace(/^v/i, '');
  const asset = pickAsset(data.assets ?? []);
  if (!latestVersion || !asset) throw new Error('В релизе GitHub нет win-x64 zip');
  return {
    success: true,
    supported: true,
    currentVersion: app.getVersion(),
    latestVersion,
    available: compareVersions(latestVersion, app.getVersion()) > 0,
    assetName: asset.name,
    assetUrl: asset.browser_download_url,
    assetDigest: asset.digest,
    size: asset.size,
    publishedAt: data.published_at,
    releaseNotes: data.body || data.name || '',
    htmlUrl: data.html_url,
  };
}

/** Резервный путь, если GitHub API недоступен или исчерпан лимит. */
async function checkViaReleasePage(): Promise<UpdateInfo> {
  const response = await net.fetch(`https://github.com/${REPO}/releases/latest`, {
    headers: { 'User-Agent': USER_AGENT },
  });
  if (!response.ok) throw new Error(`GitHub: HTTP ${response.status}`);
  const match = /\/tag\/v?([^/?#]+)/i.exec(response.url);
  const latestVersion = match?.[1] ?? '';
  if (!latestVersion) throw new Error('Не удалось определить версию релиза');
  const assetName = `WorkSpace-${latestVersion}-win-x64.zip`;
  return {
    success: true,
    supported: true,
    currentVersion: app.getVersion(),
    latestVersion,
    available: compareVersions(latestVersion, app.getVersion()) > 0,
    assetName,
    assetUrl: `https://github.com/${REPO}/releases/download/v${latestVersion}/${assetName}`,
    htmlUrl: `https://github.com/${REPO}/releases/tag/v${latestVersion}`,
  };
}

async function checkForUpdates(): Promise<UpdateInfo> {
  const currentVersion = app.getVersion();
  if (!app.isPackaged) {
    return {
      success: true,
      supported: false,
      currentVersion,
      error: 'Проверка обновлений доступна только в собранной portable-версии WorkSpace',
    };
  }
  try {
    return await checkViaApi();
  } catch (apiError) {
    try {
      return await checkViaReleasePage();
    } catch (pageError) {
      return {
        success: false,
        supported: true,
        currentVersion,
        error: `${(apiError as Error).message}; ${(pageError as Error).message}`,
      };
    }
  }
}

function runFile(command: string, args: string[], timeout = 10 * 60 * 1000): Promise<void> {
  return new Promise((resolve, reject) => {
    execFile(command, args, { windowsHide: true, timeout }, (error) => {
      if (error) reject(error);
      else resolve();
    });
  });
}

async function extractZip(zipPath: string, destination: string): Promise<void> {
  await fs.promises.rm(destination, { recursive: true, force: true });
  await fs.promises.mkdir(destination, { recursive: true });
  try {
    await runFile('tar', ['-xf', zipPath, '-C', destination]);
  } catch {
    const escaped = destination.replace(/'/g, "''");
    const escapedZip = zipPath.replace(/'/g, "''");
    await runFile('powershell', [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-Command',
      `Expand-Archive -LiteralPath '${escapedZip}' -DestinationPath '${escaped}' -Force`,
    ]);
  }
}

/** Папка, содержимое которой копируется поверх установленной программы. */
async function findSourceRoot(extracted: string): Promise<string> {
  const candidates = [path.join(extracted, 'WorkSpace'), extracted];
  for (const candidate of candidates) {
    const launcher = path.join(candidate, 'WorkSpace.exe');
    const appExe = path.join(candidate, 'app', 'WorkSpace.exe');
    if (fs.existsSync(launcher) && fs.existsSync(appExe)) return candidate;
  }
  throw new Error('В архиве обновления нет WorkSpace.exe и app\\WorkSpace.exe');
}

function sendProgress(phase: 'download' | 'extract' | 'install', received: number, total: number) {
  const percent = total > 0 ? Math.max(0, Math.min(100, (received / total) * 100)) : 0;
  for (const win of BrowserWindow.getAllWindows()) {
    if (!win.isDestroyed()) {
      win.webContents.send('update:progress', { phase, received, total, percent });
    }
  }
}

async function downloadUpdate(payload: {
  version: string;
  assetUrl: string;
  assetName: string;
  size?: number;
  digest?: string;
}): Promise<{ success: boolean; zipPath?: string; extractDir?: string; version?: string; error?: string }> {
  try {
    if (!isAllowedUrl(payload.assetUrl)) {
      throw new Error('Ссылка на обновление должна вести на GitHub');
    }
    const baseName = payload.assetName.replace(/[^a-zA-Z0-9._-]/g, '_');
    const root = path.join(app.getPath('temp'), `WorkSpace-update-${payload.version}`);
    await fs.promises.rm(root, { recursive: true, force: true });
    await fs.promises.mkdir(root, { recursive: true });
    const zipPath = path.join(root, baseName);

    const response = await net.fetch(payload.assetUrl, {
      headers: { 'User-Agent': USER_AGENT },
    });
    if (!response.ok || !response.body) {
      throw new Error(`Не удалось скачать обновление: HTTP ${response.status}`);
    }
    const total = Number(response.headers.get('content-length')) || payload.size || 0;
    const reader = response.body.getReader();
    const stream = fs.createWriteStream(zipPath);
    const hash = crypto.createHash('sha256');
    let received = 0;
    let lastSent = 0;

    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        if (value) {
          received += value.byteLength;
          hash.update(value);
          await new Promise<void>((resolve, reject) => {
            stream.write(Buffer.from(value), (error) => (error ? reject(error) : resolve()));
          });
          const now = Date.now();
          if (now - lastSent > 250 || (total > 0 && received >= total)) {
            lastSent = now;
            sendProgress('download', received, total);
          }
        }
      }
    } finally {
      await new Promise<void>((resolve) => stream.end(() => resolve()));
    }

    if (payload.digest && payload.digest.startsWith('sha256:')) {
      const actual = `sha256:${hash.digest('hex')}`;
      if (actual.toLowerCase() !== payload.digest.toLowerCase()) {
        await fs.promises.rm(root, { recursive: true, force: true });
        throw new Error('SHA-256 скачанного файла не совпадает с GitHub');
      }
    }

    sendProgress('extract', 0, 1);
    const extracted = path.join(root, 'extracted');
    await extractZip(zipPath, extracted);
    const sourceRoot = await findSourceRoot(extracted);
    sendProgress('extract', 1, 1);
    return { success: true, zipPath, extractDir: sourceRoot, version: payload.version };
  } catch (error) {
    return { success: false, error: (error as Error).message };
  }
}

/**
 * Код выхода, которым приложение просит загрузчик применить обновление.
 * Загрузчик WorkSpace.exe ждёт завершения приложения, копирует файлы и
 * запускает новую версию — сам он в job-объекте Chromium не состоит.
 */
export const UPDATE_EXIT_CODE = 42;

/** Маркер подготовленного обновления; его читает загрузчик. */
function pendingUpdatePath(): string {
  return path.join(app.getPath('userData'), 'pending-update.txt');
}

/** Новый загрузчик пишет launcher.json при каждом старте. */
function launcherSupportsRestart(): boolean {
  try {
    const raw = fs.readFileSync(path.join(app.getPath('userData'), 'launcher.json'), 'utf8');
    return /"autoRestart"\s*:\s*true/.test(raw);
  } catch {
    return false;
  }
}

export interface InstallResult {
  success: boolean;
  /** Обновление применит загрузчик и перезапустит приложение сам. */
  autoRestart?: boolean;
  /** Нужен ручной перезапуск: загрузчик старой версии не умеет применять маркер. */
  deferred?: boolean;
  error?: string;
}

/**
 * Готовит обновление к установке.
 *
 * Electron держит resources/app.asar открытым, поэтому заменить его из
 * работающего приложения нельзя (Windows: «файл занят другим процессом»).
 * Приложение пишет маркер и завершается, а файлы копирует загрузчик —
 * до старта приложения и после выхода с кодом UPDATE_EXIT_CODE.
 *
 * Процессы, запущенные приложением, живут в job-объекте Chromium с
 * KILL_ON_JOB_CLOSE и не могут пережить выход — поэтому внешний скрипт
 * установки здесь не используется.
 */
function installUpdate(payload: {
  zipPath: string;
  extractDir: string;
  version: string;
}): InstallResult {
  if (!app.isPackaged) {
    return { success: false, error: 'Обновление доступно только в собранной версии' };
  }

  const appDir = path.dirname(process.execPath);
  const target = path.dirname(appDir);
  if (
    path.basename(appDir).toLowerCase() !== 'app' ||
    !fs.existsSync(path.join(target, 'WorkSpace.exe')) ||
    !fs.existsSync(path.join(payload.extractDir, 'WorkSpace.exe')) ||
    !fs.existsSync(path.join(payload.extractDir, 'app', 'WorkSpace.exe'))
  ) {
    return {
      success: false,
      error: 'Не удалось определить папку portable-сборки. Распакуйте обновление вручную.',
    };
  }

  // Проверяем право записи до выхода: иначе приложение закроется и не вернётся.
  try {
    const probe = path.join(target, `.workspace-update-${process.pid}.tmp`);
    fs.writeFileSync(probe, 'ok');
    fs.unlinkSync(probe);
  } catch (error) {
    return {
      success: false,
      error: `Нет прав на запись в папку программы: ${(error as Error).message}`,
    };
  }

  // Загрузчик новой версии ждёт завершения приложения, поэтому его файл занят.
  // Тогда новый загрузчик кладём рядом: на старте он переименует старый образ
  // и встанет на его место.
  let launcherStaged = false;
  try {
    fs.copyFileSync(path.join(payload.extractDir, 'WorkSpace.exe'), path.join(target, 'WorkSpace.exe'));
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code !== 'EBUSY' && code !== 'EPERM' && code !== 'EACCES') {
      return {
        success: false,
        error: `Не удалось обновить загрузчик WorkSpace.exe: ${(error as Error).message}`,
      };
    }
    try {
      fs.copyFileSync(
        path.join(payload.extractDir, 'WorkSpace.exe'),
        path.join(target, 'WorkSpace.exe.new')
      );
      launcherStaged = true;
    } catch (stageError) {
      return {
        success: false,
        error: `Не удалось подготовить загрузчик: ${(stageError as Error).message}`,
      };
    }
  }
  void launcherStaged;

  // Среду Electron заменить на ходу нельзя — кладём рядом, загрузчик поставит
  // её на место перед следующим запуском.
  try {
    const newRuntime = path.join(payload.extractDir, 'app', 'WorkSpace.exe');
    const currentRuntime = path.join(appDir, 'WorkSpace.exe');
    if (fs.statSync(newRuntime).size !== fs.statSync(currentRuntime).size) {
      fs.copyFileSync(newRuntime, `${currentRuntime}.new`);
    }
  } catch {
    // Среда Electron обычно не меняется между версиями приложения.
  }

  const marker = pendingUpdatePath();
  try {
    fs.mkdirSync(path.dirname(marker), { recursive: true });
    fs.writeFileSync(marker, `source=${payload.extractDir}\nversion=${payload.version}\n`, 'utf8');
  } catch (error) {
    return { success: false, error: `Не удалось записать маркер обновления: ${(error as Error).message}` };
  }

  const autoRestart = launcherSupportsRestart();
  return { success: true, autoRestart, deferred: !autoRestart };
}

export function registerUpdater(options: { quitForUpdate: () => void }): void {
  ipcMain.handle('update:getVersion', () => app.getVersion());
  ipcMain.handle('update:check', () => checkForUpdates());
  ipcMain.handle('update:download', (_event, payload) => downloadUpdate(payload));
  ipcMain.handle('update:install', async (_event, payload) => {
    const result = installUpdate(payload);
    // Закрываемся только тогда, когда загрузчик умеет применить обновление и
    // поднять приложение: иначе окно просто исчезнет без обновления.
    if (result.success && result.autoRestart) {
      setTimeout(() => options.quitForUpdate(), 1200);
    }
    return result;
  });
}
