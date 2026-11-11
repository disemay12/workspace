/**
 * Экспорт изображений через формы сайтов-хостингов (ImgBB, Freeimage, Япикс).
 *
 * У сервисов нет публичного API загрузки, поэтому окно Electron открывает
 * страницу сайта и подставляет файлы в тот же `<input type="file">`, который
 * использовал бы человек. Дальше сайт сам загружает картинки и показывает
 * ссылки — приложение не притворяется API и не нарушает лимиты сервиса.
 *
 * Япикс (yapx.ru) собирает загрузку в альбом: лимит изображений на альбом
 * берётся со страницы, а большие наборы режутся на последовательные пачки,
 * каждая — в своём окне. Приложение возвращает ссылки на полученные альбомы.
 */
import { BrowserWindow, ipcMain, shell } from 'electron';
import * as fs from 'fs';
import * as path from 'path';

const MIME: Record<string, string> = {
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.png': 'image/png',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.bmp': 'image/bmp',
  '.apng': 'image/apng',
};

export interface ImageHostConfig {
  id: string;
  title: string;
  url: string;
  /** Селекторы файлового поля — от точного к общему. */
  selectors: string[];
  /** Разрешённые расширения без точки. */
  formats: string[];
  /** Максимальный размер файла в мегабайтах. */
  maxFileMb: number;
  /** Сколько изображений сайт принимает в один альбом. */
  perAlbum: number;
  /** Лимит изображений на альбом для авторизованного пользователя. */
  perAlbumAuthorized?: number;
  /** Лимит загрузок в час (аноним / с аккаунтом). */
  perHour?: number;
  perHourAuthorized?: number;
  /** Шаблон ссылки на результат: по нему ловим завершение загрузки. */
  resultPattern?: string;
  /** Сайт умеет принимать вставку из буфера — резервный способ инъекции. */
  supportsPaste?: boolean;
}

export const IMAGE_HOSTS: Record<string, ImageHostConfig> = {
  imgbb: {
    id: 'imgbb',
    title: 'ImgBB',
    url: 'https://ru.imgbb.com/',
    selectors: ['input[type="file"]:not([capture])', 'input[type="file"]'],
    formats: ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp'],
    maxFileMb: 32,
    perAlbum: 64,
  },
  freeimage: {
    id: 'freeimage',
    title: 'Freeimage.host',
    url: 'https://freeimage.host/ru',
    selectors: ['input[type="file"]:not([capture])', 'input[type="file"]'],
    formats: ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp'],
    maxFileMb: 16,
    perAlbum: 64,
  },
  yapx: {
    id: 'yapx',
    title: 'Япикс (yapx.ru)',
    url: 'https://yapx.ru/upload',
    selectors: [
      '#upload-input',
      '.js__main__upload-image input[type="file"]',
      '.drop-area input[type="file"]',
      'input[type="file"][multiple]',
      'input[type="file"]',
    ],
    formats: ['apng', 'bmp', 'gif', 'jpg', 'jpeg', 'png'],
    maxFileMb: 3,
    perAlbum: 10,
    perAlbumAuthorized: 20,
    perHour: 20,
    perHourAuthorized: 30,
    resultPattern: 'yapx\\.ru\\/album\\/',
    supportsPaste: true,
  },
};

export interface ImageHostOptions {
  /** Пользователь вошёл в аккаунт сайта — лимиты выше. */
  authorized?: boolean;
  /** Размер пачки, если нужно переопределить лимит сайта. */
  perAlbum?: number;
}

export interface ImageHostResult {
  success: boolean;
  count?: number;
  /** Ссылки на альбомы (по пачкам). */
  albums?: string[];
  /** Файлы, которые не подходят по формату/размеру. */
  skipped?: { path: string; reason: string }[];
  error?: string;
}

interface UploadPayload {
  name: string;
  type: string;
  data: string;
}

/** Проверка, что файл принимается сервисом. */
export function hostFileProblem(
  filePath: string,
  host: ImageHostConfig
): string | null {
  const ext = path.extname(filePath).toLowerCase().replace('.', '');
  if (!host.formats.includes(ext)) {
    return `${ext || 'без расширения'} — формат не поддерживается (${host.formats.join(', ')})`;
  }
  try {
    const stat = fs.statSync(filePath);
    if (stat.size > host.maxFileMb * 1024 * 1024) {
      return `${(stat.size / (1024 * 1024)).toFixed(1)} МБ — больше лимита ${host.maxFileMb} МБ`;
    }
  } catch (err) {
    return (err as Error).message;
  }
  return null;
}

export function batchFiles<T>(files: T[], perBatch: number): T[][] {
  const size = Math.max(1, Math.floor(perBatch));
  const batches: T[][] = [];
  for (let index = 0; index < files.length; index += size) {
    batches.push(files.slice(index, index + size));
  }
  return batches;
}

function buildInjectScript(payload: UploadPayload[], selectors: string[], allowPaste: boolean): string {
  return `(${injectedBody})(${JSON.stringify(payload)}, ${JSON.stringify(selectors)}, ${Boolean(
    allowPaste
  )});`;
}

/** Тело инъекции: собирает DataTransfer и дергает change, курсор — в резервной ветке paste. */
const injectedBody = `function (files, selectors, allowPaste) {
  if (window.__wsInjected === true) return 'skip';
  window.__wsInjected = true;

  const transfer = new DataTransfer();
  for (const item of files) {
    const binary = atob(item.data);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    transfer.items.add(new File([bytes], item.name, { type: item.type }));
  }

  let input = null;
  for (const selector of selectors) {
    const found = document.querySelector(selector);
    if (found) { input = found; break; }
  }

  if (input) {
    input.dataset.wsAttached = '1';
    try {
      input.files = transfer.files;
      input.dispatchEvent(new Event('change', { bubbles: true }));
      return 'input';
    } catch (error) {
      console.warn('workspace: input.files', error);
    }
  }

  if (allowPaste) {
    try {
      const target = document.querySelector('.drop-area') || document.body;
      const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
      Object.defineProperty(event, 'clipboardData', { value: transfer });
      target.dispatchEvent(event);
      return 'paste';
    } catch (error) {
      console.warn('workspace: paste', error);
    }
  }

  return 'failed';
}`;

/** Ожидание страницы результата: навигация или опрос адреса. */
function waitForResult(win: BrowserWindow, patternRegExp: RegExp, timeoutMs: number): Promise<string> {
  return new Promise((resolve) => {
    let settled = false;
    let poll: NodeJS.Timeout | undefined;
    let timer: NodeJS.Timeout | undefined;

    const finish = (url: string) => {
      if (settled) return;
      settled = true;
      if (timer) clearTimeout(timer);
      if (poll) clearInterval(poll);
      resolve(url);
    };

    const check = (url: string) => {
      if (patternRegExp.test(url)) finish(url);
    };

    win.webContents.on('did-navigate', (_event, url) => check(url));
    win.webContents.on('did-navigate-in-page', (_event, url) => check(url));
    win.webContents.on('did-finish-load', () => {
      void win.webContents
        .executeJavaScript('window.location.href')
        .then((url) => check(String(url)))
        .catch(() => undefined);
    });

    poll = setInterval(() => {
      void win.webContents
        .executeJavaScript('window.location.href')
        .then((url) => check(String(url)))
        .catch(() => undefined);
    }, 1200);

    timer = setTimeout(() => finish(''), timeoutMs);

    win.on('closed', () => finish(''));
  });
}

async function openBatch(
  host: ImageHostConfig,
  files: string[],
  options: { captureAlbum: boolean; authorized: boolean }
): Promise<{ album?: string; count: number }> {
  const payload: UploadPayload[] = files.map((file) => ({
    name: path.basename(file),
    type: MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
    data: fs.readFileSync(file).toString('base64'),
  }));

  const win = new BrowserWindow({
    width: 1180,
    height: 880,
    title: host.title,
    autoHideMenuBar: true,
    webPreferences: { sandbox: true, contextIsolation: true },
  });
  win.webContents.setWindowOpenHandler(({ url }) => {
    void shell.openExternal(url);
    return { action: 'deny' };
  });

  const script = buildInjectScript(payload, host.selectors, Boolean(host.supportsPaste));
  const attach = () => {
    void win.webContents.executeJavaScript(script).catch(() => undefined);
  };
  win.webContents.on('did-finish-load', attach);
  win.webContents.on('did-navigate-in-page', attach);

  const pattern = new RegExp(host.resultPattern ?? '^$');
  const albumPromise = options.captureAlbum ? waitForResult(win, pattern, 120_000) : Promise.resolve('');

  await win.loadURL(host.url);
  const album = await albumPromise;
  return { album: album || undefined, count: files.length };
}

export function registerImageHosts(): void {
  ipcMain.handle('imagehost:hosts', () => IMAGE_HOSTS);

  ipcMain.handle(
    'imagehost:open',
    async (_event, hostId: string, filePaths: string[], options: ImageHostOptions = {}) => {
      const host = IMAGE_HOSTS[hostId];
      const result: ImageHostResult = { success: false, skipped: [] };
      if (!host) return { ...result, error: 'Неизвестный сайт' };

      const existing = (filePaths || []).filter((file) => file && fs.existsSync(file));
      if (existing.length === 0) return { ...result, error: 'Не выбраны изображения' };

      const accepted: string[] = [];
      for (const file of existing) {
        const problem = hostFileProblem(file, host);
        if (problem) result.skipped!.push({ path: file, reason: problem });
        else accepted.push(file);
      }

      if (accepted.length === 0) {
        return {
          ...result,
          error: 'Ни один файл не подходит: ' + (result.skipped ?? []).map((item) => item.reason).join('; '),
        };
      }

      const limit = options.authorized
        ? host.perAlbumAuthorized ?? host.perAlbum
        : host.perAlbum;
      const perAlbum = Math.max(1, Math.min(options.perAlbum ?? limit, limit));
      const batches = batchFiles(accepted, perAlbum);
      const albums: string[] = [];
      let uploaded = 0;

      try {
        for (const batch of batches) {
          const { album, count } = await openBatch(host, batch, {
            // Альбом ловим только там, где у сайта есть страница результата:
            // иначе окно ImgBB/Freeimage ждало бы адрес, которого не будет.
            captureAlbum: Boolean(host.resultPattern),
            authorized: Boolean(options.authorized),
          });
          uploaded += count;
          if (album) albums.push(album);
        }
        return {
          success: uploaded > 0,
          count: uploaded,
          albums,
          skipped: result.skipped,
        };
      } catch (err) {
        return {
          ...result,
          success: uploaded > 0,
          count: uploaded,
          albums,
          error: (err as Error).message,
        };
      }
    }
  );
}
