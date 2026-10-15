/**
 * Интеграция с Google Drive: OAuth 2.0 (loopback-редирект) и выгрузка
 * файлов и папок целиком.
 *
 * Токены и параметры клиента хранятся в userData/gdrive.json и никогда не
 * покидают главный процесс — рендерер получает только статус и ссылки.
 */
import { app, ipcMain, shell } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import * as http from 'http';

const SCOPES = [
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/userinfo.email',
  'https://www.googleapis.com/auth/youtube.upload',
  'https://www.googleapis.com/auth/youtube.readonly',
].join(' ');
const LOOPBACK_PORT = 42813;
const REDIRECT_URI = 'http://127.0.0.1:' + LOOPBACK_PORT;

interface GDriveStore {
  clientId?: string;
  clientSecret?: string;
  refreshToken?: string;
  accessToken?: string;
  expiry?: number;
  email?: string;
}

function storePath(): string {
  return path.join(app.getPath('userData'), 'gdrive.json');
}

function readStore(): GDriveStore {
  try {
    const raw = fs.readFileSync(storePath(), 'utf-8');
    const parsed = JSON.parse(raw) as GDriveStore;
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

function writeStore(patch: Partial<GDriveStore>): GDriveStore {
  const next: GDriveStore = { ...readStore(), ...patch };
  try {
    fs.writeFileSync(storePath(), JSON.stringify(next, null, 2), 'utf-8');
  } catch {
    // Каталог может быть недоступен — не срываем операцию из-за этого
  }
  return next;
}

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  avif: 'image/avif',
  ico: 'image/x-icon',
  pdf: 'application/pdf',
  txt: 'text/plain',
  md: 'text/markdown',
  csv: 'text/csv',
  json: 'application/json',
  xml: 'application/xml',
  html: 'text/html',
  zip: 'application/zip',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

function mimeFor(filePath: string): string {
  const ext = path.extname(filePath).toLowerCase().replace('.', '');
  return MIME_BY_EXT[ext] || 'application/octet-stream';
}

const HTML_OK =
  '<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>WorkSpace</title></head>' +
  '<body style="font-family:Segoe UI,sans-serif;background:#0b0f17;color:#e2e8f0;display:flex;' +
  'align-items:center;justify-content:center;height:100vh;margin:0"><div style="text-align:center">' +
  '<h2>Готово</h2><p style="color:#94a3b8">Авторизация Google завершена. Вернитесь в WorkSpace — ' +
  'это окно можно закрыть.</p></div></body></html>';

function htmlError(message: string): string {
  return (
    '<!doctype html><html lang="ru"><head><meta charset="utf-8"></head>' +
    '<body style="font-family:Segoe UI,sans-serif;background:#0b0f17;color:#f87171;display:flex;' +
    'align-items:center;justify-content:center;height:100vh;margin:0"><div style="text-align:center">' +
    '<h2>Ошибка авторизации</h2><p>' +
    message +
    '</p></div></body></html>'
  );
}

async function exchangeCode(code: string): Promise<GDriveStore> {
  const cfg = readStore();
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      code,
      client_id: cfg.clientId || '',
      client_secret: cfg.clientSecret || '',
      redirect_uri: REDIRECT_URI,
      grant_type: 'authorization_code',
    }),
  });
  if (!res.ok) {
    throw new Error('Не удалось обменять код на токен: ' + (await res.text()).slice(0, 300));
  }
  const data = (await res.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in?: number;
  };

  const patch: Partial<GDriveStore> = {
    accessToken: data.access_token,
    expiry: Date.now() + (data.expires_in || 3600) * 1000,
  };
  if (data.refresh_token) patch.refreshToken = data.refresh_token;
  const saved = writeStore(patch);

  // Профиль — чтобы показать, к какому аккаунту подключено хранилище
  try {
    const info = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: 'Bearer ' + data.access_token },
    });
    if (info.ok) {
      const profile = (await info.json()) as { email?: string };
      if (profile.email) {
        writeStore({ email: profile.email });
        saved.email = profile.email;
      }
    }
  } catch {
    // Профиль не обязателен для загрузки
  }
  return saved;
}

export async function getGoogleAccessToken(): Promise<string> {
  const store = readStore();
  if (!store.refreshToken) throw new Error('Google Drive не подключён');
  if (store.accessToken && store.expiry && Date.now() < store.expiry - 60000) {
    return store.accessToken;
  }
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: store.clientId || '',
      client_secret: store.clientSecret || '',
      refresh_token: store.refreshToken,
      grant_type: 'refresh_token',
    }),
  });
  if (!res.ok) {
    throw new Error('Не удалось обновить токен: ' + (await res.text()).slice(0, 300));
  }
  const data = (await res.json()) as { access_token: string; expires_in?: number };
  writeStore({
    accessToken: data.access_token,
    expiry: Date.now() + (data.expires_in || 3600) * 1000,
  });
  return data.access_token;
}

async function createFolder(token: string, name: string, parentId?: string): Promise<string> {
  const res = await fetch('https://www.googleapis.com/drive/v3/files?fields=id', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: parentId ? [parentId] : undefined,
    }),
  });
  if (!res.ok) throw new Error('Не удалось создать папку: ' + (await res.text()).slice(0, 300));
  const data = (await res.json()) as { id: string };
  return data.id;
}

async function uploadFile(
  token: string,
  filePath: string,
  parentId?: string
): Promise<{ id: string; link?: string }> {
  const buffer = await fs.promises.readFile(filePath);
  const boundary = 'wsboundary' + Math.random().toString(36).slice(2);
  const metadata = { name: path.basename(filePath), parents: parentId ? [parentId] : undefined };

  // multipart/related: метаданные + содержимое одним запросом
  const head = Buffer.from(
    '--' +
      boundary +
      '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
      JSON.stringify(metadata) +
      '\r\n--' +
      boundary +
      '\r\nContent-Type: ' +
      mimeFor(filePath) +
      '\r\n\r\n'
  );
  const tail = Buffer.from('\r\n--' + boundary + '--');
  const body = Buffer.concat([head, buffer, tail]);

  const res = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink',
    {
      method: 'POST',
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'multipart/related; boundary=' + boundary,
      },
      body,
    }
  );
  if (!res.ok) throw new Error('Не удалось загрузить файл: ' + (await res.text()).slice(0, 300));
  const data = (await res.json()) as { id: string; webViewLink?: string };
  return { id: data.id, link: data.webViewLink };
}

interface UploadStats {
  uploaded: number;
  failed: string[];
}

async function uploadDirectory(
  token: string,
  dirPath: string,
  parentId: string,
  stats: UploadStats
): Promise<void> {
  let entries: fs.Dirent[] = [];
  try {
    entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
  } catch (err) {
    stats.failed.push(path.basename(dirPath) + ': ' + (err as Error).message);
    return;
  }

  for (const entry of entries) {
    const full = path.join(dirPath, entry.name);
    try {
      if (entry.isDirectory()) {
        const subId = await createFolder(token, entry.name, parentId);
        await uploadDirectory(token, full, subId, stats);
      } else {
        await uploadFile(token, full, parentId);
        stats.uploaded++;
      }
    } catch (err) {
      stats.failed.push(entry.name + ': ' + (err as Error).message);
    }
  }
}

async function uploadPaths(
  targets: string[],
  folderName?: string
): Promise<{ success: boolean; uploaded: number; failed: string[]; link?: string }> {
  const token = await getGoogleAccessToken();
  const stats: UploadStats = { uploaded: 0, failed: [] };
  let link: string | undefined;

  // Несколько отдельных файлов с именем папки складываются в один альбом.
  const fileTargets: string[] = [];
  const dirTargets: string[] = [];
  for (const target of targets) {
    try {
      const stat = await fs.promises.stat(target);
      if (stat.isDirectory()) dirTargets.push(target);
      else fileTargets.push(target);
    } catch (err) {
      stats.failed.push(path.basename(target) + ': ' + (err as Error).message);
    }
  }

  if (folderName && fileTargets.length > 0) {
    try {
      const folderId = await createFolder(token, folderName);
      link = 'https://drive.google.com/drive/folders/' + folderId;
      for (const file of fileTargets) {
        try {
          await uploadFile(token, file, folderId);
          stats.uploaded++;
        } catch (err) {
          stats.failed.push(path.basename(file) + ': ' + (err as Error).message);
        }
      }
    } catch (err) {
      stats.failed.push(folderName + ': ' + (err as Error).message);
    }
  } else {
    for (const file of fileTargets) {
      try {
        const res = await uploadFile(token, file);
        stats.uploaded++;
        if (res.link) link = res.link;
      } catch (err) {
        stats.failed.push(path.basename(file) + ': ' + (err as Error).message);
      }
    }
  }

  for (const target of dirTargets) {
    try {
      const folderId = await createFolder(token, folderName || path.basename(target));
      link = 'https://drive.google.com/drive/folders/' + folderId;
      await uploadDirectory(token, target, folderId, stats);
    } catch (err) {
      stats.failed.push(path.basename(target) + ': ' + (err as Error).message);
    }
  }

  return { success: stats.failed.length === 0, uploaded: stats.uploaded, failed: stats.failed, link };
}

/** Открывает браузер и ждёт код авторизации на loopback-адресе. */
function authorize(): Promise<string> {
  const store = readStore();
  if (!store.clientId || !store.clientSecret) {
    return Promise.reject(new Error('Укажите Client ID и Client Secret приложения Google'));
  }

  const authUrl =
    'https://accounts.google.com/o/oauth2/v2/auth?' +
    new URLSearchParams({
      client_id: store.clientId,
      redirect_uri: REDIRECT_URI,
      response_type: 'code',
      scope: SCOPES,
      access_type: 'offline',
      prompt: 'consent',
      include_granted_scopes: 'true',
    }).toString();

  return new Promise<string>((resolve, reject) => {
    let settled = false;
    let timer: NodeJS.Timeout | undefined;

    const finish = (fn: () => void) => {
      if (settled) return;
      settled = true;
      if (timer) clearTimeout(timer);
      try {
        server.close();
      } catch {
        // Сервер уже закрыт
      }
      fn();
    };

    const server = http.createServer((req, res) => {
      const url = new URL(req.url || '/', REDIRECT_URI);
      const code = url.searchParams.get('code');
      const error = url.searchParams.get('error');
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });

      if (code) {
        res.end(HTML_OK);
        finish(() => {
          exchangeCode(code)
            .then((saved) => resolve(saved.email || ''))
            .catch(reject);
        });
      } else {
        res.end(htmlError(error || 'Код авторизации не получен'));
        finish(() => reject(new Error(error || 'Авторизация отклонена')));
      }
    });

    server.on('error', (err) => finish(() => reject(err as Error)));

    timer = setTimeout(
      () => finish(() => reject(new Error('Истекло время ожидания авторизации (5 минут)'))),
      5 * 60 * 1000
    );

    server.listen(LOOPBACK_PORT, '127.0.0.1', () => {
      void shell.openExternal(authUrl);
    });
  });
}

export function registerGDrive(): void {
  ipcMain.handle('gdrive:getConfig', () => {
    const store = readStore();
    return { clientId: store.clientId || '', clientSecret: store.clientSecret || '' };
  });

  ipcMain.handle('gdrive:setConfig', (_event, cfg: { clientId?: string; clientSecret?: string }) => {
    writeStore({ clientId: cfg.clientId || '', clientSecret: cfg.clientSecret || '' });
    return { success: true };
  });

  ipcMain.handle('gdrive:status', () => {
    const store = readStore();
    return {
      configured: Boolean(store.clientId && store.clientSecret),
      connected: Boolean(store.refreshToken),
      email: store.email || '',
    };
  });

  ipcMain.handle('gdrive:connect', async () => {
    try {
      const email = await authorize();
      return { success: true, email };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  ipcMain.handle('gdrive:disconnect', () => {
    writeStore({ refreshToken: '', accessToken: '', expiry: 0, email: '' });
    return { success: true };
  });

  ipcMain.handle('gdrive:upload', async (_event, targets: string[], folderName?: string) => {
    try {
      if (!Array.isArray(targets) || targets.length === 0) {
        return { success: false, uploaded: 0, failed: [], error: 'Не выбраны файлы для загрузки' };
      }
      return await uploadPaths(targets, folderName);
    } catch (err) {
      return { success: false, uploaded: 0, failed: [], error: (err as Error).message };
    }
  });
}
