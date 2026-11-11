/**
 * Google Drive и общие IPC-каналы Google.
 *
 * Авторизация вынесена в `googleAuth.ts` (отдельные согласия Drive/YouTube,
 * PKCE, шифрование токенов). Здесь — выгрузка файлов и папок, в том числе
 * resumable-загрузка больших файлов кусками с прогрессом.
 *
 * Токены не покидают главный процесс: рендерер получает только статус и ссылки.
 */
import { ipcMain } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import {
  connectService,
  diagnose,
  disconnectService,
  getAccessToken,
  googleConfig,
  googleStatus,
  mapGoogleError,
  setGoogleConfig,
  type GoogleService,
} from './googleAuth';

/** Текущий доступ к Google API нужного сервиса. */
export async function getGoogleAccessToken(service: GoogleService = 'drive'): Promise<string> {
  return getAccessToken(service);
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

/** Порог, выше которого выгрузка идёт кусками (resumable). */
const RESUMABLE_THRESHOLD = 8 * 1024 * 1024;
const CHUNK = 8 * 1024 * 1024;

type ProgressHandler = (sent: number, total: number) => void;

async function getJson<T>(url: string, token: string): Promise<T> {
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  if (!res.ok) throw new Error(mapGoogleError(await res.text()));
  return (await res.json()) as T;
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
  if (!res.ok) throw new Error(mapGoogleError(await res.text()));
  const data = (await res.json()) as { id: string };
  return data.id;
}

/** Уникальное имя в папке: Google разрешает дубли, но человек ждёт «(2)». */
async function uniqueName(token: string, name: string, parentId?: string): Promise<string> {
  const escaped = name.replace(/'/g, "\\'");
  const query = parentId ? `name = '${escaped}' and '${parentId}' in parents and trashed = false` : `name = '${escaped}' and trashed = false`;
  try {
    const data = await getJson<{ files?: { id: string }[] }>(
      'https://www.googleapis.com/drive/v3/files?fields=files(id)&q=' + encodeURIComponent(query),
      token
    );
    if (!data.files?.length) return name;
  } catch {
    return name;
  }
  const ext = path.extname(name);
  const base = path.basename(name, ext);
  const suffix = ' ' + new Date().toLocaleTimeString('ru-RU').replace(/:/g, '-');
  return base + suffix + ext;
}

async function uploadMultipart(
  token: string,
  filePath: string,
  parentId?: string
): Promise<{ id: string; link?: string }> {
  const buffer = await fs.promises.readFile(filePath);
  const boundary = 'wsboundary' + Math.random().toString(36).slice(2);
  const metadata = {
    name: await uniqueName(token, path.basename(filePath), parentId),
    parents: parentId ? [parentId] : undefined,
  };
  const head = Buffer.from(
    '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
      JSON.stringify(metadata) +
      '\r\n--' + boundary + '\r\nContent-Type: ' + mimeFor(filePath) + '\r\n\r\n'
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
  if (!res.ok) throw new Error(mapGoogleError(await res.text()));
  const data = (await res.json()) as { id: string; webViewLink?: string };
  return { id: data.id, link: data.webViewLink };
}

/**
 * Resumable-загрузка: сервер выдаёт адрес сессии, файл уходит кусками по 8 МБ.
 * Так большой файл не нужно целиком держать в памяти и видно прогресс.
 */
async function uploadResumable(
  token: string,
  filePath: string,
  parentId: string | undefined,
  onProgress?: ProgressHandler
): Promise<{ id: string; link?: string }> {
  const stat = await fs.promises.stat(filePath);
  const name = await uniqueName(token, path.basename(filePath), parentId);
  const start = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id,name,webViewLink',
    {
      method: 'POST',
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Upload-Content-Length': String(stat.size),
        'X-Upload-Content-Type': mimeFor(filePath),
      },
      body: JSON.stringify({ name, parents: parentId ? [parentId] : undefined }),
    }
  );
  if (!start.ok) throw new Error(mapGoogleError(await start.text()));
  const location = start.headers.get('location');
  if (!location) throw new Error('Google не выдал адрес загрузки');

  const handle = await fs.promises.open(filePath, 'r');
  try {
    let offset = 0;
    while (offset < stat.size) {
      let length = Math.min(CHUNK, stat.size - offset);
      const buffer = Buffer.alloc(length);
      const read = await handle.read(buffer, 0, length, offset);
      length = read.bytesRead;
      if (length <= 0) break;
      const end = offset + length - 1;
      const put = await fetch(location, {
        method: 'PUT',
        headers: {
          Authorization: 'Bearer ' + token,
          'Content-Type': mimeFor(filePath),
          'Content-Range': `bytes ${offset}-${end}/${stat.size}`,
        },
        body: buffer.subarray(0, length),
      });
      onProgress?.(end + 1, stat.size);
      if (put.status === 200 || put.status === 201) {
        const data = (await put.json()) as { id: string; webViewLink?: string };
        return { id: data.id, link: data.webViewLink };
      }
      if (put.status !== 308) throw new Error(mapGoogleError(await put.text()));
      const range = put.headers.get('range');
      const uploaded = range ? Number(range.split('-')[1]) + 1 : end + 1;
      offset = Number.isFinite(uploaded) ? uploaded : end + 1;
    }
  } finally {
    await handle.close();
  }
  throw new Error('Google не подтвердил загрузку файла');
}

async function uploadFile(
  token: string,
  filePath: string,
  parentId?: string,
  onProgress?: ProgressHandler
): Promise<{ id: string; link?: string }> {
  const stat = await fs.promises.stat(filePath);
  if (stat.size > RESUMABLE_THRESHOLD) {
    return uploadResumable(token, filePath, parentId, onProgress);
  }
  const result = await uploadMultipart(token, filePath, parentId);
  onProgress?.(stat.size, stat.size);
  return result;
}

interface UploadStats {
  uploaded: number;
  failed: string[];
}

async function uploadDirectory(
  token: string,
  dirPath: string,
  parentId: string,
  stats: UploadStats,
  onProgress?: ProgressHandler
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
        await uploadDirectory(token, full, subId, stats, onProgress);
      } else {
        await uploadFile(token, full, parentId, onProgress);
        stats.uploaded++;
      }
    } catch (err) {
      stats.failed.push(entry.name + ': ' + mapGoogleError(err));
    }
  }
}

async function uploadPaths(
  targets: string[],
  folderName?: string,
  onProgress?: ProgressHandler
): Promise<{ success: boolean; uploaded: number; failed: string[]; link?: string }> {
  const token = await getGoogleAccessToken('drive');
  const stats: UploadStats = { uploaded: 0, failed: [] };
  let link: string | undefined;

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
          await uploadFile(token, file, folderId, onProgress);
          stats.uploaded++;
        } catch (err) {
          stats.failed.push(path.basename(file) + ': ' + mapGoogleError(err));
        }
      }
    } catch (err) {
      stats.failed.push(folderName + ': ' + mapGoogleError(err));
    }
  } else {
    for (const file of fileTargets) {
      try {
        const res = await uploadFile(token, file, undefined, onProgress);
        stats.uploaded++;
        if (res.link) link = res.link;
      } catch (err) {
        stats.failed.push(path.basename(file) + ': ' + mapGoogleError(err));
      }
    }
  }

  for (const target of dirTargets) {
    try {
      const folderId = await createFolder(token, folderName || path.basename(target));
      link = 'https://drive.google.com/drive/folders/' + folderId;
      await uploadDirectory(token, target, folderId, stats, onProgress);
    } catch (err) {
      stats.failed.push(path.basename(target) + ': ' + mapGoogleError(err));
    }
  }

  return { success: stats.failed.length === 0, uploaded: stats.uploaded, failed: stats.failed, link };
}

export function registerGDrive(): void {
  ipcMain.handle('gdrive:getConfig', () => googleConfig());

  ipcMain.handle('gdrive:setConfig', (_event, cfg: { clientId?: string; clientSecret?: string }) => {
    setGoogleConfig(cfg.clientId ?? '', cfg.clientSecret ?? '');
    return { success: true };
  });

  // Старый контракт статуса: интерфейс Диска читает {configured, connected, email}.
  ipcMain.handle('gdrive:status', () => {
    const status = googleStatus();
    return {
      configured: status.configured,
      connected: status.drive.connected,
      email: status.drive.email ?? status.email ?? '',
    };
  });

  ipcMain.handle('gdrive:connect', async () => connectService('drive'));
  ipcMain.handle('gdrive:disconnect', async () => {
    await disconnectService('drive');
    return { success: true };
  });

  ipcMain.handle('gdrive:upload', async (event, targets: string[], folderName?: string) => {
    try {
      if (!Array.isArray(targets) || targets.length === 0) {
        return { success: false, uploaded: 0, failed: [], error: 'Не выбраны файлы для загрузки' };
      }
      const result = await uploadPaths(targets, folderName, (sent, total) => {
        event.sender.send('gdrive:progress', { filePath: targets[0], sent, total, folderName });
      });
      return result;
    } catch (err) {
      return { success: false, uploaded: 0, failed: [], error: mapGoogleError(err) };
    }
  });

  /* ---------- Новый контракт: сервисы Google по отдельности ---------- */

  ipcMain.handle('google:connect', async (_event, service: GoogleService) => {
    const target: GoogleService = service === 'youtube' ? 'youtube' : 'drive';
    return connectService(target);
  });

  ipcMain.handle('google:status', () => googleStatus());

  ipcMain.handle('google:disconnect', async (_event, service: GoogleService) => {
    await disconnectService(service === 'youtube' ? 'youtube' : 'drive');
    return { success: true };
  });

  ipcMain.handle('google:diagnose', async () => diagnose());
}
