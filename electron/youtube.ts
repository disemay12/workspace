/**
 * Загрузка видео на YouTube через тот же Google-аккаунт, что и Drive.
 *
 * Используется resumable upload: файл уходит кусками, а не целиком в память.
 * Метаданные совпадают с формой сайта: название, описание, видимость,
 * категория, теги, язык, «сделано для детей», лицензия и запись в плейлист.
 */
import { ipcMain } from 'electron';
import { mapGoogleError } from './googleAuth';
import * as fs from 'fs';
import * as path from 'path';

interface YoutubeVideo {
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
  notifySubscribers?: boolean;
}

const CHUNK = 8 * 1024 * 1024;

async function token(): Promise<string> {
  const mod = await import('./gdrive');
  // Загрузка видео требует согласия именно на YouTube-скоупы.
  return mod.getGoogleAccessToken('youtube');
}

export async function uploadYoutube(
  filePath: string,
  meta: YoutubeVideo,
  onProgress?: (sent: number, total: number) => void
): Promise<{ success: boolean; id?: string; url?: string; error?: string }> {
  const access = await token();
  const stat = await fs.promises.stat(filePath);
  const snippet: Record<string, unknown> = {
    title: meta.title.slice(0, 100),
    description: (meta.description || '').slice(0, 5000),
    tags: (meta.tags || []).map((t) => t.trim()).filter(Boolean).slice(0, 30),
    categoryId: meta.categoryId || '22',
    defaultLanguage: meta.defaultLanguage || 'ru',
  };
  const status: Record<string, unknown> = {
    privacyStatus: meta.publishAt ? 'private' : meta.privacyStatus,
    selfDeclaredMadeForKids: Boolean(meta.madeForKids),
    embeddable: meta.embeddable !== false,
    license: meta.license || 'youtube',
  };
  if (meta.publishAt) status.publishAt = meta.publishAt;

  const start = await fetch(
    'https://www.googleapis.com/upload/youtube/v3/videos?uploadType=resumable&part=snippet,status',
    {
      method: 'POST',
      headers: {
        Authorization: 'Bearer ' + access,
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Upload-Content-Length': String(stat.size),
        'X-Upload-Content-Type': 'video/*',
      },
      body: JSON.stringify({ snippet, status }),
    }
  );
  if (!start.ok) {
    return { success: false, error: mapGoogleError(await start.text()) };
  }
  const location = start.headers.get('location');
  if (!location) return { success: false, error: 'YouTube не выдал адрес загрузки' };

  const fh = await fs.promises.open(filePath, 'r');
  try {
    let offset = 0;
    while (offset < stat.size) {
      const length = Math.min(CHUNK, stat.size - offset);
      const buffer = Buffer.alloc(length);
      await fh.read(buffer, 0, length, offset);
      const end = offset + length - 1;
      const put = await fetch(location, {
        method: 'PUT',
        headers: {
          Authorization: 'Bearer ' + access,
          'Content-Length': String(length),
          'Content-Type': 'video/*',
          'Content-Range': 'bytes ' + offset + '-' + end + '/' + stat.size,
        },
        body: buffer,
      });
      onProgress?.(end + 1, stat.size);
      if (put.status === 200 || put.status === 201) {
        const data = (await put.json()) as { id?: string };
        if (meta.playlistId && data.id) {
          await fetch('https://www.googleapis.com/youtube/v3/playlistItems?part=snippet', {
            method: 'POST',
            headers: { Authorization: 'Bearer ' + access, 'Content-Type': 'application/json' },
            body: JSON.stringify({
              snippet: {
                playlistId: meta.playlistId,
                resourceId: { kind: 'youtube#video', videoId: data.id },
              },
            }),
          });
        }
        return { success: true, id: data.id, url: data.id ? 'https://youtu.be/' + data.id : undefined };
      }
      if (put.status !== 308) {
        return { success: false, error: mapGoogleError(await put.text()) };
      }
      const range = put.headers.get('range');
      const uploaded = range ? Number(range.split('-')[1]) + 1 : end + 1;
      offset = Number.isFinite(uploaded) ? uploaded : end + 1;
    }
  } finally {
    await fh.close();
  }
  return { success: false, error: 'YouTube не подтвердил загрузку' };
}

export function registerYoutube(): void {
  ipcMain.handle('youtube:categories', async () => {
    try {
      const access = await token();
      const res = await fetch(
        'https://www.googleapis.com/youtube/v3/videoCategories?part=snippet&regionCode=RU',
        { headers: { Authorization: 'Bearer ' + access } }
      );
      if (!res.ok) return { success: false, categories: [], error: mapGoogleError(await res.text()) };
      const data = (await res.json()) as { items?: { id: string; snippet?: { title?: string; assignable?: boolean } }[] };
      const categories = (data.items || [])
        .filter((item) => item.snippet?.assignable !== false)
        .map((item) => ({ id: item.id, title: item.snippet?.title || item.id }));
      return { success: true, categories };
    } catch (err) {
      return { success: false, categories: [], error: (err as Error).message };
    }
  });

  ipcMain.handle('youtube:playlists', async () => {
    try {
      const access = await token();
      const res = await fetch(
        'https://www.googleapis.com/youtube/v3/playlists?part=snippet&mine=true&maxResults=50',
        { headers: { Authorization: 'Bearer ' + access } }
      );
      if (!res.ok) return { success: false, playlists: [], error: mapGoogleError(await res.text()) };
      const data = (await res.json()) as { items?: { id: string; snippet?: { title?: string } }[] };
      return {
        success: true,
        playlists: (data.items || []).map((item) => ({ id: item.id, title: item.snippet?.title || item.id })),
      };
    } catch (err) {
      return { success: false, playlists: [], error: (err as Error).message };
    }
  });

  ipcMain.handle('youtube:upload', async (event, filePath: string, meta: YoutubeVideo) => {
    try {
      if (!filePath || !fs.existsSync(filePath)) return { success: false, error: 'Файл не найден' };
      if (!meta?.title?.trim()) return { success: false, error: 'Нужно название видео' };
      return await uploadYoutube(filePath, meta, (sent, total) => {
        event.sender.send('youtube:progress', { filePath, sent, total });
      });
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });
}
