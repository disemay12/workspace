/**
 * Экспорт файлов во внешние сервисы (Imgur, Catbox, 0x0.st, Google Drive).
 *
 * Сетевые запросы выполняет главный процесс (electron/share.ts и
 * electron/gdrive.ts): браузерный fetch подчинён CORS, а у файлообменников
 * нет заголовков, разрешающих запросы из рендерера.
 */
import { imageMime } from './fileSystemService';
import { extName } from '../lib/utils';

export interface ShareServiceInfo {
  id: string;
  label: string;
  hint: string;
  /** Сервису нужен Client-ID приложения (Imgur). */
  needsClientId?: boolean;
  /** Сервис принимает только изображения. */
  imagesOnly?: boolean;
  /** Экспорт открывает сайт и подставляет файлы в форму. */
  opensSite?: boolean;
}

export const SHARE_SERVICES: ShareServiceInfo[] = [
  { id: 'catbox', label: 'Catbox', hint: 'Постоянное хранилище, без ключа' },
  { id: '0x0', label: '0x0.st', hint: 'Временная ссылка, без ключа' },
  { id: 'uguu', label: 'Uguu', hint: 'Временная ссылка (3 часа)' },
  { id: 'tmpfiles', label: 'Tmpfiles', hint: 'Временная ссылка (1 час)' },
  { id: 'telegraph', label: 'Telegraph', hint: 'Хостинг картинок, без ключа', imagesOnly: true },
  { id: 'fileio', label: 'File.io', hint: 'Одноразовая ссылка' },
  {
    id: 'imgur',
    label: 'Imgur',
    hint: 'Нужен Client-ID приложения Imgur',
    needsClientId: true,
    imagesOnly: true,
  },
  {
    id: 'imgbb',
    label: 'ImgBB',
    hint: 'Откроет сайт и добавит выбранные фото в форму',
    imagesOnly: true,
    opensSite: true,
  },
  {
    id: 'freeimage',
    label: 'Freeimage.host',
    hint: 'Откроет сайт и добавит выбранные фото в форму',
    imagesOnly: true,
    opensSite: true,
  },
  {
    id: 'yapx',
    label: 'Япикс (yapx.ru)',
    hint: 'Альбом до 10 фото; большие наборы — пачками по альбомам',
    imagesOnly: true,
    opensSite: true,
  },
];

/** Лимиты сайтового хостинга — те же, что применяет главный процесс. */
export interface ImageHostLimits {
  id: string;
  title: string;
  formats: string[];
  maxFileMb: number;
  perAlbum: number;
  perAlbumAuthorized?: number;
  perHour?: number;
}

export async function getImageHostLimits(): Promise<Record<string, ImageHostLimits>> {
  const client = window.electronAPI;
  if (!client?.imageHostLimits) return {};
  return client.imageHostLimits();
}

/** Разбиение на пачки: одна пачка — один альбом сайта. */
export function splitIntoBatches<T>(items: T[], perBatch: number): T[][] {
  const size = Math.max(1, Math.floor(perBatch));
  const batches: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    batches.push(items.slice(index, index + size));
  }
  return batches;
}

/** Файлы, которые сайт примет, и причины отказа для остальных. */
export function filterHostCompatible<T extends { path: string; name: string }>(
  files: T[],
  limits: ImageHostLimits | undefined,
  sizes: Record<string, number>
): { accepted: T[]; skipped: { path: string; name: string; reason: string }[] } {
  const accepted: T[] = [];
  const skipped: { path: string; name: string; reason: string }[] = [];
  for (const file of files) {
    const ext = (file.name.split('.').pop() ?? '').toLowerCase();
    if (limits) {
      if (!limits.formats.includes(ext)) {
        skipped.push({
          path: file.path,
          name: file.name,
          reason: `.${ext} не поддерживается Япиксом (${limits.formats.join(', ')})`,
        });
        continue;
      }
      const size = sizes[file.path] ?? 0;
      if (size > limits.maxFileMb * 1024 * 1024) {
        skipped.push({
          path: file.path,
          name: file.name,
          reason: `${(size / (1024 * 1024)).toFixed(1)} МБ — больше лимита ${limits.maxFileMb} МБ`,
        });
        continue;
      }
    }
    accepted.push(file);
  }
  return { accepted, skipped };
}

/** Загрузка на сайт, который принимает файлы через свою форму (Япикс, ImgBB). */
export async function shareViaSite(
  host: 'imgbb' | 'freeimage' | 'yapx',
  files: { path: string; name: string }[],
  options: { authorized?: boolean; perAlbum?: number } = {}
): Promise<{
  success: boolean;
  count: number;
  albums: string[];
  skipped: { path: string; reason: string }[];
  error?: string;
}> {
  const client = window.electronAPI;
  const empty = { success: false, count: 0, albums: [], skipped: [] };
  if (!client) return { ...empty, error: 'Экспорт доступен только в настольном приложении' };
  const paths = files.map((file) => file.path);
  if (host === 'imgbb' || host === 'freeimage') {
    const open = host === 'imgbb' ? client.openImgbb : client.openFreeimage;
    const result = await open?.(paths);
    return {
      success: Boolean(result?.success),
      count: result?.count ?? 0,
      albums: [],
      skipped: [],
      error: result?.error,
    };
  }
  const result = await client.openYapx?.(paths, options);
  return {
    success: Boolean(result?.success),
    count: result?.count ?? 0,
    albums: result?.albums ?? [],
    skipped: result?.skipped ?? [],
    error: result?.error,
  };
}

export const IMGUR_CLIENT_ID_KEY = 'ws_imgur_client_id';

export function getImgurClientId(): string {
  try {
    return localStorage.getItem(IMGUR_CLIENT_ID_KEY) || '';
  } catch {
    return '';
  }
}

export function setImgurClientId(value: string): void {
  try {
    localStorage.setItem(IMGUR_CLIENT_ID_KEY, value.trim());
  } catch {
    /* localStorage недоступен — значение просто не сохранится */
  }
}

export interface UploadOutcome {
  success: boolean;
  url?: string;
  error?: string;
}

export function isSharingAvailable(): boolean {
  return Boolean(window.electronAPI && window.electronAPI.uploadToService);
}

/** Загрузка уже готовых данных (например, отредактированного изображения). */
export async function uploadBase64(
  service: string,
  fileName: string,
  mime: string,
  base64: string
): Promise<UploadOutcome> {
  const client = window.electronAPI;
  if (!client || !client.uploadToService) {
    return { success: false, error: 'Экспорт доступен только в настольном приложении' };
  }
  const info = SHARE_SERVICES.find((s) => s.id === service);
  if (info && info.needsClientId && !getImgurClientId()) {
    return { success: false, error: 'Не задан Imgur Client-ID' };
  }
  const res = await client.uploadToService({
    service,
    fileName,
    mime,
    base64,
    clientId: service === 'imgur' ? getImgurClientId() : undefined,
  });
  return res;
}

export interface AlbumOutcome {
  success: boolean;
  url?: string;
  uploaded: number;
  failed: string[];
  links: string[];
  error?: string;
}

/** Несколько изображений в один альбом (сейчас — Imgur). */
export async function uploadAlbum(
  service: string,
  title: string,
  files: { path: string; name: string }[]
): Promise<AlbumOutcome> {
  const client = window.electronAPI;
  if (!client || !client.uploadAlbum) {
    return { success: false, uploaded: 0, failed: [], links: [], error: 'Альбом доступен только в настольном приложении' };
  }
  if (service === 'imgur' && !getImgurClientId()) {
    return { success: false, uploaded: 0, failed: [], links: [], error: 'Не задан Imgur Client-ID' };
  }
  const images: { fileName: string; mime: string; base64: string; title: string }[] = [];
  const failed: string[] = [];
  for (const file of files) {
    const read = await client.readFile(file.path, 'base64');
    if (!read.success || !read.content) {
      failed.push(file.name + ': ' + (read.error || 'не прочитан'));
      continue;
    }
    images.push({
      fileName: file.name,
      mime: imageMime(extName(file.name)),
      base64: read.content,
      title: file.name,
    });
  }
  if (!images.length) {
    return { success: false, uploaded: 0, failed, links: [], error: 'Не удалось прочитать изображения' };
  }
  const res = await client.uploadAlbum({
    service,
    title,
    clientId: service === 'imgur' ? getImgurClientId() : undefined,
    images,
  });
  return { ...res, failed: failed.concat(res.failed || []) };
}

/** Несколько фото на сервис без регистрации. Возвращает список прямых ссылок. */
export async function uploadMany(
  service: string,
  files: { path: string; name: string }[]
): Promise<AlbumOutcome> {
  const links: string[] = [];
  const failed: string[] = [];
  for (const file of files) {
    const result = await uploadFilePath(service, file.path);
    if (result.success && result.url) links.push(result.url);
    else failed.push(file.name + ': ' + (result.error || 'ошибка'));
  }
  return {
    success: failed.length === 0,
    uploaded: links.length,
    failed,
    links,
    url: links.join('\n'),
  };
}

/** Загрузка файла с диска. */
export async function uploadFilePath(service: string, filePath: string): Promise<UploadOutcome> {
  const client = window.electronAPI;
  if (!client) return { success: false, error: 'Экспорт доступен только в настольном приложении' };
  const read = await client.readFile(filePath, 'base64');
  if (!read.success || !read.content) {
    return { success: false, error: read.error || 'Не удалось прочитать файл' };
  }
  const name = filePath.split(/[\\/]/).pop() || 'file';
  return uploadBase64(service, name, imageMime(extName(name)), read.content);
}

/** Копирует ссылку в системный буфер обмена. */
export async function copyToClipboard(text: string): Promise<boolean> {
  const client = window.electronAPI;
  if (!client || !client.writeClipboard) return false;
  const res = await client.writeClipboard(text);
  return res.success;
}

/* ============================================================
   Google Drive
   ============================================================ */

export interface GDriveStatus {
  configured: boolean;
  connected: boolean;
  email?: string;
}

export interface GDriveUploadResult {
  success: boolean;
  uploaded: number;
  failed: string[];
  link?: string;
  error?: string;
}

export async function gdriveStatus(): Promise<GDriveStatus> {
  const client = window.electronAPI;
  if (!client || !client.gdriveStatus) {
    return { configured: false, connected: false };
  }
  return client.gdriveStatus();
}

export async function gdriveGetConfig(): Promise<{ clientId: string; clientSecret: string }> {
  const client = window.electronAPI;
  if (!client || !client.gdriveGetConfig) return { clientId: '', clientSecret: '' };
  return client.gdriveGetConfig();
}

export async function gdriveSetConfig(clientId: string, clientSecret: string): Promise<void> {
  await window.electronAPI?.gdriveSetConfig?.({ clientId, clientSecret });
}

export async function gdriveConnect(): Promise<{ success: boolean; email?: string; error?: string }> {
  const client = window.electronAPI;
  if (!client || !client.gdriveConnect) {
    return { success: false, error: 'Google Drive доступен только в настольном приложении' };
  }
  return client.gdriveConnect();
}

export async function gdriveDisconnect(): Promise<void> {
  await window.electronAPI?.gdriveDisconnect?.();
}

export async function gdriveUpload(
  paths: string[],
  folderName?: string
): Promise<GDriveUploadResult> {
  const client = window.electronAPI;
  if (!client || !client.gdriveUpload) {
    return { success: false, uploaded: 0, failed: [], error: 'Google Drive недоступен' };
  }
  return client.gdriveUpload(paths, folderName);
}

/* ============================================================
   Google: раздельные согласия, диагностика
   ============================================================ */

export type GoogleService = 'drive' | 'youtube';

export interface GoogleFullStatus {
  configured: boolean;
  drive: { connected: boolean; email?: string; scopes?: string };
  youtube: { connected: boolean; email?: string; scopes?: string };
  email?: string;
}

export interface GoogleDiagnoseResult {
  configured: boolean;
  tokenValid: boolean;
  driveApi: 'ok' | 'no-access' | 'error' | 'not-connected';
  youtubeApi: 'ok' | 'no-access' | 'error' | 'not-connected';
  quotaBytes?: number;
  channelTitle?: string;
  error?: string;
}

const EMPTY_GOOGLE_STATUS: GoogleFullStatus = {
  configured: false,
  drive: { connected: false },
  youtube: { connected: false },
};

export async function googleFullStatus(): Promise<GoogleFullStatus> {
  const client = window.electronAPI;
  if (!client?.googleStatus) return EMPTY_GOOGLE_STATUS;
  try {
    return await client.googleStatus();
  } catch {
    return EMPTY_GOOGLE_STATUS;
  }
}

export async function googleConnect(
  service: GoogleService
): Promise<{ success: boolean; email?: string; error?: string }> {
  const client = window.electronAPI;
  if (!client?.googleConnect) {
    return { success: false, error: 'Google доступен только в настольном приложении' };
  }
  return client.googleConnect(service);
}

export async function googleDisconnectService(service: GoogleService): Promise<void> {
  await window.electronAPI?.googleDisconnect?.(service);
}

export async function googleDiagnose(): Promise<GoogleDiagnoseResult | null> {
  const client = window.electronAPI;
  if (!client?.googleDiagnose) return null;
  return client.googleDiagnose();
}
