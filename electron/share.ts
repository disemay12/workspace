/**
 * Загрузка файлов в публичные сервисы обмена: Imgur, Catbox, 0x0.st,
 * Uguu, Tmpfiles, Telegraph, File.io.
 *
 * Выполняется в главном процессе: браузерный fetch подчинён политике CORS,
 * а у файлообменников нет заголовков, разрешающих запросы из рендерера.
 * Здесь же удобно передавать ключ Imgur (Client-ID).
 */
import { ipcMain } from 'electron';

export type ShareServiceId =
  | 'imgur'
  | 'catbox'
  | '0x0'
  | 'uguu'
  | 'tmpfiles'
  | 'telegraph'
  | 'fileio';

export interface ShareRequest {
  service: ShareServiceId;
  fileName: string;
  mime: string;
  base64: string;
  clientId?: string;
}

export interface ShareResponse {
  success: boolean;
  url?: string;
  error?: string;
}

export interface AlbumImage {
  fileName: string;
  mime: string;
  base64: string;
  title?: string;
}

export interface AlbumResponse {
  success: boolean;
  url?: string;
  uploaded: number;
  failed: string[];
  links: string[];
  error?: string;
}

const USER_AGENT = 'WorkSpace/1.0';

function blobFrom(base64: string, mime: string): Blob {
  const buffer = Buffer.from(base64, 'base64');
  return new Blob([buffer], { type: mime || 'application/octet-stream' });
}

async function postForm(
  url: string,
  form: FormData,
  headers: Record<string, string> = {}
): Promise<Response> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 90000);
  try {
    return await fetch(url, {
      method: 'POST',
      headers: { 'User-Agent': USER_AGENT, ...headers },
      body: form,
      signal: controller.signal,
    });
  } finally {
    clearTimeout(timer);
  }
}

async function failure(res: Response): Promise<Error> {
  let detail = '';
  try {
    detail = (await res.text()).slice(0, 300);
  } catch {
    detail = '';
  }
  return new Error('Сервис вернул ошибку ' + res.status + (detail ? ': ' + detail : ''));
}

async function upload(req: ShareRequest, blob: Blob): Promise<string> {
  switch (req.service) {
    case 'imgur': {
      if (!req.clientId) throw new Error('Не задан Imgur Client-ID');
      const form = new FormData();
      form.append('image', blob, req.fileName);
      const res = await postForm('https://api.imgur.com/3/image', form, {
        Authorization: 'Client-ID ' + req.clientId,
      });
      if (!res.ok) throw await failure(res);
      const data = (await res.json()) as { data?: { link?: string } };
      const link = data.data && data.data.link;
      if (!link) throw new Error('Imgur не вернул ссылку');
      return link;
    }
    case 'catbox': {
      const form = new FormData();
      form.append('reqtype', 'fileupload');
      form.append('fileToUpload', blob, req.fileName);
      const res = await postForm('https://catbox.moe/user/api.php', form);
      if (!res.ok) throw await failure(res);
      const text = (await res.text()).trim();
      if (!text.startsWith('http')) throw new Error(text || 'Catbox не вернул ссылку');
      return text;
    }
    case '0x0': {
      const form = new FormData();
      form.append('file', blob, req.fileName);
      const res = await postForm('https://0x0.st', form);
      if (!res.ok) throw await failure(res);
      const text = (await res.text()).trim();
      if (!text.startsWith('http')) throw new Error(text || '0x0.st не вернул ссылку');
      return text;
    }
    case 'uguu': {
      const form = new FormData();
      form.append('files[]', blob, req.fileName);
      const res = await postForm('https://uguu.se/upload.php', form);
      if (!res.ok) throw await failure(res);
      const data = (await res.json()) as { files?: { url?: string }[] };
      const link = data.files && data.files[0] && data.files[0].url;
      if (!link) throw new Error('Uguu не вернул ссылку');
      return link;
    }
    case 'tmpfiles': {
      const form = new FormData();
      form.append('file', blob, req.fileName);
      const res = await postForm('https://tmpfiles.org/api/v1/upload', form);
      if (!res.ok) throw await failure(res);
      const data = (await res.json()) as { data?: { url?: string } };
      const raw = data.data && data.data.url;
      if (!raw) throw new Error('Tmpfiles не вернул ссылку');
      return raw.replace('tmpfiles.org/', 'tmpfiles.org/dl/');
    }
    case 'telegraph': {
      const form = new FormData();
      form.append('file', blob, req.fileName);
      const res = await postForm('https://telegra.ph/upload', form);
      if (!res.ok) throw await failure(res);
      const data = (await res.json()) as { src?: string }[];
      if (Array.isArray(data) && data[0] && data[0].src) {
        return 'https://telegra.ph' + data[0].src;
      }
      throw new Error('Telegraph не вернул ссылку');
    }
    case 'fileio': {
      const form = new FormData();
      form.append('file', blob, req.fileName);
      const res = await postForm('https://file.io', form);
      if (!res.ok) throw await failure(res);
      const data = (await res.json()) as { link?: string };
      if (!data.link) throw new Error('File.io не вернул ссылку');
      return data.link;
    }
    default:
      throw new Error('Неизвестный сервис: ' + String(req.service));
  }
}

async function imgurJson(
  url: string,
  form: FormData,
  clientId: string
): Promise<{ data?: { id?: string; link?: string; deletehash?: string } }> {
  const res = await postForm(url, form, { Authorization: 'Client-ID ' + clientId });
  if (!res.ok) throw await failure(res);
  return (await res.json()) as { data?: { id?: string; link?: string; deletehash?: string } };
}

/** Создаёт публичный альбом Imgur и кладёт в него все переданные фото. */
async function uploadImgurAlbum(
  images: AlbumImage[],
  title: string,
  clientId: string
): Promise<AlbumResponse> {
  if (!clientId) throw new Error('Не задан Imgur Client-ID');
  if (!images.length) throw new Error('Нет изображений для альбома');

  const albumForm = new FormData();
  albumForm.append('title', title || 'WorkSpace');
  albumForm.append('privacy', 'public');
  const album = await imgurJson('https://api.imgur.com/3/album', albumForm, clientId);
  const albumId = album.data && album.data.id;
  const deletehash = album.data && album.data.deletehash;
  if (!albumId || !deletehash) throw new Error('Imgur не создал альбом');

  const links: string[] = [];
  const failed: string[] = [];
  for (const image of images) {
    try {
      const form = new FormData();
      form.append('image', blobFrom(image.base64, image.mime), image.fileName);
      form.append('album', deletehash);
      form.append('title', image.title || image.fileName);
      form.append('type', 'file');
      const uploaded = await imgurJson('https://api.imgur.com/3/image', form, clientId);
      if (uploaded.data && uploaded.data.link) links.push(uploaded.data.link);
      else failed.push(image.fileName + ': нет ссылки');
    } catch (err) {
      failed.push(image.fileName + ': ' + (err as Error).message);
    }
  }

  return {
    success: failed.length === 0,
    url: 'https://imgur.com/a/' + albumId,
    uploaded: links.length,
    failed,
    links,
  };
}
/** Telegraph хранит файлы по отдельности. Страница /createPage собирает их в одну галерею. */
async function uploadTelegraphAlbum(images: AlbumImage[], title: string): Promise<AlbumResponse> {
  if (!images.length) throw new Error('Нет изображений для страницы');
  const accountRes = await fetch('https://api.telegra.ph/createAccount', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ short_name: 'WorkSpace', author_name: 'WorkSpace' }),
  });
  const account = (await accountRes.json()) as { ok?: boolean; result?: { access_token?: string } };
  const token = account.result && account.result.access_token;
  if (!account.ok || !token) throw new Error('Telegraph не создал страницу');

  const links: string[] = [];
  const failed: string[] = [];
  const content: { tag: string; children: string[] }[] = [];
  for (const image of images) {
    try {
      const form = new FormData();
      form.append('file', blobFrom(image.base64, image.mime), image.fileName);
      const uploaded = await postForm('https://telegra.ph/upload', form);
      const data = (await uploaded.json()) as { src?: string }[] | { error?: string };
      const src = Array.isArray(data) && data[0] ? data[0].src : '';
      if (!uploaded.ok || !src) throw new Error('нет ссылки');
      const url = 'https://telegra.ph' + src;
      links.push(url);
      content.push({ tag: 'img', children: [url] });
      content.push({ tag: 'p', children: [image.title || image.fileName] });
    } catch (err) {
      failed.push(image.fileName + ': ' + (err as Error).message);
    }
  }
  if (!links.length) throw new Error(failed.join('; ') || 'Telegraph не принял изображения');

  const pageRes = await fetch('https://api.telegra.ph/createPage', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      access_token: token,
      title: (title || 'WorkSpace').slice(0, 256),
      content,
      return_content: false,
    }),
  });
  const page = (await pageRes.json()) as { ok?: boolean; result?: { url?: string } };
  if (!page.ok || !page.result?.url) throw new Error('Telegraph не собрал страницу');
  return { success: failed.length === 0, url: page.result.url, uploaded: links.length, failed, links };
}

export function registerShare(): void {
  ipcMain.handle(
    'net:uploadAlbum',
    async (
      _event,
      req: { service: string; title?: string; clientId?: string; images: AlbumImage[] }
    ): Promise<AlbumResponse> => {
      try {
        if (req?.service === 'telegraph') return await uploadTelegraphAlbum(req.images || [], req.title || 'WorkSpace');
        if (req?.service === 'imgur') return await uploadImgurAlbum(req.images || [], req.title || 'WorkSpace', req.clientId || '');
        throw new Error('Этот сервис не собирает несколько фото на одну страницу');
      } catch (err) {
        return {
          success: false,
          uploaded: 0,
          failed: [],
          links: [],
          error: (err as Error).message,
        };
      }
    }
  );

  ipcMain.handle('net:upload', async (_event, req: ShareRequest): Promise<ShareResponse> => {
    try {
      if (!req || !req.base64) throw new Error('Нет данных для загрузки');
      const url = await upload(req, blobFrom(req.base64, req.mime));
      return { success: true, url };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });
}
