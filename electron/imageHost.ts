import { BrowserWindow, ipcMain, shell } from 'electron';
import * as fs from 'fs';
import * as path from 'path';

const MIME: Record<string, string> = {
  '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png',
  '.gif': 'image/gif', '.webp': 'image/webp', '.bmp': 'image/bmp',
};

/** Страницы, куда фото подставляются в обычную форму, без API-ключа. */
const HOSTS = {
  imgbb: { title: 'ImgBB', url: 'https://ru.imgbb.com/' },
  freeimage: { title: 'Freeimage.host', url: 'https://freeimage.host/ru' },
};

export function registerHostedImageUpload(): void {
  ipcMain.handle('imagehost:open', async (_event, host: keyof typeof HOSTS, filePaths: string[]) => {
    const target = HOSTS[host];
    const existing = (filePaths || []).filter((file) => file && fs.existsSync(file));
    if (!target) return { success: false, error: 'Неизвестный сайт' };
    if (!existing.length) return { success: false, error: 'Не выбраны изображения' };
    const payload = existing.map((file) => ({
      name: path.basename(file),
      type: MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
      data: fs.readFileSync(file).toString('base64'),
    }));
    const win = new BrowserWindow({
      width: 1180, height: 860, title: target.title, autoHideMenuBar: true,
      webPreferences: { sandbox: true, contextIsolation: true },
    });
    win.webContents.setWindowOpenHandler(({ url }) => {
      void shell.openExternal(url);
      return { action: 'deny' };
    });
    const script = fillScript(payload);
    const attach = () => { void win.webContents.executeJavaScript(script).catch(() => undefined); };
    win.webContents.on('did-finish-load', attach);
    win.webContents.on('did-navigate-in-page', attach);
    await win.loadURL(target.url);
    return { success: true, count: existing.length };
  });
}

function fillScript(files: { name: string; type: string; data: string }[]): string {
  return '(' + functionBody + ')(' + JSON.stringify(files) + ');';
}

const functionBody = 'function (files) {' +
  'const input = document.querySelector(\'input[type="file"]:not([capture])\');' +
  'if (!input || input.dataset.wsAttached === "1") return;' +
  'input.dataset.wsAttached = "1";' +
  'const transfer = new DataTransfer();' +
  'for (const item of files) {' +
    'const binary = atob(item.data);' +
    'const bytes = new Uint8Array(binary.length);' +
    'for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);' +
    'transfer.items.add(new File([bytes], item.name, { type: item.type }));' +
  '}' +
  'try {' +
    'input.files = transfer.files;' +
    'input.dispatchEvent(new Event("change", { bubbles: true }));' +
  '} catch (error) { console.warn(error); }' +
'}';
