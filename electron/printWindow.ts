/**
 * Печать HTML из скрытого окна и предпросмотр через printToPDF.
 *
 * Один путь на «Печать» и «Предпросмотр»: предпросмотр — это ровно тот PDF,
 * который уйдёт на принтер, поэтому экран, предпросмотр и бумага не
 * расходятся. Большой документ (со встроенными картинками) не пролезает в
 * data:-URL, поэтому он пишется во временный файл и грузится с диска.
 */
import { BrowserWindow } from 'electron';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

export type PrintMode = 'print' | 'preview';

export interface PrintRequest {
  html: string;
  title?: string;
  mode: PrintMode;
  /** Размер страницы в микронах; без него размер берётся из CSS @page. */
  pageSizeMicrons?: { width: number; height: number };
  landscape?: boolean;
  /** Масштаб печати в процентах (вписывание широкой таблицы в лист). */
  scaleFactor?: number;
}

export interface PrintResult {
  success: boolean;
  canceled?: boolean;
  error?: string;
  /** PDF в base64 для предпросмотра. */
  pdfBase64?: string;
}

export interface PrintHost {
  /** Приватная политика окна: запрет навигации, всплывающих окон, webview. */
  harden?: (win: BrowserWindow) => void;
  /** Ожидание шрифтов и первого кадра перед снятием PDF/печатью. */
  waitForPaint?: (win: BrowserWindow) => Promise<void>;
}

/** Ограничение на размер HTML: защита от случайной гигантской вставки. */
const MAX_HTML_BYTES = 64 * 1024 * 1024;
/** До этого размера HTML грузится data:-URL, дальше — временным файлом. */
const DATA_URL_LIMIT = 512 * 1024;

function pageSizeFor(request: PrintRequest): { width: number; height: number } | undefined {
  const size = request.pageSizeMicrons;
  if (!size) return undefined;
  const width = Math.round(Number(size.width));
  const height = Math.round(Number(size.height));
  if (!(width > 1000) || !(height > 1000)) return undefined;
  return { width, height };
}

async function loadHtml(win: BrowserWindow, html: string): Promise<() => void> {
  if (Buffer.byteLength(html, 'utf8') <= DATA_URL_LIMIT) {
    await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    return () => {};
  }
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'ws-print-'));
  const file = path.join(dir, 'document.html');
  await fs.promises.writeFile(file, html, 'utf8');
  await win.loadFile(file);
  return () => {
    void fs.promises.rm(dir, { recursive: true, force: true }).catch(() => undefined);
  };
}

/**
 * Печатает HTML или отдаёт PDF страницы.
 * Диалог печати системный: отмена возвращается как `canceled`, а не как ошибка.
 */
export async function runPrint(request: PrintRequest, host: PrintHost = {}): Promise<PrintResult> {
  if (!request || typeof request.html !== 'string' || !request.html.trim()) {
    return { success: false, error: 'Нет содержимого для печати' };
  }
  if (Buffer.byteLength(request.html, 'utf8') > MAX_HTML_BYTES) {
    return { success: false, error: 'Документ слишком большой для печати' };
  }

  let win: BrowserWindow | null = null;
  let cleanup = () => {};
  try {
    win = new BrowserWindow({
      show: false,
      skipTaskbar: true,
      width: 1024,
      height: 768,
      webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false },
    });
    host.harden?.(win);
    cleanup = await loadHtml(win, request.html);
    await host.waitForPaint?.(win);

    const pageSize = pageSizeFor(request);
    if (request.mode === 'preview') {
      // Размер бумаги берём из CSS (@page задан в HTML документа): Chromium
      // отказывает printToPDF с нестандартным pageSize и пишет
      // «Failed to generate PDF: Printing failed».
      const pdf = await win.webContents.printToPDF({
        printBackground: true,
        preferCSSPageSize: true,
      });
      return { success: true, pdfBase64: pdf.toString('base64') };
    }

    const runPrintJob = (withSize: boolean) =>
      new Promise<string | null>((resolve) => {
        win!.webContents.print(
          {
            silent: false,
            printBackground: true,
            // Поля задаёт сам документ: страница целиком наша.
            margins: { marginType: 'none' },
            landscape: Boolean(request.landscape),
            ...(withSize && pageSize ? { pageSize } : {}),
            ...(request.scaleFactor && request.scaleFactor > 0
              ? { scaleFactor: Math.round(request.scaleFactor) }
              : {}),
          },
          (success, reason) => resolve(success ? null : reason || 'cancelled')
        );
      });

    let failure = await runPrintJob(Boolean(pageSize));
    if (failure && pageSize && /fail|invalid|unable/i.test(failure)) {
      // Нестандартный размер бумаги принтер может не принять: печатаем на
      // стандартный лист, выбрав его в системном диалоге.
      failure = await runPrintJob(false);
    }

    if (failure) {
      const canceled = /cancel/i.test(failure);
      return canceled ? { success: false, canceled: true } : { success: false, error: failure };
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  } finally {
    cleanup();
    win?.destroy();
  }
}
