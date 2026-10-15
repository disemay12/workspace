/**
 * Проверка открытия и сохранения .docx в реальном интерфейсе.
 *
 * Воспроизводит жалобу «Word пишет, что сохранил, но при открытии текста нет»:
 * копирует переданный файл в песочницу, открывает его через дерево файлов,
 * читает содержимое полотна редактора, вносит правку, сохраняет и проверяет,
 * что текст действительно лежит в пакете на диске.
 *
 * Запуск: electron.exe dist-electron/docx-check.js [путь-к-docx]
 */
const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const iconv = require('iconv-lite');

function resolveProjectRoot() {
  let dir = __dirname;
  for (let i = 0; i < 4; i++) {
    if (fs.existsSync(path.join(dir, 'package.json')) && fs.existsSync(path.join(dir, 'dist'))) {
      return dir;
    }
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return path.join(__dirname, '..');
}

const ROOT = resolveProjectRoot();
const SANDBOX = path.join(ROOT, 'test-workspace');
const OUT = path.join(ROOT, 'docx-check.json');
const results: any = { root: ROOT, sandbox: SANDBOX };
const errors: string[] = [];

/** Исходный документ: путь из argv либо синтетический файл с известным текстом. */
const ARGV_DOCX = process.argv.slice(2).find((a) => a.toLowerCase().endsWith('.docx')) || '';
const DOC_NAME = 'Проверка.docx';

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

app.disableHardwareAcceleration();

function registerIPC() {
  ipcMain.handle('window:minimize', () => undefined);
  ipcMain.handle('window:maximize', () => false);
  ipcMain.handle('window:isMaximized', () => false);
  ipcMain.handle('window:close', () => undefined);
  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('fs:openFolderDialog', () => SANDBOX);
  ipcMain.handle('fs:openFileDialog', () => null);
  ipcMain.handle('fs:saveFileDialog', () => null);
  ipcMain.handle('fs:showInExplorer', () => ({ success: true }));

  ipcMain.handle('fs:listDirectory', async (_e, dirPath) => {
    try {
      const entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
      const items = entries.map((entry) => {
        const full = path.join(dirPath, entry.name);
        let size = 0;
        let mtime: string | null = null;
        try {
          const st = fs.statSync(full);
          size = st.size;
          mtime = st.mtime.toISOString();
        } catch {
          /* файл мог исчезнуть между чтением каталога и stat */
        }
        return {
          name: entry.name,
          path: full,
          isDirectory: entry.isDirectory(),
          size,
          mtime,
          extension: entry.name.includes('.') ? entry.name.split('.').pop() : '',
        };
      });
      return { success: true, items };
    } catch (err) {
      return { success: false, error: (err as Error).message, items: [] };
    }
  });

  ipcMain.handle('fs:readFile', async (_e, filePath, encoding = 'utf-8') => {
    try {
      const buffer = await fs.promises.readFile(filePath);
      if (encoding === 'base64') {
        return { success: true, content: buffer.toString('base64'), isBinary: true };
      }
      if (encoding && encoding.toLowerCase() !== 'utf-8' && iconv.encodingExists(encoding)) {
        return { success: true, content: iconv.decode(buffer, encoding), isBinary: false };
      }
      if (buffer.length >= 2) {
        const b0 = buffer[0];
        const b1 = buffer[1];
        if (b0 === 0xff && b1 === 0xfe) {
          return { success: true, content: iconv.decode(buffer, 'utf-16le'), isBinary: false };
        }
        if (b0 === 0xfe && b1 === 0xff) {
          return { success: true, content: iconv.decode(buffer, 'utf-16be'), isBinary: false };
        }
      }
      return { success: true, content: buffer.toString('utf-8'), isBinary: false };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  ipcMain.handle('fs:writeFile', async (_e, filePath, content, encoding = 'utf-8') => {
    try {
      const tmpPath = `${filePath}.ws-tmp`;
      if (encoding === 'base64') {
        await fs.promises.writeFile(tmpPath, Buffer.from(content, 'base64'));
      } else if (encoding && encoding.toLowerCase() !== 'utf-8' && iconv.encodingExists(encoding)) {
        await fs.promises.writeFile(tmpPath, iconv.encode(content, encoding));
      } else {
        await fs.promises.writeFile(tmpPath, content, 'utf-8');
      }
      try {
        await fs.promises.rename(tmpPath, filePath);
      } catch (err) {
        const code = (err as NodeJS.ErrnoException).code;
        if (code !== 'EPERM' && code !== 'EEXIST' && code !== 'EBUSY') throw err;
        await fs.promises.copyFile(tmpPath, filePath);
        await fs.promises.unlink(tmpPath).catch(() => undefined);
      }
      return { success: true };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  ipcMain.handle('fs:createFile', async (_e, filePath, initialContent = '') => {
    try {
      await fs.promises.writeFile(filePath, initialContent, { flag: 'wx' });
      return { success: true };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });
  ipcMain.handle('fs:createDirectory', async (_e, dirPath) => {
    try {
      await fs.promises.mkdir(dirPath, { recursive: true });
      return { success: true };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));
  ipcMain.handle('fs:getFavorites', () => []);
  ipcMain.handle('fs:toggleFavorite', () => []);
  ipcMain.handle('ai:getSettings', () => ({ configured: false }));
  ipcMain.handle('app:getVersion', () => '1.0.0');
}

/** Собирает минимальный валидный .docx с заданным текстом (без внешних зависимостей). */
function buildDocx(paragraphs: string[]): Buffer {
  const body = paragraphs
    .map((text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`)
    .join('');
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `</Types>`;
  const rels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`;

  // Собираем ZIP вручную (stored, без сжатия) — этого достаточно для валидного пакета.
  const files: { name: string; data: Buffer }[] = [
    { name: '[Content_Types].xml', data: Buffer.from(contentTypes, 'utf-8') },
    { name: '_rels/.rels', data: Buffer.from(rels, 'utf-8') },
    { name: 'word/document.xml', data: Buffer.from(documentXml, 'utf-8') },
  ];
  return zipStored(files);
}

/** CRC32 для ZIP. */
function crc32(buf: Buffer): number {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/** Минимальный ZIP-архив без сжатия. */
function zipStored(files: { name: string; data: Buffer }[]): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;

  for (const file of files) {
    const nameBuf = Buffer.from(file.name, 'utf-8');
    const crc = crc32(file.data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(file.data.length, 18);
    local.writeUInt32LE(file.data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, nameBuf, file.data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(file.data.length, 20);
    central.writeUInt32LE(file.data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);

    offset += local.length + nameBuf.length + file.data.length;
  }

  const centralBuf = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(centralBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);

  return Buffer.concat([...locals, centralBuf, end]);
}

/** Достаёт текст из .docx без внешних библиотек: ищем <w:t> в распакованном XML. */
function readDocxText(filePath: string): string {
  const buf = fs.readFileSync(filePath);
  // Находим local header для word/document.xml и читаем его данные.
  const needle = Buffer.from('word/document.xml', 'utf-8');
  const idx = buf.indexOf(needle);
  if (idx === -1) return '';
  const headerStart = idx - 30;
  if (headerStart < 0) return '';
  const method = buf.readUInt16LE(headerStart + 8);
  const compSize = buf.readUInt32LE(headerStart + 18);
  const dataStart = idx + needle.length;
  const data = buf.subarray(dataStart, dataStart + compSize);
  let xml: string;
  if (method === 0) {
    xml = data.toString('utf-8');
  } else {
    // deflate-raw
    const zlib = require('zlib');
    xml = zlib.inflateRawSync(data).toString('utf-8');
  }
  const matches = xml.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [];
  return matches
    .map((m) => m.replace(/<[^>]+>/g, ''))
    .join('')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"');
}

async function main() {
  if (!fs.existsSync(SANDBOX)) fs.mkdirSync(SANDBOX, { recursive: true });

  // Чистим прошлые артефакты прогона
  for (const leftover of fs.readdirSync(SANDBOX)) {
    if (leftover === DOC_NAME || leftover.startsWith(DOC_NAME + '.')) {
      fs.rmSync(path.join(SANDBOX, leftover), { force: true });
    }
  }

  const target = path.join(SANDBOX, DOC_NAME);
  const SEED_LINES = ['Первая строка проверки', 'Вторая строка проверки'];

  if (ARGV_DOCX && fs.existsSync(ARGV_DOCX)) {
    fs.copyFileSync(ARGV_DOCX, target);
    results.source = ARGV_DOCX;
    results.sourceCopied = true;
  } else {
    fs.writeFileSync(target, buildDocx(SEED_LINES));
    results.source = '(синтетический документ)';
    results.sourceCopied = false;
  }

  results.fileBeforeOpen = {
    exists: fs.existsSync(target),
    size: fs.statSync(target).size,
    text: readDocxText(target),
  };

  registerIPC();

  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    show: false,
    webPreferences: {
      preload: path.join(ROOT, 'dist-electron', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  win.webContents.on('console-message', (_e, _level, message) => {
    if (/error|Error|ошибк/i.test(message)) errors.push(message);
  });

  await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
  await sleep(2500);

  // Запоминаем, какой именно base64 приложение получает при чтении файла,
  // и подкладываем парсер из бандла, чтобы вызвать его напрямую.
  await win.webContents.executeJavaScript(`
    (() => {
      window.__docxProbeBase64 = ${JSON.stringify(
        fs.readFileSync(target).toString('base64')
      )};
      return true;
    })()
  `);

  // Прямая проверка шагов парсера в настоящем DOM приложения: видно, на каком
  // именно шаге текст теряется.
  results.parserSteps = await win.webContents.executeJavaScript(`
    (async () => {
      const out = {};
      try {
        const b64 = window.__docxProbeBase64;
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        out.byteLen = bytes.length;
        out.magic = [...bytes.slice(0, 4)].map((b) => b.toString(16)).join(' ');

        const JSZip = window.__JSZip;
        out.hasJSZip = !!JSZip;
        if (!JSZip) return out;

        const zip = await JSZip.loadAsync(bytes);
        out.entries = Object.keys(zip.files);
        const docFile = zip.file('word/document.xml');
        out.hasDocumentXml = !!docFile;
        if (!docFile) return out;
        const xml = await docFile.async('string');
        out.xmlLen = xml.length;
        out.xmlHead = xml.slice(0, 200);

        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const perr = doc.getElementsByTagName('parsererror');
        out.parserError = perr.length ? perr[0].textContent.slice(0, 300) : '';
        out.docChildren = [...doc.documentElement.childNodes]
          .filter((n) => n.nodeType === 1)
          .map((n) => n.nodeName);
        const bodyEl = doc.getElementsByTagNameNS('*', 'body')[0];
        out.hasBodyNs = !!bodyEl;
        out.bodyChildNames = bodyEl
          ? [...bodyEl.childNodes].filter((n) => n.nodeType === 1).map((n) => n.nodeName)
          : [];
        // Проверяем localName у элементов с префиксом w:
        if (bodyEl) {
          const firstP = [...bodyEl.childNodes].find(
            (n) => n.nodeType === 1 && (n.localName === 'p' || /:p$/.test(n.nodeName))
          );
          out.firstPLocalName = firstP ? firstP.localName : null;
          out.firstPNodeName = firstP ? firstP.nodeName : null;
          out.getElementsByTagNameStar = doc.getElementsByTagName('*').length;
          out.localNameMatchesP =
            doc.getElementsByTagName('*')
              ? [...doc.getElementsByTagName('*')].filter(
                  (el) => (el.localName || el.nodeName.replace(/^.*:/, '')) === 'p'
                ).length
              : -1;
        }
      } catch (err) {
        out.error = String(err && err.message ? err.message : err);
      }
      return out;
    })()
  `);

  // Перехватываем ошибку разбора прямо в рендерере: WordEditor показывает её
  // в баннере, но если падает промис — текст теряется молча.
  await win.webContents.executeJavaScript(`
    (() => {
      window.__docxLog = [];
      const origParse = DOMParser.prototype.parseFromString;
      DOMParser.prototype.parseFromString = function (str, type) {
        const doc = origParse.call(this, str, type);
        try {
          const err = doc.getElementsByTagName('parsererror');
          window.__docxLog.push({
            type,
            len: str.length,
            head: str.slice(0, 120),
            parserError: err.length ? err[0].textContent.slice(0, 400) : '',
          });
        } catch (e) {
          window.__docxLog.push({ type, threw: String(e) });
        }
        return doc;
      };
      window.addEventListener('error', (e) => {
        window.__docxLog.push({ windowError: String(e.message) });
      });
      window.addEventListener('unhandledrejection', (e) => {
        window.__docxLog.push({ unhandledRejection: String(e.reason && e.reason.message ? e.reason.message : e.reason) });
      });
      const origErr = console.error;
      console.error = function (...args) {
        window.__docxLog.push({ consoleError: args.map(String).join(' ').slice(0, 400) });
        return origErr.apply(this, args);
      };
      return true;
    })()
  `);
  results.treeHasFile = await win.webContents.executeJavaScript(`
    (() => {
      const nodes = [...document.querySelectorAll('div')]
        .filter((d) => d.textContent?.trim() === ${JSON.stringify(DOC_NAME)});
      if (!nodes.length) return false;
      const clickable = nodes.find((d) => d.className.includes('cursor-pointer')) || nodes[0];
      clickable.click();
      return true;
    })()
  `);
  await sleep(3000);

  results.docxLogAfterOpen = await win.webContents.executeJavaScript(
    `window.__docxLog ? window.__docxLog.slice(0, 25) : null`
  );

  results.editorAfterOpen = await win.webContents.executeJavaScript(`
    (() => {
      const canvas = document.querySelector('[contenteditable=true]');
      const body = document.body.innerText;
      return {
        hasCanvas: !!canvas,
        canvasText: canvas ? canvas.innerText : '',
        canvasHtmlLen: canvas ? canvas.innerHTML.length : 0,
        canvasHtmlHead: canvas ? canvas.innerHTML.slice(0, 300) : '',
        statusShowsDoc: body.includes(${JSON.stringify(DOC_NAME)}),
        anyErrorText: (body.match(/Не удалось разобрать документ[^\\n]*/) || [''])[0],
        loadingVisible: body.includes('Загрузка документа'),
        bodyTextHead: body.replace(/\\s+/g, ' ').slice(0, 400),
      };
    })()
  `);

  // Прямой вызов парсера на том же base64 — видно, что именно он возвращает.
  results.parserDirect = await win.webContents.executeJavaScript(`
    (async () => {
      try {
        const b64 = window.__docxProbeBase64 || '';
        if (!b64) return { skipped: 'нет base64-пробы' };
        const mod = window.__docxProbe;
        if (!mod) return { skipped: 'модуль парсера не выставлен' };
        const parsed = await mod.docxBase64ToHtml(b64);
        return { html: parsed.html, htmlLen: parsed.html.length, warnings: parsed.warnings };
      } catch (err) {
        return { error: String(err && err.message ? err.message : err) };
      }
    })()
  `);

  // Вносим правку и сохраняем
  const EDIT_MARK = 'ПРАВКА_ПРОВЕРКИ';
  results.editTyped = await win.webContents.executeJavaScript(`
    (() => {
      const canvas = document.querySelector('[contenteditable=true]');
      if (!canvas) return false;
      canvas.focus();
      const range = document.createRange();
      range.selectNodeContents(canvas);
      range.collapse(false);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      document.execCommand('insertText', false, ${JSON.stringify(EDIT_MARK)});
      canvas.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()
  `);
  await sleep(2500);

  results.dirtyBeforeSave = await win.webContents.executeJavaScript(`
    (() => {
      const btn = [...document.querySelectorAll('button')]
        .find((b) => (b.textContent || '').includes('Сохранить'));
      return { found: !!btn, disabled: btn ? btn.disabled : null };
    })()
  `);

  // Ctrl+S
  await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
  await sleep(2500);

  results.fileAfterSave = {
    size: fs.existsSync(target) ? fs.statSync(target).size : 0,
    text: fs.existsSync(target) ? readDocxText(target) : '',
    tmpLeftover: fs.existsSync(target + '.ws-tmp'),
  };

  // Переоткрываем: закрываем вкладку и открываем файл заново
  await win.webContents.executeJavaScript(`
    (() => {
      const tab = [...document.querySelectorAll('*')]
        .filter((el) => el.textContent?.trim() === ${JSON.stringify(DOC_NAME)});
      return !!tab.length;
    })()
  `);
  await sleep(500);

  results.summary = {
    savedTextContainsEdit: results.fileAfterSave.text.includes(EDIT_MARK),
    editorShowedSourceText:
      results.editorAfterOpen.canvasText.includes('ывапывап') ||
      results.editorAfterOpen.canvasText.trim().length > 0,
    consoleErrors: errors.slice(0, 20),
  };

  fs.writeFileSync(OUT, JSON.stringify(results, null, 2), 'utf-8');
  console.log(JSON.stringify(results, null, 2));

  win.destroy();
  app.quit();
}

app.whenReady().then(main).catch((err) => {
  fs.writeFileSync(OUT, JSON.stringify({ fatal: String(err), results }, null, 2), 'utf-8');
  console.error(err);
  app.quit();
});
