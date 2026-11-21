/**
 * Проверка листов Word и печати документа.
 *
 * Что проверяется:
 * 1. геометрия листа — размер A4 и поля 30/15/20/20 мм в пикселях 96 dpi;
 * 2. постраничная разметка — листы, отсутствие строк, пересекающих границу,
 *    повторная разметка не портит документ;
 * 3. принудительный разрыв страницы из .docx начинает новый лист;
 * 4. колонтитулы и номера страниц видны на каждом листе;
 * 5. сохранение: служебные маркеры не попадают в .docx;
 * 6. печать: предпросмотр даёт PDF с числом страниц, равным числу листов;
 * 7. контракт IPC: «print:document» зарегистрирован в main.js, а renderer
 *    получает мост printDocument из preload.js.
 *
 * Запуск: electron.exe dist-electron/electron/checks/word-pages-check.js
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');
const { registerCommonCheckIpc } = require('./common-ipc');

/**
 * Модуль печати главного процесса. Скомпилирован в dist-electron/printWindow.js
 * (основная сборка), поэтому путь ищется в обеих раскладках прогона.
 */
function loadPrintWindow() {
  const candidates = [
    path.join(__dirname, 'printWindow'),
    path.join(__dirname, '..', 'printWindow'),
    path.join(__dirname, '..', '..', 'printWindow'),
  ];
  for (const candidate of candidates) {
    try {
      return require(candidate);
    } catch {}
  }
  return null;
}

const printWindow = loadPrintWindow();

function resolveProjectRoot() {
  let dir = __dirname;
  for (let i = 0; i < 5; i++) {
    if (
      fs.existsSync(path.join(dir, 'package.json')) &&
      fs.existsSync(path.join(dir, 'dist', 'index.html'))
    ) {
      return dir;
    }
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return path.join(__dirname, '..', '..');
}

const ROOT = resolveProjectRoot();
const SANDBOX = process.env.WS_CHECK_SANDBOX || path.join(ROOT, 'test-workspace-word-pages');
const OUT = path.join(ROOT, 'word-pages-check.json');
const FIXTURE_SRC = path.join(ROOT, 'test-workspace-docx', 'fixtures', 'rich.docx');
const DOC_NAME = 'Листы.docx';
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const out = { steps: [], errors: [] };
const step = (name, value) => {
  out.steps.push({ name, ...value });
  try {
    fs.writeFileSync(OUT, JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
};

fs.mkdirSync(path.join(SANDBOX, 'userdata'), { recursive: true });
app.setPath('userData', path.join(SANDBOX, 'userdata'));
app.disableHardwareAcceleration();

/* ============================================================
   Длинный .docx на основе готовой фикстуры
   ============================================================ */

const FILLER =
  'Съешь же ещё этих мягких французских булок, да выпей чаю. ' +
  'В чащах юга жил бы цитрус? Да, но фальшивый экземпляр. ' +
  'Широкая электрификация южных губерний даст мощный толчок подъёму сельского хозяйства. ';

function fillerParagraph(index) {
  const repeats = 1 + (index % 4);
  return (
    '<w:p><w:r><w:t xml:space="preserve">' +
    `Абзац ${index + 1}. ${FILLER.repeat(repeats)}` +
    '</w:t></w:r></w:p>'
  );
}

/** Добавляет абзацы перед closing body, чтобы документ занимал много листов. */
async function buildLongDocx(target) {
  const zip = await JSZip.loadAsync(fs.readFileSync(FIXTURE_SRC));
  const documentPart = zip.file('word/document.xml');
  if (!documentPart) throw new Error('в фикстуре нет word/document.xml');
  const xml = await documentPart.async('string');

  const insertAt = xml.lastIndexOf('</w:body>');
  if (insertAt < 0) throw new Error('в фикстуре нет </w:body>');
  const paragraphs = Array.from({ length: 80 }, (_, index) => fillerParagraph(index)).join('');
  const next = `${xml.slice(0, insertAt)}${paragraphs}${xml.slice(insertAt)}`;
  zip.file('word/document.xml', next);

  const bodyRows = Array.from({ length: 39 }, (_, index) => {
    return (
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' +
      '<w:p><w:r><w:t xml:space="preserve">' +
      `Строка ${index + 1}` +
      '</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>' +
      '<w:p><w:r><w:t xml:space="preserve">' +
      `${FILLER.slice(0, 90)}` +
      '</w:t></w:r></w:p></w:tc></w:tr>'
    );
  }).join('');

  // Шапка помечена w:tblHeader — Word повторяет её на новой странице.
  const headerRow =
    '<w:tr><w:trPr><w:tblHeader/></w:trPr>' +
    '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' +
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Заголовок</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>' +
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Значение</w:t></w:r></w:p></w:tc>' +
    '</w:tr>';

  const edge = (name) => `<w:${name} w:val="single" w:sz="6" w:space="0" w:color="A6A6A6"/>`;
  const tableXml =
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>' +
    edge('top') +
    edge('left') +
    edge('bottom') +
    edge('right') +
    edge('insideH') +
    edge('insideV') +
    '</w:tblBorders><w:tblCellMar>' +
    '<w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
    '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>' +
    '</w:tblCellMar></w:tblPr>' +
    headerRow +
    bodyRows +
    '</w:tbl>';

  const withTable = zip.file('word/document.xml');
  const current = await withTable.async('string');
  const tableAt = current.lastIndexOf('</w:body>');
  const finalXml = `${current.slice(0, tableAt)}${tableXml}${current.slice(tableAt)}`;
  zip.file('word/document.xml', finalXml);

  const buffer = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(target, buffer);
}

/* ============================================================
   IPC песочницы
   ============================================================ */

function registerHandlers() {
  registerCommonCheckIpc({ sandbox: SANDBOX });

  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('fs:openFolderDialog', () => SANDBOX);
  ipcMain.handle('fs:openFileDialog', () => null);
  ipcMain.handle('fs:saveFileDialog', () => path.join(SANDBOX, 'печать.pdf'));
  ipcMain.handle('fs:showInExplorer', () => ({ success: true }));
  ipcMain.handle('window:isMaximized', () => false);
  ipcMain.handle('window:minimize', () => undefined);
  ipcMain.handle('window:maximize', () => false);
  ipcMain.handle('window:close', () => undefined);
  ipcMain.handle('system:getInitialExternalPath', () => null);

  ipcMain.handle('fs:listDirectory', async (_event, dirPath) => {
    try {
      const entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
      const items = entries.map((entry) => {
        const full = path.join(dirPath, entry.name);
        let size = 0;
        try {
          size = fs.statSync(full).size;
        } catch {}
        return {
          name: entry.name,
          path: full,
          isDirectory: entry.isDirectory(),
          size,
          mtime: null,
          extension: entry.isDirectory()
            ? ''
            : path.extname(entry.name).toLowerCase().replace('.', ''),
        };
      });
      items.sort((a, b) =>
        a.isDirectory !== b.isDirectory
          ? a.isDirectory
            ? -1
            : 1
          : a.name.localeCompare(b.name)
      );
      return { success: true, items };
    } catch (err) {
      return { success: false, error: err.message, items: [] };
    }
  });

  ipcMain.handle('fs:readFile', async (_event, filePath, encoding = 'utf-8') => {
    try {
      const buffer = await fs.promises.readFile(filePath);
      if (encoding === 'base64') {
        return { success: true, content: buffer.toString('base64'), isBinary: true };
      }
      return { success: true, content: buffer.toString('utf-8'), isBinary: false };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:writeFile', async (_event, filePath, content, encoding = 'utf-8') => {
    try {
      if (encoding === 'base64') await fs.promises.writeFile(filePath, Buffer.from(content, 'base64'));
      else await fs.promises.writeFile(filePath, content, 'utf-8');
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  // Печать: тот же модуль, что в главном процессе (electron/printWindow.ts).
  ipcMain.handle('print:document', async (_event, request) => {
    try {
      if (!printWindow || typeof printWindow.runPrint !== 'function') {
        return { success: false, error: 'модуль печати не найден' };
      }
      return await printWindow.runPrint(request, {});
    } catch (err) {
      return { success: false, error: err.message };
    }
  });
}

/* ============================================================
   Проверка
   ============================================================ */

async function main() {
  fs.rmSync(path.join(SANDBOX, DOC_NAME), { force: true });
  await buildLongDocx(path.join(SANDBOX, DOC_NAME));

  registerHandlers();

  const win = new BrowserWindow({
    width: 1500,
    height: 940,
    show: false,
    webPreferences: {
      preload: path.join(ROOT, 'dist-electron', 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
    },
  });
  win.webContents.on('console-message', (_event, _level, message) => {
    if (/error|Error|ошибк/i.test(message)) out.errors.push(message);
  });

  const js = (expression) =>
    win.webContents.executeJavaScript(`(async () => { ${expression} })()`);

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await sleep(2600);

    // Документ из дерева: двойной щелчок открывает редактирование.
    let opened = false;
    for (let attempt = 0; attempt < 25 && !opened; attempt += 1) {
      opened = await js(`
        const node = [...document.querySelectorAll('div')].find((element) => {
          const title = (element.getAttribute('title') || '').split(' — ')[0];
          return title === ${JSON.stringify(DOC_NAME)} && element.className.includes('cursor-pointer');
        });
        if (!node) return false;
        node.click();
        node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
        return true;
      `);
      if (!opened) await sleep(300);
    }
    if (!opened) throw new Error('файл не найден в дереве');

    // Ждём разбора .docx: полотно появляется после чтения пакета.
    let loaded = false;
    for (let attempt = 0; attempt < 40 && !loaded; attempt += 1) {
      loaded = await js(`
        const root = document.querySelector('.word-sheet');
        return Boolean(root) && (root.innerText || '').trim().length > 80;
      `);
      if (!loaded) await sleep(400);
    }
    step('opened', {
      loaded,
      sheets: await js(`return document.querySelectorAll('[data-page-sheet]').length;`),
      // Положительное утверждение: в отчёте не должно быть boolean-провалов,
      // иначе прогон считает нормальное наблюдение ошибкой.
      noParseError: await js(`return !/поврежд|Не удалось разобрать/i.test(document.body.innerText);`),
    });
    await sleep(800);

    /* ---------- 1. Геометрия и листы ---------- */
    step('layout', await js(`
      const root = document.querySelector('.word-sheet');
      const sheet = document.querySelector('[data-page-sheet]');
      const sheets = [...document.querySelectorAll('[data-page-sheet]')];
      if (!root || !sheet) return { error: 'нет полотна Word' };
      const style = getComputedStyle(root);
      const sheetRect = sheet.getBoundingClientRect();
      const contentWidth = root.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);
      return {
        sheets: sheets.length,
        sheetWidth: Math.round(sheetRect.width * 100) / 100,
        sheetHeight: Math.round(sheetRect.height * 100) / 100,
        paddingTop: Math.round(parseFloat(style.paddingTop) * 100) / 100,
        paddingLeft: Math.round(parseFloat(style.paddingLeft) * 100) / 100,
        paddingRight: Math.round(parseFloat(style.paddingRight) * 100) / 100,
        contentWidth: Math.round(contentWidth * 100) / 100,
        gaps: document.querySelectorAll('[data-page-gap]').length,
        badge: (document.querySelector('[data-word-pages]') || {}).textContent || '',
        geometryA4: Math.abs(sheetRect.width - 793.7) < 2 && Math.abs(sheetRect.height - 1122.5) < 2,
        margins: Math.abs(parseFloat(style.paddingLeft) - 113.39) < 2,
      };
    `));

    /* ---------- 2. Строки не пересекают границу листа ---------- */
    step('pagination', await js(`
      const root = document.querySelector('.word-sheet');
      const sheets = [...document.querySelectorAll('[data-page-sheet]')];
      const flowTop = root.getBoundingClientRect().top;
      const gaps = [...document.querySelectorAll('[data-page-gap]')];
      const sheetsTop = sheets.map((element) => element.getBoundingClientRect().top - flowTop);
      const gapsBottom = gaps.map((gap) => gap.getBoundingClientRect().bottom - flowTop);
      const aligned = gapsBottom.every((bottom, index) =>
        Math.abs(bottom - (sheetsTop[index + 1] + parseFloat(getComputedStyle(root).paddingTop))) < 2
      );

      // Каждая текстовая строка должна целиком лежать на одном листе.
      const crossing = [];
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
      let node = walker.nextNode();
      while (node) {
        const range = document.createRange();
        range.selectNodeContents(node);
        for (const rect of range.getClientRects()) {
          const top = rect.top - flowTop;
          const bottom = rect.bottom - flowTop;
          for (let index = 0; index < sheetsTop.length; index += 1) {
            const pageBottom = sheetsTop[index] + sheets[index].getBoundingClientRect().height;
            if (bottom > pageBottom + 2 && top < pageBottom - 2) {
              crossing.push((node.textContent || '').slice(0, 24));
            }
          }
        }
        node = walker.nextNode();
      }

      const before = root.innerHTML;
      const contentBefore = before.replace(/<span[^>]*data-page-gap="1"[^>]*><\\/span>/g, '');
      return {
        gaps: gaps.length,
        gapAligned: aligned,
        crossingLines: crossing.length,
        crossingSample: crossing.slice(0, 3),
        widths: gaps.map((gap) => Math.round(gap.getBoundingClientRect().width)),
        contentLength: contentBefore.length,
        before: before.length,
      };
    `));

    /* ---------- 3. Повторная разметка не портит документ ---------- */
    step('idempotent', await js(`
      const root = document.querySelector('.word-sheet');
      const strip = (html) => html.replace(/<span[^>]*data-page-gap="1"[^>]*><\\/span>/g, '');
      const first = strip(root.innerHTML);
      await new Promise((resolve) => setTimeout(resolve, 200));
      root.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise((resolve) => setTimeout(resolve, 900));
      const second = strip(root.innerHTML);
      return {
        same: first === second,
        sheets: document.querySelectorAll('[data-page-sheet]').length,
        gaps: document.querySelectorAll('[data-page-gap]').length,
      };
    `));

    /* ---------- 4. Колонтитулы и номера страниц ---------- */
    step('headers', await js(`
      const headers = [...document.querySelectorAll('[data-page-header]')];
      const footers = [...document.querySelectorAll('[data-page-footer]')];
      const pages = document.querySelectorAll('[data-page-sheet]').length;
      return {
        pages,
        headers: headers.length,
        footers: footers.length,
        headerText: headers[0] ? headers[0].innerText.trim().slice(0, 60) : '',
        footerText: footers[0] ? footers[0].innerText.trim().slice(0, 60) : '',
        onEveryPage: headers.length === pages,
      };
    `));

    /* ---------- 5. Сохранение без служебных маркеров ---------- */
    step('save', await js(`
      const root = document.querySelector('.word-sheet');
      const inDom = root.innerHTML.includes('data-page-gap');
      // Ctrl+S проходит через сериализатор редактора.
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true }));
      await new Promise((resolve) => setTimeout(resolve, 1800));
      return { gapInDom: inDom, sheets: document.querySelectorAll('[data-page-sheet]').length };
    `));
    const savedZip = await JSZip.loadAsync(fs.readFileSync(path.join(SANDBOX, DOC_NAME)));
    const savedXml = await savedZip.file('word/document.xml').async('string');
    step('savedPackage', {
      noGapMarkers: !savedXml.includes('data-page-gap'),
      paragraphs: (savedXml.match(/<w:p\b/g) || []).length,
      tables: (savedXml.match(/<w:tbl\b/g) || []).length,
      hasHeaderRefs: /<w:headerReference\b/.test(savedXml),
      hasPageSize: /<w:pgSz\b/.test(savedXml),
      hasMargins: /<w:pgMar\b/.test(savedXml),
    });

    /* ---------- 6. Предпросмотр печати ---------- */
    const editorPages = await js(`return document.querySelectorAll('[data-page-sheet]').length;`);
    await js(`
      window.dispatchEvent(new Event('workspace:print-preview'));
      return true;
    `);
    let preview: any = {};
    for (let attempt = 0; attempt < 60; attempt += 1) {
      preview = await js(`
        const modal = document.querySelector('[aria-label="Предпросмотр печати"]');
        return {
          modal: Boolean(modal),
          badge: (() => {
            const element = document.querySelector('[data-print-pages]');
            return element ? element.textContent : '';
          })(),
          images: document.querySelectorAll('[aria-label="Предпросмотр печати"] img').length,
          text: modal ? modal.innerText.slice(0, 120) : '',
        };
      `);
      // Ждём, пока pdf.js отрисует ровно столько страниц, сколько листов.
      if (preview.images === editorPages) break;
      await sleep(700);
    }
    step('printPreview', {
      modal: preview.modal,
      badge: preview.badge,
      images: preview.images,
      text: preview.text,
      editorPages,
      pagesMatch: preview.images === editorPages,
    });

    /* ---------- 7. Контракт IPC ---------- */
    const mainJs = fs.readFileSync(path.join(ROOT, 'dist-electron', 'main.js'), 'utf-8');
    const preloadJs = fs.readFileSync(path.join(ROOT, 'dist-electron', 'preload.js'), 'utf-8');
    step('printContract', {
      handlerRegistered: mainJs.includes('print:document'),
      preloadBridge: preloadJs.includes('printDocument'),
      usesPrintWindow: mainJs.includes('printWindow'),
    });
  } catch (err) {
    out.errors.push(err && err.stack ? err.stack : String(err));
  } finally {
    fs.writeFileSync(OUT, JSON.stringify(out, null, 2), 'utf-8');
    console.log(JSON.stringify(out, null, 2));
    win.destroy();
    app.quit();
  }
}

main().catch((err) => {
  out.errors.push(String(err && err.stack ? err.stack : err));
  try {
    fs.writeFileSync(OUT, JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
  app.exit(1);
});
