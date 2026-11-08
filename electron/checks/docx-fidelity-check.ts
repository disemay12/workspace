/**
 * Проверка сохранности .docx при сохранении через редактор.
 *
 * Сценарий:
 *  1. Собирает пакет с headers/footers, styles, numbering, footnotes, comments,
 *     settings, картинкой, content control и разрывом секции.
 *  2. Открывает документ в приложении и сохраняет без правок: файл должен
 *     остаться байт в байт.
 *  3. Правит один абзац и сохраняет: меняется только word/document.xml,
 *     остальные части пакета и не поддерживаемые редактором объекты intact.
 *
 * Запуск: electron.exe --no-sandbox dist-electron/docx-fidelity-check.js
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const JSZip = require('jszip');
const { registerCommonCheckIpc } = require('./common-ipc');

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
const SANDBOX = process.env.WS_CHECK_SANDBOX || path.join(ROOT, 'test-workspace-docx');
const FIXTURES = path.join(SANDBOX, 'fixtures');
const OUT = path.join(ROOT, 'docx-fidelity-check.json');
const DOC_NAME = 'rich.docx';
const results: any = { steps: [], errors: [] };
const step = (name, value) => results.steps.push({ name, ...value });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

/** 1×1 PNG: проверяем, что media не пересобирается и ссылка не меняется. */
const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lK3hWQAAAABJRU5ErkJggg==';
/** Другой PNG: проверяем, что новое изображение получает свою media-часть. */
const INSERTED_PNG_DATA_URL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

function contentTypesXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Types xmlns="${CT_NS}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '</Types>'
  );
}

function documentRelsXml() {
  const rel = (id, type, target) =>
    `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Relationships xmlns="${REL_NS}">` +
    rel('rIdStyles', 'styles', 'styles.xml') +
    rel('rIdNumbering', 'numbering', 'numbering.xml') +
    rel('rIdHeader', 'header', 'header1.xml') +
    rel('rIdFooter', 'footer', 'footer1.xml') +
    rel('rIdFootnotes', 'footnotes', 'footnotes.xml') +
    rel('rIdComments', 'comments', 'comments.xml') +
    rel('rIdSettings', 'settings', 'settings.xml') +
    rel('rIdImage', 'image', 'media/image1.png') +
    '</Relationships>'
  );
}

function stylesXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:styles xmlns:w="${W_NS}">` +
    '<w:docDefaults>' +
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
    '</w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:after="160" w:line="240" w:lineRule="auto"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>' +
    '<w:rPr><w:b/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="MyCustomStyle"><w:name w:val="MyCustomStyle"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:rPr><w:b/><w:color w:val="112233"/></w:rPr></w:style>' +
    '</w:styles>'
  );
}

function numberingXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:numbering xmlns:w="${W_NS}">` +
    '<w:abstractNum w:abstractNumId="10"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/>' +
    '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num>' +
    '</w:numbering>'
  );
}

function headerXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:p><w:r><w:t>КОЛОНТИТУЛ ПРОВЕРКИ</w:t></w:r></w:p></w:hdr>`
  );
}

function footerXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:p><w:r><w:t>СТРАНИЦА ПРОВЕРКИ</w:t></w:r></w:p></w:ftr>`
  );
}

function footnotesXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:footnotes xmlns:w="${W_NS}"><w:footnote w:id="1"><w:p><w:r><w:t>Сноска проверки</w:t></w:r></w:p></w:footnote></w:footnotes>`
  );
}

function commentsXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:comments xmlns:w="${W_NS}"><w:comment w:id="1" w:author="Проверка"><w:p><w:r><w:t>Комментарий проверки</w:t></w:r></w:p></w:comment></w:comments>`
  );
}

function settingsXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:settings xmlns:w="${W_NS}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/></w:settings>`
  );
}

function documentXml() {
  const sectProps =
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"/>';
  const imageRun =
    '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Picture 1"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

  const body =
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Отчёт о проверке</w:t></w:r></w:p>' +
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">TOC \\o "1-3" \\h \\z \\u</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Оглавление</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
    '<w:p><w:r><w:t>Первый абзац с текстом.</w:t></w:r></w:p>' +
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Жирный фрагмент</w:t></w:r><w:r><w:t xml:space="preserve"> и </w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="C00000"/></w:rPr><w:t>курсив красным</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Пункт списка</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="A6A6A6"/><w:left w:val="single" w:sz="4" w:color="A6A6A6"/><w:bottom w:val="single" w:sz="4" w:color="A6A6A6"/><w:right w:val="single" w:sz="4" w:color="A6A6A6"/></w:tblBorders></w:tblPr>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr><w:p><w:r><w:t>Ячейка A</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr><w:p><w:r><w:t>Ячейка B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
    '<w:sdt><w:sdtPr><w:alias w:val="Поле проверки"/><w:id w:val="42"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Содержимое content control</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
    '<w:p>' + imageRun + '</w:p>' +
    '<w:p><w:pPr><w:sectPr>' + sectProps + '</w:sectPr></w:pPr><w:r><w:t>Конец первой секции</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Последний абзац.</w:t></w:r></w:p>' +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/>' +
    sectProps +
    '</w:sectPr>';

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}">` +
    `<w:body>${body}</w:body></w:document>`
  );
}

async function buildFixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypesXml());
  zip.folder('_rels').file('.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  const word = zip.folder('word');
  word.file('document.xml', documentXml());
  word.file('styles.xml', stylesXml());
  word.file('numbering.xml', numberingXml());
  word.file('header1.xml', headerXml());
  word.file('footer1.xml', footerXml());
  word.file('footnotes.xml', footnotesXml());
  word.file('comments.xml', commentsXml());
  word.file('settings.xml', settingsXml());
  word.file('media/image1.png', Buffer.from(PNG_BASE64, 'base64'));
  word.folder('_rels').file('document.xml.rels', documentRelsXml());
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/** Разворачивает пакет в карту часть → содержимое (строки и base64 для media). */
async function readPackage(filePath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
  const parts = {};
  for (const name of Object.keys(zip.files)) {
    const file = zip.files[name];
    if (file.dir) continue;
    parts[name] = /\.(png|jpe?g|gif|webp|bmp)$/i.test(name)
      ? await file.async('base64')
      : await file.async('string');
  }
  return parts;
}

function sha256(filePath) {
  return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
}

function registerIpc() {
  registerCommonCheckIpc({ sandbox: SANDBOX });
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
        try { size = fs.statSync(full).size; } catch {}
        return {
          name: entry.name,
          path: full,
          isDirectory: entry.isDirectory(),
          size,
          mtime: null,
          extension: entry.isDirectory() ? '' : path.extname(entry.name).toLowerCase().replace('.', ''),
        };
      });
      return { success: true, items };
    } catch (err) {
      return { success: false, error: String(err), items: [] };
    }
  });
  ipcMain.handle('fs:readFile', async (_e, filePath, encoding = 'utf-8') => {
    try {
      const buffer = await fs.promises.readFile(filePath);
      if (encoding === 'base64') return { success: true, content: buffer.toString('base64'), isBinary: true };
      return { success: true, content: buffer.toString('utf-8'), isBinary: false };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });
  ipcMain.handle('fs:writeFile', async (_e, filePath, content, encoding = 'utf-8') => {
    try {
      const tmp = `${filePath}.ws-tmp`;
      await fs.promises.writeFile(tmp, encoding === 'base64' ? Buffer.from(content, 'base64') : content, encoding === 'base64' ? undefined : 'utf-8');
      await fs.promises.rename(tmp, filePath).catch(async () => {
        await fs.promises.copyFile(tmp, filePath);
        await fs.promises.unlink(tmp).catch(() => {});
      });
      return { success: true };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });
  ipcMain.handle('fs:createFile', async (_e, filePath, content = '') => {
    try { await fs.promises.writeFile(filePath, content, { flag: 'wx' }); return { success: true }; }
    catch (err) { return { success: false, error: String(err) }; }
  });
  ipcMain.handle('fs:createDirectory', async () => ({ success: true }));
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));
  ipcMain.handle('fs:searchWorkspace', async () => ({ success: true, results: [] }));
  ipcMain.handle('ai:getSettings', () => ({ configured: false }));
  ipcMain.handle('app:getVersion', () => '1.0.0');
}

function sameEntries(before, after, ignore = []) {
  const names = new Set([...Object.keys(before), ...Object.keys(after)]);
  const changed = [];
  for (const name of names) {
    if (ignore.includes(name)) continue;
    if (before[name] !== after[name]) changed.push(name);
  }
  return changed;
}

async function main() {
  fs.mkdirSync(FIXTURES, { recursive: true });
  const target = path.join(SANDBOX, DOC_NAME);
  fs.writeFileSync(target, await buildFixture());
  fs.writeFileSync(path.join(FIXTURES, DOC_NAME), fs.readFileSync(target));

  const original = await readPackage(target);
  results.fixture = { path: target, parts: Object.keys(original) };

  registerIpc();

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
    if (/error|Error|ошибк/i.test(message)) results.errors.push(message);
  });

  const js = (body) => win.webContents.executeJavaScript(`(async () => { ${body} })()`);
  const clickFile = () => js(`
    const nodes = [...document.querySelectorAll('div')]
      .filter((d) => d.textContent?.trim() === ${JSON.stringify(DOC_NAME)});
    if (!nodes.length) return false;
    const clickable = nodes.find((d) => d.className.includes('cursor-pointer')) || nodes[0];
    clickable.click();
    return true;
  `);

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'), { query: { modelTest: '1' } });
    await sleep(2500);
    await clickFile();
    await sleep(3000);

    const opened = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      return {
        hasCanvas: !!canvas,
        text: canvas ? canvas.innerText : '',
        html: canvas ? canvas.innerHTML : '',
        readonlyBlocks: canvas ? canvas.querySelectorAll('[data-docx-readonly]').length : 0,
        docxBlocks: canvas ? canvas.querySelectorAll('[data-docx-id]').length : 0,
      };
    `);
    step('open', {
      hasCanvas: opened.hasCanvas,
      hasHeading: opened.text.includes('Отчёт о проверке'),
      hasTable: opened.text.includes('Ячейка A'),
      hasReadonly: opened.readonlyBlocks === 1,
      boundBlocks: opened.docxBlocks >= 6,
    });

    /* Служебные части пакета: колонтитулы, сноски, комментарии, стили, секции. */
    const fixtureBase64 = fs.readFileSync(target).toString('base64');
    const packageInfo = await win.webContents.executeJavaScript(`
      (async () => {
        const mod = window.__docx;
        if (!mod) return null;
        return await mod.docxPackageInfo(${JSON.stringify(fixtureBase64)});
      })()
    `);
    step('packageInfo', {
      available: Boolean(packageInfo),
      header: Boolean(packageInfo?.headers?.some((part) => part.text.includes('КОЛОНТИТУЛ'))),
      footer: Boolean(packageInfo?.footers?.some((part) => part.text.includes('СТРАНИЦА'))),
      footnotes: Boolean(packageInfo?.footnotes?.includes('Сноска проверки')),
      comments: Boolean(packageInfo?.comments?.includes('Комментарий проверки')),
      customStyle: Boolean(packageInfo?.styleNames?.includes('MyCustomStyle')),
      sections: packageInfo?.sections >= 2,
      page: packageInfo?.page?.widthMm === 210,
    });

    /* Писатель абзаца: интервалы и запреты разрыва попадают в OOXML. */
    const formattingHtml =
      '<p data-widow="off" style="line-height:1.5;margin-top:6pt;margin-bottom:4pt;break-after:avoid;break-before:page;background-color:#FFFF00;border:2pt dashed #000000">Абзац</p>';
    const tocHtml = '<p data-field="TOC">Оглавление</p>';
    const writerDocs = await win.webContents.executeJavaScript(`
      (async () => {
        const mod = window.__docx;
        if (!mod) return null;
        window.__customStyles?.registerCustomStyle({
          id: 'CustomFidelity',
          name: 'Врезка',
          font: 'Georgia',
          sizePt: 14,
          bold: true,
          color: '#112233',
          align: 'center',
        });
        return {
          format: await mod.htmlToDocxBase64(${JSON.stringify(formattingHtml)}),
          toc: await mod.htmlToDocxBase64(${JSON.stringify(tocHtml)}),
          link: await mod.htmlToDocxBase64('<p><a href="https://example.com">Сайт</a></p>'),
          section: await mod.htmlToDocxBase64('<p data-section-break="1"><br></p>'),
          tabs: await mod.htmlToDocxBase64('<p data-tabs="24,48">X</p>'),
          nested: await mod.htmlToDocxBase64(
            '<ul><li>Один<ul><li>Два</li></ul></li></ul>'
          ),
          nestedHtml: (
            await mod.docxBase64ToHtml(
              await mod.htmlToDocxBase64(
                '<ul><li>Один<ul><li>Два</li></ul></li></ul>'
              )
            )
          ).html,
          style: await mod.htmlToDocxBase64('<p data-style="Title">X</p>'),
          tableWidth: await mod.htmlToDocxBase64(
            '<table><tr><td style="width:100pt">A</td></tr></table>'
          ),
          tableBorder: await mod.htmlToDocxBase64(
            '<table data-table-border="2|dashed|#FF0000"><tr><td>A</td></tr></table>'
          ),
          pageField: await mod.htmlToDocxBase64('<p><span data-field="PAGE">1</span></p>'),
          bookmark: await mod.htmlToDocxBase64(
            '<p><span data-bookmark="Метка">Текст</span></p>'
          ),
          customStyle: await mod.htmlToDocxBase64(
            '<p data-style="CustomFidelity">Текст стиля</p>'
          ),
          customStyleHtml: (
            await mod.docxBase64ToHtml(
              await mod.htmlToDocxBase64(
                '<p data-style="CustomFidelity">Текст стиля</p>'
              )
            )
          ).html,
          annotations: await mod.htmlToDocxBase64(
            '<p><span data-comment-id="1" data-comment-text="Проверить" data-comment-author="AI">Текст</span> и ' +
              '<span data-footnote-id="1" data-footnote-text="Источник">сноска<sup data-footnote-ref="1">1</sup></span></p>'
          ),
          annotationsHtml: (
            await mod.docxBase64ToHtml(
              await mod.htmlToDocxBase64(
                '<p><span data-comment-id="1" data-comment-text="Проверить" data-comment-author="AI">Текст</span> и ' +
                  '<span data-footnote-id="1" data-footnote-text="Источник">сноска<sup data-footnote-ref="1">1</sup></span></p>'
              )
            )
          ).html,
          tocRefreshed: await mod.htmlToDocxBase64(
            '<p data-field="TOC">Глава\\nРаздел</p>'
          ),
          runIdsHtml: (
            await mod.docxBase64ToHtml(
              await mod.htmlToDocxBase64('<p>Первый <strong>второй</strong></p>')
            )
          ).html,
        };
      })()
    `);
    let paragraphXml = '';
    let tocXml = '';
    let linkXml = '';
    let sectionXml = '';
    let tabsXml = '';
    let nestedXml = '';
    let nestedHtml = '';
    let styleXml = '';
    let tableWidthXml = '';
    let tableBorderXml = '';
    let fieldXml = '';
    let bookmarkXml = '';
    let customStyleXml = '';
    let customStylesXml = '';
    let customStyleHtml = '';
    let annotationsXml = '';
    let commentsPartXml = '';
    let footnotesPartXml = '';
    let annotationsHtml = '';
    let tocRefreshedXml = '';
    let runIdsHtml = '';
    if (writerDocs) {
      const formatZip = await JSZip.loadAsync(Buffer.from(writerDocs.format, 'base64'));
      paragraphXml = (await formatZip.file('word/document.xml')?.async('string')) ?? '';
      const tocZip = await JSZip.loadAsync(Buffer.from(writerDocs.toc, 'base64'));
      tocXml = (await tocZip.file('word/document.xml')?.async('string')) ?? '';
      const linkZip = await JSZip.loadAsync(Buffer.from(writerDocs.link, 'base64'));
      linkXml = (await linkZip.file('word/document.xml')?.async('string')) ?? '';
      const sectionZip = await JSZip.loadAsync(Buffer.from(writerDocs.section, 'base64'));
      sectionXml = (await sectionZip.file('word/document.xml')?.async('string')) ?? '';
      const tabsZip = await JSZip.loadAsync(Buffer.from(writerDocs.tabs, 'base64'));
      tabsXml = (await tabsZip.file('word/document.xml')?.async('string')) ?? '';
      const nestedZip = await JSZip.loadAsync(Buffer.from(writerDocs.nested, 'base64'));
      nestedXml = (await nestedZip.file('word/document.xml')?.async('string')) ?? '';
      nestedHtml = String(writerDocs.nestedHtml ?? '');
      const styleZip = await JSZip.loadAsync(Buffer.from(writerDocs.style, 'base64'));
      styleXml = (await styleZip.file('word/document.xml')?.async('string')) ?? '';
      const widthZip = await JSZip.loadAsync(Buffer.from(writerDocs.tableWidth, 'base64'));
      tableWidthXml = (await widthZip.file('word/document.xml')?.async('string')) ?? '';
      const borderZip = await JSZip.loadAsync(Buffer.from(writerDocs.tableBorder, 'base64'));
      tableBorderXml = (await borderZip.file('word/document.xml')?.async('string')) ?? '';
      const fieldZip = await JSZip.loadAsync(Buffer.from(writerDocs.pageField, 'base64'));
      fieldXml = (await fieldZip.file('word/document.xml')?.async('string')) ?? '';
      const bookmarkZip = await JSZip.loadAsync(Buffer.from(writerDocs.bookmark, 'base64'));
      bookmarkXml = (await bookmarkZip.file('word/document.xml')?.async('string')) ?? '';
      const customStyleZip = await JSZip.loadAsync(
        Buffer.from(writerDocs.customStyle, 'base64')
      );
      customStyleXml =
        (await customStyleZip.file('word/document.xml')?.async('string')) ?? '';
      customStylesXml =
        (await customStyleZip.file('word/styles.xml')?.async('string')) ?? '';
      customStyleHtml = String(writerDocs.customStyleHtml ?? '');
      const annotationsZip = await JSZip.loadAsync(
        Buffer.from(writerDocs.annotations, 'base64')
      );
      annotationsXml =
        (await annotationsZip.file('word/document.xml')?.async('string')) ?? '';
      commentsPartXml =
        (await annotationsZip.file('word/comments.xml')?.async('string')) ?? '';
      footnotesPartXml =
        (await annotationsZip.file('word/footnotes.xml')?.async('string')) ?? '';
      annotationsHtml = String(writerDocs.annotationsHtml ?? '');
      const tocRefreshedZip = await JSZip.loadAsync(
        Buffer.from(writerDocs.tocRefreshed, 'base64')
      );
      tocRefreshedXml =
        (await tocRefreshedZip.file('word/document.xml')?.async('string')) ?? '';
      runIdsHtml = String(writerDocs.runIdsHtml ?? '');
    }
    step('paragraphWriter', {
      available: Boolean(writerDocs),
      line: paragraphXml.includes('w:line="360"'),
      before: paragraphXml.includes('w:before="120"'),
      after: paragraphXml.includes('w:after="80"'),
      keepNext: paragraphXml.includes('<w:keepNext/>'),
      pageBreakBefore: paragraphXml.includes('<w:pageBreakBefore/>'),
      widowControl: paragraphXml.includes('<w:widowControl w:val="0"/>'),
      shading: paragraphXml.includes('<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'),
      border:
        paragraphXml.includes('<w:pBdr>') &&
        paragraphXml.includes('w:val="dashed"') &&
        paragraphXml.includes('w:sz="16"') &&
        paragraphXml.includes('w:color="000000"'),
      tocField: /TOC \\o "1-3"/.test(tocXml) && tocXml.includes('w:fldChar'),
      hyperlink:
        linkXml.includes('HYPERLINK') &&
        linkXml.includes('https://example.com') &&
        linkXml.includes('<w:fldSimple'),
      sectionBreak: /<w:pPr>[^]*?<w:sectPr>/.test(sectionXml),
      tabs:
        tabsXml.includes('<w:tabs>') &&
        tabsXml.includes('w:pos="480"') &&
        tabsXml.includes('w:pos="960"'),
      nestedList:
        nestedXml.includes('<w:ilvl w:val="0"/>') &&
        nestedXml.includes('<w:ilvl w:val="1"/>') &&
        /Один[\s\S]*?<ul>[\s\S]*?Два/.test(nestedHtml),
      namedStyle: styleXml.includes('w:pStyle w:val="Title"'),
      tableWidth: tableWidthXml.includes('<w:tcW w:w="2000" w:type="dxa"/>'),
      tableBorder:
        tableBorderXml.includes('w:val="dashed"') &&
        tableBorderXml.includes('w:color="FF0000"'),
      pageField: fieldXml.includes('<w:fldSimple w:instr=" PAGE ">'),
      bookmark:
        bookmarkXml.includes('<w:bookmarkStart') &&
        bookmarkXml.includes('w:name="Метка"') &&
        bookmarkXml.includes('<w:bookmarkEnd'),
      customStyle:
        customStylesXml.includes('w:styleId="CustomFidelity"') &&
        customStylesXml.includes('w:customStyle="1"') &&
        customStylesXml.includes('w:name w:val="Врезка"') &&
        customStyleXml.includes('w:pStyle w:val="CustomFidelity"') &&
        /font-family:\s*&quot;?Georgia&quot;?|font-family:\s*"?Georgia"?/.test(customStyleHtml),
      annotations:
        annotationsXml.includes('<w:commentRangeStart w:id="1"/>') &&
        annotationsXml.includes('<w:commentRangeEnd w:id="1"/>') &&
        annotationsXml.includes('<w:commentReference w:id="1"/>') &&
        annotationsXml.includes('<w:footnoteReference w:id="1"/>') &&
        commentsPartXml.includes('Проверить') &&
        footnotesPartXml.includes('Источник') &&
        /data-comment-text="Проверить"/.test(annotationsHtml) &&
        /data-footnote-id="1"/.test(annotationsHtml),
      tocRefreshed:
        tocRefreshedXml.includes('TOC \\o "1-3"') &&
        tocRefreshedXml.includes('Глава') &&
        tocRefreshedXml.includes('Раздел') &&
        tocRefreshedXml.includes('<w:br/>'),
      runIds:
        runIdsHtml.includes('data-run-id="r1"') &&
        runIdsHtml.includes('data-run-id="r2"') &&
        runIdsHtml.includes('<strong>второй</strong>'),
    });

    /* Масштаб: 400 абзацев пишутся и читаются за разумное время. */
    const largeResult = await win.webContents.executeJavaScript(`
      (async () => {
        const mod = window.__docx;
        if (!mod) return null;
        const html = Array.from({ length: 400 }, (_, index) =>
          '<p>Пункт ' + index + ' большого документа</p>'
        ).join('');
        const writeStart = performance.now();
        const base64 = await mod.htmlToDocxBase64(html);
        const writeMs = Math.round(performance.now() - writeStart);
        const readStart = performance.now();
        const parsed = await mod.docxBase64ToHtml(base64);
        const readMs = Math.round(performance.now() - readStart);
        return {
          written: Boolean(base64),
          writeMs,
          readMs,
          blocks: (parsed.html.match(/<p\\b/g) || []).length,
          lastKept: parsed.html.includes('Пункт 399'),
        };
      })()
    `);
    step('largeDocumentPerf', {
      written: Boolean(largeResult && largeResult.written),
      blocks: Number(largeResult && largeResult.blocks) >= 400,
      lastKept: Boolean(largeResult && largeResult.lastKept),
      withinBudget: Boolean(
        largeResult && largeResult.writeMs < 8000 && largeResult.readMs < 8000
      ),
    });
    (results as any).metrics = {
      largeDocument: {
        writeMs: largeResult ? largeResult.writeMs : null,
        readMs: largeResult ? largeResult.readMs : null,
        blocks: largeResult ? largeResult.blocks : 0,
      },
    };

    /* Колонтитулы: правка существующего, создание нового и удаление. */
    const headerFooterResult = await win.webContents.executeJavaScript(`
      (async () => {
        const mod = window.__docx;
        if (!mod) return null;
        const source = ${JSON.stringify(fixtureBase64)};
        const edited = await mod.docxWithHeaderFooterEdits(source, [
          { kind: 'header', referenceType: 'default', text: 'НОВЫЙ ВЕРХНИЙ' },
        ]);
        const created = await mod.docxWithHeaderFooterEdits(source, [
          { kind: 'footer', referenceType: 'first', text: 'ПЕРВАЯ СТРАНИЦА' },
        ]);
        const removed = await mod.docxWithHeaderFooterEdits(source, [
          { kind: 'footer', referenceType: 'default', remove: true },
        ]);
        const even = await mod.docxWithHeaderFooterEdits(source, [
          { kind: 'header', referenceType: 'even', text: 'ЧЁТНЫЕ СТРАНИЦЫ' },
        ]);
        const fields = await mod.docxWithHeaderFooterEdits(source, [
          { kind: 'footer', referenceType: 'even', text: 'Стр. {{page}} из {{pages}}' },
        ]);
        return { edited, created, removed, even, fields };
      })()
    `);

    let headerEditOk = false;
    let footerCreateOk = false;
    let footerRemoveOk = false;
    let evenHeaderOk = false;
    let pageFieldsOk = false;
    if (headerFooterResult) {
      const editedZip = await JSZip.loadAsync(Buffer.from(headerFooterResult.edited, 'base64'));
      const editedHeader = (await editedZip.file('word/header1.xml')?.async('string')) ?? '';
      headerEditOk = editedHeader.includes('НОВЫЙ ВЕРХНИЙ');

      const createdZip = await JSZip.loadAsync(Buffer.from(headerFooterResult.created, 'base64'));
      const createdDoc = (await createdZip.file('word/document.xml')?.async('string')) ?? '';
      const createdRels =
        (await createdZip.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
      const createdTypes =
        (await createdZip.file('[Content_Types].xml')?.async('string')) ?? '';
      footerCreateOk =
        Boolean(createdZip.file('word/footer2.xml')) &&
        createdDoc.includes('w:footerReference w:type="first"') &&
        createdDoc.includes('<w:titlePg/>') &&
        createdRels.includes('footer2.xml') &&
        createdTypes.includes('/word/footer2.xml');

      const removedZip = await JSZip.loadAsync(Buffer.from(headerFooterResult.removed, 'base64'));
      const removedDoc = (await removedZip.file('word/document.xml')?.async('string')) ?? '';
      const removedRels =
        (await removedZip.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
      footerRemoveOk =
        !removedZip.file('word/footer1.xml') &&
        !/w:footerReference[^>]*w:type="default"/.test(removedDoc) &&
        !removedRels.includes('footer1.xml');

      const evenZip = await JSZip.loadAsync(Buffer.from(headerFooterResult.even, 'base64'));
      const evenDoc = (await evenZip.file('word/document.xml')?.async('string')) ?? '';
      const evenSettings =
        (await evenZip.file('word/settings.xml')?.async('string')) ?? '';
      evenHeaderOk =
        evenDoc.includes('w:headerReference w:type="even"') &&
        evenSettings.includes('<w:evenAndOddHeaders/>');

      const fieldsZip = await JSZip.loadAsync(Buffer.from(headerFooterResult.fields, 'base64'));
      const fieldsFooterNames = Object.keys(fieldsZip.files).filter((name) =>
        /^word\/footer\d+\.xml$/.test(name)
      );
      const fieldsFooters = await Promise.all(
        fieldsFooterNames.map(async (name) =>
          (await fieldsZip.file(name)?.async('string')) ?? ''
        )
      );
      pageFieldsOk = fieldsFooters.some(
        (xml) =>
          xml.includes(' PAGE ') &&
          xml.includes(' NUMPAGES ') &&
          xml.includes('w:instrText')
      );
    }
    step('headerFooterWriter', {
      headerEdited: headerEditOk,
      footerCreated: footerCreateOk,
      footerRemoved: footerRemoveOk,
      evenHeader: evenHeaderOk,
      pageFields: pageFieldsOk,
    });

    // 1. Сохранение без правок возвращает исходный пакет байт в байт.
    const hashBefore = sha256(target);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    const hashAfter = sha256(target);
    step('goldenSave', { sameFile: hashBefore === hashAfter });

    // 2. Правим один абзац и сохраняем.
    const typed = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      if (!canvas) return false;
      const paragraphs = [...canvas.querySelectorAll(':scope > p')];
      const target = paragraphs.find((p) => p.textContent.includes('Первый абзац'));
      if (!target) return false;
      target.appendChild(document.createTextNode(' ПРАВКА_СОХРАНЕНИЯ'));
      canvas.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    `);
    step('edit', { typed });
    await sleep(1800);

    // Вставляем картинку в тот же абзац: проверяем, что она попадёт в пакет.
    const imageTyped = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      if (!canvas) return false;
      const target = [...canvas.querySelectorAll(':scope > p')].find((p) =>
        p.textContent.includes('Первый абзац')
      );
      if (!target) return false;
      const image = document.createElement('img');
      image.src = '${INSERTED_PNG_DATA_URL}';
      image.style.width = '40px';
      image.style.height = '40px';
      target.appendChild(image);
      canvas.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    `);
    step('imageEdit', { inserted: imageTyped });
    await sleep(1800);

    // Разрыв страницы: в .docx должен превратиться в w:br type="page".
    const pageBreakTyped = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      if (!canvas) return false;
      const paragraph = document.createElement('p');
      paragraph.setAttribute('data-page-break', '1');
      const br = document.createElement('br');
      br.setAttribute('data-page-break', '1');
      paragraph.appendChild(br);
      canvas.appendChild(paragraph);
      canvas.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    `);
    step('pageBreakEdit', { inserted: pageBreakTyped });
    await sleep(1800);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);

    const saved = await readPackage(target);
    const changedParts = sameEntries(original, saved);
    const unexpectedParts = changedParts.filter(
      (part) =>
        part !== 'word/document.xml' &&
        part !== 'word/_rels/document.xml.rels' &&
        !/rIdImg\d+\.png$/.test(part)
    );
    const doc = saved['word/document.xml'] || '';

    step('preserveParts', {
      changedParts,
      onlyExpectedPartsChanged: unexpectedParts.length === 0,
      stylesIntact: saved['word/styles.xml'] === original['word/styles.xml'],
      numberingIntact: saved['word/numbering.xml'] === original['word/numbering.xml'],
      headerIntact: saved['word/header1.xml'] === original['word/header1.xml'],
      footerIntact: saved['word/footer1.xml'] === original['word/footer1.xml'],
      footnotesIntact: saved['word/footnotes.xml'] === original['word/footnotes.xml'],
      commentsIntact: saved['word/comments.xml'] === original['word/comments.xml'],
      settingsIntact: saved['word/settings.xml'] === original['word/settings.xml'],
      mediaIntact: saved['word/media/image1.png'] === original['word/media/image1.png'],
      insertedImageMedia: Object.keys(saved).some((name) => /rIdImg\d+\.png$/.test(name)),
      customStyleKept: (saved['word/styles.xml'] || '').includes('MyCustomStyle'),
    });

    step('documentPatch', {
      editSaved: doc.includes('ПРАВКА_СОХРАНЕНИЯ'),
      headingKept: doc.includes('w:pStyle w:val="Heading1"'),
      listKept: doc.includes('w:numId w:val="1"'),
      tableKept: doc.includes('Ячейка A') && /<w:tbl\b/.test(doc),
      contentControlKept: /<w:sdt\b/.test(doc) && doc.includes('Содержимое content control'),
      tocFieldKept: /TOC \\o/.test(doc) && doc.includes('w:instrText'),
      mixedFormattingKept: /<w:b\/>/.test(doc) && /<w:i\/>/.test(doc) && doc.includes('C00000'),
      imageLinkKept: doc.includes('r:embed="rIdImage"'),
      noDuplicatedImage: !doc.includes('rIdImg') || /r:embed="rIdImg\d+"/.test(doc),
      insertedImageSaved: /r:embed="rIdImg\d+"/.test(doc),
      pageBreakSaved: doc.includes('<w:br w:type="page"/>'),
      headerReferenceKept: doc.includes('w:headerReference'),
      footerReferenceKept: doc.includes('w:footerReference'),
      twoSections: (doc.match(/<w:sectPr\b/g) || []).length >= 2,
      otherParagraphKept: doc.includes('Последний абзац.'),
    });

    // 3. Переоткрыли приложение: правка видна, нечитаемый объект не исчез.
    await win.webContents.reload();
    await sleep(3000);
    await clickFile();
    await sleep(3000);
    const reopened = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      return {
        text: canvas ? canvas.innerText : '',
        readonly: canvas ? canvas.querySelectorAll('[data-docx-readonly]').length : 0,
        images: canvas ? canvas.querySelectorAll('img').length : 0,
      };
    `);
    step('reopen', {
      editVisible: reopened.text.includes('ПРАВКА_СОХРАНЕНИЯ'),
      contentControlVisible: reopened.text.includes('только просмотр'),
      bodyKept: reopened.text.includes('Последний абзац.'),
      imagesKept: reopened.images >= 2,
    });

    results.passed = results.steps.every((entry) =>
      Object.entries(entry).every(([key, value]) => key === 'name' || value === true || Array.isArray(value))
    );
  } catch (err) {
    results.fatal = String(err && err.stack ? err.stack : err);
  }

  fs.writeFileSync(OUT, JSON.stringify(results, null, 2), 'utf-8');
  win.destroy();
  app.quit();
}

app.whenReady().then(main).catch((err) => {
  fs.writeFileSync(OUT, JSON.stringify({ fatal: String(err), results }, null, 2), 'utf-8');
  if (!app.isReady()) app.quit();
});
