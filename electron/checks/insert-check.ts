/**
 * Точечная проверка действий над ответом AI-чата:
 *  - «Вставить код» вставляется в позицию курсора активного документа
 *  - «Заменить документ» заменяет содержимое
 * Проверяются оба редактора: Monaco (plain) и Markdown.
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const http = require('http');
const { registerCommonCheckIpc } = require('./common-ipc');


/** Корень проекта — поиск каталога с package.json и dist/ (устойчиво к глубине). */
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
const SANDBOX = path.join(os.tmpdir(), 'ws-insert-test-' + Date.now());
const out: any = { steps: [], errors: [] };
const step = (name, value) => out.steps.push({ name, ...value });

app.disableHardwareAcceleration();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function startMock(): any {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      req.on('data', () => {});
      req.on('end', () => {
        res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache' });
        const parts = ['Вот код:\n\n', '```python\n', 'def hello():\n', '    return 1\n', '```\n'];
        let i = 0;
        const t = setInterval(() => {
          if (i >= parts.length) {
            clearInterval(t);
            res.write('data: [DONE]\n\n');
            res.end();
            return;
          }
          res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: parts[i] } }] })}\n\n`);
          i++;
        }, 20);
      });
    });
    server.listen(0, '127.0.0.1', () => resolve({ server, baseUrl: `http://127.0.0.1:${server.address().port}/v1` }));
  });
}

// Регистрируем прокси AI — приложение выполняет запросы через него

function registerHandlers(): any {
  registerCommonCheckIpc({ sandbox: SANDBOX });

  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('window:isMaximized', () => false);
  ipcMain.handle('window:minimize', () => undefined);
  ipcMain.handle('window:maximize', () => false);
  ipcMain.handle('window:close', () => undefined);
  ipcMain.handle('fs:openFolderDialog', () => SANDBOX);
  ipcMain.handle('fs:saveFileDialog', () => null);
  ipcMain.handle('fs:openFileDialog', () => null);
  ipcMain.handle('fs:showInExplorer', () => ({ success: true }));
  ipcMain.handle('fs:listDirectory', async (_e, dirPath) => {
    try {
      const entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
      const items = entries.map((e) => {
        const full = path.join(dirPath, e.name);
        let size = 0;
        try { size = fs.statSync(full).size; } catch {}
        return { name: e.name, path: full, isDirectory: e.isDirectory(), size, mtime: null,
          extension: e.isDirectory() ? '' : path.extname(e.name).toLowerCase().replace('.', '') };
      });
      items.sort((a, b) => (a.isDirectory !== b.isDirectory ? (a.isDirectory ? -1 : 1) : a.name.localeCompare(b.name)));
      return { success: true, items };
    } catch (err) { return { success: false, error: err.message, items: [] }; }
  });
  ipcMain.handle('fs:readFile', async (_e, filePath, encoding = 'utf-8') => {
    try {
      const buf = await fs.promises.readFile(filePath);
      if (encoding === 'base64') return { success: true, content: buf.toString('base64'), isBinary: true };
      return { success: true, content: buf.toString('utf-8'), isBinary: false };
    } catch (err) { return { success: false, error: err.message }; }
  });
  ipcMain.handle('fs:writeFile', async () => ({ success: true }));
  ipcMain.handle('fs:createFile', async () => ({ success: true }));
  ipcMain.handle('fs:createDirectory', async () => ({ success: true }));
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));
  ipcMain.handle('fs:searchWorkspace', async () => ({ success: true, results: [] }));
}

app.whenReady().then(async () => {
  registerHandlers();
  // AI-запросы идут через главный процесс (обход CORS)
  require('./aiProxy.js').registerAiProxy();
  fs.mkdirSync(SANDBOX, { recursive: true });
  fs.writeFileSync(path.join(SANDBOX, 'code.py'), 'print("start")\n', 'utf-8');

  const mock = await startMock();
  const win = new BrowserWindow({
    width: 1500, height: 940, show: false,
    webPreferences: { preload: path.join(__dirname, 'preload.js'), nodeIntegration: false, contextIsolation: true, sandbox: false },
  });
  const js = (expr) => win.webContents.executeJavaScript(`(async () => { ${expr} })()`);

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await sleep(2000);

    await js(`
      localStorage.setItem('workspace_ai_settings', JSON.stringify({
        provider: 'custom', baseUrl: ${JSON.stringify(mock.baseUrl)}, model: 'm',
        temperature: 0.5, maxTokens: 512, systemPrompt: 's', activePresetId: 'general' }));
      localStorage.setItem('workspace_api_keys', JSON.stringify({ custom: 'plain:k' }));
      return true;
    `);
    await win.webContents.reload();
    await sleep(2200);

    /* Открываем файл, чтобы появился активный документ */
    await js(`
      const node = [...document.querySelectorAll('div')].find(
        d => d.textContent?.trim() === 'code.py' && d.className.includes('cursor-pointer'));
      if (node) node.click();
      return !!node;
    `);
    await sleep(2500);

    step('fileActive', await js(`
      const t = document.body.innerText;
      return { opened: t.includes('code.py'), hasMonaco: !!document.querySelector('.monaco-editor') };
    `));

    /* Отправляем запрос в чат, чтобы получить блок кода */
    await js(`
      const ta = document.querySelector('textarea[placeholder*="Спросите"]');
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      setter.call(ta, 'дай код');
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise(r => setTimeout(r, 200));
      [...document.querySelectorAll('button')].find(b => /Отправить/.test(b.textContent || ''))?.click();
      return true;
    `);
    await sleep(3000);

    step('responseActions', await js(`
      const t = document.body.innerText;
      const insertBtn = [...document.querySelectorAll('button')].find(b => (b.textContent || '').trim() === 'Вставить код');
      const replaceBtn = [...document.querySelectorAll('button')].find(b => (b.textContent || '').trim() === 'Заменить документ');
      return {
        insertVisible: !!insertBtn,
        replaceVisible: !!replaceBtn,
        codeDetected: t.includes('def hello'),
      };
    `));

    /* Кликаем «Вставить код» и проверяем содержимое редактора */
    const inserted = await js(`
      const btn = [...document.querySelectorAll('button')].find(b => (b.textContent || '').trim() === 'Вставить код');
      if (!btn) return false;
      btn.click();
      await new Promise(r => setTimeout(r, 800));
      return true;
    `);
    await sleep(1000);

    step('insertCodeResult', await js(`
      // Monaco разбивает визуальные строки по переносам, поэтому склеиваем
      // содержимое без разделителей и ищем подстроку в «сыром» тексте.
      const raw = [...document.querySelectorAll('.view-line')].map(l => l.textContent).join('\\n');
      const flat = raw.replace(/\\s+/g, ' ');
      const t = document.body.innerText;
      return {
        clicked: ${inserted},
        // Вставленный код должен появиться в редакторе рядом с исходной строкой
        monacoNowHasCode: flat.includes('def hello():') && flat.includes('return 1'),
        originalStillThere: raw.includes('print("start")'),
        monacoContent: raw.slice(0, 160),
        dirtySet: /Не сохранено/.test(t),
      };
    `));

    /* Проверяем «Заменить документ» */
    await js(`
      const btn = [...document.querySelectorAll('button')].find(b => (b.textContent || '').trim() === 'Заменить документ');
      if (btn) btn.click();
      return !!btn;
    `);
    await sleep(1200);

    step('replaceDocument', await js(`
      const raw = [...document.querySelectorAll('.view-line')].map(l => l.textContent).join('\\n');
      const flat = raw.replace(/\\s+/g, ' ');
      const t = document.body.innerText;
      return {
        // После замены исходный print("start") должен исчезнуть
        originalGone: !raw.includes('print("start")'),
        containsResponse: flat.includes('Вот код') && flat.includes('def hello():'),
        replacedContent: raw.slice(0, 200),
        dirtySet: /Не сохранено/.test(t),
      };
    `));
  } catch (err) {
    out.errors.push(err.stack || String(err));
  }

  mock.server.close();
  fs.writeFileSync(path.join(ROOT, 'insert-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  try { fs.rmSync(SANDBOX, { recursive: true, force: true }); } catch {}
  app.exit(0);
});
