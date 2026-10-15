/**
 * Проверка фактического рендеринга: загружает приложение, открывает реальную
 * рабочую папку, открывает файлы разных форматов и проверяет геометрию и
 * вычисленные стили ключевых элементов. Результат пишется в render-check.json.
 *
 * Запуск: electron.exe dist-electron/render-check.js
 */
const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const iconv = require('iconv-lite');
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
const SANDBOX = path.join(ROOT, 'test-workspace');
const errors = [];
const results: any = {};

app.disableHardwareAcceleration();

/* ---- Регистрируем те же IPC-обработчики, что и main.ts ---- */
function registerIPC() {
  // Каналы, которых нет ниже (избранное, буфер обмена, секреты, подтверждение
  // закрытия). Без fs:getFavorites рендерер падал и дерево файлов не грузилось.
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
        let mtime = null;
        try {
          const st = fs.statSync(full);
          size = st.size;
          mtime = st.mtime.toISOString();
        } catch {}
        return {
          name: entry.name,
          path: full,
          isDirectory: entry.isDirectory(),
          size,
          mtime,
          extension: entry.isDirectory() ? '' : path.extname(entry.name).toLowerCase().replace('.', ''),
        };
      });
      items.sort((a, b) => {
        if (a.isDirectory !== b.isDirectory) return a.isDirectory ? -1 : 1;
        return a.name.localeCompare(b.name, 'ru', { numeric: true });
      });
      return { success: true, items };
    } catch (err) {
      return { success: false, error: err.message, items: [] };
    }
  });

  ipcMain.handle('fs:readFile', async (_e, filePath, encoding = 'utf-8') => {
    try {
      const buffer = await fs.promises.readFile(filePath);
      if (encoding === 'base64') {
        return { success: true, content: buffer.toString('base64'), isBinary: true };
      }
      if (encoding !== 'utf-8' && iconv.encodingExists(encoding)) {
        return { success: true, content: iconv.decode(buffer, encoding), isBinary: false };
      }
      return { success: true, content: buffer.toString('utf-8'), isBinary: false };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:writeFile', async (_e, filePath, content, encoding = 'utf-8') => {
    try {
      if (encoding === 'base64') {
        await fs.promises.writeFile(filePath, Buffer.from(content, 'base64'));
      } else {
        await fs.promises.writeFile(filePath, content, 'utf-8');
      }
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:createFile', async (_e, filePath, initialContent = '') => {
    try {
      await fs.promises.writeFile(filePath, initialContent, { flag: 'wx' });
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });
  ipcMain.handle('fs:createDirectory', async (_e, dirPath) => {
    try {
      await fs.promises.mkdir(dirPath, { recursive: true });
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });
  ipcMain.handle('fs:deletePath', async (_e, target) => {
    try {
      await fs.promises.rm(target, { recursive: true, force: true });
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });
  ipcMain.handle('fs:renamePath', async (_e, oldPath, newPath) => {
    try {
      await fs.promises.rename(oldPath, newPath);
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:searchWorkspace', async (_e, rootDir, query, searchContent) => {
    const results: any[] = [];
    const walk = (dir) => {
      let entries = [];
      try {
        entries = fs.readdirSync(dir, { withFileTypes: true });
      } catch {
        return;
      }
      for (const entry of entries) {
        if (entry.name.startsWith('.') || entry.name === 'node_modules') continue;
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
          walk(full);
        } else if (entry.name.toLowerCase().includes(String(query).toLowerCase())) {
          results.push({ path: full, name: entry.name, isDirectory: false });
        } else if (searchContent) {
          try {
            const text = fs.readFileSync(full, 'utf-8');
            const idx = text.toLowerCase().indexOf(String(query).toLowerCase());
            if (idx !== -1) {
              const line = text.slice(0, idx).split('\n').length;
              results.push({
                path: full,
                name: entry.name,
                isDirectory: false,
                line,
                matchSnippet: text.split('\n')[line - 1]?.trim().slice(0, 120),
                totalMatches: 1,
              });
            }
          } catch {}
        }
      }
    };
    walk(rootDir);
    return { success: true, results: results.slice(0, 100) };
  });
}

/** Создаёт тестовые файлы разных форматов. */
function createFixtures() {
  fs.mkdirSync(SANDBOX, { recursive: true });
  fs.writeFileSync(
    path.join(SANDBOX, 'Заметки.md'),
    '# Тестовый документ\n\nЭто **проверка** рендеринга Markdown.\n\n- пункт один\n- пункт два\n- [ ] задача\n\n> цитата\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n```js\nconst x = 1;\n```\n',
    'utf-8'
  );
  fs.writeFileSync(
    path.join(SANDBOX, 'config.json'),
    JSON.stringify({ name: 'workspace', version: '1.0.0', nested: { enabled: true } }, null, 2),
    'utf-8'
  );
  // Файл в CP1251 для проверки смены кодировок
  fs.writeFileSync(
    path.join(SANDBOX, 'legacy.txt'),
    iconv.encode('Привет из Windows-1251\nСтрока два\n', 'windows-1251')
  );
  fs.writeFileSync(
    path.join(SANDBOX, 'server.log'),
    '2026-01-01 INFO Start\n2026-01-01 ERROR Failed to connect\n2026-01-01 FATAL Shutdown\n',
    'utf-8'
  );
  fs.writeFileSync(
    path.join(SANDBOX, 'script.py'),
    'def hello(name):\n    return f"Hello, {name}"\n\nprint(hello("World"))\n',
    'utf-8'
  );
  fs.mkdirSync(path.join(SANDBOX, 'Подпапка'), { recursive: true });
  fs.writeFileSync(path.join(SANDBOX, 'Подпапка', 'inner.txt'), 'вложенный файл\n', 'utf-8');
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

app.whenReady().then(async () => {
  registerIPC();
  createFixtures();

  const win = new BrowserWindow({
    width: 1500,
    height: 940,
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
    },
  });

  win.webContents.on('console-message', (_e, level, message, line, sourceId) => {
    if (level >= 3) errors.push(`${message} (${sourceId}:${line})`);
  });
  win.webContents.on('did-fail-load', (_e, code, desc) => {
    errors.push(`did-fail-load ${code} ${desc}`);
  });

  /** Снимок геометрии и стилей ключевых областей интерфейса. */
  const probeLayout = `
    (() => {
      const rect = (sel) => {
        const el = document.querySelector(sel);
        if (!el) return null;
        const r = el.getBoundingClientRect();
        const cs = getComputedStyle(el);
        return {
          w: Math.round(r.width), h: Math.round(r.height),
          x: Math.round(r.x), y: Math.round(r.y),
          bg: cs.backgroundColor, color: cs.color,
          visible: r.width > 0 && r.height > 0,
        };
      };
      return {
        bodyBg: getComputedStyle(document.body).backgroundColor,
        rootH: document.getElementById('root')?.getBoundingClientRect().height ?? 0,
        titlebar: rect('.titlebar-drag'),
        // Панель вкладок и строка состояния находятся через текст
        tabbar: (() => {
          const el = [...document.querySelectorAll('div')].find(d =>
            d.className && String(d.className).includes('h-9') && String(d.className).includes('border-b'));
          if (!el) return null;
          const r = el.getBoundingClientRect();
          return { h: Math.round(r.height), w: Math.round(r.width) };
        })(),
        text: document.body.innerText.replace(/\\s+/g, ' ').slice(0, 900),
      };
    })()
  `;

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await sleep(2500);

    // Дожидаемся загрузки дерева файлов из рабочей папки
    await win.webContents.executeJavaScript(`
      new Promise((resolve) => {
        const started = Date.now();
        const check = () => {
          if (document.body.innerText.includes('Заметки.md') || Date.now() - started > 8000) {
            resolve(true);
          } else setTimeout(check, 150);
        };
        check();
      })
    `);

    results.initialLayout = await win.webContents.executeJavaScript(probeLayout);
    results.fileTreeLoaded = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText;
        return {
          hasMarkdown: t.includes('Заметки.md'),
          hasJson: t.includes('config.json'),
          hasLog: t.includes('server.log'),
          hasPython: t.includes('script.py'),
          hasLegacy: t.includes('legacy.txt'),
          hasFolder: t.includes('Подпапка'),
        };
      })()
    `);

    /* ---- Открываем Markdown-файл кликом по дереву ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const node = [...document.querySelectorAll('div')].find(
          (d) => d.textContent?.trim() === 'Заметки.md' && d.className.includes('cursor-pointer')
        );
        if (node) node.click();
        return !!node;
      })()
    `);
    await sleep(1800);

    results.markdownOpened = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText;
        const preview = document.querySelector('.ws-prose');
        return {
          tabCreated: t.includes('Заметки.md'),
          hasPreview: !!preview,
          previewHtmlLen: preview ? preview.innerHTML.length : 0,
          hasHeading: !!document.querySelector('.ws-prose h1'),
          hasList: !!document.querySelector('.ws-prose ul, .ws-prose ol'),
          hasTable: !!document.querySelector('.ws-prose table'),
          hasCodeBlock: !!document.querySelector('.ws-prose pre'),
          hasCheckbox: !!document.querySelector('.ws-prose input[type=checkbox]'),
          hasTextarea: !!document.querySelector('textarea'),
        };
      })()
    `);

    /* ---- Статусная строка и вкладка ---- */
    results.statusBar = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText;
        // Статусная строка показывает либо имя модели, либо «AI не настроен»
        const statusText = [...document.querySelectorAll('button')]
          .map(b => b.textContent || '')
          .find(x => /AI не настроен|gpt|claude|gemini|deepseek/i.test(x)) || '';
        return {
          hasWordCount: /Слов:/.test(t),
          hasCharCount: /Символов:/.test(t),
          hasEncoding: /UTF-8/.test(t),
          hasAIStatus: statusText.length > 0,
          aiStatusText: statusText.trim(),
        };
      })()
    `);

    /* ---- Проверка светлой темы ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const btn = [...document.querySelectorAll('button')].find(
          (b) => (b.title || '').includes('Светлая тема') || (b.title || '').includes('Тёмная тема')
        );
        if (btn) btn.click();
        return !!btn;
      })()
    `);
    await sleep(1000);

    results.lightTheme = await win.webContents.executeJavaScript(`
      (() => {
        const cs = getComputedStyle(document.body);
        return {
          isDarkClass: document.documentElement.classList.contains('dark'),
          bodyBg: cs.backgroundColor,
          bodyColor: cs.color,
        };
      })()
    `);

    // Возвращаем тёмную тему
    await win.webContents.executeJavaScript(`
      (() => {
        const btn = [...document.querySelectorAll('button')].find(
          (b) => (b.title || '').includes('Светлая тема') || (b.title || '').includes('Тёмная тема')
        );
        if (btn) btn.click();
        return true;
      })()
    `);
    await sleep(600);

    /* ---- Открываем JSON в Monaco ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText;
        if (!t.includes('config.json')) return false;
        const node = [...document.querySelectorAll('div')].find(
          (d) => d.textContent?.trim() === 'config.json' && d.className.includes('cursor-pointer')
        );
        if (node) node.click();
        return !!node;
      })()
    `);
    await sleep(3000);

    results.monacoEditor = await win.webContents.executeJavaScript(`
      (() => ({
        hasMonaco: !!document.querySelector('.monaco-editor'),
        hasMinimap: !!document.querySelector('.minimap'),
        lineCount: document.querySelectorAll('.view-line').length,
        // Виджет поиска Monaco создаёт лениво, поэтому здесь его ещё нет —
        // работоспособность Ctrl+F проверяется отдельным шагом ниже.
        text: document.body.innerText.replace(/\\s+/g,' ').slice(0, 400),
      }))()
    `);

    /* ---- Ctrl+F внутри Monaco открывает виджет поиска ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const ta = document.querySelector('.monaco-editor textarea');
        if (ta) ta.focus();
        return !!ta;
      })()
    `);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F', modifiers: ['control'] });
    await sleep(1200);

    results.monacoFindWidget = await win.webContents.executeJavaScript(`
      (() => {
        const w = document.querySelector('.find-widget');
        const input = document.querySelector('.find-widget input.input, .find-widget textarea, .find-widget input');
        return {
          widgetAppeared: !!w,
          widgetVisible: w ? getComputedStyle(w).display !== 'none' && w.offsetHeight > 0 : false,
          hasSearchInput: !!input,
        };
      })()
    `);

    // Закрываем виджет поиска, чтобы не мешал дальнейшим шагам
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    await sleep(500);

    /* ---- Проверяем открытие Excel-редактора (создаём xlsx) ---- */
    const XLSX = require('xlsx');
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['Товар', 'Кол-во', 'Цена', 'Итого'],
      ['Яблоки', 10, 50, '=B2*C2'],
      ['Груши', 5, 80, '=B3*C3'],
      ['Итого', '=SUM(B2:B3)', '', '=SUM(D2:D3)'],
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Продажи');
    fs.writeFileSync(
      path.join(SANDBOX, 'Финансы.xlsx'),
      XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' })
    );

    // Обновляем дерево, чтобы новый файл появился
    await win.webContents.executeJavaScript(`
      (() => {
        const refreshBtn = [...document.querySelectorAll('button')].find(
          (b) => (b.title || '').includes('Обновить дерево')
        );
        if (refreshBtn) refreshBtn.click();
        return !!refreshBtn;
      })()
    `);
    await sleep(1500);

    await win.webContents.executeJavaScript(`
      (() => {
        const node = [...document.querySelectorAll('div')].find(
          (d) => d.textContent?.trim() === 'Финансы.xlsx' && d.className.includes('cursor-pointer')
        );
        if (node) node.click();
        return !!node;
      })()
    `);
    await sleep(3000);

    results.excelEditor = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText;
        const table = document.querySelector('table');
        const rows = table ? table.querySelectorAll('tbody tr').length : 0;
        const cells = table ? table.querySelectorAll('tbody td').length : 0;
        return {
          gridRendered: !!table,
          rowCount: rows,
          cellCount: cells,
          hasSheetTabs: t.includes('Продажи'),
          hasFormulaBar: t.includes('fx') || !!document.querySelector('input[placeholder*="формула"], input[placeholder*="SUM"]'),
          hasAIDialog: /AI-формула/.test(t),
          formulaEvaluated: t.includes('500') && t.includes('400'),
        };
      })()
    `);

    /* ---- Проверяем панель поиска ---- */
    const searchToggleClicked = await win.webContents.executeJavaScript(`
      (() => {
        // Переключатель «Поиск» находится в шапке боковой панели (рядом с «Файлы»)
        const filesBtn = [...document.querySelectorAll('button')]
          .find(b => (b.textContent || '').trim() === 'Файлы');
        const header = filesBtn?.parentElement;
        const searchBtn = header
          ? [...header.querySelectorAll('button')].find(b => (b.textContent || '').trim() === 'Поиск')
          : null;
        if (searchBtn) { searchBtn.click(); return true; }
        return false;
      })()
    `);
    await sleep(1000);
    results.searchPanel = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText;
        return {
          toggleFound: ${searchToggleClicked},
          opened: /ПОИСК ПО ФАЙЛАМ/i.test(t),
          hasSearchInput: !!document.querySelector('input[placeholder*="Искать текст"]'),
          hasCaseToggle: !!document.querySelector('button[title="Учитывать регистр"]'),
          hasRegexToggle: !!document.querySelector('button[title="Регулярное выражение"]'),
        };
      })()
    `);

    /* ---- Возвращаемся к дереву файлов ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const filesBtn = [...document.querySelectorAll('button')]
          .find(b => (b.textContent || '').trim() === 'Файлы');
        if (filesBtn) filesBtn.click();
        return !!filesBtn;
      })()
    `);
    await sleep(600);

    /* ---- Скриншоты ключевых состояний ---- */
    win.show();
    await sleep(700);
    const shot = await win.webContents.capturePage();
    fs.writeFileSync(path.join(ROOT, 'render-screenshot.png'), shot.toPNG());

    results.errors = errors;
    fs.writeFileSync(
      path.join(ROOT, 'render-check.json'),
      JSON.stringify(results, null, 2),
      'utf-8'
    );
    console.log('RENDER_CHECK_WRITTEN');
  } catch (err) {
    results.fatal = err.stack || err.message;
    fs.writeFileSync(
      path.join(ROOT, 'render-check.json'),
      JSON.stringify(results, null, 2),
      'utf-8'
    );
  }

  app.exit(0);
});
