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
  // Каналы есть в preload приложения, но в проверочной среде главный процесс
  // свой — без заглушек рендерер получает отказ и пишет ошибку в консоль.
  ipcMain.handle('system:getInitialExternalPath', () => null);
  ipcMain.handle('window:getAlwaysOnTop', () => false);
  ipcMain.handle('window:setAlwaysOnTop', (_e, value) => ({ success: true, alwaysOnTop: Boolean(value) }));
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

  ipcMain.handle('fs:readPreview', async (_e, filePath, maxBytes = 65536) => {
    try {
      const stat = await fs.promises.stat(filePath);
      const size = Math.min(Math.max(1024, maxBytes), stat.size);
      const handle = await fs.promises.open(filePath, 'r');
      try {
        const buffer = Buffer.alloc(size);
        const { bytesRead } = await handle.read(buffer, 0, size, 0);
        return {
          success: true,
          base64: buffer.subarray(0, bytesRead).toString('base64'),
          size: stat.size,
        };
      } finally {
        await handle.close();
      }
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('system:openPath', async () => ({ success: true }));

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
          tableTextClean: !(document.querySelector('.ws-prose table')?.innerText || '').includes('[object Object]'),
          tableHasData: (() => {
            const text = document.querySelector('.ws-prose table')?.innerText || '';
            return /A/.test(text) && /B/.test(text) && /1/.test(text) && /2/.test(text);
          })(),
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
          hasFormattingToolbar: !!document.querySelector('button[title="Полужирный (Ctrl+B)"]'),
          hasMerge: !!document.querySelector('button[title="Объединить ячейки"]'),
          hasFreeze: !!document.querySelector('button[title*="Закрепить строки"]'),
          hasFind: !!document.querySelector('button[title="Найти и заменить (Ctrl+F)"]'),
          hasNumberFormats: !!document.querySelector('select[title="Числовой формат"]'),
          hasExcelBridge: !!document.querySelector('button[title*="настоящем Excel"]'),
          hasMacroButton: !!document.querySelector('button[title*="VBA-макрос"]'),
          hasReloadButton: !!document.querySelector('button[title*="Перечитать файл"]'),
        };
      })()
    `);

    /* ---- Форматирование ячейки и сохранение в файл ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const bold = document.querySelector('button[title="Полужирный (Ctrl+B)"]');
        if (bold) bold.click();
        return !!bold;
      })()
    `);
    await sleep(700);
    results.excelFormatting = await win.webContents.executeJavaScript(`
      (() => {
        const cell = [...document.querySelectorAll('tbody td')]
          .find((td) => td.textContent.trim() === 'Товар');
        return {
          boldApplied: !!cell && getComputedStyle(cell).fontWeight === '700',
        };
      })()
    `);

    // Числовой формат: выбираем ячейку с числом и включаем «Финансовый»
    await win.webContents.executeJavaScript(`
      (() => {
        const cell = [...document.querySelectorAll('tbody td')]
          .find((td) => td.textContent.trim() === '500');
        if (cell) cell.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
        return !!cell;
      })()
    `);
    await sleep(500);
    await win.webContents.executeJavaScript(`
      (() => {
        const select = document.querySelector('select[title="Числовой формат"]');
        if (!select) return false;
        const setter = Object.getOwnPropertyDescriptor(window.HTMLSelectElement.prototype, 'value').set;
        setter.call(select, 'currency');
        select.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
      })()
    `);
    await sleep(700);
    const currencyApplied = await win.webContents.executeJavaScript(`
      (() => {
        const cell = [...document.querySelectorAll('tbody td')]
          .find((td) => td.textContent.includes('500'));
        return !!cell && (cell.textContent.includes('₽') || cell.textContent.includes('\\u20BD'));
      })()
    `);
    results.excelFormatting.currencyApplied = currencyApplied;

    // Объединение ячеек: выделяем A1:B1 и нажимаем «Объединить»
    await win.webContents.executeJavaScript(`
      (() => {
        const cells = [...document.querySelectorAll('tbody td')];
        const a1 = cells.find((td) => td.textContent.trim() === 'Товар');
        const b1 = cells.find((td) => td.textContent.trim() === 'Кол-во');
        if (a1) a1.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
        if (b1) b1.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1, shiftKey: true }));
        return Boolean(a1 && b1);
      })()
    `);
    await sleep(400);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Объединить ячейки"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(600);
    results.excelMerge = await win.webContents.executeJavaScript(`
      (() => {
        const merged = [...document.querySelectorAll('tbody td[colspan]')];
        const cell = merged.find((td) => td.textContent.includes('Товар'));
        return {
          mergeApplied: !!cell && cell.getAttribute('colspan') === '2',
        };
      })()
    `);

    // Фильтр: панель со значениями выбранного столбца
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Фильтр по значениям выбранного столбца"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(600);
    results.excelFilter = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText;
        return {
          panelOpened: t.includes('Фильтр:'),
          hasCheckboxes: document.querySelectorAll('input[type=checkbox]').length > 0,
        };
      })()
    `);
    await win.webContents.executeJavaScript(`
      (() => {
        const input = [...document.querySelectorAll('input[type=checkbox]')];
        if (input.length > 0) {
          input[0].click();
        }
        return input.length > 0;
      })()
    `);
    await sleep(500);
    results.excelFilter.rowsHidden = await win.webContents.executeJavaScript(`
      (() => {
        const visibleRows = document.querySelectorAll('tbody tr').length;
        return visibleRows < 60;
      })()
    `);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = [...document.querySelectorAll('button')]
          .find((b) => (b.textContent || '').trim() === 'Сбросить');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(400);
    await win.webContents.executeJavaScript(`
      (() => {
        window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true }));
        return true;
      })()
    `);
    await sleep(1800);
    const savedWorkbook = XLSX.read(fs.readFileSync(path.join(SANDBOX, 'Финансы.xlsx')), {
      type: 'buffer',
    });
    const savedRawMeta = (savedWorkbook.Custprops as Record<string, string> | undefined)
      ?.WorkSpaceStyles;
    const savedMeta =
      typeof savedRawMeta === 'string' && savedRawMeta.startsWith('b64:')
        ? Buffer.from(savedRawMeta.slice(4), 'base64').toString('utf8')
        : savedRawMeta;
    const savedSheet = savedWorkbook.Sheets['Продажи'] as
      | { '!merges'?: unknown[] }
      | undefined;
    results.excelPersistence = {
      stylesSaved: typeof savedMeta === 'string' && savedMeta.includes('"bold":true'),
      mergeSaved: Array.isArray(savedSheet?.['!merges']) && savedSheet['!merges'].length === 1,
    };

    // Сортировка: выбираем B2 (10) и сортируем по возрастанию
    await win.webContents.executeJavaScript(`
      (() => {
        const cell = [...document.querySelectorAll('tbody td')]
          .find((td) => td.textContent.trim() === '10');
        if (cell) cell.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
        return !!cell;
      })()
    `);
    await sleep(400);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Сортировка по возрастанию"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(700);
    results.excelSort = await win.webContents.executeJavaScript(`
      (() => {
        const values = [...document.querySelectorAll('tbody tr')]
          .map((tr) => tr.querySelectorAll('td')[2]?.textContent.trim() || '')
          .filter((value) => /^\\d+$/.test(value));
        return {
          sortedAsc: values[0] === '5' && values[1] === '10' && values[2] === '15',
        };
      })()
    `);

    // Закрепление области
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title*="Закрепить строки"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(600);
    results.excelFreeze = await win.webContents.executeJavaScript(`
      (() => {
        const sticky = [...document.querySelectorAll('tbody td')].filter(
          (td) => td.style.position === 'sticky' && td.style.top
        );
        return { freezeApplied: sticky.length > 0 };
      })()
    `);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Снять закрепление"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(300);

    /* ---- Примечание к ячейке ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        window.prompt = () => 'Проверочное примечание';
        const rows = [...document.querySelectorAll('tbody tr')];
        const cell = rows[0]?.querySelectorAll('td')[1];
        if (cell) cell.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, clientX: 120, clientY: 160 }));
        return !!cell;
      })()
    `);
    await sleep(300);
    await win.webContents.executeJavaScript(`
      (() => {
        const item = [...document.querySelectorAll('button')]
          .find((b) => (b.textContent || '').trim() === 'Примечание…');
        if (item) item.click();
        return !!item;
      })()
    `);
    await sleep(500);
    results.excelComment = await win.webContents.executeJavaScript(`
      (() => {
        const rows = [...document.querySelectorAll('tbody tr')];
        const cell = rows[0]?.querySelectorAll('td')[1];
        return {
          commentSaved: !!cell && (cell.getAttribute('title') || '').includes('Проверочное примечание'),
        };
      })()
    `);

    /* ---- Диаграмма по диапазону ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const rows = [...document.querySelectorAll('tbody tr')];
        const start = rows[0]?.querySelectorAll('td')[1];
        const end = rows[3]?.querySelectorAll('td')[rows[3].querySelectorAll('td').length - 1];
        if (start) start.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
        if (end) end.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1, shiftKey: true }));
        return Boolean(start && end);
      })()
    `);
    await sleep(400);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Построить диаграмму по выделенному диапазону"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(400);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = [...document.querySelectorAll('button')]
          .find((b) => (b.textContent || '').trim() === 'Создать');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(800);
    results.excelChart = await win.webContents.executeJavaScript(`
      (() => {
        const svg = document.querySelector('svg[role="img"]');
        const t = document.body.innerText;
        return {
          chartCreated: !!svg && t.includes('Диаграмма 1'),
        };
      })()
    `);

    // Диаграмма должна попасть в xlsx как нативный объект OOXML
    await sleep(1500);
    await win.webContents.executeJavaScript(`
      (() => {
        window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true }));
        return true;
      })()
    `);
    await sleep(2500);
    const JSZip = require('jszip');
    const chartZip = await JSZip.loadAsync(fs.readFileSync(path.join(SANDBOX, 'Финансы.xlsx')));
    results.excelChartFile = {
      chartPart: Boolean(chartZip.file('xl/charts/chart1.xml')),
      drawingPart: Boolean(chartZip.file('xl/drawings/drawing1.xml')),
      drawingRel: Boolean(chartZip.file('xl/drawings/_rels/drawing1.xml.rels')),
    };
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Удалить диаграмму"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(300);

    /* ---- Проверка данных списком ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const rows = [...document.querySelectorAll('tbody tr')];
        const cell = rows[1]?.querySelectorAll('td')[2];
        if (cell) cell.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
        return !!cell;
      })()
    `);
    await sleep(300);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Проверка данных: выпадающий список"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(300);
    await win.webContents.executeJavaScript(`
      (() => {
        const input = document.querySelector('input[placeholder*="Значения через ;"]');
        if (!input) return false;
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(input, 'Да; Нет');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        const apply = [...document.querySelectorAll('button')]
          .find((b) => (b.textContent || '').trim() === 'Применить');
        if (apply) apply.click();
        return !!apply;
      })()
    `);
    await sleep(600);
    results.excelValidation = await win.webContents.executeJavaScript(`
      (() => {
        const rows = [...document.querySelectorAll('tbody tr')];
        const cell = rows[1]?.querySelectorAll('td')[2];
        return { indicatorShown: !!cell && (cell.textContent || '').includes('▾') };
      })()
    `);

    /* ---- Условное форматирование ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const rows = [...document.querySelectorAll('tbody tr')];
        const start = rows[1]?.querySelectorAll('td')[2];
        const end = rows[3]?.querySelectorAll('td')[2];
        if (start) start.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
        if (end) end.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1, shiftKey: true }));
        return Boolean(start && end);
      })()
    `);
    await sleep(300);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Условное форматирование"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(300);
    await win.webContents.executeJavaScript(`
      (() => {
        const value = document.querySelector('input[placeholder="Значение"]');
        if (value) {
          const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
          setter.call(value, '6');
          value.dispatchEvent(new Event('input', { bubbles: true }));
        }
        const add = [...document.querySelectorAll('button')]
          .find((b) => (b.textContent || '').trim() === 'Добавить правило');
        if (add) add.click();
        return Boolean(value && add);
      })()
    `);
    await sleep(600);
    results.excelConditional = await win.webContents.executeJavaScript(`
      (() => {
        const rows = [...document.querySelectorAll('tbody tr')];
        const cells = [rows[1], rows[2], rows[3]].map((row) => row?.querySelectorAll('td')[2]);
        return {
          ruleApplied: cells.some((cell) => cell && getComputedStyle(cell).backgroundColor === 'rgb(253, 230, 138)'),
        };
      })()
    `);

    /* ---- Сводная таблица ---- */
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Сводная таблица по выделенному диапазону"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(500);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = [...document.querySelectorAll('button')]
          .find((b) => (b.textContent || '').trim() === 'Создать лист');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(900);
    results.excelPivot = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText || '';
        return {
          sheetCreated: t.includes('Сводная 1'),
          refreshButton: !!document.querySelector('button[title="Пересчитать сводную по исходным данным"]'),
        };
      })()
    `);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Пересчитать сводную по исходным данным"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(700);
    results.excelPivot.refreshed = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText || '';
        return t.includes('Сводная таблица обновлена');
      })()
    `);

    // Сохранение новых возможностей в xlsx: примечание, правило, сводный лист
    await win.webContents.executeJavaScript(`
      (() => {
        window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true }));
        return true;
      })()
    `);
    await sleep(1600);
    const finalWorkbook = XLSX.read(fs.readFileSync(path.join(SANDBOX, 'Финансы.xlsx')), {
      type: 'buffer',
    });
    const finalRawMeta = (finalWorkbook.Custprops as Record<string, string> | undefined)
      ?.WorkSpaceStyles;
    const finalMeta =
      typeof finalRawMeta === 'string' && finalRawMeta.startsWith('b64:')
        ? Buffer.from(finalRawMeta.slice(4), 'base64').toString('utf8')
        : finalRawMeta;
    results.excelExtrasPersistence = {
      commentSaved: typeof finalMeta === 'string' && finalMeta.includes('Проверочное примечание'),
      ruleSaved: typeof finalMeta === 'string' && finalMeta.includes('"operator":"greater"'),
      pivotSheetSaved: finalWorkbook.SheetNames.includes('Сводная 1'),
    };

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

    /* ---- Любой файл, включая EXE, открывается в бинарном просмотрщике ---- */
    fs.writeFileSync(
      path.join(SANDBOX, 'sample.exe'),
      Buffer.from([0x4d, 0x5a, 0x90, 0x00, 0x03, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00])
    );
    await win.webContents.executeJavaScript(`
      (() => {
        const refreshBtn = [...document.querySelectorAll('button')].find(
          (b) => (b.title || '').includes('Обновить дерево')
        );
        if (refreshBtn) refreshBtn.click();
        return !!refreshBtn;
      })()
    `);
    await sleep(1200);
    await win.webContents.executeJavaScript(`
      (() => {
        const node = [...document.querySelectorAll('div')].find(
          (d) => d.textContent?.trim() === 'sample.exe' && d.className.includes('cursor-pointer')
        );
        if (node) node.click();
        return !!node;
      })()
    `);
    await sleep(1500);
    results.binaryViewer = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText || '';
        return {
          opened: t.includes('sample.exe') && t.includes('Offset'),
          hasHex: /4D\\s+5A/.test(t),
          hasOpenExternal: t.includes('Открыть в системе'),
          hasReveal: t.includes('Проводник'),
        };
      })()
    `);

    /* ---- Скриншоты ключевых состояний ---- */
    win.show();
    await sleep(700);
    const shot = await win.webContents.capturePage();
    fs.writeFileSync(path.join(ROOT, 'render-screenshot.png'), shot.toPNG());

    /* ---- Повторное открытие файла: форматирование и метаданные из xlsx ---- */
    win.webContents.reload();
    await sleep(3500);
    await win.webContents.executeJavaScript(`
      (() => {
        const node = [...document.querySelectorAll('div')].find(
          (d) => d.textContent?.trim() === 'Финансы.xlsx' && d.className.includes('cursor-pointer')
        );
        if (node) node.click();
        return !!node;
      })()
    `);
    await sleep(2500);
    results.excelReload = await win.webContents.executeJavaScript(`
      (() => {
        const rows = [...document.querySelectorAll('tbody tr')];
        const a1 = rows[0]?.querySelectorAll('td')[1];
        const bCells = [rows[1], rows[2], rows[3]].map((row) => row?.querySelectorAll('td')[2]);
        return {
          commentRestored: !!a1 && (a1.getAttribute('title') || '').includes('Проверочное примечание'),
          validationRestored: !!bCells[0] && (bCells[0].textContent || '').includes('▾'),
          conditionalRestored: bCells.some(
            (cell) => cell && getComputedStyle(cell).backgroundColor === 'rgb(253, 230, 138)'
          ),
        };
      })()
    `);

    // Финальная диаграмма остаётся в файле для проверки реальным Excel
    await win.webContents.executeJavaScript(`
      (() => {
        const rows = [...document.querySelectorAll('tbody tr')];
        const start = rows[0]?.querySelectorAll('td')[1];
        const end = rows[3]?.querySelectorAll('td')[3];
        if (start) start.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
        if (end) end.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1, shiftKey: true }));
        return Boolean(start && end);
      })()
    `);
    await sleep(400);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = document.querySelector('button[title="Построить диаграмму по выделенному диапазону"]');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(400);
    await win.webContents.executeJavaScript(`
      (() => {
        const button = [...document.querySelectorAll('button')]
          .find((b) => (b.textContent || '').trim() === 'Создать');
        if (button) button.click();
        return !!button;
      })()
    `);
    await sleep(1800);
    await win.webContents.executeJavaScript(`
      (() => {
        window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true }));
        return true;
      })()
    `);
    await sleep(2500);
    results.excelChartFinal = await (async () => {
      const finalZip = await JSZip.loadAsync(fs.readFileSync(path.join(SANDBOX, 'Финансы.xlsx')));
      return { chartInFile: Boolean(finalZip.file('xl/charts/chart1.xml')) };
    })();

    /* ---- Кнопки «Вставить в документ» / «Заменить документ» для Excel ---- */
    results.excelChatInsert = await win.webContents.executeJavaScript(`
      (async () => {
        window.dispatchEvent(new CustomEvent('workspace:excel-insert-text', {
          detail: { text: '| Тест1 | Тест2 |\\n| --- | --- |\\n| 11 | 22 |' },
        }));
        await new Promise((r) => setTimeout(r, 900));
        const t = document.body.innerText || '';
        return { inserted: t.includes('Тест1') && t.includes('22') };
      })()
    `);
    results.excelChatReplace = await win.webContents.executeJavaScript(`
      (async () => {
        window.dispatchEvent(new CustomEvent('workspace:excel-replace-document', {
          detail: { text: '| R1 | R2 |\\n| --- | --- |\\n| 1 | 2 |' },
        }));
        await new Promise((r) => setTimeout(r, 900));
        const t = document.body.innerText || '';
        const rows = [...document.querySelectorAll('tbody tr')].slice(0, 3).map((tr) =>
          [...tr.querySelectorAll('td')].slice(0, 4).map((td) => td.textContent.trim())
        );
        return { replaced: rows[0]?.[1] === 'R1' && rows[1]?.[1] === '1', rows };
      })()
    `);

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
