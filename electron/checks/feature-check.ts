/**
 * Точечная проверка двух сценариев, которые не удалось подтвердить в
 * render-check: переключение светлой темы и открытие панели поиска.
 * Вызывает обработчики напрямую через DOM-события React.
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
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
const out: any = {};

app.disableHardwareAcceleration();

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

app.whenReady().then(async () => {
  // Общие каналы (избранное, буфер обмена, секреты) — иначе рендерер падает
  // на 'fs:getFavorites' и проверки темы/поиска не доходят до интерфейса.
  registerCommonCheckIpc({ sandbox: SANDBOX });

  ipcMain.handle('window:isMaximized', () => false);
  ipcMain.handle('window:minimize', () => undefined);
  ipcMain.handle('window:maximize', () => false);
  ipcMain.handle('window:close', () => undefined);
  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('fs:openFolderDialog', () => SANDBOX);
  ipcMain.handle('fs:showInExplorer', () => ({ success: true }));
  ipcMain.handle('fs:searchWorkspace', async (_e, rootDir, query, searchContent) => {
    const res = [];
    const walk = (dir) => {
      let entries = [];
      try {
        entries = fs.readdirSync(dir, { withFileTypes: true });
      } catch {
        return;
      }
      for (const entry of entries) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) walk(full);
        else {
          const nameHit = entry.name.toLowerCase().includes(String(query).toLowerCase());
          if (nameHit) res.push({ path: full, name: entry.name, isDirectory: false });
          else if (searchContent) {
            try {
              const text = fs.readFileSync(full, 'utf-8');
              const idx = text.toLowerCase().indexOf(String(query).toLowerCase());
              if (idx !== -1) {
                const line = text.slice(0, idx).split('\n').length;
                res.push({
                  path: full,
                  name: entry.name,
                  isDirectory: false,
                  line,
                  matchSnippet: text.split('\n')[line - 1]?.trim().slice(0, 100),
                  totalMatches: 1,
                });
              }
            } catch {}
          }
        }
      }
    };
    walk(rootDir);
    return { success: true, results: res };
  });
  ipcMain.handle('fs:listDirectory', async (_e, dirPath) => {
    try {
      const entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
      const items = entries.map((e) => {
        const full = path.join(dirPath, e.name);
        let size = 0;
        try {
          size = fs.statSync(full).size;
        } catch {}
        return {
          name: e.name,
          path: full,
          isDirectory: e.isDirectory(),
          size,
          mtime: null,
          extension: e.isDirectory() ? '' : path.extname(e.name).toLowerCase().replace('.', ''),
        };
      });
      items.sort((a, b) => (a.isDirectory !== b.isDirectory ? (a.isDirectory ? -1 : 1) : a.name.localeCompare(b.name)));
      return { success: true, items };
    } catch (err) {
      return { success: false, error: err.message, items: [] };
    }
  });
  ipcMain.handle('fs:readFile', async (_e, filePath, encoding = 'utf-8') => {
    try {
      const buf = await fs.promises.readFile(filePath);
      if (encoding === 'base64') return { success: true, content: buf.toString('base64'), isBinary: true };
      return { success: true, content: buf.toString('utf-8'), isBinary: false };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });
  ipcMain.handle('fs:writeFile', async () => ({ success: true }));
  ipcMain.handle('fs:createFile', async () => ({ success: true }));
  ipcMain.handle('fs:createDirectory', async () => ({ success: true }));
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));

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

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await sleep(2500);

    /* ---- 1. Переключение темы ---- */
    out.themeBefore = await win.webContents.executeJavaScript(`
      (() => {
        const cs = getComputedStyle(document.body);
        return {
          // Тема по умолчанию наследуется от системной настройки,
          // поэтому фиксируем фактическое состояние, а не ожидаемое.
          darkClassBeforeToggle: document.documentElement.classList.contains('dark'),
          bodyBg: cs.backgroundColor,
          bodyColor: cs.color,
          titlebarBg: document.querySelector('.titlebar-drag')
            ? getComputedStyle(document.querySelector('.titlebar-drag')).backgroundColor : null,
        };
      })()
    `);

    // Клик по кнопке темы (ищем по стабильному признаку — иконке sun/moon внутри titlebar)
    out.themeButtonFound = await win.webContents.executeJavaScript(`
      (() => {
        const btns = [...document.querySelectorAll('.titlebar-no-drag button')];
        // Кнопка темы предпоследняя перед настройками; ищем по title
        const themeBtn = btns.find(b => /тема/i.test(b.title || ''));
        if (themeBtn) { themeBtn.click(); return themeBtn.title; }
        return null;
      })()
    `);
    await sleep(900);

    out.themeAfter = await win.webContents.executeJavaScript(`
      (() => {
        const cs = getComputedStyle(document.body);
        return {
          darkClassAfterToggle: document.documentElement.classList.contains('dark'),
          bodyBg: cs.backgroundColor,
          bodyColor: cs.color,
          titlebarBg: document.querySelector('.titlebar-drag')
            ? getComputedStyle(document.querySelector('.titlebar-drag')).backgroundColor : null,
          storedTheme: localStorage.getItem('workspace_theme'),
        };
      })()
    `);

    // Проверяем, что переключение реально изменило и класс, и цвета
    out.themeToggled = {
      darkClassFlipped:
        out.themeBefore.darkClassBeforeToggle !== out.themeAfter.darkClassAfterToggle,
      colorsChanged:
        out.themeBefore.bodyBg !== out.themeAfter.bodyBg &&
        out.themeBefore.titlebarBg !== out.themeAfter.titlebarBg,
      themePersisted:
        out.themeAfter.storedTheme ===
        (out.themeAfter.darkClassAfterToggle ? 'dark' : 'light'),
    };

    // Возвращаем тёмную
    await win.webContents.executeJavaScript(`
      (() => {
        const themeBtn = [...document.querySelectorAll('.titlebar-no-drag button')]
          .find(b => /тема/i.test(b.title || ''));
        if (themeBtn) themeBtn.click();
        return true;
      })()
    `);
    await sleep(700);

    out.themeRestored = await win.webContents.executeJavaScript(`
      document.documentElement.classList.contains('dark')
    `);

    /* ---- 2. Панель поиска ---- */
    out.searchButtonFound = await win.webContents.executeJavaScript(`
      (() => {
        // Переключатель «Поиск» — сосед кнопки «Файлы» в шапке боковой панели
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

    out.searchPanelState = await win.webContents.executeJavaScript(`
      (() => {
        const t = document.body.innerText;
        const input = document.querySelector('input[placeholder*="Искать текст"]');
        return {
          panelOpen: /ПОИСК ПО ФАЙЛАМ/i.test(t),
          hasSearchInput: !!input,
          hasCaseToggle: !!document.querySelector('button[title="Учитывать регистр"]'),
          hasRegexToggle: !!document.querySelector('button[title="Регулярное выражение"]'),
          hasFilterToggle: t.includes('Фильтр'),
        };
      })()
    `);

    /* ---- 3. Выполняем реальный поиск ---- */
    out.searchExecution = await win.webContents.executeJavaScript(`
      (async () => {
        const input = document.querySelector('input[placeholder*="Искать текст"]');
        if (!input) return { error: 'no input' };
        // React слушает onChange — эмулируем ввод через native setter
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(input, 'ERROR');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        await new Promise(r => setTimeout(r, 300));
        // Отправляем форму
        const form = input.closest('form');
        if (form) form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
        await new Promise(r => setTimeout(r, 2000));
        const t = document.body.innerText;
        return {
          hasResult: t.includes('server.log'),
          hasLineBadge: /стр\\.\\s*\\d+/.test(t) || t.includes('Стр.'),
          hasSnippet: t.includes('ERROR'),
          resultsFooter: /Файлов:\\s*\\d+/.test(t),
        };
      })()
    `);

    /* ---- 4. Проверяем Quick Open (Ctrl+P) ---- */
    out.quickOpen = await win.webContents.executeJavaScript(`
      (async () => {
        window.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true }));
        await new Promise(r => setTimeout(r, 1200));
        const t = document.body.innerText;
        return {
          opened: t.includes('Начните вводить имя файла') || t.includes('Найдено:'),
          hasResults: t.includes('Заметки.md') || t.includes('config.json'),
          hasFooterHints: t.includes('навигация') || t.includes('Enter открыть'),
        };
      })()
    `);

    // Закрываем модалку
    await win.webContents.executeJavaScript(`
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      true
    `);
    await sleep(500);

    /* ---- 5. Проверяем панель настроек AI ---- */
    out.settingsModal = await win.webContents.executeJavaScript(`
      (async () => {
        // Кнопка настроек находится в строке состояния (StatusBar), а не
        // в заголовке окна: в .titlebar-no-drag её нет.
        const btn = [...document.querySelectorAll('button')]
          .find(b => /Настройки/i.test(b.title || ''));
        if (!btn) return { error: 'settings button not found' };
        btn.click();
        await new Promise(r => setTimeout(r, 900));
        const t = document.body.innerText;
        // Заголовок модального окна — «Настройки» (SettingsModal.tsx).
        // Строка «Настройки AI» существовала только как всплывающая подсказка
        // кнопки в строке состояния и в innerText не попадает.
        // Признак открытия — заголовок окна вместе с его содержимым.
        const hasHeading = t.includes('Настройки');
        return {
          opened: hasHeading && t.includes('Проверить подключение'),
          hasTabs: t.includes('Провайдер') && t.includes('Модель') && t.includes('Промт') && t.includes('Хранилище'),
          hasKeyField: !!document.querySelector('input[type=password]'),
          hasTestButton: t.includes('Проверить подключение'),
          hasProviders: t.includes('OpenAI') && t.includes('Anthropic') && t.includes('Gemini') &&
                         t.includes('DeepSeek') && t.includes('OpenRouter') && t.includes('Ollama'),
        };
      })()
    `);

    out.errors = [];
  } catch (err) {
    out.fatal = err.stack || err.message;
  }

  fs.writeFileSync(path.join(ROOT, 'feature-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  app.exit(0);
});
