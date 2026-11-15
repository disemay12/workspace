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
// Свой каталог на прогон: раньше набор читал файлы render-check из общего
// ROOT/test-workspace, поэтому зависел от порядка запусков и от чужих фикстур.
const { createCheckSandbox } = require('./sandbox');
const SANDBOX = createCheckSandbox('feature', ROOT);
const out: any = {};

app.disableHardwareAcceleration();

/** Падение до основного try должно попадать в отчёт, а не выглядеть зависанием. */
function failFast(reason) {
  out.fatal = (reason as Error)?.stack || String(reason);
  try {
    fs.writeFileSync(path.join(ROOT, 'feature-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  } catch {
    // даже запись диагностики может не удаться
  }
  app.exit(1);
}
process.on('unhandledRejection', failFast);
process.on('uncaughtException', failFast);

/**
 * Собственные фикстуры: поиск ищет «ERROR» в server.log, Quick Open — имена
 * файлов, поэтому набор не может рассчитывать на файлы другого набора.
 */
function createFixtures() {
  fs.writeFileSync(
    path.join(SANDBOX, 'server.log'),
    '2026-01-01 INFO Start\n2026-01-01 ERROR Failed to connect\n2026-01-01 FATAL Shutdown\n',
    'utf-8'
  );
  fs.writeFileSync(
    path.join(SANDBOX, 'Заметки.md'),
    '# Заметки\n\nТекст для Quick Open и поиска.\n',
    'utf-8'
  );
  fs.writeFileSync(
    path.join(SANDBOX, 'config.json'),
    JSON.stringify({ name: 'workspace', version: '1.0.0' }, null, 2),
    'utf-8'
  );
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

app.whenReady().then(async () => {
  createFixtures();

  // Общие каналы (избранное, буфер обмена, секреты) — иначе рендерер падает
  // на 'fs:getFavorites' и проверки темы/поиска не доходят до интерфейса.
  registerCommonCheckIpc({
    sandbox: SANDBOX,
    // Автопроверка обновлений должна получить доступную версию: сценарий 6
    // проверяет, что баннер появляется сам, без нажатия «Проверить обновление».
    updateCheck: {
      success: true,
      supported: true,
      currentVersion: '1.0.0',
      latestVersion: '9.9.9',
      available: true,
      assetName: 'WorkSpace-9.9.9-win-x64.zip',
      assetUrl: 'https://github.com/example/workspace/releases/download/v9.9.9/WorkSpace-9.9.9-win-x64.zip',
      assetDigest: 'sha256:' + '0'.repeat(64),
    },
  });

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
    const themeButton = await win.webContents.executeJavaScript(`
      (() => {
        const btns = [...document.querySelectorAll('.titlebar-no-drag button')];
        // Кнопка темы предпоследняя перед настройками; ищем по title
        const themeBtn = btns.find(b => /тема/i.test(b.title || ''));
        if (themeBtn) themeBtn.click();
        return { found: Boolean(themeBtn), title: themeBtn ? (themeBtn.title || '') : '' };
      })()
    `);
    // Утверждение — found; заголовок кнопки остаётся диагностикой.
    out.themeButtonFound = Boolean(themeButton.found);
    out.themeButtonTitle = themeButton.title;
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
        // Нет поля ввода — это провал именованного утверждения inputFound,
        // а не строка 'no input', которую раннер принимал за успех.
        if (!input) {
          return { inputFound: false, hasResult: false, hasLineBadge: false, hasSnippet: false, resultsFooter: false };
        }
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
          inputFound: true,
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
        // Нет кнопки — именованный провал buttonFound, а не строка-«ошибка»,
        // которую раннер принимал за успех.
        if (!btn) {
          return { buttonFound: false, opened: false, hasTabs: false, hasKeyField: false, hasTestButton: false, hasProviders: false };
        }
        btn.click();
        await new Promise(r => setTimeout(r, 900));
        const t = document.body.innerText;
        // Заголовок модального окна — «Настройки» (SettingsModal.tsx).
        // Строка «Настройки AI» существовала только как всплывающая подсказка
        // кнопки в строке состояния и в innerText не попадает.
        // Признак открытия — заголовок окна вместе с его содержимым.
        const hasHeading = t.includes('Настройки');
        return {
          buttonFound: true,
          opened: hasHeading && t.includes('Проверить подключение'),
          hasTabs: t.includes('Провайдер') && t.includes('Модель') && t.includes('Промт') && t.includes('Хранилище'),
          hasKeyField: !!document.querySelector('input[type=password]'),
          hasTestButton: t.includes('Проверить подключение'),
          hasProviders: t.includes('OpenAI') && t.includes('Anthropic') && t.includes('Gemini') &&
                         t.includes('DeepSeek') && t.includes('OpenRouter') && t.includes('Ollama'),
        };
      })()
    `);

    /* ---- 6. Баннер обновления появляется сам ---- */
    // Ни одного клика по «Проверить обновление»: версия приходит из заглушки
    // update:check, и рендерер обязан показать баннер после автопроверки.
    out.updateBanner = await win.webContents.executeJavaScript(`
      (async () => {
        window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
        await new Promise(r => setTimeout(r, 300));
        const deadline = Date.now() + 20000;
        let banner = null;
        while (Date.now() < deadline) {
          banner = document.querySelector('[data-update-banner]');
          if (banner) break;
          await new Promise(r => setTimeout(r, 250));
        }
        const buttons = banner ? [...banner.querySelectorAll('button')] : [];
        return {
          appeared: Boolean(banner),
          mentionsVersion: banner ? banner.textContent.includes('9.9.9') : false,
          hasUpdateButton: buttons.some((b) => /Обновить/.test(b.textContent || '')),
          text: banner ? banner.textContent.replace(/\\s+/g, ' ').trim().slice(0, 120) : '',
        };
      })()
    `);
    delete out.updateBanner.text;

    out.errors = [];
  } catch (err) {
    out.fatal = err.stack || err.message;
  }

  fs.writeFileSync(path.join(ROOT, 'feature-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  app.exit(out.fatal || out.errors.length > 0 ? 1 : 0);
});
