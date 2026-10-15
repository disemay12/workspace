/**
 * Сквозная проверка с реальным AI-endpoint из secrets.local.ps1:
 *  1. Модуль secrets читает ключ из локального файла
 *  2. IPC-обработчики describeSecrets / importSecretsKey работают
 *  3. Импортированный ключ реально используется приложением для запроса
 *  4. Стриминг ответа отображается в интерфейсе
 *
 * Ключ в отчёт не попадает — фиксируется только факт его наличия и длина.
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { registerCommonCheckIpc } = require('./common-ipc');

/**
 * Корень проекта. Скрипт компилируется в dist-electron/, поэтому путь
 * определяется поиском каталога с package.json и dist/ — это корректно
 * независимо от глубины вложенности исходника.
 */
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
const SANDBOX = path.join(os.tmpdir(), 'ws-live-test-' + Date.now());

const out: any = { steps: [], errors: [], progress: [] };
const step = (name, value) => out.steps.push({ name, ...value });
/** Пишем прогресс на диск после каждого шага — чтобы при зависании
 *  было видно, на чём именно остановился сценарий. */
const mark = (label) => {
  out.progress.push(`${new Date().toISOString().slice(11, 19)} ${label}`);
  try {
    fs.writeFileSync(path.join(ROOT, 'live-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
};

app.disableHardwareAcceleration();

/** Жёсткий предохранитель: приложение завершится даже при зависании сети. */
const WATCHDOG_MS = 90_000;
setTimeout(() => {
  out.errors.push('watchdog: сценарий превысил лимит времени');
  try {
    fs.writeFileSync(path.join(ROOT, 'live-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
  app.exit(1);
}, WATCHDOG_MS);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** Ожидание условия с ограничением по времени (вместо фиксированных пауз). */
async function waitFor(win, expression, timeoutMs = 20000, label = 'условие') {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const ok = await win.webContents.executeJavaScript(`(() => { try { return Boolean(${expression}); } catch { return false; } })()`);
      if (ok) return true;
    } catch {}
    await sleep(400);
  }
  out.progress.push(`TIMEOUT ожидания: ${label}`);
  return false;
}

/* ---------- 1. Модуль секретов ---------- */
function testSecretsModule() {
  const { secretsImporter } = require('./secrets.js');

  // Без ключа — метаданные
  const meta = secretsImporter.read(false);
  // С ключом
  const full = secretsImporter.read(true);

  step('secretsModule', {
    fileFound: meta.found,
    filePath: meta.filePath ? path.basename(meta.filePath) : null,
    hasBaseUrl: Boolean(meta.baseUrl),
    baseUrl: meta.baseUrl,
    textModel: meta.textModel,
    // В метаданных ключ должен быть замаскирован
    keyMaskedInMeta: meta.apiKey === '***',
    // При явном запросе — реальный ключ
    keyAvailableOnDemand: Boolean(full.apiKey) && full.apiKey.startsWith('sk-'),
    keyLength: full.apiKey ? full.apiKey.length : 0,
    detectsOtherSecrets: full.hasOtherSecrets,
  });

  return { baseUrl: full.baseUrl, apiKey: full.apiKey, model: full.textModel };
}

/* ---------- IPC для файловой системы ---------- */
// Регистрируем прокси AI — приложение выполняет запросы через него
function registerHandlers(): any {
  const { secretsImporter } = require('./secrets.js');

  ipcMain.handle('secrets:describe', () => {
    const info = secretsImporter.read(false);
    return {
      found: info.found, filePath: info.filePath, baseUrl: info.baseUrl,
      textModel: info.textModel, imageModel: info.imageModel,
      hasKey: Boolean(info.apiKey), hasOtherSecrets: info.hasOtherSecrets, error: info.error,
    };
  });
  ipcMain.handle('secrets:importKey', () => {
    const info = secretsImporter.read(true);
    if (!info.found || !info.apiKey) return { success: false, error: info.error || 'ключ не найден' };
    return { success: true, apiKey: info.apiKey, baseUrl: info.baseUrl, model: info.textModel };
  });

  // Вызывается ПОСЛЕ своих secrets:*, иначе заглушка перекрыла бы реальную
  // реализацию и повторная регистрация падала бы с «second handler».
  // Helper добавляет только недостающие каналы (избранное, буфер обмена).
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
  ipcMain.handle('fs:readFile', async () => ({ success: true, content: '', isBinary: false }));
  ipcMain.handle('fs:writeFile', async () => ({ success: true }));
  ipcMain.handle('fs:createFile', async () => ({ success: true }));
  ipcMain.handle('fs:createDirectory', async () => ({ success: true }));
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));
  ipcMain.handle('fs:searchWorkspace', async () => ({ success: true, results: [] }));
}

app.whenReady().then(async () => {
  let secrets = null;
  try {
    secrets = testSecretsModule();
  } catch (err) {
    out.errors.push('secretsModule: ' + err.message);
  }

  registerHandlers();

  // AI-запросы идут через главный процесс (обход CORS)

  require('./aiProxy.js').registerAiProxy();
  fs.mkdirSync(SANDBOX, { recursive: true });

  const win = new BrowserWindow({
    width: 1500, height: 940, show: false,
    webPreferences: { preload: path.join(__dirname, 'preload.js'), nodeIntegration: false, contextIsolation: true, sandbox: false },
  });
  const js = (expr) => win.webContents.executeJavaScript(`(async () => { ${expr} })()`);

  try {
    mark('загрузка приложения');
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await waitFor(win, `document.getElementById('root')?.children.length > 0`, 15000, 'монтирование React');

    /* ---------- 2. IPC из рендерера ---------- */
    mark('describeSecrets');
    const described = await js(`return await window.electronAPI.describeSecrets()`);
    step('ipcDescribeSecrets', {
      found: described.found,
      hasBaseUrl: Boolean(described.baseUrl),
      hasKey: described.hasKey,
      // Ключ не должен утекать в метаданных
      keyNotExposed: described.apiKey === undefined,
      hasOtherSecrets: described.hasOtherSecrets,
    });

    /* ---------- 3. Импорт ключа через UI ---------- */
    mark('открытие настроек');
    await js(`
      // Кнопка настроек находится в строке состояния, а не в заголовке окна.
      const btn = [...document.querySelectorAll('button')].find(b => /Настройки/i.test(b.title || ''));
      if (btn) btn.click();
      return true;
    `);
    // Заголовок модального окна — «Настройки» (SettingsModal.tsx);
    // «Настройки AI» — только всплывающая подсказка кнопки в строке состояния.
    await waitFor(win, `document.body.innerText.includes('Проверить подключение')`, 10000, 'модалка настроек');

    step('settingsShowsSecrets', await js(`
      const t = document.body.innerText;
      return {
        modalOpen: t.includes('Настройки') && t.includes('Проверить подключение'),
        secretsBlockShown: t.includes('Найден файл с ключами'),
        importButtonPresent: !!([...document.querySelectorAll('button')].find(b => /Импортировать/.test(b.textContent || ''))),
        warnsOtherSecrets: t.includes('Notion') || t.includes('Telegram'),
      };
    `));

    // Нажимаем «Импортировать»
    mark('импорт ключа');
    const imported = await js(`
      const btn = [...document.querySelectorAll('button')].find(b => /Импортировать/.test(b.textContent || ''));
      if (!btn) return false;
      btn.click();
      return true;
    `);
    await waitFor(win, `document.body.innerText.includes('Ключ импортирован')`, 15000, 'сообщение об импорте');

    step('importResult', await js(`
      const t = document.body.innerText;
      return {
        clicked: ${imported},
        successMessage: /Ключ импортирован/.test(t),
        messageSnippet: (t.match(/Ключ импортирован[^\\n]*/) || [''])[0],
        keyFieldFilled: (() => {
          const inp = document.querySelector('input[type=password], input[type=text].font-mono');
          return inp ? inp.value.length > 20 : false;
        })(),
      };
    `));

    // Закрываем настройки
    mark('закрытие настроек');
    await js(`
      const btn = [...document.querySelectorAll('button')].find(b => (b.textContent || '').trim() === 'Закрыть');
      if (btn) btn.click();
      return true;
    `);
    await sleep(800);

    step('statusAfterImport', await js(`
      const t = document.body.innerText;
      return {
        aiConfigured: !t.includes('AI не настроен'),
        modelShown: t.includes('deepseek'),
      };
    `));

    /* ---------- 4. Реальный запрос к нейросети ---------- */
    mark('отправка запроса к нейросети');
    const sent = await js(`
      const ta = document.querySelector('textarea[placeholder*="Спросите"]');
      if (!ta) return false;
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      setter.call(ta, 'Ответь ровно одним словом на русском: работает');
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise(r => setTimeout(r, 300));
      const send = [...document.querySelectorAll('button')].find(b => /Отправить/.test(b.textContent || ''));
      if (send) { send.click(); return true; }
      return false;
    `);

    // Ждём появления ответа (или ошибки) вместо фиксированной паузы
    const answered = await waitFor(
      win,
      `(() => {
         const t = document.body.innerText;
         const bubble = document.querySelector('.ws-prose');
         return (bubble && bubble.innerText.trim().length > 0) || /Ошибка API|Не удалось подключиться/.test(t);
       })()`,
      45000,
      'ответ нейросети'
    );

    step('liveAiResponse', await js(`
      const t = document.body.innerText;
      const bubble = document.querySelector('.ws-prose');
      return {
        requestSent: ${sent},
        answeredWithinTimeout: ${answered},
        gotResponse: !!bubble && bubble.innerText.trim().length > 0,
        responseText: bubble ? bubble.innerText.trim().slice(0, 200) : null,
        noError: !/Ошибка API|Не удалось подключиться/.test(t),
        streamFinished: !/Думает/.test(t),
      };
    `));
  } catch (err) {
    out.errors.push(err.stack || String(err));
  }

  mark('завершение');
  fs.writeFileSync(path.join(ROOT, 'live-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  try { fs.rmSync(SANDBOX, { recursive: true, force: true }); } catch {}
  app.exit(0);
});
