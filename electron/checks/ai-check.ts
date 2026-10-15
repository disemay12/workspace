/**
 * Проверка AI-подсистемы и цикла сохранения.
 *
 * 1. Поднимает локальный HTTP-сервер, эмулирующий OpenAI-совместимый
 *    streaming API (SSE), и проверяет полный путь:
 *    UI -> aiService.streamChat -> fetch -> парсинг SSE -> отрисовка.
 * 2. Проверяет кнопку «Применить» инлайн-копилота (Ctrl+K) и наличие diff.
 * 3. Проверяет цикл редактирование -> Ctrl+S -> запись на диск -> сброс dirty.
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
const SANDBOX = path.join(os.tmpdir(), 'ws-ai-test-' + Date.now());

const out: any = { steps: [], errors: [], requests: [] };
const step = (name, value) => out.steps.push({ name, ...value });

app.disableHardwareAcceleration();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Ждёт появления текста в окне, опрашивая страницу.
 *
 * Нужен вместо фиксированной паузы после reload: настройки подгружаются
 * асинхронно, и при медленном старте проверка успевала прочитать значения
 * по умолчанию. Возвращает true, если текст появился до истечения таймаута.
 */
async function waitForText(win, needle, timeoutMs = 8000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const found = await win.webContents.executeJavaScript(
        'document.body.innerText.includes(' + JSON.stringify(needle) + ')'
      );
      if (found) return true;
    } catch {
      // страница ещё перезагружается — пробуем снова
    }
    await sleep(250);
  }
  return false;
}

/* ============================================================
   Мок OpenAI-совместимого streaming API
   ============================================================ */
function startMockServer(): any {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      let body = '';
      req.on('data', (c) => (body += c));
      req.on('end', () => {
        let parsed = null;
        try {
          parsed = JSON.parse(body);
        } catch {}

        out.requests.push({
          url: req.url,
          method: req.method,
          auth: req.headers.authorization || null,
          model: parsed?.model ?? null,
          stream: parsed?.stream ?? null,
          temperature: parsed?.temperature ?? null,
          systemPromptPresent: Array.isArray(parsed?.messages) && parsed.messages[0]?.role === 'system',
          systemPromptSnippet: parsed?.messages?.[0]?.content?.slice(0, 60) ?? null,
          messageCount: parsed?.messages?.length ?? 0,
          lastUserContent: parsed?.messages?.[parsed.messages.length - 1]?.content?.slice(0, 200) ?? null,
        });

        // Проверка авторизации: без ключа возвращаем 401
        if (req.url.includes('/unauthorized')) {
          res.writeHead(401, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: { message: 'Invalid API key' } }));
          return;
        }

        res.writeHead(200, {
          'Content-Type': 'text/event-stream',
          'Cache-Control': 'no-cache',
          Connection: 'keep-alive',
        });

        // Стримим ответ по частям, как настоящий API
        const chunks = ['Привет', ', ', 'это ', '**тестовый** ', 'ответ.\n\n', '```js\n', 'const x = 42;\n', '```\n'];
        let i = 0;
        const timer = setInterval(() => {
          if (i >= chunks.length) {
            clearInterval(timer);
            res.write('data: [DONE]\n\n');
            res.end();
            return;
          }
          const payload = { choices: [{ delta: { content: chunks[i] } }] };
          res.write(`data: ${JSON.stringify(payload)}\n\n`);
          i++;
        }, 25);
      });
    });

    server.listen(0, '127.0.0.1', () => {
      const port = server.address().port;
      resolve({ server, port, baseUrl: `http://127.0.0.1:${port}/v1` });
    });
  });
}

/**
 * Настройки AI, которые «постоянное хранилище» вернёт рендереру.
 *
 * Заполняются после запуска mock-сервера, когда известен его адрес, но
 * передаются в IPC по ссылке: helper читает объект в момент вызова
 * aiSettings:load, то есть уже после того, как адрес подставлен.
 */
const seededAiSettings: { settings: Record<string, unknown>; keys: Record<string, string> } = {
  settings: {},
  keys: {},
};

/** Регистрируем IPC-обработчики файловой системы. */
// Регистрируем прокси AI — приложение выполняет запросы через него
function registerHandlers(): any {
  registerCommonCheckIpc({
    sandbox: SANDBOX,
    // Геттеры, а не значения: сценарий меняет настройки между фазами, и helper
    // должен видеть актуальные, а не снимок на момент регистрации.
    getAiSettings: () => seededAiSettings.settings,
    getAiKeys: () => seededAiSettings.keys,
  });

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

  ipcMain.handle('fs:writeFile', async (_e, filePath, content, encoding = 'utf-8') => {
    try {
      if (encoding === 'base64') await fs.promises.writeFile(filePath, Buffer.from(content, 'base64'));
      else await fs.promises.writeFile(filePath, content, 'utf-8');
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
  fs.writeFileSync(
    path.join(SANDBOX, 'test.md'),
    '# Заголовок\n\nИсходный текст для проверки AI.\n',
    'utf-8'
  );

  const mock = await startMockServer();

  // Постоянное хранилище настроек важнее localStorage, поэтому те же значения,
  // что сценарий кладёт в localStorage, задаются и здесь — иначе aiSettings:load
  // вернёт пустой объект и затрёт модель из localStorage.
  seededAiSettings.settings = {
    provider: 'custom',
    baseUrl: mock.baseUrl,
    model: 'test-model',
    temperature: 0.5,
    maxTokens: 1024,
    systemPrompt: 'Ты тестовый ассистент.',
    activePresetId: 'general',
  };
  seededAiSettings.keys = { custom: 'test-key-123' };

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

  const js = (expr) => win.webContents.executeJavaScript(`(async () => { ${expr} })()`);

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await sleep(2500);

    /* ============ 1. Настраиваем AI через localStorage ============ */
    await js(`
      localStorage.setItem('workspace_ai_settings', JSON.stringify({
        provider: 'custom',
        baseUrl: ${JSON.stringify(mock.baseUrl)},
        model: 'test-model',
        temperature: 0.5,
        maxTokens: 1024,
        systemPrompt: 'Ты тестовый ассистент.',
        activePresetId: 'general'
      }));
      return true;
    `);

    // Ключ кладём напрямую в защищённое хранилище (эмуляция сохранения пользователем)
    await js(`
      localStorage.setItem('workspace_api_keys', JSON.stringify({
        custom: 'plain:test-key-123'
      }));
      return true;
    `);

    await win.webContents.reload();

    // Ждём завершения гидратации настроек: после перезагрузки приложение сначала
    // рисует значения по умолчанию, а затем подменяет их данными из хранилища.
    // Фиксированная пауза делала проверку плавающей, поэтому ждём по условию.
    await waitForText(win, 'test-model', 10000);

    step('aiSettingsLoaded', await js(`
      const status = document.body.innerText;
      return {
        // В статусной строке должна отобразиться модель из настроек
        modelShown: status.includes('test-model'),
        notConfigured: status.includes('AI не настроен'),
      };
    `));

    /* ============ 2. Отправляем сообщение в AI-чат ============ */
    step('chatBeforeSend', await js(`
      const ta = document.querySelector('textarea[placeholder*="Спросите"]');
      return { hasInput: !!ta, hasSendButton: !!document.querySelector('button') };
    `));

    // Вводим текст и отправляем
    await js(`
      const ta = document.querySelector('textarea[placeholder*="Спросите"]');
      if (!ta) return false;
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      setter.call(ta, 'Проверка потокового ответа');
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise(r => setTimeout(r, 200));
      const sendBtn = [...document.querySelectorAll('button')].find(b => /Отправить/.test(b.textContent || ''));
      if (sendBtn) sendBtn.click();
      return !!sendBtn;
    `);

    // Ждём завершения стрима
    await sleep(3000);

    step('chatAfterStream', await js(`
      const t = document.body.innerText;
      return {
        userMessageShown: t.includes('Проверка потокового ответа'),
        assistantResponded: t.includes('Привет') && t.includes('тестовый'),
        hasCodeBlock: !!document.querySelector('.ws-prose pre code'),
        codeContent: document.querySelector('.ws-prose pre code')?.textContent?.trim() ?? null,
        hasCopyAction: /Копировать/.test(t),
        // Кнопка «Заменить документ» появляется только при открытом документе,
        // поэтому здесь проверяем лишь доступность действий копирования.
        hasCopyCapability: !!document.querySelector('button') && /Копировать/.test(t),
        stillStreaming: t.includes('Думает'),
      };
    `));

    step('mockApiRequest', {
      requestCount: out.requests.length,
      first: out.requests[0] ?? null,
    });

    /* ============ 3. Проверка ошибки API (401) ============ */
    // Постоянное хранилище приоритетнее localStorage, поэтому фаза 401 должна
    // обновить и его — иначе останется рабочий URL и запрос уйдёт на реальный API.
    seededAiSettings.settings = {
      provider: 'custom',
      baseUrl: mock.baseUrl.replace('/v1', '/v1/unauthorized'),
      model: 'test-model',
      temperature: 0.5,
      maxTokens: 512,
      systemPrompt: 'тест',
      activePresetId: 'general',
    };
    await js(`
      localStorage.setItem('workspace_ai_settings', JSON.stringify({
        provider: 'custom',
        baseUrl: ${JSON.stringify(mock.baseUrl.replace('/v1', '/v1/unauthorized'))},
        model: 'test-model',
        temperature: 0.5, maxTokens: 512,
        systemPrompt: 'тест', activePresetId: 'general'
      }));
      return true;
    `);
    await win.webContents.reload();
    // Ждём гидратацию настроек: без неё следующий шаг работает со значениями
    // по умолчанию и запрос уходит не туда.
    await waitForText(win, 'test-model', 10000);

    await js(`
      const ta = document.querySelector('textarea[placeholder*="Спросите"]');
      if (!ta) return false;
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      setter.call(ta, 'запрос с неверным ключом');
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise(r => setTimeout(r, 200));
      const sendBtn = [...document.querySelectorAll('button')].find(b => /Отправить/.test(b.textContent || ''));
      if (sendBtn) sendBtn.click();
      return true;
    `);
    await sleep(2500);

    step('apiErrorHandling', await js(`
      const t = document.body.innerText;
      return {
        errorShown: /401|Invalid API key|Ошибка API/i.test(t),
        suggestsKeyCheck: /ключ/i.test(t),
        snippet: t.match(/.{0,120}(401|Invalid API key|Ошибка API).{0,120}/i)?.[0] ?? null,
      };
    `));

    /* ============ 4. Инлайн-копилот Ctrl+K ============ */
    // Возвращаем рабочий URL — и в localStorage, и в постоянном хранилище,
    // потому что именно последнее определяет настройки после гидратации.
    seededAiSettings.settings = {
      provider: 'custom',
      baseUrl: mock.baseUrl,
      model: 'test-model',
      temperature: 0.5,
      maxTokens: 512,
      systemPrompt: 'тест',
      activePresetId: 'general',
    };
    await js(`
      localStorage.setItem('workspace_ai_settings', JSON.stringify({
        provider: 'custom',
        baseUrl: ${JSON.stringify(mock.baseUrl)},
        model: 'test-model', temperature: 0.5, maxTokens: 512,
        systemPrompt: 'тест', activePresetId: 'general'
      }));
      return true;
    `);
    await win.webContents.reload();
    // Ключ к inline-шагам: панель AI-правки отправит запрос только когда
    // приложение считает себя настроенным, то есть после гидратации.
    await waitForText(win, 'test-model', 10000);

    // Гарантируем открытую рабочую папку: если после перезагрузок приложение осталось
    // на стартовом экране, нажимаем «Открыть папку» (диалог в сценарии возвращает SANDBOX).
    await js(`
      const t = document.body.innerText;
      if (!t.includes('test.md')) {
        const btn = [...document.querySelectorAll('button')]
          .find(b => (b.textContent || '').trim() === 'Открыть папку'
                  || (b.textContent || '').trim() === 'Открыть папку сейчас'
                  || (b.title || '').indexOf('Открыть рабочую папку') !== -1);
        if (btn) btn.click();
      }
      return true;
    `);
    await sleep(2000);

    // Открываем Markdown-файл из дерева.
    // Строка файла содержит иконку и размер, поэтому textContent не равен имени:
    // ищем по атрибуту title, который начинается с имени файла.
    const fileOpened = await js(`
      const node = [...document.querySelectorAll('div')]
        .find(d => {
          const title = d.getAttribute('title') || '';
          return title.split(' — ')[0] === 'test.md'
            && d.className.includes('cursor-pointer');
        });
      if (node) node.click();
      return !!node;
    `);
    await sleep(2000);

    step('fileOpenedForInline', await js(`
      const t = document.body.innerText;
      return {
        opened: t.includes('test.md'),
        // Признак реально открытого редактора — его панель инструментов,
        // а не упоминание имени файла в дереве или списке горячих клавиш.
        // Редактор Markdown использует <textarea>, поэтому наличие только
        // поля чата не считается признаком открытого документа: проверяем,
        // что текстовых полей больше одного.
        editorAreaPresent: document.querySelectorAll('textarea').length > 1,
        editorToolbarPresent: [...document.querySelectorAll('button')]
          .some(b => (b.title || '').indexOf('Жирный') !== -1),
        hasInlineButton: [...document.querySelectorAll('button')]
          .some(b => (b.title || '').indexOf('AI-правка') !== -1),
      };
    `));

    // Нажимаем кнопку AI-правка в тулбаре редактора
    // Кнопка содержит иконку и текст, поэтому опознаём её по атрибуту title.
    const inlineOpened = await js(`
      const btn = [...document.querySelectorAll('button')]
        .find(b => (b.title || '').indexOf('AI-правка') !== -1);
      if (!btn) return false;
      btn.click();
      return true;
    `);
    await sleep(1200);

    step('inlineCopilotOpen', await js(`
      const t = document.body.innerText;
      return {
        buttonFound: ${inlineOpened},
        panelOpen: t.includes('AI-правка') && (t.includes('Сгенерировать') || t.includes('Что сделать')),
        hasQuickActions: t.includes('Исправить ошибки') && t.includes('Сократить'),
        hasTranslate: t.includes('Перевести на английский'),
        hasExplain: t.includes('Объяснить'),
        hasContinue: t.includes('Продолжить текст'),
        hasGenerateButton: /Сгенерировать/.test(t),
        hasApplyButton: /Применить/.test(t),
      };
    `));

    // Запускаем генерацию через быструю команду
    const quickClicked = await js(`
      // Быстрая команда панели AI-правки. Кнопка таблиц «Исправить ошибки формул»
      // содержит похожий текст, поэтому ищем точное совпадение подписи.
      const target = [...document.querySelectorAll('button')]
        .find(b => (b.textContent || '').trim() === 'Исправить ошибки и стиль');
      if (target) target.click();
      return Boolean(target);
    `);
    out.quickAction = quickClicked;
    await sleep(3000);

    step('inlineDiffRendered', await js(`
      const t = document.body.innerText;
      const diffContainer = [...document.querySelectorAll('div')].find(d =>
        d.className && String(d.className).includes('font-mono') && d.children.length > 2);
      return {
        // Заголовок панели: «Генерация…» во время стрима либо «Предпросмотр изменений»
        previewShown: /Предпросмотр изменений/i.test(t),
        generationOrPreview: /Предпросмотр изменений|Генерация/i.test(t),
        applyEnabled: (() => {
          const b = [...document.querySelectorAll('button')].find(x => /Применить/.test(x.textContent || ''));
          return b ? !b.disabled : false;
        })(),
        diffHasMarkers: diffContainer ? diffContainer.innerText.length > 0 : false,
        proposalVisible: t.includes('Привет') || t.includes('тестовый'),
      };
    `));

    // Применяем изменения
    const applied = await js(`
      const btn = [...document.querySelectorAll('button')].find(b => /Применить/.test(b.textContent || ''));
      if (btn && !btn.disabled) { btn.click(); return true; }
      return false;
    `);
    await sleep(1200);

    step('inlineEditApplied', await js(`
      const t = document.body.innerText;
      const ta = document.querySelector('textarea');
      return {
        applyClicked: ${applied},
        contentChanged: ta ? ta.value.includes('Привет') || ta.value.includes('тестовый') : false,
        dirtyIndicator: t.includes('Не сохранено') || t.includes('Изменено'),
        textareaValue: ta ? ta.value.slice(0, 120) : null,
      };
    `));

    /* ============ 5. Сохранение (Ctrl+S) ============ */
    const beforeSave = fs.readFileSync(path.join(SANDBOX, 'test.md'), 'utf-8');

    await js(`
      // Ставим фокус в редактор и нажимаем Ctrl+S
      const ta = document.querySelector('textarea');
      if (ta) ta.focus();
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true }));
      return true;
    `);
    await sleep(1500);

    const afterSave = fs.readFileSync(path.join(SANDBOX, 'test.md'), 'utf-8');

    step('saveRoundTrip', await js(`
      const t = document.body.innerText;
      return {
        fileChangedOnDisk: ${JSON.stringify(beforeSave)} !== ${JSON.stringify(afterSave)},
        contentOnDisk: ${JSON.stringify(afterSave.slice(0, 120))},
        dirtyCleared: !/Не сохранено/.test(t),
        savedToast: /Сохранено/.test(t),
      };
    `));

    /* ============ 6. Проверка настроек: смена провайдера ============ */
    await js(`
      // Кнопка настроек живёт в строке состояния (StatusBar), а не в заголовке окна:
      // в .titlebar-no-drag её нет, поэтому поиск идёт по всему документу.
      const btn = [...document.querySelectorAll('button')].find(b => /Настройки/i.test(b.title || ''));
      if (btn) btn.click();
      return true;
    `);
    await sleep(900);

    await js(`
      const btn = [...document.querySelectorAll('button')].find(b => b.textContent?.includes('Anthropic'));
      if (btn) btn.click();
      return !!btn;
    `);
    await sleep(900);

    step('settingsProviderSwitch', await js(`
      const inputs = [...document.querySelectorAll('input')];
      const urlInput = inputs.find(i => i.value && i.value.includes('http'));
      return {
        baseUrlUpdated: urlInput ? urlInput.value : null,
        anthropicUrlApplied: urlInput ? urlInput.value.includes('anthropic') : false,
      };
    `));
  } catch (err) {
    out.errors.push(err.stack || String(err));
  }

  mock.server.close();
  fs.writeFileSync(path.join(ROOT, 'ai-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  try {
    fs.rmSync(SANDBOX, { recursive: true, force: true });
  } catch {}
  app.exit(0);
});
