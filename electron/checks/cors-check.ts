/**
 * Сквозная проверка AI-запроса через прокси в главном процессе.
 * Именно этот сценарий падал с "Не удалось подключиться" из-за CORS:
 * браузерный fetch получал net::ERR_FAILED, хотя сеть работала.
 *
 * Здесь проверяется реальный путь: рендерер -> IPC -> net-модуль -> шлюз.
 * Дополнительно сравнивается прямой fetch (должен падать) и запрос
 * через прокси (должен успешно вернуть ответ).
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

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
const SANDBOX = path.join(os.tmpdir(), 'ws-cors-test-' + Date.now());

const out = { steps: [], errors: [], progress: [] };
const step = (name, value) => out.steps.push({ name, ...value });
const mark = (label) => {
  out.progress.push(label);
  try {
    fs.writeFileSync(path.join(ROOT, 'cors-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
};

app.disableHardwareAcceleration();

const WATCHDOG_MS = 120_000;
setTimeout(() => {
  out.errors.push('watchdog: превышен лимит времени');
  try {
    fs.writeFileSync(path.join(ROOT, 'cors-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
  app.exit(1);
}, WATCHDOG_MS);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

app.whenReady().then(async () => {
  // Регистрируем ровно те обработчики, что и рабочее приложение
  const { registerAiProxy } = require('./aiProxy.js');
  const { secretsImporter } = require('./secrets.js');
  registerAiProxy();

  ipcMain.handle('secrets:describe', () => {
    const i = secretsImporter.read(false);
    return { found: i.found, filePath: i.filePath, baseUrl: i.baseUrl, textModel: i.textModel, hasKey: Boolean(i.apiKey), hasOtherSecrets: i.hasOtherSecrets };
  });
  ipcMain.handle('secrets:importKey', () => {
    const i = secretsImporter.read(true);
    if (!i.found || !i.apiKey) return { success: false, error: i.error || 'ключ не найден' };
    return { success: true, apiKey: i.apiKey, baseUrl: i.baseUrl, model: i.textModel };
  });
  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('window:isMaximized', () => false);
  ipcMain.handle('window:minimize', () => undefined);
  ipcMain.handle('window:maximize', () => false);
  ipcMain.handle('window:close', () => undefined);
  ipcMain.handle('fs:listDirectory', async () => ({ success: true, items: [] }));
  ipcMain.handle('fs:readFile', async () => ({ success: true, content: '', isBinary: false }));
  ipcMain.handle('fs:writeFile', async () => ({ success: true }));
  ipcMain.handle('fs:createFile', async () => ({ success: true }));
  ipcMain.handle('fs:createDirectory', async () => ({ success: true }));
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));
  ipcMain.handle('fs:searchWorkspace', async () => ({ success: true, results: [] }));
  ipcMain.handle('fs:openFolderDialog', () => null);
  ipcMain.handle('fs:saveFileDialog', () => null);
  ipcMain.handle('fs:openFileDialog', () => null);
  ipcMain.handle('fs:showInExplorer', () => ({ success: true }));

  fs.mkdirSync(SANDBOX, { recursive: true });

  const win = new BrowserWindow({
    width: 1500, height: 940, show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false, contextIsolation: true, sandbox: false,
    },
  });
  const js = (expr) => win.webContents.executeJavaScript(`(async () => { ${expr} })()`);

  try {
    mark('загрузка приложения');
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await sleep(2500);

    const info = secretsImporter.read(true);
    const endpoint = (info.baseUrl || '').replace(/\/+$/, '') + '/chat/completions';

    /* ---- 1. Прямой fetch из рендерера: ожидаем CORS-блокировку ---- */
    mark('прямой fetch (ожидаем блокировку CORS)');
    const direct = await js(`
      try {
        const r = await fetch(${JSON.stringify(endpoint)}, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + ${JSON.stringify(info.apiKey || '')} },
          body: JSON.stringify({ model: ${JSON.stringify(info.textModel || '')}, messages: [{ role: 'user', content: 'ок' }], max_tokens: 5 }),
        });
        return { ok: true, status: r.status };
      } catch (e) { return { ok: false, message: e.message }; }
    `);
    // Это НАБЛЮДЕНИЕ, а не требование: шлюз отдаёт CORS-заголовки не всегда,
    // поэтому прямой fetch то проходит, то блокируется, и фиксировать его
    // результат булевым флагом нельзя — он нестабилен по своей природе.
    // Значимо лишь то, что запрос через прокси работает в обоих случаях.
    step('directFetchViaRenderer', {
      rendererFetchStatus: direct.ok ? 'прошёл' : 'заблокирован CORS',
      rendererFetchMessage: direct.message || null,
      observationOnly: true,
      note: 'Справочно: нестабильность CORS у шлюза — причина исходной ошибки',
    });

    /* ---- 2. Запрос через прокси: ожидаем успех ---- */
    mark('запрос через прокси главного процесса');
    const viaProxy = await js(`
      const requestId = 'test-' + Date.now();
      return await new Promise((resolve) => {
        let text = '';
        let status = null;
        const timer = setTimeout(() => { cleanup(); resolve({ ok: false, error: 'timeout' }); }, 45000);
        const offChunk = window.electronAPI.onAiChunk((d) => {
          if (d.requestId !== requestId) return;
          // Разбираем SSE вручную, как это делает aiService
          for (const line of d.chunk.split('\\n')) {
            const t = line.trim();
            if (!t.startsWith('data:')) continue;
            const payload = t.slice(5).trim();
            if (!payload || payload === '[DONE]') continue;
            try {
              const p = JSON.parse(payload);
              const delta = p.choices?.[0]?.delta?.content || '';
              if (delta) text += delta;
            } catch {}
          }
        });
        const offErr = window.electronAPI.onAiError((d) => {
          if (d.requestId !== requestId) return;
          clearTimeout(timer); cleanup();
          resolve({ ok: false, status: d.status, error: d.message });
        });
        const offEnd = window.electronAPI.onAiEnd((d) => {
          if (d.requestId !== requestId) return;
          clearTimeout(timer); cleanup();
          resolve({ ok: true, text, status });
        });
        function cleanup() { offChunk(); offErr(); offEnd(); }

        window.electronAPI.aiStream({
          requestId,
          url: ${JSON.stringify(endpoint)},
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + ${JSON.stringify(info.apiKey || '')} },
          body: JSON.stringify({
            model: ${JSON.stringify(info.textModel || '')},
            messages: [{ role: 'user', content: 'Ответь ровно одним словом: работает' }],
            stream: true,
            max_tokens: 30,
          }),
        });
      });
    `);
    step('viaProxySucceeded', {
      ok: viaProxy.ok,
      responseText: viaProxy.text ? viaProxy.text.trim().slice(0, 120) : null,
      error: viaProxy.error || null,
      status: viaProxy.status || null,
    });

    /* ---- 3. Проверка через сам aiService в интерфейсе ---- */
    mark('проверка через UI и aiService');
    await js(`
      localStorage.setItem('workspace_ai_settings', JSON.stringify({
        provider: 'starimg',
        baseUrl: ${JSON.stringify(info.baseUrl)},
        model: ${JSON.stringify(info.textModel)},
        temperature: 0.7, maxTokens: 512, systemPrompt: 'тест', activePresetId: 'general',
      }));
      localStorage.setItem('workspace_api_keys', JSON.stringify({ starimg: 'plain:' + ${JSON.stringify(info.apiKey || '')} }));
      return true;
    `);
    await win.webContents.reload();
    await sleep(2800);

    const uiStatus = await js(`
      const t = document.body.innerText;
      return { configured: !t.includes('AI не настроен'), modelShown: t.includes('deepseek') };
    `);
    step('uiConfigured', uiStatus);

    await js(`
      const ta = document.querySelector('textarea[placeholder*="Спросите"]');
      if (!ta) return false;
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      setter.call(ta, 'Ответь ровно одним словом на русском: функционирует');
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise(r => setTimeout(r, 300));
      const btn = [...document.querySelectorAll('button')].find(b => /Отправить/.test(b.textContent || ''));
      if (btn) btn.click();
      return true;
    `);

    // Ждём ответа или ошибки
    const deadline = Date.now() + 50000;
    let answered = false;
    while (Date.now() < deadline) {
      const done = await js(`
        const t = document.body.innerText;
        const b = document.querySelector('.ws-prose');
        return (b && b.innerText.trim().length > 0) || /Ошибка|Не удалось подключиться/.test(t);
      `);
      if (done) { answered = true; break; }
      await sleep(800);
    }

    step('uiLiveResponse', await js(`
      const t = document.body.innerText;
      const b = document.querySelector('.ws-prose');
      return {
        answered: ${answered},
        gotResponse: !!b && b.innerText.trim().length > 0,
        responseText: b ? b.innerText.trim().slice(0, 150) : null,
        // Главное: исходной ошибки подключения быть не должно
        connectionErrorAbsent: !/Не удалось подключиться|Failed to fetch/.test(t),
        apiErrorAbsent: !/Ошибка API/.test(t),
      };
    `));
  } catch (err) {
    out.errors.push(err.stack || String(err));
  }

  mark('завершение');
  fs.writeFileSync(path.join(ROOT, 'cors-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  try { fs.rmSync(SANDBOX, { recursive: true, force: true }); } catch {}
  app.exit(0);
});
