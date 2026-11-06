/**
 * Проверка агентной работы ассистента с открытым .docx.
 *
 * Поднимает приложение с синтетическим резюме, подменяет OpenAI-совместимый
 * API скриптованными tool_calls и проверяет весь путь:
 * права → инструменты → живой редактор → откат → сохранение оформления.
 *
 * Запуск: electron.exe dist-electron/doc-agent-check.js
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const http = require('http');
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

const ROOT = process.env.WS_ROOT || resolveProjectRoot();
const SANDBOX = process.env.WS_CHECK_SANDBOX || path.join(ROOT, 'test-workspace-doc-agent');
const OUT = path.join(ROOT, 'doc-agent-check.json');
const DOC_NAME = 'Резюме.docx';
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const out = { steps: [], errors: [], requests: [] };
const step = (name, value) => {
  out.steps.push({ name, ...value });
  // Промежуточная запись: при зависании видно, какой шаг последним завершился.
  try {
    fs.writeFileSync(OUT, JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
};

fs.mkdirSync(SANDBOX, { recursive: true });
fs.mkdirSync(path.join(SANDBOX, 'userdata'), { recursive: true });
app.setPath('userData', path.join(SANDBOX, 'userdata'));
app.disableHardwareAcceleration();
try {
  fs.writeFileSync(path.join(SANDBOX, 'boot.txt'), 'module loaded');
} catch {}

function writeOut(reason?: string) {
  if (reason) out.errors.push(reason);
  try {
    fs.writeFileSync(OUT, JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
  console.log(JSON.stringify(out, null, 2));
}

/* ============================================================
   Синтетический .docx
   ============================================================ */

function buildDocx() {
  const paragraph = (text, options: any = {}) => {
    const props = [];
    if (options.style) props.push(`<w:pStyle w:val="${options.style}"/>`);
    if (options.list) props.push('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    if (options.align) props.push(`<w:jc w:val="${options.align}"/>`);
    const pPr = props.length ? `<w:pPr>${props.join('')}</w:pPr>` : '';
    return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  };

  const body = [
    paragraph('Иван Иванов', { style: 'Heading1' }),
    paragraph('Москва · +7 999 000-00-00 · ivan@example.com'),
    paragraph('Опыт работы', { style: 'Heading2' }),
    paragraph('Компания А — инженер, Июль 2024 — настоящее время'),
    paragraph('Навыки', { style: 'Heading2' }),
    paragraph('TypeScript', { list: true }),
    paragraph('React', { list: true }),
    paragraph('К сожалению, я не могу редактировать документ Word. Вот рекомендации по оформлению: используйте заголовки и списки.'),
    paragraph('Хотите, я создам отдельный файл с исправленным текстом?'),
  ].join('');

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

  const stylesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>' +
    '</w:styles>';

  const numberingXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
    '<w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

  const zip = new JSZip();
  zip.file('word/document.xml', documentXml);
  zip.file('word/styles.xml', stylesXml);
  zip.file('word/numbering.xml', numberingXml);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** Текст .docx без сторонних зависимостей — для проверки сохранения. */
async function readDocxXml(filePath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
  const xml = await zip.file('word/document.xml').async('string');
  const text = (xml.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [])
    .map((match) => match.replace(/<[^>]+>/g, ''))
    .join('');
  return { xml, text };
}

/* ============================================================
   Мок OpenAI-совместимого API с tool_calls
   ============================================================ */

const ATS_HTML = [
  '<h1>Иван Иванов</h1>',
  '<p style="text-align:left">Москва · +7 999 000-00-00 · ivan@example.com</p>',
  '<h3>Опыт работы</h3>',
  '<p><strong>Компания А</strong> — инженер, Июль 2024 — настоящее время</p>',
  '<h3>Навыки</h3>',
  '<ul><li>TypeScript</li><li>React</li></ul>',
  '<h3>Образование</h3>',
  '<p>МГУ, 2018</p>',
].join('');

const FALLBACK_HTML = [
  '<h1>Иван Иванов</h1>',
  '<p style="text-align:center">Москва · +7 999 000-00-00 · ivan@example.com</p>',
  '<h2>Опыт работы</h2>',
  '<p>Компания А — инженер, Июнь 2024 — настоящее время. Оформлено через document_apply.</p>',
  '<h2>Навыки</h2>',
  '<ul><li>TypeScript</li><li>React</li></ul>',
].join('');

function toolCall(name, args, id, index = 0) {
  return {
    choices: [
      {
        delta: {
          tool_calls: [
            {
              index,
              id,
              type: 'function',
              function: { name, arguments: JSON.stringify(args || {}) },
            },
          ],
        },
      },
    ],
  };
}

function contentChunk(text) {
  return { choices: [{ delta: { content: text } }] };
}

function scenarioOf(messages) {
  const user = [...messages].reverse().find((message) => message.role === 'user');
  const text = String(user?.content || '');
  if (text.includes('запасной протокол')) return 'fallback';
  if (text.includes('Перепиши резюме')) return 'norights';
  if (text.includes('не относится к резюме')) return 'cleanup';
  if (text.includes('ATS')) return 'ats';
  if (text.includes('H2') || text.includes('по центру')) return 'format';
  if (text.includes('Июль 2024')) return 'find';
  return 'default';
}

function lastStructure(messages) {
  const tool = [...messages].reverse().find((message) => message.role === 'tool');
  if (!tool) return [];
  try {
    const data = JSON.parse(String(tool.content));
    return String(data.structure || '')
      .split('\n')
      .map((line) => {
        const match = /^\s*(b\d+)\s+\[([^\]]+)\]\s*(.*)$/.exec(line);
        return match ? { id: match[1], tag: match[2], text: match[3] } : null;
      })
      .filter(Boolean);
  } catch {
    return [];
  }
}

function scriptedResponse(scenario, stepIndex, messages) {
  const lines = lastStructure(messages);

  if (scenario === 'cleanup') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'c-get')];
    if (stepIndex === 1) {
      return [
        toolCall('delete_block', { block_id: 'b9' }, 'c-del-1', 0),
        toolCall('delete_block', { block_id: 'b10' }, 'c-del-2', 1),
      ];
    }
    return [contentChunk('Убрал из документа всё, что не относится к резюме.')];
  }

  if (scenario === 'ats') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'a-get')];
    if (stepIndex === 1) {
      return [toolCall('replace_document', { html: ATS_HTML }, 'a-replace')];
    }
    return [contentChunk('Переписал резюме под ATS: одна колонка, стандартные заголовки, списки.')];
  }

  if (scenario === 'format') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'f-get')];
    if (stepIndex === 1) {
      const calls = [];
      for (const line of lines) {
        if (!/^h[1-6]$/.test(line.tag)) continue;
        if (!/(Опыт|Навыки|Образование)/i.test(line.text)) continue;
        calls.push(
          toolCall('format', { block_id: line.id, heading: '2' }, `f-h-${line.id}`, calls.length)
        );
      }
      const contact = lines.find((line) => line.tag === 'p' && line.text.includes('@'));
      if (contact) {
        calls.push(
          toolCall('format', { block_id: contact.id, align: 'center' }, 'f-center', calls.length)
        );
      }
      return calls.length > 0 ? calls : [contentChunk('Не нашёл заголовков для оформления.')];
    }
    return [contentChunk('Сделал заголовки разделов H2 и выровнял контакты по центру.')];
  }

  if (scenario === 'find') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'r-get')];
    if (stepIndex === 1) {
      return [
        toolCall('find_replace', { find: 'Июль 2024', replace: 'Июнь 2024', all: true }, 'r-one'),
      ];
    }
    return [contentChunk('Заменил «Июль 2024» на «Июнь 2024» в опыте работы.')];
  }

  if (scenario === 'norights') {
    return [
      contentChunk(
        'Прав на запись в документ нет — изменить его не могу. Нажмите «Предоставить права на запись» в панели AI.'
      ),
    ];
  }

  if (scenario === 'fallback') {
    const action = {
      action: 'document_apply',
      operations: [{ tool: 'replace_document', html: FALLBACK_HTML }],
    };
    return [
      contentChunk(
        'Оформляю документ через запасной протокол.\n\n```workspace-action\n' +
          JSON.stringify(action) +
          '\n```'
      ),
    ];
  }

  return [contentChunk('Готово.')];
}

function startMockServer(): Promise<any> {
  return new Promise((resolve) => {
    const counters = new Map();

    const server = http.createServer((req, res) => {
      let body = '';
      req.on('data', (chunk) => (body += chunk));
      req.on('end', () => {
        let parsed = null;
        try {
          parsed = JSON.parse(body);
        } catch {}

        const messages = Array.isArray(parsed?.messages) ? parsed.messages : [];
        const last = messages[messages.length - 1];
        const scenario = scenarioOf(messages);

        // Новый пользовательский запрос начинает сценарий заново —
        // повторный прогон после отката снова проходит все шаги.
        if (!counters.has(scenario)) counters.set(scenario, { step: 0, lastRole: '' });
        const state = counters.get(scenario);
        if (last?.role === 'user' && state.lastRole !== 'user') {
          state.step = 0;
        }
        const stepIndex = state.step;
        state.step += 1;
        state.lastRole = last?.role || '';

        const systemPrompt = String(messages[0]?.content || '');
        out.requests.push({
          scenario,
          stepIndex,
          lastRole: last?.role || '',
          model: parsed?.model ?? null,
          stream: parsed?.stream ?? null,
          toolNames: Array.isArray(parsed?.tools)
            ? parsed.tools.map((tool) => tool?.function?.name).filter(Boolean)
            : [],
          lastUserContent: String(
            [...messages].reverse().find((message) => message.role === 'user')?.content || ''
          ).slice(0, 500),
          promptHasWrite: systemPrompt.includes('Права записи: ЕСТЬ'),
          promptHasNoWrite: systemPrompt.includes('Права записи: НЕТ'),
          promptMentionsDocument: systemPrompt.includes('ОТКРЫТЫЙ ДОКУМЕНТ'),
          promptFallback: systemPrompt.includes('Function calling недоступен'),
        });

        // Первая попытка fallback-сценария: модель «не поддерживает» tools.
        if (scenario === 'fallback' && Array.isArray(parsed?.tools) && parsed.tools.length > 0) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: { message: 'tools are not supported by this model' } }));
          return;
        }

        const payloads = scriptedResponse(scenario, stepIndex, messages);

        res.writeHead(200, {
          'Content-Type': 'text/event-stream',
          'Cache-Control': 'no-cache',
          Connection: 'keep-alive',
        });
        for (const payload of payloads) {
          res.write(`data: ${JSON.stringify(payload)}\n\n`);
        }
        res.write('data: [DONE]\n\n');
        res.end();
      });
    });

    server.listen(0, '127.0.0.1', () => {
      resolve({ server, baseUrl: `http://127.0.0.1:${server.address().port}/v1` });
    });
  });
}

/* ============================================================
   IPC песочницы
   ============================================================ */

const seededAiSettings = { settings: {}, keys: {} };

function registerHandlers() {
  registerCommonCheckIpc({
    sandbox: SANDBOX,
    getAiSettings: () => seededAiSettings.settings,
    getAiKeys: () => seededAiSettings.keys,
  });

  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('fs:openFolderDialog', () => SANDBOX);
  ipcMain.handle('fs:openFileDialog', () => null);
  ipcMain.handle('fs:saveFileDialog', () => null);
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

  ipcMain.handle('fs:createFile', async (_event, filePath, initialContent = '') => {
    try {
      await fs.promises.writeFile(filePath, initialContent, { flag: 'wx' });
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });
  ipcMain.handle('fs:createDirectory', async (_event, dirPath) => {
    try {
      await fs.promises.mkdir(dirPath, { recursive: true });
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));
  ipcMain.handle('fs:searchWorkspace', async () => ({ success: true, results: [] }));
}

/* ============================================================
   Сценарий
   ============================================================ */

async function main() {
  fs.rmSync(path.join(SANDBOX, DOC_NAME), { force: true });
  fs.writeFileSync(path.join(SANDBOX, DOC_NAME), await buildDocx());

  registerHandlers();
  require('./aiProxy.js').registerAiProxy();

  const mock = await startMockServer();
  seededAiSettings.settings = {
    provider: 'custom',
    baseUrl: mock.baseUrl,
    model: 'test-model',
    temperature: 0.2,
    maxTokens: 4096,
    systemPrompt: 'Ты тестовый ассистент.',
    activePresetId: 'general',
  };
  seededAiSettings.keys = { custom: 'test-key' };

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

  win.webContents.on('console-message', (_event, _level, message) => {
    if (/error|Error|ошибк/i.test(message)) out.errors.push(message);
  });

  const js = (expression) => win.webContents.executeJavaScript(`(async () => { ${expression} })()`);

  const waitIdle = async (timeout = 20000) => {
    const deadline = Date.now() + timeout;
    while (Date.now() < deadline) {
      const idle = await js(`
        const stop = [...document.querySelectorAll('button')]
          .some((button) => (button.textContent || '').trim() === 'Стоп');
        const status = /Готовлю запрос|Читаю документ|Правлю документ|Переписываю документ|Ищу и заменяю|Меняю оформление|Удаляю лишнее|Работаю через текстовые действия/.test(
          document.body.innerText
        );
        return !stop && !status;
      `);
      if (idle) return true;
      await sleep(250);
    }
    return false;
  };

  const send = async (text) => {
    const sent = await js(`
      const ta = [...document.querySelectorAll('textarea')].find((item) =>
        (item.placeholder || '').includes('Спросите')
      );
      if (!ta) return false;
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      setter.call(ta, ${JSON.stringify(text)});
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise((resolve) => setTimeout(resolve, 200));
      const button = [...document.querySelectorAll('button')].find((item) =>
        /Отправить/.test(item.textContent || '')
      );
      if (button) button.click();
      return Boolean(button);
    `);
    await sleep(500);
    const idle = await waitIdle();
    await sleep(300);
    return { sent: Boolean(sent), idle };
  };

  const documentState = () =>
    js(`
      const canvas = document.querySelector('[contenteditable=true]');
      return canvas
        ? { html: canvas.innerHTML, text: canvas.innerText }
        : { html: '', text: '' };
    `);

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await sleep(2500);

    // Открываем резюме из дерева файлов.
    let opened = false;
    for (let attempt = 0; attempt < 20 && !opened; attempt++) {
      opened = await js(`
        const node = [...document.querySelectorAll('div')].find((element) => {
          const title = (element.getAttribute('title') || '').split(' — ')[0];
          return title === ${JSON.stringify(DOC_NAME)} && element.className.includes('cursor-pointer');
        });
        if (node) { node.click(); return true; }
        return false;
      `);
      if (!opened) await sleep(300);
    }
    await sleep(2500);

    step('documentOpened', await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      return {
        opened: Boolean(canvas),
        hasResume: Boolean(canvas) && canvas.innerText.includes('Иван Иванов'),
        hasJunk: Boolean(canvas) && canvas.innerText.includes('не могу редактировать'),
        rightsShown: document.body.innerText.includes('Права записи выданы'),
        panelPresent: [...document.querySelectorAll('textarea')].some((item) =>
          (item.placeholder || '').includes('Спросите')
        ),
      };
    `));

    /* ---------- 1. Удаление мусора ---------- */
    const cleanupSend = await send('Удали из документа всё, что не относится к резюме');
    const cleanupState = await documentState();
    step('cleanup', {
      sent: cleanupSend.sent,
      idle: cleanupSend.idle,
      junkGone:
        !cleanupState.text.includes('не могу редактировать') &&
        !cleanupState.text.includes('Хотите, я создам'),
      resumeKept: cleanupState.text.includes('Иван Иванов') && cleanupState.text.includes('Опыт работы'),
      reportShown: (await js('return document.body.innerText')).includes('Убрал из документа'),
    });

    /* ---------- 2. ATS-friendly ---------- */
    const atsSend = await send('Сделай резюме ATS-friendly');
    const atsState = await documentState();
    step('ats', {
      sent: atsSend.sent,
      idle: atsSend.idle,
      structureChanged: atsState.html.includes('<h3') && atsState.html.includes('<ul'),
      resumeKept: atsState.text.includes('Иван Иванов') && atsState.text.includes('Образование'),
      reportShown: (await js('return document.body.innerText')).includes('под ATS'),
    });

    /* ---------- 3. Оформление: H2 + контакты по центру ---------- */
    const formatSend = await send('Сделай заголовки разделов H2 и выровняй контакты по центру');
    const formatState = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      const headings = canvas ? [...canvas.querySelectorAll('h2')].map((item) => item.textContent) : [];
      const contact = canvas
        ? [...canvas.querySelectorAll('p')].find((item) => item.textContent.includes('@'))
        : null;
      return {
        h2: headings,
        contactAlign: contact ? contact.style.textAlign : '',
        html: canvas ? canvas.innerHTML : '',
      };
    `);
    step('format', {
      sent: formatSend.sent,
      idle: formatSend.idle,
      sectionsAreH2:
        formatState.h2.some((text) => text.includes('Опыт')) &&
        formatState.h2.some((text) => text.includes('Навыки')),
      contactsCentered: formatState.contactAlign === 'center',
      reportShown: (await js('return document.body.innerText')).includes('выровнял контакты'),
    });

    /* ---------- 3b. Откат серии правок ---------- */
    const undoFormat = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        /Вернуть прежний текст/.test(item.textContent || '')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(900);
    step('undoSeries', {
      clicked: Boolean(undoFormat),
      h3Restored: await js(`
        const canvas = document.querySelector('[contenteditable=true]');
        return Boolean(canvas) && canvas.querySelectorAll('h3').length >= 2;
      `),
      centerCleared: await js(`
        const canvas = document.querySelector('[contenteditable=true]');
        const contact = canvas
          ? [...canvas.querySelectorAll('p')].find((item) => item.textContent.includes('@'))
          : null;
        return !contact || contact.style.textAlign !== 'center';
      `),
    });

    // Возвращаем оформление повторным запросом (проверяем и повторный агентный цикл).
    const formatAgain = await send('Сделай заголовки разделов H2 и выровняй контакты по центру');
    const formatAgainState = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      const contact = canvas
        ? [...canvas.querySelectorAll('p')].find((item) => item.textContent.includes('@'))
        : null;
      return {
        h2: canvas ? canvas.querySelectorAll('h2').length : 0,
        contactAlign: contact ? contact.style.textAlign : '',
      };
    `);
    step('formatAgain', {
      sent: formatAgain.sent,
      idle: formatAgain.idle,
      sectionsAreH2: formatAgainState.h2 >= 2,
      contactsCentered: formatAgainState.contactAlign === 'center',
    });

    /* ---------- 4. Точечная замена ---------- */
    const findSend = await send('Замени "Июль 2024" на "Июнь 2024"');
    const findState = await documentState();
    step('findReplace', {
      sent: findSend.sent,
      idle: findSend.idle,
      replaced: findState.text.includes('Июнь 2024') && !findState.text.includes('Июль 2024'),
      reportShown: (await js('return document.body.innerText')).includes('Июнь 2024'),
    });

    /* ---------- 5. Откат одного ответа ---------- */
    const undoFind = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        /Вернуть прежний текст/.test(item.textContent || '')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(900);
    const undoState = await documentState();
    step('undoSingle', {
      clicked: Boolean(undoFind),
      julyRestored: undoState.text.includes('Июль 2024') && !undoState.text.includes('Июнь 2024'),
    });

    /* ---------- 6. Сохранение оформления ---------- */
    const saved = await js(`
      const button = [...document.querySelectorAll('button')].find(
        (item) => (item.textContent || '').trim() === 'Сохранить' && !item.disabled
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(2500);
    const savedDoc = await readDocxXml(path.join(SANDBOX, DOC_NAME));
    step('save', {
      clicked: Boolean(saved),
      fileExists: fs.existsSync(path.join(SANDBOX, DOC_NAME)),
      headingsPreserved: savedDoc.xml.includes('w:pStyle w:val="Heading2"'),
      centerPreserved: savedDoc.xml.includes('<w:jc w:val="center"/>'),
      listPreserved: /<w:numId w:val="1"\/>/.test(savedDoc.xml),
      dateRestored: savedDoc.text.includes('Июль 2024'),
      cleanText:
        !savedDoc.text.includes('не могу редактировать') &&
        !savedDoc.text.includes('Хотите, я создам'),
    });

    /* ---------- 7. Fallback без function calling ---------- */
    const fallbackSend = await send('Оформи документ через запасной протокол');
    const fallbackState = await documentState();
    const fallbackRequests = out.requests.filter((request) => request.scenario === 'fallback');
    step('fallbackProtocol', {
      sent: fallbackSend.sent,
      idle: fallbackSend.idle,
      toolsRejected: fallbackRequests.some((request) => request.toolNames.length === 7),
      retriedWithoutTools: fallbackRequests.some((request) => request.toolNames.length === 0),
      promptFallback: fallbackRequests.some((request) => request.promptFallback),
      documentUpdated: fallbackState.text.includes('Оформлено через document_apply'),
      contactsCentered: fallbackState.html.includes('text-align:center'),
      reportShown: (await js('return document.body.innerText')).includes('Оформляю документ'),
    });

    /* ---------- 8. Без прав записи ---------- */
    const rightsOff = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        /Права записи выданы/.test(item.textContent || '')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(400);
    const beforeNoRights = await documentState();
    const noRightsSend = await send('Перепиши резюме заново');
    const afterNoRights = await documentState();
    const noRightsRequest = [...out.requests].reverse().find((request) => request.scenario === 'norights');

    step('noRights', {
      toggled: Boolean(rightsOff),
      sent: noRightsSend.sent,
      idle: noRightsSend.idle,
      onlyReadTool:
        Boolean(noRightsRequest) &&
        noRightsRequest.toolNames.length === 1 &&
        noRightsRequest.toolNames[0] === 'get_document',
      promptSaysNoWrite: Boolean(noRightsRequest && noRightsRequest.promptHasNoWrite),
      documentUnchanged: beforeNoRights.text === afterNoRights.text,
      honestReply: (await js('return document.body.innerText')).includes('Прав на запись'),
    });

    /* ---------- 8. Контракт запроса агента ---------- */
    const cleanupRequest = out.requests.find((request) => request.scenario === 'cleanup');
    const atsRequest = out.requests.find((request) => request.scenario === 'ats');
    step('agentRequest', {
      allToolsPassed:
        Boolean(cleanupRequest) &&
        cleanupRequest.toolNames.length === 7 &&
        cleanupRequest.toolNames.includes('replace_document') &&
        cleanupRequest.toolNames.includes('format'),
      documentNotAttached:
        Boolean(cleanupRequest) &&
        !String(cleanupRequest.lastUserContent).includes('Иван Иванов') &&
        !String(cleanupRequest.lastUserContent).includes('Москва'),
      promptHasReadWrite: Boolean(cleanupRequest && cleanupRequest.promptHasWrite),
      promptHasDocument: Boolean(cleanupRequest && cleanupRequest.promptMentionsDocument),
      atsUsedFullReplace:
        Boolean(atsRequest) && out.requests.some((request) => request.scenario === 'ats' && request.stepIndex === 1),
    });

    /* ---------- 9. Кнопка «Заменить документ» ---------- */
    const replaceOnlyContent = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        /Заменить документ/.test(item.textContent || '')
      );
      return Boolean(button);
    `);
    step('replaceButton', {
      present: Boolean(replaceOnlyContent),
    });

    (out as any).summary = {
      requests: out.requests.length,
      scenarios: [...new Set(out.requests.map((request) => request.scenario))],
    };
  } catch (err) {
    out.errors.push(err && err.stack ? err.stack : String(err));
  }

  mock.server.close();
  writeOut();
  app.exit(0);
}

app.whenReady().then(() => {
  try {
    fs.writeFileSync(path.join(SANDBOX, 'ready.txt'), 'app ready');
  } catch {}
  // Проверка не должна зависать в CI: через 3 минуты фиксируем таймаут.
  const timer = setTimeout(() => {
    writeOut('timeout: проверка не завершилась за 180 секунд');
    app.exit(2);
  }, 180000);
  void main()
    .catch((err) => {
      writeOut(String(err && err.stack ? err.stack : err));
      app.exit(1);
    })
    .finally(() => clearTimeout(timer));
});
