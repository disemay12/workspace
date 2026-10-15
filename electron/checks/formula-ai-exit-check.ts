/**
 * Проверка трёх исправлений, о которых сообщил пользователь:
 *   1. формулы Excel считаются полностью (не только SUM/AVERAGE);
 *   2. AI-формула вставляется в ячейку с подтверждением, а не уходит в чат;
 *   3. при закрытии приложения с несохранёнными файлами появляется диалог.
 *
 * Загружает реальную сборку приложения в окно Electron и взаимодействует
 * с интерфейсом так же, как пользователь. Результат — formula-ai-exit-check.json.
 */
const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

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
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

app.disableHardwareAcceleration();

/** Запрос подтверждения закрытия, отправленный в рендерер. */
let closeRequestCount = 0;
let lastConfirm = null;
/** Окно, которое реально закрылось. */
let windowClosed = false;

app.whenReady().then(async () => {
  /* ---- IPC, повторяющий main.ts ---- */
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
  ipcMain.handle('fs:createFile', async () => ({ success: true }));
  ipcMain.handle('fs:createDirectory', async () => ({ success: true }));
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));
  ipcMain.handle('fs:searchWorkspace', async () => ({ success: true, results: [] }));

  /** Считаем записи на диск — нужно для проверки «Сохранить всё и выйти». */
  const written = [];
  ipcMain.handle('fs:writeFile', async (_e, filePath, content, encoding) => {
    written.push({ filePath, length: String(content ?? '').length, encoding });
    try {
      const buf =
        encoding === 'base64'
          ? Buffer.from(content, 'base64')
          : Buffer.from(String(content), 'utf-8');
      fs.mkdirSync(path.dirname(filePath), { recursive: true });
      fs.writeFileSync(filePath, buf);
    } catch (err) {
      return { success: false, error: err.message };
    }
    return { success: true };
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

  /** Подтверждение закрытия — фиксируем, что рендерер его вызвал. */
  ipcMain.handle('window:confirmClose', (_event, proceed) => {
    lastConfirm = proceed;
    if (proceed) {
      windowClosed = true;
      // Реальный main.ts закрывает окно; здесь просто отмечаем факт
    }
    return { closed: Boolean(proceed) };
  });

  /* ---- Подготовка тестовой таблицы ---- */
  const xlsxPath = path.join(SANDBOX, 'Формулы.xlsx');
  fs.mkdirSync(SANDBOX, { recursive: true });

  const win = new BrowserWindow({
    width: 1600,
    height: 1000,
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
    },
  });

  // Перехватываем запрос закрытия так же, как это делает main.ts
  win.on('close', (event) => {
    if (windowClosed) return;
    event.preventDefault();
    closeRequestCount++;
    win.webContents.send('window:closeRequested');
  });

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await sleep(2500);

    /* ============================================================
       1. Подготовка тестовой книги и открытие её через проводник
       ============================================================ */
    // Создаём книгу с формулами разных категорий прямо в файловой системе
    const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));
    const wb = XLSX.utils.book_new();
    const data = [
      ['Товар', 'Кол-во', 'Цена', 'Оплачено'],
      ['Яблоки', 10, 50, 'Да'],
      ['Груши', 5, 80, 'Нет'],
      ['Сливы', 20, 30, 'Да'],
      ['Сумма', { f: 'SUM(B2:B4)' }, { f: 'SUM(B2:B4*C2:C4)' }, { f: 'COUNTIF(D2:D4,"Да")' }],
      ['Среднее', { f: 'ROUND(AVERAGE(B2:B4),2)' }, { f: 'MAX(C2:C4)' }, { f: 'MIN(C2:C4)' }],
      ['Условие', { f: 'IF(B2>5,"много","мало")' }, { f: 'SUMIF(D2:D4,"Да",B2:B4)' }, { f: 'VLOOKUP("Груши",A2:C4,3,FALSE)' }],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), 'Данные');
    XLSX.writeFile(wb, xlsxPath);

    // Открываем файл способом, доступным из UI: через Quick Open не выйдет
    // (файл создан после старта), поэтому перезагружаем дерево и кликаем.
    out.openFile = await win.webContents.executeJavaScript(`
      (async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        // Обновляем дерево, чтобы появился новый файл
        const refreshBtn = [...document.querySelectorAll('button')]
          .find(b => /обновить/i.test(b.title || ''));
        if (refreshBtn) refreshBtn.click();
        await sleep(1200);

        const el = [...document.querySelectorAll('*')]
          .filter(e => e.children.length === 0 && /Формулы\\.xlsx/.test(e.textContent || ''));
        if (!el.length) return { found: false, treeText: document.body.innerText.slice(0, 400) };
        el[0].click();
        await sleep(1800);
        return { found: true };
      })()
    `).catch((e) => ({ error: String(e) }));

    /* Проверяем вычисленные значения в сетке */
    out.formulaDisplay = await win.webContents.executeJavaScript(`
      (async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        await sleep(500);
        const cells = [...document.querySelectorAll('td')].map(td => (td.innerText || '').trim());
        const text = document.body.innerText;
        return {
          // Проверяем, что вычисленные значения появились, а не сырые формулы
          hasSum35: cells.includes('35'),
          hasSum1500: cells.includes('1500'),
          hasCountIf2: cells.includes('2'),
          hasAvg1167: cells.includes('11.67'),
          hasIfMnogo: cells.includes('много'),
          hasVlookup80: cells.includes('80'),
          rawFormulaVisible: text.includes('=SUM(') || text.includes('=COUNTIF('),
          stats: /Формул:\\s*\\d+/.test(text),
          sample: cells.slice(0, 24),
        };
      })()
    `).catch((e) => ({ error: String(e) }));

    /* ============================================================
       3. Формульный режим: автодополнение и выбор мышью
       ============================================================ */
    out.formulaMode = await win.webContents.executeJavaScript(`
      (async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        const input = document.querySelector('input[placeholder*="например =SUM"]');
        if (!input) return { error: 'строка формул не найдена' };

        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        // Вводим начало формулы — ожидаем подсказки функций
        setter.call(input, '=СУМ');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        await sleep(400);
        const afterTyping = document.body.innerText;
        const hasSuggestions = /СУММ/.test(afterTyping);

        // Проверяем вставку ссылки кликом по ячейке в формульном режиме
        setter.call(input, '=SUM(');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        await sleep(300);

        const tds = [...document.querySelectorAll('tbody td')];
        // Кликаем по ячейке B2 (первая строка данных, второй столбец)
        const target = tds.find(td => td.title && /^10$/.test(td.title)) || tds[2];
        if (target) {
          target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
          target.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        }
        await sleep(400);

        const finalValue = input.value;
        return {
          hasSuggestions,
          formulaAfterClick: finalValue,
          clickInsertedRef: /SUM\\([A-Z]+\\d+/.test(finalValue),
        };
      })()
    `).catch((e) => ({ error: String(e) }));

    /* ============================================================
       4. AI-формула: предложение применить, а не уход в чат
       ============================================================ */
    out.aiFormula = await win.webContents.executeJavaScript(`
      (async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        const aiBtn = [...document.querySelectorAll('button')]
          .find(b => /AI-формула/i.test(b.textContent || ''));
        if (!aiBtn) return { error: 'кнопка AI-формула не найдена' };
        aiBtn.click();
        await sleep(500);

        const opened = /Генератор формул/.test(document.body.innerText);
        const ta = document.querySelector('textarea[placeholder*="посчитай сумму"]');
        if (!ta) return { opened, error: 'поле промта не найдено' };

        const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
        setter.call(ta, 'сумма по столбцу B');
        ta.dispatchEvent(new Event('input', { bubbles: true }));
        await sleep(300);

        const gen = [...document.querySelectorAll('button')]
          .find(b => /Сгенерировать/i.test(b.textContent || ''));
        if (gen) gen.click();
        await sleep(800);

        return {
          opened,
          dialogClosed: !/Генератор формул/.test(document.body.innerText),
          // Запрос должен попасть в панель AI как контекст, а не как «диалог без действия»
          chatHasRequest: /ЯЧЕЙКА/.test(document.body.innerText),
        };
      })()
    `).catch((e) => ({ error: String(e) }));

    /* Симулируем ответ модели и проверяем диалог подтверждения */
    out.aiConfirm = await win.webContents.executeJavaScript(`
      (async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        window.dispatchEvent(new CustomEvent('workspace:ai-answer', {
          detail: { answer: 'ЯЧЕЙКА: E2\\nФОРМУЛА: =SUM(B2:B4)' },
        }));
        await sleep(900);
        const text = document.body.innerText;
        return {
          confirmShown: /Применить действие AI/.test(text),
          showsTarget: /E2/.test(text),
          showsFormula: /SUM\\(B2:B4\\)/.test(text),
          showsPreview: /Предпросмотр на текущих данных/.test(text),
          hasApply: /Применить/.test(text),
        };
      })()
    `).catch((e) => ({ error: String(e) }));

    /* Применяем и проверяем, что значение записалось в ячейку */
    out.aiApplied = await win.webContents.executeJavaScript(`
      (async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        const apply = [...document.querySelectorAll('button')]
          .find(b => (b.textContent || '').trim() === 'Применить');
        if (!apply) return { error: 'кнопка Применить не найдена' };
        apply.click();
        await sleep(1200);

        const text = document.body.innerText;
        const cells = [...document.querySelectorAll('td')].map(td => (td.innerText || '').trim());
        return {
          dialogClosed: !/Применить действие AI/.test(text),
          // E2 — сумма B2:B4 = 35
          valueInCell: cells.includes('35'),
          formulasCounted: /Формул:\\s*\\d+/.test(text),
        };
      })()
    `).catch((e) => ({ error: String(e) }));

    /* ============================================================
       5. Подтверждение выхода
       ============================================================ */
    // Делаем вкладку «грязной», чтобы появился диалог
    out.makeDirty = await win.webContents.executeJavaScript(`
      (async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        // Меняем значение ячейки через сетку: двойной клик и ввод
        const tds = [...document.querySelectorAll('tbody td')];
        const cell = tds.find(td => td.title === 'Яблоки');
        if (!cell) return { error: 'ячейка не найдена' };
        cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
        await sleep(400);
        const editor = document.querySelector('td input');
        if (!editor) return { error: 'редактор ячейки не открылся' };
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(editor, 'Яблоки красные');
        editor.dispatchEvent(new Event('input', { bubbles: true }));
        editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        await sleep(900);
        return { dirty: /не сохранено|Изменено/i.test(document.body.innerText) };
      })()
    `).catch((e) => ({ error: String(e) }));

    // Просим окно закрыться — перехват должен отправить запрос в рендерер
    win.close();
    await sleep(1500);

    out.closeIntercept = {
      closeRequested: closeRequestCount > 0,
      windowStillOpen: !win.isDestroyed(),
    };

    out.exitDialog = await win.webContents.executeJavaScript(`
      (async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        await sleep(400);
        const text = document.body.innerText;
        return {
          dialogShown: /Несохранённые изменения/.test(text),
          hasSaveAll: /Сохранить всё и выйти/.test(text),
          hasExitNoSave: /Выйти без сохранения/.test(text),
          hasCancel: /Отмена/.test(text),
          listsFile: /Формулы\\.xlsx/.test(text) || /Яблоки красные/.test(text),
        };
      })()
    `).catch((e) => ({ error: String(e) }));

    // Проверяем «Сохранить всё и выйти» — файл должен записаться
    out.saveAllAndExit = await win.webContents.executeJavaScript(`
      (async () => {
        const sleep = ms => new Promise(r => setTimeout(r, ms));
        const btn = [...document.querySelectorAll('button')]
          .find(b => /Сохранить всё и выйти/i.test(b.textContent || ''));
        if (!btn) return { error: 'кнопка не найдена' };
        btn.click();
        await sleep(2000);
        return { clicked: true };
      })()
    `).catch((e) => ({ error: String(e) }));

    out.writtenFiles = written;
    out.confirmProceed = lastConfirm;

    // Закрываем окно окончательно
    windowClosed = true;
    if (!win.isDestroyed()) win.destroy();

    out.errors = [];
  } catch (err) {
    out.fatal = err.stack || err.message;
  }

  fs.writeFileSync(
    path.join(ROOT, 'formula-ai-exit-check.json'),
    JSON.stringify(out, null, 2),
    'utf-8'
  );
  app.exit(0);
});
