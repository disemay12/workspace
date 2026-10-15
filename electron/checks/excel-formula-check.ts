/**
 * Проверка поведения строки формул в Excel-редакторе.
 *
 * Воспроизводит жалобу «формула остаётся в верхнем поле и переносится на
 * другую ячейку»: вводим формулу в ячейку, затем выбираем другую ячейку и
 * проверяем, что строка формул показывает содержимое НОВОЙ ячейки, а не
 * оставшуюся формулу.
 *
 * Запуск: electron.exe dist-electron/electron/checks/excel-formula-check.js
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const iconv = require('iconv-lite');

function resolveProjectRoot() {
  let dir = __dirname;
  for (let i = 0; i < 6; i++) {
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
const OUT = path.join(ROOT, 'excel-formula-check.json');
const results: any = { root: ROOT };
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

app.disableHardwareAcceleration();

function registerIPC() {
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
      return {
        success: true,
        items: entries.map((entry) => {
          const full = path.join(dirPath, entry.name);
          let size = 0;
          let mtime = null;
          try {
            const st = fs.statSync(full);
            size = st.size;
            mtime = st.mtime.toISOString();
          } catch {
            /* файл мог исчезнуть */
          }
          return {
            name: entry.name,
            path: full,
            isDirectory: entry.isDirectory(),
            size,
            mtime,
            extension: entry.name.includes('.') ? entry.name.split('.').pop() : '',
          };
        }),
      };
    } catch (err) {
      return { success: false, error: (err as Error).message, items: [] };
    }
  });
  ipcMain.handle('fs:readFile', async (_e, filePath, encoding = 'utf-8') => {
    try {
      const buffer = await fs.promises.readFile(filePath);
      if (encoding === 'base64') {
        return { success: true, content: buffer.toString('base64'), isBinary: true };
      }
      if (encoding && encoding.toLowerCase() !== 'utf-8' && iconv.encodingExists(encoding)) {
        return { success: true, content: iconv.decode(buffer, encoding), isBinary: false };
      }
      return { success: true, content: buffer.toString('utf-8'), isBinary: false };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });
  ipcMain.handle('fs:writeFile', async (_e, filePath, content, encoding = 'utf-8') => {
    try {
      const tmpPath = `${filePath}.ws-tmp`;
      if (encoding === 'base64') await fs.promises.writeFile(tmpPath, Buffer.from(content, 'base64'));
      else await fs.promises.writeFile(tmpPath, content, 'utf-8');
      try {
        await fs.promises.rename(tmpPath, filePath);
      } catch (err) {
        const code = (err as NodeJS.ErrnoException).code;
        if (code !== 'EPERM' && code !== 'EEXIST' && code !== 'EBUSY') throw err;
        await fs.promises.copyFile(tmpPath, filePath);
        await fs.promises.unlink(tmpPath).catch(() => undefined);
      }
      return { success: true };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });
  ipcMain.handle('fs:createFile', async () => ({ success: true }));
  ipcMain.handle('fs:createDirectory', async () => ({ success: true }));
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));
  ipcMain.handle('fs:getFavorites', () => []);
  ipcMain.handle('fs:toggleFavorite', () => []);
}

async function main() {
  if (!fs.existsSync(SANDBOX)) fs.mkdirSync(SANDBOX, { recursive: true });
  const xlsxPath = path.join(SANDBOX, 'Формулы.xlsx');

  // Готовим книгу через тот же движок, что использует приложение.
  const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Число', 'Формула'],
    [10, ''],
    [20, ''],
    [30, ''],
  ]);
  XLSX.utils.book_append_sheet(wb, ws, 'Лист1');
  XLSX.writeFile(wb, xlsxPath);

  registerIPC();

  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    show: false,
    webPreferences: {
      preload: path.join(ROOT, 'dist-electron', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
  await sleep(2500);

  // Открываем книгу из дерева файлов
  results.opened = await win.webContents.executeJavaScript(`
    (() => {
      const nodes = [...document.querySelectorAll('div')]
        .filter((d) => d.textContent?.trim() === 'Формулы.xlsx');
      const clickable = nodes.find((d) => d.className.includes('cursor-pointer')) || nodes[0];
      if (!clickable) return false;
      clickable.click();
      return true;
    })()
  `);
  await sleep(3000);

  /** Хелпер: читает содержимое строки формул и адрес активной ячейки. */
  const readBar = () =>
    win.webContents.executeJavaScript(`
      (() => {
        const bar = [...document.querySelectorAll('input[type=text]')]
          .find((i) => (i.placeholder || '').includes('формула') || (i.placeholder || '').includes('SUM'));
        // Поле адреса — первый div с текстом вида A1 (оно идёт перед строкой формул).
        const refBox = [...document.querySelectorAll('div')]
          .map((d) => (d.textContent || '').trim())
          .find((t) => /^[A-Z]{1,3}\\d{1,5}$/.test(t));
        return { barValue: bar ? bar.value : null, hasBar: !!bar, refBox: refBox || null };
      })()
    `);

  /**
   * Клик по ячейке по её адресу A1.
   *
   * Структура таблицы: tr[0] — заголовки столбцов (thead), дальше по одному
   * tr на строку листа, а первая td каждой строки — номер строки. Поэтому
   * адрес A1 соответствует trs[1] и td[1].
   */
  const clickCell = (ref: string) =>
    win.webContents.executeJavaScript(`
      (() => {
        const m = ${JSON.stringify(ref)}.match(/^([A-Z]+)(\\d+)$/);
        if (!m) return { ok: false, reason: 'bad ref' };
        let col = 0;
        for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
        const row = parseInt(m[2], 10);
        const trs = [...document.querySelectorAll('tr')];
        const tr = trs[row];
        if (!tr) return { ok: false, reason: 'no tr', rowCount: trs.length };
        const tds = tr.querySelectorAll('td');
        // tds[0] — номер строки, дальше столбцы A, B, C...
        const td = tds[col];
        if (!td) return { ok: false, reason: 'no td', tdCount: tds.length };
        td.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
        td.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        return { ok: true, text: (td.textContent || '').trim() };
      })()
    `);

  /** Установка значения строки формул с корректным событием React. */
  const setBar = (value: string) =>
    win.webContents.executeJavaScript(`
      (() => {
        const bar = [...document.querySelectorAll('input[type=text]')]
          .find((i) => (i.placeholder || '').includes('формула') || (i.placeholder || '').includes('SUM'));
        if (!bar) return false;
        bar.focus();
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(bar, ${JSON.stringify(value)});
        bar.dispatchEvent(new Event('input', { bubbles: true }));
        return true;
      })()
    `);

  /** Нажатие клавиши в строке формул. */
  const pressInBar = (key: string) =>
    win.webContents.executeJavaScript(`
      (() => {
        const bar = [...document.querySelectorAll('input[type=text]')]
          .find((i) => (i.placeholder || '').includes('формула') || (i.placeholder || '').includes('SUM'));
        if (!bar) return false;
        bar.dispatchEvent(new KeyboardEvent('keydown', { key: ${JSON.stringify(key)}, bubbles: true }));
        return true;
      })()
    `);

  results.barInitial = await readBar();

  // 1. Выбираем B2 и вводим формулу =A2+1
  results.clickB2 = await clickCell('B2');
  await sleep(700);
  results.barAfterSelectB2 = await readBar();

  results.typeFormula = await setBar('=A2+1');
  await sleep(600);
  results.barWhileTyping = await readBar();

  results.pressEnter = await pressInBar('Enter');
  await sleep(900);
  results.barAfterEnter = await readBar();

  // 2. Выбираем ДРУГУЮ ячейку (A4 = 30). Строка формул должна показать 30.
  results.clickA4 = await clickCell('A4');
  await sleep(900);
  results.barAfterSelectA4 = await readBar();

  // 3. Проверяем, что формула действительно записалась в B2
  results.clickB2Again = await clickCell('B2');
  await sleep(900);
  results.barBackInB2 = await readBar();

  // 4. Проверяем, что выбор диапазона мышью в формуле всё ещё работает:
  //    начинаем формулу =СУММ( и кликаем по A2.
  results.sumStart = await setBar('=SUM(');
  await sleep(600);
  results.clickA2ForRef = await clickCell('A2');
  await sleep(700);
  results.barAfterRangeClick = await readBar();

  results.summary = {
    formulaLeakedToOtherCell:
      typeof results.barAfterSelectA4.barValue === 'string' &&
      results.barAfterSelectA4.barValue.startsWith('='),
    barAfterEnter: results.barAfterEnter.barValue,
    barAfterSelectA4: results.barAfterSelectA4.barValue,
    expectedA4: '30',
    formulaStoredInB2: results.barBackInB2.barValue,
    rangeClickInsertsRef:
      typeof results.barAfterRangeClick.barValue === 'string' &&
      results.barAfterRangeClick.barValue.includes('A2'),
  };

  fs.writeFileSync(OUT, JSON.stringify(results, null, 2), 'utf-8');
  console.log(JSON.stringify(results, null, 2));
  win.destroy();
  app.quit();
}

app.whenReady().then(main).catch((err) => {
  fs.writeFileSync(OUT, JSON.stringify({ fatal: String(err), results }, null, 2), 'utf-8');
  console.error(err);
  app.quit();
});
