/**
 * Smoke-тест: запускает Electron, открывает окно WorkSpace, дожидается
 * отрисовки React и проверяет, что ключевые элементы интерфейса присутствуют,
 * а в консоли нет ошибок. Затем делает скриншот и завершает процесс.
 *
 * Запуск: electron.exe dist-electron/smoke.js
 */
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');

/** Корень проекта — поиск каталога с package.json (устойчиво к глубине). */
function resolveProjectRoot() {
  let dir = __dirname;
  for (let i = 0; i < 4; i++) {
    if (fs.existsSync(path.join(dir, 'package.json'))) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return path.join(__dirname, '..');
}
const ROOT = resolveProjectRoot();
const errors: string[] = [];
const logs: string[] = [];

app.disableHardwareAcceleration();

app.whenReady().then(async () => {
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
    },
  });

  win.webContents.on(
    'console-message',
    (_e: unknown, level: number, message: string, line: number, sourceId: string) => {
      const entry = `[${level}] ${message} (${sourceId}:${line})`;
      logs.push(entry);
      // level 3 = error
      if (level >= 3) errors.push(entry);
    }
  );

  win.webContents.on(
    'did-fail-load',
    (_e: unknown, code: number, desc: string, url: string) => {
      errors.push(`did-fail-load ${code} ${desc} ${url}`);
    }
  );

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    // Даём React смонтироваться и отработать эффекту initWorkspace
    await new Promise((r) => setTimeout(r, 3500));

    const probe = await win.webContents.executeJavaScript(`
      (() => {
        const root = document.getElementById('root');
        const text = document.body.innerText || '';
        return {
          rootChildren: root ? root.children.length : -1,
          htmlLength: document.documentElement.outerHTML.length,
          hasTitleBar: !!document.querySelector('.titlebar-drag'),
          hasWorkSpaceLabel: text.includes('WorkSpace'),
          hasExplorer: text.includes('Проводник'),
          hasAI: text.includes('AI Ассистент') || text.includes('Чем помочь'),
          hasEmptyState: text.includes('Открыть папку'),
          hasHotkeys: text.includes('Ctrl+K'),
          hasElectronAPI: typeof window.electronAPI === 'object' && window.electronAPI !== null,
          apiKeys: window.electronAPI ? Object.keys(window.electronAPI).length : 0,
          bodyText: text.slice(0, 600),
        };
      })()
    `);

    // Скриншот для визуальной проверки
    win.show();
    await new Promise((r) => setTimeout(r, 900));
    const image = await win.webContents.capturePage();
    fs.writeFileSync(path.join(ROOT, 'smoke-screenshot.png'), image.toPNG());

    console.log(
      'SMOKE_RESULT ' + JSON.stringify({ ...probe, errors, logCount: logs.length }, null, 2)
    );
    fs.writeFileSync(
      path.join(ROOT, 'smoke-result.json'),
      JSON.stringify({ ...probe, errors, logs }, null, 2),
      'utf-8'
    );
  } catch (err) {
    console.log('SMOKE_FAILED ' + (err as Error).message);
    console.log('SMOKE_ERRORS ' + JSON.stringify(errors));
    fs.writeFileSync(
      path.join(ROOT, 'smoke-result.json'),
      JSON.stringify({ failed: (err as Error).message, errors, logs }, null, 2),
      'utf-8'
    );
  }

  app.exit(errors.length > 0 ? 1 : 0);
});
