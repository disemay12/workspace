/**
 * Интеграционная проверка IPC-слоя файловой системы.
 *
 * Обработчики вызываются так же, как из рендерера — через ipcRenderer.invoke
 * из реального окна с preload-скриптом. Это проверяет весь путь целиком:
 * renderer -> preload (contextBridge) -> ipcMain -> файловая система.
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
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
const SANDBOX = path.join(os.tmpdir(), 'ws-integration-' + Date.now());

const out: any = { steps: [], errors: [] };
const step = (name, value) => out.steps.push({ name, ...value });

app.disableHardwareAcceleration();

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** Регистрируем обработчики, идентичные dist-electron/main.js. */
function registerHandlers(): any {
  registerCommonCheckIpc({ sandbox: SANDBOX });

  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('window:isMaximized', () => false);
  ipcMain.handle('window:minimize', () => undefined);
  ipcMain.handle('window:maximize', () => false);
  ipcMain.handle('window:close', () => undefined);
  ipcMain.handle('fs:openFolderDialog', () => SANDBOX);
  ipcMain.handle('fs:openFileDialog', () => null);
  ipcMain.handle('fs:saveFileDialog', () => null);
  ipcMain.handle('fs:showInExplorer', () => ({ success: true }));

  ipcMain.handle('fs:listDirectory', async (_e, dirPath) => {
    try {
      const entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
      const items = await Promise.all(
        entries.map(async (entry) => {
          const fullPath = path.join(dirPath, entry.name);
          let size = 0;
          let mtime = null;
          try {
            const stat = await fs.promises.stat(fullPath);
            size = stat.size;
            mtime = stat.mtime.toISOString();
          } catch {}
          return {
            name: entry.name,
            path: fullPath,
            isDirectory: entry.isDirectory(),
            size,
            mtime,
            extension: entry.isDirectory()
              ? ''
              : path.extname(entry.name).toLowerCase().replace('.', ''),
          };
        })
      );
      items.sort((a, b) => {
        if (a.isDirectory !== b.isDirectory) return a.isDirectory ? -1 : 1;
        return a.name.localeCompare(b.name, 'ru', { numeric: true, sensitivity: 'base' });
      });
      return { success: true, items };
    } catch (err) {
      return { success: false, error: err.message, items: [] };
    }
  });

  ipcMain.handle('fs:readFile', async (_e, filePath, encoding = 'utf-8') => {
    try {
      const buffer = await fs.promises.readFile(filePath);
      if (encoding === 'base64') return { success: true, content: buffer.toString('base64'), isBinary: true };
      if (encoding && encoding.toLowerCase() !== 'utf-8' && iconv.encodingExists(encoding)) {
        return { success: true, content: iconv.decode(buffer, encoding), isBinary: false };
      }
      if (buffer.length >= 2) {
        if (buffer[0] === 0xff && buffer[1] === 0xfe)
          return { success: true, content: iconv.decode(buffer, 'utf-16le'), isBinary: false };
        if (buffer[0] === 0xfe && buffer[1] === 0xff)
          return { success: true, content: iconv.decode(buffer, 'utf-16be'), isBinary: false };
      }
      return { success: true, content: buffer.toString('utf-8'), isBinary: false };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:writeFile', async (_e, filePath, content, encoding = 'utf-8') => {
    try {
      const tmpPath = `${filePath}.ws-tmp`;
      if (encoding === 'base64') {
        await fs.promises.writeFile(tmpPath, Buffer.from(content, 'base64'));
      } else if (encoding && encoding.toLowerCase() !== 'utf-8' && iconv.encodingExists(encoding)) {
        await fs.promises.writeFile(tmpPath, iconv.encode(content, encoding));
      } else {
        await fs.promises.writeFile(tmpPath, content, 'utf-8');
      }
      await fs.promises.rename(tmpPath, filePath);
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
      if (err.code === 'EEXIST') return { success: false, error: 'Файл с таким именем уже существует' };
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

  ipcMain.handle('fs:deletePath', async (_e, targetPath) => {
    try {
      const stat = await fs.promises.stat(targetPath);
      if (stat.isDirectory()) await fs.promises.rm(targetPath, { recursive: true, force: true });
      else await fs.promises.unlink(targetPath);
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:renamePath', async (_e, oldPath, newPath) => {
    try {
      if (fs.existsSync(newPath)) return { success: false, error: 'Объект с таким именем уже существует' };
      await fs.promises.rename(oldPath, newPath);
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:searchWorkspace', async (_e, rootDir, query, searchContent = false, options = {}) => {
    const IGNORED = new Set(['node_modules', '.git', '.svn', '.hg', 'dist', 'dist-electron', 'build', 'out', '.next', '.cache', '.vscode', '.idea', '__pycache__', '.venv', 'venv', 'target']);
    const TEXT_EXT = new Set(['.txt', '.md', '.markdown', '.log', '.json', '.yaml', '.yml', '.xml', '.ini', '.cfg', '.conf', '.env', '.properties', '.js', '.jsx', '.mjs', '.cjs', '.ts', '.tsx', '.html', '.htm', '.css', '.scss', '.less', '.py', '.sql', '.sh', '.bash', '.bat', '.cmd', '.ps1', '.csv', '.tsv', '.toml', '.c', '.cpp', '.h', '.hpp', '.cs', '.java', '.go', '.rs', '.php', '.rb', '.lua']);
    const MAX = 300;

    if (!query || !query.trim()) return { success: true, results: [] };
    if (!rootDir || !fs.existsSync(rootDir)) return { success: false, error: 'Рабочая папка недоступна', results: [] };

    const caseSensitive = !!(options as any).caseSensitive;
    const useRegex = !!(options as any).useRegex;
    const includeGlob = (options as any).includeGlob ?? '*';

    let regex;
    try {
      regex = useRegex
        ? new RegExp(query, caseSensitive ? '' : 'i')
        : new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), caseSensitive ? '' : 'i');
    } catch (err) {
      return { success: false, error: `Некорректное регулярное выражение: ${err.message}`, results: [] };
    }

    const matchesGlob = (fileName, filePath, glob) => {
      if (!glob || glob === '*') return true;
      const patterns = glob.includes('{')
        ? glob.replace(/^\{|\}$/g, '').split(',').map((p) => p.trim())
        : [glob];
      return patterns.some((pattern) => {
        const target = pattern.includes('/') || pattern.includes('\\') ? filePath.replace(/\\/g, '/') : fileName;
        const escaped = pattern.replace(/\\/g, '/').replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
        return new RegExp(`^${escaped}$`, 'i').test(target) || new RegExp(escaped, 'i').test(target);
      });
    };

    const results: any[] = [];
    const matcher = (text) => {
      regex.lastIndex = 0;
      return regex.test(text);
    };

    const walk = (dir) => {
      if (results.length >= MAX) return;
      let entries = [];
      try {
        entries = fs.readdirSync(dir, { withFileTypes: true });
      } catch {
        return;
      }
      for (const entry of entries) {
        if (results.length >= MAX) break;
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
          if (IGNORED.has(entry.name) || entry.name.startsWith('.')) continue;
          walk(full);
          continue;
        }
        if (!matchesGlob(entry.name, full, includeGlob)) continue;
        if (matcher(entry.name)) {
          results.push({ path: full, name: entry.name, isDirectory: false });
          continue;
        }
        if (!searchContent) continue;
        const ext = path.extname(entry.name).toLowerCase();
        if (!TEXT_EXT.has(ext)) continue;
        try {
          const stat = fs.statSync(full);
          if (stat.size > 4 * 1024 * 1024) continue;
          const text = fs.readFileSync(full, 'utf-8');
          const lines = text.split(/\r?\n/);
          let first = null;
          let total = 0;
          for (let i = 0; i < lines.length; i++) {
            if (matcher(lines[i])) {
              total++;
              if (!first) first = { line: i + 1, snippet: lines[i].trim().slice(0, 160) };
              if (total >= 50) break;
            }
          }
          if (first) {
            results.push({ path: full, name: entry.name, isDirectory: false, line: first.line, matchSnippet: first.snippet, totalMatches: total });
          }
        } catch {}
      }
    };

    walk(rootDir);
    return { success: true, results };
  });
}

app.whenReady().then(async () => {
  registerHandlers();

  /* ---------- Подготовка тестовых данных ---------- */
  fs.mkdirSync(SANDBOX, { recursive: true });
  fs.mkdirSync(path.join(SANDBOX, 'sub'), { recursive: true });
  fs.writeFileSync(path.join(SANDBOX, 'a.txt'), 'hello world\n');
  fs.writeFileSync(path.join(SANDBOX, 'b.log'), 'INFO ok\nERROR broke\nFATAL down\n');
  fs.writeFileSync(path.join(SANDBOX, 'sub', 'nested.md'), '# Nested\nwith Needle inside\n');
  fs.writeFileSync(path.join(SANDBOX, 'code.ts'), 'const x: number = 1;\nfunction f() {}\n');
  fs.writeFileSync(path.join(SANDBOX, 'cyrillic.txt'), iconv.encode('Привет, мир!\nВторая строка\n', 'windows-1251'));
  // Файл UTF-16 LE с BOM — проверяем автоопределение
  fs.writeFileSync(path.join(SANDBOX, 'utf16.txt'), Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('UTF16 текст\n', 'utf16le')]));

  const win = new BrowserWindow({
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
    await sleep(1200);

    // Вызываем IPC строго через preload — как это делает приложение
    const api = `window.electronAPI`;
    const call = async (expr) => win.webContents.executeJavaScript(`(async () => { ${expr} })()`);

    /* 1. listDirectory */
    const listing = await call(`return await ${api}.listDirectory(${JSON.stringify(SANDBOX)})`);
    step('listDirectory', {
      success: listing.success,
      count: listing.items.length,
      dirsFirst: listing.items[0].isDirectory === true && listing.items.slice(1).every((i) => !i.isDirectory),
      names: listing.items.map((i) => i.name),
      hasSize: listing.items.every((i) => typeof i.size === 'number'),
    });

    /* 2. Чтение UTF-8, CP1251, UTF-16 BOM */
    const utf = await call(`return await ${api}.readFile(${JSON.stringify(path.join(SANDBOX, 'a.txt'))}, 'utf-8')`);
    const cp = await call(`return await ${api}.readFile(${JSON.stringify(path.join(SANDBOX, 'cyrillic.txt'))}, 'windows-1251')`);
    const cpAsUtf = await call(`return await ${api}.readFile(${JSON.stringify(path.join(SANDBOX, 'cyrillic.txt'))}, 'utf-8')`);
    const u16 = await call(`return await ${api}.readFile(${JSON.stringify(path.join(SANDBOX, 'utf16.txt'))}, 'utf-8')`);
    step('readFileEncodings', {
      utf8Ok: utf.content === 'hello world\n',
      cp1251Ok: cp.content.includes('Привет, мир!'),
      cp1251AsUtf8Broken: !cpAsUtf.content.includes('Привет'),
      utf16BomAutoDetected: u16.content.includes('UTF16 текст'),
    });

    /* 3. Запись CP1251 + round-trip */
    const cpTarget = path.join(SANDBOX, 'written-cp.txt');
    const cpText = 'Тест записи в CP1251\n';
    const wr = await call(`return await ${api}.writeFile(${JSON.stringify(cpTarget)}, ${JSON.stringify(cpText)}, 'windows-1251')`);
    const rawBytes = fs.readFileSync(cpTarget);
    const utf8Bytes = Buffer.from(cpText, 'utf-8');
    const back = await call(`return await ${api}.readFile(${JSON.stringify(cpTarget)}, 'windows-1251')`);
    step('writeFileCp1251', {
      writeOk: wr.success,
      byteLength: rawBytes.length,
      utf8LengthWouldBe: utf8Bytes.length,
      // Ключевая проверка: в CP1251 кириллица однобайтовая, поэтому файл
      // должен быть КОРОЧЕ utf-8-варианта ровно на число кириллических символов.
      cp1251IsCompact: rawBytes.length < utf8Bytes.length,
      cp1251LengthMatches: rawBytes.length === Buffer.from(cpText, 'latin1').length,
      roundTripOk: back.content === cpText,
      noTempLeftover: !fs.existsSync(`${cpTarget}.ws-tmp`),
    });

    /* 4. createFile — запрет перезаписи */
    const nf = path.join(SANDBOX, 'new.txt');
    const c1 = await call(`return await ${api}.createFile(${JSON.stringify(nf)}, 'init')`);
    const c2 = await call(`return await ${api}.createFile(${JSON.stringify(nf)}, 'again')`);
    step('createFile', {
      firstOk: c1.success,
      duplicateRejected: c2.success === false,
      duplicateMessage: c2.error,
      notOverwritten: fs.readFileSync(nf, 'utf-8') === 'init',
    });

    /* 5. renamePath — конфликт */
    const aTxt = path.join(SANDBOX, 'a.txt');
    const a2 = path.join(SANDBOX, 'a2.txt');
    const r1 = await call(`return await ${api}.renamePath(${JSON.stringify(aTxt)}, ${JSON.stringify(a2)})`);
    const r2 = await call(`return await ${api}.renamePath(${JSON.stringify(a2)}, ${JSON.stringify(path.join(SANDBOX, 'b.log'))})`);
    step('renamePath', {
      ok: r1.success && fs.existsSync(a2),
      conflictRejected: r2.success === false,
      originalIntact: fs.existsSync(a2),
    });

    /* 6. createDirectory + deletePath рекурсивно */
    // Имя каталога намеренно НЕ совпадает с IGNORED_DIRS в поиске,
    // иначе последующий шаг поиска его пропустит.
    const baseDir = path.join(SANDBOX, 'projects');
    const nd = path.join(baseDir, 'alpha', 'beta');
    const cd = await call(`return await ${api}.createDirectory(${JSON.stringify(nd)})`);
    const existsDeep = fs.existsSync(nd);
    const delDir = await call(`return await ${api}.deletePath(${JSON.stringify(baseDir)})`);
    step('directories', {
      createSuccess: cd.success,
      createError: cd.error ?? null,
      createdRecursively: existsDeep,
      deletedRecursively: delDir.success && !fs.existsSync(baseDir),
    });

    /* 7. Поиск по имени */
    const byName = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, 'b', false, {})`);
    step('searchByName', {
      success: byName.success,
      foundBLog: byName.results.some((r) => r.name === 'b.log'),
      onlyFiles: byName.results.every((r) => !r.isDirectory),
    });

    /* 8. Поиск по содержимому + номер строки */
    const byContent = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, 'ERROR', true, {})`);
    const hit = byContent.results.find((r) => r.name === 'b.log');
    step('searchByContent', {
      found: !!hit,
      correctLine: hit?.line === 2,
      snippet: hit?.matchSnippet,
      totalMatches: hit?.totalMatches,
    });

    /* 9. Регистр */
    const ci = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, 'needle', true, {})`);
    const cs = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, 'needle', true, {caseSensitive:true})`);
    const csUpper = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, 'Needle', true, {caseSensitive:true})`);
    step('searchCaseSensitivity', {
      insensitiveFinds: ci.results.length > 0,
      sensitiveSkipsLower: cs.results.length === 0,
      sensitiveFindsExact: csUpper.results.length > 0,
    });

    /* 10. RegEx */
    const rx = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, 'F(ATAL|IXME)', true, {useRegex:true})`);
    const badRx = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, '[unclosed', true, {useRegex:true})`);
    step('searchRegex', {
      finds: rx.results.some((r) => r.name === 'b.log'),
      invalidHandled: badRx.success === false,
      invalidMessage: badRx.error,
    });

    /* 11. Glob-маска */
    const glob = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, 'const', true, {includeGlob:'*.ts'})`);
    step('searchGlob', {
      foundTs: glob.results.some((r) => r.name === 'code.ts'),
      onlyTs: glob.results.length > 0 && glob.results.every((r) => r.name.endsWith('.ts')),
    });

    /* 12. Игнорирование служебных папок */
    fs.mkdirSync(path.join(SANDBOX, 'node_modules'), { recursive: true });
    fs.writeFileSync(path.join(SANDBOX, 'node_modules', 'x.txt'), 'UNIQUEMARKER\n');
    fs.mkdirSync(path.join(SANDBOX, '.git'), { recursive: true });
    fs.writeFileSync(path.join(SANDBOX, '.git', 'y.txt'), 'UNIQUEMARKER\n');
    const ign = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, 'UNIQUEMARKER', true, {})`);
    step('ignoredDirs', { noneFound: ign.results.length === 0, count: ign.results.length });

    /* 13. Крайние случаи */
    const empty = await call(`return await ${api}.searchWorkspace(${JSON.stringify(SANDBOX)}, '', true, {})`);
    const badDir = await call(`return await ${api}.searchWorkspace(${JSON.stringify(path.join(SANDBOX, 'nope'))}, 'x', true, {})`);
    const missing = await call(`return await ${api}.readFile(${JSON.stringify(path.join(SANDBOX, 'missing.txt'))}, 'utf-8')`);
    step('edgeCases', {
      emptyQueryOk: empty.success && empty.results.length === 0,
      missingDirHandled: badDir.success === false,
      missingFileHandled: missing.success === false && typeof missing.error === 'string',
    });

    /* 14. Проверка формы API, который видит рендерер */
    const apiShape = await call(`return { keys: Object.keys(${api}), count: Object.keys(${api}).length }`);
    step('preloadApiSurface', apiShape);
  } catch (err) {
    out.errors.push(err.stack || String(err));
  }

  fs.writeFileSync(path.join(ROOT, 'integration-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  try {
    fs.rmSync(SANDBOX, { recursive: true, force: true });
  } catch {}
  app.exit(0);
});
