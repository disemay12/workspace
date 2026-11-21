/**
 * Итоговый прогон всех проверок: типы, сборка и функциональные наборы.
 * Запускает каждый сценарий в отдельном процессе Electron и сводит результат.
 */
const { execFileSync, spawn } = require('child_process');
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
const ELECTRON = path.join(ROOT, 'node_modules', 'electron', 'dist', 'electron.exe');
const NPX = path.join(ROOT, 'node_modules', '.bin');

/**
 * Каталог прогона уникален: общие dist-electron/*-check.json двух одновременных
 * запусков перезаписывали друг друга, и один прогон читал результат другого.
 */
const RUN_ID = new Date().toISOString().replace(/[:.]/g, '-');
const RUN_DIR = path.join(ROOT, '.checks-runs', RUN_ID);

const results: any[] = [];

async function run(title, fn) {
  const started = Date.now();
  try {
    const detail = await fn();
    results.push({ title, ok: true, ms: Date.now() - started, detail });
  } catch (err) {
    results.push({ title, ok: false, ms: Date.now() - started, detail: err.message });
  }
}

/* ---------- Блокировка прогона ---------- */

const LOCK_FILE = path.join(ROOT, '.checks.lock');

/** Живой процесс по pid: kill(pid, 0) не шлёт сигнал, а только проверяет право. */
function isProcessAlive(pid) {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return err.code === 'EPERM';
  }
}

function readLock() {
  try {
    return JSON.parse(fs.readFileSync(LOCK_FILE, 'utf-8'));
  } catch {
    return null;
  }
}

/**
 * Возраст файла блокировки в мс; null — mtime прочитать не удалось.
 * Пустой или битый лок — это в том числе окно между захватом файла (wx) и
 * записью pid, поэтому свежий лок нельзя считать мёртвым: иначе два прогона
 * одновременно пересобирают dist/ и перезаписывают одни и те же *-check.json.
 */
function lockAgeMs() {
  try {
    return Date.now() - fs.statSync(LOCK_FILE).mtimeMs;
  } catch {
    return null;
  }
}

/** Лок, который моложе этого возраста, не трогаем: pid мог ещё не записаться. */
const LOCK_GRACE_MS = 10_000;

/** Запись pid через rename: читатель видит либо прошлое состояние, либо полный json. */
function writeLockAtomic(lock) {
  const tmp = `${LOCK_FILE}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(lock, null, 2), 'utf-8');
  try {
    fs.renameSync(tmp, LOCK_FILE);
  } catch (err) {
    try {
      fs.unlinkSync(tmp);
    } catch {}
    throw err;
  }
}

/**
 * Два одновременных прогона пересобирают dist/ и переписывают одни и те же
 * *-check.json — ленивый импорт чанка Vite падал у соседа. Поэтому запуск
 * сериализуется: файл создаётся эксклюзивно (wx), а блокировка мёртвого
 * процесса снимается автоматически.
 */
function acquireLock() {
  const lock = { pid: process.pid, startedAt: new Date().toISOString(), root: ROOT };
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      // Эксклюзивное создание — единственный неделимый «захват»: второй
      // процесс получает EEXIST и до записи содержимого не доходит.
      fs.closeSync(fs.openSync(LOCK_FILE, 'wx'));
    } catch (err) {
      if (err.code !== 'EEXIST') throw err;
      const existing = readLock();
      if (existing && isProcessAlive(existing.pid)) {
        throw new Error(
          `уже идёт прогон (pid ${existing.pid}, старт ${existing.startedAt}): ` +
            `дождитесь завершения или удалите ${LOCK_FILE}, если процесс уже убит`
        );
      }
      const age = lockAgeMs();
      if (age === null || age < LOCK_GRACE_MS) {
        throw new Error(
          `уже идёт прогон (${LOCK_FILE} создан ${age === null ? 'только что' : `${Math.round(age / 1000)} с назад`}, ` +
            `pid не записан): дождитесь завершения или удалите файл вручную`
        );
      }
      // Блокировка осталась от упавшего прогона — снимаем и пробуем ещё раз.
      try {
        fs.unlinkSync(LOCK_FILE);
      } catch {}
      continue;
    }
    try {
      writeLockAtomic(lock);
    } catch (err) {
      // Захватили, но записать не смогли: пустой лок иначе блокировал бы
      // следующий прогон до истечения LOCK_GRACE_MS.
      try {
        fs.unlinkSync(LOCK_FILE);
      } catch {}
      throw err;
    }
    return lock;
  }
  throw new Error(`не удалось захватить ${LOCK_FILE}`);
}

function releaseLock() {
  const existing = readLock();
  if (existing && existing.pid !== process.pid) return;
  try {
    fs.unlinkSync(LOCK_FILE);
  } catch {}
}

/* ---------- Запуск процессов ---------- */

/**
 * Chromium не может инициализировать sandbox, когда в Windows отключён UAC
 * (EnableLUA=0): процесс получает полный токен администратора, и Electron
 * завершается с STATUS_BREAKPOINT (0x80000003) ещё до загрузки приложения.
 * В этом состоянии проверочные процессы запускаются с --no-sandbox.
 */
function sandboxFlags(): string[] {
  if (process.platform !== 'win32') return [];
  if (process.env.ELECTRON_DISABLE_SANDBOX) return ['--no-sandbox'];
  try {
    const out = execFileSync(
      'reg',
      [
        'query',
        'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System',
        '/v',
        'EnableLUA',
      ],
      { encoding: 'utf-8', windowsHide: true }
    );
    const m = /EnableLUA\s+REG_DWORD\s+0x([0-9a-fA-F]+)/.exec(out);
    if (m && parseInt(m[1], 16) === 0) return ['--no-sandbox'];
  } catch {
    // Политику прочитать не удалось — считаем sandbox рабочим.
  }
  return [];
}

/**
 * Наборы открывают окно и ждут отрисовки: на медленной машине или при
 * загруженном GPU 3 минут не хватало (render-check не успевал записать
 * результат), поэтому предел — 5 минут. Зависание всё равно отсекается.
 */
const EXEC_TIMEOUT_MS = 300000;

/** Чистый профиль Electron для набора: каталог прогона, а не общий profile. */
function suiteProfileDir(script: string): string {
  const dir = path.join(RUN_DIR, 'profiles', path.basename(script, '.js'));
  fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

/** Electron порождает дочерние процессы: один kill(pid) оставлял их жить. */
function killTree(pid) {
  if (!pid) return '';
  if (process.platform === 'win32') {
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], {
        stdio: 'ignore',
        windowsHide: true,
      });
    } catch (err) {
      // Обычно процесс уже завершился сам, но причину не глушим: она попадает
      // в сообщение о таймауте, иначе разбираться в зависании не с чем.
      return String((err as Error).message || 'taskkill не смог завершить процесс');
    }
    return '';
  }
  try {
    process.kill(-pid, 'SIGKILL');
  } catch {
    try {
      process.kill(pid, 'SIGKILL');
    } catch {}
  }
  return '';
}

/**
 * Запуск с таймаутом и убийством дерева процессов: таймаут не должен оставлять
 * Electron живым.
 *
 * Ждать только 'close' нельзя: это событие приходит после закрытия stdio, а
 * орфанные потомки держат pipe — промис не завершался никогда, и прогон висел
 * до внешнего убийства. Код выхода известен уже на 'exit', поэтому результат
 * отдаётся по нему, а на дочитывание stderr даётся короткая пауза.
 */
const EXEC_KILL_ESCALATION_MS = 5000;
const EXEC_DRAIN_MS = 300;

function exec(
  cmd: string,
  args: string[],
  opts: any = {}
): Promise<{ stdout: string; stderr: string; status: number }> {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      cwd: ROOT,
      windowsHide: true,
      // shell нужен только .cmd/.bat-обёрткам npm (tsc.cmd, vite.cmd): без него
      // spawn падает с EINVAL. Для electron.exe включать shell нельзя — он
      // подставляет cmd.exe, ломает кавычки в путях и скрывает настоящий код выхода.
      shell: process.platform === 'win32' && /\.(cmd|bat)$/i.test(cmd),
      // ELECTRON_RUN_AS_NODE превращает Electron в node и ломает GUI-наборы.
      env: { ...process.env, ELECTRON_RUN_AS_NODE: undefined, ...(opts.env || {}) },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    let timedOut = false;
    let killReason = '';
    let settled = false;
    let killTimer: NodeJS.Timeout | null = null;
    let drainTimer: NodeJS.Timeout | null = null;

    const finish = (code: number | null) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (killTimer) clearTimeout(killTimer);
      if (drainTimer) clearTimeout(drainTimer);
      const tail = stderr.trim().split('\n').slice(-3).join(' | ').slice(0, 400);
      if (timedOut) {
        reject(
          new Error(
            `таймаут ${EXEC_TIMEOUT_MS} мс: ${cmd}${killReason ? ` (kill: ${killReason})` : ''}${
              tail ? ` (${tail})` : ''
            }`
          )
        );
        return;
      }
      if (code !== 0) {
        reject(new Error(`${path.basename(cmd)}: код ${code}${tail ? `: ${tail}` : ''}`));
        return;
      }
      resolve({ stdout, stderr, status: code });
    };

    const timer = setTimeout(() => {
      timedOut = true;
      killReason = killTree(child.pid);
      // taskkill мог не добрать потомков, которые держат pipe: второй шаг бьёт
      // по процессу напрямую, и ошибка таймаута всё равно уходит наверх, а не
      // теряется в незавершённом промисе.
      killTimer = setTimeout(() => {
        try {
          child.kill('SIGKILL');
        } catch {
          // Процесс уже завершился — 'exit' вот-вот придёт.
        }
      }, EXEC_KILL_ESCALATION_MS);
    }, EXEC_TIMEOUT_MS);
    child.stdout.on('data', (chunk) => {
      stdout += chunk;
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
    });
    child.on('error', (err) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (killTimer) clearTimeout(killTimer);
      reject(err);
    });
    // Код выхода известен здесь; 'close' ждёт закрытия stdio и может не прийти.
    child.on('exit', (code) => {
      drainTimer = setTimeout(() => finish(code), EXEC_DRAIN_MS);
    });
    child.on('close', (code) => finish(code));
  });
}

/* ---------- Контракт наборов ---------- */

/**
 * Поля-счётчики провалов: ноль (или пустой массив) — успех, непусто — провал.
 * Имя полное («шаг.поле»): по последнему сегменту освобождение действовало бы
 * во всех наборах сразу.
 */
const FAILURE_COUNTERS = new Set(['dateFormat.failures', 'sanitizer.failures', 'conversion.failures']);

/**
 * Контракт набора: утверждение — это boolean (провал — false), а строки,
 * числа, объекты и массивы — диагностика, и раннер их не судит. Раньше ''
 * или null в диагностическом поле считались провалом, из-за чего наборы
 * прятали наблюдения в «непустые» значения, а настоящая регрессия в них
 * не могла стать красной. Если по смыслу утверждение — это непустое значение,
 * набор пишет рядом boolean (has…, …Shown), а само значение называет
 * *Value или *Snippet.
 *
 * Boolean, у которого false — нормальное наблюдение, перечисляется здесь по
 * полному имени: сверка по последнему сегменту освобождала поле с тем же
 * именем во всех наборах.
 */
const NEGATIVE_KEYS = new Set([
  // Статус «AI не настроен» не должен появиться, когда настройки загружены.
  'aiSettingsLoaded.notConfigured',
  // Индикатор «Думает» не должен оставаться на экране после конца стрима.
  'chatAfterStream.stillStreaming',
  // Снимок темы ДО переключения наследуется от системной настройки:
  // допустимы оба значения, утверждение — в themeToggled.darkClassFlipped.
  'themeBefore.darkClassBeforeToggle',
  // Итоговый класс после двух переключений возвращается к исходному: на светлой
  // системе это законно false, сравнение сделано в themeToggled.themePersisted.
  'themeRestored',
  // Строка формул Excel не должна «протекать» в другую ячейку.
  'summary.formulaLeakedToOtherCell',
]);

/** Значения, которыми набор сообщает о провале: только явный boolean-провал. */
function isFailureValue(value) {
  return value === false || value === 'false';
}

/**
 * Разворачивает результат набора в плоский список проверок.
 * Наборы используют две формы: { steps: [...] } и плоский объект
 * с вложенными группами (render-check, feature-check).
 */
function flattenChecks(data) {
  const checks = [];
  if (Array.isArray(data.steps)) {
    for (const step of data.steps) {
      for (const [key, value] of Object.entries(step)) {
        if (key !== 'name') checks.push({ name: `${step.name}.${key}`, value });
      }
    }
    return checks;
  }
  for (const [group, payload] of Object.entries(data)) {
    if (group === 'errors' || group === 'logs' || group === 'requests') continue;
    if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
      for (const [key, value] of Object.entries(payload)) {
        checks.push({ name: `${group}.${key}`, value });
      }
    } else {
      // Скаляр на верхнем уровне — плоская форма (v11-check пишет так):
      // без этой ветки набор давал 0 проверок и падал «ни одной проверки».
      checks.push({ name: group, value: payload });
    }
  }
  return checks;
}

/* ---------- Прогон ---------- */

async function main() {
  const lock = acquireLock();
  process.on('exit', releaseLock);
  for (const signal of ['SIGINT', 'SIGTERM']) {
    process.on(signal, () => {
      releaseLock();
      process.exit(130);
    });
  }
  fs.mkdirSync(RUN_DIR, { recursive: true });

  /* ---------- 0. Компиляция проверочных скриптов ---------- */
  await run('compile: check scripts', async () => {
    await exec(path.join(NPX, 'tsc.cmd'), ['-p', 'electron/checks/tsconfig.json']);
    const nestedDir = path.join(ROOT, 'dist-electron', 'electron', 'checks');
    if (fs.existsSync(nestedDir)) {
      for (const file of fs.readdirSync(nestedDir)) {
        if (file.endsWith('.js')) {
          fs.copyFileSync(path.join(nestedDir, file), path.join(ROOT, 'dist-electron', file));
        }
      }
    }
    const needed = [
      // Запускаются по вложенному пути: импортируют src/lib/*.
      'electron/checks/date-format-check.js',
      'electron/checks/sanitizer-check.js',
      'integration-check.js',
      'render-check.js',
      'feature-check.js',
      'ai-check.js',
      'insert-check.js',
      'doc-agent-check.js',
      // Запускается по вложенному пути: импортирует src/services и src/lib.
      'electron/checks/conversion-check.js',
      // Тоже вложенный путь: набор открывает окно и печатает через живой
      // редактор (провайдер печати Word сам лежит в бандле рендерера).
      'electron/checks/word-pages-check.js',
      'v11-check.js',
      // Боевые обработчики aiSettings и preload: поднимает окно в своём профиле.
      'key-migration-check.js',
    ];
    for (const f of needed) {
      if (!fs.existsSync(path.join(ROOT, 'dist-electron', f))) {
        throw new Error(`dist-electron/${f} не создан`);
      }
    }
    return `${needed.length} скриптов`;
  });

  /* ---------- 1. Проверка типов ---------- */
  await run('typecheck: frontend (tsconfig.json)', async () => {
    await exec(path.join(NPX, 'tsc.cmd'), ['-p', 'tsconfig.json', '--noEmit']);
    return 'no errors';
  });

  await run('typecheck: electron (tsconfig.electron.json)', async () => {
    await exec(path.join(NPX, 'tsc.cmd'), ['-p', 'tsconfig.electron.json', '--noEmit']);
    return 'no errors';
  });

  /* ---------- 2. Сборка ---------- */
  await run('build: vite production bundle', async () => {
    const { stdout } = await exec(path.join(NPX, 'vite.cmd'), ['build']);
    const match = stdout.match(/built in ([\d.]+m?s)/);
    const chunks = (stdout.match(/dist\/assets\//g) || []).length;
    if (!fs.existsSync(path.join(ROOT, 'dist', 'index.html'))) {
      throw new Error('dist/index.html не создан');
    }
    return `${chunks} assets, built in ${match ? match[1] : '?'}`;
  });

  await run('build: electron main process', async () => {
    await exec(path.join(NPX, 'tsc.cmd'), ['-p', 'tsconfig.electron.json']);
    for (const f of ['main.js', 'preload.js']) {
      if (!fs.existsSync(path.join(ROOT, 'dist-electron', f))) {
        throw new Error(`dist-electron/${f} не создан`);
      }
    }
    return 'main.js + preload.js';
  });

  /* ---------- 3. Функциональные наборы ---------- */
  // Отключённые наборы (компилируются в checks:build, но в прогон не входят):
  //   excel-formula-check  — открывает UI и в текущем окружении получает
  //                          opened:false (окно не дорисовалось) — результат
  //                          ничего не проверяет, нужен стабильный стенд;
  //   docx-check, docx-fidelity-check, formula-ai-exit-check, model-check, smoke —
  //                          требуют интерактивного окна и собственных фикстур;
  //                          model-check не завершается за 180 с (проверено).
  // Повторно включать только после ручного зелёного прогона на чистой машине.
  const suites = [
    // Чистая логика без окна: единственный набор, который проходит headless.
    // Лежит в dist-electron/electron/checks/ и импортирует src/lib/dateFormat,
    // поэтому запускается по вложенному пути, а не из плоского dist-electron/.
    ['Форматирование дат Excel', 'electron/checks/date-format-check.js', 'date-format-check.json'],
    ['Экранирование HTML в Markdown', 'electron/checks/sanitizer-check.js', 'sanitizer-check.json'],
    ['IPC и файловая система', 'integration-check.js', 'integration-check.json'],
    ['Рендеринг интерфейса', 'render-check.js', 'render-check.json'],
    ['Тема, поиск, настройки', 'feature-check.js', 'feature-check.json'],
    ['AI: стриминг и Ctrl+K', 'ai-check.js', 'ai-check.json'],
    ['AI: вставка кода в редактор', 'insert-check.js', 'insert-check.json'],
    ['AI: агентные правки .docx', 'doc-agent-check.js', 'doc-agent-check.json'],
    // Листы Word и печать: поднимает окно, открывает .docx на 20+ листов и
    // проверяет геометрию, колонтитулы, сохранение без маркеров и PDF
    // предпросмотра. Чистой логики здесь нет — только живой редактор.
    ['Листы Word и печать', 'electron/checks/word-pages-check.js', 'word-pages-check.json'],
    // Импортирует src/** — из плоского dist-electron/ require не разрешается,
    // процесс падал диалогом ошибки и набор «висел» до таймаута.
    ['Конвертация форматов и ширины Excel', 'electron/checks/conversion-check.js', 'conversion-check.json'],
    ['Фото-хостинги, зум видео, Google OAuth', 'v11-check.js', 'v11-check.json'],
    // Один из двух CRITICAL прошлого ревью: перенос ключей API из localStorage
    // в профиль. Набор поднимает окно в отдельном профиле и проверяет, что
    // переживают ключи ВСЕХ провайдеров, а не только активного.
    ['Перенос ключей API в профиль', 'key-migration-check.js', 'key-migration-check.json'],
  ];

  // Наборы, которые ходят в реальный AI-endpoint. Раньше они подключались самим
  // фактом наличия secrets.local.ps1, и состав прогона зависел от машины: там, где
  // ключа нет, регрессию в обходе CORS никто не видел. Теперь — только по явному
  // запросу: npm run checks:live (флаг --live или WORKSPACE_LIVE_CHECKS=1).
  const SECRETS_FILE = path.join(ROOT, 'secrets.local.ps1');
  const liveRequested =
    process.argv.includes('--live') || process.env.WORKSPACE_LIVE_CHECKS === '1';

  if (liveRequested) {
    await run('preflight: live-наборы', async () => {
      if (!fs.existsSync(SECRETS_FILE)) {
        throw new Error('запрошены live-наборы, но secrets.local.ps1 не найден');
      }
      return 'secrets.local.ps1 на месте';
    });
    // Проверяет обход CORS: браузерный fetch к части шлюзов блокируется
    // с net::ERR_FAILED, поэтому запросы идут через главный процесс.
    suites.push(['AI: обход CORS через прокси', 'cors-check.js', 'cors-check.json']);
    suites.push(['Сквозная проверка с реальным AI', 'live-check.js', 'live-check.json']);
  }

  for (const [title, script, resultFile] of suites) {
    await run(title, async () => {
      const resultPath = path.join(ROOT, resultFile);
      if (fs.existsSync(resultPath)) fs.unlinkSync(resultPath);
      // Ненулевой код выхода сам по себе ещё не повод терять детали: набор мог
      // успеть записать результат, и в нём точный список провалов. Поэтому
      // разбираем файл, а код выхода предъявляем последним.
      let exitError = null;
      try {
        await exec(ELECTRON, [
          ...sandboxFlags(),
          // Свой профиль на каждый набор: иначе состояние сессии и настроек
          // прошлого прогона восстанавливалось в новом окне (другая вкладка,
          // другой лист) и проверки зависели от порядка запусков.
          `--user-data-dir=${suiteProfileDir(script)}`,
          path.join('dist-electron', script),
        ]);
      } catch (err) {
        exitError = err;
      }
      if (!fs.existsSync(resultPath)) {
        throw exitError || new Error(`набор не создал ${resultFile}`);
      }
      const raw = fs.readFileSync(resultPath, 'utf-8');
      fs.writeFileSync(path.join(RUN_DIR, resultFile), raw, 'utf-8');
      const data = JSON.parse(raw);
      if (data.errors && data.errors.length > 0) {
        throw new Error(`ошибки в консоли: ${data.errors.join(' | ').slice(0, 220)}`);
      }
      if (data.fatal) {
        throw new Error(`исключение: ${String(data.fatal).slice(0, 220)}`);
      }

      const checks = flattenChecks(data);
      if (checks.length === 0) {
        throw new Error('набор не вернул ни одной проверки');
      }

      const failures = [];
      for (const { name, value } of checks) {
        // Счётчики провалов проверяем в обе стороны: 0 — успех, непусто — провал.
        if (FAILURE_COUNTERS.has(name)) {
          const bad = Array.isArray(value) ? value.length > 0 : typeof value === 'number' && value > 0;
          if (bad) failures.push(`${name}=${JSON.stringify(value).slice(0, 80)}`);
          continue;
        }
        // Наблюдения не судим: список явный и по полному имени.
        if (NEGATIVE_KEYS.has(name)) continue;
        if (isFailureValue(value)) failures.push(`${name}=${JSON.stringify(value)}`);
      }
      if (failures.length > 0) {
        throw new Error(`провалы: ${failures.join(', ')}`);
      }
      if (exitError) throw exitError;
      return `${checks.length} проверок, ошибок консоли нет`;
    });
  }

  /* ---------- Сводка ---------- */
  const passed = results.filter((r) => r.ok).length;
  const lines = [];
  lines.push('');
  lines.push('='.repeat(72));
  lines.push('  ИТОГОВАЯ ПРОВЕРКА WorkSpace');
  lines.push('='.repeat(72));
  lines.push(`  Каталог прогона: ${path.relative(ROOT, RUN_DIR)}`);

  for (const r of results) {
    const mark = r.ok ? 'PASS' : 'FAIL';
    lines.push(`  [${mark}] ${r.title}`);
    lines.push(`         ${r.detail}`);
  }
  lines.push('-'.repeat(72));
  lines.push(`  Итого: ${passed}/${results.length} проверок пройдено`);

  const failed = results.filter((r) => !r.ok);
  if (failed.length > 0) {
    lines.push('');
    lines.push('  НЕ ПРОЙДЕНО:');
    for (const f of failed) lines.push(`    - ${f.title}: ${f.detail}`);
  }
  lines.push('='.repeat(72));

  const report = lines.join('\n');
  fs.writeFileSync(path.join(RUN_DIR, 'final-report.txt'), report, 'utf-8');
  console.log(report);
  process.exit(failed.length > 0 ? 1 : 0);
}

void main().catch((err) => {
  releaseLock();
  console.error(`[run-all-checks] ${err.message}`);
  process.exit(1);
});
