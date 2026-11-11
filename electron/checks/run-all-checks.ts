/**
 * Итоговый прогон всех проверок: типы, сборка и функциональные наборы.
 * Запускает каждый сценарий в отдельном процессе Electron и сводит результат.
 */
const { execFileSync } = require('child_process');
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

const results: any[] = [];

function run(title, fn) {
  const started = Date.now();
  try {
    const detail = fn();
    results.push({ title, ok: true, ms: Date.now() - started, detail });
  } catch (err) {
    results.push({ title, ok: false, ms: Date.now() - started, detail: err.message });
  }
}

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

function exec(cmd, args, opts = {}) {
  return execFileSync(cmd, args, {
    cwd: ROOT,
    encoding: 'utf-8',
    stdio: ['ignore', 'pipe', 'pipe'],
    // На Windows .cmd-обёртки (tsc.cmd, vite.cmd) требуют shell,
    // иначе spawn завершается с EINVAL.
    shell: process.platform === 'win32',
    env: { ...process.env, ELECTRON_RUN_AS_NODE: undefined },
    ...opts,
  });
}

/* ---------- 0. Компиляция проверочных скриптов ---------- */
run('compile: check scripts', () => {
  exec(path.join(NPX, 'tsc.cmd'), ['-p', 'electron/checks/tsconfig.json']);
  const nestedDir = path.join(ROOT, 'dist-electron', 'electron', 'checks');
  if (fs.existsSync(nestedDir)) {
    for (const file of fs.readdirSync(nestedDir)) {
      if (file.endsWith('.js')) {
        fs.copyFileSync(path.join(nestedDir, file), path.join(ROOT, 'dist-electron', file));
      }
    }
  }
  const needed = [
    'integration-check.js',
    'render-check.js',
    'feature-check.js',
    'ai-check.js',
    'insert-check.js',
    'doc-agent-check.js',
    'conversion-check.js',
    'v11-check.js',
  ];
  for (const f of needed) {
    if (!fs.existsSync(path.join(ROOT, 'dist-electron', f))) {
      throw new Error(`dist-electron/${f} не создан`);
    }
  }
  return `${needed.length} скриптов`;
});

/* ---------- 1. Проверка типов ---------- */
run('typecheck: frontend (tsconfig.json)', () => {
  exec(path.join(NPX, 'tsc.cmd'), ['-p', 'tsconfig.json', '--noEmit']);
  return 'no errors';
});

run('typecheck: electron (tsconfig.electron.json)', () => {
  exec(path.join(NPX, 'tsc.cmd'), ['-p', 'tsconfig.electron.json', '--noEmit']);
  return 'no errors';
});

/* ---------- 2. Сборка ---------- */
run('build: vite production bundle', () => {
  const out = exec(path.join(NPX, 'vite.cmd'), ['build']);
  const match = out.match(/built in ([\d.]+m?s)/);
  const chunks = (out.match(/dist\/assets\//g) || []).length;
  if (!fs.existsSync(path.join(ROOT, 'dist', 'index.html'))) {
    throw new Error('dist/index.html не создан');
  }
  return `${chunks} assets, built in ${match ? match[1] : '?'}`;
});

run('build: electron main process', () => {
  exec(path.join(NPX, 'tsc.cmd'), ['-p', 'tsconfig.electron.json']);
  for (const f of ['main.js', 'preload.js']) {
    if (!fs.existsSync(path.join(ROOT, 'dist-electron', f))) {
      throw new Error(`dist-electron/${f} не создан`);
    }
  }
  return 'main.js + preload.js';
});

/* ---------- 3. Функциональные наборы ---------- */
const suites = [
  ['IPC и файловая система', 'integration-check.js', 'integration-check.json'],
  ['Рендеринг интерфейса', 'render-check.js', 'render-check.json'],
  ['Тема, поиск, настройки', 'feature-check.js', 'feature-check.json'],
  ['AI: стриминг и Ctrl+K', 'ai-check.js', 'ai-check.json'],
  ['AI: вставка кода в редактор', 'insert-check.js', 'insert-check.json'],
  ['AI: агентные правки .docx', 'doc-agent-check.js', 'doc-agent-check.json'],
  ['Конвертация форматов и ширины Excel', 'conversion-check.js', 'conversion-check.json'],
    ['Фото-хостинги, зум видео, Google OAuth', 'v11-check.js', 'v11-check.json'],
];

// Сквозные проверки с реальным API подключаются только при наличии ключа,
// чтобы обычный прогон не зависел от сети.
if (fs.existsSync(path.join(ROOT, 'secrets.local.ps1'))) {
  // Проверяет обход CORS: браузерный fetch к части шлюзов блокируется
  // с net::ERR_FAILED, поэтому запросы идут через главный процесс.
  suites.push(['AI: обход CORS через прокси', 'cors-check.js', 'cors-check.json']);
  suites.push(['Сквозная проверка с реальным AI', 'live-check.js', 'live-check.json']);
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
    }
  }
  return checks;
}

for (const [title, script, resultFile] of suites) {
  run(title, () => {
    const resultPath = path.join(ROOT, resultFile);
    if (fs.existsSync(resultPath)) fs.unlinkSync(resultPath);
    exec(ELECTRON, [...sandboxFlags(), path.join('dist-electron', script)]);
    if (!fs.existsSync(resultPath)) {
      throw new Error(`набор не создал ${resultFile}`);
    }
    const data = JSON.parse(fs.readFileSync(resultPath, 'utf-8'));
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

    // Провалами считаются только утвердительные ожидания (true),
    // которые оказались false. Часть полей фиксирует состояние, которое
    // по смыслу должно быть false: сломанный текст при неверной кодировке,
    // отклонённые конфликты, отсутствие совпадений в игнорируемых папках,
    // а также состояние темы ДО её переключения.
    const negative =
      /Broken|Skips|Rejected|Handled|noneFound|onlyTs|Gone|notConfigured|stillStreaming|originalGone|caseSensitive|BeforeToggle/i;
    const failures = [];
    for (const { name, value } of checks) {
      if (value === false && !negative.test(name)) failures.push(name);
    }
    if (failures.length > 0) {
      throw new Error(`провалы: ${failures.join(', ')}`);
    }
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
fs.writeFileSync(path.join(ROOT, 'final-report.txt'), report, 'utf-8');
console.log(report);
process.exit(failed.length > 0 ? 1 : 0);
