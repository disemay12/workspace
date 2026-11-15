/**
 * Границы записи и чтения для запросов из рендерера.
 *
 * Рендерер — граница доверия: любая XSS-инъекция в документе или ответе модели
 * получает мост electronAPI, то есть произвольные пути. Проверки ниже стоят на
 * каждом IPC-обработчике, который пишет, удаляет или отдаёт файл.
 *
 * Сравнение путей идёт по канонической форме: Windows трактует `C:\Windows.` и
 * `C:\PROGRA~1` как те же каталоги, а `\\?\C:\...` — как префикс, обходящий
 * любое сравнение строк.
 */
import * as fs from 'fs';
import * as path from 'path';

/** Корень открытой рабочей папки: путь к нему нельзя удалять и переименовывать. */
let workspaceRoot = '';
/** Системные каталоги, которые добавляет главный процесс (каталог приложения). */
let extraProtected: string[] = [];
/** Точные файлы программы: их загрузчик применяет при обновлении. */
let protectedFiles: string[] = [];

export function setWorkspaceRoot(value: string): void {
  workspaceRoot = value ? canonicalPath(value) : '';
}

export function getWorkspaceRoot(): string {
  return workspaceRoot;
}

export function setProtectedDirectories(dirs: string[]): void {
  extraProtected = dirs.filter(Boolean).map(canonicalPath);
}

/**
 * Файлы, которые загрузчик и Electron загружают при следующем старте.
 * Рендерер мог записать `WorkSpace.exe.new` рядом с программой, и загрузчик
 * подменял им сам себя — это запуск произвольного кода без обновления.
 */
export function setProtectedFiles(files: string[]): void {
  protectedFiles = files.filter(Boolean).map((file) => canonicalPath(file).toLowerCase());
}

/** Убирает завершающие точки и пробелы: Windows их в именах отбрасывает. */
function trimNameSegments(value: string): string {
  const root = path.parse(value).root;
  const parts = value
    .slice(root.length)
    .split(/[\\/]+/)
    .filter(Boolean)
    .map((part) => part.replace(/[. ]+$/, '') || part);
  return parts.length > 0 ? path.join(root, ...parts) : root;
}

/**
 * Каноническая форма пути: без префиксов выравнивания, с реальными именами
 * существующих каталогов (снимает 8.3-имена и junction'ы).
 */
export function canonicalPath(value: string): string {
  let prepared = String(value ?? '').trim();
  if (!prepared) return '';
  if (/^\\\\\?\\/.test(prepared)) prepared = prepared.replace(/^\\\\\?\\/, '');
  if (/^\\\\\.\\/.test(prepared)) prepared = prepared.replace(/^\\\\\.\\/, '');
  const normalized = trimNameSegments(path.resolve(prepared));
  try {
    return fs.realpathSync.native(normalized);
  } catch {
    // Путь ещё не существует: канонизируем существующего родителя.
  }
  const parent = path.dirname(normalized);
  if (!parent || parent === normalized) return normalized;
  try {
    return path.join(fs.realpathSync.native(parent), path.basename(normalized));
  } catch {
    return normalized;
  }
}

/** Совпадение пути с каталогом (или вложенность в него) без учёта регистра. */
export function isInsidePath(child: string, parent: string): boolean {
  const normalizedChild = canonicalPath(child).replace(/[\\/]+$/, '').toLowerCase();
  const normalizedParent = canonicalPath(parent).replace(/[\\/]+$/, '').toLowerCase();
  if (!normalizedParent) return false;
  return (
    normalizedChild === normalizedParent ||
    normalizedChild.startsWith(normalizedParent + path.sep) ||
    normalizedChild.startsWith(normalizedParent + '/')
  );
}

/** Системные каталоги, в которых приложению нечего писать и удалять. */
export function protectedDirectories(): string[] {
  const env = process.env;
  const dirs: string[] = [];
  const add = (value?: string) => {
    if (value && value.trim()) dirs.push(canonicalPath(value));
  };
  add(env.APPDATA ? path.join(env.APPDATA, 'Microsoft', 'Windows', 'Start Menu') : '');
  add(env.ProgramData ? path.join(env.ProgramData, 'Microsoft', 'Windows', 'Start Menu') : '');
  add(env.SystemRoot ?? env.windir);
  add(env.ProgramFiles);
  add(env['ProgramFiles(x86)']);
  add(env.ProgramW6432);
  return dirs.concat(extraProtected);
}

/**
 * Проверяет, что путь не является корнем диска, корнем рабочей папки при
 * удалении и не лежит в системном каталоге.
 */
export function assertMutablePath(target: string, operation: 'write' | 'delete'): void {
  if (!target || typeof target !== 'string') throw new Error('Не указан путь');
  const resolved = canonicalPath(target);
  const driveRoot = path.parse(resolved).root.replace(/[\\/]+$/, '');
  if (!resolved || resolved.replace(/[\\/]+$/, '').length <= driveRoot.length) {
    throw new Error('Операция над корнем диска отклонена');
  }
  if (operation === 'delete' && workspaceRoot && isInsidePath(workspaceRoot, resolved)) {
    throw new Error('Нельзя удалить или переименовать корень рабочей папки');
  }
  if (protectedFiles.includes(resolved.toLowerCase())) {
    throw new Error('Файл программы защищён от изменения из интерфейса');
  }
  for (const protectedDir of protectedDirectories()) {
    if (isInsidePath(resolved, protectedDir)) {
      throw new Error(`Каталог защищён: ${protectedDir}`);
    }
  }
}

/**
 * Файлы и каталоги, которые рендереру читать нельзя ни при каких условиях:
 * это ключи и настройки самого приложения. Путь к ним уже известен рендереру
 * (диагностика показывает userData), поэтому запрет нужен именно на чтении.
 */
const APP_SECRET_RE =
  /(^|[\\/])(secrets\.local\.[^\\/]*|ai-settings\.json|gdrive\.json|workspace-session\.json|pending-update\.txt|launcher\.json|update\.log)$/i;
/**
 * Секреты пользователя. Вне открытой рабочей папки они закрыты, внутри —
 * открываются: редактор кода обязан показывать `.env` и `.pem` проекта,
 * иначе работа с ним невозможна.
 */
const USER_SECRET_RE =
  /(^|[\\/])(\.env[^\\/]*|id_(rsa|dsa|ecdsa|ed25519)[^\\/]*|[^\\/]+\.(pem|ppk|pfx|p12|kdbx))$/i;
const SENSITIVE_DIRS = ['.ssh', '.aws', '.gnupg', '.password-store', 'gcloud'];

function matchesSensitiveDir(canonical: string): boolean {
  return SENSITIVE_DIRS.some(
    (dir) => canonical.includes(`${path.sep}${dir}${path.sep}`) || canonical.endsWith(`${path.sep}${dir}`)
  );
}

export function isSecretPath(target: string): boolean {
  const canonical = canonicalPath(target).toLowerCase();
  if (!canonical) return false;
  // Настройки самого приложения закрыты везде: рабочей папкой может оказаться
  // профиль пользователя, и тогда запрет «снаружи» ничего бы не значил.
  if (APP_SECRET_RE.test(canonical)) return true;
  if (matchesSensitiveDir(canonical) && !isInsidePath(canonical, workspaceRoot)) return true;
  if (!USER_SECRET_RE.test(canonical)) return false;
  return !workspaceRoot || !isInsidePath(canonical, workspaceRoot);
}

/** Бросает, если рендерер пытается прочитать секреты приложения. */
export function assertReadablePath(target: string): void {
  if (!target || typeof target !== 'string') throw new Error('Не указан путь');
  if (isSecretPath(target)) {
    throw new Error('Файл с ключами или настройками приложения недоступен из интерфейса');
  }
}
