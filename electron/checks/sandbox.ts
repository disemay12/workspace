/**
 * Каталог для тестовых данных проверочных наборов.
 *
 * os.tmpdir() не всегда доступен на запись: на этой машине TEMP указывает на
 * C:\Temp, где mkdir падает с EPERM (низкий уровень целостности процесса).
 * Раньше это давало unhandled rejection и набор молча висел без результата.
 * Поэтому пробуем temp, а при отказе — каталог проекта (он в .gitignore
 * шаблоном test-workspace*).
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

/**
 * Префикс — единственная часть пути, приходящая из кода набора: без проверки
 * опечатка (`../`, `C:\`) уводила бы тестовые данные за пределы проекта.
 */
const PREFIX_RE = /^[a-z0-9-]+$/;

export function createCheckSandbox(prefix: string, projectRoot: string): string {
  if (!PREFIX_RE.test(prefix)) {
    throw new Error(`Неподходящий префикс каталога проверки: ${JSON.stringify(prefix)}`);
  }
  // Date.now() не уникален: два набора, стартовавших в одну миллисекунду,
  // работали бы в общем каталоге, поэтому добавлены pid и случайный суффикс.
  const stamp = `${Date.now()}-${process.pid}-${Math.random().toString(36).slice(2, 8)}`;
  const candidates = [
    path.join(os.tmpdir(), `ws-${prefix}-test-${stamp}`),
    path.join(projectRoot, `test-workspace-${prefix}-${stamp}`),
  ];
  const failures: string[] = [];
  for (const candidate of candidates) {
    try {
      fs.mkdirSync(candidate, { recursive: true });
      return candidate;
    } catch (err) {
      // Причина первого отказа (EPERM на C:\Temp) иначе пропадает в общем
      // «не удалось создать каталог», и разбираться приходится наугад.
      failures.push(`${candidate}: ${(err as Error).message}`);
    }
  }
  throw new Error(
    `Не удалось создать каталог тестовых данных для ${prefix}: ${failures.join('; ')}`
  );
}
