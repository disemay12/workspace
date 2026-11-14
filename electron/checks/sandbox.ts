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

export function createCheckSandbox(prefix: string, projectRoot: string): string {
  const stamp = Date.now();
  const candidates = [
    path.join(os.tmpdir(), `ws-${prefix}-test-${stamp}`),
    path.join(projectRoot, `test-workspace-${prefix}-${stamp}`),
  ];
  for (const candidate of candidates) {
    try {
      fs.mkdirSync(candidate, { recursive: true });
      return candidate;
    } catch {
      // пробуем следующий вариант
    }
  }
  throw new Error(`Не удалось создать каталог тестовых данных для ${prefix}`);
}
