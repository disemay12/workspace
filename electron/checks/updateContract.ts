/**
 * Подключение контракта ответа `update:check` (electron/updateInfo.ts).
 *
 * Проверочные сценарии существуют в двух раскладках: скомпилированные
 * dist-electron/electron/checks/ и скопированные в плоский dist-electron/
 * (checks:sync). Один относительный путь не работает в обеих — из плоской
 * копии `../updateInfo` уходит на уровень выше сборки, — поэтому путь ищется
 * по очереди, а ошибка показывает оба варианта.
 */
import type { UnsupportedUpdateInfo } from '../updateInfo';

export interface UpdateInfoContract {
  unsupportedUpdateInfo(currentVersion: string): UnsupportedUpdateInfo;
}

let cached: UpdateInfoContract | null = null;

export function loadUpdateInfoContract(): UpdateInfoContract {
  if (cached) return cached;
  const failures: string[] = [];
  // checks:build кладёт electron/updateInfo.ts в dist-electron/electron/, поэтому
  // из вложенной раскладки это ../updateInfo, а из плоской копии (checks:sync) —
  // ./electron/updateInfo. Путь ../updateInfo из плоской копии ведёт за пределы
  // сборки, поэтому он идёт первым только для вложенного случая.
  for (const candidate of ['../updateInfo', './electron/updateInfo', './updateInfo']) {
    try {
      cached = require(candidate) as UpdateInfoContract;
      return cached;
    } catch (err) {
      failures.push(`${candidate}: ${(err as Error).message}`);
    }
  }
  throw new Error(`не найден контракт update:check: ${failures.join('; ')}`);
}
