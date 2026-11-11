/**
 * Выборочное применение правок из ответа AI: списки операций и исключения.
 *
 * Чистые функции: интерфейс помечает снятые галочки, а перед выполнением
 * операции фильтруются — так превью и фактическое применение не расходятся.
 */
export interface DiffOperationLike {
  tool: string;
  [key: string]: unknown;
}

/** Переключает индекс в списке исключённых (без дублей и мусора). */
export function toggleExcluded(excluded: number[], index: number): number[] {
  const clean = [...new Set(excluded.filter((item) => Number.isInteger(item) && item >= 0))];
  return clean.includes(index) ? clean.filter((item) => item !== index) : [...clean, index].sort((a, b) => a - b);
}

/** Оставляет только выбранные (не исключённые) операции. */
export function filterOperations<T>(operations: T[], excluded: number[]): T[] {
  const drop = new Set(excluded);
  return operations.filter((_, index) => !drop.has(index));
}

/** Подпись вида «2 из 3» для карточки правок. */
export function describeOperations(total: number, excluded: number[]): string {
  const applied = Math.max(0, total - excluded.filter((index) => index >= 0 && index < total).length);
  return `${applied} из ${total}`;
}

/** Итоговая строка для подтверждения: какие инструменты будут выполнены. */
export function summarizeOperations(operations: DiffOperationLike[], excluded: number[]): string {
  const kept = filterOperations(operations, excluded).map((operation) => operation.tool);
  return kept.length === 0 ? 'Правки не выбраны' : `Будет выполнено: ${kept.join(', ')}`;
}
