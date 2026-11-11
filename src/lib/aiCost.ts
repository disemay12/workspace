/**
 * Оценка стоимости запросов к AI: таблица цен за 1M токенов и расчёт по usage.
 *
 * Цены — ориентир для отображения в интерфейсе (провайдеры меняют их чаще,
 * чем выходят релизы WorkSpace), поэтому таблица живёт отдельно от кода чата
 * и может быть переопределена настройками.
 */
export interface TokenUsage {
  promptTokens: number;
  completionTokens: number;
}

export interface ModelPrice {
  /** Доллары за 1M входных токенов. */
  input: number;
  /** Доллары за 1M выходных токенов. */
  output: number;
}

/** Сопоставление по префиксу модели: порядок важен (сначала длинные). */
const PRICES: { match: string; price: ModelPrice }[] = [
  { match: 'gpt-4.1-mini', price: { input: 0.4, output: 1.6 } },
  { match: 'gpt-4.1-nano', price: { input: 0.1, output: 0.4 } },
  { match: 'gpt-4.1', price: { input: 2, output: 8 } },
  { match: 'gpt-4o-mini', price: { input: 0.15, output: 0.6 } },
  { match: 'gpt-4o', price: { input: 2.5, output: 10 } },
  { match: 'o3-mini', price: { input: 1.1, output: 4.4 } },
  { match: 'o1', price: { input: 15, output: 60 } },
  { match: 'claude-3-5-haiku', price: { input: 0.8, output: 4 } },
  { match: 'claude-3-5-sonnet', price: { input: 3, output: 15 } },
  { match: 'claude-sonnet-4', price: { input: 3, output: 15 } },
  { match: 'claude-opus', price: { input: 15, output: 75 } },
  { match: 'gemini-2.0-flash', price: { input: 0.1, output: 0.4 } },
  { match: 'gemini-1.5-flash', price: { input: 0.075, output: 0.3 } },
  { match: 'gemini-1.5-pro', price: { input: 1.25, output: 5 } },
  { match: 'deepseek-chat', price: { input: 0.27, output: 1.1 } },
  { match: 'deepseek-reasoner', price: { input: 0.55, output: 2.19 } },
];

export const DEFAULT_PRICE: ModelPrice = { input: 0, output: 0 };

export function priceForModel(model: string): ModelPrice {
  const name = (model || '').toLowerCase();
  for (const entry of PRICES) {
    if (name.startsWith(entry.match)) return entry.price;
  }
  return DEFAULT_PRICE;
}

/** Стоимость в долларах; при неизвестной модели — 0 (цена не показывается). */
export function estimateCost(model: string, usage: TokenUsage): number {
  const price = priceForModel(model);
  if (price.input === 0 && price.output === 0) return 0;
  const prompt = Math.max(0, Number(usage.promptTokens) || 0);
  const completion = Math.max(0, Number(usage.completionTokens) || 0);
  const cost = (prompt / 1_000_000) * price.input + (completion / 1_000_000) * price.output;
  return Math.round(cost * 1_000_000) / 1_000_000;
}

/** Компактная строка: «$0.0123» или «< $0.0001» для микроскопических сумм. */
export function formatCost(usd: number): string {
  const value = Number(usd) || 0;
  if (value <= 0) return '';
  if (value < 0.0001) return '< $0.0001';
  return `$${value.toFixed(4)}`;
}

/** Итоговая строка для интерфейса: токены и деньги (если цена известна). */
export function describeUsage(model: string, usage: TokenUsage): string {
  const total = (Number(usage.promptTokens) || 0) + (Number(usage.completionTokens) || 0);
  const cost = formatCost(estimateCost(model, usage));
  return cost ? `${total} токенов · ${cost}` : `${total} токенов`;
}
