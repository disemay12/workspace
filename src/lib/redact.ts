/**
 * Редакция секретов в диагностических отчётах.
 *
 * Отчёт можно отправить в поддержку/приложить к issue, поэтому из текста
 * вычищаются ключи API, токены, адреса почты и домашний каталог пользователя.
 */
const PATTERNS: { pattern: RegExp; replace: string }[] = [
  // OpenAI/Anthropic/OpenRouter/DeepSeek и совместимые ключи.
  { pattern: /\bsk-[A-Za-z0-9_-]{12,}\b/g, replace: 'sk-<redacted>' },
  { pattern: /\bsk-ant-[A-Za-z0-9_-]{12,}\b/g, replace: 'sk-ant-<redacted>' },
  { pattern: /\bsk-or-[A-Za-z0-9_-]{12,}\b/g, replace: 'sk-or-<redacted>' },
  // Google API/OAuth.
  { pattern: /\bAIza[0-9A-Za-z_-]{20,}\b/g, replace: 'AIza<redacted>' },
  { pattern: /\bya29\.[0-9A-Za-z_-]{10,}\b/g, replace: 'ya29.<redacted>' },
  // Заголовки авторизации.
  { pattern: /(Bearer\s+)[A-Za-z0-9._-]{10,}/gi, replace: '$1<redacted>' },
  // Пары «ключ: значение» для токенов и паролей.
  {
    pattern: /("?(?:api[_-]?key|token|secret|password|access_token|refresh_token)"?\s*[:=]\s*"?)([^"\s,}]{6,})/gi,
    replace: '$1<redacted>',
  },
  // Почта: оставляем только домен, чтобы понять провайдера.
  { pattern: /\b([A-Za-z0-9._%+-])[A-Za-z0-9._%+-]*@([A-Za-z0-9.-]+\.[A-Za-z]{2,})\b/g, replace: '$1***@$2' },
];

/** Домашний каталог и профиль заменяются на переменные окружения. */
export function redactPaths(text: string, homeDir = '', profile = ''): string {
  let result = text;
  const pairs: [string, string][] = [
    [homeDir, '%USERPROFILE%'],
    [profile, '%APPDATA%'],
  ];
  for (const [path, replacement] of pairs) {
    if (!path || path.length < 4) continue;
    const escaped = path.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    result = result.replace(new RegExp(escaped, 'gi'), replacement);
  }
  return result;
}

/** Полная редакция: секреты, почта, домашние пути. */
export function redact(text: string, options: { homeDir?: string; profile?: string } = {}): string {
  let result = String(text ?? '');
  for (const { pattern, replace } of PATTERNS) {
    result = result.replace(pattern, replace);
  }
  return redactPaths(result, options.homeDir ?? '', options.profile ?? '');
}

/** Короткая сводка: сколько секретов и путей вычищено — для отчёта. */
export function redactionStats(text: string, redacted: string): { changed: boolean; removedChars: number } {
  return {
    changed: text !== redacted,
    removedChars: Math.max(0, text.length - redacted.length),
  };
}
