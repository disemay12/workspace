/**
 * Безопасное разрешение пути внутри рабочей папки.
 *
 * Ассистент читает только файлы workspace: `..`, абсолютные пути вне корня и
 * переходы по символическим ссылкам не должны выводить за его пределы.
 */
export function resolveWorkspacePath(root: string, requested: string): string | null {
  const raw = String(requested ?? '')
    .trim()
    .replace(/^["']|["']$/g, '');
  if (!raw || !root.trim()) return null;

  const normalizedRoot = root.replace(/\\/g, '/').replace(/\/+$/, '');
  const absolute = /^[a-zA-Z]:[\\/]/.test(raw) || raw.startsWith('/')
    ? raw
    : `${normalizedRoot}/${raw}`;

  const stack: string[] = [];
  for (const part of absolute.replace(/\\/g, '/').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') {
      if (stack.length === 0) return null;
      stack.pop();
      continue;
    }
    stack.push(part);
  }

  const rootParts = normalizedRoot.split('/').filter(Boolean);
  if (stack.length < rootParts.length) return null;
  const prefix = stack.slice(0, rootParts.length).join('/').toLowerCase();
  if (prefix !== rootParts.join('/').toLowerCase()) return null;

  return stack.join('/').replace(/\//g, '\\');
}
