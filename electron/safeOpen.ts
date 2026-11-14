/**
 * Открытие внешних адресов средствами Windows.
 *
 * shell.openExternal принимает ЛЮБУЮ схему: file: открывает локальный файл,
 * а зарегистрированные протоколы (ms-settings:, search-ms:, приложения)
 * запускают обработчик. Поэтому ссылку из документа, ответа модели или
 * сторонней страницы пропускаем только как http(s).
 */
import { shell } from 'electron';

const EXTERNAL_SCHEMES = new Set(['http:', 'https:', 'mailto:']);

/** true, если адрес можно передать оболочке Windows. */
export function isSafeExternalUrl(value: string): boolean {
  try {
    const url = new URL(String(value ?? ''));
    return EXTERNAL_SCHEMES.has(url.protocol);
  } catch {
    return false;
  }
}

/** Открывает адрес в браузере/почте; опасные схемы игнорирует. */
export function openExternalSafe(value: string): boolean {
  if (!isSafeExternalUrl(value)) return false;
  void shell.openExternal(value).catch(() => undefined);
  return true;
}
