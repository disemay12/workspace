import { useCallback, useEffect, useState } from 'react';

/** Чтение ключа; битый JSON не затираем молча — оставляем копию для диагностики. */
function readStored<T>(key: string, fallback: T): T {
  let raw: string | null = null;
  try {
    raw = window.localStorage.getItem(key);
  } catch {
    return fallback;
  }
  if (raw === null) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch {
    try {
      window.localStorage.setItem(`${key}.__broken`, raw);
    } catch {
      // Quota или приватный режим — копию сохранить некуда
    }
    return fallback;
  }
}

/**
 * Хук для состояния, синхронизированного с localStorage.
 * Используется для UI-настроек: тема, ширины панелей, активный вид.
 */
export function useLocalStorage<T>(
  key: string,
  initialValue: T
): [T, (value: T | ((prev: T) => T)) => void] {
  // Значение вместе с ключом: иначе при смене key в новый ключ уходило старое значение.
  const [entry, setEntry] = useState<{ key: string; value: T }>(() => ({
    key,
    value: readStored(key, initialValue),
  }));

  useEffect(() => {
    setEntry((prev) => (prev.key === key ? prev : { key, value: readStored(key, initialValue) }));
  }, [key, initialValue]);

  useEffect(() => {
    // Пишем только когда значение уже прочитано для текущего ключа.
    if (entry.key !== key) return;
    try {
      window.localStorage.setItem(key, JSON.stringify(entry.value));
    } catch {
      // Quota или приватный режим — настройка просто не сохранится
    }
  }, [key, entry]);

  useEffect(() => {
    // Соседнее окно сохранило настройку — иначе окна расходятся до перезапуска.
    const onStorage = (event: StorageEvent) => {
      if (event.key !== key) return;
      setEntry({
        key,
        value: event.newValue === null ? initialValue : readStored(key, initialValue),
      });
    };
    window.addEventListener('storage', onStorage);
    return () => window.removeEventListener('storage', onStorage);
  }, [key, initialValue]);

  const setValue = useCallback(
    (value: T | ((prev: T) => T)) => {
      setEntry((prev) => ({
        key,
        value:
          typeof value === 'function' ? (value as (prev: T) => T)(prev.value) : value,
      }));
    },
    [key]
  );

  return [entry.value, setValue];
}
