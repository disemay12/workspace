import { useCallback, useEffect, useState } from 'react';

/**
 * Хук для состояния, синхронизированного с localStorage.
 * Используется для UI-настроек: тема, ширины панелей, активный вид.
 */
export function useLocalStorage<T>(
  key: string,
  initialValue: T
): [T, (value: T | ((prev: T) => T)) => void] {
  const [stored, setStored] = useState<T>(() => {
    try {
      const raw = window.localStorage.getItem(key);
      if (raw === null) return initialValue;
      return JSON.parse(raw) as T;
    } catch {
      return initialValue;
    }
  });

  useEffect(() => {
    try {
      window.localStorage.setItem(key, JSON.stringify(stored));
    } catch {
      // Quota или приватный режим — настройка просто не сохранится
    }
  }, [key, stored]);

  return [stored, setStored];
}
