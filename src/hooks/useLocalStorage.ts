import { useCallback, useEffect, useRef, useState } from 'react';

interface StoredRead<T> {
  value: T;
  /** Значение было битым: копия сохранена рядом, исходный ключ ещё не перезаписан. */
  broken: boolean;
}

/** Чтение ключа; битый JSON не затираем молча — оставляем копию для диагностики. */
function readStored<T>(key: string, fallback: T): StoredRead<T> {
  let raw: string | null = null;
  try {
    raw = window.localStorage.getItem(key);
  } catch {
    return { value: fallback, broken: false };
  }
  if (raw === null) return { value: fallback, broken: false };
  try {
    return { value: JSON.parse(raw) as T, broken: false };
  } catch (err) {
    try {
      window.localStorage.setItem(`${key}.__broken`, raw);
    } catch {
      // Quota или приватный режим — копию сохранить некуда
    }
    // Сброс настройки должен быть виден: иначе пользователь не понимает,
    // почему вид приложения вернулся к значениям по умолчанию.
    void window.electronAPI
      ?.diagnosticsLog?.(
        'warn',
        `Настройка ${key} повреждена и сброшена (${(err as Error).message}); копия: ${key}.__broken`
      )
      ?.catch?.(() => undefined);
    return { value: fallback, broken: true };
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
  const [entry, setEntry] = useState<{ key: string; value: T; broken: boolean }>(() => {
    const read = readStored(key, initialValue);
    return { key, value: read.value, broken: read.broken };
  });
  /** Битое значение не перезаписываем, пока пользователь сам не поменяет настройку. */
  const pendingUserChange = useRef(entry.broken);

  useEffect(() => {
    setEntry((prev) => {
      if (prev.key === key) return prev;
      const read = readStored(key, initialValue);
      pendingUserChange.current = read.broken;
      return { key, value: read.value, broken: read.broken };
    });
  }, [key, initialValue]);

  useEffect(() => {
    // Пишем только когда значение уже прочитано для текущего ключа.
    if (entry.key !== key) return;
    // Битый исходник остаётся на месте: копия уже сохранена в .__broken,
    // а затирать её дефолтами без ведома пользователя нельзя.
    if (pendingUserChange.current) return;
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
      const read =
        event.newValue === null
          ? { value: initialValue, broken: false }
          : readStored(key, initialValue);
      pendingUserChange.current = false;
      setEntry({ key, value: read.value, broken: read.broken });
    };
    window.addEventListener('storage', onStorage);
    return () => window.removeEventListener('storage', onStorage);
  }, [key, initialValue]);

  const setValue = useCallback(
    (value: T | ((prev: T) => T)) => {
      pendingUserChange.current = false;
      setEntry((prev) => {
        const next =
          typeof value === 'function' ? (value as (prev: T) => T)(prev.value) : value;
        if (prev.broken) {
          // Пользователь задал значение сам — копия повреждённого больше не нужна.
          try {
            window.localStorage.removeItem(`${key}.__broken`);
          } catch {
            // Хранилище недоступно — копия останется, это не ошибка.
          }
        }
        return { key, value: next, broken: false };
      });
    },
    [key]
  );

  return [entry.value, setValue];
}
