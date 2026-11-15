import { useCallback, useEffect, useRef, useState } from 'react';
import type { UpdateCheckResult } from '../types';

/**
 * Автоматическая проверка обновлений.
 *
 * Раньше проверка запускалась один раз через 4 секунды после старта и молча
 * умирала, если в этот момент не было сети (или GitHub ответил ошибкой): баннер
 * «доступна версия X» появлялся только после ручного нажатия «Проверить
 * обновление» в настройках. Здесь проверка повторяется с backoff, обновляется
 * при возврате в окно и повторяется раз в несколько часов — но без спама
 * запросов при запуске в разработке, где обновление невозможно.
 */

/** Не соревнуемся с загрузкой интерфейса: первый запрос после первого кадра. */
const FIRST_CHECK_DELAY_MS = 3000;
/** Повторы после неудачи: нет сети в момент старта, лимит GitHub, прокси. */
const RETRY_DELAYS_MS = [20_000, 90_000, 300_000];
/** Обновлений нет — перепроверяем редко, но перепроверяем. */
const IDLE_RECHECK_MS = 6 * 60 * 60 * 1000;
/** Возврат в окно после долгой паузы (спящий ноутбук, другой рабочий стол). */
const FOCUS_RECHECK_AFTER_MS = 30 * 60 * 1000;
/** Ручное нажатие не чаще: иначе кнопка и авто-проверка дублируют запрос. */
const MIN_MANUAL_GAP_MS = 10_000;

export interface UpdateCheckState {
  info: UpdateCheckResult | null;
  checking: boolean;
  error: string;
  /** Баннер скрыт для уже показанной версии; новая версия показывается снова. */
  dismissed: boolean;
  dismiss: () => void;
  checkNow: (options?: { manual?: boolean }) => Promise<UpdateCheckResult | null>;
}

export function useUpdateCheck(enabled: boolean): UpdateCheckState {
  const [info, setInfo] = useState<UpdateCheckResult | null>(null);
  const [checking, setChecking] = useState(false);
  const [error, setError] = useState('');
  const [dismissedVersion, setDismissedVersion] = useState('');

  const inFlightRef = useRef<Promise<UpdateCheckResult | null> | null>(null);
  const timerRef = useRef<number | null>(null);
  const attemptsRef = useRef(0);
  const lastAttemptRef = useRef(0);
  /** Ошибка одной и той же причины пишется в диагностику один раз. */
  const loggedErrorRef = useRef('');
  const availableRef = useRef(false);
  const cancelledRef = useRef(false);

  const clearTimer = useCallback(() => {
    if (timerRef.current !== null) {
      window.clearTimeout(timerRef.current);
      timerRef.current = null;
    }
  }, []);

  const schedule = useCallback(
    (delay: number, action: () => void) => {
      clearTimer();
      if (cancelledRef.current) return;
      timerRef.current = window.setTimeout(() => {
        timerRef.current = null;
        if (!cancelledRef.current) action();
      }, delay);
    },
    [clearTimer]
  );

  const checkNow = useCallback(
    async (options?: { manual?: boolean }): Promise<UpdateCheckResult | null> => {
      const manual = options?.manual === true;
      if (!enabled) return null;
      const api = window.electronAPI?.checkForUpdates;
      if (!api) {
        if (manual) setError('Проверка обновлений доступна только в настольной версии');
        return null;
      }
      // Параллельные запросы склеиваются: баннер и настройки могут попросить
      // проверку одновременно.
      if (inFlightRef.current) return inFlightRef.current;
      if (!manual && Date.now() - lastAttemptRef.current < MIN_MANUAL_GAP_MS) return null;

      lastAttemptRef.current = Date.now();
      setChecking(true);
      const task = (async () => {
        try {
          const result = await api();
          if (cancelledRef.current) return result ?? null;
          if (!result) return null;
          // В разработке обновления недоступны: повторять и показывать ошибку
          // «только в собранной версии» бессмысленно.
          if (result.supported === false) {
            setError('');
            setInfo(result);
            return result;
          }
          if (result.success) {
            attemptsRef.current = 0;
            setError('');
            setInfo(result);
            availableRef.current = Boolean(result.available);
            if (result.available) {
              // Нашли обновление — дальнейшие запросы не нужны.
              clearTimer();
            } else {
              schedule(IDLE_RECHECK_MS, () => void checkNow());
            }
            return result;
          }
          const message = result.error || 'Не удалось проверить обновления';
          setError(message);
          if (loggedErrorRef.current !== message) {
            loggedErrorRef.current = message;
            void Promise.resolve(window.electronAPI?.diagnosticsLog?.('warn', `Обновления: ${message}`)).catch(
              () => undefined
            );
          }
          const retryDelay = RETRY_DELAYS_MS[attemptsRef.current];
          if (retryDelay !== undefined) {
            attemptsRef.current += 1;
            schedule(retryDelay, () => void checkNow());
          }
          return result;
        } catch (err) {
          const message = (err as Error).message || 'Ошибка проверки обновлений';
          setError(message);
          if (loggedErrorRef.current !== message) {
            loggedErrorRef.current = message;
            void Promise.resolve(window.electronAPI?.diagnosticsLog?.('warn', `Обновления: ${message}`)).catch(
              () => undefined
            );
          }
          const retryDelay = RETRY_DELAYS_MS[attemptsRef.current];
          if (retryDelay !== undefined) {
            attemptsRef.current += 1;
            schedule(retryDelay, () => void checkNow());
          }
          return null;
        } finally {
          inFlightRef.current = null;
          if (!cancelledRef.current) setChecking(false);
        }
      })();
      inFlightRef.current = task;
      return task;
    },
    [clearTimer, enabled, schedule]
  );

  const dismiss = useCallback(() => {
    setDismissedVersion(info?.latestVersion ?? '');
  }, [info?.latestVersion]);

  // Первая проверка — один раз за запуск окна.
  useEffect(() => {
    if (!enabled) return undefined;
    cancelledRef.current = false;
    attemptsRef.current = 0;
    schedule(FIRST_CHECK_DELAY_MS, () => void checkNow());
    return () => {
      cancelledRef.current = true;
      clearTimer();
    };
  }, [enabled, checkNow, schedule, clearTimer]);

  // Возврат в окно после паузы: сеть могла появиться позже запуска.
  useEffect(() => {
    if (!enabled) return undefined;
    const onFocus = () => {
      if (cancelledRef.current || availableRef.current) return;
      if (document.visibilityState === 'hidden') return;
      if (Date.now() - lastAttemptRef.current < FOCUS_RECHECK_AFTER_MS) return;
      attemptsRef.current = 0;
      void checkNow();
    };
    window.addEventListener('focus', onFocus);
    document.addEventListener('visibilitychange', onFocus);
    return () => {
      window.removeEventListener('focus', onFocus);
      document.removeEventListener('visibilitychange', onFocus);
    };
  }, [enabled, checkNow]);

  return {
    info,
    checking,
    error,
    dismissed: Boolean(info?.latestVersion && info.latestVersion === dismissedVersion),
    dismiss,
    checkNow,
  };
}
