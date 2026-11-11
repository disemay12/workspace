/**
 * Правила уведомлений: какие события показывать тостом и когда молчать.
 *
 * Отдельные чистые функции — их проверяет model-check (включая переходы
 * через полночь и режим тишины).
 */
import type { NotificationPreferences } from '../state/preferences';

export type NotificationKind = keyof NotificationPreferences['events'];

export interface NotificationInput {
  kind: NotificationKind;
  prefs: NotificationPreferences;
  now: Date;
}

/** «ЧЧ:ММ» → минуты от начала суток или null. */
export function parseClock(value: string): number | null {
  const match = /^(\d{1,2}):(\d{2})$/.exec((value || '').trim());
  if (!match) return null;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (hours > 23 || minutes > 59) return null;
  return hours * 60 + minutes;
}

/** Попадает ли время в окно тишины (в т.ч. окно через полночь). */
export function inQuietHours(from: string, to: string, now: Date): boolean {
  const start = parseClock(from);
  const end = parseClock(to);
  if (start === null || end === null || start === end) return false;
  const current = now.getHours() * 60 + now.getMinutes();
  return start < end ? current >= start && current < end : current >= start || current < end;
}

export function shouldNotify({ kind, prefs, now }: NotificationInput): boolean {
  // «Без звука» — это 'off', а не пустая строка; иначе проверка врала.
  if (!prefs.toasts && prefs.sound === 'off') return false;
  if (!prefs.events[kind]) return false;
  return !inQuietHours(prefs.quietFrom, prefs.quietTo, now);
}

/** Нужен ли звук для события: тихий сигнал не играет для ошибок. */
export function shouldPlaySound(prefs: NotificationPreferences, kind: NotificationKind): boolean {
  if (prefs.sound === 'off' || !prefs.sound) return false;
  if (!prefs.events[kind]) return false;
  return kind !== 'error';
}
