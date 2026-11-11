/**
 * Оформление интерфейса: тема, акцент, плотность, масштаб.
 *
 * Все расчёты — чистые функции, чтобы их можно было проверить в model-check;
 * `applyAppearance` пишет результат в `<html>` (класс `dark` и CSS-переменные).
 */
import type { AppearancePreferences } from '../state/preferences';

export const UI_SCALE_MIN = 0.85;
export const UI_SCALE_MAX = 1.3;

/** Итоговая тема с учётом системной. */
export function resolveDark(theme: AppearancePreferences['theme'], systemDark: boolean): boolean {
  if (theme === 'dark') return true;
  if (theme === 'light') return false;
  return systemDark;
}

export function clampScale(value: number): number {
  const scale = Number(value);
  if (!Number.isFinite(scale)) return 1;
  return Math.min(UI_SCALE_MAX, Math.max(UI_SCALE_MIN, Math.round(scale * 100) / 100));
}

/** Нормализация акцента: #RRGGBB, короткий #RGB или пустая строка. */
export function normalizeAccent(value: string): string {
  const raw = (value || '').trim().replace(/^#/, '');
  if (!raw) return '';
  if (/^[0-9a-fA-F]{6}$/.test(raw)) return `#${raw.toUpperCase()}`;
  if (/^[0-9a-fA-F]{3}$/.test(raw)) {
    return `#${raw
      .split('')
      .map((char) => char + char)
      .join('')
      .toUpperCase()}`;
  }
  return '';
}

/** Акцент с прозрачностью для мягких подложек (--ws-accent-soft). */
export function hexToRgba(hex: string, alpha: number): string {
  const normalized = normalizeAccent(hex);
  if (!normalized) return '';
  const value = normalized.slice(1);
  const red = parseInt(value.slice(0, 2), 16);
  const green = parseInt(value.slice(2, 4), 16);
  const blue = parseInt(value.slice(4, 6), 16);
  const safeAlpha = Math.min(1, Math.max(0, Number(alpha) || 0));
  return `rgba(${red}, ${green}, ${blue}, ${safeAlpha})`;
}

/**
 * Применяет оформление к документу. Возвращает итоговое состояние — по нему
 * UI понимает, тёмная ли сейчас тема.
 */
export function applyAppearance(
  appearance: AppearancePreferences,
  systemDark: boolean,
  root: HTMLElement = document.documentElement
): { dark: boolean; scale: number } {
  const dark = resolveDark(appearance.theme, systemDark);
  const scale = clampScale(appearance.uiScale);
  root.classList.toggle('dark', dark);
  root.style.colorScheme = dark ? 'dark' : 'light';
  root.dataset.density = appearance.density;
  root.dataset.motion = appearance.reduceMotion ? 'reduced' : 'full';
  root.style.setProperty('--ws-ui-scale', String(scale));
  const accent = normalizeAccent(appearance.accent);
  if (accent) {
    root.style.setProperty('--ws-accent', accent);
    root.style.setProperty('--ws-accent-soft', hexToRgba(accent, dark ? 0.22 : 0.14));
  } else {
    root.style.removeProperty('--ws-accent');
    root.style.removeProperty('--ws-accent-soft');
  }
  return { dark, scale };
}
