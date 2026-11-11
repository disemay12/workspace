/**
 * Пользовательские предпочтения интерфейса (не AI-настройки).
 *
 * AI-ключи и модели живут в `state/settings.ts` (шифрованное хранилище),
 * а здесь — то, что можно держать в localStorage: поведение больших таблиц,
 * режим Япикса, зум видео, подсказки гайдов и раскрытие дерева из Проводника.
 */
import { useCallback, useEffect, useState } from 'react';

export type LargeSheetMode = 'auto' | 'always' | 'never';

export interface LargeSheetPreferences {
  /** auto — по порогам ниже, always — всегда оконный режим, never — как раньше. */
  mode: LargeSheetMode;
  /** Порог размера файла в мегабайтах для авто-режима. */
  thresholdMb: number;
  /** Порог количества строк (если известно заранее) для авто-режима. */
  thresholdRows: number;
}

/** Вид проводника, сохраняется между запусками. */
export interface ExplorerPreferences {
  sortBy: 'name' | 'mtime' | 'size' | 'type';
  dir: 'asc' | 'desc';
  foldersFirst: boolean;
  showHidden: boolean;
  /** Показывать размер и дату справа от имени. */
  showMeta: boolean;
  /** Миниатюры изображений вместо значка типа. */
  showThumbs: boolean;
}


/** Оформление интерфейса: тема, акцент, плотность, масштаб. */
export interface AppearancePreferences {
  theme: 'system' | 'dark' | 'light';
  /** HEX без #; пустая строка — акцент по умолчанию. */
  accent: string;
  density: 'comfortable' | 'compact';
  /** Масштаб интерфейса, 0.85…1.3. */
  uiScale: number;
  reduceMotion: boolean;
}

/** Уведомления внутри приложения: тосты и тихий сигнал. */
export interface NotificationPreferences {
  toasts: boolean;
  sound: 'off' | 'knock' | 'soft';
  /** 0…1. */
  volume: number;
  /** Типы событий, о которых сообщать. */
  events: Record<'save' | 'export' | 'ai' | 'update' | 'error', boolean>;
  /** Режим тишины, локальное время «ЧЧ:ММ», пустые строки — выключено. */
  quietFrom: string;
  quietTo: string;
}
export interface Preferences {
  largeSheet: LargeSheetPreferences;
  /** Пользователь вошёл в аккаунт Япикса — лимиты выше (20 в альбоме вместо 10). */
  yapxAccountMode: boolean;
  /** Начальный зум превью видео. */
  videoZoomDefault: number;
  /** При открытии файла из контекстного меню раскрывать дерево до него. */
  explorerRevealOnOpen: boolean;
  /** Вид проводника: сортировка и скрытые файлы. */
  explorer: ExplorerPreferences;
  /** Пройденные мини-гайды: id → true. */
  guidesSeen: Record<string, boolean>;
  appearance: AppearancePreferences;
  notifications: NotificationPreferences;
}

export const DEFAULT_PREFERENCES: Preferences = {
  largeSheet: { mode: 'auto', thresholdMb: 8, thresholdRows: 50000 },
  yapxAccountMode: false,
  videoZoomDefault: 1,
  explorerRevealOnOpen: true,
  explorer: { sortBy: 'name', dir: 'asc', foldersFirst: true, showHidden: false, showMeta: true, showThumbs: true },
  guidesSeen: {},
  appearance: {
    theme: 'system',
    accent: '',
    density: 'comfortable',
    uiScale: 1,
    reduceMotion: false,
  },
  notifications: {
    toasts: true,
    sound: 'knock',
    volume: 0.25,
    events: { save: true, export: true, ai: true, update: true, error: true },
    quietFrom: '',
    quietTo: '',
  },
};

const STORAGE_KEY = 'workspace_preferences';

/** Слияние с умолчаниями: старые записи не должны терять новые поля. */
function merge(input: unknown): Preferences {
  const stored = input && typeof input === 'object' ? (input as Partial<Preferences>) : {};
  const largeSheet = stored.largeSheet && typeof stored.largeSheet === 'object'
    ? stored.largeSheet
    : ({} as Partial<LargeSheetPreferences>);
  const mode: LargeSheetMode =
    largeSheet.mode === 'always' || largeSheet.mode === 'never' || largeSheet.mode === 'auto'
      ? largeSheet.mode
      : DEFAULT_PREFERENCES.largeSheet.mode;
  return {
    largeSheet: {
      mode,
      thresholdMb:
        typeof largeSheet.thresholdMb === 'number' && largeSheet.thresholdMb > 0
          ? largeSheet.thresholdMb
          : DEFAULT_PREFERENCES.largeSheet.thresholdMb,
      thresholdRows:
        typeof largeSheet.thresholdRows === 'number' && largeSheet.thresholdRows > 0
          ? largeSheet.thresholdRows
          : DEFAULT_PREFERENCES.largeSheet.thresholdRows,
    },
    yapxAccountMode: Boolean(stored.yapxAccountMode),
    videoZoomDefault:
      typeof stored.videoZoomDefault === 'number' && stored.videoZoomDefault > 0
        ? stored.videoZoomDefault
        : DEFAULT_PREFERENCES.videoZoomDefault,
    explorerRevealOnOpen:
      stored.explorerRevealOnOpen === undefined
        ? DEFAULT_PREFERENCES.explorerRevealOnOpen
        : Boolean(stored.explorerRevealOnOpen),
    explorer: {
      sortBy:
        stored.explorer?.sortBy === 'mtime' ||
        stored.explorer?.sortBy === 'size' ||
        stored.explorer?.sortBy === 'type'
          ? stored.explorer.sortBy
          : DEFAULT_PREFERENCES.explorer.sortBy,
      dir: stored.explorer?.dir === 'desc' ? 'desc' : 'asc',
      foldersFirst:
        stored.explorer?.foldersFirst === undefined
          ? DEFAULT_PREFERENCES.explorer.foldersFirst
          : Boolean(stored.explorer.foldersFirst),
      showHidden: Boolean(stored.explorer?.showHidden),
      showMeta:
        stored.explorer?.showMeta === undefined
          ? DEFAULT_PREFERENCES.explorer.showMeta
          : Boolean(stored.explorer.showMeta),
      showThumbs:
        stored.explorer?.showThumbs === undefined
          ? DEFAULT_PREFERENCES.explorer.showThumbs
          : Boolean(stored.explorer.showThumbs),
    },
    guidesSeen:
      stored.guidesSeen && typeof stored.guidesSeen === 'object'
        ? { ...(stored.guidesSeen as Record<string, boolean>) }
        : {},
    appearance: { ...DEFAULT_PREFERENCES.appearance, ...(stored.appearance ?? {}) },
    notifications: {
      ...DEFAULT_PREFERENCES.notifications,
      ...(stored.notifications ?? {}),
      events: { ...DEFAULT_PREFERENCES.notifications.events, ...(stored.notifications?.events ?? {}) },
    },
  };
}

/** Чтение предпочтений вне React (сервисы, обработчики событий). */
export function readPreferences(): Preferences {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_PREFERENCES };
    return merge(JSON.parse(raw));
  } catch {
    return { ...DEFAULT_PREFERENCES };
  }
}

export function writePreferences(patch: Partial<Preferences>): Preferences {
  const next = merge({ ...readPreferences(), ...patch });
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  } catch {
    // Quota или приватный режим — настройка просто не сохранится
  }
  return next;
}

/**
 * Решение об оконном режиме для таблицы.
 * Размер в байтах и (если известно) число строк вместе с пользовательским режимом.
 */
export function shouldUseWindowedSheet(
  preferences: Preferences,
  sizeBytes: number,
  rows?: number
): boolean {
  const { mode, thresholdMb, thresholdRows } = preferences.largeSheet;
  if (mode === 'always') return true;
  if (mode === 'never') return false;
  if (sizeBytes > thresholdMb * 1024 * 1024) return true;
  return typeof rows === 'number' && rows > thresholdRows;
}

/** Хук предпочтений: одно хранилище, обновление патчем. */
export function usePreferences(): [Preferences, (patch: Partial<Preferences>) => void] {
  const [preferences, setPreferences] = useState<Preferences>(readPreferences);

  // Другое окно могло изменить значения — подхватываем.
  useEffect(() => {
    const onStorage = (event: StorageEvent) => {
      if (event.key && event.key !== STORAGE_KEY) return;
      setPreferences(readPreferences());
    };
    window.addEventListener('storage', onStorage);
    return () => window.removeEventListener('storage', onStorage);
  }, []);

  const update = useCallback((patch: Partial<Preferences>) => {
    setPreferences(writePreferences(patch));
  }, []);

  return [preferences, update];
}
