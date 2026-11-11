/**
 * Раскладка окна: нормализация сохранённого состояния и отпечаток изменений.
 *
 * В workspace-session.json может лежать что угодно (старая версия, ручная
 * правка, другой билд), поэтому читаем защитно: только известные поля и типы.
 */
export interface LayoutState {
  aiPanelOpen: boolean;
  terminalOpen: boolean;
  searchPanelOpen: boolean;
  openPaths: string[];
  activePath: string;
}

export const DEFAULT_LAYOUT: LayoutState = {
  aiPanelOpen: true,
  terminalOpen: false,
  searchPanelOpen: false,
  openPaths: [],
  activePath: '',
};

export const MAX_RESTORED_TABS = 12;

function bool(value: unknown, fallback: boolean): boolean {
  return typeof value === 'boolean' ? value : fallback;
}

/** Приводит произвольный объект к LayoutState, отбрасывая мусор. */
export function normalizeLayout(raw: unknown): LayoutState {
  if (!raw || typeof raw !== 'object') return { ...DEFAULT_LAYOUT };
  const source = raw as Record<string, unknown>;
  const paths = Array.isArray(source.openPaths)
    ? source.openPaths
        .filter((path): path is string => typeof path === 'string' && path.trim().length > 0)
        .slice(0, MAX_RESTORED_TABS)
    : [];
  const active = typeof source.activePath === 'string' ? source.activePath : '';
  return {
    aiPanelOpen: bool(source.aiPanelOpen, DEFAULT_LAYOUT.aiPanelOpen),
    terminalOpen: bool(source.terminalOpen, DEFAULT_LAYOUT.terminalOpen),
    searchPanelOpen: bool(source.searchPanelOpen, DEFAULT_LAYOUT.searchPanelOpen),
    openPaths: [...new Set(paths)],
    activePath: active,
  };
}

/** Отпечаток раскладки: по нему решаем, нужно ли писать в файл сессии. */
export function layoutFingerprint(layout: LayoutState): string {
  return JSON.stringify([
    layout.aiPanelOpen ? 1 : 0,
    layout.terminalOpen ? 1 : 0,
    layout.searchPanelOpen ? 1 : 0,
    layout.openPaths,
    layout.activePath,
  ]);
}
