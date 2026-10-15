import { useCallback, useEffect, useState } from 'react';

export type ThemeMode = 'dark' | 'light';

const STORAGE_KEY = 'workspace_theme';

/**
 * Управление темой оформления.
 * Класс .dark на <html> управляет CSS-переменными (см. src/index.css).
 */
export function useTheme() {
  const [theme, setTheme] = useState<ThemeMode>(() => {
    const saved = window.localStorage.getItem(STORAGE_KEY);
    if (saved === 'light' || saved === 'dark') return saved;
    return window.matchMedia?.('(prefers-color-scheme: light)').matches ? 'light' : 'dark';
  });

  useEffect(() => {
    const root = document.documentElement;
    root.classList.toggle('dark', theme === 'dark');
    root.style.colorScheme = theme;
    try {
      window.localStorage.setItem(STORAGE_KEY, theme);
    } catch {
      // игнорируем недоступность localStorage
    }
  }, [theme]);

  const toggleTheme = useCallback(() => {
    setTheme((prev) => (prev === 'dark' ? 'light' : 'dark'));
  }, []);

  return { theme, setTheme, toggleTheme, isDark: theme === 'dark' };
}
