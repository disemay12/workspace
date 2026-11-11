/**
 * Пресеты оформления презентации: цвета и шрифты темы.
 *
 * Цвета пишутся в `a:clrScheme` (dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink),
 * шрифты — в `a:fontScheme` (major/minor). Это ровно те части, которые меняет
 * PowerPoint при смене «Темы» / «Варианта»: элементы, оформленные цветами схемы,
 * перекрашиваются, а заданные явным `srgbClr` — нет.
 */

export interface ThemeColorSet {
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  folHlink: string;
}

export type ThemeColorSlot = keyof ThemeColorSet;

export const THEME_COLOR_SLOTS: { id: ThemeColorSlot; label: string }[] = [
  { id: 'dk1', label: 'Текст 1' },
  { id: 'lt1', label: 'Фон 1' },
  { id: 'dk2', label: 'Текст 2' },
  { id: 'lt2', label: 'Фон 2' },
  { id: 'accent1', label: 'Акцент 1' },
  { id: 'accent2', label: 'Акцент 2' },
  { id: 'accent3', label: 'Акцент 3' },
  { id: 'accent4', label: 'Акцент 4' },
  { id: 'accent5', label: 'Акцент 5' },
  { id: 'accent6', label: 'Акцент 6' },
  { id: 'hlink', label: 'Ссылка' },
  { id: 'folHlink', label: 'Прослеженная ссылка' },
];

export interface ThemeFontPair {
  id: string;
  label: string;
  major: string;
  minor: string;
}

export interface ThemeEffectPreset {
  id: string;
  label: string;
  /** Тень по умолчанию для новых фигур. */
  shadow: boolean;
}

export interface ThemePreset {
  id: string;
  label: string;
  colors: ThemeColorSet;
  fonts: ThemeFontPair;
}

/** Шрифтовые пары: только гарнитуры, которые есть в Windows без установки. */
export const FONT_PRESETS: ThemeFontPair[] = [
  { id: 'calibri', label: 'Calibri / Calibri Light', major: 'Calibri Light', minor: 'Calibri' },
  { id: 'segoe', label: 'Segoe UI', major: 'Segoe UI Semilight', minor: 'Segoe UI' },
  { id: 'arial', label: 'Arial', major: 'Arial', minor: 'Arial' },
  { id: 'georgia', label: 'Georgia / Calibri', major: 'Georgia', minor: 'Calibri' },
  { id: 'times', label: 'Times New Roman', major: 'Times New Roman', minor: 'Times New Roman' },
  { id: 'verdana', label: 'Verdana / Tahoma', major: 'Verdana', minor: 'Tahoma' },
];

export const EFFECT_PRESETS: ThemeEffectPreset[] = [
  { id: 'office', label: 'Office (мягкая тень)', shadow: true },
  { id: 'flat', label: 'Плоские (без тени)', shadow: false },
  { id: 'crisp', label: 'Строгие (тень + контур)', shadow: true },
];

const colors = (
  dk1: string,
  lt1: string,
  dk2: string,
  lt2: string,
  accents: [string, string, string, string, string, string],
  hlink = '0563C1',
  folHlink = '954F72'
): ThemeColorSet => ({
  dk1,
  lt1,
  dk2,
  lt2,
  accent1: accents[0],
  accent2: accents[1],
  accent3: accents[2],
  accent4: accents[3],
  accent5: accents[4],
  accent6: accents[5],
  hlink,
  folHlink,
});

export const THEME_PRESETS: ThemePreset[] = [
  {
    id: 'office',
    label: 'Office',
    colors: colors('000000', 'FFFFFF', '44546A', 'E7E6E6', [
      '4472C4',
      'ED7D31',
      'A5A5A5',
      'FFC000',
      '5B9BD5',
      '70AD47',
    ]),
    fonts: FONT_PRESETS[0],
  },
  {
    id: 'graphite',
    label: 'Графит',
    colors: colors('0F1115', 'FAFAFA', '3A3F47', 'D9DDE3', [
      '4C8BF5',
      'E4572E',
      '8D99AE',
      'F2A65A',
      '2EC4B6',
      '6A994E',
    ]),
    fonts: FONT_PRESETS[1],
  },
  {
    id: 'ocean',
    label: 'Океан',
    colors: colors('04283A', 'F4FBFF', '0B4F6C', 'DDF3FA', [
      '0B6E99',
      '01BAEF',
      '20C997',
      'F9C74F',
      '577590',
      '43AA8B',
    ]),
    fonts: FONT_PRESETS[1],
  },
  {
    id: 'warm',
    label: 'Тепло',
    colors: colors('2B1D14', 'FFF8F1', '6B4A30', 'F4E3D1', [
      'D96459',
      'F2AE72',
      '8C5347',
      'E8B04B',
      'A98467',
      '6C584C',
    ]),
    fonts: FONT_PRESETS[3],
  },
  {
    id: 'forest',
    label: 'Лес',
    colors: colors('0B1E14', 'F3FAF5', '244A38', 'D6EBDD', [
      '2D6A4F',
      '40916C',
      '74C69D',
      'B7E4C7',
      'A3B18A',
      '588157',
    ]),
    fonts: FONT_PRESETS[3],
  },
  {
    id: 'neon',
    label: 'Неон',
    colors: colors('0A0A12', 'F7F7FF', '1B1B2F', 'E4E4F5', [
      '7C3AED',
      '06B6D4',
      'F43F5E',
      '22D3EE',
      'A3E635',
      'F59E0B',
    ]),
    fonts: FONT_PRESETS[1],
  },
  {
    id: 'sunset',
    label: 'Закат',
    colors: colors('2A1024', 'FFF5F7', '6A2C4A', 'FBE0E7', [
      'E05780',
      'F07B3F',
      'FFB627',
      '9B5DE5',
      'B5838D',
      '6D597A',
    ]),
    fonts: FONT_PRESETS[4],
  },
  {
    id: 'mono',
    label: 'Моно',
    colors: colors('111111', 'FFFFFF', '4D4D4D', 'E6E6E6', [
      '1A1A1A',
      '595959',
      '8C8C8C',
      'BFBFBF',
      'D9D9D9',
      'F2F2F2',
    ]),
    fonts: FONT_PRESETS[5],
  },
  {
    id: 'retro',
    label: 'Ретро',
    colors: colors('23203A', 'FDF6EC', '4B3F72', 'F3E8D8', [
      '3D5A80',
      'EE6C4D',
      '98C1D9',
      'E0FBFC',
      'B5838D',
      'C9ADA7',
    ]),
    fonts: FONT_PRESETS[2],
  },
  {
    id: 'campus',
    label: 'Кампус',
    colors: colors('1B2432', 'F6F8FB', '2F3E52', 'E3E9F2', [
      '1F4E79',
      'C00000',
      'FFC000',
      '70AD47',
      '5B9BD5',
      '7F7F7F',
    ]),
    fonts: FONT_PRESETS[2],
  },
];

export function presetById(id: string): ThemePreset {
  return THEME_PRESETS.find((preset) => preset.id === id) ?? THEME_PRESETS[0];
}

export function fontPresetById(id: string): ThemeFontPair {
  return FONT_PRESETS.find((preset) => preset.id === id) ?? FONT_PRESETS[0];
}

export function effectPresetById(id: string): ThemeEffectPreset {
  return EFFECT_PRESETS.find((preset) => preset.id === id) ?? EFFECT_PRESETS[0];
}

/** Нормализация цвета: '#A1B2C3' | 'A1B2C3' | 'a1b2c3' → 'A1B2C3'. */
export function normalizeHexColor(value: string, fallback = '000000'): string {
  const hex = (value || '').replace(/^#/, '').trim();
  if (/^[0-9a-fA-F]{6}$/.test(hex)) return hex.toUpperCase();
  if (/^[0-9a-fA-F]{3}$/.test(hex)) {
    return hex
      .split('')
      .map((ch) => ch + ch)
      .join('')
      .toUpperCase();
  }
  return fallback;
}

/** Тёмный ли цвет: '#FFFFFF' → false. */
export function isDarkHex(value: string): boolean {
  const hex = normalizeHexColor(value, 'FFFFFF');
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b < 140;
}
