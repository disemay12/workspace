import React, { useMemo, useState } from 'react';
import { Check, Palette, Type, Sparkles, X } from 'lucide-react';
import { cn } from '../../lib/utils';
import { Button } from '../ui';
import {
  EFFECT_PRESETS,
  FONT_PRESETS,
  THEME_COLOR_SLOTS,
  THEME_PRESETS,
  isDarkHex,
  normalizeHexColor,
  type ThemeColorSet,
  type ThemeColorSlot,
  type ThemeFontPair,
} from '../../lib/pptxThemes';
import type { PptxThemeModel } from '../../lib/pptxDoc';

/**
 * Панель темы: пресеты «цвета + шрифты», ручная правка 12 слотов схемы
 * и выбор шрифтовой пары. Применение меняет `a:clrScheme`/`a:fontScheme`
 * в `ppt/theme/themeN.xml` — так же, как «Дизайн → Тема» в PowerPoint.
 */

export interface ThemePanelProps {
  open: boolean;
  theme: PptxThemeModel;
  onClose: () => void;
  onApplyPreset: (presetId: string, label: string) => void;
  onApplyColors: (colors: Partial<ThemeColorSet>, label: string) => void;
  onApplyFonts: (fonts: ThemeFontPair) => void;
  onApplyEffects: (effectId: string) => void;
  shadowDefault: boolean;
}

type Tab = 'presets' | 'colors' | 'fonts';

export const ThemePanel: React.FC<ThemePanelProps> = ({
  open,
  theme,
  onClose,
  onApplyPreset,
  onApplyColors,
  onApplyFonts,
  onApplyEffects,
  shadowDefault,
}) => {
  const [tab, setTab] = useState<Tab>('presets');
  const [draft, setDraft] = useState<Partial<ThemeColorSet>>({});

  const colors = useMemo<ThemeColorSet>(() => ({ ...theme.colors, ...draft }), [theme.colors, draft]);
  const dirty = Object.keys(draft).length > 0;

  if (!open) return null;

  return (
    <div className="fixed inset-0 z-[80] flex items-center justify-center bg-black/50 p-6" onClick={onClose}>
      <div
        className="flex max-h-[86vh] w-full max-w-3xl flex-col overflow-hidden rounded-xl border border-ws bg-ws-surface shadow-xl"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="flex items-center justify-between border-b border-ws px-4 py-2.5">
          <div className="flex items-center gap-2 text-sm font-medium text-ws">
            <Palette size={15} className="text-ws-accent" />
            Тема оформления
            <span className="text-[11px] font-normal text-ws-faint">{theme.name}</span>
          </div>
          <button className="rounded p-1 text-ws-muted hover:bg-ws-hover hover:text-ws" onClick={onClose}>
            <X size={15} />
          </button>
        </div>

        <div className="flex shrink-0 gap-1 border-b border-ws px-3 py-1.5 text-[11px]">
          {(
            [
              ['presets', 'Пресеты', <Sparkles key="p" size={12} />],
              ['colors', 'Цвета', <Palette key="c" size={12} />],
              ['fonts', 'Шрифты', <Type key="f" size={12} />],
            ] as [Tab, string, React.ReactNode][]
          ).map(([id, label, icon]) => (
            <button
              key={id}
              onClick={() => setTab(id)}
              className={cn(
                'flex items-center gap-1.5 rounded px-2.5 py-1',
                tab === id ? 'bg-ws-accent-soft text-ws-accent' : 'text-ws-muted hover:bg-ws-hover'
              )}
            >
              {icon}
              {label}
            </button>
          ))}
        </div>

        <div className="min-h-0 flex-1 overflow-y-auto p-4">
          {tab === 'presets' && (
            <div className="grid gap-2 sm:grid-cols-2">
              {THEME_PRESETS.map((preset) => {
                const active = preset.label === theme.name;
                return (
                  <button
                    key={preset.id}
                    onClick={() => onApplyPreset(preset.id, preset.label)}
                    className={cn(
                      'flex items-center gap-3 rounded-lg border p-2.5 text-left transition-colors',
                      active
                        ? 'border-ws-accent bg-ws-accent-soft/40'
                        : 'border-ws bg-ws-elevated hover:border-ws-strong'
                    )}
                  >
                    <div
                      className="flex h-11 w-14 shrink-0 items-center justify-center rounded border border-black/10"
                      style={{ background: `#${preset.colors.lt1}` }}
                    >
                      <span
                        className="text-[15px] font-semibold leading-none"
                        style={{
                          color: `#${preset.colors.dk1}`,
                          fontFamily: `"${preset.fonts.major}", "Segoe UI", sans-serif`,
                        }}
                      >
                        Aa
                      </span>
                    </div>
                    <div className="min-w-0">
                      <div className="flex items-center gap-1 text-[12px] font-medium text-ws">
                        {preset.label}
                        {active && <Check size={12} className="text-ws-accent" />}
                      </div>
                      <div className="mt-1 flex gap-1">
                        {(
                          ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const
                        ).map((slot) => (
                          <span
                            key={slot}
                            className="h-2.5 w-2.5 rounded-full border border-black/10"
                            style={{ background: `#${preset.colors[slot]}` }}
                          />
                        ))}
                      </div>
                      <div className="mt-0.5 truncate text-[10px] text-ws-faint">{preset.fonts.label}</div>
                    </div>
                  </button>
                );
              })}
            </div>
          )}

          {tab === 'colors' && (
            <div className="space-y-3">
              <p className="text-[11px] leading-relaxed text-ws-muted">
                Меняются только элементы, оформленные цветами схемы. У фигур с явно заданным цветом
                заливка и текст останутся прежними — как в PowerPoint.
              </p>
              <div className="grid gap-2 sm:grid-cols-2">
                {THEME_COLOR_SLOTS.map(({ id, label }) => (
                  <label
                    key={id}
                    className="flex items-center gap-2 rounded border border-ws bg-ws-elevated px-2 py-1.5"
                  >
                    <input
                      type="color"
                      value={`#${normalizeHexColor(colors[id], 'FFFFFF').toLowerCase()}`}
                      onChange={(event) =>
                        setDraft((prev) => ({ ...prev, [id as ThemeColorSlot]: normalizeHexColor(event.target.value) }))
                      }
                      className="h-6 w-8 cursor-pointer rounded border border-ws bg-transparent"
                    />
                    <span className="flex-1 text-[11px] text-ws">{label}</span>
                    <span className="font-mono text-[10px] text-ws-faint">
                      #{normalizeHexColor(colors[id], 'FFFFFF')}
                    </span>
                  </label>
                ))}
              </div>
              <div className="flex flex-wrap gap-2">
                <Button
                  variant="primary"
                  disabled={!dirty}
                  onClick={() => {
                    onApplyColors(draft, 'Свои цвета');
                    setDraft({});
                  }}
                >
                  Применить цвета
                </Button>
                <Button variant="secondary" disabled={!dirty} onClick={() => setDraft({})}>
                  Сбросить
                </Button>
              </div>
            </div>
          )}

          {tab === 'fonts' && (
            <div className="space-y-3">
              <div className="grid gap-2 sm:grid-cols-2">
                {FONT_PRESETS.map((pair) => {
                  const active = pair.label === theme.fonts.label;
                  return (
                    <button
                      key={pair.id}
                      onClick={() => onApplyFonts(pair)}
                      className={cn(
                        'rounded-lg border p-2.5 text-left transition-colors',
                        active
                          ? 'border-ws-accent bg-ws-accent-soft/40'
                          : 'border-ws bg-ws-elevated hover:border-ws-strong'
                      )}
                    >
                      <div className="flex items-center justify-between text-[11px] font-medium text-ws">
                        {pair.label}
                        {active && <Check size={12} className="text-ws-accent" />}
                      </div>
                      <div
                        className="mt-1 text-[18px] leading-tight text-ws"
                        style={{ fontFamily: `"${pair.major}", "Segoe UI", sans-serif` }}
                      >
                        Заголовок
                      </div>
                      <div
                        className="text-[11px] text-ws-muted"
                        style={{ fontFamily: `"${pair.minor}", "Segoe UI", sans-serif` }}
                      >
                        Основной текст слайда
                      </div>
                    </button>
                  );
                })}
              </div>

              <div className="rounded-lg border border-ws bg-ws-elevated p-3">
                <div className="mb-1.5 text-[11px] font-medium text-ws">Эффекты для новых фигур</div>
                <div className="flex flex-wrap gap-1.5">
                  {EFFECT_PRESETS.map((effect) => {
                    const active = effect.shadow === shadowDefault;
                    return (
                      <button
                        key={effect.id}
                        onClick={() => onApplyEffects(effect.id)}
                        className={cn(
                          'rounded border px-2 py-1 text-[11px]',
                          active
                            ? 'border-ws-accent bg-ws-accent-soft text-ws-accent'
                            : 'border-ws text-ws-muted hover:bg-ws-hover'
                        )}
                      >
                        {effect.label}
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
          )}
        </div>

        <div className="flex shrink-0 items-center justify-between border-t border-ws px-4 py-2 text-[10px] text-ws-faint">
          <span>
            Схема: {isDarkHex(colors.lt1) ? 'тёмная' : 'светлая'} · шрифты {theme.fonts.major} /{' '}
            {theme.fonts.minor}
          </span>
          <Button variant="ghost" onClick={onClose}>
            Готово
          </Button>
        </div>
      </div>
    </div>
  );
};
