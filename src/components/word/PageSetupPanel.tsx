/**
 * Параметры листа Word: бумага, ориентация, поля, колонтитулы.
 *
 * Значения уходят в существующий `applyPageSetup`: он меняет sectPr пакета,
 * поэтому документ открывается в Word с теми же полями и бумагой.
 */
import React from 'react';
import type { PageSetup } from '../../types';
import { PAPER_SIZES, type PageGeometry, type PaperName } from '../../lib/units';

interface Props {
  page: PageSetup;
  geometry: PageGeometry;
  pages: number;
  onChange: (patch: Partial<PageSetup>) => void;
  onHeaderFooterText: (kind: 'header' | 'footer', text: string) => void;
  headerText: string;
  footerText: string;
}

const PAPER_OPTIONS: Array<{ id: PaperName; label: string }> = [
  { id: 'A4', label: 'A4 · 210 × 297 мм' },
  { id: 'A3', label: 'A3 · 297 × 420 мм' },
  { id: 'A5', label: 'A5 · 148 × 210 мм' },
  { id: 'Letter', label: 'Letter · 216 × 279 мм' },
  { id: 'Legal', label: 'Legal · 216 × 356 мм' },
];

const MARGIN_PRESETS: Array<{ id: string; label: string; margins: [number, number, number, number] }> = [
  { id: 'normal', label: 'Обычные', margins: [20, 15, 20, 30] },
  { id: 'narrow', label: 'Узкие', margins: [12.7, 12.7, 12.7, 12.7] },
  { id: 'wide', label: 'Широкие', margins: [25.4, 25.4, 25.4, 25.4] },
  { id: 'reference', label: 'По ГОСТ', margins: [20, 15, 20, 30] },
];

const marginKeys = [
  ['marginTopMm', 'Верхнее'],
  ['marginBottomMm', 'Нижнее'],
  ['marginLeftMm', 'Левое'],
  ['marginRightMm', 'Правое'],
] as const;

const inputClass =
  'w-full rounded border border-ws bg-ws-elevated px-1.5 py-1 text-[11px] text-ws focus:border-ws-accent focus:outline-none';

export const PageSetupPanel: React.FC<Props> = ({
  page,
  geometry,
  pages,
  onChange,
  onHeaderFooterText,
  headerText,
  footerText,
}) => {
  const paperId: PaperName =
    geometry.paper === 'custom' ? 'custom' : geometry.paper;
  const landscape = !geometry.portrait;

  const applyPaper = (name: PaperName) => {
    if (name === 'custom') return;
    const size = PAPER_SIZES[name];
    const width = landscape ? size.heightMm : size.widthMm;
    const height = landscape ? size.widthMm : size.heightMm;
    onChange({ widthMm: Math.round(width * 10) / 10, heightMm: Math.round(height * 10) / 10 });
  };

  const setOrientation = (portrait: boolean) => {
    if (portrait === geometry.portrait) return;
    onChange({ widthMm: geometry.heightMm, heightMm: geometry.widthMm });
  };

  const numberValue = (
    key: 'marginTopMm' | 'marginRightMm' | 'marginBottomMm' | 'marginLeftMm' | 'headerMm' | 'footerMm',
    fallback: number
  ): number => {
    const value = page[key];
    return typeof value === 'number' && value >= 0
      ? Math.round(value * 10) / 10
      : Math.round(fallback * 10) / 10;
  };

  return (
    <div className="space-y-2 p-2 text-[11px] text-ws" data-page-setup="1">
      <div>
        <label className="mb-0.5 block text-ws-muted">Бумага</label>
        <select
          value={paperId}
          onChange={(event) => applyPaper(event.target.value as PaperName)}
          className={inputClass}
        >
          {PAPER_OPTIONS.map((option) => (
            <option key={option.id} value={option.id}>
              {option.label}
            </option>
          ))}
          <option value="custom">
            Своя · {Math.round(geometry.widthMm)} × {Math.round(geometry.heightMm)} мм
          </option>
        </select>
      </div>

      <div className="flex items-center gap-1">
        <button
          type="button"
          onClick={() => setOrientation(true)}
          className={`flex-1 rounded border px-2 py-1 ${
            geometry.portrait ? 'border-ws-accent bg-ws-accent-soft text-ws-accent' : 'border-ws'
          }`}
        >
          Книжная
        </button>
        <button
          type="button"
          onClick={() => setOrientation(false)}
          className={`flex-1 rounded border px-2 py-1 ${
            !geometry.portrait ? 'border-ws-accent bg-ws-accent-soft text-ws-accent' : 'border-ws'
          }`}
        >
          Альбомная
        </button>
      </div>

      <div>
        <div className="mb-1 flex items-center justify-between text-ws-muted">
          <span>Поля, мм</span>
          <select
            value=""
            onChange={(event) => {
              const preset = MARGIN_PRESETS.find((item) => item.id === event.target.value);
              if (!preset) return;
              const [top, right, bottom, left] = preset.margins;
              onChange({
                marginTopMm: top,
                marginRightMm: right,
                marginBottomMm: bottom,
                marginLeftMm: left,
              });
            }}
            className="rounded border border-ws bg-ws-elevated px-1 py-0.5 text-[10px] text-ws"
          >
            <option value="">Пресет…</option>
            {MARGIN_PRESETS.map((preset) => (
              <option key={preset.id} value={preset.id}>
                {preset.label}
              </option>
            ))}
          </select>
        </div>
        <div className="grid grid-cols-2 gap-1">
          {marginKeys.map(([key, label]) => (
            <label key={key} className="flex items-center gap-1">
              <span className="w-[54px] shrink-0 text-ws-faint">{label}</span>
              <input
                type="number"
                min={0}
                max={100}
                step={0.5}
                value={numberValue(key, key === 'marginRightMm' ? 15 : key === 'marginLeftMm' ? 30 : 20)}
                onChange={(event) => {
                  const value = Number(event.target.value.replace(',', '.'));
                  if (!Number.isFinite(value) || value < 0 || value > 100) return;
                  onChange({ [key]: value } as Partial<PageSetup>);
                }}
                className={inputClass}
              />
            </label>
          ))}
        </div>
      </div>

      <div>
        <div className="mb-1 text-ws-muted">Колонтитулы, мм от края</div>
        <div className="grid grid-cols-2 gap-1">
          <label className="flex items-center gap-1">
            <span className="w-[54px] shrink-0 text-ws-faint">Верхний</span>
            <input
              type="number"
              min={0}
              max={60}
              step={0.5}
              value={numberValue('headerMm', 12.5)}
              onChange={(event) => {
                const value = Number(event.target.value.replace(',', '.'));
                if (!Number.isFinite(value) || value < 0 || value > 60) return;
                onChange({ headerMm: value });
              }}
              className={inputClass}
            />
          </label>
          <label className="flex items-center gap-1">
            <span className="w-[54px] shrink-0 text-ws-faint">Нижний</span>
            <input
              type="number"
              min={0}
              max={60}
              step={0.5}
              value={numberValue('footerMm', 12.5)}
              onChange={(event) => {
                const value = Number(event.target.value.replace(',', '.'));
                if (!Number.isFinite(value) || value < 0 || value > 60) return;
                onChange({ footerMm: value });
              }}
              className={inputClass}
            />
          </label>
        </div>
        <label className="mt-1 block text-ws-muted">Текст верхнего колонтитула</label>
        <input
          value={headerText}
          onChange={(event) => onHeaderFooterText('header', event.target.value)}
          placeholder="например: Отчёт · {PAGE}"
          className={inputClass}
        />
        <label className="mt-1 block text-ws-muted">Текст нижнего колонтитула</label>
        <input
          value={footerText}
          onChange={(event) => onHeaderFooterText('footer', event.target.value)}
          placeholder="например: страница {PAGE} из {NUMPAGES}"
          className={inputClass}
        />
      </div>

      <div className="border-t border-ws pt-1 text-[10px] text-ws-faint">
        Полоса набора: {Math.round(geometry.contentWidthPx)} × {Math.round(geometry.contentHeightPx)} px ·
        листов: {pages}
      </div>
    </div>
  );
};
