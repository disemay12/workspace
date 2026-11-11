import React from 'react';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  ArrowDown,
  ArrowUp,
  Bold,
  Copy,
  Italic,
  List,
  ListOrdered,
  Strikethrough,
  Trash2,
  Underline,
} from 'lucide-react';
import { cn } from '../../lib/utils';
import { normalizeHexColor } from '../../lib/pptxThemes';
import { FONT_PRESETS } from '../../lib/pptxThemes';
import type {
  PptxDocModel,
  PptxParagraph,
  PptxShapeModel,
  PptxSlideModel,
  PptxTextBody,
  PptxTextRun,
  PptxTransitionKind,
  PptxTransitionModel,
} from '../../lib/pptxDoc';

const EMU_PER_CM = 360000;

export interface TextFormatPatch {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  size?: number;
  color?: string;
  family?: string;
}

export interface InspectorProps {
  shape: PptxShapeModel | null;
  slide: PptxSlideModel;
  model: PptxDocModel;
  onGeometry: (patch: { x?: number; y?: number; w?: number; h?: number; rotation?: number }) => void;
  onFill: (color: string | null) => void;
  onOutline: (outline: { color?: string; width?: number } | null) => void;
  onShadow: (enabled: boolean) => void;
  onTextFormat: (patch: TextFormatPatch) => void;
  onParagraph: (patch: Partial<PptxParagraph>) => void;
  onTextBody: (patch: Partial<Pick<PptxTextBody, 'verticalAnchor' | 'autofit'>>) => void;
  onZOrder: (mode: 'front' | 'back' | 'forward' | 'backward') => void;
  onDeleteShape: () => void;
  onDuplicateShape: () => void;
  onBackground: (color: string, all: boolean) => void;
  onTransition: (transition: PptxTransitionModel | undefined) => void;
}

const TRANSITIONS: { id: PptxTransitionKind; label: string }[] = [
  { id: 'none', label: 'Нет' },
  { id: 'fade', label: 'Затемнение' },
  { id: 'push', label: 'Сдвиг' },
  { id: 'wipe', label: 'Стирание' },
  { id: 'cut', label: 'Резкая смена' },
];

export const Inspector: React.FC<InspectorProps> = ({
  shape,
  slide,
  model,
  onGeometry,
  onFill,
  onOutline,
  onShadow,
  onTextFormat,
  onParagraph,
  onTextBody,
  onZOrder,
  onDeleteShape,
  onDuplicateShape,
  onBackground,
  onTransition,
}) => {
  const firstRun: PptxTextRun | undefined = shape?.text?.paragraphs[0]?.runs[0];
  const firstParagraph = shape?.text?.paragraphs[0];
  const cm = (emu: number) => Math.round((emu / EMU_PER_CM) * 10) / 10;
  const fromCm = (value: string) => Math.round((Number(value) || 0) * EMU_PER_CM);

  return (
    <aside className="flex w-64 shrink-0 flex-col overflow-y-auto border-l border-ws bg-ws-surface text-[11px]">
      {shape ? (
        <>
          <Section title="Фигура">
            <div className="truncate rounded bg-ws-elevated px-2 py-1 text-[10px] text-ws-muted" title={shape.name}>
              {shape.name || shape.id} · {shape.kind}
            </div>
            <div className="mt-1.5 grid grid-cols-2 gap-1.5">
              {(
                [
                  ['X, см', shape.x, (value: string) => onGeometry({ x: fromCm(value) })],
                  ['Y, см', shape.y, (value: string) => onGeometry({ y: fromCm(value) })],
                  ['Ш, см', shape.w, (value: string) => onGeometry({ w: fromCm(value) })],
                  ['В, см', shape.h, (value: string) => onGeometry({ h: fromCm(value) })],
                ] as [string, number, (value: string) => void][]
              ).map(([label, value, handler]) => (
                <label key={label} className="flex flex-col gap-0.5">
                  <span className="text-[10px] text-ws-faint">{label}</span>
                  <input
                    type="number"
                    step={0.1}
                    value={cm(value)}
                    onChange={(event) => handler(event.target.value)}
                    className="rounded border border-ws bg-ws-editor px-1.5 py-1 text-[11px] text-ws outline-none focus:border-ws-accent"
                  />
                </label>
              ))}
            </div>
            <label className="mt-1.5 flex items-center gap-2">
              <span className="w-14 text-[10px] text-ws-faint">Поворот</span>
              <input
                type="range"
                min={-180}
                max={180}
                step={1}
                value={Math.round(shape.rotation ?? 0)}
                onChange={(event) => onGeometry({ rotation: Number(event.target.value) })}
                className="flex-1 accent-[var(--ws-accent)]"
              />
              <span className="w-8 text-right font-mono text-[10px] text-ws-muted">
                {Math.round(shape.rotation ?? 0)}°
              </span>
            </label>
          </Section>

          <Section title="Оформление">
            <div className="flex items-center gap-1.5">
              <span className="w-14 text-[10px] text-ws-faint">Заливка</span>
              <input
                type="color"
                value={`#${normalizeHexColor(shape.fill ?? model.theme.colors.accent1, 'FFFFFF').toLowerCase()}`}
                onChange={(event) => onFill(normalizeHexColor(event.target.value))}
                className="h-6 w-8 cursor-pointer rounded border border-ws"
              />
              <button
                onClick={() => onFill(null)}
                className="rounded border border-ws px-1.5 py-0.5 text-[10px] text-ws-muted hover:bg-ws-hover"
              >
                Нет
              </button>
              <div className="ml-auto flex gap-1">
                {(
                  ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const
                ).map((slot) => (
                  <button
                    key={slot}
                    title={slot}
                    onClick={() => onFill(model.theme.colors[slot])}
                    className="h-3.5 w-3.5 rounded-full border border-black/20"
                    style={{ background: `#${model.theme.colors[slot]}` }}
                  />
                ))}
              </div>
            </div>

            <div className="mt-1.5 flex items-center gap-1.5">
              <span className="w-14 text-[10px] text-ws-faint">Контур</span>
              <input
                type="color"
                value={`#${normalizeHexColor(shape.outline?.color ?? '000000', '000000').toLowerCase()}`}
                onChange={(event) =>
                  onOutline({ color: normalizeHexColor(event.target.value), width: shape.outline?.width ?? 1 })
                }
                className="h-6 w-8 cursor-pointer rounded border border-ws"
              />
              <input
                type="number"
                min={0}
                max={12}
                step={0.25}
                value={shape.outline?.width ?? 0}
                onChange={(event) =>
                  onOutline(
                    Number(event.target.value) > 0
                      ? { color: shape.outline?.color ?? '000000', width: Number(event.target.value) }
                      : null
                  )
                }
                className="w-14 rounded border border-ws bg-ws-editor px-1 py-0.5 text-[11px] text-ws"
              />
              <button
                onClick={() => onOutline(null)}
                className="rounded border border-ws px-1.5 py-0.5 text-[10px] text-ws-muted hover:bg-ws-hover"
              >
                Нет
              </button>
            </div>

            <label className="mt-1.5 flex items-center gap-2">
              <input type="checkbox" checked={Boolean(shape.shadow)} onChange={(event) => onShadow(event.target.checked)} />
              Тень
            </label>

            <div className="mt-1.5 flex gap-1">
              <IconButton title="На передний план" onClick={() => onZOrder('front')}>
                <ArrowUp size={12} />
              </IconButton>
              <IconButton title="Выше на один" onClick={() => onZOrder('forward')}>
                <ArrowUp size={10} />
              </IconButton>
              <IconButton title="Ниже на один" onClick={() => onZOrder('backward')}>
                <ArrowDown size={10} />
              </IconButton>
              <IconButton title="На задний план" onClick={() => onZOrder('back')}>
                <ArrowDown size={12} />
              </IconButton>
              <IconButton title="Дублировать (Ctrl+D)" onClick={onDuplicateShape}>
                <Copy size={12} />
              </IconButton>
              <IconButton title="Удалить (Del)" onClick={onDeleteShape} danger>
                <Trash2 size={12} />
              </IconButton>
            </div>
          </Section>

          <Section title="Текст">
            <div className="flex flex-wrap gap-1">
              <Toggle active={Boolean(firstRun?.bold)} onClick={() => onTextFormat({ bold: !firstRun?.bold })} title="Полужирный">
                <Bold size={12} />
              </Toggle>
              <Toggle active={Boolean(firstRun?.italic)} onClick={() => onTextFormat({ italic: !firstRun?.italic })} title="Курсив">
                <Italic size={12} />
              </Toggle>
              <Toggle active={Boolean(firstRun?.underline)} onClick={() => onTextFormat({ underline: !firstRun?.underline })} title="Подчёркнутый">
                <Underline size={12} />
              </Toggle>
              <Toggle active={Boolean(firstRun?.strike)} onClick={() => onTextFormat({ strike: !firstRun?.strike })} title="Зачёркнутый">
                <Strikethrough size={12} />
              </Toggle>
              <Toggle active={firstParagraph?.align === 'l'} onClick={() => onParagraph({ align: 'l' })} title="По левому краю">
                <AlignLeft size={12} />
              </Toggle>
              <Toggle active={firstParagraph?.align === 'ctr'} onClick={() => onParagraph({ align: 'ctr' })} title="По центру">
                <AlignCenter size={12} />
              </Toggle>
              <Toggle active={firstParagraph?.align === 'r'} onClick={() => onParagraph({ align: 'r' })} title="По правому краю">
                <AlignRight size={12} />
              </Toggle>
              <Toggle active={firstParagraph?.align === 'just'} onClick={() => onParagraph({ align: 'just' })} title="По ширине">
                <AlignJustify size={12} />
              </Toggle>
              <Toggle
                active={firstParagraph?.bullet === 'bullet'}
                onClick={() => onParagraph({ bullet: firstParagraph?.bullet === 'bullet' ? 'none' : 'bullet' })}
                title="Маркированный список"
              >
                <List size={12} />
              </Toggle>
              <Toggle
                active={firstParagraph?.bullet === 'number'}
                onClick={() => onParagraph({ bullet: firstParagraph?.bullet === 'number' ? 'none' : 'number' })}
                title="Нумерованный список"
              >
                <ListOrdered size={12} />
              </Toggle>
            </div>

            <div className="mt-1.5 grid grid-cols-2 gap-1.5">
              <label className="flex flex-col gap-0.5">
                <span className="text-[10px] text-ws-faint">Размер, пт</span>
                <input
                  type="number"
                  min={6}
                  max={200}
                  value={Math.round(firstRun?.size ?? 18)}
                  onChange={(event) => onTextFormat({ size: Math.max(6, Number(event.target.value) || 18) })}
                  className="rounded border border-ws bg-ws-editor px-1.5 py-1 text-[11px] text-ws"
                />
              </label>
              <label className="flex flex-col gap-0.5">
                <span className="text-[10px] text-ws-faint">Цвет</span>
                <input
                  type="color"
                  value={`#${normalizeHexColor(firstRun?.color ?? model.theme.colors.dk1, '000000').toLowerCase()}`}
                  onChange={(event) => onTextFormat({ color: normalizeHexColor(event.target.value) })}
                  className="h-[26px] w-full cursor-pointer rounded border border-ws"
                />
              </label>
            </div>

            <label className="mt-1.5 flex flex-col gap-0.5">
              <span className="text-[10px] text-ws-faint">Шрифт</span>
              <select
                value={firstRun?.family ?? ''}
                onChange={(event) => onTextFormat({ family: event.target.value || undefined })}
                className="rounded border border-ws bg-ws-editor px-1.5 py-1 text-[11px] text-ws"
              >
                <option value="">Из темы ({model.theme.fonts.minor})</option>
                {[...new Set(FONT_PRESETS.flatMap((pair) => [pair.major, pair.minor]))].map((family) => (
                  <option key={family} value={family}>
                    {family}
                  </option>
                ))}
              </select>
            </label>

            <label className="mt-1.5 flex items-center gap-2">
              <span className="w-14 text-[10px] text-ws-faint">Интервал</span>
              <input
                type="number"
                min={50}
                max={300}
                step={5}
                value={Math.round(firstParagraph?.lineSpacing ?? 100)}
                onChange={(event) => onParagraph({ lineSpacing: Math.max(50, Number(event.target.value) || 100) })}
                className="w-16 rounded border border-ws bg-ws-editor px-1 py-0.5 text-[11px] text-ws"
              />
              <span className="text-[10px] text-ws-faint">%</span>
              <select
                value={shape.text?.verticalAnchor ?? 't'}
                onChange={(event) => onTextBody({ verticalAnchor: event.target.value as 't' | 'ctr' | 'b' })}
                className="ml-auto rounded border border-ws bg-ws-editor px-1 py-0.5 text-[10px] text-ws"
              >
                <option value="t">Сверху</option>
                <option value="ctr">По центру</option>
                <option value="b">Снизу</option>
              </select>
            </label>
          </Section>
        </>
      ) : (
        <Section title="Слайд">
          <label className="flex items-center gap-2">
            <span className="w-16 text-[10px] text-ws-faint">Фон</span>
            <input
              type="color"
              value={`#${normalizeHexColor(slide.background, 'FFFFFF').toLowerCase()}`}
              onChange={(event) => onBackground(normalizeHexColor(event.target.value), false)}
              className="h-6 w-8 cursor-pointer rounded border border-ws"
            />
            <button
              onClick={() => onBackground(normalizeHexColor(model.theme.colors.lt1), true)}
              className="rounded border border-ws px-1.5 py-0.5 text-[10px] text-ws-muted hover:bg-ws-hover"
            >
              Для всех
            </button>
          </label>
          <div className="mt-1.5 flex gap-1">
            {(['accent1', 'accent2', 'accent3', 'accent4', 'dark'] as const).map((slot) => {
              const color =
                slot === 'dark' ? model.theme.colors.dk1 : model.theme.colors[slot as keyof typeof model.theme.colors];
              return (
                <button
                  key={slot}
                  onClick={() => onBackground(color, false)}
                  className="h-4 w-4 rounded-full border border-black/20"
                  style={{ background: `#${color}` }}
                />
              );
            })}
          </div>

          <label className="mt-3 flex flex-col gap-0.5">
            <span className="text-[10px] text-ws-faint">Переход</span>
            <select
              value={slide.transition?.type ?? 'none'}
              onChange={(event) => {
                const type = event.target.value as PptxTransitionKind;
                onTransition(type === 'none' ? undefined : { type, speed: 'med', advanceOnClick: true });
              }}
              className="rounded border border-ws bg-ws-editor px-1.5 py-1 text-[11px] text-ws"
            >
              {TRANSITIONS.map((item) => (
                <option key={item.id} value={item.id}>
                  {item.label}
                </option>
              ))}
            </select>
          </label>

          {slide.transition && slide.transition.type !== 'none' && (
            <>
              <label className="mt-1.5 flex flex-col gap-0.5">
                <span className="text-[10px] text-ws-faint">Скорость</span>
                <select
                  value={slide.transition.speed ?? 'med'}
                  onChange={(event) =>
                    onTransition({
                      ...slide.transition!,
                      speed: event.target.value as 'slow' | 'med' | 'fast',
                    })
                  }
                  className="rounded border border-ws bg-ws-editor px-1.5 py-1 text-[11px] text-ws"
                >
                  <option value="slow">Медленно</option>
                  <option value="med">Средне</option>
                  <option value="fast">Быстро</option>
                </select>
              </label>
              <label className="mt-1.5 flex items-center gap-2">
                <span className="w-16 text-[10px] text-ws-faint">Интервал</span>
                <input
                  type="number"
                  min={0}
                  max={120}
                  value={Math.round((slide.transition.advanceAfterMs ?? 0) / 1000)}
                  onChange={(event) =>
                    onTransition({
                      ...slide.transition!,
                      advanceAfterMs: Number(event.target.value) > 0 ? Number(event.target.value) * 1000 : undefined,
                    })
                  }
                  className="w-16 rounded border border-ws bg-ws-editor px-1 py-0.5 text-[11px] text-ws"
                />
                <span className="text-[10px] text-ws-faint">сек</span>
              </label>
            </>
          )}

          <div className="mt-3 rounded border border-ws bg-ws-elevated p-2 text-[10px] leading-relaxed text-ws-muted">
            Ничего не выбрано. Кликните по фигуре, чтобы править её; двойной клик — текст.
          </div>
        </Section>
      )}
    </aside>
  );
};

const Section: React.FC<{ title: string; children: React.ReactNode }> = ({ title, children }) => (
  <section className="border-b border-ws px-3 py-2.5">
    <div className="mb-1.5 text-[10px] font-semibold uppercase tracking-wider text-ws-muted">{title}</div>
    {children}
  </section>
);

const IconButton: React.FC<{
  title: string;
  onClick: () => void;
  children: React.ReactNode;
  danger?: boolean;
}> = ({ title, onClick, children, danger }) => (
  <button
    title={title}
    onClick={onClick}
    className={cn(
      'rounded border border-ws px-1.5 py-1 text-ws-muted hover:bg-ws-hover hover:text-ws',
      danger && 'hover:border-ws-danger hover:text-ws-danger'
    )}
  >
    {children}
  </button>
);

const Toggle: React.FC<{
  active?: boolean;
  onClick: () => void;
  title: string;
  children: React.ReactNode;
}> = ({ active, onClick, title, children }) => (
  <button
    title={title}
    onClick={onClick}
    className={cn(
      'rounded border px-1.5 py-1',
      active
        ? 'border-ws-accent bg-ws-accent-soft text-ws-accent'
        : 'border-ws text-ws-muted hover:bg-ws-hover'
    )}
  >
    {children}
  </button>
);
