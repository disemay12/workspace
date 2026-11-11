import React, { useMemo } from 'react';
import { cn } from '../../lib/utils';
import { isDarkHex } from '../../lib/pptxThemes';
import type { PptxDocModel, PptxParagraph, PptxShapeModel, PptxSlideModel, PptxTextBody } from '../../lib/pptxDoc';

/**
 * Рендер слайда: реальные координаты OOXML (EMU → проценты), тексты по
 * абзацам и run'ам, изображения, таблицы, заливки и обводки фигур.
 *
 * Размер шрифта пересчитывается из пунктов: 1 пункт = 12700 EMU, поэтому
 * `pt * widthPx / (slideWidthEmu / 12700)` даёт точку на экране при любом
 * масштабе — от миниатюры до режима показа.
 */

export interface SlideCanvasProps {
  slide: PptxSlideModel;
  model: PptxDocModel;
  /** Ширина холста в пикселях: остальное считается от неё. */
  width: number;
  className?: string;
  selectedId?: string | null;
  /** Двойной клик по фигуре — вход в правку текста. */
  onShapeDoubleClick?: (shape: PptxShapeModel) => void;
  onShapeMouseDown?: (shape: PptxShapeModel, event: React.PointerEvent) => void;
  /** Скрывать текст при перетаскивании, чтобы не мигали каретки. */
  hideCaret?: boolean;
  /** Временная геометрия при перетаскивании, до фиксации в документе. */
  geometryOverride?: Record<string, { x: number; y: number; w: number; h: number }>;
}

const EMU_PER_PT = 12700;

export const SlideCanvas: React.FC<SlideCanvasProps> = ({
  slide,
  model,
  width,
  className,
  selectedId,
  onShapeDoubleClick,
  onShapeMouseDown,
  geometryOverride,
}) => {
  const height = Math.round((width * model.height) / model.width);
  const ptToPx = useMemo(() => width / (model.width / EMU_PER_PT), [width, model.width]);
  const dark = isDarkHex(slide.background);
  const pct = (value: number, total: number) => `${(value / total) * 100}%`;

  return (
    <div
      className={cn('relative overflow-hidden', className)}
      style={{
        width: '100%',
        aspectRatio: `${model.width} / ${model.height}`,
        background: slide.backgroundImage
          ? `#000 url(${slide.backgroundImage}) center / cover no-repeat`
          : `#${slide.background}`,
        height,
      }}
    >
      {slide.shapes.map((shape) => {
        const selected = selectedId === shape.id;
        const geometry = geometryOverride?.[shape.id] ?? shape;
        const style: React.CSSProperties = {
          position: 'absolute',
          left: pct(geometry.x, model.width),
          top: pct(geometry.y, model.height),
          width: pct(geometry.w, model.width),
          height: pct(geometry.h, model.height),
          transform: shape.rotation ? `rotate(${shape.rotation}deg)` : undefined,
          transformOrigin: 'center center',
        };

        const interactive = Boolean(onShapeDoubleClick || onShapeMouseDown);

        if (shape.kind === 'image') {
          return (
            <div
              key={shape.id}
              style={style}
              className={cn('select-none', interactive && 'cursor-move')}
              onDoubleClick={() => onShapeDoubleClick?.(shape)}
              onPointerDown={(event) => onShapeMouseDown?.(shape, event)}
            >
              {shape.image?.src ? (
                <img
                  src={shape.image.src}
                  alt={shape.name}
                  draggable={false}
                  className="h-full w-full object-contain"
                  style={{ boxShadow: shape.shadow ? '0 6px 18px rgba(0,0,0,0.28)' : undefined }}
                />
              ) : (
                <div className="flex h-full w-full items-center justify-center rounded border border-dashed border-ws-faint/40 text-[10px] text-ws-faint">
                  Изображение
                </div>
              )}
              {selected && <SelectionFrame />}
            </div>
          );
        }

        if (shape.kind === 'table' && shape.table) {
          const table = shape.table;
          return (
            <div
              key={shape.id}
              style={style}
              className={cn('select-none', interactive && 'cursor-move')}
              onDoubleClick={() => onShapeDoubleClick?.(shape)}
              onPointerDown={(event) => onShapeMouseDown?.(shape, event)}
            >
              <table className="h-full w-full table-fixed border-collapse">
                <tbody>
                  {table.rows.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {row.map((cell, cellIndex) => (
                        <td
                          key={cellIndex}
                          className="border border-slate-300 p-1 align-middle"
                          style={{ background: cell.fill ? `#${cell.fill}` : '#FFFFFF' }}
                        >
                          <TextBody
                            body={cell.text}
                            ptToPx={ptToPx}
                            defaultColor={dark ? 'FFFFFF' : '000000'}
                            theme={model.theme}
                          />
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {selected && <SelectionFrame />}
            </div>
          );
        }

        // Текстовая фигура, надпись или автофигура с заливкой.
        return (
          <div
            key={shape.id}
            style={{
              ...style,
              background:
                shape.kind === 'preset' && shape.fill === null
                  ? 'transparent'
                  : shape.fill
                    ? `#${shape.fill}`
                    : undefined,
              border: shape.outline?.color
                ? `${Math.max(1, (shape.outline.width ?? 1) * ptToPx)}px solid #${shape.outline.color}`
                : undefined,
              boxShadow: shape.shadow ? '0 6px 18px rgba(0,0,0,0.28)' : undefined,
            }}
            className={cn(
              'select-none overflow-hidden p-1',
              interactive && 'cursor-move',
              shape.geom === 'ellipse' && 'rounded-full'
            )}
            onDoubleClick={() => onShapeDoubleClick?.(shape)}
            onPointerDown={(event) => onShapeMouseDown?.(shape, event)}
          >
            {shape.text && (
              <TextBody
                body={shape.text}
                ptToPx={ptToPx}
                defaultColor={dark ? 'F8FAFC' : '111827'}
                theme={model.theme}
              />
            )}
            {selected && <SelectionFrame />}
          </div>
        );
      })}
    </div>
  );
};

const SelectionFrame: React.FC = () => (
  <div className="pointer-events-none absolute inset-0 border-2 border-blue-500/90">
    {[
      ['-top-1 -left-1', 'nw'],
      ['-top-1 -right-1', 'ne'],
      ['-bottom-1 -left-1', 'sw'],
      ['-bottom-1 -right-1', 'se'],
    ].map(([position, key]) => (
      <span
        key={key}
        className={cn('absolute h-2 w-2 rounded-sm border border-white bg-blue-500', position)}
      />
    ))}
  </div>
);

/* ============================================================
   Текст абзацев и run'ов
   ============================================================ */

const TextBody: React.FC<{
  body: PptxTextBody;
  ptToPx: number;
  defaultColor: string;
  theme: PptxDocModel['theme'];
}> = ({ body, ptToPx, defaultColor, theme }) => (
  <div
    className="flex h-full w-full flex-col"
    style={{
      justifyContent:
        body.verticalAnchor === 'ctr' ? 'center' : body.verticalAnchor === 'b' ? 'flex-end' : 'flex-start',
    }}
  >
    {body.paragraphs.map((paragraph, index) => (
      <Paragraph
        key={index}
        paragraph={paragraph}
        ptToPx={ptToPx}
        defaultColor={defaultColor}
        theme={theme}
      />
    ))}
  </div>
);

const Paragraph: React.FC<{
  paragraph: PptxParagraph;
  ptToPx: number;
  defaultColor: string;
  theme: PptxDocModel['theme'];
}> = ({ paragraph, ptToPx, defaultColor, theme }) => {
  const level = paragraph.level ?? 0;
  const align = paragraph.align ?? 'l';
  const runs = paragraph.runs.length > 0 ? paragraph.runs : [];
  const bullet = paragraph.bullet === 'bullet' || paragraph.bullet === 'number';
  const fallbackSize = runs[0]?.size ?? 18;

  return (
    <div
      style={{
        textAlign: align === 'ctr' ? 'center' : align === 'r' ? 'right' : align === 'just' ? 'justify' : 'left',
        paddingLeft: bullet ? `${(level + 1) * 1.2}em` : level > 0 ? `${level * 1.2}em` : undefined,
        textIndent: bullet ? '-1.2em' : undefined,
        lineHeight: paragraph.lineSpacing ? paragraph.lineSpacing / 100 : 1.2,
        marginTop: paragraph.spaceBefore ? paragraph.spaceBefore * ptToPx : undefined,
        marginBottom: paragraph.spaceAfter ? paragraph.spaceAfter * ptToPx : undefined,
        whiteSpace: 'pre-wrap',
        overflowWrap: 'anywhere',
      }}
    >
      {bullet && (
        <span style={{ fontSize: fallbackSize * ptToPx, color: `#${defaultColor}` }}>
          {paragraph.bullet === 'number' ? '1. ' : '• '}
        </span>
      )}
      {runs.length === 0 ? (
        <span style={{ fontSize: fallbackSize * ptToPx }}>&nbsp;</span>
      ) : (
        runs.map((run, index) => {
          const size = run.size ?? fallbackSize;
          const color = run.color ?? defaultColor;
          const family =
            run.family ??
            (run.size && run.size >= 28 ? theme.fonts.major : theme.fonts.minor);
          return (
            <span
              key={index}
              style={{
                fontSize: size * ptToPx,
                color: `#${color}`,
                fontFamily: `"${family}", "Segoe UI", sans-serif`,
                fontWeight: run.bold ? 700 : 400,
                fontStyle: run.italic ? 'italic' : undefined,
                textDecoration: [
                  run.underline ? 'underline' : '',
                  run.strike ? 'line-through' : '',
                ]
                  .filter(Boolean)
                  .join(' ') || undefined,
                background: run.highlight ? `#${run.highlight}` : undefined,
              }}
            >
              {run.text}
            </span>
          );
        })
      )}
    </div>
  );
};
