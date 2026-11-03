import React, { useMemo } from 'react';
import { ChartSpec } from '../../lib/sheetFormatting';

export interface ChartSeries {
  name: string;
  values: (number | null)[];
}

interface Props {
  spec: ChartSpec;
  labels: string[];
  series: ChartSeries[];
}

const COLORS = ['#10b981', '#3b82f6', '#f59e0b', '#a855f7', '#ef4444', '#06b6d4'];

const WIDTH = 480;
const HEIGHT = 260;

/** Диаграмма листа: столбчатая, линейная или круговая. */
export const SheetChart: React.FC<Props> = ({ spec, labels, series }) => {
  const data = useMemo(() => {
    const all = series.flatMap((item) => item.values).filter((value): value is number => value !== null);
    const min = all.length ? Math.min(0, ...all) : 0;
    const max = all.length ? Math.max(1, ...all) : 1;
    return { min, max, count: Math.max(1, labels.length) };
  }, [labels.length, series]);

  if (!series.length || !labels.length) {
    return (
      <div className="flex h-full items-center justify-center text-xs text-ws-faint">
        Нет числовых данных в диапазоне
      </div>
    );
  }

  if (spec.type === 'pie') {
    const values = series[0].values.map((value) => Math.max(0, value ?? 0));
    const total = values.reduce((sum, value) => sum + value, 0) || 1;
    let angle = -Math.PI / 2;
    const cx = 130;
    const cy = 120;
    const radius = 92;
    return (
      <svg viewBox={`0 0 ${WIDTH} ${HEIGHT}`} className="h-full w-full" role="img" aria-label={spec.title}>
        {values.map((value, index) => {
          const slice = (value / total) * Math.PI * 2;
          const end = angle + slice;
          const large = slice > Math.PI ? 1 : 0;
          const x1 = cx + radius * Math.cos(angle);
          const y1 = cy + radius * Math.sin(angle);
          const x2 = cx + radius * Math.cos(end);
          const y2 = cy + radius * Math.sin(end);
          angle = end;
          if (value <= 0) return null;
          return (
            <path
              key={index}
              d={`M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2} Z`}
              fill={COLORS[index % COLORS.length]}
              stroke="rgba(255,255,255,0.65)"
              strokeWidth={1}
            />
          );
        })}
        {labels.map((label, index) => (
          <g key={label + index}>
            <rect x={270} y={34 + index * 22} width={10} height={10} rx={2} fill={COLORS[index % COLORS.length]} />
            <text x={286} y={43 + index * 22} fontSize={11} fill="currentColor">
              {label.length > 24 ? `${label.slice(0, 24)}…` : label} — {Math.round(values[index] ?? 0)}
            </text>
          </g>
        ))}
      </svg>
    );
  }

  const plotLeft = 44;
  const plotRight = WIDTH - 14;
  const plotTop = 16;
  const plotBottom = HEIGHT - 34;
  const plotWidth = plotRight - plotLeft;
  const plotHeight = plotBottom - plotTop;
  const scaleY = (value: number) =>
    plotBottom - ((value - data.min) / Math.max(0.0001, data.max - data.min)) * plotHeight;
  const categoryWidth = plotWidth / data.count;

  return (
    <svg viewBox={`0 0 ${WIDTH} ${HEIGHT}`} className="h-full w-full" role="img" aria-label={spec.title}>
      {[0, 0.25, 0.5, 0.75, 1].map((fraction) => {
        const y = plotTop + plotHeight * fraction;
        const value = data.max - (data.max - data.min) * fraction;
        return (
          <g key={fraction}>
            <line x1={plotLeft} y1={y} x2={plotRight} y2={y} stroke="currentColor" strokeOpacity={0.12} />
            <text x={4} y={y + 4} fontSize={9} fill="currentColor" fillOpacity={0.55}>
              {Math.round(value * 100) / 100}
            </text>
          </g>
        );
      })}
      {spec.type === 'bar'
        ? series.map((item, seriesIndex) => {
            const barWidth = Math.max(4, (categoryWidth * 0.72) / series.length);
            return item.values.map((value, index) => {
              if (value === null) return null;
              const x =
                plotLeft +
                index * categoryWidth +
                categoryWidth * 0.14 +
                seriesIndex * barWidth;
              const y = scaleY(value);
              const height = Math.max(1, plotBottom - y);
              return (
                <rect
                  key={`${seriesIndex}-${index}`}
                  x={x}
                  y={y}
                  width={barWidth - 2}
                  height={height}
                  rx={2}
                  fill={COLORS[seriesIndex % COLORS.length]}
                />
              );
            });
          })
        : series.map((item, seriesIndex) => {
            const points = item.values
              .map((value, index) => {
                if (value === null) return null;
                const x = plotLeft + index * categoryWidth + categoryWidth / 2;
                return `${x},${scaleY(value)}`;
              })
              .filter(Boolean)
              .join(' ');
            return (
              <g key={seriesIndex}>
                <polyline
                  points={points}
                  fill="none"
                  stroke={COLORS[seriesIndex % COLORS.length]}
                  strokeWidth={2}
                  strokeLinejoin="round"
                />
                {item.values.map((value, index) =>
                  value === null ? null : (
                    <circle
                      key={index}
                      cx={plotLeft + index * categoryWidth + categoryWidth / 2}
                      cy={scaleY(value)}
                      r={2.5}
                      fill={COLORS[seriesIndex % COLORS.length]}
                    />
                  )
                )}
              </g>
            );
          })}
      {labels.map((label, index) => (
        <text
          key={label + index}
          x={plotLeft + index * categoryWidth + categoryWidth / 2}
          y={HEIGHT - 12}
          fontSize={9}
          textAnchor="middle"
          fill="currentColor"
          fillOpacity={0.65}
        >
          {label.length > 10 ? `${label.slice(0, 10)}…` : label}
        </text>
      ))}
    </svg>
  );
};
