/**
 * Метрики шрифта для точного межстрочного интервала.
 *
 * Word задаёт «множитель» (w:lineRule="auto", w:line="259") как кратность
 * естественной высоте строки шрифта, а CSS `line-height: 1.08` — кратность
 * кеглю. Для Calibri 11 pt это 11,9 pt вместо 14,9 pt, то есть строки на 20 %
 * теснее и весь документ «съезжает» на страницы раньше. Поэтому множитель
 * переводится через измеренное отношение: CSS = множитель × (естественная
 * высота / кегль). Один замер на семейство шрифта, дальше кэш.
 */

const FALLBACK_RATIO = 1.221;
const cache = new Map<string, number>();
let host: HTMLElement | null = null;

function measure(fontFamily: string): number {
  if (typeof document === 'undefined' || !document.body) return FALLBACK_RATIO;
  if (!host) {
    host = document.createElement('div');
    host.setAttribute('aria-hidden', 'true');
    host.style.cssText =
      'position:fixed;left:-10000px;top:0;visibility:hidden;pointer-events:none;contain:strict;';
    document.body.append(host);
  }
  const probe = document.createElement('div');
  probe.style.cssText = 'font-size:100px;line-height:normal;white-space:nowrap;';
  probe.style.fontFamily = `"${fontFamily.replace(/["']/g, '')}", serif`;
  probe.textContent = 'Ag';
  host.append(probe);
  const ratio = probe.getBoundingClientRect().height / 100;
  probe.remove();
  return Number.isFinite(ratio) && ratio > 0.8 && ratio < 2.6 ? ratio : FALLBACK_RATIO;
}

/** Отношение естественной высоты строки к кеглю для семейства шрифта. */
export function naturalLineHeightRatio(fontFamily: string): number {
  const key = (fontFamily || '').trim().toLowerCase();
  const cached = cache.get(key);
  if (cached !== undefined) return cached;
  const ratio = key ? measure(key) : FALLBACK_RATIO;
  cache.set(key, ratio);
  return ratio;
}

/**
 * CSS межстрочного интервала из w:spacing.
 * `line` — twips, `rule` — w:lineRule (auto / exact / atLeast).
 */
export function lineHeightCss(
  lineTwips: number,
  rule: string,
  fontFamily: string
): string | null {
  if (!(lineTwips > 0)) return null;
  const points = Math.round((lineTwips / 20) * 100) / 100;
  if (rule === 'exact') return `line-height:${points}pt`;
  if (rule === 'atLeast') return `line-height:max(${points}pt, 1.15em)`;
  const multiple = Math.round((lineTwips / 240) * 1000) / 1000;
  if (Math.abs(multiple - 1) < 0.005) return 'line-height:normal';
  const ratio = naturalLineHeightRatio(fontFamily);
  return `line-height:${Math.round(multiple * ratio * 1000) / 1000}`;
}
