/**
 * Печатная HTML-раскладка презентации: слайд — страница своего формата.
 *
 * Вынесено из PowerPointEditor, чтобы тем же кодом печаталась и презентация,
 * открытая из проводника (без живого редактора).
 */
import type { PptxDocModel, PptxTextBody } from './pptxDoc';

const EMU_PER_INCH = 914400;

/** HTML всех слайдов (или одного) с размером страницы под слайд. */
export function buildPptxPrintHtml(model: PptxDocModel, only?: number): string {
  const widthIn = model.width / EMU_PER_INCH;
  const heightIn = model.height / EMU_PER_INCH;
  const px = (emu: number) => `${(emu / EMU_PER_INCH).toFixed(4)}in`;

  const slides = model.slides
    .filter((_, index) => only === undefined || index === only)
    .map((slide) => {
      const shapes = slide.shapes
        .map((shape) => {
          const box = `position:absolute;left:${px(shape.x)};top:${px(shape.y)};width:${px(
            shape.w
          )};height:${px(shape.h)};overflow:hidden;`;
          const rotate = shape.rotation ? `transform:rotate(${shape.rotation}deg);` : '';
          if (shape.kind === 'image' && shape.image?.src) {
            return `<img src="${shape.image.src}" style="${box}${rotate}object-fit:contain">`;
          }
          if (shape.kind === 'table' && shape.table) {
            const rows = shape.table.rows
              .map(
                (row) =>
                  '<tr>' +
                  row
                    .map(
                      (cell) =>
                        `<td style="border:1px solid #cbd5e1;padding:2pt;vertical-align:middle;background:${
                          cell.fill ? '#' + cell.fill : '#ffffff'
                        }">${textBodyHtml(cell.text)}</td>`
                    )
                    .join('') +
                  '</tr>'
              )
              .join('');
            return `<div style="${box}${rotate}"><table style="width:100%;height:100%;border-collapse:collapse;table-layout:fixed">${rows}</table></div>`;
          }
          const background = shape.fill ? `background:#${shape.fill};` : '';
          const border = shape.outline?.color
            ? `border:${(shape.outline.width ?? 1).toFixed(2)}pt solid #${shape.outline.color};`
            : '';
          return `<div style="${box}${rotate}${background}${border}padding:2pt">${
            shape.text ? textBodyHtml(shape.text) : ''
          }</div>`;
        })
        .join('');

      const background = slide.backgroundImage
        ? `background:#000 url(${slide.backgroundImage}) center/cover no-repeat;`
        : `background:#${slide.background};`;

      return `<div class="slide" style="${background}">${shapes}</div>`;
    })
    .join('');

  return (
    `<!doctype html><html><head><meta charset="utf-8"><style>` +
    `@page{size:${widthIn.toFixed(4)}in ${heightIn.toFixed(4)}in;margin:0}` +
    `html,body{margin:0;padding:0}` +
    `.slide{position:relative;width:${widthIn.toFixed(4)}in;height:${heightIn.toFixed(4)}in;` +
    `page-break-after:always;break-after:page;overflow:hidden}` +
    `.slide:last-child{page-break-after:auto;break-after:auto}` +
    `</style></head><body>` +
    slides +
    `</body></html>`
  );
}

/** Абзацы текстового тела фигуры с оформлением run'ов. */
export function textBodyHtml(body: PptxTextBody): string {
  return body.paragraphs
    .map((paragraph) => {
      const align =
        paragraph.align === 'ctr'
          ? 'center'
          : paragraph.align === 'r'
            ? 'right'
            : paragraph.align === 'just'
              ? 'justify'
              : 'left';
      const runs = paragraph.runs
        .map((run) => {
          const style = [
            run.size ? `font-size:${run.size}pt` : '',
            run.color ? `color:#${run.color}` : '',
            run.bold ? 'font-weight:700' : '',
            run.italic ? 'font-style:italic' : '',
            run.underline ? 'text-decoration:underline' : '',
            run.family ? `font-family:'${run.family}','Segoe UI',sans-serif` : '',
          ]
            .filter(Boolean)
            .join(';');
          const text = run.text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
          return `<span style="${style}">${text}</span>`;
        })
        .join('');
      const bullet =
        paragraph.bullet === 'bullet' ? '• ' : paragraph.bullet === 'number' ? '1. ' : '';
      return `<div style="text-align:${align};margin:0;line-height:${
        paragraph.lineSpacing ? paragraph.lineSpacing / 100 : 1.2
      }">${bullet}${runs || '&nbsp;'}</div>`;
    })
    .join('');
}

/** Размер страницы презентации в микронах — для системного диалога печати. */
export function pptxPageMicrons(model: PptxDocModel): { width: number; height: number } {
  return {
    width: Math.round((model.width / EMU_PER_INCH) * 25400),
    height: Math.round((model.height / EMU_PER_INCH) * 25400),
  };
}
