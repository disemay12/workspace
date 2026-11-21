/**
 * Фон листов под полотном Word: сами листы, колонтитулы и номер страницы.
 *
 * Слой не интерактивен (`pointer-events: none`) и лежит под текстом, поэтому
 * не мешает ни выделению, ни каретке. Координаты — в пикселях полотна, вместе
 * с полотном масштабируются CSS zoom.
 */
import React from 'react';
import type { DocxHeaderFooterHtml } from '../../lib/docx';
import { fillHeaderFooterFields, pickHeaderFooter } from '../../lib/docxFields';
import type { PageBox } from '../../lib/pageLayout';
import type { PageGeometry } from '../../lib/units';

interface Props {
  geometry: PageGeometry;
  pages: PageBox[];
  headers?: DocxHeaderFooterHtml;
  footers?: DocxHeaderFooterHtml;
}

export const PageSheets: React.FC<Props> = ({ geometry, pages, headers, footers }) => {
  const total = pages.length;
  return (
    <div className="ws-page-layer" aria-hidden="true" data-page-layer="1">
      {pages.map((page, index) => {
        const number = index + 1;
        const header = pickHeaderFooter(headers, number);
        const footer = pickHeaderFooter(footers, number);
        return (
          <div
            key={page.index}
            className="ws-page-sheet"
            data-page-sheet={number}
            style={{
              top: page.topPx,
              left: 0,
              width: geometry.widthPx,
              height: page.heightPx,
            }}
          >
            {header ? (
              <div
                className="ws-page-header"
                data-page-header={number}
                style={{
                  top: geometry.headerPx,
                  left: geometry.marginLeftPx,
                  width: geometry.contentWidthPx,
                }}
                dangerouslySetInnerHTML={{
                  __html: fillHeaderFooterFields(header, number, total),
                }}
              />
            ) : null}
            {footer ? (
              <div
                className="ws-page-footer"
                data-page-footer={number}
                style={{
                  bottom: geometry.footerPx,
                  left: geometry.marginLeftPx,
                  width: geometry.contentWidthPx,
                }}
                dangerouslySetInnerHTML={{
                  __html: fillHeaderFooterFields(footer, number, total),
                }}
              />
            ) : null}
          </div>
        );
      })}
    </div>
  );
};
