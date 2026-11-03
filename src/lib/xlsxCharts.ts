/**
 * Нативный экспорт диаграмм WorkSpace в xlsx (OOXML).
 *
 * SheetJS CE не умеет писать диаграммы, поэтому после формирования книги
 * добавляем части xl/charts и xl/drawings прямо в zip. Excel открывает такие
 * диаграммы как обычные объекты листа — без установленного Excel при сохранении.
 */
import JSZip from 'jszip';
import { ChartSpec, parseRange, toColumnName } from './sheetFormatting';

const EMU_PER_PIXEL = 9525;
const DEFAULT_COL_PX = 104;
const DEFAULT_ROW_PX = 28;
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const MAIN_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function quoteSheet(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

function colLetter(index: number): string {
  return toColumnName(index);
}

function rangeRef(sheet: string, startRow: number, startCol: number, endRow: number, endCol: number): string {
  const start = `$${colLetter(startCol)}$${startRow + 1}`;
  const end = `$${colLetter(endCol)}$${endRow + 1}`;
  return `${quoteSheet(sheet)}!${start}:${end}`;
}

function cellRef(sheet: string, row: number, col: number): string {
  return `${quoteSheet(sheet)}!$${colLetter(col)}$${row + 1}`;
}

function chartTitleXml(title: string): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ru-RU" sz="1200" b="1"/><a:t>${escapeXml(
    title
  )}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

function seriesXml(
  sheet: string,
  range: ReturnType<typeof parseRange>,
  seriesIndex: number,
  col: number
): string {
  if (!range) return '';
  const dataStart = range.startRow + 1;
  const name = cellRef(sheet, range.startRow, col);
  const categories = rangeRef(sheet, dataStart, range.startCol, range.endRow, range.startCol);
  const values = rangeRef(sheet, dataStart, col, range.endRow, col);
  return [
    '<c:ser>',
    `<c:idx val="${seriesIndex}"/><c:order val="${seriesIndex}"/>`,
    `<c:tx><c:strRef><c:f>${escapeXml(name)}</c:f></c:strRef></c:tx>`,
    `<c:cat><c:strRef><c:f>${escapeXml(categories)}</c:f></c:strRef></c:cat>`,
    `<c:val><c:numRef><c:f>${escapeXml(values)}</c:f></c:numRef></c:val>`,
    '</c:ser>',
  ].join('');
}

function buildChartXml(sheet: string, chart: ChartSpec): string {
  const range = parseRange(chart.range);
  if (!range) return '';
  const series: string[] = [];
  const lastCol = range.endCol;
  const firstSeriesCol = range.startCol + 1;
  if (chart.type === 'pie') {
    series.push(seriesXml(sheet, range, 0, firstSeriesCol));
  } else {
    for (let col = firstSeriesCol; col <= lastCol; col++) {
      series.push(seriesXml(sheet, range, col - firstSeriesCol, col));
    }
  }

  const axisIds = '<c:axId val="111111111"/><c:axId val="222222222"/>';
  const axes =
    chart.type === 'pie'
      ? ''
      : [
          '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>',
          '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>',
        ].join('');

  const plot =
    chart.type === 'bar'
      ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series.join(
          ''
        )}${axisIds}</c:barChart>`
      : chart.type === 'line'
        ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series.join(
            ''
          )}${axisIds}</c:lineChart>`
        : `<c:pieChart><c:varyColors val="1"/>${series.join('')}</c:pieChart>`;

  const legend =
    chart.type === 'pie' || series.length > 1
      ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
      : '';

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${MAIN_NS}" xmlns:r="${REL_NS}">`,
    '<c:chart>',
    chartTitleXml(chart.title),
    '<c:autoTitleDeleted val="0"/>',
    '<c:plotArea><c:layout/>',
    plot,
    axes,
    '</c:plotArea>',
    legend,
    '<c:plotVisOnly val="1"/>',
    '<c:dispBlanksAs val="gap"/>',
    '</c:chart>',
    '</c:chartSpace>',
  ].join('');
}

function buildDrawingXml(sheet: string, charts: ChartSpec[], chartFileIndex: number): {
  xml: string;
  rels: string;
  chartFiles: { name: string; xml: string }[];
} {
  const anchors: string[] = [];
  const rels: string[] = [];
  const chartFiles: { name: string; xml: string }[] = [];

  charts.forEach((chart, index) => {
    const chartIndex = chartFileIndex + index;
    const chartName = `xl/charts/chart${chartIndex}.xml`;
    chartFiles.push({ name: chartName, xml: buildChartXml(sheet, chart) });
    rels.push(
      `<Relationship Id="rId${index + 1}" Type="${REL_NS}/chart" Target="../charts/chart${chartIndex}.xml"/>`
    );

    const col = Math.max(0, Math.floor(chart.x / DEFAULT_COL_PX));
    const colOff = Math.max(0, Math.round((chart.x % DEFAULT_COL_PX) * EMU_PER_PIXEL));
    const row = Math.max(0, Math.floor(chart.y / DEFAULT_ROW_PX));
    const rowOff = Math.max(0, Math.round((chart.y % DEFAULT_ROW_PX) * EMU_PER_PIXEL));
    const width = Math.round(chart.width * EMU_PER_PIXEL);
    const height = Math.round(chart.height * EMU_PER_PIXEL);

    anchors.push(
      [
        '<xdr:oneCellAnchor>',
        `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:from>`,
        `<xdr:ext cx="${width}" cy="${height}"/>`,
        '<xdr:graphicFrame macro="">',
        `<xdr:nvGraphicFramePr><xdr:cNvPr id="${index + 2}" name="Chart ${index + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`,
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>',
        `<a:graphic><a:graphicData uri="${CHART_NS}"><c:chart xmlns:c="${CHART_NS}" xmlns:r="${REL_NS}" r:id="rId${index + 1}"/></a:graphicData></a:graphic>`,
        '</xdr:graphicFrame>',
        '<xdr:clientData/>',
        '</xdr:oneCellAnchor>',
      ].join('')
    );
  });

  const xml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<xdr:wsDr xmlns:xdr="${DRAWING_NS}" xmlns:a="${MAIN_NS}" xmlns:r="${REL_NS}">`,
    anchors.join(''),
    '</xdr:wsDr>',
  ].join('');

  const relsXml = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`,
    rels.join(''),
    '</Relationships>',
  ].join('');

  return { xml, rels: relsXml, chartFiles };
}

interface SheetEntry {
  name: string;
  path: string;
}

async function readSheetMap(zip: JSZip): Promise<SheetEntry[]> {
  const workbookFile = zip.file('xl/workbook.xml');
  const relsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (!workbookFile || !relsFile) return [];
  const workbookXml = await workbookFile.async('string');
  const relsXml = await relsFile.async('string');

  const targets = new Map<string, string>();
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const tag = match[0];
    const id = /\bId="([^"]+)"/.exec(tag)?.[1];
    const target = /\bTarget="([^"]+)"/.exec(tag)?.[1];
    if (id && target) targets.set(id, target);
  }

  const result: SheetEntry[] = [];
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*\/?>/g)) {
    const tag = match[0];
    const name = /\bname="([^"]+)"/.exec(tag)?.[1];
    const rid = /\br:id="([^"]+)"/.exec(tag)?.[1];
    if (!name || !rid) continue;
    const target = targets.get(rid);
    if (!target) continue;
    const normalized = target.startsWith('/')
      ? target.slice(1)
      : `xl/${target.replace(/^\.\//, '')}`;
    result.push({ name, path: normalized.replace(/\/\.\//g, '/') });
  }
  return result;
}

function decodeEntities(value: string): string {
  return value
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

export interface ChartExportEntry {
  sheet: string;
  charts: ChartSpec[];
}

export async function injectChartsIntoXlsx(
  base64: string,
  entries: ChartExportEntry[]
): Promise<string> {
  const relevant = entries.filter((entry) => entry.charts.length > 0);
  if (relevant.length === 0) return base64;

  const zip = await JSZip.loadAsync(base64, { base64: true });
  const sheetMap = await readSheetMap(zip);
  const contentTypesFile = zip.file('[Content_Types].xml');
  if (!contentTypesFile) return base64;
  let contentTypes = await contentTypesFile.async('string');
  const overrides: string[] = [];
  let chartCursor = 1;
  let drawingCursor = 1;

  for (const entry of relevant) {
    const sheetEntry = sheetMap.find(
      (item) => decodeEntities(item.name) === entry.sheet
    );
    if (!sheetEntry) continue;
    const sheetFile = zip.file(sheetEntry.path);
    if (!sheetFile) continue;
    let sheetXml = await sheetFile.async('string');
    // Если лист уже содержит drawing, не рискуем ломать чужую графику.
    if (/<drawing\b/.test(sheetXml)) continue;

    const drawingIndex = drawingCursor++;
    const drawingPath = `xl/drawings/drawing${drawingIndex}.xml`;
    const relsPath = sheetEntry.path.replace(/worksheets\/([^/]+)$/, 'worksheets/_rels/$1.rels');
    const drawing = buildDrawingXml(entry.sheet, entry.charts, chartCursor);
    chartCursor += entry.charts.length;

    zip.file(drawingPath, drawing.xml);
    zip.file(
      `xl/drawings/_rels/drawing${drawingIndex}.xml.rels`,
      drawing.rels
    );
    for (const chart of drawing.chartFiles) {
      zip.file(chart.name, chart.xml);
      overrides.push(
        `<Override PartName="/${chart.name}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
      );
    }
    overrides.push(
      `<Override PartName="/${drawingPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
    );

    // Связь листа с drawing
    let relsXml = '';
    const existingRels = zip.file(relsPath);
    if (existingRels) relsXml = await existingRels.async('string');
    let nextRel = 1;
    for (const match of relsXml.matchAll(/\bId="rId(\d+)"/g)) {
      nextRel = Math.max(nextRel, Number(match[1]) + 1);
    }
    if (!relsXml) {
      relsXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '</Relationships>';
    }
    const relation = `<Relationship Id="rId${nextRel}" Type="${REL_NS}/drawing" Target="../drawings/drawing${drawingIndex}.xml"/>`;
    relsXml = relsXml.replace('</Relationships>', `${relation}</Relationships>`);
    zip.file(relsPath, relsXml);

    sheetXml = sheetXml.replace('</worksheet>', `<drawing r:id="rId${nextRel}"/></worksheet>`);
    zip.file(sheetEntry.path, sheetXml);
  }

  if (overrides.length === 0) return base64;
  contentTypes = contentTypes.replace('</Types>', `${overrides.join('')}</Types>`);
  zip.file('[Content_Types].xml', contentTypes);

  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}
