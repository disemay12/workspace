/**
 * Конвертация документов между форматами WorkSpace.
 *
 * Все функции работают в рендерере: Word/Excel/PowerPoint уже умеют
 * разбирать и собирать свои пакеты. Здесь только выбирается нужный путь
 * (Markdown → DOCX, CSV → XLSX, DOCX → MD и т.п.).
 */
import * as XLSX from 'xlsx';
import type { FileType } from '../types';
import { baseName, extName } from '../lib/utils';
import { docxBase64ToHtml, htmlToDocxBase64 } from '../lib/docx';
import { markdownToDocumentHtml } from '../lib/aiDocument';
import { htmlToPlainText, pptxText } from '../lib/documentText';
import { extractLegacyDocText, plainTextToHtml } from '../lib/legacyDoc';
import { detectMarkdownTables } from './aiActions';
import { applyAutoFitColumns } from '../lib/xlsxColumns';
import { fileSystemService, getFileType } from './fileSystemService';

export type ConversionFormat = 'xlsx' | 'csv' | 'tsv' | 'docx' | 'md' | 'txt' | 'html' | 'pdf' | 'pptx';

export interface ConversionSource {
  path: string;
  name: string;
  extension: string;
  type: FileType;
  content: string;
  encoding: string;
  /** Активный лист таблицы, если он известен. */
  activeSheet?: string;
}

export type ConversionResult =
  | {
      kind: 'data';
      content: string;
      encoding: string;
      extension: string;
      note?: string;
    }
  | {
      kind: 'pdf';
      html: string;
      extension: 'pdf';
      note?: string;
    };

const FORMAT_LABELS: Record<ConversionFormat, string> = {
  xlsx: 'Excel (.xlsx)',
  csv: 'CSV (.csv)',
  tsv: 'TSV (.tsv)',
  docx: 'Word (.docx)',
  md: 'Markdown (.md)',
  txt: 'Текст (.txt)',
  html: 'HTML (.html)',
  pdf: 'PDF (.pdf)',
  pptx: 'PowerPoint (.pptx)',
};

const BINARY_EXTENSIONS = new Set([
  'docx',
  'doc',
  'xlsx',
  'xls',
  'xlsm',
  'pptx',
  'ppt',
  'pdf',
  'png',
  'jpg',
  'jpeg',
  'jpe',
  'jfif',
  'gif',
  'webp',
  'bmp',
  'ico',
  'tif',
  'tiff',
  'avif',
  'svg',
]);

export function formatLabel(format: ConversionFormat): string {
  return FORMAT_LABELS[format] ?? format.toUpperCase();
}

export function formatExtension(format: ConversionFormat): string {
  return format;
}

/** Приводит «excel», «word», «markdown» и т.п. к каноническому формату. */
export function normalizeConversionFormat(value?: string): ConversionFormat | null {
  const raw = (value || '').trim().toLowerCase().replace(/^\./, '');
  const aliases: Record<string, ConversionFormat> = {
    xlsx: 'xlsx',
    excel: 'xlsx',
    xls: 'xlsx',
    xlsm: 'xlsx',
    csv: 'csv',
    tsv: 'tsv',
    docx: 'docx',
    word: 'docx',
    doc: 'docx',
    md: 'md',
    markdown: 'md',
    mdx: 'md',
    txt: 'txt',
    text: 'txt',
    html: 'html',
    htm: 'html',
    pdf: 'pdf',
    pptx: 'pptx',
    powerpoint: 'pptx',
    ppt: 'pptx',
  };
  return aliases[raw] ?? null;
}

export function guessFormat(path: string): ConversionFormat | null {
  const ext = extName(path);
  if (ext === 'markdown' || ext === 'mdx') return 'md';
  if (ext === 'htm') return 'html';
  if (ext === 'jpeg' || ext === 'jpg') return null;
  return (['xlsx', 'csv', 'tsv', 'docx', 'md', 'txt', 'html', 'pdf', 'pptx'] as ConversionFormat[]).includes(
    ext as ConversionFormat
  )
    ? (ext as ConversionFormat)
    : null;
}

/** Форматы, в которые можно сохранить открытый файл. */
export function formatsFor(source: { extension: string; type: FileType }): ConversionFormat[] {
  const ext = source.extension.toLowerCase();
  if (source.type === 'word') return ['docx', 'md', 'txt', 'html', 'pdf'];
  if (source.type === 'excel') {
    if (ext === 'csv' || ext === 'tsv') return ['csv', 'tsv', 'xlsx', 'md', 'html', 'pdf'];
    return ['xlsx', 'csv', 'tsv', 'md', 'html', 'pdf'];
  }
  if (source.type === 'markdown') return ['md', 'docx', 'txt', 'html', 'pdf', 'xlsx'];
  if (source.type === 'richtext' || source.type === 'text' || source.type === 'code') {
    return ['txt', 'md', 'docx', 'html', 'pdf'];
  }
  if (source.type === 'powerpoint') return ['pptx', 'md', 'txt', 'docx', 'html', 'pdf'];
  return [];
}

export function defaultTargetName(sourceName: string, format: ConversionFormat): string {
  const base = sourceName.replace(/\.[^./\\]+$/, '') || 'Документ';
  return `${base}.${formatExtension(format)}`;
}

/** Читает файл с диска как источник конвертации. */
export async function readFileSource(filePath: string): Promise<ConversionSource> {
  const extension = extName(filePath);
  const type = getFileType(extension);
  const encoding = BINARY_EXTENSIONS.has(extension) ? 'base64' : 'utf-8';
  const res = await fileSystemService.readFile(filePath, encoding);
  return {
    path: filePath,
    name: baseName(filePath),
    extension,
    type,
    content: res.content,
    encoding,
  };
}

function isZipBase64(content: string): boolean {
  return /^UEsDB/.test(content.replace(/\s+/g, '').slice(0, 8));
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function stripMarkdown(text: string): string {
  return text
    .replace(/^#{1,6}\s+/gm, '')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\*([^*]+)\*/g, '$1')
    .replace(/__([^_]+)__/g, '$1')
    .replace(/~~([^~]+)~~/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/^\s*[-*•]\s+/gm, '')
    .replace(/^\s*\d+[.)]\s+/gm, '')
    .replace(/^\s*>\s?/gm, '')
    .replace(/\|/g, '\t')
    .replace(/^[-:\t ]+$/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function matrixFromCsv(text: string, delimiter: ',' | '\t'): string[][] {
  const workbook = XLSX.read(text, { type: 'string', FS: delimiter });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    blankrows: false,
    raw: false,
  }) as string[][];
}

function matrixFromMarkdown(text: string): string[][] | null {
  const table = detectMarkdownTables(text).sort((a, b) => b.rows.length - a.rows.length)[0];
  if (!table) return null;
  return [table.columns, ...table.rows];
}

async function htmlFromSource(source: ConversionSource): Promise<string> {
  const ext = source.extension.toLowerCase();
  if (source.type === 'word') {
    if (ext === 'doc' && !isZipBase64(source.content)) {
      return plainTextToHtml(await extractLegacyDocText(source.content));
    }
    return (await docxBase64ToHtml(source.content)).html;
  }
  if (source.type === 'markdown') return markdownToDocumentHtml(source.content);
  if (source.type === 'richtext' || source.type === 'text' || source.type === 'code') {
    return plainTextToHtml(source.content);
  }
  if (ext === 'html' || ext === 'htm') return source.content;
  if (source.type === 'excel') return htmlFromWorkbook(source);
  if (source.type === 'powerpoint') {
    return markdownToDocumentHtml(await pptxText(source.content));
  }
  throw new Error(`Формат .${ext || source.type} не поддерживает конвертацию документа`);
}

async function markdownFromSource(source: ConversionSource): Promise<string> {
  const ext = source.extension.toLowerCase();
  if (source.type === 'markdown') return source.content;
  if (source.type === 'richtext' || source.type === 'text' || source.type === 'code') {
    return source.content;
  }
  if (ext === 'html' || ext === 'htm') return htmlToPlainText(source.content);
  if (source.type === 'word') {
    if (ext === 'doc' && !isZipBase64(source.content)) {
      return stripMarkdown(await extractLegacyDocText(source.content));
    }
    return htmlToPlainText((await docxBase64ToHtml(source.content)).html);
  }
  if (source.type === 'excel') return markdownFromWorkbook(source);
  if (source.type === 'powerpoint') return await pptxText(source.content);
  throw new Error(`Формат .${ext || source.type} не поддерживается`);
}

function workbookFromSource(source: ConversionSource): XLSX.WorkBook {
  return XLSX.read(source.content, {
    type: source.encoding === 'base64' ? 'base64' : 'string',
    bookVBA: source.extension.toLowerCase() === 'xlsm',
    cellStyles: true,
  });
}

function matrixFromSheet(sheet: XLSX.WorkSheet): string[][] {
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    blankrows: false,
    raw: false,
  }) as string[][];
}

function matrixToMarkdown(matrix: string[][]): string {
  if (matrix.length === 0) return '';
  const width = Math.max(...matrix.map((row) => row.length));
  const grid = matrix.map((row) => [
    ...row.map((cell) => String(cell ?? '').replace(/\|/g, '\\|').replace(/\n/g, '<br>')),
    ...Array(Math.max(0, width - row.length)).fill(''),
  ]);
  const head = `| ${grid[0].join(' | ')} |`;
  const divider = `| ${Array(width).fill('---').join(' | ')} |`;
  const body = grid.slice(1).map((row) => `| ${row.join(' | ')} |`);
  return [head, divider, ...body].join('\n');
}

function markdownFromWorkbook(source: ConversionSource): string {
  const workbook = workbookFromSource(source);
  const parts = workbook.SheetNames.map((name) => {
    const table = matrixToMarkdown(matrixFromSheet(workbook.Sheets[name]));
    return workbook.SheetNames.length > 1 ? `## ${name}\n\n${table}` : table;
  });
  return parts.join('\n\n');
}

function htmlFromWorkbook(source: ConversionSource): string {
  const workbook = workbookFromSource(source);
  return workbook.SheetNames.map((name) => {
    const html = XLSX.utils.sheet_to_html(workbook.Sheets[name]);
    return workbook.SheetNames.length > 1 ? `<h2>${escapeHtml(name)}</h2>${html}` : html;
  }).join('');
}

function htmlFromMatrix(matrix: string[][], title?: string): string {
  if (matrix.length === 0) return '<p></p>';
  const rows = matrix
    .map(
      (row, index) =>
        `<tr>${row
          .map((cell) =>
            index === 0 ? `<th>${escapeHtml(String(cell ?? ''))}</th>` : `<td>${escapeHtml(String(cell ?? ''))}</td>`
          )
          .join('')}</tr>`
    )
    .join('');
  return `${title ? `<h2>${escapeHtml(title)}</h2>` : ''}<table>${rows}</table>`;
}

export function htmlDocumentShell(title: string, bodyHtml: string): string {
  return [
    '<!doctype html><html><head><meta charset="utf-8">',
    `<title>${escapeHtml(title)}</title>`,
    '<style>',
    "body{font-family:'Times New Roman','Segoe UI',serif;font-size:12pt;line-height:1.45;color:#111;margin:0}",
    'h1{font-size:18pt}h2{font-size:15pt}h3{font-size:13pt}',
    'table{border-collapse:collapse;width:100%;margin:8pt 0}',
    'td,th{border:1px solid #888;padding:4pt 6pt;vertical-align:top;text-align:left}',
    'th{background:#eef1f5;font-weight:600}',
    'pre{white-space:pre-wrap;font-family:Consolas,monospace;font-size:10pt}',
    'blockquote{margin:8pt 0 8pt 14pt;color:#444}',
    'img{max-width:100%}',
    '</style></head><body>',
    bodyHtml || '<p></p>',
    '</body></html>',
  ].join('');
}

function xlsxBase64FromMatrix(matrix: string[][], sheetName: string): string {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(matrix.length > 0 ? matrix : [['']]);
  applyAutoFitColumns(sheet, matrix, { headerRows: 1 });
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName || 'Лист 1');
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
}

function sheetNameFor(source: ConversionSource): string {
  return source.name.replace(/\.[^./\\]+$/, '').slice(0, 31) || 'Лист 1';
}

async function toXlsx(source: ConversionSource): Promise<string> {
  const ext = source.extension.toLowerCase();
  if (source.type === 'excel') {
    if (ext === 'xlsx') return source.content;
    if (ext === 'xls' || ext === 'xlsm') {
      const workbook = workbookFromSource(source);
      for (const name of workbook.SheetNames) {
        applyAutoFitColumns(workbook.Sheets[name], matrixFromSheet(workbook.Sheets[name]), {
          headerRows: 1,
        });
      }
      return XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
    }
    const matrix = matrixFromCsv(source.content, ext === 'tsv' ? '\t' : ',');
    return xlsxBase64FromMatrix(matrix, sheetNameFor(source));
  }
  if (source.type === 'markdown') {
    const matrix = matrixFromMarkdown(source.content);
    if (!matrix) throw new Error('В Markdown не найдена таблица для сохранения в Excel');
    return xlsxBase64FromMatrix(matrix, sheetNameFor(source));
  }
  if (ext === 'html' || ext === 'htm') {
    const doc = new DOMParser().parseFromString(source.content, 'text/html');
    const table = doc.querySelector('table');
    if (!table) throw new Error('В HTML не найдена таблица для сохранения в Excel');
    const matrix = [...table.querySelectorAll('tr')].map((tr) =>
      [...tr.querySelectorAll('th,td')].map((cell) => cell.textContent?.trim() ?? '')
    );
    return xlsxBase64FromMatrix(matrix, sheetNameFor(source));
  }
  throw new Error('В Excel можно сохранить таблицу, CSV, TSV или Markdown-таблицу');
}

function toDelimited(source: ConversionSource, delimiter: ',' | '\t'): { content: string; note?: string } {
  const ext = source.extension.toLowerCase();
  if (source.type === 'excel' && source.encoding === 'base64') {
    const workbook = workbookFromSource(source);
    const name =
      source.activeSheet && workbook.SheetNames.includes(source.activeSheet)
        ? source.activeSheet
        : workbook.SheetNames[0];
    const sheet = name ? workbook.Sheets[name] : undefined;
    if (!sheet) throw new Error('В книге нет листов');
    const content = XLSX.utils.sheet_to_csv(sheet, { FS: delimiter, blankrows: false });
    return {
      content,
      note:
        workbook.SheetNames.length > 1
          ? `Экспортирован лист «${name}» из ${workbook.SheetNames.length}`
          : undefined,
    };
  }
  if (source.type === 'excel') {
    const matrix = matrixFromCsv(source.content, ext === 'tsv' ? '\t' : ',');
    const sheet = XLSX.utils.aoa_to_sheet(matrix);
    return { content: XLSX.utils.sheet_to_csv(sheet, { FS: delimiter, blankrows: false }) };
  }
  if (source.type === 'markdown') {
    const matrix = matrixFromMarkdown(source.content);
    if (!matrix) throw new Error('В Markdown не найдена таблица');
    const sheet = XLSX.utils.aoa_to_sheet(matrix);
    return { content: XLSX.utils.sheet_to_csv(sheet, { FS: delimiter, blankrows: false }) };
  }
  if (ext === 'html' || ext === 'htm') {
    const doc = new DOMParser().parseFromString(source.content, 'text/html');
    const table = doc.querySelector('table');
    if (!table) throw new Error('В HTML не найдена таблица');
    const matrix = [...table.querySelectorAll('tr')].map((tr) =>
      [...tr.querySelectorAll('th,td')].map((cell) => cell.textContent?.trim() ?? '')
    );
    const sheet = XLSX.utils.aoa_to_sheet(matrix);
    return { content: XLSX.utils.sheet_to_csv(sheet, { FS: delimiter, blankrows: false }) };
  }
  throw new Error('В CSV/TSV можно сохранить таблицу, Markdown-таблицу или HTML-таблицу');
}

function sameFormatResult(source: ConversionSource, format: ConversionFormat): ConversionResult | null {
  const ext = source.extension.toLowerCase();
  if (format === 'docx' && source.type === 'word' && (ext === 'docx' || (ext === 'doc' && isZipBase64(source.content)))) {
    return { kind: 'data', content: source.content, encoding: 'base64', extension: 'docx' };
  }
  if (format === 'xlsx' && ext === 'xlsx') {
    return { kind: 'data', content: source.content, encoding: 'base64', extension: 'xlsx' };
  }
  if (format === 'pptx' && ext === 'pptx') {
    return { kind: 'data', content: source.content, encoding: 'base64', extension: 'pptx' };
  }
  if (format === 'csv' && ext === 'csv') {
    return { kind: 'data', content: source.content, encoding: 'utf-8-bom', extension: 'csv' };
  }
  if (format === 'tsv' && ext === 'tsv') {
    return { kind: 'data', content: source.content, encoding: 'utf-8-bom', extension: 'tsv' };
  }
  if (format === 'md' && (ext === 'md' || ext === 'markdown' || ext === 'mdx')) {
    return { kind: 'data', content: source.content, encoding: 'utf-8', extension: 'md' };
  }
  if (format === 'txt' && (source.type === 'text' || source.type === 'richtext' || source.type === 'code')) {
    return { kind: 'data', content: source.content, encoding: 'utf-8', extension: 'txt' };
  }
  if (format === 'html' && (ext === 'html' || ext === 'htm')) {
    return { kind: 'data', content: source.content, encoding: 'utf-8', extension: 'html' };
  }
  return null;
}

/** Главная функция: конвертирует содержимое открытого документа. */
export async function convertDocument(
  source: ConversionSource,
  format: ConversionFormat
): Promise<ConversionResult> {
  const same = sameFormatResult(source, format);
  if (same) return same;

  if (format === 'pdf') {
    const body = await htmlFromSource(source);
    return {
      kind: 'pdf',
      html: htmlDocumentShell(source.name.replace(/\.[^.]+$/, ''), body),
      extension: 'pdf',
    };
  }

  if (format === 'xlsx') {
    return { kind: 'data', content: await toXlsx(source), encoding: 'base64', extension: 'xlsx' };
  }

  if (format === 'csv' || format === 'tsv') {
    const result = toDelimited(source, format === 'tsv' ? '\t' : ',');
    return {
      kind: 'data',
      content: result.content,
      encoding: 'utf-8-bom',
      extension: format,
      note: result.note,
    };
  }

  if (format === 'docx') {
    const html = await htmlFromSource(source);
    return {
      kind: 'data',
      content: await htmlToDocxBase64(html),
      encoding: 'base64',
      extension: 'docx',
    };
  }

  if (format === 'md') {
    return {
      kind: 'data',
      content: await markdownFromSource(source),
      encoding: 'utf-8',
      extension: 'md',
    };
  }

  if (format === 'html') {
    const body = await htmlFromSource(source);
    return {
      kind: 'data',
      content: htmlDocumentShell(source.name.replace(/\.[^.]+$/, ''), body),
      encoding: 'utf-8',
      extension: 'html',
    };
  }

  if (format === 'txt') {
    const markdown = await markdownFromSource(source);
    return {
      kind: 'data',
      content: stripMarkdown(markdown),
      encoding: 'utf-8',
      extension: 'txt',
    };
  }

  if (format === 'pptx') {
    if (source.type === 'powerpoint' && source.extension === 'pptx') {
      return { kind: 'data', content: source.content, encoding: 'base64', extension: 'pptx' };
    }
    throw new Error('PowerPoint можно сохранить из открытой презентации, но не собрать из другого формата');
  }

  throw new Error(`Конвертация в ${format} не поддерживается`);
}
