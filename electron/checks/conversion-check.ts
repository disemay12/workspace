/**
 * Проверка конвертации форматов, CSV/UTF-8 и ширины колонок.
 * Запускается в главном процессе Electron: DOM-функции здесь не вызываются,
 * проверяются чистые ветки конвертера и запись XLSX.
 */
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';
import {
  convertDocument,
  defaultTargetName,
  formatsFor,
  normalizeConversionFormat,
  type ConversionSource,
} from '../../src/services/documentConversion';
import { compareVersions } from '../../electron/updater';
import { applyAutoFitColumns, textWidthChars } from '../../src/lib/xlsxColumns';
import { parseAIActions } from '../../src/services/aiActions';

const failures: string[] = [];
let passed = 0;

function check(name: string, actual: unknown, expected: unknown) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) passed++;
  else failures.push(`${name}\n     ожидалось: ${e}\n     получено:   ${a}`);
}

const source = (overrides: Partial<ConversionSource>): ConversionSource => ({
  path: '',
  name: 'file',
  extension: 'txt',
  type: 'text',
  content: '',
  encoding: 'utf-8',
  ...overrides,
});

async function main() {
  /* ---- Матрица форматов ---- */
  check('Markdown → DOCX доступен', formatsFor({ extension: 'md', type: 'markdown' }).includes('docx'), true);
  check('CSV → XLSX доступен', formatsFor({ extension: 'csv', type: 'excel' }).includes('xlsx'), true);
  check('XLSX → CSV доступен', formatsFor({ extension: 'xlsx', type: 'excel' }).includes('csv'), true);
  check('DOCX → MD доступен', formatsFor({ extension: 'docx', type: 'word' }).includes('md'), true);
  check('PDF → DOCX недоступен', formatsFor({ extension: 'pdf', type: 'pdf' }).includes('docx'), false);
  check('формат excel → xlsx', normalizeConversionFormat('excel'), 'xlsx');
  check('формат word → docx', normalizeConversionFormat('word'), 'docx');
  check('имя отчёта', defaultTargetName('отчёт.md', 'docx'), 'отчёт.docx');

  /* ---- Markdown → текст ---- */
  const md = source({
    name: 'notes.md',
    extension: 'md',
    type: 'markdown',
    content: '# Заголовок\n\n- пункт **жирный**\n\n| A | B |\n| --- | --- |\n| 1 | 2 |',
  });
  const txt = await convertDocument(md, 'txt');
  check('Markdown → TXT: без #', txt.kind === 'data' && !txt.content.includes('# '), true);
  check('Markdown → TXT: текст сохранён', txt.kind === 'data' && txt.content.includes('пункт жирный'), true);

  /* ---- CSV → XLSX: русский текст и ширины ---- */
  const csv = source({
    name: 'данные.csv',
    extension: 'csv',
    type: 'excel',
    content: 'Товар,Описание\nЯблоки,Очень длинное описание товара\nГруши,Коротко',
  });
  const xlsx = await convertDocument(csv, 'xlsx');
  check('CSV → XLSX: создан base64', xlsx.kind === 'data' && xlsx.content.length > 100, true);
  if (xlsx.kind === 'data') {
    const workbook = XLSX.read(xlsx.content, { type: 'base64', cellStyles: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    check('CSV → XLSX: заголовок на месте', sheet?.['A1']?.v, 'Товар');
    check('CSV → XLSX: кириллица не потеряна', sheet?.['A2']?.v, 'Яблоки');
    const cols = sheet?.['!cols'] ?? [];
    check('CSV → XLSX: ширина описания', (cols[1]?.wch ?? 0) >= textWidthChars('Очень длинное описание товара'), true);
  }

  /* ---- CSV → CSV остаётся UTF-8 с BOM ---- */
  const csvSame = await convertDocument(csv, 'csv');
  check('CSV → CSV: кодировка utf-8-bom', csvSame.kind === 'data' && csvSame.encoding, 'utf-8-bom');

  /* ---- XLSX → CSV ---- */
  const xlsxSource = source({
    name: 'данные.xlsx',
    extension: 'xlsx',
    type: 'excel',
    encoding: 'base64',
    content: xlsx.kind === 'data' ? xlsx.content : '',
  });
  const csvOut = await convertDocument(xlsxSource, 'csv');
  check('XLSX → CSV: строка Товар', csvOut.kind === 'data' && csvOut.content.includes('Товар'), true);
  check('XLSX → CSV: BOM включён в кодировку', csvOut.kind === 'data' && csvOut.encoding, 'utf-8-bom');

  /* ---- Markdown-таблица → XLSX ---- */
  const mdTable = source({
    name: 'table.md',
    extension: 'md',
    type: 'markdown',
    content: '| Товар | Цена |\n| --- | --- |\n| Яблоки | 120 |',
  });
  const tableXlsx = await convertDocument(mdTable, 'xlsx');
  if (tableXlsx.kind === 'data') {
    const tableWb = XLSX.read(tableXlsx.content, { type: 'base64', cellStyles: true });
    const table = tableWb.Sheets[tableWb.SheetNames[0]];
    check('Markdown-таблица → XLSX: значение', table?.['A2']?.v, 'Яблоки');
  } else {
    check('Markdown-таблица → XLSX: создана', false, true);
  }

  /* ---- Автоширина ---- */
  const wideSheet = XLSX.utils.aoa_to_sheet([
    ['Описание'],
    ['Очень длинный текст, который должен поместиться в колонку Excel'],
  ]);
  applyAutoFitColumns(wideSheet, [
    ['Описание'],
    ['Очень длинный текст, который должен поместиться в колонку Excel'],
  ]);
  const wideZip = await JSZip.loadAsync(
    XLSX.write(
      (() => {
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, wideSheet, 'S');
        return wb;
      })(),
      { bookType: 'xlsx', type: 'base64' }
    ),
    { base64: true }
  );
  const sheetXml = (await wideZip.file('xl/worksheets/sheet1.xml')?.async('string')) ?? '';
  check('XLSX: тег cols записан', /<cols>/.test(sheetXml), true);

  /* ---- AI-действия ---- */
  const actions = parseAIActions(
    [
      '```workspace-action',
      '{"action":"save_as","format":"xlsx","path":"Отчёт.xlsx"}',
      '```',
      '```workspace-action',
      '{"action":"convert_file","source":"notes.md","target":"notes.docx","format":"docx"}',
      '```',
    ].join('\n')
  );
  check('AI: save_as распознан', actions.some((item) => item.action === 'save_as'), true);
  check('AI: convert_file распознан', actions.some((item) => item.action === 'convert_file'), true);

  /* ---- Версии обновлений ---- */
  check('1.0.12 новее 1.0.11', compareVersions('1.0.12', '1.0.11') > 0, true);
  check('v1.0.11 равна 1.0.11', compareVersions('v1.0.11', '1.0.11'), 0);
  check('1.0.9 старее 1.0.10', compareVersions('1.0.9', '1.0.10') < 0, true);

  console.log('='.repeat(72));
  console.log('  ПРОВЕРКА КОНВЕРТАЦИИ, CSV BOM И ШИРИН КОЛОНОК');
  console.log('='.repeat(72));
  console.log(`  Пройдено: ${passed}, провалов: ${failures.length}`);
  if (failures.length > 0) {
    console.log('-'.repeat(72));
    for (const failure of failures) console.log(`   - ${failure}`);
  }
  console.log('='.repeat(72));
  console.log(failures.length === 0 ? '  ИТОГ: ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ' : '  ИТОГ: ЕСТЬ ПРОВАЛЫ');

  const resultPath = path.join(process.cwd(), 'conversion-check.json');
  fs.writeFileSync(
    resultPath,
    JSON.stringify(
      {
        steps: [
          {
            name: 'conversion',
            ok: failures.length === 0,
            passed,
            failures: failures.length,
            details: failures,
          },
        ],
        errors: [],
      },
      null,
      2
    ),
    'utf-8'
  );
  process.exit(failures.length === 0 ? 0 : 1);
}

void main();
