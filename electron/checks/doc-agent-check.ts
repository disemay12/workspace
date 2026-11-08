/**
 * Проверка агентной работы ассистента с открытым .docx.
 *
 * Поднимает приложение с синтетическим резюме, подменяет OpenAI-совместимый
 * API скриптованными tool_calls и проверяет весь путь:
 * права → инструменты → живой редактор → откат → сохранение оформления.
 *
 * Запуск: electron.exe dist-electron/doc-agent-check.js
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const http = require('http');
const JSZip = require('jszip');
const XLSX = require('xlsx');
const { registerCommonCheckIpc } = require('./common-ipc');

function resolveProjectRoot() {
  let dir = __dirname;
  for (let i = 0; i < 4; i++) {
    if (fs.existsSync(path.join(dir, 'package.json')) && fs.existsSync(path.join(dir, 'dist'))) {
      return dir;
    }
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return path.join(__dirname, '..');
}

const ROOT = process.env.WS_ROOT || resolveProjectRoot();
const SANDBOX = process.env.WS_CHECK_SANDBOX || path.join(ROOT, 'test-workspace-doc-agent');
const OUT = path.join(ROOT, 'doc-agent-check.json');
const DOC_NAME = 'Резюме.docx';
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const out = { steps: [], errors: [], requests: [] };
const step = (name, value) => {
  out.steps.push({ name, ...value });
  // Промежуточная запись: при зависании видно, какой шаг последним завершился.
  try {
    fs.writeFileSync(OUT, JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
};

fs.mkdirSync(SANDBOX, { recursive: true });
fs.mkdirSync(path.join(SANDBOX, 'userdata'), { recursive: true });
app.setPath('userData', path.join(SANDBOX, 'userdata'));
app.disableHardwareAcceleration();
try {
  fs.writeFileSync(path.join(SANDBOX, 'boot.txt'), 'module loaded');
} catch {}

function writeOut(reason?: string) {
  if (reason) out.errors.push(reason);
  try {
    fs.writeFileSync(OUT, JSON.stringify(out, null, 2), 'utf-8');
  } catch {}
  console.log(JSON.stringify(out, null, 2));
}

/* ============================================================
   Синтетический .docx
   ============================================================ */

function buildDocx() {
  const paragraph = (text, options: any = {}) => {
    const props = [];
    if (options.style) props.push(`<w:pStyle w:val="${options.style}"/>`);
    if (options.list) props.push('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    if (options.align) props.push(`<w:jc w:val="${options.align}"/>`);
    const pPr = props.length ? `<w:pPr>${props.join('')}</w:pPr>` : '';
    return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  };

  const body = [
    paragraph('Иван Иванов', { style: 'Heading1' }),
    paragraph('Москва · +7 999 000-00-00 · ivan@example.com'),
    paragraph('Опыт работы', { style: 'Heading2' }),
    // Фраза разбита на runs, как после ручного форматирования в Word:
    // точечная замена должна найти её, не переписывая документ целиком.
    '<w:p><w:r><w:t xml:space="preserve">Компания А — инженер, </w:t></w:r>' +
      '<w:r><w:t xml:space="preserve">Июль </w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">2024</w:t></w:r>' +
      '<w:r><w:t xml:space="preserve"> — настоящее время</w:t></w:r></w:p>',
    paragraph('Навыки', { style: 'Heading2' }),
    paragraph('TypeScript', { list: true }),
    paragraph('React', { list: true }),
    paragraph('К сожалению, я не могу редактировать документ Word. Вот рекомендации по оформлению: используйте заголовки и списки.'),
    paragraph('Хотите, я создам отдельный файл с исправленным текстом?'),
  ].join('');

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

  const stylesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>' +
    '</w:styles>';

  const numberingXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
    '<w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

  const zip = new JSZip();
  zip.file('word/document.xml', documentXml);
  zip.file('word/styles.xml', stylesXml);
  zip.file('word/numbering.xml', numberingXml);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** Текст .docx без сторонних зависимостей — для проверки сохранения. */
async function readDocxXml(filePath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
  const xml = await zip.file('word/document.xml').async('string');
  const stylesEntry = zip.file('word/styles.xml');
  const styles = stylesEntry ? await stylesEntry.async('string') : '';
  const text = (xml.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [])
    .map((match) => match.replace(/<[^>]+>/g, ''))
    .join('');
  return { xml, styles, text };
}

/** Все XML-части пакета — для проверки колонтитулов и rels. */
async function readDocxParts(filePath): Promise<Record<string, string>> {
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
  const parts: Record<string, string> = {};
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir || !name.endsWith('.xml')) continue;
    parts[name] = await zip.files[name].async('string');
  }
  return parts;
}

/* ============================================================
   Мок OpenAI-совместимого API с tool_calls
   ============================================================ */

let IMAGE_URL = '';

/** 1×1 PNG для проверки загрузки изображения по URL. */
const URL_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lK3hWQAAAABJRU5ErkJggg==';

const LONG_HTML = Array.from(
  { length: 200 },
  (_, index) =>
    `<p>Блок ${index}: строка длинного документа для проверки масштабирования.</p>`
).join('');

const TABLE_HTML =
  '<h2>Навыки</h2><table><tbody><tr><th>Навык</th><th>Уровень</th></tr>' +
  '<tr><td>JavaScript</td><td>5</td></tr><tr><td>TypeScript</td><td>4</td></tr></tbody></table>';

const ATS_HTML = [
  '<h1>Иван Иванов</h1>',
  '<p style="text-align:left">Москва · +7 999 000-00-00 · ivan@example.com</p>',
  '<h3>Опыт работы</h3>',
  '<p><strong>Компания А</strong> — инженер, Июль 2024 — настоящее время</p>',
  '<h3>Навыки</h3>',
  '<ul><li>TypeScript</li><li>React</li></ul>',
  '<h3>Образование</h3>',
  '<p>МГУ, 2018</p>',
].join('');

const FALLBACK_HTML = [
  '<h1>Иван Иванов</h1>',
  '<p style="text-align:center">Москва · +7 999 000-00-00 · ivan@example.com</p>',
  '<h2>Опыт работы</h2>',
  '<p>Компания А — инженер, Июнь 2024 — настоящее время. Оформлено через document_apply.</p>',
  '<h2>Навыки</h2>',
  '<ul><li>TypeScript</li><li>React</li></ul>',
].join('');

/** Резюме, переписанное под ATS, с намеренно оставленными отчётами модели. */
const ATS_REWRITE_HTML = [
  '<h1>Иван Сидоров</h1>',
  '<p>Оформил документ в деловом стиле и заменил профессию. Вот итоговая версия:</p>',
  '<h2>Опыт работы</h2>',
  '<p>Компания А — ведущий инженер, июль 2024 — настоящее время.</p>',
  '<h2>Навыки</h2>',
  '<ul><li>TypeScript</li><li>React</li><li>Диагностика оборудования</li></ul>',
  '<h2>Образование</h2>',
  '<p>МГУ, 2018</p>',
  '<p>Документ Test.docx обновлён. Он по-прежнему помечен как несохранённый — сохранить сейчас?</p>',
].join('');

const TABLE_ANSWER = [
  '| Навык | Уровень |',
  '| --- | --- |',
  '| TypeScript | высокий |',
  '| React | средний |',
].join('\n');

/** Ответ «напиши сказку» без вызова инструментов — проверка принудительной вставки. */
const STORY_TEXT = [
  'Сказка о забытом фонаре',
  '',
  'В маленьком городке на краю света жил старый фонарщик Тим. Каждый вечер он зажигал фонари на главной улице — один за другим, не спеша.',
  '',
  'Но был среди них один — самый последний фонарь на площади. Он стоял в стороне от других, и его стекло помутнело от времени.',
  '',
  'Однажды в городок пришла девочка по имени Айя. Она спросила: «Почему не горит тот фонарь?» Взрослые пожимали плечами, а Тим улыбнулся и ответил: «Ему не нужен мой огонь. Ему нужен свой».',
].join('\n');

function toolCall(name, args, id, index = 0) {
  return {
    choices: [
      {
        delta: {
          tool_calls: [
            {
              index,
              id,
              type: 'function',
              function: { name, arguments: JSON.stringify(args || {}) },
            },
          ],
        },
      },
    ],
  };
}

function contentChunk(text) {
  return { choices: [{ delta: { content: text } }] };
}

function scenarioOf(messages) {
  const users = messages.filter((message) => message.role === 'user');
  let text = String(users[users.length - 1]?.content || '');
  // Повтор после текстового ответа: сам запрос — предыдущее сообщение.
  if (/не внёс изменения|Ты не внёс текст/.test(text) && users.length > 1) {
    text = String(users[users.length - 2]?.content || '');
  }
  if (text.includes('Измени заголовок на «Иван Петров»')) return 'quickTitle';
  if (text.includes('Замени строку с должностью')) return 'blockedLine';
  if (text.includes('Замени выделенное')) return 'selection';
  if (text.includes('второй файл')) return 'multidoc';
  if (text.includes('верхний колонтитул')) return 'headerfooter';
  if (text.includes('Заполни A4')) return 'excelWrite';
  if (text.includes('Добавь лист Итоги')) return 'excelSheet';
  if (text.includes('Объедини A1:B1')) return 'excelOps';
  if (text.includes('Что сейчас выделено в таблице')) return 'excelSelection';
  if (text.includes('Замени текст на первом слайде')) return 'pptxText';
  if (text.includes('Добавь слайд «Итоги»')) return 'pptxAdd';
  if (text.includes('Добавь текст на первый слайд')) return 'pptxTextbox';
if (text.includes('Измени заметки')) return 'pptxNotes';
if (text.includes('Вставь картинку на слайд')) return 'pptxImage';
if (text.includes('Смени макет')) return 'pptxLayout';
  if (text.includes('Что выделено на слайде')) return 'pptxSelection';
  if (text.includes('Запиши в третий файл')) return 'writeDoc';
if (text.includes('Запиши книгу')) return 'writeBook';
if (text.includes('Сделай таблицу навыков')) return 'tableOps';
if (text.includes('Длинный документ')) return 'longDoc';
if (text.includes('Создай стиль')) return 'customStyle';
if (text.includes('Добавь комментарий и сноску')) return 'annotations';
  if (text.includes('Вставь картинку по URL')) return 'imageUrl';
if (text.includes('Оформи Markdown')) return 'mdStructure';
  if (text.includes('Замени в Markdown')) return 'mdReplace';
  if (text.includes('атомарность')) return 'atomic';
  if (text.includes('Сделай выделенное жирным')) return 'selRestore';
  if (text.includes('Поменяй строку с опытом')) return 'preciseRetry';
  if (text.includes('ATC-фильтры')) return 'atsRewrite';
  if (text.includes('Перепиши документ полностью')) return 'noopRewrite';
  if (text.includes('Покажи таблицу навыков')) return 'tableAnswer';
  if (text.includes('мини-сказку')) return 'story';
  if (text.includes('запасной протокол')) return 'fallback';
  if (text.includes('Перепиши резюме')) return 'norights';
  if (text.includes('как реальный документ')) return 'pages';
  if (text.includes('заявление на отпуск')) return 'statement';
  if (/абзац[а-яё]*\s+про\s+отпуск/i.test(text)) return 'insert';
  if (text.includes('не относится к резюме')) return 'cleanup';
  if (text.includes('ATS')) return 'ats';
  if (text.includes('H2') || text.includes('по центру')) return 'format';
  if (text.includes('Июль 2024')) return 'find';
  return 'default';
}

function lastStructure(messages) {
  const tool = [...messages].reverse().find((message) => message.role === 'tool');
  if (!tool) return [];
  try {
    const data = JSON.parse(String(tool.content));
    return String(data.structure || '')
      .split('\n')
      .map((line) => {
        const match = /^\s*(b\d+)\s+\[([^\]]+)\]\s*(.*)$/.exec(line);
        return match ? { id: match[1], tag: match[2], text: match[3] } : null;
      })
      .filter(Boolean);
  } catch {
    return [];
  }
}

function scriptedResponse(scenario, stepIndex, messages) {
  const lines = lastStructure(messages);
  const lastMessage = messages[messages.length - 1];
  const lastRole = String(lastMessage?.role || '');
  const lastUser = [...messages].reverse().find((message) => message.role === 'user');
  const lastUserText = String(lastUser?.content || '');

  if (scenario === 'story') {
    return [
      contentChunk(
        STORY_TEXT +
          '\n\nГотово — сказка записана в Test.docx. Если хочешь, могу изменить имена героев, сделать версию для детей помладше или, наоборот, добавить главы и превратить её в целую книгу. Что скажешь?'
      ),
    ];
  }

  if (scenario === 'quickTitle') {
    return [toolCall('find_replace', { find: 'Иван Иванов', replace: 'Иван Петров' }, 'q-title')];
  }

  if (scenario === 'selection') {
    const toolMessages = messages.filter((message) => message.role === 'tool');
    if (toolMessages.length === 0) {
      return [toolCall('get_selection', {}, 'sel-get')];
    }
    if (
      !toolMessages.some((message) =>
        String(message.content || '').includes('Выделенный фрагмент заменён')
      )
    ) {
      return [
        toolCall('replace_selection', { html: 'Санкт-Петербург' }, 'sel-replace'),
      ];
    }
    return [contentChunk('Заменил выделенный фрагмент.')];
  }

  if (scenario === 'multidoc') {
    if (lastRole === 'user') {
      return [toolCall('read_document', { path: 'Второй.md' }, 'md-read')];
    }
    const toolContent = String(messages[messages.length - 1]?.content || '');
    return [
      contentChunk(
        toolContent.includes('Секретный текст второго файла')
          ? 'Прочитал второй файл: секретный текст найден.'
          : 'Не удалось прочитать второй файл.'
      ),
    ];
  }

  if (scenario === 'headerfooter') {
    const toolMessages = messages.filter((message) => message.role === 'tool');
    if (toolMessages.length === 0) {
      return [toolCall('get_document_info', {}, 'hf-info')];
    }
    if (toolMessages.length === 1) {
      return [
        toolCall(
          'set_header_footer',
          { kind: 'header', type: 'default', text: 'Шапка отдела' },
          'hf-set'
        ),
      ];
    }
    return [contentChunk('Обновил верхний колонтитул.')];
  }

  if (scenario === 'excelWrite') {
    const toolMessages = messages.filter((message) => message.role === 'tool');
    if (toolMessages.length === 0) {
      return [toolCall('excel_read', { range: 'A1:B3' }, 'xl-read')];
    }
    if (toolMessages.length === 1) {
      const readResult = String(toolMessages[0].content || '');
      if (!readResult.includes('Яблоко')) {
        return [contentChunk('Не удалось прочитать таблицу.')];
      }
      return [
        toolCall(
          'excel_write_cells',
          {
            cells: [
              { address: 'A4', value: 'Апельсин' },
              { address: 'B4', value: '=B2+B3' },
            ],
          },
          'xl-write'
        ),
      ];
    }
    return [contentChunk('Заполнил A4:B4.')];
  }

  if (scenario === 'excelSheet') {
    if (lastRole === 'user') {
      return [toolCall('excel_add_sheet', { name: 'Итоги' }, 'xl-sheet')];
    }
    return [contentChunk('Добавил лист Итоги.')];
  }

  if (scenario === 'excelOps') {
    const toolMessages = messages.filter((message) => message.role === 'tool');
    if (toolMessages.length === 0) {
      return [
        toolCall(
          'excel_format_range',
          { sheet: 'Данные', range: 'A1:B1', bold: true, fill: '#FFFF00' },
          'xl-format'
        ),
      ];
    }
    if (toolMessages.length === 1) {
      return [
        toolCall('excel_merge_cells', { sheet: 'Данные', range: 'A1:B1' }, 'xl-merge'),
      ];
    }
    return [contentChunk('Выделил и объединил A1:B1.')];
  }

  if (scenario === 'excelSelection') {
    if (lastRole === 'user') {
      return [toolCall('excel_get_selection', {}, 'xl-selection')];
    }
    const toolContent = String(messages[messages.length - 1]?.content || '');
    return [
      contentChunk(
        toolContent.includes('Выделено')
          ? 'Выделение получено: диапазон на активном листе.'
          : 'Выделение не получено.'
      ),
    ];
  }

  if (scenario === 'tableOps') {
    if (stepIndex === 0) {
      return [toolCall('replace_document', { html: TABLE_HTML }, 'tb-replace')];
    }
    if (stepIndex === 1) {
      return [toolCall('get_document', {}, 'tb-get')];
    }
    const toolMessages = messages.filter((message) => message.role === 'tool');
    const structure = String(toolMessages[toolMessages.length - 1]?.content || '');
    const tableId = /b\d+(?=\s*\[table)/.exec(structure)?.[0] ?? '';
    if (stepIndex === 2 && tableId) {
      return [
        toolCall(
          'table_set_column_width',
          { block_id: tableId, col: 0, width_pt: 100 },
          'tb-width'
        ),
      ];
    }
    if (stepIndex === 3 && tableId) {
      return [
        toolCall(
          'table_set_borders',
          { block_id: tableId, color: '#FF0000', size_pt: 2, style: 'single' },
          'tb-border'
        ),
      ];
    }
    return [contentChunk('Таблица навыков оформлена.')];
  }

  if (scenario === 'longDoc') {
    if (stepIndex === 0) {
      return [toolCall('replace_document', { html: LONG_HTML }, 'ld-replace')];
    }
    if (stepIndex === 1) {
      return [
        toolCall(
          'find_replace',
          { find: 'Блок 199:', replace: 'Блок 199 изменён:', all: false },
          'ld-edit'
        ),
      ];
    }
    return [contentChunk('Длинный документ обновлён.')];
  }

  if (scenario === 'customStyle') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'cs-get')];
    const toolMessages = messages.filter((message) => message.role === 'tool');
    const structure = String(toolMessages[toolMessages.length - 1]?.content || '');
    const blockId = /b\d+/.exec(structure)?.[0] ?? '';
    if (stepIndex === 1 && blockId) {
      return [
        toolCall(
          'define_style',
          {
            block_id: blockId,
            name: 'Врезка AI',
            style_id: 'CustomAI',
            font: 'Georgia',
            size_pt: 14,
            bold: true,
            color: '#112233',
            align: 'center',
          },
          'cs-define'
        ),
      ];
    }
    return [contentChunk('Стиль «Врезка AI» создан.')];
  }

  if (scenario === 'annotations') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'an-get')];
    const toolMessages = messages.filter((message) => message.role === 'tool');
    const structure = String(toolMessages[toolMessages.length - 1]?.content || '');
    const blockId = /b\d+/.exec(structure)?.[0] ?? '';
    if (stepIndex === 1 && blockId) {
      return [
        toolCall(
          'add_comment',
          { block_id: blockId, text: 'Проверить формулировку', author: 'AI' },
          'an-comment'
        ),
      ];
    }
    if (stepIndex === 2 && blockId) {
      return [
        toolCall(
          'add_footnote',
          { block_id: blockId, text: 'Источник для блока' },
          'an-footnote'
        ),
      ];
    }
    return [contentChunk('Комментарий и сноска добавлены.')];
  }

  if (scenario === 'pptxText') {
    const toolMessages = messages.filter((message) => message.role === 'tool');
    if (toolMessages.length === 0) {
      return [toolCall('pptx_read', {}, 'pp-read')];
    }
    if (toolMessages.length === 1) {
      const readResult = String(toolMessages[0].content || '');
      if (!readResult.includes('Заметка докладчика')) {
        return [contentChunk('Не удалось прочитать заметки слайда.')];
      }
      return [
        toolCall(
          'pptx_set_text',
          { slide: 1, shape_id: '2', text: 'Новый заголовок' },
          'pp-text'
        ),
      ];
    }
    return [contentChunk('Заменил заголовок первого слайда.')];
  }

  if (scenario === 'pptxAdd') {
    if (lastRole === 'user') {
      return [toolCall('pptx_add_slide', { index: 1, title: 'Итоги' }, 'pp-add')];
    }
    return [contentChunk('Добавил слайд «Итоги».')];
  }

  if (scenario === 'pptxTextbox') {
    if (lastRole === 'user') {
      return [
        toolCall(
          'pptx_add_text_box',
          { slide: 1, text: 'Добавленный текст' },
          'pp-textbox'
        ),
      ];
    }
    return [contentChunk('Добавил текст на первый слайд.')];
  }

  if (scenario === 'pptxNotes') {
    if (lastRole === 'user') {
      return [
        toolCall(
          'pptx_set_notes',
          { slide: 1, text: 'Тезисы для доклада' },
          'pp-notes'
        ),
      ];
    }
    return [contentChunk('Заметки обновлены.')];
  }

  if (scenario === 'pptxImage') {
    if (lastRole === 'user') {
      return [
        toolCall(
          'pptx_insert_image',
          { slide: 1, src: IMAGE_URL, width: 120, height: 80, alt: 'PPT image' },
          'pp-image'
        ),
      ];
    }
    return [contentChunk('Картинка добавлена на слайд.')];
  }

  if (scenario === 'pptxLayout') {
    if (lastRole === 'user') {
      return [
        toolCall(
          'pptx_set_layout',
          { slide: 1, layout: 'Раздел' },
          'pp-layout'
        ),
      ];
    }
    return [contentChunk('Макет слайда изменён.')];
  }

  if (scenario === 'pptxSelection') {
    if (lastRole === 'user') {
      return [toolCall('pptx_get_selection', {}, 'pp-selection')];
    }
    const toolContent = String(messages[messages.length - 1]?.content || '');
    return [
      contentChunk(
        toolContent.includes('Новый заголовок')
          ? 'Выделена фигура заголовка.'
          : 'Выделенная фигура не получена.'
      ),
    ];
  }

  if (scenario === 'writeDoc') {
    if (lastRole === 'user') {
      return [
        toolCall(
          'write_document',
          { path: 'Третий.md', text: '# Изменено AI\n\nНовый текст.\n' },
          'wd-write'
        ),
      ];
    }
    return [contentChunk('Файл обновлён.')];
  }

  if (scenario === 'writeBook') {
    if (lastRole === 'user') {
      return [
        toolCall(
          'write_document',
          {
            path: 'Третья.xlsx',
            columns: ['Поле', 'Значение'],
            rows: [['A', '1'], ['B', '2']],
            sheet: 'Данные',
          },
          'wb-write'
        ),
      ];
    }
    return [contentChunk('Книга записана.')];
  }

  if (scenario === 'imageUrl') {
    if (lastRole === 'user') {
      return [
        toolCall(
          'insert_image',
          { src: IMAGE_URL, width: 40, height: 40, alt: 'URL image' },
          'iu-insert'
        ),
      ];
    }
    return [contentChunk('Изображение по URL вставлено.')];
  }

  if (scenario === 'mdStructure') {
    if (stepIndex === 0) {
      return [toolCall('md_set_heading', { find: 'Открытый текст', level: '2' }, 'md-heading')];
    }
    if (stepIndex === 1) {
      return [
        toolCall(
          'md_insert_link',
          { find: 'второго файла', text: 'второго файла', url: 'https://example.com' },
          'md-link'
        ),
      ];
    }
    if (stepIndex === 2) {
      return [
        toolCall(
          'md_insert_table',
          { headers: ['Поле', 'Значение'], rows: [['A', '1']] },
          'md-table'
        ),
      ];
    }
    return [contentChunk('Markdown оформлен.')];
  }

  if (scenario === 'mdReplace') {
    const toolMessages = messages.filter((message) => message.role === 'tool');
    if (toolMessages.length === 0) {
      return [toolCall('md_read', {}, 'md-read')];
    }
    if (toolMessages.length === 1) {
      return [
        toolCall(
          'md_replace',
          { find: 'Секретный', replace: 'Открытый', all: true },
          'md-replace'
        ),
      ];
    }
    return [contentChunk('Заменил слово в Markdown.')];
  }

  if (scenario === 'atomic') {
    const toolMessages = messages.filter((message) => message.role === 'tool');
    if (toolMessages.length === 0) {
      return [
        toolCall(
          'find_replace',
          { find: 'ТестПоиска', replace: 'АТОМНАЯ_ПРАВКА' },
          'at-good',
          0
        ),
        toolCall('delete_block', { block_id: 'b999' }, 'at-bad', 1),
      ];
    }
    return [contentChunk('Правка отменена из-за ошибки.')];
  }

  if (scenario === 'selRestore') {
    const toolMessages = messages.filter((message) => message.role === 'tool');
    if (toolMessages.length === 0) {
      return [toolCall('get_document', {}, 'sr-get')];
    }
    if (toolMessages.length === 1) {
      const raw = String(toolMessages[0].content || '');
      let structure = raw;
      try {
        structure = String(JSON.parse(raw)?.structure ?? raw);
      } catch {
        /* оставляем raw */
      }
      const line = structure.split('\n').find((item) => item.includes('ТестПоиска'));
      const blockId = /^b\d+/.exec(line ?? '')?.[0] ?? 'b1';
      return [
        toolCall('format', { block_id: blockId, bold: true }, 'sr-format'),
      ];
    }
    return [contentChunk('Выделение сохранено.')];
  }

  if (scenario === 'blockedLine') {
    if (stepIndex === 0) {
      return [
        toolCall(
          'replace_document',
          { html: '<h1>Иван Петров</h1><p>Полная замена документа.</p>' },
          'b-bad'
        ),
      ];
    }
    if (stepIndex === 1) {
      return [
        toolCall('find_replace', { find: 'инженер', replace: 'ведущий инженер' }, 'b-good'),
      ];
    }
    return [contentChunk('Уточнил должность в опыте работы.')];
  }

  if (scenario === 'preciseRetry') {
    if (lastUserText.includes('не внёс') && lastRole === 'user') {
      return [
        toolCall('find_replace', { find: 'Навыки', replace: 'Ключевые навыки' }, 'p-good'),
      ];
    }
    if (lastRole === 'user') {
      return [
        contentChunk(
          'Понял задачу — нужно изменить строку с опытом работы. Если нужно поменять ещё что-то — скажи.'
        ),
      ];
    }
    return [contentChunk('Изменил строку с опытом работы.')];
  }

  if (scenario === 'atsRewrite') {
    if (lastUserText.includes('не внёс') && lastRole === 'user') {
      return [toolCall('replace_document', { html: ATS_REWRITE_HTML }, 'ats-replace')];
    }
    if (lastRole === 'user') {
      return [
        contentChunk(
          'Вот текущая версия резюме — она уже должна проходить ATS-фильтры. Если нужно, могу что-то поправить.'
        ),
      ];
    }
    return [contentChunk('Переписал резюме под ATS: одна колонка, стандартные заголовки, списки.')];
  }

  if (scenario === 'noopRewrite') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'n-get')];
    if (stepIndex === 1) {
      const sameHtml = lines
        .filter((line) => line.tag !== 'hr')
        .map((line) => `<p>${line.text}</p>`)
        .join('');
      return [
        toolCall(
          'replace_document',
          { html: sameHtml || '<p></p>' },
          'n-bad'
        ),
      ];
    }
    return [contentChunk('Содержимое не изменилось — правка не применена.')];
  }

  if (scenario === 'tableAnswer') {
    return [contentChunk(TABLE_ANSWER)];
  }

  if (scenario === 'cleanup') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'c-get')];
    if (stepIndex === 1) {
      return [
        toolCall('delete_block', { block_id: 'b9' }, 'c-del-1', 0),
        toolCall('delete_block', { block_id: 'b10' }, 'c-del-2', 1),
      ];
    }
    return [contentChunk('Убрал из документа всё, что не относится к резюме.')];
  }

  if (scenario === 'ats') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'a-get')];
    if (stepIndex === 1) {
      return [toolCall('replace_document', { html: ATS_HTML }, 'a-replace')];
    }
    return [contentChunk('Переписал резюме под ATS: одна колонка, стандартные заголовки, списки.')];
  }

  if (scenario === 'format') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'f-get')];
    if (stepIndex === 1) {
      const calls = [];
      for (const line of lines) {
        if (!/^h[1-6]$/.test(line.tag)) continue;
        if (!/(Опыт|Навыки|Образование)/i.test(line.text)) continue;
        calls.push(
          toolCall('format', { block_id: line.id, heading: '2' }, `f-h-${line.id}`, calls.length)
        );
      }
      const contact = lines.find((line) => line.tag === 'p' && line.text.includes('@'));
      if (contact) {
        calls.push(
          toolCall('format', { block_id: contact.id, align: 'center' }, 'f-center', calls.length)
        );
      }
      return calls.length > 0 ? calls : [contentChunk('Не нашёл заголовков для оформления.')];
    }
    return [contentChunk('Сделал заголовки разделов H2 и выровнял контакты по центру.')];
  }

  if (scenario === 'find') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'r-get')];
    if (stepIndex === 1) {
      return [
        toolCall('find_replace', { find: 'Июль   2024', replace: 'Июнь 2024', all: true }, 'r-one'),
      ];
    }
    return [contentChunk('Заменил «Июль 2024» на «Июнь 2024» в опыте работы.')];
  }

  if (scenario === 'pages') {
    if (stepIndex === 0) return [toolCall('get_document', {}, 'p-get')];
    if (stepIndex === 1) {
      return [toolCall('page_setup', { template: 'russian_business' }, 'p-setup')];
    }
    return [contentChunk('Оформил как деловой документ: поля 3/1,5/2/2 см, Times New Roman 14.')];
  }

  if (scenario === 'statement') {
    return [
      contentChunk(
        [
          'Понял Вас. Ниже — готовый текст заявления.',
          '',
          'Рекомендуемые настройки страницы и шрифта (для Word):',
          '- Поля: левое 3 см, правое 1,5 см, верхнее и нижнее 2 см',
          '- Шрифт: Times New Roman, 14 пт',
          '',
          '```',
          'Директору ООО «Ромашка»',
          'Иванову И.И.',
          '',
          'ЗАЯВЛЕНИЕ',
          '',
          'Прошу предоставить отпуск с 1 декабря 2026 года на 14 календарных дней.',
          '',
          'Дата: 07.11.2026',
          '```',
        ].join('\n')
      ),
    ];
  }

  if (scenario === 'insert') {
    return [
      contentChunk(
        [
          'Дополнительный абзац: прошу согласовать отпуск с 1 декабря 2026 года.',
          '',
          'Готово — абзац добавлен в Test.docx. Если хочешь, могу изменить формулировку или добавить подпись. Что скажешь?',
        ].join('\n')
      ),
    ];
  }

  if (scenario === 'norights') {
    return [
      contentChunk(
        'Прав на запись в документ нет — изменить его не могу. Нажмите «Предоставить права на запись» в панели AI.'
      ),
    ];
  }

  if (scenario === 'fallback') {
    const action = {
      action: 'document_apply',
      operations: [{ tool: 'replace_document', html: FALLBACK_HTML }],
    };
    return [
      contentChunk(
        'Оформляю документ через запасной протокол.\n\n```workspace-action\n' +
          JSON.stringify(action) +
          '\n```'
      ),
    ];
  }

  return [contentChunk('Готово.')];
}

function startMockServer(): Promise<any> {
  return new Promise((resolve) => {
    const counters = new Map();

    const server = http.createServer((req, res) => {
      let body = '';
      req.on('data', (chunk) => (body += chunk));
      req.on('end', () => {
        let parsed = null;
        try {
          parsed = JSON.parse(body);
        } catch {}

        const messages = Array.isArray(parsed?.messages) ? parsed.messages : [];
        const last = messages[messages.length - 1];
        const scenario = scenarioOf(messages);

        // Новый пользовательский запрос начинает сценарий заново —
        // повторный прогон после отката снова проходит все шаги.
        if (!counters.has(scenario)) counters.set(scenario, { step: 0, lastRole: '' });
        const state = counters.get(scenario);
        if (last?.role === 'user' && state.lastRole !== 'user') {
          state.step = 0;
        }
        const stepIndex = state.step;
        state.step += 1;
        state.lastRole = last?.role || '';

        const systemPrompt = String(messages[0]?.content || '');
        out.requests.push({
          scenario,
          stepIndex,
          lastRole: last?.role || '',
          model: parsed?.model ?? null,
          stream: parsed?.stream ?? null,
          toolNames: Array.isArray(parsed?.tools)
            ? parsed.tools.map((tool) => tool?.function?.name).filter(Boolean)
            : [],
          lastUserContent: String(
            [...messages].reverse().find((message) => message.role === 'user')?.content || ''
          ).slice(0, 500),
          promptHasWrite: systemPrompt.includes('Права записи: ЕСТЬ'),
          promptHasNoWrite: systemPrompt.includes('Права записи: НЕТ'),
          promptMentionsDocument: systemPrompt.includes('ОТКРЫТЫЙ ДОКУМЕНТ'),
          promptFallback: systemPrompt.includes('Function calling недоступен'),
          toolContents: messages
            .filter((message) => message.role === 'tool')
            .slice(-6)
            .map((message) => String(message.content || '').slice(0, 300)),
        });

        // Первая попытка fallback-сценария: модель «не поддерживает» tools.
        if (scenario === 'fallback' && Array.isArray(parsed?.tools) && parsed.tools.length > 0) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: { message: 'tools are not supported by this model' } }));
          return;
        }

        const payloads = scriptedResponse(scenario, stepIndex, messages);

        res.writeHead(200, {
          'Content-Type': 'text/event-stream',
          'Cache-Control': 'no-cache',
          Connection: 'keep-alive',
        });
        for (const payload of payloads) {
          res.write(`data: ${JSON.stringify(payload)}\n\n`);
        }
        res.write('data: [DONE]\n\n');
        res.end();
      });
    });

    server.listen(0, '127.0.0.1', () => {
      resolve({ server, baseUrl: `http://127.0.0.1:${server.address().port}/v1` });
    });
  });
}

/* ============================================================
   IPC песочницы
   ============================================================ */

const seededAiSettings = { settings: {}, keys: {} };

function registerHandlers() {
  registerCommonCheckIpc({
    sandbox: SANDBOX,
    getAiSettings: () => seededAiSettings.settings,
    getAiKeys: () => seededAiSettings.keys,
  });

  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('fs:openFolderDialog', () => SANDBOX);
  ipcMain.handle('fs:openFileDialog', () => null);
  ipcMain.handle('fs:saveFileDialog', () => null);
  ipcMain.handle(
    'pdf:export',
    async (_event, html, _suggestedName, targetPath) => {
      if (!targetPath) return { success: false, canceled: true };
      const pdfWindow = new BrowserWindow({ show: false });
      try {
        await pdfWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
        const pdf = await pdfWindow.webContents.printToPDF({
          printBackground: true,
          pageSize: 'A4',
        });
        fs.writeFileSync(targetPath, pdf);
        return { success: true, path: targetPath };
      } catch (err) {
        return { success: false, error: String(err) };
      } finally {
        pdfWindow.destroy();
      }
    }
  );
  ipcMain.handle('fs:showInExplorer', () => ({ success: true }));
  ipcMain.handle('window:isMaximized', () => false);
  ipcMain.handle('window:minimize', () => undefined);
  ipcMain.handle('window:maximize', () => false);
  ipcMain.handle('window:close', () => undefined);
  ipcMain.handle('system:getInitialExternalPath', () => null);

  ipcMain.handle('fs:listDirectory', async (_event, dirPath) => {
    try {
      const entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
      const items = entries.map((entry) => {
        const full = path.join(dirPath, entry.name);
        let size = 0;
        try {
          size = fs.statSync(full).size;
        } catch {}
        return {
          name: entry.name,
          path: full,
          isDirectory: entry.isDirectory(),
          size,
          mtime: null,
          extension: entry.isDirectory()
            ? ''
            : path.extname(entry.name).toLowerCase().replace('.', ''),
        };
      });
      items.sort((a, b) =>
        a.isDirectory !== b.isDirectory
          ? a.isDirectory
            ? -1
            : 1
          : a.name.localeCompare(b.name)
      );
      return { success: true, items };
    } catch (err) {
      return { success: false, error: err.message, items: [] };
    }
  });

  ipcMain.handle('fs:readFile', async (_event, filePath, encoding = 'utf-8') => {
    try {
      const buffer = await fs.promises.readFile(filePath);
      if (encoding === 'base64') {
        return { success: true, content: buffer.toString('base64'), isBinary: true };
      }
      return { success: true, content: buffer.toString('utf-8'), isBinary: false };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:writeFile', async (_event, filePath, content, encoding = 'utf-8') => {
    try {
      if (encoding === 'base64') {
        await fs.promises.writeFile(filePath, Buffer.from(content, 'base64'));
      } else {
        await fs.promises.writeFile(filePath, content, 'utf-8');
      }
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:createFile', async (_event, filePath, initialContent = '') => {
    try {
      await fs.promises.writeFile(filePath, initialContent, { flag: 'wx' });
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });
  ipcMain.handle('fs:createDirectory', async (_event, dirPath) => {
    try {
      await fs.promises.mkdir(dirPath, { recursive: true });
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async (_event, oldPath, newPath) => {
    try {
      await fs.promises.rename(oldPath, newPath);
      return { success: true };
    } catch (error) {
      return { success: false, error: (error as Error).message };
    }
  });
  ipcMain.handle('fs:searchWorkspace', async () => ({ success: true, results: [] }));
}

/* ============================================================
   Сценарий
   ============================================================ */

/** Минимальная презентация: заголовок и тело на одном слайде. */
async function writePptxFixture(filePath) {
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const zip = new JSZip();

  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Types xmlns="${CT}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
      '<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
      '</Types>'
  );
  zip.folder('_rels').file(
    '.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${REL}">` +
      `<Relationship Id="rId1" Type="${R}/officeDocument" Target="ppt/presentation.xml"/>` +
      '</Relationships>'
  );
  zip.folder('ppt').file(
    'presentation.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<p:presentation xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">` +
      '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
      '<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>' +
      '</p:presentation>'
  );
  zip.folder('ppt').folder('_rels').file(
    'presentation.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${REL}">` +
      `<Relationship Id="rId2" Type="${R}/slide" Target="slides/slide1.xml"/>` +
      '</Relationships>'
  );
  zip.folder('ppt').folder('slides').file(
    'slide1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">` +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="609600" y="365760"/><a:ext cx="10972800" cy="1371600"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ru-RU" sz="3200"/><a:t>Старый заголовок</a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Text 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="609600" y="1828800"/><a:ext cx="10972800" cy="3657600"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ru-RU" sz="1800"/><a:t>Основной текст слайда</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
  );
  zip.folder('ppt').folder('slides').folder('_rels').file(
    'slide1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${REL}">` +
      `<Relationship Id="rId1" Type="${R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId2" Type="${R}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>` +
      '</Relationships>'
  );
  zip.folder('ppt').folder('slideLayouts').file(
    'slideLayout1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
  );
  zip.folder('ppt').folder('slideLayouts').file(
    'slideLayout2.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" type="titleOnly"><p:cSld name="Раздел"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
  );

  zip.folder('ppt').folder('notesSlides').file(
    'notesSlide1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<p:notes xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ru-RU"/><a:t>Заметка докладчика</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld></p:notes>'
  );

  fs.writeFileSync(filePath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

async function main() {
  fs.rmSync(path.join(SANDBOX, DOC_NAME), { force: true });
  fs.writeFileSync(path.join(SANDBOX, DOC_NAME), await buildDocx());
  fs.writeFileSync(
    path.join(SANDBOX, 'Второй.md'),
    '# Второй документ\n\nСекретный текст второго файла.\n',
    'utf-8'
  );
  fs.writeFileSync(
    path.join(SANDBOX, 'Третий.md'),
    '# Третий документ\n\nИсходный текст.\n',
    'utf-8'
  );

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ['Товар', 'Цена'],
      ['Яблоко', 100],
      ['Груша', 150],
    ]),
    'Данные'
  );
  fs.writeFileSync(
    path.join(SANDBOX, 'Книга.xlsx'),
    XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
  );

  await writePptxFixture(path.join(SANDBOX, 'Презентация.pptx'));

  registerHandlers();
  require('./aiProxy.js').registerAiProxy();

  const mock = await startMockServer();
  const imageServer = http.createServer((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'image/png' });
    response.end(Buffer.from(URL_PNG_BASE64, 'base64'));
  });
  await new Promise<void>((resolve) => imageServer.listen(0, '127.0.0.1', () => resolve()));
  const imageUrl = `http://127.0.0.1:${(imageServer.address() as any).port}/image.png`;
  IMAGE_URL = imageUrl;
  seededAiSettings.settings = {
    provider: 'custom',
    baseUrl: mock.baseUrl,
    model: 'test-model',
    temperature: 0.2,
    maxTokens: 4096,
    systemPrompt: 'Ты тестовый ассистент.',
    activePresetId: 'general',
  };
  seededAiSettings.keys = { custom: 'test-key' };

  const win = new BrowserWindow({
    width: 1500,
    height: 940,
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
    },
  });

  win.webContents.on('console-message', (_event, _level, message) => {
    if (/error|Error|ошибк/i.test(message)) out.errors.push(message);
  });

  const js = (expression) => win.webContents.executeJavaScript(`(async () => { ${expression} })()`);

  const waitIdle = async (timeout = 20000) => {
    const deadline = Date.now() + timeout;
    while (Date.now() < deadline) {
      const idle = await js(`
        const stop = [...document.querySelectorAll('button')]
          .some((button) => (button.textContent || '').trim() === 'Стоп');
        const status = /Готовлю запрос|Читаю документ|Правлю документ|Переписываю документ|Ищу и заменяю|Меняю оформление|Настраиваю страницу|Повторяю правку|Удаляю лишнее|Работаю через текстовые действия/.test(
          document.body.innerText
        );
        return !stop && !status;
      `);
      if (idle) return true;
      await sleep(250);
    }
    return false;
  };

  const send = async (text) => {
    const sent = await js(`
      const ta = [...document.querySelectorAll('textarea')].find((item) =>
        (item.placeholder || '').includes('Спросите')
      );
      if (!ta) return false;
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      setter.call(ta, ${JSON.stringify(text)});
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise((resolve) => setTimeout(resolve, 200));
      const button = [...document.querySelectorAll('button')].find((item) =>
        /Отправить/.test(item.textContent || '')
      );
      if (button) button.click();
      return Boolean(button);
    `);
    await sleep(500);
    const idle = await waitIdle();
    await sleep(300);
    return { sent: Boolean(sent), idle };
  };

  const documentState = () =>
    js(`
      const canvas = document.querySelector('[contenteditable=true]');
      return canvas
        ? { html: canvas.innerHTML, text: canvas.innerText }
        : { html: '', text: '' };
    `);

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'));
    await sleep(2500);

    // Права записи AI — persisted-настройка: после прогона «без прав» она
    // осталась бы false и следующий запуск не смог править документ.
    await js(`
      localStorage.setItem('ws_ai_doc_write', 'true');
      return true;
    `);
    await win.webContents.reload();
    await sleep(2200);

    // Открываем резюме из дерева файлов.
    let opened = false;
    for (let attempt = 0; attempt < 20 && !opened; attempt++) {
      opened = await js(`
        const node = [...document.querySelectorAll('div')].find((element) => {
          const title = (element.getAttribute('title') || '').split(' — ')[0];
          return title === ${JSON.stringify(DOC_NAME)} && element.className.includes('cursor-pointer');
        });
        if (node) { node.click(); return true; }
        return false;
      `);
      if (!opened) await sleep(300);
    }
    await sleep(2500);

    step('documentOpened', await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      return {
        opened: Boolean(canvas),
        hasResume: Boolean(canvas) && canvas.innerText.includes('Иван Иванов'),
        hasJunk: Boolean(canvas) && canvas.innerText.includes('не могу редактировать'),
        rightsShown: document.body.innerText.includes('Права записи выданы'),
        panelPresent: [...document.querySelectorAll('textarea')].some((item) =>
          (item.placeholder || '').includes('Спросите')
        ),
      };
    `));

    /* ---------- 1. Удаление мусора ---------- */
    const cleanupSend = await send('Удали из документа всё, что не относится к резюме');
    const cleanupState = await documentState();
    step('cleanup', {
      sent: cleanupSend.sent,
      idle: cleanupSend.idle,
      junkGone:
        !cleanupState.text.includes('не могу редактировать') &&
        !cleanupState.text.includes('Хотите, я создам'),
      resumeKept: cleanupState.text.includes('Иван Иванов') && cleanupState.text.includes('Опыт работы'),
      reportShown: (await js('return document.body.innerText')).includes('Убрал из документа'),
    });

    /* ---------- 2. ATS-friendly ---------- */
    const atsSend = await send('Сделай резюме ATS-friendly');
    const atsState = await documentState();
    step('ats', {
      sent: atsSend.sent,
      idle: atsSend.idle,
      structureChanged: atsState.html.includes('<h3') && atsState.html.includes('<ul'),
      resumeKept: atsState.text.includes('Иван Иванов') && atsState.text.includes('Образование'),
      reportShown: (await js('return document.body.innerText')).includes('под ATS'),
    });

    /* ---------- 3. Оформление: H2 + контакты по центру ---------- */
    const formatSend = await send('Сделай заголовки разделов H2 и выровняй контакты по центру');
    const formatState = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      const headings = canvas ? [...canvas.querySelectorAll('h2')].map((item) => item.textContent) : [];
      const contact = canvas
        ? [...canvas.querySelectorAll('p')].find((item) => item.textContent.includes('@'))
        : null;
      return {
        h2: headings,
        contactAlign: contact ? contact.style.textAlign : '',
        html: canvas ? canvas.innerHTML : '',
      };
    `);
    step('format', {
      sent: formatSend.sent,
      idle: formatSend.idle,
      sectionsAreH2:
        formatState.h2.some((text) => text.includes('Опыт')) &&
        formatState.h2.some((text) => text.includes('Навыки')),
      contactsCentered: formatState.contactAlign === 'center',
      reportShown: (await js('return document.body.innerText')).includes('выровнял контакты'),
    });

    /* ---------- 3b. Откат серии правок ---------- */
    const undoFormat = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        /Вернуть прежний текст/.test(item.textContent || '')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(900);
    step('undoSeries', {
      clicked: Boolean(undoFormat),
      h3Restored: await js(`
        const canvas = document.querySelector('[contenteditable=true]');
        return Boolean(canvas) && canvas.querySelectorAll('h3').length >= 2;
      `),
      centerCleared: await js(`
        const canvas = document.querySelector('[contenteditable=true]');
        const contact = canvas
          ? [...canvas.querySelectorAll('p')].find((item) => item.textContent.includes('@'))
          : null;
        return !contact || contact.style.textAlign !== 'center';
      `),
    });

    /* ---------- 3c. Панель истории AI-транзакций ---------- */
    const historyOpened = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        (item.getAttribute('title') || '').includes('История правок ассистента')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(400);
    const historyPanelText = await js('return document.body.innerText');
    step('historyPanel', {
      opened: Boolean(historyOpened),
      hasRecords: /История ассистента · [1-9]/.test(historyPanelText),
      hasDiffCount: historyPanelText.includes('блоков:'),
      hasUndo: historyPanelText.includes('Отменить'),
      hasRedo: historyPanelText.includes('Повторить'),
    });
    await js(`
      const close = [...document.querySelectorAll('button')].find(
        (item) => (item.getAttribute('title') || '') === 'Закрыть историю'
      );
      if (close) close.click();
    `);

    // Возвращаем оформление повторным запросом (проверяем и повторный агентный цикл).
    const formatAgain = await send('Сделай заголовки разделов H2 и выровняй контакты по центру');
    const formatAgainState = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      const contact = canvas
        ? [...canvas.querySelectorAll('p')].find((item) => item.textContent.includes('@'))
        : null;
      return {
        h2: canvas ? canvas.querySelectorAll('h2').length : 0,
        contactAlign: contact ? contact.style.textAlign : '',
      };
    `);
    step('formatAgain', {
      sent: formatAgain.sent,
      idle: formatAgain.idle,
      sectionsAreH2: formatAgainState.h2 >= 2,
      contactsCentered: formatAgainState.contactAlign === 'center',
    });

    /* ---------- 3d. Панель структуры документа ---------- */
    const outlineState = await js(`
      const panel = document.querySelector('[data-outline-panel]');
      return {
        visible: !!panel,
        items: panel
          ? [...panel.querySelectorAll('[data-outline-item]')].map((item) => (item.textContent || '').trim())
          : [],
      };
    `);
    step('outlinePanel', {
      visible: outlineState.visible,
      hasItems: outlineState.items.length >= 2,
      hasExperience: outlineState.items.some((text) => text.includes('Опыт')),
      hasSkills: outlineState.items.some((text) => text.includes('Навыки') || text.includes('Ключевые')),
    });

    /* ---------- 4. Точечная замена ---------- */
    const findSend = await send('Замени "Июль 2024" на "Июнь 2024"');
    const findState = await documentState();
    step('findReplace', {
      sent: findSend.sent,
      idle: findSend.idle,
      replaced: findState.text.includes('Июнь 2024') && !findState.text.includes('Июль 2024'),
      reportShown: (await js('return document.body.innerText')).includes('Июнь 2024'),
    });

    /* ---------- 5. Откат одного ответа ---------- */
    const undoFind = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        /Вернуть прежний текст/.test(item.textContent || '')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(900);
    const undoState = await documentState();
    step('undoSingle', {
      clicked: Boolean(undoFind),
      julyRestored: undoState.text.includes('Июль 2024') && !undoState.text.includes('Июнь 2024'),
    });

    /* ---------- 6. Сохранение оформления ---------- */
    const saved = await js(`
      const button = [...document.querySelectorAll('button')].find(
        (item) => (item.textContent || '').trim() === 'Сохранить' && !item.disabled
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(2500);
    const savedDoc = await readDocxXml(path.join(SANDBOX, DOC_NAME));
    step('save', {
      clicked: Boolean(saved),
      fileExists: fs.existsSync(path.join(SANDBOX, DOC_NAME)),
      headingsPreserved: savedDoc.xml.includes('w:pStyle w:val="Heading2"'),
      centerPreserved: savedDoc.xml.includes('<w:jc w:val="center"/>'),
      listPreserved: /<w:numId w:val="1"\/>/.test(savedDoc.xml),
      dateRestored: savedDoc.text.includes('Июль 2024'),
      cleanText:
        !savedDoc.text.includes('не могу редактировать') &&
        !savedDoc.text.includes('Хотите, я создам'),
    });

    /* ---------- 6b. Быстрая локальная правка заголовка без запроса к модели ---------- */
    const requestsBeforeTitle = out.requests.length;
    const quickTitleSend = await send('Измени заголовок на «Иван Петров»');
    await sleep(1500);
    const quickTitleState = await documentState();
    step('quickTitleEdit', {
      sent: quickTitleSend.sent,
      idle: quickTitleSend.idle,
      titleChanged: quickTitleState.text.includes('Иван Петров'),
      oldTitleGone: !quickTitleState.text.includes('Иван Иванов'),
      bodyKept:
        quickTitleState.text.includes('Опыт работы') && quickTitleState.text.includes('Москва'),
      noModelCall: out.requests.length === requestsBeforeTitle,
    });

    const titleTrace = await js(`
      const chips = [...document.querySelectorAll('[data-tool-trace]')];
      return { names: chips.map((chip) => chip.getAttribute('data-tool-trace')) };
    `);
    step('toolTracePersisted', {
      shown: titleTrace.names.length > 0,
      findReplace: titleTrace.names.includes('find_replace'),
    });

    /* ---------- 6b. Выделенный фрагмент: get_selection → replace_selection ---------- */
    const selected = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      if (!canvas) return null;
      const walker = document.createTreeWalker(canvas, NodeFilter.SHOW_TEXT);
      let node = null;
      while (walker.nextNode()) {
        const current = walker.currentNode;
        if ((current.textContent || '').includes('Москва')) { node = current; break; }
      }
      if (!node) return null;
      const index = (node.textContent || '').indexOf('Москва');
      const range = document.createRange();
      range.setStart(node, index);
      range.setEnd(node, index + 6);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.dispatchEvent(new Event('selectionchange'));
      return { text: selection.toString(), index };
    `);
    await sleep(300);
    const selectionSend = await send('Замени выделенное на «Санкт-Петербург»');
    await sleep(2000);
    const selectionState = await documentState();
    const selectionRequests = out.requests.filter((request) => request.scenario === 'selection');
    step('selectionEdit', {
      selected: Boolean(selected) && selected.text === 'Москва',
      sent: selectionSend.sent,
      idle: selectionSend.idle,
      replaced: selectionState.text.includes('Санкт-Петербург'),
      oldGone: !selectionState.text.includes('Москва'),
      bodyKept:
        selectionState.text.includes('Опыт работы') &&
        selectionState.text.includes('Иван Петров'),
      toolsAdvertised: selectionRequests.some(
        (request) =>
          request.toolNames.includes('get_selection') &&
          request.toolNames.includes('replace_selection')
      ),
    });

    /* ---------- 6b2. Другой файл workspace: чтение без права записи ---------- */
    const secondBefore = fs.readFileSync(path.join(SANDBOX, 'Второй.md'), 'utf-8');
    const multiSend = await send('Прочитай второй файл и скажи, что в нём');
    await sleep(1500);
    const multiPageText = await js('return document.body.innerText');
    const multiRequests = out.requests.filter((request) => request.scenario === 'multidoc');
    const multiState = await documentState();
    const multiChip = await js(`
      return Boolean(document.querySelector('[data-context-file="Второй.md"]'));
    `);
    step('multiDocRead', {
      sent: multiSend.sent,
      idle: multiSend.idle,
      readSucceeded: multiPageText.includes('секретный текст найден'),
      toolsAdvertised: multiRequests.some((request) =>
        request.toolNames.includes('read_document')
      ),
      secondFileUnchanged: fs.readFileSync(path.join(SANDBOX, 'Второй.md'), 'utf-8') === secondBefore,
      activeDocumentKept: multiState.text.includes('Санкт-Петербург'),
      contextChip: Boolean(multiChip),
    });

    /* ---------- 6c. Точечная правка строки: replace_document запрещён ---------- */
    const blockedSend = await send('Замени строку с должностью');
    await sleep(1500);
    const blockedState = await documentState();
    step('preciseLineNoFullReplace', {
      sent: blockedSend.sent,
      idle: blockedSend.idle,
      lineChanged: blockedState.text.includes('ведущий инженер'),
      bodyKept: blockedState.text.includes('Опыт работы') && blockedState.text.includes('Санкт-Петербург'),
      noFullReplace: !blockedState.text.includes('Полная замена документа'),
    });

    /* ---------- 6d. Повтор после текстового ответа: правка всё равно применяется ---------- */
    const retrySend = await send('Поменяй строку с опытом');
    await sleep(1500);
    const retryState = await documentState();
    step('preciseLineRetry', {
      sent: retrySend.sent,
      idle: retrySend.idle,
      headingChanged: retryState.text.includes('Ключевые навыки'),
      bodyKept: retryState.text.includes('Опыт работы') && retryState.text.includes('Санкт-Петербург'),
    });

    /* ---------- 6e. ATS/ATC: копия не принимается, резюме переписывается ---------- */
    const atsRewriteSend = await send('Резюме должно проходить ATC-фильтры');
    await sleep(2000);
    const atsRewriteState = await documentState();
    step('atsRewrite', {
      sent: atsRewriteSend.sent,
      idle: atsRewriteSend.idle,
      titleChanged: atsRewriteState.text.includes('Иван Сидоров'),
      atsBody: atsRewriteState.text.includes('Диагностика оборудования'),
      educationAdded: atsRewriteState.text.includes('МГУ, 2018'),
      oldHeadingGone: !atsRewriteState.text.includes('Ключевые навыки'),
      reportGone: !atsRewriteState.text.includes('Оформил документ'),
      saveOfferGone: !atsRewriteState.text.includes('сохранить сейчас'),
      noTables: !atsRewriteState.html.includes('<table'),
    });

    /* ---------- 6f. No-op замена отклоняется ---------- */
    const beforeNoop = await documentState();
    const noopSend = await send('Перепиши документ полностью');
    await sleep(2000);
    const afterNoop = await documentState();
    step('noopReplaceRejected', {
      sent: noopSend.sent,
      idle: noopSend.idle,
      htmlUnchanged: afterNoop.html === beforeNoop.html,
      textUnchanged: afterNoop.text === beforeNoop.text,
    });

    /* ---------- 6g. Word + Markdown-таблица: не предлагать Excel ---------- */
    const tableSend = await send('Покажи таблицу навыков');
    await sleep(1200);
    const pageTextAfterTable = await js('return document.body.innerText');
    step('noExcelForWord', {
      sent: tableSend.sent,
      idle: tableSend.idle,
      noExcelButton: !pageTextAfterTable.includes('Создать Excel'),
      noTableBanner: !pageTextAfterTable.includes('Обнаружена таблица данных'),
    });

    /* ---------- 6g. «Напиши сказку» без инструментов: текст попадает в документ ---------- */
    const storySend = await send('Напиши в этом документе мини-сказку');
    await sleep(2000);
    const storyState = await documentState();
    step('storyAutoApply', {
      sent: storySend.sent,
      idle: storySend.idle,
      inserted: storyState.text.includes('Сказка о забытом фонаре'),
      storyBody: storyState.text.includes('Айя'),
      reportGone: !storyState.text.includes('сказка записана'),
      offerGone:
        !storyState.text.includes('Если хочешь') &&
        !storyState.text.includes('Что скажешь'),
      previousTextKept: storyState.text.includes('Опыт работы'),
    });

    /* ---------- 7. Fallback без function calling ---------- */
    const fallbackSend = await send('Оформи документ через запасной протокол');
    const fallbackState = await documentState();
    const fallbackRequests = out.requests.filter((request) => request.scenario === 'fallback');
    step('fallbackProtocol', {
      sent: fallbackSend.sent,
      idle: fallbackSend.idle,
      toolsRejected: fallbackRequests.some((request) => request.toolNames.length > 0),
      retriedWithoutTools: fallbackRequests.some((request) => request.toolNames.length === 0),
      promptFallback: fallbackRequests.some((request) => request.promptFallback),
      documentUpdated: fallbackState.text.includes('Оформлено через document_apply'),
      contactsCentered: fallbackState.html.includes('text-align:center'),
      reportShown: (await js('return document.body.innerText')).includes('Оформляю документ'),
    });

    /* ---------- 8. Оформление страницы инструментом page_setup ---------- */
    const pageSend = await send('Оформи как реальный документ для деловой переписки');
    await sleep(1500);
    const pageState = await documentState();
    const pageSave = await js(`
      const button = [...document.querySelectorAll('button')].find(
        (item) => (item.textContent || '').trim() === 'Сохранить' && !item.disabled
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(2500);
    const pageDoc = await readDocxXml(path.join(SANDBOX, DOC_NAME));
    step('pageSetup', {
      sent: pageSend.sent,
      idle: pageSend.idle,
      saved: Boolean(pageSave),
      margins: /<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"\/>/.test(
        pageDoc.xml
      ),
      serifFont: pageDoc.styles.includes('w:ascii="Times New Roman"'),
      font14: /<w:sz w:val="28"\/>/.test(pageDoc.styles),
      singleSpacing: /w:line="240"/.test(pageDoc.styles),
      justified: pageState.html.includes('text-align:justify'),
      authoredCenterKept: pageState.html.includes('text-align:center'),
      headingScaleKept: (pageState.html.match(/<h[1-6][^>]*>/g) || []).every(
        (tag) => !tag.includes('font-size')
      ),
    });

    /* ---------- 8b. Колонтитул через get_document_info → set_header_footer ---------- */
    const bodyBeforeHeader = await documentState();
    const headerSend = await send('Измени верхний колонтитул на «Шапка отдела»');
    await sleep(2000);
    const headerSave = await js(`
      const button = [...document.querySelectorAll('button')].find(
        (item) => (item.textContent || '').trim() === 'Сохранить' && !item.disabled
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(2500);
    const headerParts = await readDocxParts(path.join(SANDBOX, DOC_NAME));
    const headerEntry = Object.entries(headerParts).find(([name]) =>
      /^word\/header\d+\.xml$/.test(name)
    );
    const bodyAfterHeader = await documentState();
    step('headerFooter', {
      sent: headerSend.sent,
      idle: headerSend.idle,
      saved: Boolean(headerSave),
      headerPartCreated: Boolean(headerEntry),
      headerText: Boolean(headerEntry && headerEntry[1].includes('Шапка отдела')),
      referenceAdded: Boolean(
        headerParts['word/document.xml']?.includes('w:headerReference')
      ),
      bodyUnchanged: bodyBeforeHeader.text === bodyAfterHeader.text,
    });

    /* ---------- 8c. Excel: чтение и запись ячеек через AI ---------- */
    const excelOpened = await js(`
      const nodes = [...document.querySelectorAll('div')]
        .filter((d) => d.textContent?.trim() === 'Книга.xlsx');
      if (!nodes.length) return false;
      const clickable = nodes.find((d) => d.className.includes('cursor-pointer')) || nodes[0];
      clickable.click();
      return true;
    `);
    await sleep(3000);
    const excelWriteSend = await send('Заполни A4:B4 значениями Апельсин и сумму B2+B3');
    await sleep(2500);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);

    let excelWriteState: any = null;
    try {
      const savedWorkbook = XLSX.read(
        fs.readFileSync(path.join(SANDBOX, 'Книга.xlsx')),
        { type: 'buffer' }
      );
      const dataSheet = savedWorkbook.Sheets['Данные'];
      excelWriteState = {
        a4: dataSheet?.A4?.v === 'Апельсин',
        b4Formula: dataSheet?.B4?.f === 'B2+B3',
        otherKept: dataSheet?.A2?.v === 'Яблоко',
      };
    } catch {
      excelWriteState = { a4: false, b4Formula: false, otherKept: false };
    }
    const excelDiffText = await js(`
      const node = document.querySelector('[data-ai-diff]');
      return node ? node.innerText : '';
    `);
    step('excelWriteCells', {
      opened: Boolean(excelOpened),
      sent: excelWriteSend.sent,
      idle: excelWriteSend.idle,
      aiDiff: /A4/.test(excelDiffText) && /B4/.test(excelDiffText),
      ...excelWriteState,
    });

    const excelZoomState = await js(`
      const table = document.querySelector('table[style*="table-layout"]');
      const target = table?.parentElement;
      if (!table || !target) return null;
      target.dispatchEvent(new WheelEvent('wheel', { deltaY: -200, ctrlKey: true, bubbles: true, cancelable: true }));
      await new Promise((resolve) => setTimeout(resolve, 120));
      const status = [...document.querySelectorAll('span')]
        .map((node) => node.textContent || '')
        .find((text) => text.includes('Масштаб:'));
      const zoom = table.style.zoom;
      const styleAttr = table.getAttribute('style') || '';
      const computed = getComputedStyle(table).zoom || '';
      target.dispatchEvent(new WheelEvent('wheel', { deltaY: 200, ctrlKey: true, bubbles: true, cancelable: true }));
      return { zoom, styleAttr, computed, status };
    `);
    step('excelZoom', {
      zoomed: /(?:^|;)\s*zoom:\s*(?:1\.[1-9]|[2-9])/.test(excelZoomState?.styleAttr || '') ||
        Number.parseFloat(excelZoomState?.computed || '1') > 1,
      status: /Масштаб: \d+%/.test(excelZoomState?.status || ''),
    });

    const excelSheetSend = await send('Добавь лист Итоги');
    await sleep(2000);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let excelSheetState: any = null;
    try {
      const savedWorkbook = XLSX.read(
        fs.readFileSync(path.join(SANDBOX, 'Книга.xlsx')),
        { type: 'buffer' }
      );
      excelSheetState = {
        sheetAdded: savedWorkbook.SheetNames.includes('Итоги'),
        originalKept: savedWorkbook.SheetNames.includes('Данные'),
      };
    } catch {
      excelSheetState = { sheetAdded: false, originalKept: false };
    }
    step('excelAddSheet', {
      sent: excelSheetSend.sent,
      idle: excelSheetSend.idle,
      ...excelSheetState,
    });

    /* ---------- 8c2. Excel: формат и объединение ---------- */
    const excelOpsSend = await send('Объедини A1:B1 и выдели жирным');
    await sleep(2200);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let excelOpsState: any = null;
    try {
      const savedWorkbook = XLSX.read(
        fs.readFileSync(path.join(SANDBOX, 'Книга.xlsx')),
        { type: 'buffer', cellStyles: true }
      );
      const dataSheet = savedWorkbook.Sheets['Данные'];
      const merges = dataSheet?.['!merges'] ?? [];
      const encodedStyles = String(savedWorkbook.Custprops?.WorkSpaceStyles ?? '');
      const styles = encodedStyles.startsWith('b64:')
        ? Buffer.from(encodedStyles.slice(4), 'base64').toString('utf-8')
        : encodedStyles;
      excelOpsState = {
        merged: merges.some(
          (merge: any) =>
            merge.s.r === 0 &&
            merge.s.c === 0 &&
            merge.e.r === 0 &&
            merge.e.c === 1
        ),
        boldSaved: styles.includes('"bold":true'),
        fillSaved: styles.includes('#FFFF00'),
      };
    } catch {
      excelOpsState = { merged: false, boldSaved: false, fillSaved: false };
    }
    step('excelFormatMerge', {
      sent: excelOpsSend.sent,
      idle: excelOpsSend.idle,
      ...excelOpsState,
    });

    const excelSelectionSend = await send('Что сейчас выделено в таблице');
    await sleep(1500);
    const excelSelectionText = await js('return document.body.innerText');
    step('excelSelection', {
      sent: excelSelectionSend.sent,
      idle: excelSelectionSend.idle,
      read: excelSelectionText.includes('Выделение получено'),
    });

    /* ---------- 8d. PowerPoint: текст слайда и новый слайд ---------- */
    const pptOpened = await js(`
      const nodes = [...document.querySelectorAll('div')]
        .filter((d) => d.textContent?.trim() === 'Презентация.pptx');
      if (!nodes.length) return false;
      const clickable = nodes.find((d) => d.className.includes('cursor-pointer')) || nodes[0];
      clickable.click();
      return true;
    `);
    await sleep(3000);
    const pptTextSend = await send('Замени текст на первом слайде на «Новый заголовок»');
    await sleep(2500);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);

    let pptTextState: any = null;
    try {
      const savedZip = await JSZip.loadAsync(
        fs.readFileSync(path.join(SANDBOX, 'Презентация.pptx'))
      );
      const slideXml = (await savedZip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
      pptTextState = {
        replaced: slideXml.includes('Новый заголовок'),
        oldGone: !slideXml.includes('Старый заголовок'),
      };
    } catch {
      pptTextState = { replaced: false, oldGone: false };
    }
    step('pptxSetText', {
      opened: Boolean(pptOpened),
      sent: pptTextSend.sent,
      idle: pptTextSend.idle,
      ...pptTextState,
    });

    const pptAddSend = await send('Добавь слайд «Итоги»');
    await sleep(2000);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let pptAddState: any = null;
    try {
      const savedZip = await JSZip.loadAsync(
        fs.readFileSync(path.join(SANDBOX, 'Презентация.pptx'))
      );
      const presentationXml =
        (await savedZip.file('ppt/presentation.xml')?.async('string')) ?? '';
      const slide2Xml = (await savedZip.file('ppt/slides/slide2.xml')?.async('string')) ?? '';
      pptAddState = {
        slideAdded: Boolean(savedZip.file('ppt/slides/slide2.xml')),
        inPresentation: (presentationXml.match(/<p:sldId\b/g) || []).length === 2,
        titleSaved: slide2Xml.includes('Итоги'),
      };
    } catch {
      pptAddState = { slideAdded: false, inPresentation: false, titleSaved: false };
    }
    step('pptxAddSlide', {
      sent: pptAddSend.sent,
      idle: pptAddSend.idle,
      ...pptAddState,
    });

    const pptTextboxSend = await send('Добавь текст на первый слайд');
    await sleep(2000);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2200);
    let pptTextboxState: any = null;
    try {
      const savedZip = await JSZip.loadAsync(
        fs.readFileSync(path.join(SANDBOX, 'Презентация.pptx'))
      );
      const slideXml = (await savedZip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
      pptTextboxState = {
        textAdded: slideXml.includes('Добавленный текст'),
        oldTitleKept: slideXml.includes('Новый заголовок'),
      };
    } catch {
      pptTextboxState = { textAdded: false, oldTitleKept: false };
    }
    step('pptxAddTextBox', {
      sent: pptTextboxSend.sent,
      idle: pptTextboxSend.idle,
      ...pptTextboxState,
    });

    const pptNotesSend = await send('Измени заметки первого слайда');
    await sleep(2000);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let pptNotesState: any = null;
    try {
      const savedZip = await JSZip.loadAsync(
        fs.readFileSync(path.join(SANDBOX, 'Презентация.pptx'))
      );
      const notesXml =
        (await savedZip.file('ppt/notesSlides/notesSlide1.xml')?.async('string')) ?? '';
      pptNotesState = {
        updated: notesXml.includes('Тезисы для доклада'),
        oldGone: !notesXml.includes('Заметка докладчика'),
        slideKept: Boolean(savedZip.file('ppt/slides/slide1.xml')),
      };
    } catch {
      pptNotesState = { updated: false, oldGone: false, slideKept: false };
    }
    step('pptxSetNotes', {
      sent: pptNotesSend.sent,
      idle: pptNotesSend.idle,
      ...pptNotesState,
    });

    const pptImageSend = await send('Вставь картинку на слайд');
    await sleep(2200);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let pptImageState: any = null;
    try {
      const savedZip = await JSZip.loadAsync(
        fs.readFileSync(path.join(SANDBOX, 'Презентация.pptx'))
      );
      const slideXml = (await savedZip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
      const media = Object.keys(savedZip.files).filter((name) =>
        /^ppt\/media\/image\d+\./i.test(name)
      );
      pptImageState = {
        picture: slideXml.includes('<p:pic>') && /r:embed="rId\d+"/.test(slideXml),
        mediaSaved: media.length > 0,
        textKept: slideXml.includes('Новый заголовок'),
      };
    } catch {
      pptImageState = { picture: false, mediaSaved: false, textKept: false };
    }
    step('pptxInsertImage', {
      sent: pptImageSend.sent,
      idle: pptImageSend.idle,
      ...pptImageState,
    });

    const pptLayoutSend = await send('Смени макет первого слайда на «Раздел»');
    await sleep(2200);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let pptLayoutState: any = null;
    try {
      const savedZip = await JSZip.loadAsync(
        fs.readFileSync(path.join(SANDBOX, 'Презентация.pptx'))
      );
      const relsXml =
        (await savedZip.file('ppt/slides/_rels/slide1.xml.rels')?.async('string')) ?? '';
      const slideXml = (await savedZip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
      pptLayoutState = {
        layoutApplied: relsXml.includes('slideLayout2.xml'),
        mediaKept: Object.keys(savedZip.files).some((name) =>
          /^ppt\/media\/image\d+\./i.test(name)
        ),
        textKept: slideXml.includes('Новый заголовок'),
      };
    } catch {
      pptLayoutState = { layoutApplied: false, mediaKept: false, textKept: false };
    }
    step('pptxSetLayout', {
      sent: pptLayoutSend.sent,
      idle: pptLayoutSend.idle,
      ...pptLayoutState,
    });

    const pptSelectFocused = await js(`
      const firstSlide = [...document.querySelectorAll('aside button')].find((item) =>
        /^[1-9]/.test((item.textContent || '').trim())
      );
      if (firstSlide) firstSlide.click();
      return Boolean(firstSlide);
    `);
    await sleep(400);
    const pptTextareaFocused = await js(`
      const canvas = [...document.querySelectorAll('textarea')].find((item) =>
        item.closest('[style*="aspect-ratio"]')
      );
      if (!canvas) return false;
      canvas.focus();
      return true;
    `);
    await sleep(400);
    const pptSelectionSend = await send('Что выделено на слайде?');
    await sleep(1600);
    const pptPageText = await js('return document.body.innerText');
    step('pptxSelection', {
      focused: Boolean(pptSelectFocused && pptTextareaFocused),
      sent: pptSelectionSend.sent,
      idle: pptSelectionSend.idle,
      read: pptPageText.includes('Выделена фигура заголовка'),
    });

    /* ---------- 8e. Markdown: точечная замена через md_replace ---------- */
    const mdOpened = await js(`
      const nodes = [...document.querySelectorAll('div')]
        .filter((d) => d.textContent?.trim() === 'Второй.md');
      if (!nodes.length) return false;
      const clickable = nodes.find((d) => d.className.includes('cursor-pointer')) || nodes[0];
      clickable.click();
      return true;
    `);
    await sleep(2500);
    const mdSend = await send('Замени в Markdown «Секретный» на «Открытый»');
    await sleep(2000);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2200);
    let mdState: any = null;
    try {
      const content = fs.readFileSync(path.join(SANDBOX, 'Второй.md'), 'utf-8');
      mdState = {
        replaced: content.includes('Открытый текст второго файла'),
        oldGone: !content.includes('Секретный'),
        headingKept: content.includes('# Второй документ'),
      };
    } catch {
      mdState = { replaced: false, oldGone: false, headingKept: false };
    }
    step('markdownReplace', {
      opened: Boolean(mdOpened),
      sent: mdSend.sent,
      idle: mdSend.idle,
      ...mdState,
    });

    const mdHistoryOpened = await js(`
      const button = document.querySelector('[data-md-history-toggle]');
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(400);
    const mdHistoryText = await js(`
      const panel = document.querySelector('[data-md-history-panel]');
      return panel ? panel.innerText : '';
    `);
    step('markdownHistory', {
      opened: Boolean(mdHistoryOpened),
      hasEntries: /История Markdown · [1-9]/.test(mdHistoryText),
      hasDiff: /\+\d+ −\d+/.test(mdHistoryText),
    });
    await js(`
      const panel = document.querySelector('[data-md-history-panel]');
      const close = panel ? [...panel.querySelectorAll('button')].find((button) => (button.textContent || '').trim() === '×') : null;
      if (close) close.click();
    `);

    const mdStructureSend = await send('Оформи Markdown структурно');
    await sleep(2500);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let mdStructureState: any = null;
    try {
      const content = fs.readFileSync(path.join(SANDBOX, 'Второй.md'), 'utf-8');
      mdStructureState = {
        heading: /^## Открытый текст/m.test(content),
        link: content.includes('[второго файла](https://example.com)'),
        table: content.includes('| Поле | Значение |') && content.includes('| A | 1 |'),
      };
    } catch {
      mdStructureState = { heading: false, link: false, table: false };
    }
    step('markdownStructure', {
      sent: mdStructureSend.sent,
      idle: mdStructureSend.idle,
      ...mdStructureState,
    });

    /* ---------- 9. Возврат к Word и проверка прав ---------- */
    await js(`
      const nodes = [...document.querySelectorAll('div')]
        .filter((d) => d.textContent?.trim() === ${JSON.stringify(DOC_NAME)});
      const clickable = nodes.find((d) => d.className.includes('cursor-pointer')) || nodes[0];
      if (clickable) { clickable.click(); return true; }
      return false;
    `);
    await sleep(2500);

    /* ---------- 8e. Поиск и замена в панели Word ---------- */
    const findOpened = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        (item.getAttribute('title') || '').includes('Найти и заменить')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(400);
    const findSetup = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      const word = canvas && canvas.innerText.includes('Москва') ? 'Москва' : 'Опыт';
      const setValue = (element, value) => {
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        setter.call(element, value);
        element.dispatchEvent(new Event('input', { bubbles: true }));
      };
      const findInput = document.querySelector('[data-find-input]');
      const replaceInput = document.querySelector('[data-replace-input]');
      if (findInput) setValue(findInput, word);
      if (replaceInput) setValue(replaceInput, 'ТестПоиска');
      return { word, hasFind: !!findInput, hasReplace: !!replaceInput };
    `);
    await sleep(300);
    const findReplaceClicked = await js(`
      const button = document.querySelector('[data-find-replace-all]');
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(600);
    const findUiState = await documentState();
    step('findReplaceUi', {
      wordFound: Boolean(findSetup.word),
      inputsShown: Boolean(findSetup.hasFind && findSetup.hasReplace),
      replaced: Boolean(findReplaceClicked) && findUiState.text.includes('ТестПоиска'),
      oldGone: Boolean(findSetup.word) && !findUiState.text.includes(findSetup.word),
    });

    const crumbSelected = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      if (!canvas) return false;
      const heading = [...canvas.querySelectorAll('h1, h2, h3, h4, h5, h6')].find((item) =>
        /Опыт|Навыки|ТестПоиска/.test(item.textContent || '')
      );
      const next = heading ? heading.nextElementSibling : null;
      if (!next) return false;
      const walker = document.createTreeWalker(next, NodeFilter.SHOW_TEXT);
      const node = walker.nextNode();
      if (!node) return false;
      const range = document.createRange();
      range.setStart(node, 0);
      range.setEnd(node, Math.min((node.textContent || '').length, 4));
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.dispatchEvent(new Event('selectionchange'));
      return true;
    `);
    await sleep(400);
    const crumbText = await js(`
      const node = document.querySelector('[data-breadcrumb]');
      return node ? node.innerText : '';
    `);
    step('breadcrumb', {
      selected: Boolean(crumbSelected),
      shown: Boolean(crumbText && crumbText.length > 0),
      hasHeading: /ТестПоиска|Опыт|Навыки/.test(crumbText || ''),
    });

    const atomicSend = await send('Проверь атомарность: измени слово и удали несуществующий блок');
    await sleep(2000);
    const atomicState = await documentState();
    step('atomicTransaction', {
      sent: atomicSend.sent,
      idle: atomicSend.idle,
      rolledBack:
        atomicState.text.includes('ТестПоиска') && !atomicState.text.includes('АТОМНАЯ_ПРАВКА'),
    });

    const selectionRestoreSetup = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      if (!canvas) return false;
      const walker = document.createTreeWalker(canvas, NodeFilter.SHOW_TEXT);
      let node = null;
      while (walker.nextNode()) {
        if ((walker.currentNode.textContent || '').includes('ТестПоиска')) {
          node = walker.currentNode;
          break;
        }
      }
      if (!node) return false;
      const index = node.textContent.indexOf('ТестПоиска');
      const range = document.createRange();
      range.setStart(node, index);
      range.setEnd(node, index + 'ТестПоиска'.length);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.dispatchEvent(new Event('selectionchange'));
      return true;
    `);
    await sleep(300);
    const selRestoreSend = await send('Сделай выделенное жирным');
    await sleep(2200);
    const selectionAfter = await js('return window.getSelection().toString()');
    step('selectionRestore', {
      setup: Boolean(selectionRestoreSetup),
      sent: selRestoreSend.sent,
      idle: selRestoreSend.idle,
      selectionKept: String(selectionAfter || '').includes('ТестПоиска'),
    });

    const versionsOpened = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        (item.getAttribute('title') || '').includes('Версии документа')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(400);
    const versionsText = await js('return document.body.innerText');
    step('versionsPanel', {
      opened: Boolean(versionsOpened),
      hasEntries: /Версии документа · [1-9]/.test(versionsText),
    });
    await js(`
      const button = [...document.querySelectorAll('button')].find(
        (item) => (item.getAttribute('title') || '') === 'Закрыть версии'
      );
      if (button) button.click();
    `);

    const pdfPath = path.join(SANDBOX, 'export-check.pdf');
    fs.rmSync(pdfPath, { force: true });
    const pdfResult = await js(`
      return await window.electronAPI.exportPdf(
        '<html><body><h1>PDF TEST</h1><p>Проверка экспорта</p></body></html>',
        'export-check.pdf',
        ${JSON.stringify(pdfPath)}
      );
    `);
    await sleep(1200);
    step('pdfExport', {
      success: Boolean(pdfResult && pdfResult.success),
      fileExists: fs.existsSync(pdfPath),
      fileSize: fs.existsSync(pdfPath) ? fs.statSync(pdfPath).size > 500 : false,
    });
    fs.rmSync(pdfPath, { force: true });

    await win.webContents.sendInputEvent({
      type: 'keyDown',
      keyCode: 'P',
      modifiers: ['control', 'shift'],
    });
    await sleep(500);
    const paletteText = await js(`
      const palette = document.querySelector('[data-command-palette]');
      return palette ? palette.innerText : '';
    `);
    step('commandPalette', {
      opened: Boolean(paletteText),
      hasSave: paletteText.includes('Сохранить активный'),
      hasFind: paletteText.includes('Найти и заменить'),
      hasVersions: paletteText.includes('Версии документа'),
    });
    await js(`
      const palette = document.querySelector('[data-command-palette]');
      if (palette && palette.parentElement) palette.parentElement.click();
    `);

    const rightsOff = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        /Права записи выданы/.test(item.textContent || '')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(400);
    const beforeNoRights = await documentState();
    const noRightsSend = await send('Перепиши резюме заново');
    const afterNoRights = await documentState();
    const noRightsRequest = [...out.requests].reverse().find((request) => request.scenario === 'norights');

    step('noRights', {
      toggled: Boolean(rightsOff),
      sent: noRightsSend.sent,
      idle: noRightsSend.idle,
      onlyReadTool:
        Boolean(noRightsRequest) &&
        noRightsRequest.toolNames.length > 0 &&
        noRightsRequest.toolNames.every((name) =>
          ['get_document', 'get_document_info', 'get_outline', 'get_block', 'get_range', 'search_document', 'get_selection', 'table_read', 'read_document'].includes(name)
        ),
      promptSaysNoWrite: Boolean(noRightsRequest && noRightsRequest.promptHasNoWrite),
      documentUnchanged: beforeNoRights.text === afterNoRights.text,
      honestReply: (await js('return document.body.innerText')).includes('Прав на запись'),
    });

    /* ---------- 8. Контракт запроса агента ---------- */
    const cleanupRequest = out.requests.find((request) => request.scenario === 'cleanup');
    const atsRequest = out.requests.find((request) => request.scenario === 'ats');
    step('agentRequest', {
      allToolsPassed:
        Boolean(cleanupRequest) &&
        cleanupRequest.toolNames.length >= 8 &&
        ['get_document', 'replace_document', 'format', 'page_setup'].every((name) =>
          cleanupRequest.toolNames.includes(name)
        ),
      documentNotAttached:
        Boolean(cleanupRequest) &&
        !String(cleanupRequest.lastUserContent).includes('Иван Иванов') &&
        !String(cleanupRequest.lastUserContent).includes('Москва'),
      promptHasReadWrite: Boolean(cleanupRequest && cleanupRequest.promptHasWrite),
      promptHasDocument: Boolean(cleanupRequest && cleanupRequest.promptMentionsDocument),
      atsUsedFullReplace:
        Boolean(atsRequest) && out.requests.some((request) => request.scenario === 'ats' && request.stepIndex === 1),
    });

    /* ---------- 9. «Заменить документ»: только текст документа ---------- */
    // Ответ без вызова инструментов — единственный, где кнопка замены имеет
    // смысл: отчёты агента уже применили правки и кнопку не показывают.
    const statementSend = await send('Подготовь заявление на отпуск');
    const replaceCount = await js(`
      return [...document.querySelectorAll('button')].filter((item) =>
        /Заменить документ/.test(item.textContent || '')
      ).length;
    `);
    step('replaceButton', { present: replaceCount > 0 });

    const replaceClicked = await js(`
      const buttons = [...document.querySelectorAll('button')].filter((item) =>
        /Заменить документ/.test(item.textContent || '')
      );
      const button = buttons[buttons.length - 1];
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(2000);
    const saveAfterReplace = await js(`
      const button = [...document.querySelectorAll('button')].find(
        (item) => (item.textContent || '').trim() === 'Сохранить' && !item.disabled
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(2500);
    const statementDoc = await readDocxXml(path.join(SANDBOX, DOC_NAME));
    step('replaceCleanContent', {
      sent: statementSend.sent,
      clicked: Boolean(replaceClicked),
      saved: Boolean(saveAfterReplace),
      statementKept: statementDoc.text.includes('Прошу предоставить отпуск'),
      preambleGone: !statementDoc.text.includes('Понял Вас'),
      metaGone: !statementDoc.text.includes('Рекомендуемые настройки'),
      fencesGone: !statementDoc.text.includes('```'),
      fileTitleGone: !statementDoc.text.includes('Резюме'),
      marginsKept: /<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"\/>/.test(
        statementDoc.xml
      ),
    });

    /* ---------- 11. «Вставить в документ»: текст реально вставляется ---------- */
    const insertSend = await send('Дай текст абзаца про отпуск для вставки');
    const insertClicked = await js(`
      const buttons = [...document.querySelectorAll('button')].filter((item) =>
        /Вставить в документ/.test(item.textContent || '')
      );
      let target = null;
      let targetSize = Infinity;
      for (const button of buttons) {
        let node = button.parentElement;
        for (let depth = 0; depth < 8 && node; depth += 1, node = node.parentElement) {
          const text = node.textContent || '';
          if (text.includes('Дополнительный абзац')) {
            if (text.length < targetSize) {
              target = button;
              targetSize = text.length;
            }
            break;
          }
        }
      }
      if (target) { target.click(); return true; }
      return false;
    `);
    await sleep(1500);
    const insertedState = await documentState();
    const saveAfterInsert = await js(`
      const button = [...document.querySelectorAll('button')].find(
        (item) => (item.textContent || '').trim() === 'Сохранить' && !item.disabled
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(2500);
    const insertedDoc = await readDocxXml(path.join(SANDBOX, DOC_NAME));
    step('insertIntoDocument', {
      sent: insertSend.sent,
      clicked: Boolean(insertClicked),
      saved: Boolean(saveAfterInsert),
      insertedKept: insertedDoc.text.includes('прошу согласовать отпуск'),
      previousTextKept: insertedDoc.text.includes('Прошу предоставить отпуск'),
      preambleGone: !insertedDoc.text.includes('Понял Вас'),
      metaTailGone:
        !insertedDoc.text.includes('абзац добавлен') &&
        !insertedDoc.text.includes('Что скажешь') &&
        !insertedDoc.text.includes('Если хочешь'),
      fencesGone: !insertedDoc.text.includes('```'),
      appendedAtEnd:
        insertedDoc.text.indexOf('Прошу предоставить отпуск') <
        insertedDoc.text.indexOf('прошу согласовать отпуск'),
      canvasHasInserted: insertedState.text.includes('прошу согласовать отпуск'),
    });

    /* ---------- 11. Шаблоны документов ---------- */
    const beforeTemplates = new Set(fs.readdirSync(SANDBOX));
    await win.webContents.sendInputEvent({
      type: 'keyDown',
      keyCode: 'P',
      modifiers: ['control', 'shift'],
    });
    await sleep(400);
    const templateCommandClicked = await js(`
      const item = [...document.querySelectorAll('[data-command-item]')].find((button) =>
        (button.textContent || '').includes('Создать из шаблона')
      );
      if (item) { item.click(); return true; }
      return false;
    `);
    await sleep(400);
    const templatePaletteShown = await js(`
      const palette = document.querySelector('[data-template-palette]');
      return palette ? palette.innerText : '';
    `);
    const templateClicked = await js(`
      const item = document.querySelector('[data-template-item="resume"]');
      if (item) { item.click(); return true; }
      return false;
    `);
    await sleep(3000);
    const templateFile = fs
      .readdirSync(SANDBOX)
      .find((name) => /^Резюме.*\.docx$/i.test(name) && !beforeTemplates.has(name));
    step('templateCreate', {
      commandOpened: Boolean(templateCommandClicked),
      paletteShown: Boolean(templatePaletteShown && templatePaletteShown.includes('Резюме')),
      templateClicked: Boolean(templateClicked),
      fileCreated: Boolean(templateFile),
    });

    /* ---------- 12. Запись в дополнительный файл после разрешения ---------- */
    const extraWriteToggled = await js(`
      const button = [...document.querySelectorAll('button')].find((item) =>
        (item.textContent || '').includes('Доп. файлы: чтение')
      );
      if (button) { button.click(); return true; }
      return false;
    `);
    await sleep(400);
    const extraWriteEnabled = await js(
      `return document.body.innerText.includes('Доп. файлы: запись')`
    );
    const extraWriteSend = await send('Запиши в третий файл новый текст');
    await sleep(2200);
    const thirdContent = fs.readFileSync(path.join(SANDBOX, 'Третий.md'), 'utf-8');
    step('extraDocumentWrite', {
      toggled: Boolean(extraWriteToggled) || Boolean(extraWriteEnabled),
      enabled: Boolean(extraWriteEnabled),
      sent: extraWriteSend.sent,
      idle: extraWriteSend.idle,
      written: thirdContent.includes('Изменено AI'),
    });

    const bookSend = await send('Запиши книгу с итогами');
    await sleep(2200);
    let bookState: any = null;
    try {
      const bookZip = await JSZip.loadAsync(fs.readFileSync(path.join(SANDBOX, 'Третья.xlsx')));
      const workbookXml = (await bookZip.file('xl/workbook.xml')?.async('string')) ?? '';
      const sheetXml =
        (await bookZip.file('xl/worksheets/sheet1.xml')?.async('string')) ?? '';
      bookState = {
        written: Boolean(bookZip.file('xl/workbook.xml')),
        sheetNamed: workbookXml.includes('Данные'),
        cells: /<c\b[^>]*>/.test(sheetXml),
        values: sheetXml.includes('Поле') && sheetXml.includes('Значение'),
      };
    } catch {
      bookState = { written: false, sheetNamed: false, cells: false, values: false };
    }
    step('extraWorkbookWrite', {
      sent: bookSend.sent,
      idle: bookSend.idle,
      ...bookState,
    });

    const imageUrlSend = await send('Вставь картинку по URL');
    await sleep(2500);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let imageUrlState: any = null;
    try {
      const createdDoc = fs
        .readdirSync(SANDBOX)
        .find((name) => /^Резюме.*\.docx$/i.test(name) && !beforeTemplates.has(name));
      if (!createdDoc) {
        imageUrlState = { documentFound: false, drawingSaved: false, mediaSaved: false };
      } else {
        const savedZip = await JSZip.loadAsync(
          fs.readFileSync(path.join(SANDBOX, createdDoc))
        );
        const documentXml =
          (await savedZip.file('word/document.xml')?.async('string')) ?? '';
        const media = Object.keys(savedZip.files).filter((name) =>
          /^word\/media\/.+\.(png|jpe?g|gif|webp)$/i.test(name)
        );
        imageUrlState = {
          documentFound: true,
          drawingSaved: documentXml.includes('<w:drawing>'),
          mediaSaved: media.length > 0,
        };
      }
    } catch {
      imageUrlState = { documentFound: false, drawingSaved: false, mediaSaved: false };
    }
    step('imageFromUrl', {
      sent: imageUrlSend.sent,
      idle: imageUrlSend.idle,
      ...imageUrlState,
    });

    const beforeRestore = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      return {
        image: Boolean(canvas && canvas.querySelector('img')),
        text: canvas ? canvas.innerText : '',
      };
    `);
    const restoreClick = await js(`
      const open = [...document.querySelectorAll('button')].find(
        (item) => (item.getAttribute('title') || '').includes('Версии документа')
      );
      if (open) open.click();
      await new Promise((resolve) => setTimeout(resolve, 300));
      const buttons = [...document.querySelectorAll('button')].filter(
        (item) => (item.textContent || '').trim() === 'Восстановить'
      );
      if (!buttons.length) {
        const close = [...document.querySelectorAll('button')].find(
          (item) => (item.getAttribute('title') || '') === 'Закрыть версии'
        );
        if (close) close.click();
        return { count: 0, clicked: false };
      }
      buttons[buttons.length - 1].click();
      return { count: buttons.length, clicked: true };
    `);
    await sleep(900);
    const afterRestore = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      return {
        image: Boolean(canvas && canvas.querySelector('img')),
        text: canvas ? canvas.innerText : '',
      };
    `);
    step('versionsRestore', {
      hadImage: Boolean(beforeRestore.image),
      versions: restoreClick.count,
      clicked: Boolean(restoreClick.clicked),
      imageGone: Boolean(beforeRestore.image) && !afterRestore.image,
      textKept: String(afterRestore.text || '').length > 20,
    });

    const tableFormatSend = await send('Сделай таблицу навыков');
    await sleep(3000);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let tableState: any = null;
    try {
      const latest = fs
        .readdirSync(SANDBOX)
        .filter((name) => /\.docx$/i.test(name))
        .map((name) => ({
          name,
          time: fs.statSync(path.join(SANDBOX, name)).mtimeMs,
        }))
        .sort((left, right) => right.time - left.time)[0];
      if (!latest) {
        tableState = { saved: false, table: false, width: false, border: false };
      } else {
        const savedZip = await JSZip.loadAsync(fs.readFileSync(path.join(SANDBOX, latest.name)));
        const documentXml =
          (await savedZip.file('word/document.xml')?.async('string')) ?? '';
        tableState = {
          saved: true,
          table: documentXml.includes('<w:tbl>') && documentXml.includes('JavaScript'),
          width: documentXml.includes('<w:tcW w:w="2000" w:type="dxa"/>'),
          border: documentXml.includes('w:color="FF0000"'),
        };
      }
    } catch {
      tableState = { saved: false, table: false, width: false, border: false };
    }
    step('tableFormatE2e', {
      sent: tableFormatSend.sent,
      idle: tableFormatSend.idle,
      ...tableState,
    });

    const longSend = await send('Длинный документ на двести блоков');
    await sleep(3500);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let longState: any = null;
    try {
      const longLatest = fs
        .readdirSync(SANDBOX)
        .filter((name) => /\.docx$/i.test(name))
        .map((name) => ({
          name,
          time: fs.statSync(path.join(SANDBOX, name)).mtimeMs,
        }))
        .sort((left, right) => right.time - left.time)[0];
      if (!longLatest) {
        longState = { saved: false, blocks: 0, edited: false, firstKept: false };
      } else {
        const longZip = await JSZip.loadAsync(fs.readFileSync(path.join(SANDBOX, longLatest.name)));
        const longXml =
          (await longZip.file('word/document.xml')?.async('string')) ?? '';
        longState = {
          saved: true,
          blocks: (longXml.match(/<w:p\b/g) || []).length,
          edited: longXml.includes('Блок 199 изменён'),
          firstKept: longXml.includes('Блок 0:'),
        };
      }
    } catch {
      longState = { saved: false, blocks: 0, edited: false, firstKept: false };
    }
    step('longDocumentE2e', {
      sent: longSend.sent,
      idle: longSend.idle,
      ...longState,
    });

    const styleSend = await send('Создай стиль врезки и примени к первому абзацу');
    await sleep(3000);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let styleState: any = null;
    try {
      const styleLatest = fs
        .readdirSync(SANDBOX)
        .filter((name) => /\.docx$/i.test(name))
        .map((name) => ({
          name,
          time: fs.statSync(path.join(SANDBOX, name)).mtimeMs,
        }))
        .sort((left, right) => right.time - left.time)[0];
      if (!styleLatest) {
        styleState = { saved: false, definition: false, applied: false };
      } else {
        const styleZip = await JSZip.loadAsync(
          fs.readFileSync(path.join(SANDBOX, styleLatest.name))
        );
        const definitionXml =
          (await styleZip.file('word/styles.xml')?.async('string')) ?? '';
        const documentXml =
          (await styleZip.file('word/document.xml')?.async('string')) ?? '';
        styleState = {
          saved: true,
          definition:
            definitionXml.includes('w:styleId="CustomAI"') &&
            definitionXml.includes('w:name w:val="Врезка AI"'),
          applied: documentXml.includes('w:pStyle w:val="CustomAI"'),
        };
      }
    } catch {
      styleState = { saved: false, definition: false, applied: false };
    }
    step('customStyleE2e', {
      sent: styleSend.sent,
      idle: styleSend.idle,
      ...styleState,
    });

    const annotationSend = await send('Добавь комментарий и сноску к первому абзацу');
    await sleep(3200);
    await win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    await sleep(2500);
    let annotationState: any = null;
    try {
      const annotationLatest = fs
        .readdirSync(SANDBOX)
        .filter((name) => /\.docx$/i.test(name))
        .map((name) => ({
          name,
          time: fs.statSync(path.join(SANDBOX, name)).mtimeMs,
        }))
        .sort((left, right) => right.time - left.time)[0];
      if (!annotationLatest) {
        annotationState = { saved: false, comment: false, footnote: false, anchors: false };
      } else {
        const annotationZip = await JSZip.loadAsync(
          fs.readFileSync(path.join(SANDBOX, annotationLatest.name))
        );
        const documentXml =
          (await annotationZip.file('word/document.xml')?.async('string')) ?? '';
        const commentsXml =
          (await annotationZip.file('word/comments.xml')?.async('string')) ?? '';
        const footnotesXml =
          (await annotationZip.file('word/footnotes.xml')?.async('string')) ?? '';
        annotationState = {
          saved: true,
          comment: commentsXml.includes('Проверить формулировку'),
          footnote: footnotesXml.includes('Источник для блока'),
          anchors:
            documentXml.includes('<w:commentRangeStart') &&
            documentXml.includes('<w:footnoteReference'),
        };
      }
    } catch {
      annotationState = { saved: false, comment: false, footnote: false, anchors: false };
    }
    step('annotationsE2e', {
      sent: annotationSend.sent,
      idle: annotationSend.idle,
      ...annotationState,
    });

    const wordZoomState = await js(`
      const canvas = document.querySelector('[contenteditable=true]');
      const target = canvas?.parentElement;
      if (!canvas || !target) return null;
      target.dispatchEvent(new WheelEvent('wheel', { deltaY: -200, ctrlKey: true, bubbles: true, cancelable: true }));
      await new Promise((resolve) => setTimeout(resolve, 120));
      const status = [...document.querySelectorAll('span')]
        .map((node) => node.textContent || '')
        .find((text) => text.includes('Масштаб:'));
      const zoom = canvas.style.zoom;
      const styleAttr = canvas.getAttribute('style') || '';
      const computed = getComputedStyle(canvas).zoom || '';
      target.dispatchEvent(new WheelEvent('wheel', { deltaY: 200, ctrlKey: true, bubbles: true, cancelable: true }));
      return { zoom, styleAttr, computed, status };
    `);
    step('wordZoom', {
      zoomed: /(?:^|;)\s*zoom:\s*(?:1\.[1-9]|[2-9])/.test(wordZoomState?.styleAttr || '') ||
        Number.parseFloat(wordZoomState?.computed || '1') > 1,
      status: /Масштаб: \d+%/.test(wordZoomState?.status || ''),
    });

    const renameMenuOpened = await js(`
      const nodes = [...document.querySelectorAll('div')]
        .filter((node) => (node.textContent || '').trim() === 'Третий.md');
      const node = nodes.find((item) => item.className.includes('cursor-pointer')) || nodes[0];
      if (!node) return false;
      node.dispatchEvent(new MouseEvent('contextmenu', {
        bubbles: true,
        cancelable: true,
        clientX: 120,
        clientY: 120,
      }));
      return true;
    `);
    await sleep(300);
    const renameMenuClicked = await js(`
      const button = [...document.querySelectorAll('button')]
        .find((item) => (item.textContent || '').trim() === 'Переименовать');
      if (!button) return false;
      button.click();
      return true;
    `);
    await sleep(300);
    const renameSubmitted = await js(`
      const input = [...document.querySelectorAll('input')]
        .find((item) => item.placeholder === 'Новое имя');
      if (!input) return false;
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      setter.call(input, 'Третий-renamed.md');
      input.dispatchEvent(new Event('input', { bubbles: true }));
      const button = [...document.querySelectorAll('button')]
        .find((item) => (item.textContent || '').trim() === 'Переименовать');
      if (!button) return false;
      button.click();
      return true;
    `);
    await sleep(1000);
    step('renameFile', {
      menu: Boolean(renameMenuOpened),
      clicked: Boolean(renameMenuClicked),
      submitted: Boolean(renameSubmitted),
      renamed:
        fs.existsSync(path.join(SANDBOX, 'Третий-renamed.md')) &&
        !fs.existsSync(path.join(SANDBOX, 'Третий.md')),
    });

    (out as any).summary = {
      requests: out.requests.length,
      scenarios: [...new Set(out.requests.map((request) => request.scenario))],
    };
  } catch (err) {
    out.errors.push(err && err.stack ? err.stack : String(err));
  }

  mock.server.close();
  imageServer.close();
  writeOut();
  app.exit(0);
}

app.whenReady().then(() => {
  try {
    fs.writeFileSync(path.join(SANDBOX, 'ready.txt'), 'app ready');
  } catch {}
  // Проверка не должна зависать в CI: через 5 минут фиксируем таймаут.
  const timer = setTimeout(() => {
    writeOut('timeout: проверка не завершилась за 300 секунд');
    app.exit(2);
  }, 300000);
  void main()
    .catch((err) => {
      writeOut(String(err && err.stack ? err.stack : err));
      app.exit(1);
    })
    .finally(() => clearTimeout(timer));
});
