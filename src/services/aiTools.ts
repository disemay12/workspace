/**
 * Инструменты ассистента для работы с открытым документом.
 *
 * Схемы отдаются OpenAI-совместимому API как `tools` (function calling).
 * Для моделей и провайдеров без поддержки tools те же операции доступны
 * текстовым протоколом workspace-action → document_apply.
 */
import type { FileType } from '../types';

export interface ToolDefinition {
  type: 'function';
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

const BLOCK_ID = {
  type: 'string',
  description: 'Идентификатор блока из get_document или результата предыдущего инструмента, например b3',
};

const HTML = {
  type: 'string',
  description:
    'HTML фрагмента: <p>, <h1>–<h3>, <ul>/<ol> с <li>, <table>, <blockquote>, <strong>, <em>, <u>, <s>, <br>. ' +
    'Оформление абзаца — inline style: text-align, font-family, font-size (pt), margin-left (pt), text-indent (pt). Без Markdown.',
};

const FORMAT_PROPERTIES: Record<string, unknown> = {
  block_id: {
    ...BLOCK_ID,
    description:
      'Блок для форматирования. Не указывайте вместе с find, если нужно отформатировать найденный текст во всём документе.',
  },
  find: {
    type: 'string',
    description: 'Точный текст внутри блока для точечного форматирования. Без block_id ищет по всему документу.',
  },
  heading: {
    type: 'string',
    description: 'Уровень заголовка: "1", "2", "3" или "0" — обычный абзац.',
  },
  bold: { type: 'boolean', description: 'true — жирный, false — снять жирный' },
  italic: { type: 'boolean', description: 'true — курсив, false — снять курсив' },
  underline: { type: 'boolean', description: 'true — подчёркивание, false — снять' },
  list: {
    type: 'string',
    description: 'Тип списка: "bullet" — маркированный, "numbered" — нумерованный, "none" — обычный абзац.',
  },
  align: {
    type: 'string',
    description: 'Выравнивание: "left", "center", "right", "justify".',
  },
  font: { type: 'string', description: 'Гарнитура, например "Calibri" или "Times New Roman".' },
  size_pt: { type: 'string', description: 'Размер шрифта в пунктах, например "14".' },
  indent_pt: { type: 'string', description: 'Отступ слева в пунктах, например "24".' },
  first_line_indent_pt: { type: 'string', description: 'Красная строка в пунктах, например "24".' },
};

const ALL_TOOLS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'get_document',
      description:
        'Возвращает структуру открытого документа: идентификаторы блоков и элементов списков, тип блока, текст и оформление. ' +
        'Вызывайте первым, если структура ещё неизвестна или документ мог измениться.',
      parameters: { type: 'object', properties: {}, additionalProperties: false },
    },
  },
  {
    type: 'function',
    function: {
      name: 'replace_document',
      description:
        'Полностью заменяет содержимое открытого документа. Используйте только когда меняется весь документ (переписать, сократить, сделать ATS-friendly).',
      parameters: {
        type: 'object',
        properties: { html: HTML },
        required: ['html'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'replace_block',
      description: 'Заменяет один блок или элемент списка. Для точечных правок текста, заголовка или пункта.',
      parameters: {
        type: 'object',
        properties: { block_id: BLOCK_ID, html: HTML },
        required: ['block_id', 'html'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_after',
      description: 'Вставляет новый блок или пункт списка после указанного блока.',
      parameters: {
        type: 'object',
        properties: { block_id: BLOCK_ID, html: HTML },
        required: ['block_id', 'html'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_block',
      description: 'Удаляет блок, элемент списка или таблицу по идентификатору.',
      parameters: {
        type: 'object',
        properties: { block_id: BLOCK_ID },
        required: ['block_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'find_replace',
      description:
        'Заменяет текст на другой. По умолчанию только первое вхождение; all=true — все. Регистр по умолчанию не важен.',
      parameters: {
        type: 'object',
        properties: {
          find: { type: 'string', description: 'Что заменить' },
          replace: { type: 'string', description: 'На что заменить' },
          all: { type: 'boolean', description: 'Заменить все вхождения' },
        },
        required: ['find', 'replace'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'format',
      description:
        'Меняет оформление блока или найденного фрагмента: заголовок H1–H3, жирный, курсив, подчёркивание, список, выравнивание, шрифт, размер, отступы.',
      parameters: {
        type: 'object',
        properties: FORMAT_PROPERTIES,
        additionalProperties: false,
      },
    },
  },
];

/** Полный набор для Word, урезанный — для Markdown/текста, только чтение — без прав. */
export function buildDocumentTools(writeAccess: boolean, tabType: FileType | string): ToolDefinition[] {
  if (!writeAccess) return ALL_TOOLS.filter((tool) => tool.function.name === 'get_document');
  if (tabType === 'word') return ALL_TOOLS;
  return ALL_TOOLS.filter((tool) =>
    ['get_document', 'replace_document', 'find_replace'].includes(tool.function.name)
  );
}

/**
 * Нативный function calling есть только у OpenAI-совместимых провайдеров.
 * Anthropic и Gemini используют другой формат — для них работает текстовый
 * протокол document_apply.
 */
export function supportsNativeDocumentTools(provider: string): boolean {
  return provider !== 'anthropic' && provider !== 'gemini';
}

const TYPE_LABEL: Record<string, string> = {
  word: 'Word (.docx)',
  markdown: 'Markdown',
  richtext: 'текст',
  text: 'текст',
};

export interface DocumentPromptContext {
  name: string;
  type: FileType | string;
  writeAccess: boolean;
  /** Модель без function calling: правки через document_apply. */
  fallback?: boolean;
}

/** Системный блок про открытый документ: права, инструменты, правила. */
export function buildDocumentSystemPrompt(context: DocumentPromptContext): string {
  const typeLabel = TYPE_LABEL[context.type] ?? context.type;
  const header = [
    'ОТКРЫТЫЙ ДОКУМЕНТ',
    `- Файл: «${context.name}» (${typeLabel}).`,
    `- Права записи: ${context.writeAccess ? 'ЕСТЬ' : 'НЕТ'}.`,
  ];

  if (!context.writeAccess) {
    return [
      ...header,
      '- Инструменты записи недоступны. Документ менять нельзя.',
      '- Если просят изменить, оформить, переписать или удалить что-то в документе — одной фразой скажи, что нужны права записи (кнопка «Предоставить права на запись» в панели AI), и не меняй документ.',
      context.fallback
        ? '- Содержимое документа приложено к сообщению — используй его для ответов на вопросы.'
        : '- Для вопросов по содержимому используй get_document.',
      '- Не создавай новые файлы вместо правки документа.',
    ].join('\n');
  }

  const rules = [...header, '- Работай только с этим документом.'];

  if (context.fallback) {
    rules.push(
      '- Function calling недоступен. Содержимое документа приложено к сообщению выше — не вызывай get_document, работай с приложенным текстом.',
      '- Любые правки возвращай только блоком workspace-action с действием document_apply:',
      '```workspace-action',
      '{"action":"document_apply","operations":[{"tool":"replace_document","html":"<h1>Заголовок</h1><p>Текст</p>"}]}',
      '```',
      '- Для переписывания и оформления используй replace_document с полным HTML документа, для точечных замен — find_replace. Блочные операции replace_block/insert_after/delete_block применяй только когда точно знаешь id блока.',
      '- После блока — одна короткая строка отчёта. Не вставляй текст документа в обычное сообщение.'
    );
  } else {
    rules.push(
      '- Работай инструментами: get_document, replace_document, replace_block, insert_after, delete_block, find_replace, format.',
      '- Сначала вызови get_document, если структура ещё не известна. Идентификаторы блоков (b1, b2, …) бери из последнего ответа get_document или из результата предыдущего инструмента.',
      '- При просьбе изменить, оформить, переписать, сократить, удалить лишнее или поправить текст — вызывай инструменты. Не пиши готовый текст документа в чат.'
    );
  }

  rules.push(
    '- Не создавай новые файлы (create_file, create_table, create_document) и не предлагай скачать документ: правка идёт в открытый файл.',
    '- Не отвечай «не могу редактировать документ Word»: правка применяется сразу.',
    '- В документ клади только содержимое документа: без вступлений («Вот вариант…»), пояснений, вопросов, советов и списков «что изменено».',
    '- В чат верни короткий отчёт: 1–3 пункта о том, что сделано. Текст документа в чат не дублируй.',
    context.fallback
      ? '- Для точечных правок используй find_replace; replace_document — когда меняется весь документ.'
      : '- Для точечных правок используй replace_block / insert_after / delete_block / find_replace / format. replace_document — только когда меняется весь документ.',
    '- Формат HTML: <h1>–<h3>, <p>, <ul>/<ol> с <li>, <table>, <blockquote>, <strong>, <em>, <u>, <s>, <br>. Оформление абзаца — inline style: text-align, font-family, font-size (pt), margin-left (pt), text-indent (pt). Markdown внутри HTML не используй.'
  );

  if (context.type !== 'word') {
    rules.push(
      '- Для этого типа файла доступны только get_document, replace_document и find_replace: содержимое хранится как текст, блочное форматирование не поддерживается.'
    );
  }

  return rules.join('\n');
}
