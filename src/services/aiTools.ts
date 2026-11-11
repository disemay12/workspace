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
  color: { type: 'string', description: 'Цвет текста: "#RRGGBB" или имя цвета.' },
  highlight: { type: 'string', description: 'Цвет фона/выделения: "#RRGGBB" или имя цвета.' },
  fill: { type: 'string', description: 'Заливка всего блока: "#RRGGBB" или имя цвета.' },
  border: { type: 'boolean', description: 'true — включить рамку блока, false — снять.' },
  border_color: { type: 'string', description: 'Цвет рамки блока: "#RRGGBB".' },
  border_width_pt: { type: 'string', description: 'Толщина рамки в пунктах, например "1".' },
  border_style: { type: 'string', description: 'Стиль рамки: solid, dashed, dotted, double.' },
  link: {
    type: 'string',
    description: 'URL ссылки для найденного фрагмента (http/https/mailto/tel). Требуется find.',
  },
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
      name: 'get_outline',
      description:
        'Возвращает структуру документа постранично: id блоков, тип, текст и оформление. Для длинного документа вызывайте с offset, пока nextOffset не станет null.',
      parameters: {
        type: 'object',
        properties: {
          offset: { type: 'string', description: 'Номер первого блока, с которого начать (0 по умолчанию).' },
          limit: { type: 'string', description: 'Сколько блоков вернуть, 1–100 (40 по умолчанию).' },
        },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_block',
      description: 'Возвращает полный HTML и текст одного блока. Используйте перед точечной правкой, если из outline непонятна разметка.',
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
      name: 'get_range',
      description: 'Возвращает HTML и текст диапазона внутри блока по смещениям символов. Смещения считаются по тексту блока без разметки.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          start: { type: 'string', description: 'Начало диапазона, номер символа от 0.' },
          end: { type: 'string', description: 'Конец диапазона, не включается.' },
        },
        required: ['block_id', 'start', 'end'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'search_document',
      description: 'Ищет текст по всему документу и возвращает блоки и смещения совпадений. Используйте, когда фраза может быть в любом месте.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Текст для поиска.' },
          match_case: { type: 'boolean', description: 'Учитывать регистр.' },
          limit: { type: 'string', description: 'Сколько совпадений вернуть, 1–100 (20 по умолчанию).' },
        },
        required: ['query'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_selection',
      description:
        'Возвращает текст, HTML и точные смещения выделенного пользователем фрагмента. Вызывайте, когда просят изменить «выделенное», «этот фрагмент» или «здесь».',
      parameters: { type: 'object', properties: {}, additionalProperties: false },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_document_info',
      description:
        'Возвращает служебные части Word: колонтитулы, сноски, комментарии, именованные стили, число секций и список частей пакета. Используйте для вопросов о колонтитулах и оформлении, которые не видны в теле документа.',
      parameters: { type: 'object', properties: {}, additionalProperties: false },
    },
  },
  {
    type: 'function',
    function: {
      name: 'read_document',
      description:
        'Читает другой файл рабочей папки (Word, Excel, PowerPoint, Markdown, текст) без права записи. Путь можно указать относительно корня workspace.',
      parameters: {
        type: 'object',
        properties: {
          path: {
            type: 'string',
            description: 'Путь к файлу: относительный от корня workspace или абсолютный внутри него.',
          },
        },
        required: ['path'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'write_document',
      description:
        'Записывает другой файл workspace (Word, Excel или Markdown/текст). Доступно только после явного разрешения записи в дополнительные файлы.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string', description: 'Путь к файлу внутри workspace.' },
          html: { type: 'string', description: 'HTML содержимого для .docx.' },
          text: { type: 'string', description: 'Текст для Markdown/текстового файла.' },
          columns: {
            type: 'array',
            items: { type: 'string' },
            description: 'Заголовки столбцов для .xlsx.',
          },
          rows: {
            type: 'array',
            items: { type: 'array', items: { type: 'string' } },
            description: 'Строки для .xlsx.',
          },
          sheet: { type: 'string', description: 'Имя листа для .xlsx.' },
        },
        required: ['path'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'replace_document',
      description:
        'Полностью заменяет содержимое открытого документа. Только когда меняется весь документ (переписать, сократить, сделать ATS-friendly). ' +
        'Для одной строки, фразы, даты, имени или пункта используйте find_replace или replace_block — не переписывайте документ целиком.',
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
      name: 'insert_before',
      description: 'Вставляет новый блок или пункт списка перед указанным блоком.',
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
      name: 'append_to_block',
      description:
        'Дописывает текст в конец блока или новые пункты <li> в конец списка, не создавая новый абзац.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          html: { ...HTML, description: 'Inline HTML для абзаца или <li> для списка.' },
        },
        required: ['block_id', 'html'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'move_block',
      description:
        'Перемещает блок: перед/после другого блока или на позицию index среди верхнеуровневых блоков.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          before_id: { type: 'string', description: 'Переместить перед этим блоком.' },
          after_id: { type: 'string', description: 'Переместить после этого блока.' },
          index: { type: 'string', description: 'Позиция среди верхнеуровневых блоков, 0 — начало.' },
        },
        required: ['block_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'duplicate_block',
      description: 'Дублирует блок сразу после оригинала. Копия получает новый id.',
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
      name: 'split_block',
      description: 'Разделяет абзац или пункт списка на два по смещению символа.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          offset: { type: 'string', description: 'Номер символа, с которого начинается второй блок.' },
        },
        required: ['block_id', 'offset'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'merge_blocks',
      description: 'Объединяет блок со следующим (или с указанным next_id), сохраняя оформление первого.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          next_id: { type: 'string', description: 'Необязательно: id второго блока.' },
        },
        required: ['block_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'table_read',
      description: 'Возвращает таблицу по ячейкам: строки, столбцы, текст, объединения, плюс полный HTML.',
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
      name: 'table_insert_row',
      description: 'Вставляет строку в таблицу. index — позиция с 0; без index строка добавляется в конец.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          index: { type: 'string', description: 'Позиция строки, 0 — первая.' },
          cells: {
            type: 'array',
            items: { type: 'string' },
            description: 'Необязательные тексты ячеек по порядку.',
          },
        },
        required: ['block_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'table_delete_row',
      description: 'Удаляет строку таблицы по индексу.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          index: { type: 'string', description: 'Позиция строки, 0 — первая.' },
        },
        required: ['block_id', 'index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'table_insert_column',
      description: 'Вставляет столбец во все строки таблицы.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          index: { type: 'string', description: 'Позиция столбца, 0 — первый.' },
        },
        required: ['block_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'table_delete_column',
      description: 'Удаляет столбец таблицы по индексу.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          index: { type: 'string', description: 'Позиция столбца, 0 — первый.' },
        },
        required: ['block_id', 'index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'table_set_cell',
      description: 'Заменяет содержимое ячейки таблицы по строке и столбцу (с 0).',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          row: { type: 'string', description: 'Строка, с 0.' },
          col: { type: 'string', description: 'Столбец, с 0.' },
          html: HTML,
        },
        required: ['block_id', 'row', 'col', 'html'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'table_merge_cells',
      description: 'Объединяет ячейки таблицы, начиная с row/col, на rowspan строк и colspan столбцов.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          row: { type: 'string', description: 'Строка, с 0.' },
          col: { type: 'string', description: 'Столбец, с 0.' },
          rowspan: { type: 'string', description: 'Сколько строк объединить.' },
          colspan: { type: 'string', description: 'Сколько столбцов объединить.' },
        },
        required: ['block_id', 'row', 'col'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_image',
      description:
        'Вставляет изображение из файла workspace или data URL. Без block_id — в конец документа, с block_id — после блока.',
      parameters: {
        type: 'object',
        properties: {
          src: {
            type: 'string',
            description:
              'Путь к изображению внутри workspace, data:image/...;base64,... или http(s) URL.',
          },
          block_id: { ...BLOCK_ID, description: 'Необязательно: вставить после этого блока.' },
          alt: { type: 'string', description: 'Альтернативный текст.' },
          width: { type: 'string', description: 'Ширина в пикселях.' },
          height: { type: 'string', description: 'Высота в пикселях.' },
          align: { type: 'string', description: 'Выравнивание: left, center, right.' },
        },
        required: ['src'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'update_image',
      description: 'Меняет размер, выравнивание или alt изображения в блоке.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          width: { type: 'string', description: 'Ширина в пикселях.' },
          height: { type: 'string', description: 'Высота в пикселях.' },
          alt: { type: 'string', description: 'Альтернативный текст.' },
          align: { type: 'string', description: 'Выравнивание: left, center, right.' },
        },
        required: ['block_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_image',
      description: 'Удаляет изображение из блока или первое изображение документа.',
      parameters: {
        type: 'object',
        properties: { block_id: BLOCK_ID },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_page_break',
      description: 'Вставляет разрыв страницы после блока или в конец документа.',
      parameters: {
        type: 'object',
        properties: { block_id: { ...BLOCK_ID, description: 'Необязательно: после этого блока.' } },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_section_break',
      description: 'Вставляет разрыв секции после блока или в конец документа.',
      parameters: {
        type: 'object',
        properties: { block_id: { ...BLOCK_ID, description: 'Необязательно: после этого блока.' } },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_section_break',
      description:
        'Удаляет разрыв секции по block_id; без block_id — первый найденный разрыв.',
      parameters: {
        type: 'object',
        properties: { block_id: BLOCK_ID },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_paragraph_format',
      description:
        'Задаёт межстрочный интервал, отступы до/после, запрет разрыва и разрыв перед абзацем. Для всего документа используйте page_setup.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          line_spacing: {
            type: 'string',
            description: 'Множитель: "1", "1.5", "2" или число.',
          },
          space_before_pt: { type: 'string', description: 'Отступ перед абзацем в пунктах.' },
          space_after_pt: { type: 'string', description: 'Отступ после абзаца в пунктах.' },
          keep_with_next: { type: 'boolean', description: 'Не разрывать абзац со следующим.' },
          keep_together: { type: 'boolean', description: 'Не разрывать абзац между страницами.' },
          page_break_before: { type: 'boolean', description: 'Начать абзац с новой страницы.' },
          widow_control: {
            type: 'boolean',
            description: 'Запретить одиночные строки абзаца на странице.',
          },
        },
        required: ['block_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_toc',
      description:
        'Вставляет поле оглавления Word (TOC) в начало документа или после указанного блока. Word обновляет его при открытии.',
      parameters: {
        type: 'object',
        properties: { block_id: { ...BLOCK_ID, description: 'Необязательно: после этого блока.' } },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_tabs',
      description: 'Задаёт позиции табуляции абзаца в пунктах.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          positions: {
            type: 'string',
            description: 'Позиции через запятую или пробел, например "24,48".',
          },
        },
        required: ['block_id', 'positions'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_paragraph_style',
      description:
        'Назначает именованный стиль абзаца (например Title, Subtitle, Quote, Heading1, Normal). Стиль должен существовать в документе.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          style: { type: 'string', description: 'Имя или styleId стиля Word.' },
        },
        required: ['block_id', 'style'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'define_style',
      description:
        'Создаёт пользовательский стиль абзаца и сохраняет его определение в styles.xml. block_id применяет стиль сразу.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          name: { type: 'string', description: 'Имя стиля, например «Врезка».' },
          style_id: { type: 'string', description: 'Необязательный styleId.' },
          based_on: { type: 'string', description: 'Базовый стиль, по умолчанию Normal.' },
          font: { type: 'string', description: 'Гарнитура.' },
          size_pt: { type: 'string', description: 'Кегль в пунктах.' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          color: { type: 'string', description: 'Цвет текста #RRGGBB.' },
          align: {
            type: 'string',
            enum: ['left', 'center', 'right', 'justify'],
            description: 'Выравнивание.',
          },
          line_height: { type: 'string', description: 'Множитель межстрочного интервала.' },
          space_before_pt: { type: 'string', description: 'Отступ до абзаца в пунктах.' },
          space_after_pt: { type: 'string', description: 'Отступ после абзаца в пунктах.' },
        },
        required: ['name'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_header_footer',
      description:
        'Меняет текст верхнего или нижнего колонтитула. Сначала вызовите get_document_info, чтобы увидеть текущие колонтитулы.',
      parameters: {
        type: 'object',
        properties: {
          kind: { type: 'string', description: '"header" — верхний, "footer" — нижний.' },
          type: {
            type: 'string',
            description: 'Тип колонтитула: "default", "first" (первая страница) или "even" (чётные).',
          },
          text: {
            type: 'string',
            description:
              'Новый текст; переносы строк — новые абзацы. {{page}} — номер страницы, {{pages}} — всего страниц.',
          },
        },
        required: ['kind', 'text'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_header_footer',
      description: 'Удаляет верхний или нижний колонтитул выбранного типа.',
      parameters: {
        type: 'object',
        properties: {
          kind: { type: 'string', description: '"header" или "footer".' },
          type: { type: 'string', description: '"default", "first" или "even".' },
        },
        required: ['kind'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'table_unmerge_cells',
      description: 'Снимает объединение с ячейки, восстанавливая пустые ячейки.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          row: { type: 'string', description: 'Строка, с 0.' },
          col: { type: 'string', description: 'Столбец, с 0.' },
        },
        required: ['block_id', 'row', 'col'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'table_set_column_width',
      description: 'Задаёт ширину столбца таблицы в пунктах или пикселях.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          col: { type: 'string', description: 'Столбец, с 0.' },
          width_pt: { type: 'string', description: 'Ширина в пунктах.' },
          width_px: { type: 'string', description: 'Ширина в пикселях.' },
        },
        required: ['block_id', 'col'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'table_set_borders',
      description:
        'Задаёт границы таблицы: цвет, толщину и стиль. style="none" убирает границы.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          color: { type: 'string', description: 'Цвет границ, например #000000.' },
          size_pt: { type: 'string', description: 'Толщина в пунктах.' },
          style: {
            type: 'string',
            enum: ['single', 'double', 'dashed', 'dotted', 'none'],
            description: 'Стиль границ.',
          },
        },
        required: ['block_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'refresh_toc',
      description:
        'Пересобирает кэш оглавления из текущих заголовков документа; страницы обновит Word.',
      parameters: { type: 'object', properties: {}, additionalProperties: false },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_field',
      description:
        'Вставляет поле Word с номером страницы (PAGE) или общим числом страниц (NUMPAGES).',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          kind: {
            type: 'string',
            enum: ['PAGE', 'NUMPAGES'],
            description: 'Тип поля.',
          },
          at: {
            type: 'string',
            enum: ['after', 'before', 'end'],
            description: 'Куда вставить относительно block_id; end — в конец документа.',
          },
        },
        required: ['kind'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_bookmark',
      description:
        'Добавляет закладку Word: оборачивает найденный текст или ставит метку в начало блока.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          name: { type: 'string', description: 'Имя закладки без пробелов.' },
          find: { type: 'string', description: 'Текст, который станет закладкой.' },
        },
        required: ['name'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'add_comment',
      description:
        'Добавляет комментарий Word к найденному тексту блока или к началу блока.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          find: { type: 'string', description: 'Фрагмент текста для комментария.' },
          text: { type: 'string', description: 'Текст комментария.' },
          author: { type: 'string', description: 'Автор, по умолчанию AI.' },
        },
        required: ['block_id', 'text'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'add_footnote',
      description:
        'Добавляет сноску Word к найденному тексту блока или к началу блока.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          find: { type: 'string', description: 'Фрагмент текста, к которому относится сноска.' },
          text: { type: 'string', description: 'Текст сноски.' },
        },
        required: ['block_id', 'text'],
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
        'Заменяет текст на другой. По умолчанию только первое вхождение; all=true — все. Регистр по умолчанию не важен. ' +
        'Работает и когда Word разбил фразу на несколько фрагментов форматирования. Используйте для точечной правки строки вместо replace_document.',
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
      name: 'replace_range',
      description:
        'Заменяет диапазон символов внутри блока. Точный инструмент для правки части абзаца: смещения берите из get_range или search_document.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          start: { type: 'string', description: 'Начало диапазона, номер символа от 0.' },
          end: { type: 'string', description: 'Конец диапазона, не включается.' },
          html: { ...HTML, description: 'Новый текст (inline HTML): <strong>, <em>, <u>, <span style>, <br>.' },
        },
        required: ['block_id', 'start', 'end', 'html'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_range',
      description: 'Удаляет диапазон символов внутри блока, не трогая соседний текст.',
      parameters: {
        type: 'object',
        properties: {
          block_id: BLOCK_ID,
          start: { type: 'string', description: 'Начало диапазона, номер символа от 0.' },
          end: { type: 'string', description: 'Конец диапазона, не включается.' },
        },
        required: ['block_id', 'start', 'end'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'replace_selection',
      description:
        'Заменяет выделенный пользователем фрагмент. Сначала вызовите get_selection, чтобы увидеть текст и смещения.',
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
      name: 'insert_at_cursor',
      description: 'Вставляет фрагмент в позицию каретки или в конец выделения.',
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
  {
    type: 'function',
    function: {
      name: 'page_setup',
      description:
        'Оформляет страницу и основной текст всего документа: поля, размер бумаги, шрифт, кегль, межстрочный интервал, выравнивание и красная строка. ' +
        'Вызывайте вместо советов «настройте Word»: это меняет сам .docx. Для заявления, письма, договора, приказа используйте template "russian_business".',
      parameters: {
        type: 'object',
        properties: {
          template: {
            type: 'string',
            description:
              'Готовый шаблон: "russian_business" — поля 3/1,5/2/2 см, Times New Roman 14, одинарный интервал, по ширине, красная строка 1,25 см.',
          },
          margins: {
            type: 'object',
            description: 'Поля страницы в сантиметрах.',
            properties: {
              top: { type: 'string', description: 'Верхнее поле, см' },
              right: { type: 'string', description: 'Правое поле, см' },
              bottom: { type: 'string', description: 'Нижнее поле, см' },
              left: { type: 'string', description: 'Левое поле, см' },
            },
            additionalProperties: false,
          },
          font: { type: 'string', description: 'Гарнитура основного текста, например "Times New Roman".' },
          size_pt: { type: 'string', description: 'Кегль основного текста в пунктах, например "14".' },
          line_spacing: {
            type: 'string',
            description: 'Межстрочный интервал: "single", "1.5" или число.',
          },
          align: { type: 'string', description: 'Выравнивание абзацев: "left", "center", "right", "justify".' },
          first_line_indent_cm: { type: 'string', description: 'Красная строка в сантиметрах, например "1.25".' },
        },
        additionalProperties: false,
      },
    },
  },
];

/** Инструменты открытой книги Excel. */
const EXCEL_TOOLS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'excel_read',
      description:
        'Читает лист или диапазон книги. Возвращает значения, формулы и отображаемый текст ячеек.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Имя листа; по умолчанию активный.' },
          range: { type: 'string', description: 'Диапазон A1, например A1:D20. Без него — заполненная область.' },
          limit: { type: 'string', description: 'Максимум ячеек, 1–1200 (400 по умолчанию).' },
        },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_get_selection',
      description: 'Возвращает текущее выделение пользователя на активном листе.',
      parameters: { type: 'object', properties: {}, additionalProperties: false },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_write_cells',
      description:
        'Записывает значения или формулы в ячейки. Формула начинается с "=". Можно передать список ячеек или диапазон и матрицу значений.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Лист; по умолчанию активный.' },
          cells: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                address: { type: 'string', description: 'Адрес A1, например B4.' },
                value: { type: 'string', description: 'Значение или формула.' },
              },
              required: ['address', 'value'],
              additionalProperties: false,
            },
            description: 'Список ячеек для записи.',
          },
          range: { type: 'string', description: 'Диапазон A1 для матрицы values.' },
          values: {
            type: 'array',
            items: { type: 'array', items: { type: 'string' } },
            description: 'Матрица значений для range.',
          },
        },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_add_sheet',
      description: 'Добавляет лист в книгу.',
      parameters: {
        type: 'object',
        properties: { name: { type: 'string', description: 'Имя нового листа.' } },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_rename_sheet',
      description: 'Переименовывает лист.',
      parameters: {
        type: 'object',
        properties: {
          old_name: { type: 'string', description: 'Текущее имя.' },
          new_name: { type: 'string', description: 'Новое имя.' },
        },
        required: ['old_name', 'new_name'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_delete_sheet',
      description: 'Удаляет лист, если в книге останется хотя бы один лист.',
      parameters: {
        type: 'object',
        properties: { name: { type: 'string', description: 'Имя листа.' } },
        required: ['name'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_set_active_sheet',
      description: 'Переключает активный лист книги.',
      parameters: {
        type: 'object',
        properties: { name: { type: 'string', description: 'Имя листа.' } },
        required: ['name'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_insert_rows',
      description: 'Вставляет строки в лист. index — номер строки с 1, count — сколько.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Лист; по умолчанию активный.' },
          index: { type: 'string', description: 'Номер строки с 1.' },
          count: { type: 'string', description: 'Сколько строк вставить (по умолчанию 1).' },
        },
        required: ['index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_delete_rows',
      description: 'Удаляет строки. index — номер первой строки с 1.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Лист; по умолчанию активный.' },
          index: { type: 'string', description: 'Номер строки с 1.' },
          count: { type: 'string', description: 'Сколько строк удалить (по умолчанию 1).' },
        },
        required: ['index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_insert_columns',
      description: 'Вставляет столбцы. index — номер столбца с 1.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Лист; по умолчанию активный.' },
          index: { type: 'string', description: 'Номер столбца с 1.' },
          count: { type: 'string', description: 'Сколько столбцов вставить (по умолчанию 1).' },
        },
        required: ['index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_delete_columns',
      description: 'Удаляет столбцы. index — номер первого столбца с 1.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Лист; по умолчанию активный.' },
          index: { type: 'string', description: 'Номер столбца с 1.' },
          count: { type: 'string', description: 'Сколько столбцов удалить (по умолчанию 1).' },
        },
        required: ['index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_merge_cells',
      description: 'Объединяет диапазон ячеек, например A1:B1.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Лист; по умолчанию активный.' },
          range: { type: 'string', description: 'Диапазон A1:B1.' },
        },
        required: ['range'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_unmerge_cells',
      description: 'Снимает объединение с диапазона.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Лист; по умолчанию активный.' },
          range: { type: 'string', description: 'Диапазон A1:B1.' },
        },
        required: ['range'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'excel_format_range',
      description: 'Форматирует диапазон: жирный, курсив, заливка, цвет, размер, выравнивание, перенос.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Лист; по умолчанию активный.' },
          range: { type: 'string', description: 'Диапазон A1:D10.' },
          bold: { type: 'boolean', description: 'Жирный.' },
          italic: { type: 'boolean', description: 'Курсив.' },
          fill: { type: 'string', description: 'Заливка #RRGGBB.' },
          color: { type: 'string', description: 'Цвет текста #RRGGBB.' },
          font_size: { type: 'string', description: 'Размер шрифта в pt.' },
          align: { type: 'string', description: 'left, center или right.' },
          wrap: { type: 'boolean', description: 'Перенос текста.' },
        },
        required: ['range'],
        additionalProperties: false,
      },
    },
  },
];

/** Инструменты открытой презентации. */
const PPTX_TOOLS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'pptx_read',
      description: 'Читает презентацию: слайды, фигуры с текстом, заметки.',
      parameters: {
        type: 'object',
        properties: { slide: { type: 'string', description: 'Необязательно: номер слайда с 1.' } },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_get_selection',
      description: 'Возвращает выделенную текстовую фигуру текущего слайда.',
      parameters: { type: 'object', properties: {}, additionalProperties: false },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_set_text',
      description:
        'Заменяет текст фигуры на слайде. Переносы строк — новые абзацы. Сначала вызовите pptx_read, чтобы получить shape_id.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'string', description: 'Номер слайда с 1.' },
          shape_id: { type: 'string', description: 'id фигуры из pptx_read.' },
          text: { type: 'string', description: 'Новый текст фигуры; \\n — новый абзац.' },
        },
        required: ['slide', 'shape_id', 'text'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_add_slide',
      description: 'Добавляет слайд после указанного номера (с 1) или в конец.',
      parameters: {
        type: 'object',
        properties: {
          index: { type: 'string', description: 'После какого слайда вставить, с 1.' },
          title: { type: 'string', description: 'Заголовок нового слайда.' },
        },
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_delete_slide',
      description: 'Удаляет слайд по номеру с 1.',
      parameters: {
        type: 'object',
        properties: { index: { type: 'string', description: 'Номер слайда с 1.' } },
        required: ['index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_move_slide',
      description: 'Меняет порядок слайдов.',
      parameters: {
        type: 'object',
        properties: {
          from: { type: 'string', description: 'Текущий номер слайда с 1.' },
          to: { type: 'string', description: 'Новый номер с 1.' },
        },
        required: ['from', 'to'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_add_text_box',
      description: 'Добавляет текстовую фигуру на слайд.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'string', description: 'Номер слайда с 1.' },
          text: { type: 'string', description: 'Текст фигуры.' },
        },
        required: ['slide', 'text'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_insert_image',
      description:
        'Вставляет изображение на слайд. src — путь внутри workspace, data URL или http(s) URL.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'string', description: 'Номер слайда с 1.' },
          src: { type: 'string', description: 'Путь, data URL или http(s) URL.' },
          width: { type: 'string', description: 'Ширина в пикселях.' },
          height: { type: 'string', description: 'Высота в пикселях.' },
          x: { type: 'string', description: 'Отступ слева в пикселях.' },
          y: { type: 'string', description: 'Отступ сверху в пикселях.' },
          alt: { type: 'string', description: 'Описание изображения.' },
        },
        required: ['slide', 'src'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_set_layout',
      description:
        'Меняет макет слайда по имени или номеру из pptx_read. Содержимое слайда сохраняется.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'string', description: 'Номер слайда с 1.' },
          layout: { type: 'string', description: 'Имя или номер макета.' },
        },
        required: ['slide', 'layout'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_set_notes',
      description:
        'Заменяет заметки докладчика слайда. Пустая строка очищает заметки; переносы строк — абзацы.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'string', description: 'Номер слайда с 1.' },
          text: { type: 'string', description: 'Текст заметок.' },
        },
        required: ['slide', 'text'],
        additionalProperties: false,
      },
    },
  },
];

/** Инструменты Markdown/текста. */
const MARKDOWN_TOOLS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'md_read',
      description: 'Читает Markdown/текстовый документ целиком.',
      parameters: { type: 'object', properties: {}, additionalProperties: false },
    },
  },
  {
    type: 'function',
    function: {
      name: 'md_replace',
      description:
        'Заменяет текст в Markdown без пересборки разметки. По умолчанию первое вхождение; all=true — все.',
      parameters: {
        type: 'object',
        properties: {
          find: { type: 'string', description: 'Точный текст для поиска.' },
          replace: { type: 'string', description: 'Новый текст.' },
          all: { type: 'boolean', description: 'Заменить все вхождения.' },
        },
        required: ['find', 'replace'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'md_append',
      description: 'Дописывает Markdown-текст в конец документа.',
      parameters: {
        type: 'object',
        properties: { text: { type: 'string', description: 'Текст для добавления.' } },
        required: ['text'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'md_replace_document',
      description: 'Полностью заменяет Markdown/текстовый документ.',
      parameters: {
        type: 'object',
        properties: { text: { type: 'string', description: 'Новый текст документа.' } },
        required: ['text'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'md_set_heading',
      description:
        'Делает строку заголовком Markdown. find — фрагмент строки; пусто — первый непустой абзац.',
      parameters: {
        type: 'object',
        properties: {
          find: { type: 'string', description: 'Фрагмент строки для поиска.' },
          level: {
            type: 'string',
            enum: ['1', '2', '3', '4', '5', '6'],
            description: 'Уровень заголовка.',
          },
        },
        required: ['level'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'md_toggle_list',
      description:
        'Превращает строку в пункт списка или убирает маркер, если он уже есть.',
      parameters: {
        type: 'object',
        properties: {
          find: { type: 'string', description: 'Фрагмент строки.' },
          style: {
            type: 'string',
            enum: ['bullet', 'numbered'],
            description: 'bullet — «-», numbered — «1.».',
          },
        },
        required: ['find', 'style'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'md_insert_link',
      description:
        'Вставляет Markdown-ссылку вместо найденного текста или в конец документа.',
      parameters: {
        type: 'object',
        properties: {
          find: { type: 'string', description: 'Текст, который станет ссылкой.' },
          text: { type: 'string', description: 'Подпись ссылки.' },
          url: { type: 'string', description: 'http(s), mailto, #anchor или путь.' },
        },
        required: ['url'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'md_insert_table',
      description: 'Добавляет Markdown-таблицу в конец документа.',
      parameters: {
        type: 'object',
        properties: {
          headers: {
            type: 'array',
            items: { type: 'string' },
            description: 'Заголовки столбцов.',
          },
          rows: {
            type: 'array',
            items: { type: 'array', items: { type: 'string' } },
            description: 'Строки таблицы.',
          },
        },
        required: ['headers', 'rows'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'search_document',
      description: 'Ищет текст в Markdown и возвращает позиции совпадений.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Текст для поиска.' },
          match_case: { type: 'boolean', description: 'Учитывать регистр.' },
        },
        required: ['query'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'read_document',
      description: 'Читает другой файл рабочей папки без права записи.',
      parameters: {
        type: 'object',
        properties: { path: { type: 'string', description: 'Путь к файлу.' } },
        required: ['path'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'write_document',
      description:
        'Записывает другой файл workspace (Word, Excel или Markdown/текст). Доступно только после явного разрешения записи в дополнительные файлы.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string', description: 'Путь к файлу внутри workspace.' },
          html: { type: 'string', description: 'HTML содержимого для .docx.' },
          text: { type: 'string', description: 'Текст для Markdown/текстового файла.' },
          columns: {
            type: 'array',
            items: { type: 'string' },
            description: 'Заголовки столбцов для .xlsx.',
          },
          rows: {
            type: 'array',
            items: { type: 'array', items: { type: 'string' } },
            description: 'Строки для .xlsx.',
          },
          sheet: { type: 'string', description: 'Имя листа для .xlsx.' },
        },
        required: ['path'],
        additionalProperties: false,
      },
    },
  },
];

/** Расширенное управление слайдами: всё, что умеет модель документа. */
const PPTX_ADVANCED_TOOLS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'pptx_get_slide',
      description:
        'Читает один слайд: фон, переход, заметки и фигуры с shape_id, типом, геометрией в см и текстом.',
      parameters: {
        type: 'object',
        properties: { slide: { type: 'number', description: 'Номер слайда с 1.' } },
        required: ['slide'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_duplicate_slide',
      description: 'Дублирует слайд со всем содержимым.',
      parameters: {
        type: 'object',
        properties: { index: { type: 'number', description: 'Номер слайда с 1.' } },
        required: ['index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_hide_slide',
      description: 'Скрывает или показывает слайд при показе.',
      parameters: {
        type: 'object',
        properties: {
          index: { type: 'number', description: 'Номер слайда с 1.' },
          hidden: { type: 'boolean', description: 'true — скрыть, false — показать.' },
        },
        required: ['index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_set_slide_size',
      description: 'Размер слайда: пресет 16:9, 4:3, 16:10 или явные размеры в сантиметрах.',
      parameters: {
        type: 'object',
        properties: {
          preset: { type: 'string', description: '16:9, 4:3 или 16:10.' },
          width_cm: { type: 'number', description: 'Ширина в см (вместо пресета).' },
          height_cm: { type: 'number', description: 'Высота в см (вместо пресета).' },
        },
        required: [],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_set_transition',
      description: 'Переход слайда: none, fade, push, wipe, cut.',
      parameters: {
        type: 'object',
        properties: {
          index: { type: 'number', description: 'Номер слайда с 1.' },
          type: { type: 'string', description: 'none | fade | push | wipe | cut.' },
          speed: { type: 'string', description: 'slow | med | fast.' },
        },
        required: ['index', 'type'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_set_background',
      description: 'Фон слайда: сплошной цвет HEX или изображение из workspace/URL.',
      parameters: {
        type: 'object',
        properties: {
          index: { type: 'number', description: 'Номер слайда с 1.' },
          color: { type: 'string', description: 'Цвет без #, например 0F172A.' },
          image: { type: 'string', description: 'Путь к изображению внутри workspace.' },
          all: { type: 'boolean', description: 'true — применить ко всем слайдам.' },
        },
        required: ['index'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_format_text',
      description:
        'Форматирует текст фигуры целиком: жирный, курсив, подчёркивание, зачёркивание, размер, цвет, шрифт, выравнивание, списки, уровень.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'number', description: 'Номер слайда с 1.' },
          shape_id: { type: 'string', description: 'Идентификатор фигуры из pptx_get_slide.' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          underline: { type: 'boolean' },
          strike: { type: 'boolean' },
          size: { type: 'number', description: 'Размер в пунктах.' },
          color: { type: 'string', description: 'Цвет без #.' },
          font: { type: 'string', description: 'Семейство шрифта, например Calibri.' },
          align: { type: 'string', description: 'l | ctr | r | just.' },
          bullet: { type: 'string', description: 'none | bullet | number.' },
          level: { type: 'number', description: 'Уровень списка 0–4.' },
          line_spacing: { type: 'number', description: 'Межстрочный интервал.' },
        },
        required: ['slide', 'shape_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_find_replace',
      description: 'Ищет и заменяет текст по всем слайдам (при необходимости и в заметках).',
      parameters: {
        type: 'object',
        properties: {
          find: { type: 'string' },
          replace: { type: 'string' },
          include_notes: { type: 'boolean' },
        },
        required: ['find'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_add_shape',
      description:
        'Добавляет фигуру: geom (rect, ellipse, roundRect, triangle, arrow, line…), позиция и размер, заливка, контур, тень, текст.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'number', description: 'Номер слайда с 1.' },
          geom: { type: 'string', description: 'Пресет геометрии PowerPoint.' },
          x: { type: 'number' },
          y: { type: 'number' },
          w: { type: 'number' },
          h: { type: 'number' },
          unit: { type: 'string', description: 'cm или emu (по умолчанию emu).' },
          fill: { type: 'string', description: 'Цвет заливки без #.' },
          outline_color: { type: 'string' },
          outline_width: { type: 'number', description: 'Толщина контура в пунктах.' },
          shadow: { type: 'boolean' },
          text: { type: 'string' },
        },
        required: ['slide'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_set_shape_geometry',
      description: 'Меняет позицию, размер или поворот фигуры.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'number' },
          shape_id: { type: 'string' },
          x: { type: 'number' },
          y: { type: 'number' },
          w: { type: 'number' },
          h: { type: 'number' },
          rotation: { type: 'number', description: 'Поворот в градусах.' },
          unit: { type: 'string', description: 'cm или emu.' },
        },
        required: ['slide', 'shape_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_delete_shape',
      description: 'Удаляет фигуру слайда по shape_id.',
      parameters: {
        type: 'object',
        properties: { slide: { type: 'number' }, shape_id: { type: 'string' } },
        required: ['slide', 'shape_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_duplicate_shape',
      description: 'Дублирует фигуру со смещением.',
      parameters: {
        type: 'object',
        properties: { slide: { type: 'number' }, shape_id: { type: 'string' } },
        required: ['slide', 'shape_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_reorder_shape',
      description: 'Меняет порядок наложения фигуры: front, back, forward, backward.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'number' },
          shape_id: { type: 'string' },
          mode: { type: 'string', description: 'front | back | forward | backward.' },
        },
        required: ['slide', 'shape_id', 'mode'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_set_shape_style',
      description: 'Заливка, контур и тень фигуры. "none" убирает заливку или контур.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'number' },
          shape_id: { type: 'string' },
          fill: { type: 'string', description: 'Цвет без # или none.' },
          outline: { type: 'string', description: 'Цвет без # или none.' },
          outline_width: { type: 'number' },
          shadow: { type: 'boolean' },
        },
        required: ['slide', 'shape_id'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_add_table',
      description: 'Вставляет таблицу с данными; размеры в см при unit=cm.',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'number' },
          rows: { type: 'number' },
          cols: { type: 'number' },
          data: {
            type: 'array',
            description: 'Матрица строк текста (может задать rows/cols автоматически).',
            items: { type: 'array', items: { type: 'string' } },
          },
          x: { type: 'number' },
          y: { type: 'number' },
          w: { type: 'number' },
          h: { type: 'number' },
          unit: { type: 'string', description: 'cm или emu.' },
        },
        required: ['slide'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_set_table_data',
      description: 'Заменяет текст в существующей таблице (shape_id из pptx_get_slide).',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'number' },
          shape_id: { type: 'string' },
          data: {
            type: 'array',
            description: 'Матрица новых значений.',
            items: { type: 'array', items: { type: 'string' } },
          },
        },
        required: ['slide', 'shape_id', 'data'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_apply_theme',
      description:
        'Применяет тему: пресет (office, ink, ocean, forest, sunset, mono…), набор цветов схемы и шрифтовую пару.',
      parameters: {
        type: 'object',
        properties: {
          preset: { type: 'string', description: 'Идентификатор пресета темы.' },
          colors: {
            type: 'object',
            description: 'Переопределение слотов схемы: dk1, lt1, dk2, lt2, accent1…accent6, hlink, folHlink.',
            additionalProperties: { type: 'string' },
          },
          font_pair: { type: 'string', description: 'Идентификатор шрифтовой пары (calibri, segoe, arial…).' },
        },
        required: [],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_present',
      description: 'Запускает полноэкранный показ слайдов в приложении.',
      parameters: {
        type: 'object',
        properties: { slide: { type: 'number', description: 'Слайд, с которого начать.' } },
        required: [],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_export_pdf',
      description: 'Экспортирует презентацию в PDF (страницы размером со слайд).',
      parameters: {
        type: 'object',
        properties: { path: { type: 'string', description: 'Куда сохранить PDF внутри workspace.' } },
        required: [],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'pptx_export_images',
      description: 'Сохраняет каждый слайд в PNG (по умолчанию в «Загрузки»): файлы slide-01.png и далее.',
      parameters: {
        type: 'object',
        properties: {
          dir: { type: 'string', description: 'Папка для PNG внутри workspace. Без неё — «Загрузки».' },
        },
        required: [],
        additionalProperties: false,
      },
    },
  },
];

/** Управление изображениями: работает и для открытой вкладки, и для файлов. */
const IMAGE_TOOLS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'image_info',
      description: 'Размеры, формат, вес файла изображения. Без path — открытое изображение.',
      parameters: {
        type: 'object',
        properties: { path: { type: 'string' } },
        required: [],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'image_transform',
      description:
        'Правит изображение: rotate (90/-90/180), flip (h/v), crop (x,y,w,h, unit px или %), ' +
        'resize (mode fit/exact/max, width/height/max, keep_aspect), convert (png/jpeg/webp + quality), ' +
        'compress (max_kb), thumbnail (max). Операции применяются по порядку.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string', description: 'Файл; без него — открытое изображение.' },
          ops: {
            type: 'array',
            description:
              'Список операций. Пример: [{"op":"rotate","degrees":90},{"op":"resize","mode":"max","max":1600},' +
              '{"op":"compress","max_kb":300}]',
            items: { type: 'object', additionalProperties: true },
          },
          output: { type: 'string', description: 'overwrite | copy | путь к новому файлу (по умолчанию copy).' },
          suffix: { type: 'string', description: 'Суффикс имени при copy, по умолчанию -edit.' },
        },
        required: ['ops'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'image_batch',
      description: 'Применяет те же операции к нескольким файлам или папке; по умолчанию создаёт копии.',
      parameters: {
        type: 'object',
        properties: {
          paths: { type: 'array', items: { type: 'string' }, description: 'Список файлов.' },
          dir: { type: 'string', description: 'Папка (без рекурсии).' },
          ops: { type: 'array', items: { type: 'object', additionalProperties: true } },
          suffix: { type: 'string' },
          overwrite: { type: 'boolean', description: 'true — перезаписывать оригиналы.' },
        },
        required: ['ops'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'image_insert',
      description: 'Вставляет изображение в открытый Word-документ или слайд презентации.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string' },
          target: { type: 'string', description: 'word | slide.' },
          slide: { type: 'number', description: 'Номер слайда, если target=slide.' },
        },
        required: ['path', 'target'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'image_upload',
      description:
        'Загружает изображения на хостинг и возвращает ссылки: imgbb, freeimage, yapx, imgur, catbox, teleph.',
      parameters: {
        type: 'object',
        properties: {
          paths: { type: 'array', items: { type: 'string' } },
          host: { type: 'string', description: 'imgbb | freeimage | yapx | imgur | catbox | teleph.' },
        },
        required: ['paths'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'image_open',
      description: 'Открывает изображение во вкладке редактора.',
      parameters: {
        type: 'object',
        properties: { path: { type: 'string' } },
        required: ['path'],
        additionalProperties: false,
      },
    },
  },
];

/** Управление видео: обработка через ffmpeg главного процесса. */
const VIDEO_TOOLS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'video_info',
      description: 'Длительность, разрешение, частота кадров, дорожки аудио и субтитров.',
      parameters: {
        type: 'object',
        properties: { path: { type: 'string' } },
        required: ['path'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'video_operate',
      description:
        'Обрезает, ускоряет, сжимает, убирает звук или кадрирует видео (зум). ' +
        'Без вкладки обрабатывает файл по path, с открытой вкладкой — применяет к ней.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string' },
          start: { type: 'number', description: 'Начало, секунды.' },
          end: { type: 'number', description: 'Конец, секунды.' },
          speed: { type: 'number', description: 'Множитель скорости 0.25–4.' },
          speed_segments: {
            type: 'array',
            description: 'Сегменты разной скорости: [{start,end,speed}].',
            items: { type: 'object', additionalProperties: true },
          },
          strip_audio: { type: 'boolean' },
          compress: { type: 'boolean' },
          zoom_scale: { type: 'number', description: 'Кадрирование 1–4.' },
          zoom_x: { type: 'number', description: 'Смещение по X -1…1.' },
          zoom_y: { type: 'number', description: 'Смещение по Y -1…1.' },
          output: { type: 'string', description: 'Путь результата.' },
        },
        required: [],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'video_extract_audio',
      description: 'Извлекает звук в mp3/aac (copy — без перекодирования, если дорожка подходит).',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string' },
          format: { type: 'string', description: 'mp3 | aac | copy.' },
          output: { type: 'string' },
        },
        required: ['path'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'video_frames',
      description: 'Делает кадры-миниатюры (по умолчанию 12 равномерно) в папку рядом с видео.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string' },
          count: { type: 'number', description: 'Сколько кадров, 1–24.' },
          at: { type: 'array', items: { type: 'number' }, description: 'Точные секунды кадров.' },
          max_width: { type: 'number' },
          dir: { type: 'string', description: 'Куда сохранить кадры.' },
        },
        required: ['path'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'video_convert',
      description: 'Конвертирует видео в mp4, webm или gif.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string' },
          format: { type: 'string', description: 'mp4 | webm | gif.' },
          quality: { type: 'string', description: 'low | medium | high.' },
          output: { type: 'string' },
        },
        required: ['path', 'format'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'video_gif',
      description: 'Создаёт GIF из отрезка видео.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string' },
          start: { type: 'number' },
          duration: { type: 'number', description: 'Длительность отрезка, секунды.' },
          fps: { type: 'number', description: 'Кадров в секунду, 5–30.' },
          width: { type: 'number', description: 'Ширина GIF, px.' },
          output: { type: 'string' },
        },
        required: ['path'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'video_publish',
      description: 'Публикует видео на YouTube (нужно подключение YouTube в настройках).',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string' },
          title: { type: 'string' },
          description: { type: 'string' },
          privacy: { type: 'string', description: 'private | unlisted | public.' },
          category: { type: 'string' },
          playlist: { type: 'string' },
        },
        required: ['path', 'title'],
        additionalProperties: false,
      },
    },
  },
];

/** Сборка презентации из markdown — доступна всегда. */
const CREATE_PRESENTATION_TOOL: ToolDefinition = {
  type: 'function',
  function: {
    name: 'create_presentation',
    description:
      'Создаёт .pptx из Markdown-разметки: "«# Заголовок»" начинает слайд, "---" — принудительный разрыв, ' +
      'списки и абзацы идут в тело, таблица Markdown становится таблицей PowerPoint, ```код``` — моноширинным блоком. ' +
      'Используй это действие, когда просят создать презентацию, а не Markdown-файл.',
    parameters: {
      type: 'object',
      properties: {
        markdown: { type: 'string', description: 'Содержимое будущих слайдов.' },
        name: { type: 'string', description: 'Имя файла, например Презентация.pptx.' },
        path: { type: 'string', description: 'Папка внутри workspace.' },
        title: { type: 'string', description: 'Заголовок титульного слайда.' },
        theme: { type: 'string', description: 'Пресет темы: office, ink, ocean, forest, sunset, mono…' },
      },
      required: ['markdown'],
      additionalProperties: false,
    },
  },
};


const FILE_TOOLS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'save_as',
      description:
        'Сохраняет открытый документ под новым именем или в другом формате. ' +
        'Используй, когда просят «сохрани как», «сделай docx/xlsx», «пересохрани CSV в Excel». ' +
        'Форматы: xlsx, csv, tsv, docx, md, txt, html, pdf. Не переписывай содержимое вручную.',
      parameters: {
        type: 'object',
        properties: {
          format: {
            type: 'string',
            description: 'Целевой формат: xlsx, csv, tsv, docx, md, txt, html или pdf.',
          },
          path: {
            type: 'string',
            description:
              'Файл или папка назначения внутри workspace. Без path — рядом с исходным файлом.',
          },
          name: { type: 'string', description: 'Новое имя файла без папки.' },
        },
        required: ['format'],
        additionalProperties: false,
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'convert_file',
      description:
        'Конвертирует любой файл workspace: Markdown/текст/HTML в DOCX, CSV/TSV в XLSX, ' +
        'DOCX/XLSX/PPTX в Markdown/текст/HTML/PDF, Markdown-таблицу в XLSX, ' +
        'Markdown/текст/HTML в PPTX (слайды по заголовкам). ' +
        'Создаёт новый файл, исходный не меняет.',
      parameters: {
        type: 'object',
        properties: {
          source: { type: 'string', description: 'Исходный файл внутри workspace.' },
          target: { type: 'string', description: 'Целевой файл или папка.' },
          format: {
            type: 'string',
            description: 'Формат результата: xlsx, csv, tsv, docx, md, txt, html, pdf.',
          },
        },
        required: ['source'],
        additionalProperties: false,
      },
    },
  },
];

function withFileTools(tools: ToolDefinition[]): ToolDefinition[] {
  return [...tools, ...FILE_TOOLS, CREATE_PRESENTATION_TOOL];
}

export function buildDocumentTools(writeAccess: boolean, tabType: FileType | string): ToolDefinition[] {
  if (tabType === 'excel') {
    return writeAccess
      ? withFileTools(EXCEL_TOOLS)
      : withFileTools(
          EXCEL_TOOLS.filter((tool) =>
            ['excel_read', 'excel_get_selection'].includes(tool.function.name)
          )
        );
  }
  if (tabType === 'powerpoint') {
    const all = [...PPTX_TOOLS, ...PPTX_ADVANCED_TOOLS];
    return writeAccess
      ? withFileTools(all)
      : withFileTools(
          all.filter((tool) => ['pptx_read', 'pptx_get_selection'].includes(tool.function.name))
        );
  }
  if (tabType === 'image') {
    return writeAccess
      ? withFileTools(IMAGE_TOOLS)
      : withFileTools(IMAGE_TOOLS.filter((tool) => tool.function.name === 'image_info'));
  }
  if (tabType === 'video') {
    return writeAccess
      ? withFileTools(VIDEO_TOOLS)
      : withFileTools(VIDEO_TOOLS.filter((tool) => tool.function.name === 'video_info'));
  }
  if (tabType === 'markdown' || tabType === 'richtext' || tabType === 'text') {
    return writeAccess
      ? withFileTools(MARKDOWN_TOOLS)
      : withFileTools(
          MARKDOWN_TOOLS.filter((tool) =>
            ['md_read', 'search_document', 'read_document'].includes(tool.function.name)
          )
        );
  }
  const readOnly = [
    'get_document',
    'get_outline',
    'get_block',
    'get_range',
    'search_document',
    'get_selection',
    'get_document_info',
    'table_read',
    'read_document',
  ];
  if (!writeAccess) {
    return withFileTools(ALL_TOOLS.filter((tool) => readOnly.includes(tool.function.name)));
  }
  if (tabType === 'word') return withFileTools(ALL_TOOLS);
  return withFileTools(
    ALL_TOOLS.filter((tool) =>
      ['get_document', 'replace_document', 'find_replace', 'search_document', 'read_document'].includes(
        tool.function.name
      )
    )
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
  excel: 'Excel (.xlsx)',
  powerpoint: 'PowerPoint (.pptx)',
  image: 'изображение',
  video: 'видео',
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
  return `${buildDocumentTypePrompt(context)}\n${DOCUMENT_COMPLETION_RULES}`;
}

/**
 * Правила доведения работы до конца: короткий лимит шагов и ответы вида
 * «начинаю править» оставляли задачу недоделанной.
 */
const DOCUMENT_COMPLETION_RULES = [
  'РАБОТА ДО РЕЗУЛЬТАТА',
  '- Задача выполнена, когда инструмент реально изменил файл. Не пиши «начинаю», «сейчас проверю», «приступаю» — сразу вызывай инструмент.',
  '- Доводи шаг до конца: прочитал данные → внёс правки → проверил результат инструментом чтения, что изменения применились.',
  '- Большую задачу (много строк, листов, абзацев) выполняй последовательными вызовами в одном ответе, пока не закончишь.',
  '- Пометки «[Уже выполнено: …]» в истории диалога — это твои прошлые действия. Продолжай с этого места и не повторяй их.',
].join('\n');

function buildDocumentTypePrompt(context: DocumentPromptContext): string {
  const typeLabel = TYPE_LABEL[context.type] ?? context.type;
  const header = [
    'ОТКРЫТЫЙ ДОКУМЕНТ',
    `- Файл: «${context.name}» (${typeLabel}).`,
    `- Права записи: ${context.writeAccess ? 'ЕСТЬ' : 'НЕТ'}.`,
  ];

  if (context.type === 'excel') {
    if (!context.writeAccess) {
      return [
        ...header,
        '- Доступно только чтение: используй excel_read. Изменять книгу нельзя.',
        '- Если просят изменить данные, скажи, что нужны права записи.',
      ].join('\n');
    }
    return [
      ...header,
      '- Работай с книгой через excel_read и excel_write_cells. Не создавай новый файл и не предлагай CSV, если просят изменить открытую таблицу.',
      '- excel_read: лист и диапазон A1 (например A1:F30). Возвращает значения, формулы и отображение.',
      '- excel_get_selection: текущее выделение пользователя на листе.',
      '- excel_write_cells: список {address, value}; формула начинается с "=". Для диапазона передай range и матрицу values.',
      '- Листы: excel_add_sheet, excel_rename_sheet, excel_delete_sheet, excel_set_active_sheet.',
      '- Строки/столбцы: excel_insert_rows, excel_delete_rows, excel_insert_columns, excel_delete_columns; объединение и формат: excel_merge_cells, excel_unmerge_cells, excel_format_range.',
      '- Не выводи таблицу в чат: изменения применяются к открытой книге. В чате — короткий отчёт.',
    ].join('\n');
  }

  if (context.type === 'powerpoint') {
    if (!context.writeAccess) {
      return [
        ...header,
        '- Доступно только чтение: используй pptx_read. Менять презентацию нельзя.',
        '- Если просят изменить слайды, скажи, что нужны права записи.',
      ].join('\n');
    }
    return [
      ...header,
      '- Работай с презентацией через pptx_read, pptx_set_text, pptx_add_slide, pptx_delete_slide, pptx_move_slide, pptx_add_text_box, pptx_insert_image, pptx_set_notes, pptx_set_layout.',
      '- pptx_get_selection: выделенная текстовая фигура текущего слайда.',
      '- pptx_set_text: slide (с 1), shape_id из pptx_read, text; переносы строк — новые абзацы.',
      '- pptx_insert_image: slide и src (путь workspace, data URL или http-ссылка); width/height в px.',
      '- pptx_set_notes: заметки докладчика без изменения текста слайда.',
      '- pptx_set_layout: имя или номер макета из pptx_read.',
      `- Полный набор: ${PPTX_ADVANCED_TOOLS.map((tool) => tool.function.name).join(', ')}.`,
      '- Слайды: pptx_get_slide (shape_id и геометрия в см), pptx_duplicate_slide, pptx_hide_slide, pptx_set_slide_size, pptx_set_transition, pptx_set_background.',
      '- Текст: pptx_format_text (жирный/курсив/размер/цвет/шрифт/выравнивание/списки) по всей фигуре, pptx_find_replace по всем слайдам.',
      '- Фигуры: pptx_add_shape (geom, заливка, контур, тень, текст), pptx_set_shape_geometry (unit "cm" или "emu"), pptx_set_shape_style, pptx_reorder_shape, pptx_duplicate_shape, pptx_delete_shape.',
      '- Таблицы и тема: pptx_add_table (data — матрица строк), pptx_set_table_data, pptx_apply_theme (preset/colors/font_pair).',
      '- Показ и экспорт: pptx_present, pptx_export_pdf. Новая презентация из текста — create_presentation, а не Markdown-файл.',
      '- Не создавай новый файл и не переписывай презентацию целиком из-за одного слайда.',
      '- В чате — короткий отчёт; текст слайдов в чат не дублируй.',
    ].join('\n');
  }

  if (context.type === 'image') {
    return [
      ...header,
      '- Это изображение. Правки — инструментами image_transform и image_batch, а не переписыванием файла.',
      `- Доступно: ${IMAGE_TOOLS.map((tool) => tool.function.name).join(', ')}.`,
      '- image_transform: ops по порядку — rotate/flip/crop/resize/convert/compress/thumbnail; по умолчанию результат в копию с суффиксом -edit.',
      '- image_info: размеры, формат и вес. image_upload: ссылки с imgbb/freeimage/yapx/imgur/catbox. image_insert: в открытый Word или слайд.',
      '- В чат — только отчёт: что сделано, новые размеры и путь файла.',
    ].join('\n');
  }

  if (context.type === 'video') {
    return [
      ...header,
      '- Это видео. Обработка идёт через ffmpeg; без ffmpeg инструменты вернут ошибку, превью работает.',
      `- Доступно: ${VIDEO_TOOLS.map((tool) => tool.function.name).join(', ')}.`,
      '- video_operate: start/end, speed, speed_segments, strip_audio, compress, zoom_scale/zoom_x/zoom_y, output — обрезка, скорость, сжатие и кадрирование.',
      '- video_extract_audio, video_frames, video_convert, video_gif — производные файлы; оригинал не меняется.',
      '- video_publish — публикация на YouTube (нужно подключение в настройках).',
      '- В чат — отчёт: длительность, что сделано и путь результата.',
    ].join('\n');
  }

  if (
    !context.fallback &&
    (context.type === 'markdown' || context.type === 'richtext' || context.type === 'text')
  ) {
    if (!context.writeAccess) {
      return [
        ...header,
        '- Доступно только чтение: используй md_read. Изменять документ нельзя.',
        '- Если просят изменить текст, скажи, что нужны права записи.',
      ].join('\n');
    }
    return [
      ...header,
      '- Работай с Markdown напрямую через md_read, md_replace, md_append, md_replace_document, md_set_heading, md_toggle_list, md_insert_link, md_insert_table и search_document.',
      '- Сохраняй Markdown-разметку: заголовки #, списки, ссылки, таблицы. Не переводи документ в HTML.',
      '- Для точечной правки используй md_replace, а не md_replace_document.',
      '- Другой файл меняй только write_document и только при включённом «Доп. файлы: запись».',
      '- В чате — короткий отчёт; текст документа не дублируй.',
    ].join('\n');
  }

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

  const rules = [...header, '- Работай только с этим документом; другие файлы workspace можно читать через read_document, но не менять.'];

  if (context.fallback) {
    rules.push(
      '- Function calling недоступен. Содержимое документа приложено к сообщению выше — не вызывай get_document, работай с приложенным текстом.',
      '- Любые правки возвращай только блоком workspace-action с действием document_apply:',
      '```workspace-action',
      '{"action":"document_apply","operations":[{"tool":"replace_document","html":"<h1>Заголовок</h1><p>Текст</p>"},{"tool":"page_setup","template":"russian_business"}]}',
      '```',
      '- Для переписывания и оформления используй replace_document с полным HTML документа, для точечных замен — find_replace. Блочные операции replace_block/insert_after/delete_block применяй только когда точно знаешь id блока.',
      '- Поля страницы и базовый шрифт меняй операцией page_setup, а не советами пользователю.',
      '- После блока — одна короткая строка отчёта. Не вставляй текст документа в обычное сообщение.'
    );
  } else {
    rules.push(
      '- Работай инструментами: get_document, replace_document, replace_block, insert_after, delete_block, find_replace, format, page_setup.',
      '- Сначала вызови get_document, если структура ещё не известна. Идентификаторы блоков (b1, b2, …) бери из последнего ответа get_document или из результата предыдущего инструмента.',
      '- При просьбе изменить, оформить, переписать, сократить, удалить лишнее или поправить текст — вызывай инструменты. Не пиши готовый текст документа в чат.'
    );
  }

  rules.push(
    '- Не создавай и не предлагай Excel/CSV (create_file, create_table, create_document), если открыт Word/Markdown и пользователь не попросил именно таблицу. Правка идёт только в открытый документ.',
    '- Не отвечай «не могу редактировать документ Word»: правка применяется сразу.',
    '- В документ клади только содержимое документа: без вступлений («Вот вариант…»), пояснений, вопросов, советов и списков «что изменено».',
    '- В чат верни короткий отчёт: 1–3 пункта о том, что сделано. Текст документа в чат не дублируй.',
    '- Никогда не советуй настраивать Word вручную (поля, шрифт, интервал, «скопируйте в Word»): для этого есть page_setup, он меняет сам файл.',
    '- Другой файл рабочей папки читай через read_document по пути. Дополнительные файлы доступны только для чтения; запись в них — только инструментом write_document после разрешения «Доп. файлы: запись».',
    '- Если просят «оформи как реальный/официальный документ», «как в Word», заявление, письмо, договор или приказ — вызови page_setup с template "russian_business" (поля 3/1,5/2/2 см, Times New Roman 14, одинарный интервал, по ширине, красная строка 1,25 см).',
    '- ATS/ATC-фильтры: полностью замени документ плоским одноколоночным текстом. Заголовки «Опыт работы», «Навыки», «Образование», обычные абзацы и маркированные списки; ключевые слова из вакансии; без таблиц, колонок, картинок и текстовых блоков. Не возвращай копию исходного текста — формулировки и структура должны измениться.',
    '- Если текст документа уже приложен к сообщению (полная переработка/ATS), не вызывай get_document — сразу replace_document с изменённым текстом.',
    context.fallback
      ? '- Для одной строки, фразы, даты, имени или пункта используй find_replace (find — точный фрагмент). replace_document — только если меняется весь документ.'
      : '- Для одной строки, фразы, даты, имени или пункта используй find_replace или replace_block. replace_document — только при полной переработке документа; одна строка не повод переписывать весь документ.\n' +
        '- «Выделенное», «этот фрагмент», «здесь»: сначала get_selection, затем replace_selection. Не переписывай весь документ из-за выделения.\n' +
        '- Правка части абзаца: get_range (смещения) → replace_range/delete_range. Полный HTML блока: get_block. Длинный документ: get_outline с offset/limit.\n' +
        '- Колонтитулы, сноски, комментарии, именованные стили и число секций — get_document_info.\n' +
        '- Верхний/нижний колонтитул меняется set_header_footer {kind, type, text}, удаляется delete_header_footer. Не переписывай тело документа из-за колонтитула.',
    '- В html для replace_block/insert_after передавай только сам фрагмент: без «Вот текст…», отчётов, списков изменений и обёртки ```.',
    '- Формат HTML: <h1>–<h3>, <p>, <ul>/<ol> с <li>, <table>, <blockquote>, <strong>, <em>, <u>, <s>, <br>. Оформление абзаца — inline style: text-align, font-family, font-size (pt), margin-left (pt), text-indent (pt). Markdown внутри HTML не используй.',
    '- Не заворачивай HTML документа в блок кода и не добавляй к нему заголовки-комментарии.'
  );

  if (context.type !== 'word') {
    rules.push(
      '- Для этого типа файла доступны только get_document, replace_document и find_replace: содержимое хранится как текст, блочное форматирование не поддерживается.'
    );
  }

  return rules.join('\n');
}
