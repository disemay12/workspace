/** Готовые шаблоны документов для быстрого создания файла. */
export interface DocumentTemplate {
  id: string;
  name: string;
  description: string;
  extension: 'docx' | 'md';
  html?: string;
  markdown?: string;
}

export const DOCUMENT_TEMPLATES: DocumentTemplate[] = [
  {
    id: 'letter',
    name: 'Деловое письмо',
    description: 'Шапка, обращение, основной текст и подпись.',
    extension: 'docx',
    html: [
      '<p style="text-align:right">Генеральному директору<br/>ООО «Компания»<br/>Иванову И.И.</p>',
      '<p style="text-align:center"><strong>ИСХОДЯЩЕЕ ПИСЬМО</strong></p>',
      '<p>Уважаемый Иван Иванович!</p>',
      '<p>Прошу рассмотреть предложение о сотрудничестве. Подробности приложены отдельным документом.</p>',
      '<p>С уважением,<br/>Петров П.П.<br/>Телефон: +7 900 000-00-00</p>',
    ].join(''),
  },
  {
    id: 'resume',
    name: 'Резюме',
    description: 'Структура резюме под ATS: контакты, опыт, навыки, образование.',
    extension: 'docx',
    html: [
      '<h1>Иван Иванов</h1>',
      '<p>Москва · +7 900 000-00-00 · ivan@example.com</p>',
      '<h2>Опыт работы</h2>',
      '<p><strong>Компания</strong> — инженер, 2022 — настоящее время</p>',
      '<h2>Навыки</h2>',
      '<ul><li>Навык 1</li><li>Навык 2</li></ul>',
      '<h2>Образование</h2>',
      '<p>Университет, 2021</p>',
    ].join(''),
  },
  {
    id: 'report',
    name: 'Отчёт',
    description: 'Титульная часть, разделы и итоги.',
    extension: 'docx',
    html: [
      '<h1 style="text-align:center">ОТЧЁТ</h1>',
      '<p style="text-align:center">за период 2026 года</p>',
      '<h2>1. Введение</h2>',
      '<p>Цель отчёта — подвести итоги и зафиксировать выводы.</p>',
      '<h2>2. Основные результаты</h2>',
      '<ul><li>Результат 1</li><li>Результат 2</li></ul>',
      '<h2>3. Выводы</h2>',
      '<p>Выводы и рекомендации.</p>',
    ].join(''),
  },
  {
    id: 'note',
    name: 'Markdown-заметка',
    description: 'Заголовок и разделы для быстрых заметок.',
    extension: 'md',
    markdown: '# Заметка\n\n## Что сделано\n\n- \n\n## Следующие шаги\n\n- \n',
  },
  {
    id: 'blank',
    name: 'Пустой документ Word',
    description: 'Чистый .docx с базовыми стилями.',
    extension: 'docx',
    html: '<p></p>',
  },
];
