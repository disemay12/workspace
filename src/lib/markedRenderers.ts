/**
 * Рендереры marked, собирающие разметку Markdown в приложении.
 *
 * Модуль без зависимостей от Prism/KaTeX/Mermaid: подсветку кода и inline-разбор
 * ячеек таблицы приложение передаёт адаптерами. Причина выноса — проверочный
 * набор (electron/checks/sanitizer-check.ts) импортирует этот же файл: пока он
 * держал собственную копию рендереров, поломка боевого html/link/image-рендерера
 * оставляла проверки зелёными.
 */
import {
  escapeHtml,
  escapeRawHtml,
  isSafeImageSrc,
  isSafeLinkHref,
  sanitizeFenceLanguage,
} from './sanitizeMarkdownHtml';

/** Точки расширения, которые в приложении закрывают Prism и сам экземпляр marked. */
export interface MarkedRendererAdapters {
  /** Подсветка кода (Prism в приложении); null — подсветки для языка нет. */
  highlight?: (text: string, language: string) => string | null;
  /** Inline-разбор ячейки таблицы тем же экземпляром marked. */
  parseInline?: (text: string) => string;
}

/**
 * Срез экземпляра marked, нужный рендерерам. Собственные типы вместо импорта из
 * 'marked': пакет — ES-модуль, а этот файл компилируется и как CommonJS
 * (electron/checks/tsconfig: Node16), где type-only импорт ESM-типов запрещён.
 */
export interface MarkedRendererContext {
  parser: {
    parseInline(tokens: unknown): string;
  };
}

/** Ячейка таблицы в marked v18: текст и выравнивание колонки. */
interface MarkedTableCell {
  text: string;
  align?: string | null;
}

/** Правила разметки, совпадающие по форме с renderer'ом marked. */
export interface MarkedRendererSet {
  html(this: MarkedRendererContext, token: { text: string }): string;
  link(
    this: MarkedRendererContext,
    token: { href: string; title?: string | null; tokens: unknown }
  ): string;
  image(token: { href: string; title?: string | null; text?: string }): string;
  code(token: { text: string; lang?: string }): string;
  table(token: { header: MarkedTableCell[]; rows: MarkedTableCell[][] }): string;
}

/**
 * Правила разметки, отвечающие за пользовательский ввод: сырой HTML, ссылки,
 * картинки, информационная строка код-фенса и ячейки таблицы.
 */
export function createMarkedRenderers(
  adapters: MarkedRendererAdapters = {}
): MarkedRendererSet {
  const { highlight, parseInline } = adapters;
  return {
    /** Единственная точка, где в вывод попадает пользовательский HTML. */
    html({ text }) {
      return escapeRawHtml(text);
    },
    /**
     * Ссылка тоже приходит из документа, а переход по `javascript:` исполняет
     * код в окне с мостом electronAPI. Разрешены только явные схемы.
     */
    link({ href, title, tokens }) {
      const text = this.parser.parseInline(tokens);
      if (!isSafeLinkHref(href)) return text;
      const titleAttr = title ? ` title="${escapeHtml(title)}"` : '';
      return `<a href="${escapeHtml(href)}"${titleAttr} class="ws-link text-ws-accent underline decoration-dotted">${text}</a>`;
    },
    image({ href, title, text }) {
      if (!isSafeImageSrc(href)) return escapeHtml(text || '');
      const titleAttr = title ? ` title="${escapeHtml(title)}"` : '';
      return `<img src="${escapeHtml(href)}" alt="${escapeHtml(text || '')}"${titleAttr} class="max-w-full rounded my-1" />`;
    },
    code({ text, lang }) {
      // info-string фенса приходит из документа: вставлять его в разметку
      // без экранирования — тот же XSS, что и сырой HTML.
      const language = sanitizeFenceLanguage(lang || '');

      // Mermaid диаграмма: разметку-обёртку рендерит главный процесс проверки
      // так же, как и приложение, — сам mermaid.js здесь не нужен.
      if (language === 'mermaid') {
        const id = 'mermaid_' + Math.random().toString(36).substring(2, 9);
        return `<div class="mermaid-diagram my-3 overflow-x-auto rounded border border-ws bg-ws-elevated p-3 text-center" data-id="${id}" data-mermaid="${encodeURIComponent(text)}"><pre class="mermaid text-xs font-mono">${escapeHtml(text)}</pre></div>`;
      }

      // Без адаптера подсветки код остаётся экранированным текстом — тот же
      // безопасный путь, что и при отказе Prism.
      const highlighted = (highlight && highlight(text, language)) || escapeHtml(text);

      return `<div class="relative group my-2 overflow-hidden rounded border border-ws bg-ws-elevated text-xs font-mono">
        <div class="flex items-center justify-between border-b border-ws bg-ws-panel px-3 py-1 text-[10px] text-ws-muted uppercase tracking-wider">
          <span>${language || 'text'}</span>
          <button class="copy-code-btn rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:text-ws hover:bg-ws-hover transition-colors" data-code="${encodeURIComponent(text)}">Копировать</button>
        </div>
        <pre class="p-3 overflow-x-auto language-${language || 'text'}"><code>${highlighted}</code></pre>
      </div>`;
    },
    table({ header, rows }) {
      // marked v18 передаёт ячейки объектами { text, align, tokens }.
      // Прямая вставка массива в шаблон давала «[object Object]» в каждой ячейке.
      const renderCell = (cell: { text: string; align?: string | null }, tag: 'th' | 'td') => {
        const align =
          cell.align === 'center'
            ? 'text-center'
            : cell.align === 'right'
              ? 'text-right'
              : 'text-left';
        const inner = parseInline ? parseInline(cell.text ?? '') : escapeHtml(cell.text ?? '');
        return `<${tag} class="border border-ws px-2 py-1 ${align}">${inner}</${tag}>`;
      };
      const head = header.map((cell) => renderCell(cell, 'th')).join('');
      const body = rows
        .map((row) => `<tr>${row.map((cell) => renderCell(cell, 'td')).join('')}</tr>`)
        .join('');
      return `<div class="my-3 overflow-x-auto"><table class="w-full border-collapse border border-ws text-xs"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
    },
  };
}
