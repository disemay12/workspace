/**
 * Продвинутый Markdown-движок для WorkSpace.
 *
 * Поддерживает:
 * - Полную спецификацию GitHub Flavored Markdown (таблицы, списки задач, заголовки, переносы) через marked
 * - Подсветку синтаксиса кода с помощью Prism.js (15+ языков)
 * - Математические формулы LaTeX: инлайн $...$ и блочные $$...$$ через KaTeX
 * - Интерактивные диаграммы и схемы через Mermaid
 */
import { Marked } from 'marked';
import katex from 'katex';
import Prism from 'prismjs';
import mermaid from 'mermaid';
import 'katex/dist/katex.min.css';
import 'prismjs/themes/prism-tomorrow.min.css';

// Языки Prism
import 'prismjs/components/prism-javascript';
import 'prismjs/components/prism-typescript';
import 'prismjs/components/prism-jsx';
import 'prismjs/components/prism-tsx';
import 'prismjs/components/prism-json';
import 'prismjs/components/prism-python';
import 'prismjs/components/prism-sql';
import 'prismjs/components/prism-bash';
import 'prismjs/components/prism-markdown';
import 'prismjs/components/prism-yaml';
import 'prismjs/components/prism-css';
import 'prismjs/components/prism-diff';

export {
  escapeHtml,
  escapeRawHtml,
  isSafeImageSrc,
  isSafeLinkHref,
  sanitizeFenceLanguage,
} from './sanitizeMarkdownHtml';
import {
  escapeHtml,
  escapeRawHtml,
  isSafeImageSrc,
  isSafeLinkHref,
  sanitizeFenceLanguage,
} from './sanitizeMarkdownHtml';

// Инициализируем mermaid один раз
let mermaidInitialized = false;
export function initMermaid(isDark = true): void {
  try {
    mermaid.initialize({
      startOnLoad: false,
      theme: isDark ? 'dark' : 'default',
      // 'loose' разрешал HTML в подписях и click-обработчики: диаграмма из
      // документа исполняла скрипт в окне с мостом electronAPI.
      securityLevel: 'strict',
      fontFamily: 'Segoe UI, system-ui, sans-serif',
    });
    mermaidInitialized = true;
  } catch (err) {
    console.warn('[Mermaid] Ошибка инициализации:', err);
  }
}

/** Обработка формул KaTeX до разбора markdown */
function processMath(text: string): { text: string; mathMap: Map<string, string> } {
  const mathMap = new Map<string, string>();
  let mathCounter = 0;

  // 1. Блочные формулы: $$ ... $$
  let processed = text.replace(/\$\$([\s\S]+?)\$\$/g, (_, formula) => {
    const key = `%%MATH_BLOCK_${mathCounter++}%%`;
    try {
      const rendered = katex.renderToString(formula.trim(), {
        displayMode: true,
        throwOnError: false,
      });
      mathMap.set(key, `<div class="katex-display-wrapper my-2 select-text">${rendered}</div>`);
    } catch {
      mathMap.set(key, `<pre class="text-ws-danger font-mono text-xs">$$${escapeHtml(formula)}$$</pre>`);
    }
    return key;
  });

  // 2. Инлайн формулы: $ ... $ (без пробелов у знаков доллара и без экранирования \$)
  processed = processed.replace(/(?<!\\)\$([^\$\n]+?)(?<!\\)\$/g, (_, formula) => {
    const key = `%%MATH_INLINE_${mathCounter++}%%`;
    try {
      const rendered = katex.renderToString(formula.trim(), {
        displayMode: false,
        throwOnError: false,
      });
      mathMap.set(key, `<span class="katex-inline-wrapper select-text">${rendered}</span>`);
    } catch {
      mathMap.set(key, `<code class="text-ws-danger">$${escapeHtml(formula)}$</code>`);
    }
    return key;
  });

  return { text: processed, mathMap };
}

// Настройка кастомного экземпляра Marked
const markedInstance = new Marked({
  gfm: true,
  breaks: true,
});

markedInstance.use({
  renderer: {
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

      // Mermaid диаграмма
      if (language === 'mermaid') {
        const id = 'mermaid_' + Math.random().toString(36).substring(2, 9);
        return `<div class="mermaid-diagram my-3 overflow-x-auto rounded border border-ws bg-ws-elevated p-3 text-center" data-id="${id}" data-mermaid="${encodeURIComponent(text)}"><pre class="mermaid text-xs font-mono">${escapeHtml(text)}</pre></div>`;
      }

      // Подсветка Prism
      let highlighted = '';
      if (language && Prism.languages[language]) {
        try {
          highlighted = Prism.highlight(text, Prism.languages[language], language);
        } catch {
          highlighted = escapeHtml(text);
        }
      } else {
        highlighted = escapeHtml(text);
      }

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
        return `<${tag} class="border border-ws px-2 py-1 ${align}">${markedInstance.parseInline(
          cell.text ?? ''
        )}</${tag}>`;
      };
      const head = header.map((cell) => renderCell(cell, 'th')).join('');
      const body = rows
        .map((row) => `<tr>${row.map((cell) => renderCell(cell, 'td')).join('')}</tr>`)
        .join('');
      return `<div class="my-3 overflow-x-auto"><table class="w-full border-collapse border border-ws text-xs"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
    },
    blockquote({ text }) {
      return `<blockquote class="border-l-4 border-ws-accent bg-ws-accent-soft/30 px-3 py-1.5 my-2 text-xs italic text-ws-muted rounded-r">${text}</blockquote>`;
    },
  },
});

/**
 * Основная функция рендеринга Markdown в HTML
 */
export function renderRichMarkdown(markdown: string): string {
  if (!markdown) return '';

  const { text: processed, mathMap } = processMath(markdown);

  let html = '';
  try {
    html = markedInstance.parse(processed) as string;
  } catch (err) {
    html = `<p>${escapeHtml(markdown)}</p>`;
  }

  // Восстанавливаем отрендеренные формулы KaTeX
  for (const [key, value] of mathMap.entries()) {
    html = html.replace(key, value);
  }

  return html;
}

/**
 * Активное содержимое в готовом SVG. Проверка идёт по разобранному дереву, а
 * не регуляркой: шаблон `\son[a-z]+\s*=` срабатывал на обычный текст подписи
 * (`onset=1`, `online=true`) и превращал валидную диаграмму в ошибку.
 */
function svgHasActiveContent(svg: string): boolean {
  let doc: Document | null = null;
  try {
    doc = new DOMParser().parseFromString(svg, 'image/svg+xml');
  } catch {
    doc = null;
  }
  if (!doc || doc.getElementsByTagName('parsererror').length > 0) {
    // Разобрать не удалось — пропускаем только заведомо безопасный текст.
    return true;
  }
  for (const element of Array.from(doc.getElementsByTagName('*'))) {
    const tag = element.tagName.toLowerCase().replace(/^svg:/, '');
    if (tag === 'script' || tag === 'foreignobject' || tag === 'iframe' || tag === 'embed') return true;
    for (const attribute of Array.from(element.attributes)) {
      const name = attribute.name.toLowerCase();
      if (name.startsWith('on')) return true;
      if ((name === 'href' || name === 'xlink:href' || name === 'src') && !isSafeLinkHref(attribute.value)) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Рендерит все Mermaid блоки внутри переданного DOM-контейнера
 */
export async function renderMermaidBlocks(container: HTMLElement | null): Promise<void> {
  if (!container) return;
  if (!mermaidInitialized) initMermaid(true);

  const blocks = container.querySelectorAll('.mermaid-diagram');
  if (blocks.length === 0) return;

  for (let i = 0; i < blocks.length; i++) {
    const el = blocks[i] as HTMLElement;
    const rawCode = el.getAttribute('data-mermaid');
    if (!rawCode) continue;

    const code = decodeURIComponent(rawCode);
    const id = 'mermaid_svg_' + Math.random().toString(36).substring(2, 9);

    try {
      const { svg } = await mermaid.render(id, code);
      // Даже при securityLevel:'strict' проверяем результат: SVG вставляется
      // в живой DOM, а не в изолированный контейнер.
      if (svgHasActiveContent(svg)) {
        throw new Error('в SVG найдены активные элементы');
      }
      el.innerHTML = svg;
    } catch (err) {
      el.innerHTML = `<div class="text-ws-danger text-xs text-left p-2 font-mono">Ошибка диаграммы: ${escapeHtml(
        (err as Error).message
      )}</div><pre class="text-left font-mono text-[11px] text-ws-faint p-2">${escapeHtml(code)}</pre>`;
    }
  }
}
