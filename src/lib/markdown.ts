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
import { escapeHtml, isSafeLinkHref } from './sanitizeMarkdownHtml';
import { createMarkedRenderers } from './markedRenderers';

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

// Единственное определение боевых рендереров живёт в markedRenderers.ts:
// этот же файл импортирует проверочный набор, иначе он проверял копию.
markedInstance.use({
  renderer: {
    ...createMarkedRenderers({
      // Prism подключается адаптером: модуль рендереров общий с проверочным
      // набором, который запускается вне браузера и подсветку не тянет.
      highlight(text, language) {
        if (!language || !Prism.languages[language]) return escapeHtml(text);
        try {
          return Prism.highlight(text, Prism.languages[language], language);
        } catch {
          return escapeHtml(text);
        }
      },
      parseInline(text) {
        return String(markedInstance.parseInline(text));
      },
    }),
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
