/**
 * Проверка экранирования HTML в Markdown.
 *
 * Сырой HTML из документа или ответа модели исполняется в окне с мостом
 * electronAPI, поэтому набор страхует именно этот путь: разрешённая разметка
 * без атрибутов проходит, всё остальное становится текстом.
 */
import * as fs from 'fs';
import * as path from 'path';
import {
  escapeHtml,
  escapeRawHtml,
  isSafeImageSrc,
  isSafeLinkHref,
  sanitizeFenceLanguage,
} from '../../src/lib/sanitizeMarkdownHtml';

const { Marked } = require('marked');

/**
 * Корень проекта — поиск package.json вверх: набор запускают и из
 * dist-electron/, и из dist-electron/electron/checks/.
 */
function resolveProjectRoot(): string {
  let dir = __dirname;
  for (let i = 0; i < 4; i++) {
    if (fs.existsSync(path.join(dir, 'package.json'))) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return path.join(__dirname, '..', '..', '..');
}
const ROOT = resolveProjectRoot();

const failures: string[] = [];
let passed = 0;

function check(name: string, actual: unknown, expected: unknown) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) passed++;
  else failures.push(`${name}\n     ожидалось: ${e}\n     получено:   ${a}`);
}

/**
 * Реальный marked с теми же настройками, что и renderRichMarkdown: проверка
 * изолированного escapeRawHtml не поймала бы регрессию в самом конвейере
 * (html-токен до рендерера может не доходить вовсе).
 */
const marked = new Marked({ gfm: true, breaks: true });
marked.use({
  renderer: {
    html({ text }: { text: string }) {
      return escapeRawHtml(text);
    },
    link({ href, tokens }: { href: string; tokens: unknown[] }) {
      const text = this.parser.parseInline(tokens);
      return isSafeLinkHref(href) ? `<a href="${escapeHtml(href)}">${text}</a>` : text;
    },
  },
});
const renderMarkdown = (source: string): string => String(marked.parse(source));

/* ---------- Активные конструкции уходят в текст ---------- */
check('script', escapeRawHtml('<script>alert(1)</script>'), '&lt;script&gt;alert(1)&lt;/script&gt;');
check('img с onerror', escapeRawHtml('<img src=x onerror=alert(1)>'), '&lt;img src=x onerror=alert(1)&gt;');
check(
  'javascript: в ссылке',
  escapeRawHtml('<a href="javascript:alert(1)">x</a>'),
  '&lt;a href=&quot;javascript:alert(1)&quot;&gt;x&lt;/a&gt;'
);
check(
  'iframe',
  escapeRawHtml('<iframe src="https://evil"></iframe>'),
  '&lt;iframe src=&quot;https://evil&quot;&gt;&lt;/iframe&gt;'
);
check('svg с onload', escapeRawHtml('<svg onload=alert(1)></svg>'), '&lt;svg onload=alert(1)&gt;&lt;/svg&gt;');
check('style', escapeRawHtml('<style>body{display:none}</style>'), '&lt;style&gt;body{display:none}&lt;/style&gt;');
check(
  'form и input',
  escapeRawHtml('<form action="/"><input name=q></form>'),
  '&lt;form action=&quot;/&quot;&gt;&lt;input name=q&gt;&lt;/form&gt;'
);
// Разрешённый тег остаётся, но без атрибутов: обработчик снимается целиком.
check('обработчик снимается', escapeRawHtml('<span onclick="alert(1)">t</span>'), '<span>t</span>');
check(
  'script с атрибутом на новой строке',
  escapeRawHtml('<script\nsrc=x>alert(1)</script>'),
  '&lt;script\nsrc=x&gt;alert(1)&lt;/script&gt;'
);
// `<b/onmouseover=...>` старый шаблон «тег + пробелы + атрибуты» пропускал как
// живой тег: слэш вместо пробела обходил проверку атрибутов.
const bypassTag = escapeRawHtml('<b/onmouseover=alert(1)>');
check('разрешённый тег со слэшем: атрибут срезан', bypassTag, '<b>');
check('разрешённый тег со слэшем: обработчика нет', bypassTag.includes('onmouseover'), false);
check('svg со слэшем и onload', escapeRawHtml('<svg/onload=alert(1)>'), '&lt;svg/onload=alert(1)&gt;');

/* ---------- Безобидная разметка проходит ---------- */
check('жирный', escapeRawHtml('<b>текст</b>'), '<b>текст</b>');
check('подчёркнутый закрывается', escapeRawHtml('<u>a</u>'), '<u>a</u>');
check('абзац и div', escapeRawHtml('<div><p>x</p></div>'), '<div><p>x</p></div>');
check('перенос строки', escapeRawHtml('a<br>b'), 'a<br>b');
check('атрибуты снимаются', escapeRawHtml('<span class="x">t</span>'), '<span>t</span>');
check('комментарий удаляется', escapeRawHtml('a<!-- секрет -->b'), 'ab');
// Одиночные < и > без имени тега браузер показывает текстом — так и оставляем.
check('математика в тексте', escapeRawHtml('2 < 3 и 4 > 1'), '2 < 3 и 4 > 1');

/* ---------- Конвейер marked ---------- */
const renderedImg = renderMarkdown('<img src=x onerror=alert(1)>');
check('marked: живой <img> не появляется', /<img[\s/>]/i.test(renderedImg), false);
check('marked: onerror не остаётся атрибутом', /<img[^>]*onerror/i.test(renderedImg), false);
const renderedScript = renderMarkdown('<script>alert(1)</script>');
check('marked: живой <script> не появляется', /<script[\s/>]/i.test(renderedScript), false);
const renderedSlashTag = renderMarkdown('a <b/onmouseover=alert(1)> b');
check('marked: <b/onmouseover> не даёт живого обработчика', /<b[^>]*onmouseover/i.test(renderedSlashTag), false);
const renderedSvg = renderMarkdown('<svg/onload=alert(1)>');
check('marked: живой <svg> не появляется', /<svg[\s/>]/i.test(renderedSvg), false);
const renderedLink = renderMarkdown('[x](javascript:alert(1))');
check('marked: javascript: ссылка не становится <a>', /<a[\s>]/i.test(renderedLink), false);

/* ---------- Схемы ссылок и картинок ---------- */
// Renderer ссылок в markdown.ts пропускает только эти схемы; проверяем сам
// allowlist, чтобы тест не зависел от katex/mermaid в браузерном модуле.
check('javascript: ссылка отклонена', isSafeLinkHref('javascript:alert(1)'), false);
check('data:text/html отклонён', isSafeLinkHref('data:text/html,<script>alert(1)</script>'), false);
check('file: отклонён', isSafeLinkHref('file:///C:/Windows/win.ini'), false);
check('https: разрешён', isSafeLinkHref('https://example.com/page'), true);
check('mailto: разрешён', isSafeLinkHref('mailto:test@example.com'), true);
check('относительный путь разрешён', isSafeLinkHref('./docs/page.html'), true);
check('javascript: картинка отклонена', isSafeImageSrc('javascript:alert(1)'), false);
check('data:image/png разрешена', isSafeImageSrc('data:image/png;base64,AAAA'), true);
check('https: картинка разрешена', isSafeImageSrc('https://example.com/a.png'), true);

/* ---------- Экранирование атрибутов и кавычек ---------- */
check('кавычки', escapeHtml('"a" \'b\''), '&quot;a&quot; &#39;b&#39;');
check('амперсанд', escapeHtml('a & b'), 'a &amp; b');

/* ---------- Информационная строка код-фенса ---------- */
check('язык фенса', sanitizeFenceLanguage('TypeScript'), 'typescript');
check('язык фенса с инъекцией', sanitizeFenceLanguage('<img src=x onerror=alert(1)>'), 'imgsrcxonerroralert1');
check('язык фенса с кавычкой', sanitizeFenceLanguage('js" autofocus onfocus="'), 'jsautofocusonfocus');
check('язык фенса с дефисом', sanitizeFenceLanguage('objective-c'), 'objective-c');

console.log('='.repeat(72));
console.log('  ПРОВЕРКА ЭКРАНИРОВАНИЯ HTML В MARKDOWN');
console.log('='.repeat(72));
console.log(`  Пройдено: ${passed}, провалов: ${failures.length}`);
if (failures.length > 0) {
  console.log('-'.repeat(72));
  for (const f of failures) console.log(`   - ${f}`);
}
console.log('='.repeat(72));
console.log(failures.length === 0 ? '  ИТОГ: ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ' : '  ИТОГ: ЕСТЬ ПРОВАЛЫ');

fs.writeFileSync(
  path.join(ROOT, 'sanitizer-check.json'),
  JSON.stringify(
    {
      steps: [
        {
          name: 'sanitizer',
          ok: failures.length === 0,
          passed,
          failures,
        },
      ],
    },
    null,
    2
  ),
  'utf-8'
);

process.exit(failures.length === 0 ? 0 : 1);
