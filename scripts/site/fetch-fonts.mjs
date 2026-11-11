// Скачивает шрифты сайта в docs/assets/fonts и пишет docs/fonts.css.
// Нужен только при первом наполнении: страница не ходит за шрифтами в сеть.
//
//   node scripts/site/fetch-fonts.mjs
//
// Host Grotesk и Atkinson Hyperlegible Mono (тема Hallmark Console) не имеют
// кириллицы, поэтому для русской страницы взяты Golos Text и JetBrains Mono:
// тот же регистр — нейтральный гротеск и моноширинный для измеренных величин.

import { mkdirSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const outDir = join(root, 'docs');
const fontDir = join(outDir, 'assets', 'fonts');

const FAMILIES = [
  { family: 'Golos Text', slug: 'golos-text', weights: [400, 500] },
  { family: 'JetBrains Mono', slug: 'jetbrains-mono', weights: [400, 500] },
];
const KEEP = new Set(['latin', 'latin-ext', 'cyrillic']);
const UA =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

const cssUrl = (family, weights) =>
  `https://fonts.googleapis.com/css2?family=${family.replace(/ /g, '+')}:wght@${weights.join(';')}&display=swap`;

async function main() {
  mkdirSync(fontDir, { recursive: true });
  const rules = [];

  for (const { family, slug, weights } of FAMILIES) {
    const css = await (await fetch(cssUrl(family, weights), { headers: { 'User-Agent': UA } })).text();
    const blocks = css.split('/*').slice(1);
    for (const block of blocks) {
      const subset = block.slice(0, block.indexOf('*/')).trim();
      if (!KEEP.has(subset)) continue;
      const body = block.slice(block.indexOf('*/') + 2);
      const weight = /font-weight:\s*(\d+)/.exec(body)?.[1];
      const url = /url\((https:[^)]+\.woff2)\)/.exec(body)?.[1];
      const range = /unicode-range:\s*([^;]+);/.exec(body)?.[1];
      if (!weight || !url) continue;

      const file = `${slug}-${subset}-${weight}.woff2`;
      const bytes = new Uint8Array(await (await fetch(url, { headers: { 'User-Agent': UA } })).arrayBuffer());
      writeFileSync(join(fontDir, file), bytes);
      rules.push(
        `@font-face {\n  font-family: "${family}";\n  font-style: normal;\n  font-weight: ${weight};\n` +
          `  font-display: swap;\n  src: url("./assets/fonts/${file}") format("woff2");\n` +
          (range ? `  unicode-range: ${range.trim()};\n` : '') +
          '}'
      );
      console.log(`${file}  ${(bytes.length / 1024).toFixed(1)} КБ  ${subset} ${weight}`);
    }
  }

  writeFileSync(join(outDir, 'fonts.css'), `/* Шрифты страницы: локальные woff2, без обращений к CDN. */\n${rules.join('\n')}\n`);
  console.log(`docs/fonts.css — ${rules.length} правил`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
