/**
 * Сборка презентации .pptx из размеченного текста (Markdown, txt, md-конверсии).
 *
 * Использует модель документа `PptxDocument` и пустой шаблон пакета: слайды
 * создаются операциями OOXML, поэтому результат открывается в PowerPoint без
 * «восстановления» и дальше редактируется обычными средствами.
 *
 * Правила разбиения на слайды (детерминированные, их же описывает подсказка
 * в интерфейсе):
 *   `# Заголовок`   — новый слайд с этим заголовком;
 *   `## Подраздел`  — новый слайд, если на текущем уже есть содержимое;
 *   `---`           — принудительный разрыв слайда;
 *   списки          — маркеры, обычные строки — абзацы;
 *   ```блок```      — моноширинный текст в теле слайда;
 *   таблица Markdown — настоящая таблица PowerPoint.
 */
import { BLANK_PPTX_BASE64 } from './pptxTemplate';
import { openPptxDocument, type PptxDocument, type PptxParagraph } from './pptxDoc';
import { normalizeHexColor, presetById, type ThemeColorSet } from './pptxThemes';

export interface SlideBlock {
  kind: 'bullet' | 'number' | 'paragraph' | 'code';
  text: string;
}

export interface SlidePlan {
  title: string;
  subtitle?: string;
  blocks: SlideBlock[];
  table?: string[][];
  notes?: string;
}

const MAX_SLIDES = 80;
const MAX_BLOCKS_PER_SLIDE = 40;
const MAX_TABLE_COLS = 8;

/** Снятие инлайн-разметки: **жирный**, *курсив*, `код`, [текст](ссылка). */
export function stripInlineMarkdown(value: string): string {
  return value
    .replace(/!\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
    .replace(/`{1,2}([^`]*)`{1,2}/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/(^|[^*])\*([^*]+)\*/g, '$1$2')
    .replace(/__([^_]+)__/g, '$1')
    .replace(/(^|[^_])_([^_]+)_/g, '$1$2')
    .replace(/~~([^~]+)~~/g, '$1')
    .trim();
}

function isTableSeparator(line: string): boolean {
  return /^\s*\|?[\s:-]+\|[\s|:-]*$/.test(line) && line.includes('-');
}

function splitTableRow(line: string): string[] {
  const trimmed = line.trim().replace(/^\|/, '').replace(/\|$/, '');
  return trimmed.split('|').map((cell) => stripInlineMarkdown(cell));
}

/**
 * Разбиение Markdown на слайды. Чистая функция: её можно проверить отдельно
 * от сборки пакета.
 */
export function planSlidesFromMarkdown(markdown: string, fallbackTitle = 'Презентация'): SlidePlan[] {
  const lines = String(markdown || '').replace(/\r\n?/g, '\n').split('\n');
  const plans: SlidePlan[] = [];
  let current: SlidePlan | null = null;
  let paragraph: string[] = [];
  let codeOpen = false;
  let codeLines: string[] = [];
  let tableRows: string[][] = [];
  /** Черновик заметок: исходные строки Markdown текущего слайда. */
  let pendingNotes: string[] = [];

  const finalizeNotes = () => {
    if (current && pendingNotes.length > 0 && !current.notes) {
      current.notes = pendingNotes.slice(0, 24).join('\n');
    }
    pendingNotes = [];
  };

  const ensureSlide = (title?: string): SlidePlan => {
    if (!current) {
      current = { title: title?.trim() || fallbackTitle, blocks: [] };
      plans.push(current);
      return current;
    }
    if (title && current.title === fallbackTitle && current.blocks.length === 0) {
      // Пустая заготовка без содержимого — просто уточняем заголовок.
      current.title = title.trim();
      return current;
    }
    if (title) {
      finalizeNotes();
      current = { title: title.trim(), blocks: [] };
      plans.push(current);
    }
    return current;
  };

  const flushParagraph = () => {
    if (paragraph.length === 0) return;
    const text = stripInlineMarkdown(paragraph.join(' '));
    paragraph = [];
    if (!text) return;
    ensureSlide().blocks.push({ kind: 'paragraph', text });
  };

  const flushTable = () => {
    if (tableRows.length === 0) return;
    const slide = ensureSlide();
    slide.table = tableRows.slice(0, 30);
    tableRows = [];
  };

  for (const raw of lines) {
    const line = raw.trimEnd();
    pendingNotes.push(raw);

    if (codeOpen) {
      if (/^\s*```/.test(line)) {
        codeOpen = false;
        if (codeLines.length > 0) {
          ensureSlide().blocks.push({ kind: 'code', text: codeLines.join('\n') });
          codeLines = [];
        }
        continue;
      }
      codeLines.push(raw);
      continue;
    }

    if (/^\s*```/.test(line)) {
      flushParagraph();
      flushTable();
      codeOpen = true;
      continue;
    }

    const heading = /^(#{1,6})\s+(.*)$/.exec(line);
    if (heading) {
      flushParagraph();
      flushTable();
      const level = heading[1].length;
      const text = stripInlineMarkdown(heading[2]);
      if (!text) continue;
      if (level === 1) {
        const slide = ensureSlide(text);
        slide.title = text;
      } else if (current && ((current as SlidePlan).blocks.length > 0 || (current as SlidePlan).table)) {
        ensureSlide(text);
      } else {
        ensureSlide(text);
      }
      continue;
    }

    if (/^\s*(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) {
      flushParagraph();
      flushTable();
      finalizeNotes();
      current = { title: fallbackTitle, blocks: [] };
      plans.push(current);
      continue;
    }

    if (/^\s*\|/.test(line)) {
      flushParagraph();
      if (isTableSeparator(line)) continue;
      const cells = splitTableRow(line);
      if (cells.length >= 2) tableRows.push(cells.slice(0, MAX_TABLE_COLS));
      continue;
    }
    flushTable();

    const bullet = /^\s*[-*+]\s+(.*)$/.exec(line);
    if (bullet) {
      flushParagraph();
      const text = stripInlineMarkdown(bullet[1]);
      if (!text) continue;
      const slide = ensureSlide();
      if (slide.blocks.length < MAX_BLOCKS_PER_SLIDE) slide.blocks.push({ kind: 'bullet', text });
      continue;
    }

    const numbered = /^\s*\d+[.)]\s+(.*)$/.exec(line);
    if (numbered) {
      flushParagraph();
      const text = stripInlineMarkdown(numbered[1]);
      if (!text) continue;
      const slide = ensureSlide();
      if (slide.blocks.length < MAX_BLOCKS_PER_SLIDE) slide.blocks.push({ kind: 'number', text });
      continue;
    }

    if (!line.trim()) {
      flushParagraph();
      continue;
    }

    // Обычный текст: копим абзац, пока не встретим пустую строку или разметку.
    paragraph.push(line.trim());
  }

  if (codeOpen && codeLines.length > 0) {
    ensureSlide().blocks.push({ kind: 'code', text: codeLines.join('\n') });
  }
  flushParagraph();
  flushTable();
  finalizeNotes();

  const cleaned = plans
    .map((plan) => ({
      ...plan,
      title: plan.title && plan.title !== fallbackTitle ? plan.title : plan.title || fallbackTitle,
      blocks: plan.blocks.filter((block) => block.text.trim() !== ''),
    }))
    .filter((plan) => plan.title || plan.blocks.length > 0 || plan.table);

  if (cleaned.length === 0) {
    return [{ title: fallbackTitle, blocks: [{ kind: 'paragraph', text: 'Пустой документ' }] }];
  }

  // Сверхлимитные слайды отбрасываем, но не молча: об этом сообщает вызывающий.
  return cleaned.slice(0, MAX_SLIDES);
}

export interface BuildPptxOptions {
  title?: string;
  /** Пресет темы из `pptxThemes.ts`. */
  themePresetId?: string;
}

/** Сборка пакета .pptx из Markdown. Возвращает base64. */
export async function markdownToPptxBase64(
  markdown: string,
  options: BuildPptxOptions = {}
): Promise<string> {
  const title = (options.title || 'Презентация').trim();
  const plans = planSlidesFromMarkdown(markdown, title);
  const doc = await openPptxDocument(BLANK_PPTX_BASE64);
  buildSlides(doc, plans, options);
  return doc.toBase64();
}

/** Сборка слайдов в уже открытый документ (проверяется в model-check). */
export function buildSlides(doc: PptxDocument, plans: SlidePlan[], options: BuildPptxOptions = {}): void {
  const preset = presetById(options.themePresetId || 'office');
  doc.applyTheme({ presetId: preset.id, label: preset.label });
  // Презентации из текста собираем в 16:9: шаблон пакета — 4:3.
  doc.setSlideSize(12192000, 6858000);
  const colors = doc.model.theme.colors;

  plans.forEach((plan) => {
    const slideIndex = doc.addSlide(doc.model.slides.length, undefined, plan.title);
    const slide = doc.model.slides[slideIndex];
    if (!slide) return;
    const titleShape = slide.shapes.find((shape) => /заголовок|title/i.test(shape.name)) ?? slide.shapes[0];
    setShapeText(doc, slideIndex, titleShape?.id, plan.title, {
      size: preset.id === 'mono' ? 28 : 32,
      bold: true,
      color: normalizeHexColor(colors.dk1),
    });

    if (plan.table && plan.table.length > 0) {
      const rows = plan.table.length;
      const cols = Math.max(1, ...plan.table.map((row) => row.length));
      // Тело слайда очищаем: иначе под таблицей остаётся текст-заглушка шаблона.
      const bodyShape =
        slide.shapes.find((shape) => /текст|body/i.test(shape.name) && shape.id !== titleShape?.id) ??
        slide.shapes.find((shape) => shape.id !== titleShape?.id);
      if (bodyShape?.id) doc.setShapeText(slideIndex, bodyShape.id, { paragraphs: [{ runs: [] }] });
      doc.addTable(slideIndex, {
        rows,
        cols,
        data: plan.table,
        x: Math.round(doc.model.width * 0.08),
        y: Math.round(doc.model.height * 0.3),
        w: Math.round(doc.model.width * 0.84),
        h: Math.round(doc.model.height * 0.5),
      });
      return;
    }

    const bodyShape =
      slide.shapes.find((shape) => /текст|body/i.test(shape.name) && shape.id !== titleShape?.id) ??
      slide.shapes.find((shape) => shape.id !== titleShape?.id);
    if (bodyShape?.id) {
      const paragraphs = paragraphsFor(plan.blocks, colors);
      doc.setShapeText(slideIndex, bodyShape.id, { paragraphs, verticalAnchor: 't' });
    }
  });

  // Первый лист шаблона («Новый слайд») больше не нужен.
  if (doc.model.slides.length > plans.length) {
    doc.deleteSlide(0);
  }
}

function setShapeText(
  doc: PptxDocument,
  slideIndex: number,
  shapeId: string | undefined,
  text: string,
  style: { size: number; bold?: boolean; color?: string; family?: string }
): void {
  if (!shapeId) return;
  doc.setShapeText(slideIndex, shapeId, {
    paragraphs: text
      .split('\n')
      .map((line) => ({
        runs: line
          ? [
              {
                text: line,
                size: style.size,
                bold: style.bold,
                color: style.color ? style.color.replace('#', '') : undefined,
                family: style.family,
              },
            ]
          : [],
      })),
    verticalAnchor: 't',
  });
}

function paragraphsFor(blocks: SlideBlock[], colors: ThemeColorSet): PptxParagraph[] {
  const bodyColor = normalizeHexColor(colors.dk2);
  const paragraphs: PptxParagraph[] = [];
  let numberCounter = 1;

  for (const block of blocks) {
    if (block.kind === 'code') {
      for (const line of block.text.split('\n').slice(0, 24)) {
        paragraphs.push({
          runs: [{ text: line || ' ', size: 12, family: 'Consolas', color: '1F2937' }],
          bullet: 'none',
        });
      }
      continue;
    }
    if (block.kind === 'paragraph') {
      paragraphs.push({
        runs: [{ text: block.text, size: 18, color: bodyColor }],
        bullet: 'none',
      });
      continue;
    }
    if (block.kind === 'number') {
      paragraphs.push({
        runs: [{ text: `${numberCounter}. ${block.text}`, size: 18, color: bodyColor }],
        bullet: 'none',
      });
      numberCounter++;
      continue;
    }
    paragraphs.push({
      runs: [{ text: block.text, size: 18, color: bodyColor }],
      bullet: 'bullet',
    });
  }
  return paragraphs.length > 0 ? paragraphs : [{ runs: [] }];
}
