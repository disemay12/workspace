/**
 * Разбор ответа модели в оформленный HTML документа Word.
 *
 * Модель просят ответить Markdown: заголовки, списки, цитаты, таблицы и
 * выделения. Здесь это превращается в те же теги, которые лента Word уже
 * умеет читать и сохранять обратно в .docx.
 */

export function markdownToDocumentHtml(markdown: string): string {
  const source = normalizeHtmlTags(stripEmoji(unwrapFences(markdown))).replace(/\r\n/g, '\n').trim();
  if (!source) return '';

  const blocks: string[] = [];
  const lines = source.split('\n');
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) {
      i += 1;
      continue;
    }

    const heading = /^(#{1,4})\s+(.+)$/.exec(line.trim());
    if (heading) {
      const level = heading[1].length;
      blocks.push('<h' + level + '>' + inline(heading[2]) + '</h' + level + '>');
      i += 1;
      continue;
    }

    if (/^(-{3,}|\*{3,})$/.test(line.trim())) {
      blocks.push('<hr>');
      i += 1;
      continue;
    }

    if (line.trim().startsWith('>')) {
      const quoted: string[] = [];
      while (i < lines.length && lines[i].trim().startsWith('>')) {
        quoted.push(lines[i].trim().replace(/^>\s?/, ''));
        i += 1;
      }
      blocks.push('<blockquote><p>' + inline(quoted.join(' ')) + '</p></blockquote>');
      continue;
    }

    if (/^\s*([-*•]|\d+[.)])\s+/.test(line)) {
      const ordered = /^\s*\d+[.)]\s+/.test(line);
      const items: string[] = [];
      while (i < lines.length && /^\s*([-*•]|\d+[.)])\s+/.test(lines[i])) {
        items.push(lines[i].replace(/^\s*([-*•]|\d+[.)])\s+/, ''));
        i += 1;
      }
      const tag = ordered ? 'ol' : 'ul';
      blocks.push('<' + tag + '>' + items.map((item) => '<li>' + inline(item) + '</li>').join('') + '</' + tag + '>');
      continue;
    }

    if (line.trim().startsWith('|') && i + 1 < lines.length && /^\s*\|?\s*:?-{3,}/.test(lines[i + 1])) {
      const rows: string[][] = [];
      const header = splitRow(line);
      i += 2;
      while (i < lines.length && lines[i].trim().startsWith('|')) {
        rows.push(splitRow(lines[i]));
        i += 1;
      }
      const head = '<tr>' + header.map((cell) => '<th>' + inline(cell) + '</th>').join('') + '</tr>';
      const body = rows
        .map((row) => '<tr>' + row.map((cell) => '<td>' + inline(cell) + '</td>').join('') + '</tr>')
        .join('');
      blocks.push('<table><thead>' + head + '</thead><tbody>' + body + '</tbody></table>');
      continue;
    }

    const paragraph: string[] = [line.trim()];
    i += 1;
    while (
      i < lines.length &&
      lines[i].trim() &&
      !/^(#{1,4}\s|>|\s*([-*•]|\d+[.)])\s+|\|)/.test(lines[i]) &&
      !/^(-{3,}|\*{3,})$/.test(lines[i].trim())
    ) {
      paragraph.push(lines[i].trim());
      i += 1;
    }
    blocks.push('<p>' + inline(paragraph.join(' ')) + '</p>');
  }

  return blocks.join('');
}

/**
 * Модель часто отвечает HTML-разметкой (<b>, <details>, <summary>).
 * В документе такие теги не поддерживаются, и без этой замены они попадают
 * в текст буквально: «<details> <summary><b>...».
 */
function normalizeHtmlTags(source: string): string {
  return source
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<(b|strong)(\s[^>]*)?>/gi, '**')
    .replace(/<\/(b|strong)>/gi, '**')
    .replace(/<(i|em)(\s[^>]*)?>/gi, '*')
    .replace(/<\/(i|em)>/gi, '*')
    .replace(/<u(\s[^>]*)?>/gi, '__')
    .replace(/<\/u>/gi, '__')
    .replace(/<(s|strike|del)(\s[^>]*)?>/gi, '~~')
    .replace(/<\/(s|strike|del)>/gi, '~~')
    .replace(/<(code)(\s[^>]*)?>/gi, '`')
    .replace(/<\/code>/gi, '`')
    .replace(/<summary(\s[^>]*)?>/gi, '\n')
    .replace(/<\/summary>/gi, '\n')
    .replace(/<\/?details(\s[^>]*)?>/gi, '\n')
    .replace(/<\/?(p|div|section|article)(\s[^>]*)?>/gi, '\n')
    .replace(/<\/(li|h[1-6]|tr|blockquote)>/gi, '\n')
    .replace(/<(li|h[1-6]|tr|blockquote)(\s[^>]*)?>/gi, '\n')
    .replace(/<a\s[^>]*>([\s\S]*?)<\/a>/gi, '$1')
    // Остальные теги (таблицы, картинки, script/style) в документе не нужны.
    .replace(/<[^>]+>/g, '');
}

/**
 * Эмодзи и служебные символы их отрисовки. Модель любит добавлять их в
 * заголовки документов, где они выглядят как мусор.
 */
const EMOJI_RE = /[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{FE0F}\u{200D}]/gu;

export function stripEmoji(text: string): string {
  return text
    .replace(EMOJI_RE, '')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/[ \t]+$/gm, '');
}

/** Хвостовые фразы-предложения, которые не являются содержимым документа. */
const META_TAIL_RE =
  /^(готово|сделано|записано|хотите|хочешь|если хочешь|если хотите|если нужно|если надо|могу|либо|уточните|уточни|что дальше|что скажешь|скажи|варианты|дайте знать|дай знать|сообщите|сообщи|примечание|p\.s\.|ps:)(?=\s|,|:|—|$)/i;

function isMetaTailBlock(block: string): boolean {
  const first = block
    .trim()
    .split('\n')[0]
    .trim()
    .replace(/^[#>*\-\s]+/, '')
    .replace(/^\d+[.)]\s*/, '');
  return META_TAIL_RE.test(first) || isAssistantReportText(block);
}

/** Срезает хвост ответа: «Готово…», «Если хочешь…», «Что скажешь?». */
function stripTrailingMeta(text: string): string {
  const blocks = text.split(/\n{2,}/).map((block) => block.trim());
  while (blocks.length > 0 && isMetaTailBlock(blocks[blocks.length - 1])) {
    blocks.pop();
  }
  return blocks.join('\n\n').trim();
}

/**
 * Убирает из ответа модели эмодзи и хвост «Готово / Хотите, я создам файл? /
 * Уточните…», чтобы в документ попадал только его текст.
 */
export function stripAssistantNoise(text: string): string {
  const cleaned = stripEmoji(text).replace(/\r\n/g, '\n').trim();
  const blocks = cleaned.split(/\n{2,}/);
  if (blocks.length < 2) return cleaned;

  const isMeta = (block: string) => {
    const first = block
      .trim()
      .split('\n')[0]
      .trim()
      .replace(/^[#>*\-\s]+/, '')
      .replace(/^\d+[.)]\s*/, '');
    return META_TAIL_RE.test(first);
  };

  let cut = -1;
  for (let i = 1; i < blocks.length; i += 1) {
    if (isMeta(blocks[i]) && i / blocks.length >= 0.5) {
      cut = i;
      break;
    }
  }
  if (cut > 0) blocks.length = cut;
  return blocks.join('\n\n').trim();
}

/**
 * Снимает все ```-обёртки, оставляя содержимое.
 *
 * Модель часто отвечает «вот документ» и заворачивает его в блок кода с
 * языком text/plain. Если снять только внешнюю пару кавычек, литералы ```
 * и метка языка попадают в документ.
 */
function unwrapFences(text: string): string {
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  const kept: string[] = [];
  for (const line of lines) {
    const fence = /^\s*```[^\s`]*\s?(.*)$/.exec(line);
    if (fence) {
      // «```Директор ООО…» — модель склеила маркер и текст в одну строку.
      if (fence[1].trim()) kept.push(fence[1]);
      continue;
    }
    kept.push(line);
  }
  return kept.join('\n');
}

/** Блоки действий WorkSpace не являются содержимым документа. */
const ACTION_FENCE_RE = /```(?:workspace-action|workspace_action)[^\n]*\n[\s\S]*?```/gi;

/** Вступление чата и короткий отчёт о правке: «Понял Вас», «Исправил дату». */
const PREAMBLE_RE =
  /^(?:#{1,6}\s*)?(?:\*\*)?(?:понял|поняла|вот|ниже|готово|конечно|разумеется|держите|представляю|хорошо|с\s+удовольствием|обновил|обновила|изменил|изменила|исправил|исправила|заменил|заменила|убрал|убрала|удалил|удалила|добавил|добавила|оформил|оформила|настроил|настроила|сделал|сделала|переписал|переписала|сократил|сократила|вставил|вставила)(?![а-яёa-z])/i;

/**
 * Служебный блок ответа: рекомендации по настройке Word, советы «как
 * скопировать», списки полей/шрифтов. Это не текст документа.
 */
export function isDocumentMetaText(raw: string): boolean {
  const text = raw
    .trim()
    .split('\n')[0]
    .replace(/^[#>*\-\s]+/, '')
    .replace(/\*\*/g, '')
    .replace(/^\d+[.)]\s*/, '')
    .trim();

  if (!text) return true;
  if (/^(-{3,}|\*{3,}|_{3,})$/.test(text)) return true;
  if (/^```\s*[\w-]*\s*$/.test(text)) return true;
  if (
    /^(?:что\s+(?:изменено|сделано)|список\s+изменений|изменения\s+в\s+документе|внес[её]нн?ые\s+изменения|кратко\s+об\s+изменениях|отч[её]т\s+об\s+изменениях)(?![а-яёa-z])/i.test(
      text
    )
  ) {
    return true;
  }

  const settings = /(настройк|оформлен|форматир|инструкц|скопируйте|вставьте|откройте|сохраните)/i;
  const target = /(word|ворд|поля|шрифт|интервал|выравнивание|красная строка|страниц)/i;
  return settings.test(text) && target.test(text);
}

function isMetaBlock(block: string): boolean {
  return isDocumentMetaText(block);
}

/**
 * Служебный отчёт ассистента о проделанной работе. Модель иногда кладёт его
 * в HTML документа: «Оформил документ… Вот итоговая версия:», «Документ
 * Test.docx обновлён… сохранить сейчас?».
 */
const REPORT_START_RE =
  /^(?:оформил|оформила|обновил|обновила|изменил|изменила|заменил|заменила|исправил|исправила|переписал|переписала|сделал|сделала|привёл|привел|привела|подготовил|подготовила|сохранил|сохранила|меняю|перерабатываю|обновляю|переписываю|оформляю|правлю|заменяю|добавляю|удаляю|сокращаю|делаю|вношу|привожу|готовлю|все\s+(?:разделы|пункты|части)\s+(?:обновлен|обновлены|переписан|переписаны|изменен|изменены))(?![а-яёa-z])/i;

const REPORT_PHRASE_RE =
  /(вот\s+итогов|итоговая\s+верси|готов(?:ая|ый)\s+верси|документ\s+[^\n]{0,100}\s+обновл[её]н|помечен\s+как\s+несохран[её]нный|сохранить\s+сейчас|файл\s+[^\n]{0,40}измен[её]н[^\n]{0,80}(?:не\s+сохран|сохран)|^обратите\s+внимание(?![а-яёa-z])|^остал[а-я]*\s+техническ)/i;

export function isAssistantReportText(raw: string): boolean {
  const text = raw
    .replace(/\r\n/g, '\n')
    .replace(/^[#>*\-\s]+/, '')
    .replace(/\*\*/g, '')
    .trim();
  if (!text) return false;
  if (REPORT_PHRASE_RE.test(text)) return true;
  return (
    REPORT_START_RE.test(text) &&
    /(документ|текст|резюме|верси|стил|файл|професси|должност|раздел|заголов|назван)/i.test(text)
  );
}

function isAssistantReportListText(block: string): boolean {
  const lines = block
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => /^(?:[-*•]|\d+[.)])\s+/.test(line));
  if (lines.length === 0) return false;
  const changed = lines.filter((line) =>
    /(→|->|(?:обновлен|переименован|переписан|заменен|добавлен|удален|исправлен|сокращен|настроен))/i.test(
      line
    )
  ).length;
  return changed >= Math.ceil(lines.length * 0.6);
}

function isBlankBlock(block: string): boolean {
  const text = block
    .replace(/^[#>*\-\s]+/, '')
    .replace(/\*\*/g, '')
    .replace(/^```\s*[\w-]*\s*$/, '')
    .trim();
  return !text || /^(-{3,}|\*{3,}|_{3,})$/.test(text);
}

/**
 * Содержимое всех ```-блоков. Метку языка отбрасываем, а текст, склеенный
 * с открывающим маркером, сохраняем: «```Директор ООО…».
 */
function collectFencedBlocks(text: string): string[] {
  const blocks: string[] = [];
  let buffer: string[] | null = null;

  for (const line of text.split('\n')) {
    const fence = /^\s*```[^\s`]*\s?(.*)$/.exec(line);
    if (fence) {
      if (buffer === null) {
        buffer = [];
        if (fence[1].trim()) buffer.push(fence[1].trim());
      } else {
        const body = buffer.join('\n').trim();
        if (body) blocks.push(body);
        buffer = null;
      }
      continue;
    }
    if (buffer !== null) buffer.push(line);
  }

  if (buffer !== null) {
    const body = buffer.join('\n').trim();
    if (body) blocks.push(body);
  }

  // JSON-действия и ответы инструментов — не документ.
  return blocks.filter((body) => !/^\s*\{\s*"(?:action|actions|tool|operations)"/.test(body));
}

/**
 * Вытаскивает из ответа ассистента именно содержимое документа.
 *
 * Приоритет — готовый блок кода: модель кладёт туда текст заявления, письма
 * или договора. Если блока нет, отбрасываются вступление, разделители и
 * служебные советы по оформлению, из-за которых в .docx раньше попадали
 * фразы «Понял Вас», «Рекомендуемые настройки страницы» и литералы ```.
 */
export function extractDocumentContent(answer: string): string {
  const source = stripEmoji(answer.replace(/\r\n/g, '\n')).trim();
  if (!source) return '';

  const withoutActions = source.replace(ACTION_FENCE_RE, '').trim();
  if (!withoutActions) return '';

  const fenced = collectFencedBlocks(withoutActions);
  if (fenced.length > 0) {
    const text = stripTrailingMeta(fenced.join('\n\n'));
    if (text) return text;
  }

  const blocks = withoutActions.split(/\n{2,}/).map((block) => block.trim()).filter(Boolean);
  if (blocks.length === 0) return '';

  let start = 0;
  while (start < blocks.length - 1 && start < 2 && PREAMBLE_RE.test(blocks[start])) start += 1;

  const kept: string[] = [];
  // Список сразу после «Рекомендуемые настройки…» — тоже служебный:
  // маркеры полей/шрифтов идут отдельным блоком от своего заголовка.
  let afterMeta = false;
  for (const block of blocks.slice(start)) {
    if (isMetaBlock(block)) {
      afterMeta = true;
      continue;
    }
    if (isAssistantReportText(block)) {
      afterMeta = false;
      continue;
    }
    if (isAssistantReportListText(block)) {
      afterMeta = false;
      continue;
    }
    // Отчёт агента списком: «— Исправил дату», «— Заменил сумму».
    if (
      /^\s*(?:[-*•]|\d+[.)])\s+(?:исправил|исправила|заменил|заменила|обновил|обновила|добавил|добавила|удалил|удалила|изменил|изменила|оформил|оформила|настроил|настроила|сделал|сделала|переписал|переписала|сократил|сократила|вставил|вставила)(?![а-яёa-z])/i.test(
        block
      )
    ) {
      continue;
    }
    if (afterMeta && /^\s*([-*•]|\d+[.)])\s+/.test(block)) continue;
    afterMeta = false;
    if (isBlankBlock(block)) continue;
    kept.push(block);
  }

  const result = stripTrailingMeta(kept.join('\n\n'));
  return result || stripTrailingMeta(withoutActions);
}

function splitRow(line: string): string[] {
  return line
    .trim()
    .replace(/^\|/, '')
    .replace(/\|$/, '')
    .split('|')
    .map((cell) => cell.trim());
}

/** Жирный, курсив, подчёркивание, зачёркивание и код внутри строки. */
function inline(text: string): string {
  let html = escapeHtml(text);
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/__([^_]+)__/g, '<u>$1</u>');
  html = html.replace(/(^|[^*])\*([^*\n]+)\*(?![*])/g, '$1<em>$2</em>');
  html = html.replace(/~~([^~]+)~~/g, '<s>$1</s>');
  html = html.replace(/`([^`]+)`/g, '<span style="font-family:Consolas;font-size:11pt">$1</span>');
  return html;
}

function escapeHtml(text: string): string {
  return text.replace(/[&<>]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[char] ?? char));
}
