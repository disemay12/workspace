/**
 * Разбор ответа модели в оформленный HTML документа Word.
 *
 * Модель просят ответить Markdown: заголовки, списки, цитаты, таблицы и
 * выделения. Здесь это превращается в те же теги, которые лента Word уже
 * умеет читать и сохранять обратно в .docx.
 */

export function markdownToDocumentHtml(markdown: string): string {
  const source = stripFence(markdown).replace(/\r\n/g, '\n').trim();
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

/** Убирает обёртку из тройных кавычек, если модель всё-таки её добавила. */
function stripFence(text: string): string {
  const trimmed = text.trim();
  if (!trimmed.startsWith('```')) return text;
  const end = trimmed.lastIndexOf('```');
  if (end <= 3) return text;
  const firstBreak = trimmed.indexOf('\n');
  return trimmed.slice(firstBreak + 1, end);
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
