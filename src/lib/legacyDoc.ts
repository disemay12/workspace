/**
 * Устаревший формат .doc (Word 97-2003).
 *
 * Редактор работает с .docx, поэтому документ при открытии разбирается в
 * текст и упаковывается в OOXML-представление, а при сохранении рядом
 * создаётся файл .docx — оригинал .doc остаётся нетронутым.
 */

/** Текст .doc: разбор OLE2 в главном процессе (нужен Buffer). */
export async function extractLegacyDocText(base64: string): Promise<string> {
  const api = window.electronAPI;
  if (!api?.extractDocText) {
    throw new Error('Чтение .doc доступно только в настольном приложении WorkSpace');
  }
  const result = await api.extractDocText({ base64 });
  if (!result.success || result.text === undefined) {
    throw new Error(result.error || 'Не удалось прочитать документ .doc');
  }
  return result.text;
}

/** Простой текст → абзацы HTML для полотна Word. */
export function plainTextToHtml(text: string): string {
  const escape = (value: string) =>
    value.replace(/[&<>]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[char] ?? char));
  const blocks = text
    .split(/\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => `<p>${escape(line)}</p>`);
  return blocks.join('') || '<p></p>';
}

/** Путь для сохранения рядом: документ.doc → документ.docx. */
export function docxPathFor(docPath: string): string {
  return /\.doc$/i.test(docPath) ? `${docPath.slice(0, -4)}.docx` : `${docPath}.docx`;
}
