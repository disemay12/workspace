/**
 * Пользовательские стили абзаца.
 *
 * Определения живут в реестре до сохранения, а при импорте .docx читаются
 * из styles.xml. Это позволяет применять стиль к абзацу и сохранять его
 * определение в пакет, не пересобирая styles.xml целиком.
 */

export interface CustomStyleDefinition {
  id: string;
  name: string;
  basedOn?: string;
  font?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  lineHeight?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
}

const styles = new Map<string, CustomStyleDefinition>();

/** Стабильный styleId из имени: Word не любит пробелы и кириллицу в id. */
export function styleIdFor(name: string): string {
  let hash = 0;
  for (let index = 0; index < name.length; index += 1) {
    hash = (hash * 31 + name.charCodeAt(index)) >>> 0;
  }
  return `CustomStyle${hash.toString(36)}`;
}

export function registerCustomStyle(definition: CustomStyleDefinition): void {
  if (!definition.id) return;
  styles.set(definition.id, { ...definition });
}

export function getCustomStyle(id: string): CustomStyleDefinition | undefined {
  return styles.get(id);
}

export function listCustomStyles(): CustomStyleDefinition[] {
  return [...styles.values()];
}

export function resetCustomStyles(): void {
  styles.clear();
}

function safeColor(value: string | undefined): string {
  if (!value) return '';
  const normalized = value.trim().replace(/^#/, '');
  return /^[0-9a-f]{3,8}$/i.test(normalized) ? `#${normalized}` : '';
}

/** CSS для отображения пользовательского стиля в contenteditable. */
export function customStyleCss(definition: CustomStyleDefinition): string {
  const parts: string[] = [];
  if (definition.font) parts.push(`font-family:"${definition.font.replace(/["']/g, '')}"`);
  if (definition.sizePt) parts.push(`font-size:${definition.sizePt}pt`);
  if (definition.bold) parts.push('font-weight:700');
  if (definition.italic) parts.push('font-style:italic');
  const color = safeColor(definition.color);
  if (color) parts.push(`color:${color}`);
  if (definition.align) parts.push(`text-align:${definition.align}`);
  if (definition.lineHeight) parts.push(`line-height:${definition.lineHeight}`);
  if (definition.spaceBeforePt !== undefined) {
    parts.push(`margin-top:${definition.spaceBeforePt}pt`);
  }
  if (definition.spaceAfterPt !== undefined) {
    parts.push(`margin-bottom:${definition.spaceAfterPt}pt`);
  }
  return parts.join(';');
}
