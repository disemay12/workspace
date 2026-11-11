/**
 * Профили промптов: готовые наборы дополнительных инструкций к системному
 * промпту (документы, код, редактура, таблицы) с возможностью своих.
 */
import type { PromptProfile } from '../types';

export const DEFAULT_PROFILES: PromptProfile[] = [
  {
    id: 'general',
    label: 'Обычный помощник',
    prompt: '',
  },
  {
    id: 'editor',
    label: 'Редактор текста',
    prompt:
      'Стиль: короткие абзацы, активный залог, без канцелярита и штампов. ' +
      'Правки предлагай минимальные, сохраняя авторский голос; факты и цифры не меняй.',
  },
  {
    id: 'coder',
    label: 'Разработка',
    prompt:
      'Пиши код в принятом в проекте стиле: строгая типизация, явные имена, без лишних абстракций. ' +
      'Перед изменением читай существующий код; в ответе — что изменено и почему, без пересказа файла.',
  },
  {
    id: 'analyst',
    label: 'Таблицы и данные',
    prompt:
      'Данные проверяй на пропуски и дубли, формулы оставляй формулами (не заменяй значениями). ' +
      'Итоги считай в ячейках, а не в тексте ответа; в чате — короткая сводка и допущения.',
  },
  {
    id: 'deck',
    label: 'Презентации',
    prompt:
      'Один слайд — одна мысль. Заголовки до 60 символов, в теле 3–5 пунктов по 8–10 слов. ' +
      'Таблицы и код — только если без них теряется смысл.',
  },
];

/** Ищет профиль по id, учитывая пользовательские профили и удалённые системные. */
export function resolveProfile(
  id: string | undefined,
  profiles: PromptProfile[] | undefined
): PromptProfile | undefined {
  const list = profiles && profiles.length > 0 ? profiles : DEFAULT_PROFILES;
  if (!id) return undefined;
  return list.find((profile) => profile.id === id);
}

/** Дополнительные инструкции активного профиля ('' для обычного). */
export function profilePrompt(id: string | undefined, profiles: PromptProfile[] | undefined): string {
  return (resolveProfile(id, profiles)?.prompt || '').trim();
}

/** Слияние сохранённых профилей с умолчаниями: новые системные добавляются. */
export function mergeProfiles(saved: PromptProfile[] | undefined): PromptProfile[] {
  if (!saved || saved.length === 0) return [...DEFAULT_PROFILES];
  const byId = new Map(saved.map((profile) => [profile.id, profile]));
  for (const fallback of DEFAULT_PROFILES) {
    if (!byId.has(fallback.id)) byId.set(fallback.id, fallback);
  }
  return [...byId.values()];
}

/** Уникальный id для нового профиля. */
export function nextProfileId(profiles: PromptProfile[]): string {
  let index = 1;
  const taken = new Set(profiles.map((profile) => profile.id));
  while (taken.has(`custom-${index}`)) index += 1;
  return `custom-${index}`;
}
