/**
 * Общие IPC-заглушки для проверочных сценариев.
 *
 * Каждый GUI-набор (render-check, feature-check, ai-check, insert-check,
 * integration-check, live-check) поднимает окно Electron и сам регистрирует
 * обработчики IPC вместо main.ts. Списки обработчиков в этих файлах
 * дублировались и со временем разошлись с main.ts: в них не хватало
 * fs:getFavorites, fs:toggleFavorite, clipboard:read/write и других.
 *
 * Из-за отсутствия fs:getFavorites рендерер падал на старте
 * («No handler registered for 'fs:getFavorites'»), дерево файлов не
 * загружалось, и вместе с ним падали все зависимые проверки.
 *
 * Здесь собраны недостающие обработчики. Вызывайте registerCommonCheckIpc()
 * внутри registerIPC() каждого набора — тогда новые обработчики main.ts
 * достаточно добавить один раз, и все сценарии получат их автоматически.
 */
import { ipcMain, clipboard } from 'electron';

/** Состояние избранного на время прогона (в памяти, как в main.ts). */
const favorites: { path: string; name: string }[] = [];

/**
 * Каналы, уже зарегистрированные helper'ом. Electron не даёт публичного способа
 * узнать, занят ли канал для invoke, поэтому ведём учёт самостоятельно.
 */
const registeredChannels = new Set<string>();

export interface CommonCheckIpcOptions {
  /**
   * Корневая папка сценария. Используется для fs:getFavorites, чтобы вернуть
   * непустой список и проверить, что раздел избранного действительно рисуется.
   */
  sandbox?: string;
  /** Вызывается, когда рендерер подтверждает закрытие окна. */
  onConfirmClose?: (proceed: boolean) => void;
  /**
   * Начальные настройки AI, которые вернёт aiSettings:load.
   *
   * Постоянное хранилище важнее localStorage (settings.ts: «постоянный файл
   * важнее localStorage»), поэтому сценарий, которому нужны конкретные настройки,
   * должен задать их здесь — иначе пустой ответ перекроет значения из localStorage.
   */
  aiSettings?: Record<string, unknown>;
  /** Начальные ключи провайдеров для aiSettings:load. */
  aiKeys?: Record<string, string>;
  /**
   * Функция, возвращающая актуальные настройки на момент вызова aiSettings:load.
   *
   * Нужна, когда сценарий меняет настройки между фазами: объект, переданный в
   * aiSettings, захватывается по ссылке, и его замена через присваивание
   * свойства осталась бы невидимой.
   */
  getAiSettings?: () => Record<string, unknown>;
  /** Функция, возвращающая актуальные ключи провайдеров. */
  getAiKeys?: () => Record<string, string>;
}

/**
 * Регистрирует обработчики, которых обычно не хватает проверочным сценариям.
 * Безопасно вызывать повторно: повторная регистрация того же канала
 * пропускается, чтобы не получить «second handler» ошибку.
 */
export function registerCommonCheckIpc(options: CommonCheckIpcOptions = {}): void {
  const {
    sandbox = '',
    onConfirmClose,
    aiSettings: initialAiSettings,
    aiKeys,
    getAiSettings,
    getAiKeys,
  } = options;

  /**
   * Регистрирует канал, если он ещё не занят.
   *
   * Проверять занятость через ipcMain.listenerCount нельзя: handle() не создаёт
   * слушателя события, поэтому список всегда пуст и повторная регистрация падала
   * бы с «Attempted to register a second handler». Учитываем каналы сами, а также
   * каналы, уже зарегистрированные этим helper'ом в предыдущем вызове.
   */
  const handle = (channel: string, listener: (event: any, ...args: any[]) => any) => {
    if (registeredChannels.has(channel)) return;
    registeredChannels.add(channel);
    try {
      ipcMain.handle(channel, listener);
    } catch (err) {
      // Канал уже занят вызывающим сценарием — это нормально, оставляем его версию.
      registeredChannels.add(channel);
      void err;
    }
  };

  /* ---------- Избранное ---------- */
  handle('fs:getFavorites', () => {
    // Возвращаем песочницу как избранное: наборы проверяют, что список
    // отрисовался, поэтому пустой массив сделал бы проверку бессмысленной.
    if (sandbox && !favorites.some((f) => f.path === sandbox)) {
      return [{ path: sandbox, name: 'Избранное' }];
    }
    return favorites.slice();
  });

  handle('fs:toggleFavorite', (_event, dirPath: string) => {
    const exists = favorites.some((f) => f.path === dirPath);
    const next = exists
      ? favorites.filter((f) => f.path !== dirPath)
      : favorites.concat({ path: dirPath, name: dirPath.split(/[\\/]/).pop() || dirPath });
    favorites.length = 0;
    favorites.push(...next);
    return { success: true, favorites: favorites.slice() };
  });

  /* ---------- Буфер обмена ---------- */
  handle('clipboard:read', async () => {
    try {
      const items = await clipboard.read();
      let html = '';
      for (const item of items) {
        if (item.types.includes('text/html')) {
          const blob = await item.getType('text/html');
          html = await blob.text();
          break;
        }
      }
      return { success: true, text: await clipboard.readText(), html };
    } catch (err) {
      return { success: false, text: '', html: '', error: (err as Error).message };
    }
  });

  handle('clipboard:write', async (_event, text: string) => {
    try {
      clipboard.writeText(String(text ?? ''));
      return { success: true };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  /* ---------- Подтверждение закрытия окна ---------- */
  handle('window:confirmClose', (_event, proceed: boolean) => {
    onConfirmClose?.(Boolean(proceed));
    return { closed: Boolean(proceed) };
  });

  /* ---------- Настройки AI ---------- */
  // В проверочном сценарии постоянного хранилища нет: приложение получает пустые
  // настройки и сохраняет их только в памяти процесса.
  // Настройки читаются в момент вызова: сценарий может менять их между фазами
  // (например, подменять baseUrl для проверки ошибки авторизации).
  const readSettings = () => ({ ...(initialAiSettings ?? {}), ...(getAiSettings?.() ?? {}) });
  const readKeys = () => ({ ...(aiKeys ?? {}), ...(getAiKeys?.() ?? {}) });

  /** Значения, сохранённые самим приложением через aiSettings:save. */
  const savedSettings: Record<string, unknown> = {};

  handle('aiSettings:load', () => ({
    // Сохранённое приложением дополняется настройками сценария: последние
    // имитируют постоянный файл и потому авторитетнее.
    settings: { ...savedSettings, ...readSettings() },
    keys: { ...readKeys() },
  }));

  handle('aiSettings:save', (_event, patch: Record<string, unknown> = {}) => {
    const { apiKey: _apiKey, keys: _keys, ...rest } = patch;
    Object.assign(savedSettings, rest);
    return { success: true };
  });

  /* ---------- Секреты ---------- */
  handle('secrets:describe', () => ({
    found: false,
    filePath: '',
    baseUrl: '',
    textModel: '',
    imageModel: '',
    hasKey: false,
    hasOtherSecrets: false,
    error: '',
  }));

  handle('secrets:importKey', () => ({ success: false, error: 'Недоступно в проверочном сценарии' }));
}
