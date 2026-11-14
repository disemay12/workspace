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
import { app, ipcMain, clipboard } from 'electron';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Версия приложения — из package.json: заглушка отдавала '0.0.0-check', и
 * проверка формы ответа update:check не могла сверить её с боевой.
 */
function resolveAppVersion(): string {
  let dir = __dirname;
  for (let i = 0; i < 5; i++) {
    try {
      const pkg = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf-8'));
      if (pkg && typeof pkg.version === 'string') return pkg.version;
    } catch {
      // Идём уровнем выше.
    }
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return '0.0.0';
}
const APP_VERSION = resolveAppVersion();

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

  /* ---------- Сессия окна ----------
   * Приложение сохраняет раскладку дебаунсом через session:patch, а при старте
   * читает session:get. Без этих каналов рендерер получал «No handler
   * registered» уже на первых секундах и набор падал на ошибках консоли.
   * Храним в файле песочницы, как main.ts в userData: заглушка в памяти
   * расходилась с боевой и не проверяла запись. */
  const sessionFile = sandbox ? path.join(sandbox, 'workspace-session.json') : '';
  // Раскладка прошлого прогона (открытые вкладки, активный лист) не должна
  // переноситься в текущий: иначе проверки зависят от порядка запусков —
  // например, «Полужирный» снимал жирный, уже сохранённый предыдущим прогоном.
  if (sessionFile) {
    try {
      fs.rmSync(sessionFile, { force: true });
    } catch {
      // Сессии нет или файл занят — стартуем с пустой раскладки.
    }
  }
  const sessionMemory: Record<string, unknown> = {};
  const readSession = (): Record<string, unknown> => {
    if (!sessionFile) return { ...sessionMemory };
    try {
      const parsed = JSON.parse(fs.readFileSync(sessionFile, 'utf-8'));
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch {
      return {};
    }
  };
  handle('session:get', () => readSession());
  handle('session:patch', (_event, patch: Record<string, unknown> = {}) => {
    const next = { ...readSession(), ...(patch ?? {}) };
    if (!sessionFile) Object.assign(sessionMemory, next);
    else {
      try {
        fs.mkdirSync(path.dirname(sessionFile), { recursive: true });
        fs.writeFileSync(sessionFile, JSON.stringify(next, null, 2), 'utf-8');
      } catch {
        // Каталог данных может быть недоступен — сессия остаётся в памяти.
        Object.assign(sessionMemory, next);
      }
    }
    return next;
  });

  /* ---------- Автосохранение ----------
   * Снимки читаются на старте (уборка просроченных), поэтому каналы нужны
   * каждому сценарию, а не только тем, что проверяют восстановление. */
  const snapshots = new Map<string, unknown>();
  handle('autosave:write', (_event, name: string, payload: unknown) => {
    snapshots.set(String(name), payload);
    return { success: true };
  });
  handle('autosave:list', () => ({ success: true, items: [...snapshots.values()] }));
  /**
   * Метаданные снимков без содержимого: рендерер опрашивает их заметно чаще
   * полного списка, и форма совпадает с main.ts (content не отдаётся).
   */
  handle('autosave:listMeta', () => ({
    success: true,
    items: [...snapshots.entries()].map(([name, payload]) => {
      const snapshot = (payload ?? {}) as Record<string, unknown>;
      return {
        id: snapshot.id,
        path: snapshot.path,
        name: snapshot.name ?? name,
        type: snapshot.type,
        savedAt: snapshot.savedAt,
        size: typeof snapshot.content === 'string' ? snapshot.content.length : 0,
      };
    }),
  }));
  handle('autosave:read', (_event, name: string) => {
    const snapshot = snapshots.get(String(name));
    return snapshot === undefined
      ? { success: false, error: 'снимок не найден' }
      : { success: true, snapshot };
  });
  handle('autosave:remove', (_event, name: string) => {
    snapshots.delete(String(name));
    return { success: true };
  });

  /* ---------- Диагностика и обновления ----------
   * Рендерер пишет ошибки в diagnostics:log и проверяет обновления на старте.
   * Форма ответов взята из main.ts/updater.ts: расходившаяся заглушка
   * (`{success:false}`, версия '0.0.0-check') проходила проверку, хотя боевой
   * обработчик отдаёт отчёт и настоящую версию. */
  const diagnostics: { at: number; level: string; message: string }[] = [];
  handle('diagnostics:log', (_event, level: string, message: string) => {
    diagnostics.push({
      at: Date.now(),
      level: String(level || 'error'),
      message: String(message ?? '').slice(0, 2000),
    });
    if (diagnostics.length > 200) diagnostics.shift();
    return { success: true };
  });
  handle('diagnostics:report', () => ({
    success: true,
    report: {
      app: { version: APP_VERSION, name: 'WorkSpace', packaged: false, locale: app.getLocale() },
      system: {
        platform: process.platform,
        arch: process.arch,
        release: (process as any).getSystemVersion?.() ?? '',
        versions: {
          electron: process.versions.electron,
          chrome: process.versions.chrome,
          node: process.versions.node,
        },
        paths: { userData: app.getPath('userData'), home: app.getPath('home'), appData: app.getPath('appData') },
        memoryMb: Math.round(process.memoryUsage().rss / (1024 * 1024)),
      },
      autosave: { dir: sandbox ? path.join(sandbox, 'autosave') : '', snapshots: snapshots.size },
      logs: [...diagnostics],
    },
  }));
  handle('update:getVersion', () => APP_VERSION);
  handle('update:check', () => ({
    success: true,
    supported: false,
    currentVersion: APP_VERSION,
    error: 'Проверка обновлений доступна только в собранной portable-версии WorkSpace',
  }));

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
  /**
   * Ключи, сохранённые приложением. Боевой обработчик aiSettings:save
   * шифрует patch.apiKey в data.keys[provider] — заглушка должна делать то же
   * самое, иначе после импорта ключа приложение выглядит «не настроенным».
   */
  const savedKeys: Record<string, string> = {};

  handle('aiSettings:load', () => ({
    // Сохранённое приложением дополняется настройками сценария: последние
    // имитируют постоянный файл и потому авторитетнее.
    settings: { ...savedSettings, ...readSettings() },
    keys: { ...readKeys(), ...savedKeys },
    keyErrors: [],
    encryptionUnavailable: false,
  }));

  // Ключ конкретного провайдера: замена `keys` в load — там только активный.
  handle('aiSettings:key', (_event, provider: string) => ({
    key: savedKeys[provider] || readKeys()[provider] || '',
    unreadable: false,
  }));

  handle('aiSettings:save', (_event, patch: Record<string, unknown> = {}) => {
    const { apiKey, keys: _keys, ...rest } = patch;
    Object.assign(savedSettings, rest);
    const provider = typeof rest.provider === 'string' ? rest.provider : '';
    if (provider && typeof apiKey === 'string') {
      if (apiKey.trim()) savedKeys[provider] = apiKey.trim();
      else delete savedKeys[provider];
    }
    return { success: true };
  });

  /* ---------- Секреты ---------- */
  handle('secrets:describe', () => ({
    found: false,
    // Путь наружу не отдаётся: интерфейс показывает только имя файла.
    fileName: '',
    baseUrl: '',
    textModel: '',
    imageModel: '',
    hasKey: false,
    hasOtherSecrets: false,
    error: '',
  }));

  handle('secrets:importKey', () => ({ success: false, error: 'Недоступно в проверочном сценарии' }));

  /* ---------- Обновления ---------- */
  // В проверках сети нет: интерфейс должен получить понятный ответ без запроса.
  handle('fs:pathExists', (_event, targetPath: string) => {
    try {
      return fs.existsSync(String(targetPath));
    } catch {
      return false;
    }
  });
  handle('update:download', () => ({ success: false, error: 'Недоступно в проверочном сценарии' }));
  handle('update:install', () => ({ success: false, error: 'Недоступно в проверочном сценарии' }));
}
