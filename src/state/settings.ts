import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  AISettings,
  AIProvider,
  ProviderPreset,
  ChatMessage,
  ChatSession,
} from '../types';
import { uid } from '../lib/utils';

/* ============================================================
   Пресеты AI-провайдеров
   ============================================================ */

export const PROVIDER_PRESETS: Record<AIProvider, ProviderPreset> = {
  openai: {
    id: 'openai',
    label: 'OpenAI',
    baseUrl: 'https://api.openai.com/v1',
    models: ['gpt-4o', 'gpt-4o-mini', 'gpt-4.1', 'gpt-4.1-mini', 'o3-mini', 'o1'],
    keyHint: 'sk-...',
    docsUrl: 'https://platform.openai.com/api-keys',
    requiresKey: true,
  },
  anthropic: {
    id: 'anthropic',
    label: 'Anthropic Claude',
    baseUrl: 'https://api.anthropic.com/v1',
    models: [
      'claude-sonnet-4-20250514',
      'claude-3-7-sonnet-20250219',
      'claude-3-5-sonnet-20241022',
      'claude-3-5-haiku-20241022',
    ],
    keyHint: 'sk-ant-...',
    docsUrl: 'https://console.anthropic.com/settings/keys',
    requiresKey: true,
  },
  gemini: {
    id: 'gemini',
    label: 'Google Gemini',
    baseUrl: 'https://generativelanguage.googleapis.com/v1beta',
    models: ['gemini-2.0-flash', 'gemini-1.5-pro', 'gemini-1.5-flash'],
    keyHint: 'AIza...',
    docsUrl: 'https://aistudio.google.com/app/apikey',
    requiresKey: true,
  },
  deepseek: {
    id: 'deepseek',
    label: 'DeepSeek',
    baseUrl: 'https://api.deepseek.com/v1',
    models: ['deepseek-chat', 'deepseek-reasoner'],
    keyHint: 'sk-...',
    docsUrl: 'https://platform.deepseek.com/api_keys',
    requiresKey: true,
  },
  openrouter: {
    id: 'openrouter',
    label: 'OpenRouter (100+ моделей)',
    baseUrl: 'https://openrouter.ai/api/v1',
    models: [
      'anthropic/claude-3.5-sonnet',
      'openai/gpt-4o',
      'google/gemini-2.0-flash-001',
      'deepseek/deepseek-chat',
      'meta-llama/llama-3.3-70b-instruct',
    ],
    keyHint: 'sk-or-...',
    docsUrl: 'https://openrouter.ai/keys',
    requiresKey: true,
  },
  opencode: {
    id: 'opencode',
    label: 'OpenCode (Go)',
    baseUrl: 'https://opencode.ai/inference/go/openai/v1',
    models: [
      'deepseek-v4.1-flash',
      'grok-4.7',
      'glm-5.3',
      'claude-haiku-5-5',
      'gpt-5.6-luna',
      'kimi-k3',
      'qwen3.8-flash',
      'gemini-4-argon',
    ],
    keyHint: 'oc_sk_...',
    docsUrl: 'https://opencode.ai/docs/console/models/',
    requiresKey: true,
  },
  ollama: {
    id: 'ollama',
    label: 'Ollama (локально)',
    baseUrl: 'http://localhost:11434/v1',
    models: ['llama3.2', 'qwen2.5-coder', 'mistral', 'phi4'],
    keyHint: 'не требуется',
    docsUrl: 'https://ollama.com/download',
    requiresKey: false,
  },
  custom: {
    id: 'custom',
    label: 'Свой OpenAI-совместимый endpoint',
    baseUrl: 'http://localhost:1234/v1',
    models: [],
    keyHint: 'ключ или пусто',
    requiresKey: false,
  },
  starimg: {
    id: 'starimg',
    label: 'Starimg AI (шлюз)',
    baseUrl: 'https://ai.starimg.ru/v1',
    models: ['deepseek-v4.1-flash', 'deepseek-v4.1', 'grok-imagine-image'],
    keyHint: 'sk-star-...',
    requiresKey: true,
  },
};

export const PROVIDER_ORDER: AIProvider[] = [
  'openai',
  'anthropic',
  'gemini',
  'deepseek',
  'openrouter',
  'opencode',
  'ollama',
  'starimg',
  'custom',
];

export const DEFAULT_AI_SETTINGS: AISettings = {
  provider: 'openai',
  apiKey: '',
  baseUrl: PROVIDER_PRESETS.openai.baseUrl,
  model: 'gpt-4o-mini',
  temperature: 0.7,
  maxTokens: 4096,
  systemPrompt:
    'Ты — профессиональный ассистент и оператор рабочего пространства WorkSpace. Ты можешь отвечать на вопросы, писать код, анализировать логи и формулы Excel, а также НАПРЯМУЮ УПРАВЛЯТЬ приложением (создавать таблицы и файлы через блок workspace-action). Отвечай структурированно и по существу.',
  activePresetId: 'general',
};

const PRESET_URLS = new Set(
  Object.values(PROVIDER_PRESETS).map((preset) => preset.baseUrl.replace(/\/+$/, ''))
);

/**
 * Приводит настройки в согласованный вид: адрес и модель должны принадлежать
 * выбранному провайдеру.
 *
 * Классическая поломка «перестало подключаться»: провайдер Starimg, а baseUrl
 * остался от OpenAI (403/401 на каждом запросе). Чужой адрес из списка
 * пресетов заменяется адресом провайдера, вместе с ним — модель по умолчанию.
 * Произвольные адреса (свой endpoint) не трогаем.
 */
export function normalizeSettings(input: AISettings): AISettings {
  // Свой endpoint и локальная Ollama — адрес задаёт пользователь.
  if (input.provider === 'custom' || input.provider === 'ollama') return input;
  const preset = PROVIDER_PRESETS[input.provider];
  if (!preset) return input;

  const presetUrl = preset.baseUrl.replace(/\/+$/, '');
  const currentUrl = (input.baseUrl || '').trim().replace(/\/+$/, '');
  const foreignPreset = Boolean(currentUrl) && currentUrl !== presetUrl && PRESET_URLS.has(currentUrl);

  if (!currentUrl) return { ...input, baseUrl: preset.baseUrl };
  if (!foreignPreset) return input;

  return {
    ...input,
    baseUrl: preset.baseUrl,
    model: preset.models[0] || input.model,
  };
}

/* ============================================================
   Системные промты (пресеты ролей)
   ============================================================ */

export interface SystemPreset {
  id: string;
  name: string;
  description: string;
  prompt: string;
}

export const SYSTEM_PRESETS: SystemPreset[] = [
  {
    id: 'general',
    name: 'Универсальный ассистент',
    description: 'Общие вопросы, тексты, анализ и управление файлами',
    prompt:
      'Ты — универсальный интеллектуальный помощник и оператор WorkSpace. Отвечай структурированно и информативно на русском языке. Когда просят создать файл или таблицу, обязательно используй блок workspace-action для автоматического создания в рабочей области.',
  },
  {
    id: 'excel',
    name: 'Мастер формул и таблиц Excel',
    description: 'Создание таблиц, формулы, анализ данных',
    prompt:
      'Ты — эксперт по электронным таблицам Excel и оператор WorkSpace. Составляй формулы и структуры таблиц. Когда пользователь просит создать таблицу со случайными или реальными данными, обязательно генерируй блок workspace-action для создания файла таблицы (.xlsx или .csv) в проекте.',
  },
  {
    id: 'logs',
    name: 'Аналитик логов',
    description: 'Поиск причин сбоев, ERROR/FATAL',
    prompt:
      'Ты — старший системный инженер и аналитик логов. Анализируй фрагменты логов и стектрейсы, находи причины сбоев (ERROR, FATAL, Exception, timeout), выявляй аномалии и предлагай пошаговое исправление.',
  },
  {
    id: 'copywriter',
    name: 'Редактор и копирайтер',
    description: 'Стиль, грамматика, деловая переписка',
    prompt:
      'Ты — профессиональный редактор деловых и технических документов. Улучшай читаемость, исправляй грамматические и пунктуационные ошибки, делай текст убедительным, лаконичным и грамотным, сохраняя исходный смысл.',
  },
  {
    id: 'coder',
    name: 'Программист',
    description: 'Код, скрипты, ревью',
    prompt:
      'Ты — senior software engineer. Пиши надёжный, чистый и идиоматичный код (TypeScript, Python, PowerShell, Bash, SQL). Добавляй необходимые комментарии, указывай обработку ошибок и объясняй ключевые решения.',
  },
];

const SETTINGS_KEY = 'workspace_ai_settings';
const SESSIONS_KEY = 'workspace_ai_sessions';

/* ============================================================
   Хранилище ключей
   Ключи хранятся локально в зашифрованном виде (Web Crypto AES-GCM).
   Открытый ключ никогда не покидает машину пользователя.
   ============================================================ */

const KEY_STORE = 'workspace_api_keys';
const CRYPTO_KEY_NAME = 'workspace_key_encryption_v1';

async function getCryptoKey(): Promise<CryptoKey | null> {
  try {
    if (!window.crypto?.subtle) return null;
    const raw = window.localStorage.getItem(CRYPTO_KEY_NAME);
    if (raw) {
      const bytes = Uint8Array.from(atob(raw), (c) => c.charCodeAt(0));
      return await window.crypto.subtle.importKey('raw', bytes, 'AES-GCM', true, [
        'encrypt',
        'decrypt',
      ]);
    }
    const key = await window.crypto.subtle.generateKey({ name: 'AES-GCM', length: 256 }, true, [
      'encrypt',
      'decrypt',
    ]);
    const exported = await window.crypto.subtle.exportKey('raw', key);
    window.localStorage.setItem(
      CRYPTO_KEY_NAME,
      btoa(String.fromCharCode(...new Uint8Array(exported)))
    );
    return key;
  } catch {
    return null;
  }
}

/** Шифрование строки. При недоступности Web Crypto возвращает plain-строку. */
async function encryptString(plain: string): Promise<string> {
  const key = await getCryptoKey();
  if (!key) return `plain:${plain}`;
  try {
    const iv = window.crypto.getRandomValues(new Uint8Array(12));
    const encoded = new TextEncoder().encode(plain);
    const cipher = await window.crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, encoded);
    const ivB64 = btoa(String.fromCharCode(...iv));
    const dataB64 = btoa(String.fromCharCode(...new Uint8Array(cipher)));
    return `enc:${ivB64}:${dataB64}`;
  } catch {
    return `plain:${plain}`;
  }
}

/** Расшифровка строки, созданной encryptString. */
async function decryptString(stored: string): Promise<string> {
  if (stored.startsWith('plain:')) return stored.slice(6);
  if (!stored.startsWith('enc:')) return stored;
  const key = await getCryptoKey();
  if (!key) return '';
  try {
    const [, ivB64, dataB64] = stored.split(':');
    const iv = Uint8Array.from(atob(ivB64), (c) => c.charCodeAt(0));
    const data = Uint8Array.from(atob(dataB64), (c) => c.charCodeAt(0));
    const plain = await window.crypto.subtle.decrypt({ name: 'AES-GCM', iv }, key, data);
    return new TextDecoder().decode(plain);
  } catch {
    return '';
  }
}

/** Менеджер локального хранилища API-ключей по провайдерам. */
export const keyVault = {
  async get(provider: AIProvider): Promise<string> {
    try {
      const raw = window.localStorage.getItem(KEY_STORE);
      if (!raw) return '';
      const map = JSON.parse(raw) as Record<string, string>;
      const stored = map[provider];
      if (!stored) return '';
      return await decryptString(stored);
    } catch {
      return '';
    }
  },

  async set(provider: AIProvider, apiKey: string): Promise<void> {
    try {
      const raw = window.localStorage.getItem(KEY_STORE);
      const map = raw ? (JSON.parse(raw) as Record<string, string>) : {};
      if (apiKey) {
        map[provider] = await encryptString(apiKey);
      } else {
        delete map[provider];
      }
      window.localStorage.setItem(KEY_STORE, JSON.stringify(map));
    } catch {
      // Игнорируем ошибки записи
    }
  },

  async listConfigured(): Promise<AIProvider[]> {
    try {
      const raw = window.localStorage.getItem(KEY_STORE);
      if (!raw) return [];
      const map = JSON.parse(raw) as Record<string, string>;
      return Object.keys(map).filter((k) => Boolean(map[k])) as AIProvider[];
    } catch {
      return [];
    }
  },
};

/* ============================================================
   Хук настроек
   ============================================================ */

function loadSettings(): AISettings {
  try {
    const raw = window.localStorage.getItem(SETTINGS_KEY);
    if (raw) {
      return normalizeSettings({ ...DEFAULT_AI_SETTINGS, ...(JSON.parse(raw) as Partial<AISettings>) });
    }
  } catch {
    // повреждённые настройки — используем значения по умолчанию
  }
  return DEFAULT_AI_SETTINGS;
}

/** Постоянное хранилище доступно только в собранном приложении. */
function persistentStore() {
  return window.electronAPI?.loadAiSettings && window.electronAPI?.saveAiSettings
    ? window.electronAPI
    : null;
}

/** Пишет настройки и, если передан, ключ активного провайдера на диск. */
async function persistPersistent(settings: AISettings, apiKey?: string): Promise<void> {
  const store = persistentStore();
  if (!store?.saveAiSettings) return;
  const { apiKey: _omit, ...safe } = settings;
  await store.saveAiSettings(apiKey === undefined ? safe : { ...safe, apiKey });
}

export function useAISettings() {
  const [settings, setSettingsState] = useState<AISettings>(loadSettings);
  /** Актуальные настройки для колбэков, вызываемых сразу после смены провайдера. */
  const settingsRef = useRef<AISettings>(settings);
  settingsRef.current = settings;
  const [keyLoaded, setKeyLoaded] = useState(false);

  // Ключ активного провайдера подгружается из защищённого хранилища
  const [providerKey, setProviderKey] = useState('');
  /** Сообщение о нечитаемом ключе — показывается один раз при старте. */
  const [keyProblem, setKeyProblem] = useState('');

  const persist = useCallback((next: AISettings) => {
    setSettingsState(next);
    try {
      const { apiKey: _omit, ...safe } = next;
      window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(safe));
    } catch {
      // игнорируем
    }
    persistPersistent(next);
  }, []);

  // Загрузка ключа при смене провайдера
  const loadKeyFor = useCallback(async (provider: AIProvider) => {
    const key = await keyVault.get(provider);
    setProviderKey(key);
    setSettingsState((prev) => ({ ...prev, apiKey: key }));
    setKeyLoaded(true);
  }, []);

  const updateSettings = useCallback(
    (patch: Partial<AISettings>) => {
      setSettingsState((prev) => {
        const next = { ...prev, ...patch };
        try {
          const { apiKey: _omit, ...safe } = next;
          window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(safe));
        } catch {
          // игнорируем
        }
        persistPersistent(next);
        return next;
      });
    },
    []
  );

  // Кэш моделей, загруженных от API провайдеров
  const [modelsCache, setModelsCache] = useState<Record<string, string[]>>(() => {
    try {
      const raw = window.localStorage.getItem('workspace_ai_models_cache');
      if (raw) return JSON.parse(raw);
    } catch {}
    return {};
  });

  const fetchModels = useCallback(
    async (override?: Partial<AISettings>) => {
      const target = { ...settings, ...override };
      try {
        const { aiService } = await import('../services/aiService');
        const res = await aiService.fetchModels(target);
        if (res.success && res.models && res.models.length > 0) {
          setModelsCache((prev) => {
            const next = { ...prev, [target.provider]: res.models };
            try {
              window.localStorage.setItem('workspace_ai_models_cache', JSON.stringify(next));
            } catch {}
            persistPersistent({ ...target, modelsCache: next });
            return next;
          });

          // Если текущей модели нет в новом списке, автоматически выбираем первую модель
          if (
            !target.model ||
            (!res.models.includes(target.model) &&
              !PROVIDER_PRESETS[target.provider]?.models.includes(target.model))
          ) {
            setSettingsState((prev) => {
              const updated = { ...prev, model: res.models[0] };
              try {
                const { apiKey: _omit, ...safe } = updated;
                window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(safe));
              } catch {}
              persistPersistent(updated);
              return updated;
            });
          }
        }
        return res;
      } catch (err) {
        return { success: false, models: [], error: (err as Error).message };
      }
    },
    [settings]
  );

  const availableModels = useMemo(() => {
    const fromApi = modelsCache[settings.provider] || [];
    const fromPreset = PROVIDER_PRESETS[settings.provider]?.models || [];
    const combined = Array.from(new Set([...fromApi, ...fromPreset]));
    if (settings.model && !combined.includes(settings.model)) {
      combined.push(settings.model);
    }
    return combined;
  }, [modelsCache, settings.provider, settings.model]);

  /** Смена провайдера с подстановкой базового URL и модели по умолчанию. */
  const changeProvider = useCallback(
    async (provider: AIProvider) => {
      const preset = PROVIDER_PRESETS[provider];
      const store = persistentStore();
      const key = store?.loadAiSettings
        ? (await store.loadAiSettings()).keys?.[provider] || (await keyVault.get(provider))
        : await keyVault.get(provider);
      setProviderKey(key);
      setSettingsState((prev) => {
        const cached = modelsCache[provider] || [];
        const next: AISettings = {
          ...prev,
          provider,
          apiKey: key,
          baseUrl: preset.baseUrl,
          model: cached[0] || preset.models[0] || prev.model,
        };
        try {
          const { apiKey: _omit, ...safe } = next;
          window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(safe));
        } catch {
          // игнорируем
        }
        persistPersistent(next);
        return next;
      });
      setKeyLoaded(true);
    },
    [modelsCache]
  );

  /** Сохранение ключа: в профиль пользователя и в локальный кэш. */
  const saveApiKey = useCallback(
    async (apiKey: string, patch?: Partial<AISettings>) => {
      setKeyProblem('');
      // Патч нужен при импорте из файла секретов: провайдер и адрес меняются
      // в том же обработчике, до перерисовки состояния.
      const snapshot = normalizeSettings({ ...settingsRef.current, ...patch, apiKey });
      void keyVault.set(snapshot.provider, apiKey);
      setProviderKey(apiKey);
      setSettingsState(snapshot);
      await persistPersistent(snapshot, apiKey);
      // При сохранении ключа сразу автоматически опрашиваем API на список моделей
      if (apiKey.trim()) {
        void fetchModels({ ...snapshot, apiKey });
      }
    },
    [fetchModels]
  );

  // Первичная загрузка: постоянный файл важнее localStorage,
  // потому что localStorage обнуляется при пересборке приложения.
  useEffect(() => {
    const store = persistentStore();
    if (!store?.loadAiSettings) {
      void loadKeyFor(settings.provider);
      return;
    }
    let cancelled = false;
    void store.loadAiSettings().then(async (stored) => {
      if (cancelled) return;
      const saved = stored.settings ?? {};
      if (saved.modelsCache) {
        setModelsCache((prev) => ({ ...prev, ...saved.modelsCache }));
      }
      const provider = (saved.provider as AIProvider) || settings.provider;
      let key = stored.keys?.[provider] ?? '';
      let problem = '';

      // Ключ хранится в зашифрованном виде и привязан к профилю Windows.
      // После переноса portable-папки на другой компьютер он не читается —
      // восстанавливаем ключ из secrets.local.ps1, если файл есть рядом.
      const unreadable = (stored.keyErrors ?? []).includes(provider);
      if (!key && unreadable && window.electronAPI?.importSecretsKey) {
        const imported = await window.electronAPI.importSecretsKey().catch(() => null);
        if (imported?.success && imported.apiKey) {
          key = imported.apiKey;
          const restored = normalizeSettings({
            ...DEFAULT_AI_SETTINGS,
            ...saved,
            provider: imported.baseUrl?.includes('starimg') ? 'starimg' : provider,
            baseUrl: imported.baseUrl || saved.baseUrl || PROVIDER_PRESETS[provider]?.baseUrl || DEFAULT_AI_SETTINGS.baseUrl,
            model: imported.model || saved.model || PROVIDER_PRESETS[provider]?.models[0] || DEFAULT_AI_SETTINGS.model,
            apiKey: key,
          });
          await persistPersistent(restored, key);
          if (!cancelled) {
            setProviderKey(key);
            setSettingsState((prev) => ({ ...prev, ...restored }));
            setKeyLoaded(true);
          }
          return;
        }
        problem =
          'Ключ API сохранён на другом компьютере и не читается здесь. ' +
          'Откройте Настройки → AI и вставьте ключ заново.';
      }

      // Ключ из старого хранилища переносим в постоянное один раз.
      if (!key) {
        key = await keyVault.get(provider);
        if (key) persistPersistent({ ...DEFAULT_AI_SETTINGS, ...saved, provider, apiKey: key }, key);
      }
      if (cancelled) return;
      if (problem) setKeyProblem(problem);
      setProviderKey(key);
      setSettingsState((prev) => normalizeSettings({ ...prev, ...saved, provider, apiKey: key }));
      setKeyLoaded(true);
    }).catch(() => {
      if (!cancelled) void loadKeyFor(settings.provider);
    });
    return () => {
      cancelled = true;
    };
    // Загрузка нужна один раз при запуске.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const preset = useMemo(() => PROVIDER_PRESETS[settings.provider], [settings.provider]);

  const isConfigured = useMemo(
    () => !preset.requiresKey || Boolean(providerKey),
    [preset.requiresKey, providerKey]
  );

  return {
    settings,
    preset,
    availableModels,
    modelsCache,
    fetchModels,
    updateSettings,
    changeProvider,
    saveApiKey,
    persist,
    isConfigured,
    hasStoredKey: Boolean(providerKey),
    keyProblem,
  };
}

/* ============================================================
   История чатов
   ============================================================ */

function loadSessions(): ChatSession[] {
  try {
    const raw = window.localStorage.getItem(SESSIONS_KEY);
    if (raw) return JSON.parse(raw) as ChatSession[];
  } catch {
    // игнорируем
  }
  return [];
}

export function useChatSessions() {
  const [sessions, setSessions] = useState<ChatSession[]>(loadSessions);
  const [activeId, setActiveId] = useState<string>('');

  const persistSessions = useCallback((next: ChatSession[]) => {
    setSessions(next);
    try {
      // Ограничиваем историю, чтобы не переполнить localStorage
      window.localStorage.setItem(SESSIONS_KEY, JSON.stringify(next.slice(0, 40)));
    } catch {
      // игнорируем
    }
  }, []);

  const createSession = useCallback(
    (title = 'Новый диалог'): string => {
      const session: ChatSession = {
        id: uid('chat'),
        title,
        messages: [],
        createdAt: Date.now(),
        updatedAt: Date.now(),
      };
      setSessions((prev) => {
        const next = [session, ...prev];
        try {
          window.localStorage.setItem(SESSIONS_KEY, JSON.stringify(next.slice(0, 40)));
        } catch {
          // игнорируем
        }
        return next;
      });
      setActiveId(session.id);
      return session.id;
    },
    []
  );

  const updateSessionMessages = useCallback(
    (sessionId: string, messages: ChatMessage[]) => {
      setSessions((prev) => {
        const next = prev.map((s) =>
          s.id === sessionId
            ? {
                ...s,
                messages,
                // Заголовок берём из первого вопроса пользователя
                title:
                  s.messages.length === 0 && messages[0]?.role === 'user'
                    ? messages[0].content.slice(0, 42)
                    : s.title,
                updatedAt: Date.now(),
              }
            : s
        );
        try {
          window.localStorage.setItem(SESSIONS_KEY, JSON.stringify(next.slice(0, 40)));
        } catch {
          // игнорируем
        }
        return next;
      });
    },
    []
  );

  const deleteSession = useCallback(
    (sessionId: string) => {
      setSessions((prev) => {
        const next = prev.filter((s) => s.id !== sessionId);
        try {
          window.localStorage.setItem(SESSIONS_KEY, JSON.stringify(next));
        } catch {
          // игнорируем
        }
        return next;
      });
      setActiveId((prev) => (prev === sessionId ? '' : prev));
    },
    []
  );

  const clearAll = useCallback(() => {
    persistSessions([]);
    setActiveId('');
  }, [persistSessions]);

  return {
    sessions,
    activeId,
    setActiveId,
    createSession,
    updateSessionMessages,
    deleteSession,
    clearAll,
  };
}
