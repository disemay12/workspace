import { AISettings, ChatMessage } from '../types';
import { PROVIDER_PRESETS } from '../state/settings';
import { uid } from '../lib/utils';

/** Максимальный размер прикрепляемого к запросу файла (символов). */
export const MAX_ATTACHMENT_CHARS = 40000;

export interface StreamCallbacks {
  onChunk: (partialText: string) => void;
  onFinish: (fullText: string) => void;
  onError: (error: Error) => void;
}

const isElectron = () => typeof window !== 'undefined' && Boolean(window.electronAPI);

/** Понятное описание сетевой ошибки либо ошибки API. */
function describeError(err: unknown, providerLabel: string): Error {
  const message = err instanceof Error ? err.message : String(err);

  if (/Failed to fetch|NetworkError|net::ERR_/i.test(message)) {
    return new Error(
      `Не удалось подключиться к ${providerLabel}. Проверьте интернет-соединение и базовый URL.`
    );
  }
  if (/ENOTFOUND|EAI_AGAIN|ECONNREFUSED/i.test(message)) {
    return new Error(
      `Сервер ${providerLabel} недоступен. Проверьте базовый URL и доступность сервиса.`
    );
  }
  return err instanceof Error ? err : new Error(message);
}

/** Формирует текст сообщения с учётом прикреплённого файла. */
function buildMessageText(message: ChatMessage): string {
  if (!message.attachedFile) return message.content;
  const { name, path, content, truncated } = message.attachedFile;
  const suffix = truncated ? '\n… (фрагмент обрезан)' : '';
  return (
    `[Прикреплённый файл: ${name}]\n` +
    `Путь: ${path}\n` +
    '```\n' +
    `${content.slice(0, MAX_ATTACHMENT_CHARS)}${suffix}\n` +
    '```\n\n' +
    message.content
  );
}

/**
 * Разбор SSE-потока. Универсальный парсер: каждая строка `data: {...}`
 * передаётся в extractDelta, которая возвращает приращение текста.
 * Возвращает накопленный текст.
 */
function createSSEParser(extractDelta: (payload: any) => string, onChunk: (partial: string) => void) {
  let buffer = '';
  let full = '';

  return {
    /** Обрабатывает очередной фрагмент потока. */
    push(text: string) {
      buffer += text;
      const lines = buffer.split('\n');
      // Последняя строка может быть неполной — оставляем в буфере
      buffer = lines.pop() ?? '';

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith('data:')) continue;
        const data = trimmed.slice(5).trim();
        if (!data || data === '[DONE]') continue;
        try {
          const payload = JSON.parse(data);
          const delta = extractDelta(payload);
          if (delta) {
            full += delta;
            onChunk(full);
          }
        } catch {
          // Неполный JSON-чанк — пропускаем
        }
      }
    },
    get result() {
      return full;
    },
  };
}

/** Извлекает понятное сообщение об ошибке из тела ответа API. */
function parseApiError(status: number, body: string, label: string): Error {
  let detail = body;
  try {
    const parsed = JSON.parse(body);
    detail = parsed?.error?.message || parsed?.message || body;
  } catch {
    // тело не JSON — используем как есть
  }

  const hint =
    status === 401
      ? ' Проверьте правильность API-ключа.'
      : status === 403
        ? ' Доступ запрещён — проверьте права ключа.'
        : status === 404
          ? ' Проверьте название модели и базовый URL.'
          : status === 429
            ? ' Превышен лимит запросов, попробуйте позже.'
            : status >= 500
              ? ' Ошибка на стороне провайдера, попробуйте позже.'
              : '';

  return new Error(`${label} (${status}): ${detail.slice(0, 400)}${hint}`);
}
export const WORKSPACE_OPERATOR_SYSTEM_PROMPT = `Ты являешься интеллектуальным помощником и оператором приложения WorkSpace.
Ты можешь не только отвечать текстом, но и НАПРЯМУЮ УПРАВЛЯТЬ приложением (создавать файлы, таблицы, открывать документы).

Когда пользователь просит создать таблицу, создать файл, открыть документ или выполнить действие в программе, обязательно дай краткий ответ И добавь специальный блок действия в конце ответа:

1. Создание таблицы (Excel или CSV):
\`\`\`workspace-action
{
  "action": "create_table",
  "name": "Название.xlsx",
  "sheetName": "Лист 1",
  "columns": ["Колонка 1", "Колонка 2", "Сумма"],
  "rows": [
    ["Значение A", 100, "=B2*1.2"],
    ["Значение B", 250, "=B3*1.2"]
  ]
}
\`\`\`
(Если просили CSV: укажи "name": "данные.csv" и "format": "csv"). Для формул используй имена функций Excel (SUM, AVERAGE, IF) и формулы вида "=C2*D2".

2. Создание файла с текстом или кодом:
\`\`\`workspace-action
{
  "action": "create_file",
  "name": "script.py",
  "content": "print('Hello world')"
}
\`\`\`

3. Открытие существующего файла:
\`\`\`workspace-action
{
  "action": "open_file",
  "path": "имя_файла.xlsx"
}
\`\`\`

4. Вставка в активный документ:
\`\`\`workspace-action
{
  "action": "insert_into_editor",
  "text": "текст"
}
\`\`\`

Всегда добавляй блок \`\`\`workspace-action, когда тебя просят создать таблицу или файл, чтобы программа автоматически создала его и открыла в редакторе!`;

export function resolveSystemPrompt(userPrompt?: string): string {
  const base = (userPrompt || '').trim();
  if (base.includes('workspace-action')) return base;
  return base ? `${base}\n\n${WORKSPACE_OPERATOR_SYSTEM_PROMPT}` : WORKSPACE_OPERATOR_SYSTEM_PROMPT;
}

export const aiService = {
  /** Проверка соединения и валидности ключа — для кнопки «Проверить подключение». */
  async testConnection(settings: AISettings): Promise<{ ok: boolean; message: string }> {
    const preset = PROVIDER_PRESETS[settings.provider];
    if (preset.requiresKey && !settings.apiKey.trim()) {
      return { ok: false, message: 'API-ключ не указан' };
    }

    try {
      let received = '';
      const probe: ChatMessage[] = [
        { id: 'probe', role: 'user', content: 'Ответь одним словом: готов', timestamp: Date.now() },
      ];
      await this.streamChat(probe, settings, {
        onChunk: (t) => {
          received = t;
        },
        onFinish: () => undefined,
        onError: (e) => {
          throw e;
        },
      });
      return {
        ok: true,
        message: `Подключение успешно. Модель ответила: «${received.trim().slice(0, 40) || 'ок'}»`,
      };
    } catch (err) {
      return { ok: false, message: (err as Error).message };
    }
  },

  /**
   * Запрос списка доступных моделей от API провайдера.
   * В среде Electron запрос выполняется в главном процессе через IPC (обход CORS).
   */
  async fetchModels(
    settings: AISettings
  ): Promise<{ success: boolean; models: string[]; error?: string }> {
    const provider = settings.provider;
    const baseUrl = settings.baseUrl || PROVIDER_PRESETS[provider]?.baseUrl;
    const apiKey = settings.apiKey;

    if (isElectron() && window.electronAPI?.aiListModels) {
      return await window.electronAPI.aiListModels({
        provider,
        baseUrl,
        apiKey,
      });
    }

    // Браузерный fallback
    try {
      if (provider === 'gemini') {
        const base = (baseUrl || 'https://generativelanguage.googleapis.com/v1beta').replace(/\/+$/, '');
        const keyParam = apiKey ? `?key=${encodeURIComponent(apiKey.trim())}` : '';
        const res = await fetch(`${base}/models${keyParam}`);
        if (!res.ok) throw new Error(`HTTP ${res.status}: ${await res.text()}`);
        const data = await res.json();
        const rawModels: string[] = (data.models || [])
          .filter((m: any) => !m.supportedGenerationMethods || m.supportedGenerationMethods.includes('generateContent'))
          .map((m: any) => String(m.name || '').replace(/^models\//, ''))
          .filter((m: string) => Boolean(m));
        const unique = Array.from(new Set(rawModels)).sort((a, b) => a.localeCompare(b));
        return { success: true, models: unique };
      }

      if (provider === 'anthropic') {
        const base = (baseUrl || 'https://api.anthropic.com/v1').replace(/\/+$/, '');
        const res = await fetch(`${base}/models`, {
          headers: {
            'x-api-key': apiKey.trim(),
            'anthropic-version': '2023-06-01',
            'anthropic-dangerous-direct-browser-access': 'true',
          },
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}: ${await res.text()}`);
        const data = await res.json();
        const rawModels: string[] = (data.data || []).map((m: any) => String(m.id || '')).filter(Boolean);
        const unique = Array.from(new Set(rawModels)).sort((a, b) => a.localeCompare(b));
        return { success: true, models: unique };
      }

      // OpenAI и совместимые
      const cleanBase = (baseUrl || 'https://api.openai.com/v1')
        .replace(/\/+$/, '')
        .replace(/\/chat\/completions$/, '');
      const headers: Record<string, string> = { Accept: 'application/json' };
      if (apiKey.trim()) headers['Authorization'] = `Bearer ${apiKey.trim()}`;
      if (provider === 'openrouter') {
        headers['HTTP-Referer'] = 'https://workspace.local';
        headers['X-Title'] = 'WorkSpace';
      }
      const res = await fetch(`${cleanBase}/models`, { headers });
      if (!res.ok) throw new Error(`HTTP ${res.status}: ${await res.text()}`);
      const data = await res.json();
      let raw: any[] = [];
      if (Array.isArray(data?.data)) raw = data.data;
      else if (Array.isArray(data?.models)) raw = data.models;
      else if (Array.isArray(data)) raw = data;

      const rawModels: string[] = raw.map((m: any) => String(typeof m === 'string' ? m : m.id || m.name || '')).filter(Boolean);
      const unique = Array.from(new Set(rawModels)).sort((a, b) => a.localeCompare(b));
      return { success: true, models: unique };
    } catch (err) {
      return { success: false, models: [], error: (err as Error).message };
    }
  },

  resolveEndpoint(settings: AISettings): string {
    const base = (settings.baseUrl || PROVIDER_PRESETS[settings.provider].baseUrl)
      .trim()
      .replace(/\/+$/, '');
    if (/\/chat\/completions$/.test(base)) return base;
    return `${base}/chat/completions`;
  },

  /** Проверка доступности endpoint до отправки запроса. */
  async probeEndpoint(settings: AISettings): Promise<{ reachable: boolean; error?: string }> {
    if (!isElectron() || !window.electronAPI?.aiProbe) return { reachable: true };
    const preset = PROVIDER_PRESETS[settings.provider];
    const base = (settings.baseUrl || preset.baseUrl).trim().replace(/\/+$/, '');
    const res = await window.electronAPI.aiProbe(base);
    return { reachable: res.reachable, error: res.error };
  },

  /** Основная точка входа: маршрутизация по провайдеру со стримингом. */
  async streamChat(
    messages: ChatMessage[],
    settings: AISettings,
    callbacks: StreamCallbacks,
    signal?: AbortSignal
  ): Promise<void> {
    const preset = PROVIDER_PRESETS[settings.provider];
    if (preset.requiresKey && !settings.apiKey.trim()) {
      callbacks.onError(
        new Error(
          `API-ключ для ${preset.label} не указан. Откройте Настройки → AI и введите ваш персональный ключ.`
        )
      );
      return;
    }

    // Пустые сообщения и системные роли не отправляем
    const clean = messages.filter((m) => m.role !== 'system' && m.content.trim().length > 0);
    if (clean.length === 0) {
      callbacks.onError(new Error('Пустой запрос'));
      return;
    }

    try {
      if (settings.provider === 'anthropic') {
        await this.streamAnthropic(clean, settings, callbacks, signal);
      } else if (settings.provider === 'gemini') {
        await this.streamGemini(clean, settings, callbacks, signal);
      } else {
        await this.streamOpenAICompatible(clean, settings, callbacks, signal);
      }
    } catch (err) {
      if ((err as Error)?.name === 'AbortError') return;
      callbacks.onError(describeError(err, preset.label));
    }
  },

  /**
   * Выполняет POST-запрос со стримингом.
   *
   * В Electron запрос идёт через главный процесс (`net`-модуль), потому что
   * браузерный fetch блокируется политикой CORS: часть OpenAI-совместимых
   * шлюзов не отдаёт `Access-Control-Allow-Origin`, и Chromium прерывает
   * ответ с `net::ERR_FAILED`.
   *
   * В браузере (режим отладки без Electron) используется обычный fetch.
   */
  async requestStream(
    endpoint: string,
    headers: Record<string, string>,
    body: unknown,
    extractDelta: (payload: any) => string,
    onChunk: (partial: string) => void,
    signal?: AbortSignal,
    label = 'Ошибка API'
  ): Promise<string> {
    const bodyText = JSON.stringify(body);

    /* ---------- Путь через Electron ---------- */
    if (isElectron() && window.electronAPI?.aiStream) {
      const requestId = uid('ai');

      return new Promise<string>((resolve, reject) => {
        const parser = createSSEParser(extractDelta, onChunk);
        let settled = false;

        const cleanup = () => {
          offChunk();
          offEnd();
          offError();
        };

        const finish = (fn: () => void) => {
          if (settled) return;
          settled = true;
          cleanup();
          fn();
        };

        const offChunk = window.electronAPI!.onAiChunk((data) => {
          if (data.requestId !== requestId) return;
          parser.push(data.chunk);
        });

        const offEnd = window.electronAPI!.onAiEnd((data) => {
          if (data.requestId !== requestId) return;
          finish(() => resolve(parser.result));
        });

        const offError = window.electronAPI!.onAiError((data) => {
          if (data.requestId !== requestId) return;
          finish(() => {
            if (data.status) {
              reject(parseApiError(data.status, data.message, label));
            } else {
              reject(new Error(data.message || 'Неизвестная ошибка сети'));
            }
          });
        });

        // Отмена по кнопке «Стоп»
        signal?.addEventListener(
          'abort',
          () => {
            void window.electronAPI!.aiAbort(requestId);
            const abortError = new Error('Aborted');
            abortError.name = 'AbortError';
            finish(() => reject(abortError));
          },
          { once: true }
        );

        void window.electronAPI!
          .aiStream({ requestId, url: endpoint, method: 'POST', headers, body: bodyText })
          .then((res) => {
            if (!res.started) {
              finish(() => reject(new Error('Не удалось запустить запрос')));
            }
          })
          .catch((err) => finish(() => reject(err)));
      });
    }

    /* ---------- Путь через браузерный fetch ---------- */
    const response = await fetch(endpoint, {
      method: 'POST',
      headers,
      body: bodyText,
      signal,
    });

    if (!response.ok) {
      const errText = await response.text();
      throw parseApiError(response.status, errText, label);
    }
    if (!response.body) throw new Error('Сервер вернул пустой ответ');

    const reader = response.body.getReader();
    const decoder = new TextDecoder('utf-8');
    const parser = createSSEParser(extractDelta, onChunk);

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      parser.push(decoder.decode(value, { stream: true }));
    }

    return parser.result;
  },

  async streamOpenAICompatible(
    messages: ChatMessage[],
    settings: AISettings,
    { onChunk, onFinish }: StreamCallbacks,
    signal?: AbortSignal
  ): Promise<void> {
    const endpoint = this.resolveEndpoint(settings);

    const headers: Record<string, string> = { 'Content-Type': 'application/json' };
    if (settings.apiKey.trim()) headers['Authorization'] = `Bearer ${settings.apiKey.trim()}`;
    // OpenRouter требует идентификацию клиента
    if (settings.provider === 'openrouter') {
      headers['HTTP-Referer'] = 'https://workspace.local';
      headers['X-Title'] = 'WorkSpace';
    }

    const full = await this.requestStream(
      endpoint,
      headers,
      {
        model: settings.model || 'gpt-4o-mini',
        messages: [
          { role: 'system', content: resolveSystemPrompt(settings.systemPrompt) },
          ...messages.map((m) => ({ role: m.role, content: buildMessageText(m) })),
        ],
        temperature: settings.temperature,
        max_tokens: settings.maxTokens,
        stream: true,
      },
      (p) => p?.choices?.[0]?.delta?.content ?? p?.choices?.[0]?.message?.content ?? '',
      onChunk,
      signal,
      'Ошибка API'
    );

    onFinish(full);
  },

  async streamAnthropic(
    messages: ChatMessage[],
    settings: AISettings,
    { onChunk, onFinish }: StreamCallbacks,
    signal?: AbortSignal
  ): Promise<void> {
    const full = await this.requestStream(
      'https://api.anthropic.com/v1/messages',
      {
        'Content-Type': 'application/json',
        'x-api-key': settings.apiKey.trim(),
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      {
        model: settings.model || 'claude-3-5-sonnet-20241022',
        system: resolveSystemPrompt(settings.systemPrompt),
        messages: messages.map((m) => ({
          role: m.role === 'assistant' ? 'assistant' : 'user',
          content: buildMessageText(m),
        })),
        max_tokens: settings.maxTokens,
        temperature: settings.temperature,
        stream: true,
      },
      (p) => (p?.type === 'content_block_delta' ? (p?.delta?.text ?? '') : ''),
      onChunk,
      signal,
      'Ошибка Claude API'
    );

    onFinish(full);
  },

  async streamGemini(
    messages: ChatMessage[],
    settings: AISettings,
    { onChunk, onFinish }: StreamCallbacks,
    signal?: AbortSignal
  ): Promise<void> {
    const model = settings.model || 'gemini-1.5-flash';
    const endpoint =
      `https://generativelanguage.googleapis.com/v1beta/models/${model}:streamGenerateContent` +
      `?alt=sse&key=${encodeURIComponent(settings.apiKey.trim())}`;

    const full = await this.requestStream(
      endpoint,
      { 'Content-Type': 'application/json' },
      {
        contents: messages.map((m) => ({
          role: m.role === 'assistant' ? 'model' : 'user',
          parts: [{ text: buildMessageText(m) }],
        })),
        systemInstruction: { parts: [{ text: resolveSystemPrompt(settings.systemPrompt) }] },
        generationConfig: {
          temperature: settings.temperature,
          maxOutputTokens: settings.maxTokens,
        },
      },
      (p) => (p?.candidates?.[0]?.content?.parts ?? []).map((x: any) => x?.text ?? '').join(''),
      onChunk,
      signal,
      'Ошибка Gemini API'
    );

    onFinish(full);
  },
};
