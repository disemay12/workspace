import { AISettings, ChatMessage } from '../types';
import { PROVIDER_PRESETS, normalizeSettings } from '../state/settings';
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
export const WORKSPACE_OPERATOR_SYSTEM_PROMPT = `Ты — интеллектуальный помощник и оператор приложения WorkSpace. Ты не только отвечаешь текстом, но и НАПРЯМУЮ УПРАВЛЯЕШЬ приложением через блоки действий.

ПРОТОКОЛ ДЕЙСТВИЙ
Когда пользователь просит что-то сделать в приложении (создать, открыть, изменить, переименовать, удалить, сохранить, переключить панель, выполнить команду), дай краткий ответ и добавь в конце ответа один или несколько блоков:

\`\`\`workspace-action
{ "action": "имя_действия", ... }
\`\`\`

Можно вернуть несколько блоков подряд или один блок с массивом: { "actions": [ {...}, {...} ] }.
Не описывай действие словами как выполненное, если не добавил блок. Не оборачивай JSON в таблицу.

ДОСТУПНЫЕ ДЕЙСТВИЯ

1) create_table — таблица Excel/CSV
{ "action": "create_table", "name": "Отчет.xlsx", "sheetName": "Лист 1", "columns": ["Товар","Цена","Кол-во","Сумма"], "rows": [["Яблоки",120,3,"=B2*C2"]], "format": "xlsx" }
Для CSV: "name": "данные.csv", "format": "csv". Формулы Excel — строками "=SUM(D2:D10)".

2) create_document — Word (.docx), content в Markdown. Без "name" имя берётся из "title".
{ "action": "create_document", "name": "Отчет.docx", "title": "Отчет о продажах", "content": "# Заголовок\\n\\nТекст, **жирный**, списки..." }

3) create_file — любой текстовый файл (код, конфиг, заметка)
{ "action": "create_file", "name": "script.py", "path": "src", "content": "print('Hello')" }
"path" — папка относительно корня рабочей области (необязательно).

4) open_file — открыть существующий файл
{ "action": "open_file", "path": "docs/Отчет.docx" }

5) open_folder — сменить рабочую папку (без path откроется диалог выбора)
{ "action": "open_folder", "path": "C:\\\\Projects\\\\MyProject" }

6) create_folder — новая папка
{ "action": "create_folder", "path": "src/components" }

7) rename_path — переименовать файл или папку
{ "action": "rename_path", "path": "old.txt", "newName": "new.txt" }

8) delete_path — удалить файл или папку (потребуется подтверждение пользователя)
{ "action": "delete_path", "path": "temp/old.txt" }

9) insert_into_editor — вставить текст в активный документ в позицию курсора
{ "action": "insert_into_editor", "text": "текст для вставки" }

10) replace_document — заменить содержимое активного документа
{ "action": "replace_document", "text": "новое содержимое" }

11) open_settings — открыть настройки на вкладке provider|model|prompt|keys|integrations|system
{ "action": "open_settings", "tab": "provider" }

12) toggle_panel — показать/скрыть панель sidebar|ai|terminal|search, state open|close|toggle
{ "action": "toggle_panel", "panel": "terminal", "state": "open" }

13) select_tab — переключиться на вкладку по path, name или index (0 — первая)
{ "action": "select_tab", "name": "Отчет.docx" }

14) close_tab — закрыть вкладку (текущую или указанную)
{ "action": "close_tab" }

15) save — сохранить текущий файл или все изменённые
{ "action": "save", "all": true }

16) notify — показать уведомление
{ "action": "notify", "message": "Готово" }

17) run_command — выполнить команду в встроенном терминале (потребуется подтверждение)
{ "action": "run_command", "command": "npm run build", "cwd": "." }

18) edit_table — изменить существующую таблицу Excel/CSV (заменить или дописать строки)
{ "action": "edit_table", "path": "comka-trash.xlsx", "mode": "append", "columns": ["Поле1","Поле2"], "rows": [["значение1","значение2"]] }
Без "path" изменяется активная таблица. "mode": "replace" (по умолчанию) или "append".

19) import_table — импортировать таблицу из файла Markdown/CSV в xlsx (без вывода строк в ответе)
{ "action": "import_table", "source": "comka-report.md", "target": "comka-trash.xlsx", "sheetName": "Комиссионка" }
Используй это действие, когда просят «перенеси все данные из файла в таблицу». Не переписывай строки вручную: приложение прочитает файл целиком.
Если пользователь просит «перенеси данные в нашу таблицу», используй edit_table для открытой таблицы или import_table с её именем в target.

РЕДАКТИРОВАНИЕ ДОКУМЕНТА
- Содержимое активного документа приходит в сообщении блоком «[Файл: имя]» или «[Прикреплённый файл: имя]». Ты уже видишь текст документа и имеешь право записи в него.
- Правки применяются к Word (.docx/.doc), таблицам (.xlsx/.csv), тексту, Markdown и коду. Презентации (.pptx) приходят текстом слайдов только для чтения — переписать их нельзя, предложи вместо этого новый .docx или .md. PDF в чат не передаётся: попроси открыть его как текст или DOCX.
- «Исправь ошибки», «перепиши», «сократи», «оформи», «приведи в порядок», «сделай красиво» — это команда заменить документ целиком: верни replace_document с полным новым текстом.
- Дополнить документ или вставить фрагмент — insert_into_editor.
- Запрещено отвечать «не могу редактировать документ Word», «вот рекомендации по оформлению» и вообще давать советы вместо правки. Если содержимого документа в сообщении нет, одной строкой попроси открыть файл и приложить его.
- Оформление задавай Markdown-разметкой, приложение превратит её в настоящие стили Word: # — заголовок 1, ## — заголовок 2, ### — заголовок 3, - — маркированный список, 1. — нумерованный, **жирный**, > — цитата, | таблица |, --- — разделитель.
- Перед блоком действия — одна короткая строка о том, что изменено. Сам текст документа в чат не дублируй: он попадёт в документ из блока действия.
- Не добавляй в конец разбор «Что улучшено / что изменено», предложения («Хотите, я создам файл…», «Могу также…»), вопросы и списки вариантов: документ заканчивается содержательным текстом.
- Не используй эмодзи ни в документе, ни в ответе.
- Если блок действия вернуть не получается, отдай готовый текст документа обычным сообщением без вступления — приложение применит его к открытому документу.

ВЛОЖЕННЫЕ ФАЙЛЫ
- Если в сообщении пользователя есть блок «[Файл: имя]» — это уже прочитанное содержимое файла. Используй его напрямую.
- Никогда не отвечай, что у тебя нет доступа к файлам или файловой системе: ты управляешь WorkSpace и можешь создавать и изменять файлы через действия.
- Если данных в файле не хватает, попроси конкретный файл или уточнение, но не отказывайся от работы.

ПРАВИЛА
- Пути указывай относительно корня рабочей области, если не нужен абсолютный.
- Если параметр не задан пользователем, выбери разумное значение (имя файла, лист, папку).
- Для нескольких шагов возвращай несколько действий в правильном порядке.
- Учитывай состояние приложения из блока "Текущее состояние WorkSpace": активную вкладку, открытые файлы, панели.
- Отвечай на русском языке, кратко и по делу.`;

export function resolveSystemPrompt(userPrompt?: string, appContext?: string): string {
  const base = (userPrompt || '').trim();
  const operator = base.includes('workspace-action')
    ? base
    : base
      ? `${base}\n\n${WORKSPACE_OPERATOR_SYSTEM_PROMPT}`
      : WORKSPACE_OPERATOR_SYSTEM_PROMPT;

  const context = (appContext || '').trim();
  if (!context) return operator;

  return `${operator}\n\nТекущее состояние WorkSpace (актуально на момент запроса):\n${context}`;
}

export const aiService = {
  /** Проверка соединения и валидности ключа — для кнопки «Проверить подключение». */
  async testConnection(settings: AISettings): Promise<{ ok: boolean; message: string }> {
    const preset = PROVIDER_PRESETS[settings.provider];
    if (preset.requiresKey && !settings.apiKey.trim()) {
      return { ok: false, message: 'API-ключ не указан' };
    }

    const endpoint = this.resolveEndpoint(settings);
    const model = settings.model || 'не выбрана';

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
        message: `Подключение успешно: ${endpoint} · модель ${model}. Ответ: «${received.trim().slice(0, 40) || 'ок'}»`,
      };
    } catch (err) {
      return { ok: false, message: `${endpoint} · модель ${model}: ${(err as Error).message}` };
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
    signal?: AbortSignal,
    appContext?: string
  ): Promise<void> {
    // Адрес и модель обязаны соответствовать провайдеру: иначе ключ Starimg
    // уходит на api.openai.com и приложение «перестаёт подключаться».
    const effective = normalizeSettings(settings);
    const preset = PROVIDER_PRESETS[effective.provider];
    if (preset.requiresKey && !effective.apiKey.trim()) {
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
      if (effective.provider === 'anthropic') {
        await this.streamAnthropic(clean, effective, callbacks, signal, appContext);
      } else if (effective.provider === 'gemini') {
        await this.streamGemini(clean, effective, callbacks, signal, appContext);
      } else {
        await this.streamOpenAICompatible(clean, effective, callbacks, signal, appContext);
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
    signal?: AbortSignal,
    appContext?: string
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
          { role: 'system', content: resolveSystemPrompt(settings.systemPrompt, appContext) },
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
    signal?: AbortSignal,
    appContext?: string
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
        system: resolveSystemPrompt(settings.systemPrompt, appContext),
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
    signal?: AbortSignal,
    appContext?: string
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
        systemInstruction: { parts: [{ text: resolveSystemPrompt(settings.systemPrompt, appContext) }] },
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
