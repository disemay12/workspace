import { AISettings, ChatMessage } from '../types';
import { PROVIDER_PRESETS, normalizeSettings } from '../state/settings';
import { uid } from '../lib/utils';
import type { ToolDefinition } from './aiTools';

/** Максимальный размер прикрепляемого к запросу файла (символов). */
export const MAX_ATTACHMENT_CHARS = 40000;

export interface StreamCallbacks {
  onChunk: (partialText: string) => void;
  onFinish: (fullText: string) => void;
  onError: (error: Error) => void;
}

/** Провайдер/модель отвергли tools — нужен текстовый протокол document_apply. */
export class ToolsUnsupportedError extends Error {
  constructor(message = 'Провайдер не поддерживает function calling') {
    super(message);
    this.name = 'ToolsUnsupportedError';
  }
}

export interface DocumentAgentOptions {
  messages: ChatMessage[];
  settings: AISettings;
  appContext?: string;
  systemPromptExtra?: string;
  tools: ToolDefinition[];
  onToolCall: (name: string, args: Record<string, unknown>) => Promise<unknown>;
  onStatus?: (status: string) => void;
  signal?: AbortSignal;
  /** Страховка от бесконечного цикла: каждая итерация — один ответ модели. */
  maxIterations?: number;
}

export interface DocumentAgentResult {
  text: string;
  iterations: number;
  toolCalls: number;
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

/** Разбор SSE-потока на уровне JSON-событий. */
function createPayloadParser(onPayload: (payload: any) => void) {
  let buffer = '';
  let count = 0;

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
          count += 1;
          onPayload(payload);
        } catch {
          // Неполный JSON-чанк — пропускаем
        }
      }
    },
    get count() {
      return count;
    },
  };
}

/**
 * Разбор SSE-потока с извлечением приращения текста.
 * Возвращает накопленный текст.
 */
function createSSEParser(extractDelta: (payload: any) => string, onChunk: (partial: string) => void) {
  let full = '';
  const parser = createPayloadParser((payload) => {
    const delta = extractDelta(payload);
    if (delta) {
      full += delta;
      onChunk(full);
    }
  });

  return {
    push: (text: string) => parser.push(text),
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
Если API предоставил инструменты (tools), выполняй действия через них, а блоки workspace-action используй только для действий, которых среди инструментов нет.
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

20) save_as — сохранить открытый документ под новым именем или в другом формате
{ "action": "save_as", "format": "xlsx", "path": "Отчёт.xlsx" }
Форматы: xlsx, csv, tsv, docx, md, txt, html, pdf. Если открыт CSV и просят «сохрани в Excel» — вызывай save_as с "format": "xlsx". Не создавай новый файл вручную и не переписывай данные. Без path файл создаётся рядом с исходным. Существующий файл не перезаписывается: агент добавит « (2)».

21) convert_file — конвертировать любой файл рабочей папки
{ "action": "convert_file", "source": "notes.md", "target": "notes.docx", "format": "docx" }
Поддерживается: md/txt/html ↔ docx, csv/tsv ↔ xlsx, docx/xlsx/pptx → md/txt/html/pdf, Markdown-таблица → xlsx. Для PDF используй format "pdf".

РЕДАКТИРОВАНИЕ ДОКУМЕНТА
- Имя открытого документа и права записи — в блоке «ОТКРЫТЫЙ ДОКУМЕНТ» ниже. Не додумывай права: если там написано «НЕТ», правку выполнять нельзя.
- Если тебе переданы инструменты документа (get_document, get_outline, get_block, get_range, get_selection, get_document_info, search_document, replace_document, replace_block, insert_after, insert_before, append_to_block, move_block, duplicate_block, split_block, merge_blocks, table_read, table_insert_row, table_delete_row, table_insert_column, table_delete_column, table_set_cell, table_merge_cells, table_unmerge_cells, table_set_column_width, table_set_borders, insert_image, update_image, delete_image, insert_page_break, insert_section_break, delete_section_break, insert_toc, refresh_toc, insert_field, insert_bookmark, add_comment, add_footnote, set_paragraph_format, set_tabs, set_paragraph_style, define_style, set_header_footer, delete_header_footer, delete_block, find_replace, replace_range, delete_range, replace_selection, insert_at_cursor, format, page_setup, read_document) — правь документ только ими.
- Если инструменты не переданы (модель без function calling), используй действие document_apply с массивом operations — те же операции, что и у инструментов.
- Не создавай вместо правки новые файлы и не пиши «не могу редактировать документ Word». Если открыт Word/Markdown, не предлагай и не создавай Excel/CSV (create_file, create_table), пока пользователь прямо не попросит таблицу.
- ATS/ATC-фильтры: полностью перепиши документ плоским одноколоночным текстом — заголовки «Опыт работы», «Навыки», «Образование», ключевые слова, маркированные списки; без таблиц, колонок и картинок. Не возвращай копию исходного текста.
- Поля страницы, шрифт, кегль и межстрочный интервал меняет операция page_setup. Никогда не советуй пользователю настраивать Word вручную («установите поля», «скопируйте в Word», «выберите шрифт»): «оформи как реальный документ», заявление, письмо, договор и приказ оформляются через page_setup с template "russian_business".
- Точечная правка — одна строка, фраза, дата, имя или пункт — выполняется find_replace (find — точный фрагмент из документа) или replace_block. replace_document в этом случае запрещён: не переписывай весь документ из-за одной строки.
- «Выделенное», «этот фрагмент», «здесь»: get_selection, затем replace_selection. Правка части абзаца: get_range → replace_range/delete_range. Другой файл рабочей папки: read_document; он доступен только для чтения.
- Оформление в инструментах документа передавай HTML-тегами: <h1>–<h3>, <p>, <ul>/<ol> с <li>, <table>, <strong>, <em>, <u>, <s>, inline style для выравнивания, шрифта, размера и отступов. Markdown внутри HTML не применяй.
- Сам текст документа в чат не дублируй: он попадает в документ. В чате — одна короткая строка отчёта о том, что изменено.
- В replace_document передавай только сам документ. Отчёт «что изменено», список правок со стрелками и вопрос «сохранить?» — это сообщение чата, а не содержимое документа.
- insert_into_editor вставляй только с самим текстом: без вступления, отчёта, советов и обёртки из обратных кавычек. Для точечной правки строки используй find_replace или replace_block, а не replace_document.
- Не добавляй в конец разбор «Что улучшено / что изменено», предложения («Хотите, я создам файл…», «Могу также…»), вопросы и списки вариантов.
- Не используй эмодзи ни в документе, ни в ответе.

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

/** Системный промпт = пользовательский/встроенный + состояние приложения + блок документа. */
function composeSystemPrompt(
  settings: AISettings,
  appContext?: string,
  extraSystem?: string
): string {
  const base = resolveSystemPrompt(settings.systemPrompt, appContext);
  const extra = (extraSystem || '').trim();
  return extra ? `${base}\n\n${extra}` : base;
}

/** Провайдер ответил, что tools/function calling не поддерживаются. */
function isToolsUnsupported(error: unknown): boolean {
  const status = (error as { status?: number } | null)?.status;
  if (status !== 400 && status !== 404 && status !== 422) return false;
  const message = String((error as Error)?.message || '');
  return /tool|function|unsupported|not supported|invalid parameter/i.test(message);
}

function describeToolStatus(name: string): string {
  switch (name) {
    case 'get_document':
      return 'Читаю документ…';
    case 'replace_document':
      return 'Переписываю документ…';
    case 'find_replace':
      return 'Ищу и заменяю текст…';
    case 'format':
      return 'Меняю оформление…';
    case 'page_setup':
      return 'Настраиваю страницу и шрифт…';
    case 'delete_block':
      return 'Удаляю лишнее…';
    default:
      return 'Правлю документ…';
  }
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
    appContext?: string,
    extraSystem?: string
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
        await this.streamAnthropic(clean, effective, callbacks, signal, appContext, extraSystem);
      } else if (effective.provider === 'gemini') {
        await this.streamGemini(clean, effective, callbacks, signal, appContext, extraSystem);
      } else {
        await this.streamOpenAICompatible(clean, effective, callbacks, signal, appContext, extraSystem);
      }
    } catch (err) {
      if ((err as Error)?.name === 'AbortError') return;
      callbacks.onError(describeError(err, preset.label));
    }
  },

  /**
   * Низкоуровневый POST-запрос: отдаёт сырые чанки тела ответа.
   *
   * В Electron запрос идёт через главный процесс (`net`-модуль), потому что
   * браузерный fetch блокируется политикой CORS: часть OpenAI-совместимых
   * шлюзов не отдаёт `Access-Control-Allow-Origin`, и Chromium прерывает
   * ответ с `net::ERR_FAILED`.
   *
   * В браузере (режим отладки без Electron) используется обычный fetch.
   */
  async transportStream(
    endpoint: string,
    headers: Record<string, string>,
    bodyText: string,
    onChunk: (chunk: string) => void,
    signal?: AbortSignal,
    label = 'Ошибка API'
  ): Promise<void> {
    /* ---------- Путь через Electron ---------- */
    if (isElectron() && window.electronAPI?.aiStream) {
      const requestId = uid('ai');

      return new Promise<void>((resolve, reject) => {
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
          onChunk(data.chunk);
        });

        const offEnd = window.electronAPI!.onAiEnd((data) => {
          if (data.requestId !== requestId) return;
          finish(() => resolve());
        });

        const offError = window.electronAPI!.onAiError((data) => {
          if (data.requestId !== requestId) return;
          finish(() => {
            if (data.status) {
              const error = parseApiError(data.status, data.message, label);
              (error as Error & { status?: number }).status = data.status;
              reject(error);
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
      const error = parseApiError(response.status, errText, label);
      (error as Error & { status?: number }).status = response.status;
      throw error;
    }
    if (!response.body) throw new Error('Сервер вернул пустой ответ');

    const reader = response.body.getReader();
    const decoder = new TextDecoder('utf-8');

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      onChunk(decoder.decode(value, { stream: true }));
    }
  },

  /** Стриминговый запрос с накоплением текста. */
  async requestStream(
    endpoint: string,
    headers: Record<string, string>,
    body: unknown,
    extractDelta: (payload: any) => string,
    onChunk: (partial: string) => void,
    signal?: AbortSignal,
    label = 'Ошибка API'
  ): Promise<string> {
    const parser = createSSEParser(extractDelta, onChunk);
    await this.transportStream(
      endpoint,
      headers,
      JSON.stringify(body),
      (chunk) => parser.push(chunk),
      signal,
      label
    );
    return parser.result;
  },

  /**
   * Стриминговый запрос с доступом к каждому SSE-событию (нужен для tool_calls).
   * Если провайдер проигнорировал stream и вернул обычный JSON — разбирает его.
   */
  async requestEventStream(
    endpoint: string,
    headers: Record<string, string>,
    body: unknown,
    onPayload: (payload: any) => void,
    signal?: AbortSignal,
    label = 'Ошибка API'
  ): Promise<void> {
    const parser = createPayloadParser(onPayload);
    let raw = '';
    await this.transportStream(
      endpoint,
      headers,
      JSON.stringify(body),
      (chunk) => {
        raw += chunk;
        parser.push(chunk);
      },
      signal,
      label
    );

    if (parser.count === 0 && raw.trim()) {
      try {
        onPayload(JSON.parse(raw));
      } catch {
        throw new Error('Сервер вернул ответ не в формате SSE или JSON');
      }
    }
  },

  /**
   * Агентный цикл с нативным function calling.
   *
   * Модель вызывает инструменты, приложение выполняет их над живым документом,
   * результат возвращается модели, и так до финального текстового ответа.
   * Число итераций ограничено.
   */
  async runDocumentAgent(options: DocumentAgentOptions): Promise<DocumentAgentResult> {
    const effective = normalizeSettings(options.settings);
    const endpoint = this.resolveEndpoint(effective);

    const headers: Record<string, string> = { 'Content-Type': 'application/json' };
    if (effective.apiKey.trim()) headers['Authorization'] = `Bearer ${effective.apiKey.trim()}`;
    if (effective.provider === 'openrouter') {
      headers['HTTP-Referer'] = 'https://workspace.local';
      headers['X-Title'] = 'WorkSpace';
    }

    const apiMessages: any[] = [
      {
        role: 'system',
        content: composeSystemPrompt(effective, options.appContext, options.systemPromptExtra),
      },
      ...options.messages
        .filter((message) => message.role !== 'system' && message.content.trim())
        .map((message) => ({ role: message.role, content: buildMessageText(message) })),
    ];

    const maxIterations = Math.max(1, Math.min(options.maxIterations ?? 8, 16));
    let lastText = '';
    let toolCallsCount = 0;

    for (let iteration = 0; iteration < maxIterations; iteration += 1) {
      const textParts: string[] = [];
      const calls = new Map<number, { id: string; name: string; args: string }>();

      try {
        await this.requestEventStream(
          endpoint,
          headers,
          {
            model: effective.model || 'gpt-4o-mini',
            messages: apiMessages,
            tools: options.tools,
            tool_choice: 'auto',
            temperature: effective.temperature,
            max_tokens: Math.max(effective.maxTokens, 4096),
            stream: true,
          },
          (payload) => {
            const choice = payload?.choices?.[0];
            if (!choice) return;

            const delta = choice.delta ?? choice.message ?? {};
            if (typeof delta.content === 'string') textParts.push(delta.content);

            const toolDeltas = Array.isArray(delta.tool_calls) ? delta.tool_calls : [];
            for (const part of toolDeltas) {
              const index = typeof part.index === 'number' ? part.index : calls.size;
              const entry = calls.get(index) ?? { id: '', name: '', args: '' };
              if (typeof part.id === 'string' && part.id) entry.id = part.id;

              const fn = part.function ?? {};
              if (typeof fn.name === 'string' && fn.name) {
                // Имя обычно приходит целиком в одном чанке, но некоторые
                // шлюзы дробят его — дописываем только новые фрагменты.
                if (!entry.name) entry.name = fn.name;
                else if (fn.name !== entry.name && !entry.name.endsWith(fn.name)) entry.name += fn.name;
              }
              if (typeof fn.arguments === 'string') entry.args += fn.arguments;
              else if (fn.arguments && typeof fn.arguments === 'object') {
                entry.args = JSON.stringify(fn.arguments);
              }
              calls.set(index, entry);
            }
          },
          options.signal,
          'Ошибка API'
        );
      } catch (error) {
        if ((error as Error)?.name === 'AbortError') throw error;
        if (isToolsUnsupported(error)) {
          throw new ToolsUnsupportedError((error as Error).message);
        }
        throw error;
      }

      const text = textParts.join('');
      if (text.trim()) lastText = text.trim();

      if (calls.size === 0) {
        return { text: lastText, iterations: iteration + 1, toolCalls: toolCallsCount };
      }

      const toolCalls = [...calls.entries()]
        .sort((a, b) => a[0] - b[0])
        .map(([, call], order) => ({
          id: call.id || `call_${iteration}_${order}`,
          type: 'function' as const,
          function: { name: call.name, arguments: call.args || '{}' },
        }))
        .filter((call) => call.function.name);

      if (toolCalls.length === 0) {
        return { text: lastText, iterations: iteration + 1, toolCalls: toolCallsCount };
      }

      apiMessages.push({ role: 'assistant', content: text || null, tool_calls: toolCalls });

      for (const call of toolCalls) {
        toolCallsCount += 1;
        let result: unknown;

        let args: Record<string, unknown>;
        try {
          args = call.function.arguments.trim()
            ? (JSON.parse(call.function.arguments) as Record<string, unknown>)
            : {};
        } catch {
          result = {
            ok: false,
            message: `Аргументы инструмента ${call.function.name} — некорректный JSON. Вызови инструмент заново.`,
          };
          apiMessages.push({
            role: 'tool',
            tool_call_id: call.id,
            content: JSON.stringify(result),
          });
          continue;
        }

        options.onStatus?.(describeToolStatus(call.function.name));
        try {
          result = await options.onToolCall(call.function.name, args);
        } catch (error) {
          result = {
            ok: false,
            message: (error as Error).message || 'Ошибка выполнения инструмента',
          };
        }

        const serialized = JSON.stringify(result ?? { ok: true });
        apiMessages.push({
          role: 'tool',
          tool_call_id: call.id,
          content: serialized.length > 24000 ? `${serialized.slice(0, 24000)}…` : serialized,
        });
      }
    }

    return {
      text: lastText || 'Достигнут лимит шагов работы с документом. Уточните задачу или продолжите отдельным сообщением.',
      iterations: maxIterations,
      toolCalls: toolCallsCount,
    };
  },

  async streamOpenAICompatible(
    messages: ChatMessage[],
    settings: AISettings,
    { onChunk, onFinish }: StreamCallbacks,
    signal?: AbortSignal,
    appContext?: string,
    extraSystem?: string
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
          { role: 'system', content: composeSystemPrompt(settings, appContext, extraSystem) },
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
    appContext?: string,
    extraSystem?: string
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
        system: composeSystemPrompt(settings, appContext, extraSystem),
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
    appContext?: string,
    extraSystem?: string
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
        systemInstruction: {
          parts: [{ text: composeSystemPrompt(settings, appContext, extraSystem) }],
        },
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
