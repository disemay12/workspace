import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Sparkles,
  Send,
  Square,
  Paperclip,
  Trash2,
  Plus,
  Copy,
  Check,
  Code2,
  CornerDownLeft,
  AlertTriangle,
  User,
  History,
  X,
  FileText,
  FileSpreadsheet,
  FilePlus,
  FolderOpen,
  Play,
  ExternalLink,
} from 'lucide-react';
import { AISettings, ChatMessage, ChatSession, OpenTab } from '../../types';
import { aiService } from '../../services/aiService';
import { SYSTEM_PRESETS, PROVIDER_PRESETS } from '../../state/settings';
import { cn, extractCodeBlocks, uid } from '../../lib/utils';
import { renderRichMarkdown, renderMermaidBlocks } from '../../lib/markdown';
import {
  AIActionPayload,
  ParsedMarkdownTable,
  parseAIActions,
  detectMarkdownTables,
  stripActionBlocks,
} from '../../services/aiActions';
import { AIBadge, Button, IconButton, Spinner, Textarea } from '../ui';

interface Props {
  settings: AISettings;
  isConfigured: boolean;
  onUpdateModel?: (model: string) => void;
  availableModels?: string[];
  activeTab: OpenTab | null;
  sessions: ChatSession[];
  activeSessionId: string;
  onSelectSession: (id: string) => void;
  onCreateSession: () => string;
  onUpdateMessages: (sessionId: string, messages: ChatMessage[]) => void;
  onDeleteSession: (id: string) => void;
  onClearAll: () => void;
  onOpenSettings: () => void;
  onInsertIntoEditor: (text: string) => void;
  onApplyToDocument: (text: string) => void;
  /** Внешний запрос от других компонентов (анализ лога, формула Excel и т.п.). */
  pendingPrompt: { prompt: string; context: string; token: number } | null;
  onPendingPromptConsumed: () => void;
  /**
   * Готовый ответ ассистента. Используется редактором таблиц, чтобы
   * предложить применить сгенерированную формулу.
   */
  onAnswerComplete?: (answer: string) => void;
  onCreateTable?: (
    name: string,
    columns: string[],
    rows: (string | number | boolean | null)[][],
    sheetName?: string,
    format?: 'xlsx' | 'csv'
  ) => Promise<string | null>;
  onCreateFile?: (name: string, content: string) => Promise<string | null>;
  onOpenFile?: (path: string) => Promise<void>;
}

export const AIChatPanel: React.FC<Props> = ({
  settings,
  isConfigured,
  onUpdateModel,
  availableModels,
  activeTab,
  sessions,
  activeSessionId,
  onSelectSession,
  onCreateSession,
  onUpdateMessages,
  onDeleteSession,
  onClearAll,
  onOpenSettings,
  onInsertIntoEditor,
  onApplyToDocument,
  pendingPrompt,
  onPendingPromptConsumed,
  onAnswerComplete,
  onCreateTable,
  onCreateFile,
  onOpenFile,
}) => {
  const [input, setInput] = useState('');
  const [isStreaming, setIsStreaming] = useState(false);
  const [attachFile, setAttachFile] = useState(false);
  const [attachSelection, setAttachSelection] = useState(false);
  const [selectionText, setSelectionText] = useState('');
  const [copiedId, setCopiedId] = useState('');
  const [showHistory, setShowHistory] = useState(false);

  // Состояние выполнения действий AI (создание таблиц, файлов и т.п.)
  const [actionStates, setActionStates] = useState<
    Record<string, { status: 'idle' | 'running' | 'done' | 'error'; result?: string; error?: string }>
  >({});
  const executedActionsRef = useRef<Set<string>>(new Set());

  const scrollRef = useRef<HTMLDivElement>(null);
  const abortRef = useRef<AbortController | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const activeSession = useMemo(
    () => sessions.find((s) => s.id === activeSessionId) || null,
    [sessions, activeSessionId]
  );
  const messages = activeSession?.messages ?? [];

  /* ---------- Автопрокрутка к последнему сообщению ---------- */
  useEffect(() => {
    const el = scrollRef.current;
    if (el) el.scrollTop = el.scrollHeight;
  }, [messages.length, messages[messages.length - 1]?.content]);

  /* ---------- Отслеживание выделения в активном документе ---------- */
  useEffect(() => {
    const handler = () => {
      const sel = window.getSelection()?.toString() ?? '';
      setSelectionText(sel.trim());
    };
    document.addEventListener('selectionchange', handler);
    return () => document.removeEventListener('selectionchange', handler);
  }, []);

  /* ---------- Отправка сообщения ---------- */

  const submit = useCallback(
    async (rawText: string, attached?: ChatMessage['attachedFile']) => {
      const text = rawText.trim();
      if (!text || isStreaming) return;

      if (!isConfigured) {
        onOpenSettings();
        return;
      }

      let sessionId = activeSessionId;
      if (!sessionId) sessionId = onCreateSession();

      const userMessage: ChatMessage = {
        id: uid('msg'),
        role: 'user',
        content: text,
        timestamp: Date.now(),
        attachedFile: attached,
      };
      const assistantId = uid('msg');
      const assistantMessage: ChatMessage = {
        id: assistantId,
        role: 'assistant',
        content: '',
        timestamp: Date.now(),
        streaming: true,
      };

      // Актуальная история на момент отправки
      const base = sessionId === activeSessionId ? messages : [];
      const nextMessages = [...base, userMessage, assistantMessage];
      onUpdateMessages(sessionId, nextMessages);
      setInput('');
      setIsStreaming(true);

      const controller = new AbortController();
      abortRef.current = controller;

      // Локальный буфер для стрима — состояние обновляется по мере поступления текста
      let buffer = '';
      let lastPaint = 0;

      const paint = (force = false) => {
        const now = Date.now();
        // Ограничиваем частоту перерисовки для длинных ответов
        if (!force && now - lastPaint < 60) return;
        lastPaint = now;
        onUpdateMessages(
          sessionId,
          nextMessages.map((m) => (m.id === assistantId ? { ...m, content: buffer } : m))
        );
      };

      try {
        await aiService.streamChat(
          [...base, userMessage],
          settings,
          {
            onChunk: (partial) => {
              buffer = partial;
              paint();
            },
            onFinish: (full) => {
              buffer = full;
              paint(true);
            },
            onError: (error) => {
              throw error;
            },
          },
          controller.signal
        );

        onUpdateMessages(
          sessionId,
          nextMessages.map((m) =>
            m.id === assistantId ? { ...m, content: buffer, streaming: false } : m
          )
        );

        // Сообщаем наружу готовый ответ — редактор таблицы предложит
        // применить сгенерированную формулу.
        if (buffer.trim()) onAnswerComplete?.(buffer);
      } catch (err) {
        const message = (err as Error).message || 'Неизвестная ошибка';
        onUpdateMessages(
          sessionId,
          nextMessages.map((m) =>
            m.id === assistantId
              ? {
                  ...m,
                  content: buffer,
                  streaming: false,
                  error: message,
                }
              : m
          )
        );
      } finally {
        setIsStreaming(false);
        abortRef.current = null;
      }
    },
    [
      isStreaming,
      isConfigured,
      activeSessionId,
      messages,
      settings,
      onCreateSession,
      onUpdateMessages,
      onOpenSettings,
      onAnswerComplete,
    ]
  );

  /* ---------- Выполнение действий управления программой ---------- */

  const executeAction = useCallback(
    async (action: AIActionPayload, actionKey: string) => {
      setActionStates((prev) => ({
        ...prev,
        [actionKey]: { status: 'running' },
      }));

      try {
        let resultPath: string | null = null;
        if (action.action === 'create_table') {
          if (onCreateTable) {
            resultPath = await onCreateTable(
              action.name || 'Таблица.xlsx',
              action.columns || [],
              action.rows || [],
              action.sheetName,
              action.format
            );
          }
        } else if (action.action === 'create_file') {
          if (onCreateFile) {
            resultPath = await onCreateFile(
              action.name || action.path || 'файл.txt',
              action.content || ''
            );
          }
        } else if (action.action === 'open_file') {
          if (onOpenFile) {
            await onOpenFile(action.path);
            resultPath = action.path;
          }
        } else if (action.action === 'insert_into_editor') {
          onInsertIntoEditor(action.text);
          resultPath = 'Вставлено в редактор';
        } else if (action.action === 'replace_document') {
          onApplyToDocument(action.text);
          resultPath = 'Документ заменён';
        }

        setActionStates((prev) => ({
          ...prev,
          [actionKey]: { status: 'done', result: resultPath || undefined },
        }));
      } catch (err) {
        setActionStates((prev) => ({
          ...prev,
          [actionKey]: { status: 'error', error: (err as Error).message },
        }));
      }
    },
    [onCreateTable, onCreateFile, onOpenFile, onInsertIntoEditor, onApplyToDocument]
  );

  // Автоматическое выполнение действий из готовых ответов модели
  useEffect(() => {
    messages.forEach((msg) => {
      if (msg.role !== 'assistant' || msg.streaming) return;
      const acts = parseAIActions(msg.content);
      acts.forEach((act, idx) => {
        const actionKey = `${msg.id}_${idx}`;
        if (!executedActionsRef.current.has(actionKey)) {
          executedActionsRef.current.add(actionKey);
          void executeAction(act, actionKey);
        }
      });
    });
  }, [messages, executeAction]);

  const handleCreateTableFromMarkdown = useCallback(
    async (table: ParsedMarkdownTable, format: 'xlsx' | 'csv' = 'xlsx') => {
      if (!onCreateTable) return;
      const name = format === 'csv' ? 'Таблица_данных.csv' : 'Таблица_данных.xlsx';
      await onCreateTable(name, table.columns, table.rows, 'Лист 1', format);
    },
    [onCreateTable]
  );

  const stopStreaming = () => {
    abortRef.current?.abort();
    setIsStreaming(false);
  };

  /* ---------- Обработка внешних запросов ---------- */
  useEffect(() => {
    if (!pendingPrompt) return;
    const { prompt, context } = pendingPrompt;
    void submit(
      prompt,
      context
        ? {
            name: activeTab?.name || 'Выделенный фрагмент',
            path: activeTab?.path || '',
            content: context,
            truncated: context.length > 40000,
          }
        : undefined
    );
    onPendingPromptConsumed();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pendingPrompt?.token]);

  /* ---------- Сборка вложения ---------- */
  const buildAttachment = (): ChatMessage['attachedFile'] | undefined => {
    if (!activeTab) return undefined;
    if (attachSelection && selectionText) {
      return {
        name: `${activeTab.name} (выделенное)`,
        path: activeTab.path,
        content: selectionText,
        truncated: selectionText.length > 40000,
      };
    }
    if (attachFile) {
      return {
        name: activeTab.name,
        path: activeTab.path,
        content: activeTab.content,
        truncated: activeTab.content.length > 40000,
      };
    }
    return undefined;
  };

  const handleSend = () => {
    void submit(input.trim(), buildAttachment());
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const copyMessage = async (id: string, content: string) => {
    try {
      await navigator.clipboard.writeText(content);
      setCopiedId(id);
      window.setTimeout(() => setCopiedId(''), 1600);
    } catch {
      // Буфер обмена недоступен
    }
  };

  const providerLabel = PROVIDER_PRESETS[settings.provider]?.label ?? settings.provider;

  return (
    <div className="flex h-full w-full flex-col border-l border-ws bg-ws-panel">
      {/* Заголовок панели */}
      <div className="flex items-center justify-between border-b border-ws px-3 py-2">
        <div className="flex items-center gap-2">
          <Sparkles size={14} className="text-ws-accent" />
          <span className="text-[11px] font-semibold tracking-wider text-ws uppercase">
            AI Ассистент
          </span>
        </div>
        <div className="flex items-center gap-0.5">
          <IconButton
            title="История диалогов"
            active={showHistory}
            onClick={() => setShowHistory((v) => !v)}
          >
            <History size={14} />
          </IconButton>
          <IconButton title="Новый диалог" onClick={() => onCreateSession()}>
            <Plus size={14} />
          </IconButton>
        </div>
      </div>

      {/* Быстрый выбор модели */}
      <div className="flex items-center justify-between border-b border-ws bg-ws-elevated px-3 py-1.5 text-[11px]">
        <div className="flex items-center gap-1.5 text-ws-muted min-w-0">
          <Sparkles size={11} className="text-ws-accent shrink-0" />
          <span className="font-medium text-ws truncate">{providerLabel}:</span>
        </div>
        <select
          value={settings.model}
          onChange={(e) => onUpdateModel?.(e.target.value)}
          title="Быстрая смена модели"
          className="rounded border border-ws bg-ws-surface px-2 py-0.5 text-[11px] font-mono text-ws focus:border-ws-accent focus:outline-none max-w-[200px] truncate"
        >
          {(availableModels && availableModels.length > 0
            ? availableModels
            : PROVIDER_PRESETS[settings.provider]?.models || [settings.model]
          ).map((m) => (
            <option key={m} value={m}>
              {m}
            </option>
          ))}
          {!(
            (availableModels && availableModels.length > 0
              ? availableModels
              : PROVIDER_PRESETS[settings.provider]?.models) || []
          ).includes(settings.model) && (
            <option value={settings.model}>{settings.model}</option>
          )}
        </select>
      </div>

      {/* Список сохранённых диалогов */}
      {showHistory && (
        <div className="animate-fade-in max-h-52 overflow-y-auto border-b border-ws bg-ws-surface">
          <div className="flex items-center justify-between px-3 py-2">
            <span className="text-[11px] font-semibold text-ws-muted">
              Диалогов: {sessions.length}
            </span>
            {sessions.length > 0 && (
              <button
                onClick={() => {
                  onClearAll();
                  setShowHistory(false);
                }}
                className="text-[11px] text-ws-danger hover:underline"
              >
                Очистить всё
              </button>
            )}
          </div>
          {sessions.length === 0 && (
            <p className="px-3 pb-3 text-[11px] text-ws-faint">История пуста</p>
          )}
          {sessions.map((s) => (
            <div
              key={s.id}
              onClick={() => {
                onSelectSession(s.id);
                setShowHistory(false);
              }}
              className={cn(
                'group flex cursor-pointer items-center justify-between gap-2 px-3 py-1.5 text-[11px] transition-colors',
                s.id === activeSessionId
                  ? 'bg-ws-accent-soft text-ws-accent'
                  : 'text-ws-muted hover:bg-ws-hover'
              )}
            >
              <span className="flex items-center gap-1.5 truncate">
                <FileText size={11} className="shrink-0" />
                <span className="truncate">{s.title || 'Без названия'}</span>
              </span>
              <button
                onClick={(e) => {
                  e.stopPropagation();
                  onDeleteSession(s.id);
                }}
                title="Удалить диалог"
                className="shrink-0 opacity-0 transition-opacity group-hover:opacity-100"
              >
                <X size={12} className="hover:text-ws-danger" />
              </button>
            </div>
          ))}
        </div>
      )}

      {/* Лента сообщений */}
      <div ref={scrollRef} className="flex-1 space-y-3 overflow-y-auto p-3">
        {messages.length === 0 && (
          <div className="space-y-3 pt-4">
            <div className="text-center">
              <div className="mx-auto mb-2 flex h-9 w-9 items-center justify-center rounded-lg bg-ws-accent-soft">
                <Sparkles size={18} className="text-ws-accent" />
              </div>
              <p className="text-xs font-medium text-ws">Чем помочь?</p>
              <p className="mt-1 text-[11px] leading-relaxed text-ws-muted">
                Модель: <span className="text-ws-accent">{settings.model}</span> · {providerLabel}
              </p>
            </div>

            {/* Быстрые подсказки */}
            <div className="space-y-1.5 pt-2">
              {SYSTEM_PRESETS.slice(0, 4).map((preset) => (
                <button
                  key={preset.id}
                  onClick={() => {
                    void submit(preset.prompt);
                  }}
                  className="w-full rounded-lg border border-ws bg-ws-surface p-2.5 text-left transition-colors hover:border-ws-accent hover:bg-ws-hover"
                >
                  <div className="text-[11px] font-medium text-ws">{preset.name}</div>
                  <div className="mt-0.5 text-[10px] text-ws-faint">{preset.description}</div>
                </button>
              ))}
            </div>

            {!isConfigured && (
              <button
                onClick={onOpenSettings}
                className="flex w-full items-center gap-2 rounded-lg border border-amber-500/40 bg-amber-500/10 p-2.5 text-left text-[11px] text-ws-warning"
              >
                <AlertTriangle size={13} className="shrink-0" />
                <span>API-ключ не настроен. Нажмите, чтобы указать ключ.</span>
              </button>
            )}
          </div>
        )}

        {messages.map((msg) => (
          <MessageBubble
            key={msg.id}
            message={msg}
            isCopied={copiedId === msg.id}
            onCopy={() => copyMessage(msg.id, msg.content)}
            onInsert={(text) => onInsertIntoEditor(text)}
            onReplace={(text) => onApplyToDocument(text)}
            hasActiveTab={Boolean(activeTab)}
            actionStates={actionStates}
            onExecuteAction={executeAction}
            onOpenFile={onOpenFile}
            onCreateTableFromMarkdown={handleCreateTableFromMarkdown}
          />
        ))}
      </div>

      {/* Область ввода */}
      <div className="border-t border-ws p-3">
        {/* Индикаторы контекста */}
        {activeTab && (
          <div className="mb-2 flex flex-wrap items-center gap-1.5">
            <button
              onClick={() => setAttachFile((v) => !v)}
              title="Прикрепить содержимое открытого файла к запросу"
              className={cn(
                'flex items-center gap-1 rounded border px-1.5 py-0.5 text-[10px] transition-colors',
                attachFile
                  ? 'border-ws-accent bg-ws-accent-soft text-ws-accent'
                  : 'border-ws text-ws-faint hover:text-ws'
              )}
            >
              <Paperclip size={10} />
              {activeTab.name}
            </button>

            {selectionText && (
              <button
                onClick={() => setAttachSelection((v) => !v)}
                title="Прикрепить только выделенный фрагмент"
                className={cn(
                  'flex items-center gap-1 rounded border px-1.5 py-0.5 text-[10px] transition-colors',
                  attachSelection
                    ? 'border-ws-accent bg-ws-accent-soft text-ws-accent'
                    : 'border-ws text-ws-faint hover:text-ws'
                )}
              >
                <Code2 size={10} />
                Выделено: {selectionText.length} симв.
              </button>
            )}
          </div>
        )}

        <div className="relative rounded-lg border border-ws bg-ws-elevated focus-within:border-ws-accent">
          <Textarea
            ref={textareaRef}
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={handleKeyDown}
            rows={3}
            placeholder="Спросите что угодно… (Enter — отправить, Shift+Enter — новая строка)"
            className="border-none bg-transparent px-3 py-2 text-xs focus:border-none"
          />

          <div className="flex items-center justify-between px-2 pb-1.5">
            <span className="text-[10px] text-ws-faint">
              {input.length > 0 ? `${input.length} симв.` : 'Ctrl+K — инлайн-правка в редакторе'}
            </span>

            {isStreaming ? (
              <Button variant="danger" onClick={stopStreaming} className="gap-1">
                <Square size={11} />
                Стоп
              </Button>
            ) : (
              <Button
                variant="primary"
                onClick={handleSend}
                disabled={!input.trim()}
                className="gap-1"
              >
                <Send size={11} />
                Отправить
              </Button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

/* ============================================================
   Пузырь сообщения
   ============================================================ */

/* ============================================================
   Карточка действия AI (создание таблицы, файла и т.п.)
   ============================================================ */

const AIActionCard: React.FC<{
  action: AIActionPayload;
  status: 'idle' | 'running' | 'done' | 'error';
  result?: string;
  error?: string;
  onExecute: () => void;
  onOpen?: () => void;
}> = ({ action, status, result, error, onExecute, onOpen }) => {
  if (action.action === 'create_table') {
    const colCount = action.columns?.length || 0;
    const rowCount = action.rows?.length || 0;
    const isCsv = action.format === 'csv' || (action.name && action.name.toLowerCase().endsWith('.csv'));
    const fileName = action.name || (isCsv ? 'Таблица.csv' : 'Таблица.xlsx');

    return (
      <div className="mt-2 rounded-lg border border-emerald-500/30 bg-emerald-500/10 p-2.5 text-xs">
        <div className="flex items-start justify-between gap-2">
          <div className="flex items-start gap-2 min-w-0">
            <div className="mt-0.5 rounded bg-emerald-500/20 p-1 text-emerald-400 shrink-0">
              <FileSpreadsheet size={15} />
            </div>
            <div className="min-w-0">
              <div className="font-semibold text-ws truncate">
                Таблица: <span className="text-emerald-400 font-mono">{fileName}</span>
              </div>
              <div className="text-[11px] text-ws-muted mt-0.5">
                {colCount > 0 && `${colCount} столбцов`}
                {rowCount > 0 && ` · ${rowCount} строк`}
                {action.sheetName && ` · Лист: «${action.sheetName}»`}
                {` · ${isCsv ? 'CSV' : 'Excel (.xlsx)'}`}
              </div>
            </div>
          </div>

          <div className="shrink-0">
            {status === 'running' ? (
              <span className="flex items-center gap-1.5 text-[11px] text-ws-accent">
                <Spinner size={12} />
                Создание...
              </span>
            ) : status === 'done' ? (
              <div className="flex items-center gap-1.5">
                <span className="flex items-center gap-1 text-[11px] font-medium text-emerald-400">
                  <Check size={12} />
                  Создана
                </span>
                {onOpen && (
                  <button
                    onClick={onOpen}
                    title="Перейти к открытой таблице"
                    className="flex items-center gap-1 rounded border border-ws bg-ws-surface px-1.5 py-0.5 text-[10px] text-ws hover:bg-ws-hover transition-colors"
                  >
                    <ExternalLink size={10} />
                    Открыть
                  </button>
                )}
              </div>
            ) : status === 'error' ? (
              <div className="flex items-center gap-1.5">
                <span className="text-[10px] text-ws-danger" title={error}>
                  Ошибка
                </span>
                <button
                  onClick={onExecute}
                  className="rounded bg-ws-surface px-2 py-0.5 text-[10px] text-ws hover:bg-ws-hover"
                >
                  Повторить
                </button>
              </div>
            ) : (
              <button
                onClick={onExecute}
                className="flex items-center gap-1 rounded bg-emerald-600 px-2 py-1 text-[11px] font-medium text-white hover:bg-emerald-500 shadow-sm transition-all"
              >
                <Play size={10} />
                Создать и открыть
              </button>
            )}
          </div>
        </div>
      </div>
    );
  }

  if (action.action === 'create_file') {
    const fileName = action.name || action.path || 'Новый файл';
    const contentLen = action.content?.length || 0;

    return (
      <div className="mt-2 rounded-lg border border-blue-500/30 bg-blue-500/10 p-2.5 text-xs">
        <div className="flex items-start justify-between gap-2">
          <div className="flex items-start gap-2 min-w-0">
            <div className="mt-0.5 rounded bg-blue-500/20 p-1 text-blue-400 shrink-0">
              <FilePlus size={15} />
            </div>
            <div className="min-w-0">
              <div className="font-semibold text-ws truncate">
                Файл: <span className="text-blue-400 font-mono">{fileName}</span>
              </div>
              <div className="text-[11px] text-ws-muted mt-0.5">
                {contentLen > 0 ? `${contentLen} символов` : 'Новый документ'}
              </div>
            </div>
          </div>

          <div className="shrink-0">
            {status === 'running' ? (
              <span className="flex items-center gap-1.5 text-[11px] text-ws-accent">
                <Spinner size={12} />
                Создание...
              </span>
            ) : status === 'done' ? (
              <div className="flex items-center gap-1.5">
                <span className="flex items-center gap-1 text-[11px] font-medium text-emerald-400">
                  <Check size={12} />
                  Создан
                </span>
                {onOpen && (
                  <button
                    onClick={onOpen}
                    className="flex items-center gap-1 rounded border border-ws bg-ws-surface px-1.5 py-0.5 text-[10px] text-ws hover:bg-ws-hover transition-colors"
                  >
                    <ExternalLink size={10} />
                    Открыть
                  </button>
                )}
              </div>
            ) : (
              <button
                onClick={onExecute}
                className="flex items-center gap-1 rounded bg-blue-600 px-2 py-1 text-[11px] font-medium text-white hover:bg-blue-500 shadow-sm transition-all"
              >
                <Play size={10} />
                Создать
              </button>
            )}
          </div>
        </div>
      </div>
    );
  }

  if (action.action === 'open_file') {
    return (
      <div className="mt-2 rounded-lg border border-ws bg-ws-surface p-2 text-xs flex items-center justify-between">
        <div className="flex items-center gap-2 truncate">
          <FolderOpen size={14} className="text-amber-400 shrink-0" />
          <span className="text-ws truncate">
            Открыть: <span className="font-mono text-ws-accent">{action.path}</span>
          </span>
        </div>
        <button
          onClick={onExecute}
          className="rounded bg-ws-elevated border border-ws px-2 py-0.5 text-[10px] text-ws hover:bg-ws-hover"
        >
          {status === 'done' ? 'Открыт' : 'Открыть'}
        </button>
      </div>
    );
  }

  return null;
};

/* ============================================================
   Пузырь сообщения
   ============================================================ */

const MessageBubble: React.FC<{
  message: ChatMessage;
  isCopied: boolean;
  onCopy: () => void;
  onInsert: (text: string) => void;
  onReplace: (text: string) => void;
  hasActiveTab: boolean;
  actionStates: Record<string, { status: 'idle' | 'running' | 'done' | 'error'; result?: string; error?: string }>;
  onExecuteAction: (action: AIActionPayload, actionKey: string) => void;
  onOpenFile?: (path: string) => void;
  onCreateTableFromMarkdown?: (table: ParsedMarkdownTable, format: 'xlsx' | 'csv') => void;
}> = ({
  message,
  isCopied,
  onCopy,
  onInsert,
  onReplace,
  hasActiveTab,
  actionStates,
  onExecuteAction,
  onOpenFile,
  onCreateTableFromMarkdown,
}) => {
  const isUser = message.role === 'user';
  const codeBlocks = useMemo(
    () => (isUser ? [] : extractCodeBlocks(message.content)),
    [isUser, message.content]
  );

  const actions = useMemo(
    () => (isUser ? [] : parseAIActions(message.content)),
    [isUser, message.content]
  );

  const mdTables = useMemo(
    () => (!isUser && actions.length === 0 ? detectMarkdownTables(message.content) : []),
    [isUser, actions.length, message.content]
  );

  const displayContent = useMemo(() => {
    if (isUser) return message.content;
    const cleaned = stripActionBlocks(message.content);
    if (!cleaned && actions.length > 0) {
      return 'Выполняю команду по управлению рабочей областью...';
    }
    return cleaned;
  }, [isUser, message.content, actions.length]);

  if (isUser) {
    return (
      <div className="flex justify-end">
        <div className="max-w-[92%] rounded-lg rounded-br-sm bg-ws-accent px-3 py-2 text-xs text-white">
          {message.attachedFile && (
            <div className="mb-1.5 flex items-center gap-1 rounded bg-white/15 px-1.5 py-0.5 text-[10px]">
              <Paperclip size={9} />
              <span className="truncate">{message.attachedFile.name}</span>
            </div>
          )}
          <p className="leading-relaxed whitespace-pre-wrap select-text">{message.content}</p>
        </div>
      </div>
    );
  }

  return (
    <div className="group flex gap-2">
      <div className="mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center rounded bg-ws-accent-soft">
        <Sparkles size={12} className="text-ws-accent" />
      </div>

      <div className="min-w-0 flex-1">
        <div className="rounded-lg rounded-tl-sm border border-ws bg-ws-surface px-3 py-2">
          {displayContent ? (
            <MarkdownLite content={displayContent} />
          ) : message.streaming ? (
            <span className="flex items-center gap-1.5 text-xs text-ws-faint">
              <span className="h-1.5 w-1.5 animate-pulse-soft rounded-full bg-ws-accent" />
              Думает…
            </span>
          ) : null}

          {/* Карточки действий AI */}
          {actions.map((act, idx) => {
            const actionKey = `${message.id}_${idx}`;
            const state = actionStates[actionKey] || { status: 'idle' };
            const targetPath =
              state.result ||
              (act.action === 'create_table'
                ? act.name
                : act.action === 'create_file'
                  ? act.name || act.path
                  : act.action === 'open_file'
                    ? act.path
                    : undefined);

            return (
              <AIActionCard
                key={actionKey}
                action={act}
                status={state.status}
                result={state.result}
                error={state.error}
                onExecute={() => onExecuteAction(act, actionKey)}
                onOpen={targetPath && onOpenFile ? () => onOpenFile(targetPath) : undefined}
              />
            );
          })}

          {/* Быстрое превращение Markdown-таблицы в рабочий файл */}
          {mdTables.length > 0 && !message.streaming && (
            <div className="mt-2.5 rounded-lg border border-emerald-500/30 bg-emerald-500/10 p-2.5">
              <div className="flex items-center gap-1.5 text-[11px] font-semibold text-emerald-400">
                <FileSpreadsheet size={13} />
                <span>
                  Обнаружена таблица данных ({mdTables[0].columns.length} столбцов, {mdTables[0].rows.length} строк)
                </span>
              </div>
              <p className="mt-0.5 text-[10px] text-ws-muted">
                Создайте полноценный рабочий файл таблицы в проекте в один клик:
              </p>
              <div className="mt-2 flex items-center gap-1.5">
                <button
                  onClick={() => onCreateTableFromMarkdown?.(mdTables[0], 'xlsx')}
                  className="flex items-center gap-1 rounded bg-emerald-600 px-2 py-0.5 text-[10px] font-medium text-white hover:bg-emerald-500 transition-colors shadow-sm"
                >
                  <Play size={9} /> Создать Excel (.xlsx)
                </button>
                <button
                  onClick={() => onCreateTableFromMarkdown?.(mdTables[0], 'csv')}
                  className="flex items-center gap-1 rounded border border-ws bg-ws-surface px-2 py-0.5 text-[10px] text-ws hover:bg-ws-hover transition-colors"
                >
                  Создать CSV (.csv)
                </button>
              </div>
            </div>
          )}
        </div>

        {message.error && (
          <div className="mt-1.5 flex items-start gap-1.5 rounded border border-red-500/30 bg-red-500/10 px-2 py-1.5 text-[11px] text-ws-danger">
            <AlertTriangle size={11} className="mt-0.5 shrink-0" />
            <span className="select-text">{message.error}</span>
          </div>
        )}

        {/* Действия над ответом */}
        {message.content && !message.streaming && (
          <div className="mt-1.5 flex flex-wrap items-center gap-1 opacity-60 transition-opacity group-hover:opacity-100 focus-within:opacity-100">
            <button
              onClick={onCopy}
              title="Скопировать ответ"
              className="flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
            >
              {isCopied ? <Check size={10} className="text-ws-success" /> : <Copy size={10} />}
              {isCopied ? 'Скопировано' : 'Копировать'}
            </button>

            {codeBlocks.length > 0 && hasActiveTab && (
              <button
                onClick={() => onInsert(codeBlocks[0].code)}
                title="Вставить первый блок кода в позицию курсора"
                className="flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
              >
                <CornerDownLeft size={10} />
                Вставить код
              </button>
            )}

            {codeBlocks.length > 1 && hasActiveTab && (
              <button
                onClick={() => onInsert(codeBlocks.map((b) => b.code).join('\n\n'))}
                title={`Вставить все блоки кода (${codeBlocks.length})`}
                className="flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
              >
                <Code2 size={10} />
                Все блоки ({codeBlocks.length})
              </button>
            )}

            {hasActiveTab && (
              <button
                onClick={() => onReplace(displayContent)}
                title="Заменить содержимое документа этим ответом"
                className="flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
              >
                <FileText size={10} />
                Заменить документ
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

/* ============================================================
   Продвинутый Markdown-рендерер для ответов модели (KaTeX + Prism + Mermaid)
   ============================================================ */

export const MarkdownLite: React.FC<{ content: string }> = ({ content }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const html = useMemo(() => renderRichMarkdown(content), [content]);

  useEffect(() => {
    void renderMermaidBlocks(containerRef.current);
  }, [html]);

  const handleClick = (e: React.MouseEvent) => {
    const target = (e.target as HTMLElement).closest('.copy-code-btn') as HTMLElement;
    if (!target) return;
    const raw = target.getAttribute('data-code');
    if (!raw) return;
    void navigator.clipboard.writeText(decodeURIComponent(raw));
    const orig = target.innerText;
    target.innerText = 'Скопировано!';
    setTimeout(() => {
      target.innerText = orig;
    }, 1500);
  };

  return (
    <div
      ref={containerRef}
      onClick={handleClick}
      className="ws-prose select-text text-xs"
      dangerouslySetInnerHTML={{ __html: html }}
    />
  );
};

export function renderMarkdown(markdown: string): string {
  return renderRichMarkdown(markdown);
}

export { User };
