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
} from 'lucide-react';
import { AISettings, ChatMessage, ChatSession, OpenTab } from '../../types';
import { aiService } from '../../services/aiService';
import { SYSTEM_PRESETS, PROVIDER_PRESETS } from '../../state/settings';
import { cn, extractCodeBlocks, uid } from '../../lib/utils';
import { renderRichMarkdown, renderMermaidBlocks } from '../../lib/markdown';
import { AIBadge, Button, IconButton, Textarea } from '../ui';

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
}) => {
  const [input, setInput] = useState('');
  const [isStreaming, setIsStreaming] = useState(false);
  const [attachFile, setAttachFile] = useState(false);
  const [attachSelection, setAttachSelection] = useState(false);
  const [selectionText, setSelectionText] = useState('');
  const [copiedId, setCopiedId] = useState('');
  const [showHistory, setShowHistory] = useState(false);

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

const MessageBubble: React.FC<{
  message: ChatMessage;
  isCopied: boolean;
  onCopy: () => void;
  onInsert: (text: string) => void;
  onReplace: (text: string) => void;
  hasActiveTab: boolean;
}> = ({ message, isCopied, onCopy, onInsert, onReplace, hasActiveTab }) => {
  const isUser = message.role === 'user';
  const codeBlocks = useMemo(
    () => (isUser ? [] : extractCodeBlocks(message.content)),
    [isUser, message.content]
  );

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
          {message.content ? (
            <MarkdownLite content={message.content} />
          ) : (
            <span className="flex items-center gap-1.5 text-xs text-ws-faint">
              <span className="h-1.5 w-1.5 animate-pulse-soft rounded-full bg-ws-accent" />
              Думает…
            </span>
          )}
        </div>

        {message.error && (
          <div className="mt-1.5 flex items-start gap-1.5 rounded border border-red-500/30 bg-red-500/10 px-2 py-1.5 text-[11px] text-ws-danger">
            <AlertTriangle size={11} className="mt-0.5 shrink-0" />
            <span className="select-text">{message.error}</span>
          </div>
        )}

        {/* Действия над ответом. Всегда видимы с пониженной непрозрачностью —
            hover делает их контрастнее, но они остаются доступными
            (в т.ч. при управлении с клавиатуры). */}
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
                onClick={() => onReplace(message.content)}
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
