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
  Settings,
  PanelLeft,
  PanelBottom,
  FolderPlus,
  Pencil,
  Save,
  Terminal,
  Bell,
  MousePointerClick,
  FolderInput,
  ShieldCheck,
  ShieldAlert,
} from 'lucide-react';
import { AISettings, ChatAttachment, ChatMessage, ChatSession, OpenTab } from '../../types';
import { aiService } from '../../services/aiService';
import { SYSTEM_PRESETS, PROVIDER_PRESETS } from '../../state/settings';
import { cn, extractCodeBlocks, uid } from '../../lib/utils';
import { renderRichMarkdown, renderMermaidBlocks } from '../../lib/markdown';
import { stripAssistantNoise, stripEmoji } from '../../lib/aiDocument';
import {
  AIActionPayload,
  AppActionKind,
  AppActionResult,
  ParsedMarkdownTable,
  parseAIActions,
  detectMarkdownTables,
  requiresConfirmation,
  stripActionBlocks,
} from '../../services/aiActions';
import { AIBadge, Button, IconButton, Spinner, Textarea } from '../ui';

/** Действия, меняющие содержимое открытого документа. */
const DOC_WRITE_ACTIONS = new Set<string>(['insert_into_editor', 'replace_document']);

/** Просьба поработать с текстом документа, а не ответить на вопрос. */
const EDIT_INTENT =
  /(исправ|перепиш|перепис|оформ|сократ|приведи|сделай|структурир|отредакт|редактир|улучш|вычит|перевод|переведи|дополн|добавь|убери|замени|измени|обнови|сформат|выровняй)/i;

/**
 * Убирает из ответа вступительную фразу-комментарий («Вот вариант текста…»),
 * чтобы она не попала в документ.
 */
function buildDocumentBody(answer: string): string {
  const lines = stripAssistantNoise(answer).split('\n');
  const first = lines[0]?.trim() ?? '';
  if (
    lines.length > 1 &&
    first.length < 160 &&
    /^(вот|ниже|готово|конечно|разумеется|держите|представляю)\s/i.test(first)
  ) {
    lines.shift();
    while (lines.length > 0 && !lines[0].trim()) lines.shift();
  }
  return lines.join('\n').trim();
}

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
  /** Единая точка выполнения действий AI — полное управление приложением. */
  onRunAction: (action: AIActionPayload) => Promise<AppActionResult>;
  /** Актуальное состояние приложения для системного промта. */
  appContext?: string;
  /** Читает файл, упомянутый в сообщении, чтобы модель увидела его содержимое. */
  onResolveAttachment?: (
    name: string
  ) => Promise<{ name: string; path: string; content: string } | null>;
  /** Читает файлы, перетащенные из Проводника. */
  onReadDroppedFiles?: (
    paths: string[]
  ) => Promise<{ name: string; path: string; content: string; truncated?: boolean }[]>;
  /** Текст активного документа: Word-вкладка хранит base64, его нужно разобрать. */
  onResolveActiveDocument?: () => Promise<ChatAttachment | null>;
  /** Выданы ли ассистенту права на запись в открытый документ. */
  docWriteAccess: boolean;
  onToggleDocWriteAccess: (value: boolean) => void;
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
  onRunAction,
  appContext,
  onResolveAttachment,
  onReadDroppedFiles,
  onResolveActiveDocument,
  docWriteAccess,
  onToggleDocWriteAccess,
}) => {
  const [input, setInput] = useState('');
  const [isStreaming, setIsStreaming] = useState(false);
  const [attachFile, setAttachFile] = useState(true);
  const [attachSelection, setAttachSelection] = useState(false);
  const [selectionText, setSelectionText] = useState('');
  const [copiedId, setCopiedId] = useState('');
  const [showHistory, setShowHistory] = useState(false);
  /** Файлы, перетащенные из Проводника в поле чата. */
  const [droppedFiles, setDroppedFiles] = useState<
    { name: string; path: string; content: string; truncated?: boolean }[]
  >([]);
  const [dragActive, setDragActive] = useState(false);

  // Состояние выполнения действий AI (создание таблиц, файлов и т.п.)
  const [actionStates, setActionStates] = useState<
    Record<
      string,
      { status: 'idle' | 'running' | 'done' | 'error'; result?: string; error?: string; path?: string }
    >
  >({});
  const executedActionsRef = useRef<Set<string>>(new Set());
  /** Ответы, уже применённые к документу как готовый текст. */
  const autoAppliedRef = useRef<Set<string>>(new Set());
  /** Действия из истории до открытия панели повторно не выполняются. */
  const mountedAtRef = useRef(Date.now());
  const autoRunBusyRef = useRef(false);
  const [actionTick, setActionTick] = useState(0);

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
    async (
      rawText: string,
      attached?: ChatMessage['attachedFile'],
      dropped?: ChatMessage['attachedFiles']
    ) => {
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
        attachedFiles: dropped && dropped.length > 0 ? dropped : undefined,
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

      // Если в сообщении назван файл — читаем его и прикладываем к запросу,
      // чтобы модель работала с реальным содержимым, а не просила его вставить.
      let apiContent = text;
      if (onResolveAttachment) {
        const matches = Array.from(
          new Set(
            (text.match(/[A-Za-zА-Яа-я0-9_@()\-./\\]+\.(?:md|markdown|xlsx|xlsm|csv|tsv|txt|log|json|xml|yaml|yml|ini|cfg|docx?|pptx?|pdf)/gi) ?? [])
              .map((value) => value.replace(/[),.;:!?]+$/, ''))
              .filter(Boolean)
          )
        ).slice(0, 3);
        const files: { name: string; path: string; content: string }[] = [];
        for (const name of matches) {
          const file = await onResolveAttachment(name);
          if (file) files.push(file);
        }
        if (files.length > 0) {
          let budget = 60000;
          const blocks: string[] = [];
          for (const file of files) {
            if (budget <= 0) break;
            const content = file.content.slice(0, budget);
            budget -= content.length;
            blocks.push(`[Файл: ${file.name}]\n${content}`);
          }
          apiContent += '\n\n' + blocks.join('\n\n');
          apiContent =
            '[Инструкция WorkSpace: содержимое файлов уже вложено ниже. Используй его. Если просят перенести данные в таблицу — верни import_table с source и target, не переписывай строки вручную. Не отвечай, что у тебя нет доступа к файлам.]\n\n' +
            apiContent;
        }
      }
      if (dropped && dropped.length > 0) {
        let budget = 60000;
        const blocks: string[] = [];
        for (const file of dropped) {
          if (budget <= 0) break;
          const content = file.content.slice(0, budget);
          budget -= content.length;
          blocks.push(`[Файл: ${file.name}]\n${content}`);
        }
        apiContent += '\n\n' + blocks.join('\n\n');
        apiContent =
          '[Инструкция WorkSpace: пользователь перетащил файлы в чат, их содержимое ниже. Если просят перенести данные в таблицу — верни import_table с source и target, не переписывай строки вручную. Не отвечай, что нет доступа к файлам.]\n\n' +
          apiContent;
      }
      const apiUserMessage: ChatMessage = { ...userMessage, content: apiContent };

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
        // Переписывание целого документа не влезает в 4096 токенов:
        // ответ обрывается на середине и в документ попадает половина текста.
        const requestSettings = EDIT_INTENT.test(text)
          ? { ...settings, maxTokens: Math.max(settings.maxTokens, 8192) }
          : settings;

        await aiService.streamChat(
          [...base, apiUserMessage],
          requestSettings,
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
          controller.signal,
          appContext
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
      appContext,
      onResolveAttachment,
    ]
  );

  /* ---------- Выполнение действий управления программой ---------- */

  const executeAction = useCallback(
    async (action: AIActionPayload, actionKey: string): Promise<void> => {
      setActionStates((prev) => ({
        ...prev,
        [actionKey]: { status: 'running' },
      }));

      try {
        const result = await onRunAction(action);
        setActionStates((prev) => ({
          ...prev,
          [actionKey]: {
            status: result.ok ? 'done' : 'error',
            result: result.message,
            path: result.path,
            error: result.ok ? undefined : result.message,
          },
        }));
      } catch (err) {
        setActionStates((prev) => ({
          ...prev,
          [actionKey]: { status: 'error', error: (err as Error).message },
        }));
      }
    },
    [onRunAction]
  );

  // Автоматическое выполнение безопасных действий из новых ответов модели.
  // Действия из истории и разрушительные операции ждут явного клика.
  useEffect(() => {
    if (autoRunBusyRef.current) return;

    const pending: { action: AIActionPayload; key: string }[] = [];
    for (const msg of messages) {
      if (msg.role !== 'assistant' || msg.streaming) continue;
      if (msg.timestamp < mountedAtRef.current) continue;
      parseAIActions(msg.content).forEach((act, idx) => {
        const key = `${msg.id}_${idx}`;
        if (executedActionsRef.current.has(key)) return;
        if (requiresConfirmation(act)) return;
        if (DOC_WRITE_ACTIONS.has(act.action) && !docWriteAccess) return;
        pending.push({ action: act, key });
      });
    }

    if (pending.length === 0) return;

    autoRunBusyRef.current = true;
    void (async () => {
      for (const item of pending) {
        if (executedActionsRef.current.has(item.key)) continue;
        executedActionsRef.current.add(item.key);
        await executeAction(item.action, item.key);
      }
    })().finally(() => {
      autoRunBusyRef.current = false;
      setActionTick((v) => v + 1);
    });
  }, [messages, executeAction, actionTick, docWriteAccess]);

  /**
   * Ответ без блока действия, но с готовым текстом документа.
   *
   * Модели часто игнорируют протокол и присылают переписанный документ прямо
   * в чат. Тогда применяем ответ к открытому документу сами — иначе правка
   * остаётся только сообщением, и «в документе ничего не меняется».
   */
  useEffect(() => {
    if (!docWriteAccess) return;
    const tab = activeTab;
    if (!tab) return;
    if (!['word', 'markdown', 'richtext', 'text'].includes(tab.type)) return;

    for (let index = 0; index < messages.length; index++) {
      const msg = messages[index];
      if (msg.role !== 'assistant' || msg.streaming) continue;
      if (msg.timestamp < mountedAtRef.current) continue;
      if (autoAppliedRef.current.has(msg.id)) continue;
      if (parseAIActions(msg.content).length > 0) continue;

      const answer = stripActionBlocks(msg.content).trim();
      if (answer.length < 300) continue;

      const request = [...messages.slice(0, index)].reverse().find((item) => item.role === 'user');
      if (!request || !EDIT_INTENT.test(request.content)) continue;
      // Вопрос или отказ модели — не повод перезаписывать документ.
      if (/не могу|не имеет смысла|уточните|какой именно|пришлите/i.test(answer.slice(0, 300))) continue;

      autoAppliedRef.current.add(msg.id);
      void onApplyToDocument(buildDocumentBody(answer));
      break;
    }
  }, [messages, docWriteAccess, activeTab, onApplyToDocument]);

  const handleCreateTableFromMarkdown = useCallback(
    async (table: ParsedMarkdownTable, format: 'xlsx' | 'csv' = 'xlsx') => {
      const name = format === 'csv' ? 'Таблица_данных.csv' : 'Таблица_данных.xlsx';
      await onRunAction({
        action: 'create_table',
        name,
        columns: table.columns,
        rows: table.rows,
        sheetName: 'Лист 1',
        format,
      });
    },
    [onRunAction]
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
  const buildAttachment = async (): Promise<ChatMessage['attachedFile'] | undefined> => {
    if (!activeTab || !attachFile) return undefined;
    if (attachSelection && selectionText) {
      return {
        name: `${activeTab.name} (выделенное)`,
        path: activeTab.path,
        content: selectionText,
        truncated: selectionText.length > 40000,
      };
    }
    // Word/Excel/PowerPoint хранят содержимое вкладки в base64 — модели нужен
    // извлечённый текст, иначе она получает ZIP-пакет и «не видит» файл.
    if (!['markdown', 'richtext', 'text', 'code'].includes(activeTab.type)) {
      return (await onResolveActiveDocument?.()) ?? undefined;
    }
    return {
      name: activeTab.name,
      path: activeTab.path,
      content: activeTab.content,
      truncated: activeTab.content.length > 40000,
    };
  };

  const handleSend = () => {
    const text = input.trim() || (droppedFiles.length > 0 ? 'Проанализируй приложенные файлы' : '');
    if (!text) return;
    const dropped = droppedFiles.length > 0 ? droppedFiles : undefined;
    setDroppedFiles([]);
    void (async () => {
      const attachment = await buildAttachment();
      await submit(text, attachment, dropped);
    })();
  };

  /* ---------- Перетаскивание файлов из Проводника ---------- */
  const handleDragOver = (event: React.DragEvent) => {
    const types = Array.from(event.dataTransfer.types);
    if (types.includes('Files') || types.includes('application/x-workspace-path')) {
      event.preventDefault();
      event.stopPropagation();
      event.dataTransfer.dropEffect = 'copy';
      setDragActive(true);
    }
  };

  const handleDragLeave = (event: React.DragEvent) => {
    if (event.currentTarget === event.target) setDragActive(false);
  };

  const handleDrop = async (event: React.DragEvent) => {
    event.preventDefault();
    event.stopPropagation();
    setDragActive(false);
    if (!onReadDroppedFiles) return;
    const files = Array.from(event.dataTransfer.files || []);
    const internalPath = event.dataTransfer.getData('application/x-workspace-path');
    const paths = [
      ...(internalPath ? [internalPath] : []),
      ...files
        .map((file) => {
          let fromElectron = '';
          try {
            fromElectron = window.electronAPI?.getPathForFile?.(file) || '';
          } catch {
            fromElectron = '';
          }
          return fromElectron || (file as File & { path?: string }).path || '';
        })
        .filter(Boolean),
    ];
    const uniquePaths = [...new Set(paths)];
    if (uniquePaths.length === 0) return;
    const attachments = await onReadDroppedFiles(uniquePaths);
    if (attachments.length > 0) {
      setDroppedFiles((prev) => [...prev, ...attachments].slice(0, 5));
    }
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
    <div
      className="relative flex h-full w-full flex-col border-l border-ws bg-ws-panel"
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={(event) => void handleDrop(event)}
    >
      {dragActive && (
        <div className="pointer-events-none absolute inset-0 z-50 flex flex-col items-center justify-center gap-2 rounded border-2 border-dashed border-ws-accent bg-ws-panel/90 text-xs text-ws-accent">
          <Paperclip size={22} />
          Отпустите файлы — они прикрепятся к сообщению
        </div>
      )}
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
            onInsert={(text) => {
              if (!docWriteAccess) onToggleDocWriteAccess(true);
              onInsertIntoEditor(text);
            }}
            onReplace={(text) => {
              if (!docWriteAccess) onToggleDocWriteAccess(true);
              onApplyToDocument(text);
            }}
            hasActiveTab={Boolean(activeTab)}
            actionStates={actionStates}
            onExecuteAction={(action, key) => {
              // Ручной запуск действия — это и есть согласие на запись:
              // фиксируем ключ, чтобы автозапуск не выполнил его повторно.
              executedActionsRef.current.add(key);
              if (DOC_WRITE_ACTIONS.has(action.action) && !docWriteAccess) {
                onToggleDocWriteAccess(true);
              }
              void executeAction(action, key);
            }}
            onOpenFile={(path) => {
              void onRunAction({ action: 'open_file', path });
            }}
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
              title="Содержимое открытого файла прикладывается к запросу. Нажмите, чтобы отключить."
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

            {(activeTab.type === 'word' ||
              activeTab.type === 'markdown' ||
              activeTab.type === 'richtext' ||
              activeTab.type === 'text') && (
              <button
                onClick={() => onToggleDocWriteAccess(!docWriteAccess)}
                title={
                  docWriteAccess
                    ? 'Ассистент правит открытый документ сам. Нажмите, чтобы запретить запись.'
                    : 'Запись запрещена. Нажмите, чтобы разрешить ассистенту изменять документ.'
                }
                className={cn(
                  'flex items-center gap-1 rounded border px-1.5 py-0.5 text-[10px] transition-colors',
                  docWriteAccess
                    ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-400'
                    : 'border-amber-500/40 bg-amber-500/10 text-ws-warning'
                )}
              >
                {docWriteAccess ? <ShieldCheck size={10} /> : <ShieldAlert size={10} />}
                {docWriteAccess ? 'Права записи выданы' : 'Предоставить права на запись'}
              </button>
            )}
          </div>
        )}

        <div className="relative rounded-lg border border-ws bg-ws-elevated focus-within:border-ws-accent">
          {droppedFiles.length > 0 && (
            <div className="flex flex-wrap gap-1 border-b border-ws/60 px-2 pt-2 pb-1">
              {droppedFiles.map((file, index) => (
                <span
                  key={`${file.path}_${index}`}
                  className="flex items-center gap-1 rounded border border-ws-accent/40 bg-ws-accent-soft px-1.5 py-0.5 text-[10px] text-ws-accent"
                  title={file.path}
                >
                  <Paperclip size={9} />
                  <span className="max-w-[180px] truncate">{file.name}</span>
                  {file.truncated && <span className="text-ws-warning">…</span>}
                  <button
                    type="button"
                    title="Убрать файл"
                    onClick={() =>
                      setDroppedFiles((prev) => prev.filter((_, item) => item !== index))
                    }
                    className="rounded p-0.5 hover:bg-ws-hover"
                  >
                    <X size={9} />
                  </button>
                </span>
              ))}
            </div>
          )}
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
                disabled={!input.trim() && droppedFiles.length === 0}
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

const ACTION_META: Record<
  AppActionKind,
  { label: string; icon: React.ReactNode; box: string; iconBox: string }
> = {
  create_table: {
    label: 'Создать таблицу',
    icon: <FileSpreadsheet size={15} />,
    box: 'border-emerald-500/30 bg-emerald-500/10',
    iconBox: 'bg-emerald-500/20 text-emerald-400',
  },
  import_table: {
    label: 'Импортировать таблицу из файла',
    icon: <FileSpreadsheet size={15} />,
    box: 'border-emerald-500/30 bg-emerald-500/10',
    iconBox: 'bg-emerald-500/20 text-emerald-400',
  },
  edit_table: {
    label: 'Изменить таблицу',
    icon: <FileSpreadsheet size={15} />,
    box: 'border-emerald-500/30 bg-emerald-500/10',
    iconBox: 'bg-emerald-500/20 text-emerald-400',
  },
  create_document: {    label: 'Создать документ Word',
    icon: <FileText size={15} />,
    box: 'border-blue-500/30 bg-blue-500/10',
    iconBox: 'bg-blue-500/20 text-blue-400',
  },
  create_file: {
    label: 'Создать файл',
    icon: <FilePlus size={15} />,
    box: 'border-blue-500/30 bg-blue-500/10',
    iconBox: 'bg-blue-500/20 text-blue-400',
  },
  open_file: {
    label: 'Открыть файл',
    icon: <FolderOpen size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-amber-500/15 text-amber-400',
  },
  open_folder: {
    label: 'Открыть рабочую папку',
    icon: <FolderInput size={15} />,
    box: 'border-amber-500/30 bg-amber-500/10',
    iconBox: 'bg-amber-500/20 text-amber-400',
  },
  create_folder: {
    label: 'Создать папку',
    icon: <FolderPlus size={15} />,
    box: 'border-amber-500/30 bg-amber-500/10',
    iconBox: 'bg-amber-500/20 text-amber-400',
  },
  rename_path: {
    label: 'Переименовать',
    icon: <Pencil size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-sky-500/15 text-sky-400',
  },
  delete_path: {
    label: 'Удалить',
    icon: <Trash2 size={15} />,
    box: 'border-red-500/30 bg-red-500/10',
    iconBox: 'bg-red-500/20 text-red-400',
  },
  insert_into_editor: {
    label: 'Вставить в документ',
    icon: <CornerDownLeft size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-ws-accent-soft text-ws-accent',
  },
  replace_document: {
    label: 'Заменить документ',
    icon: <FileText size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-ws-accent-soft text-ws-accent',
  },
  open_settings: {
    label: 'Открыть настройки',
    icon: <Settings size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-ws-accent-soft text-ws-accent',
  },
  toggle_panel: {
    label: 'Переключить панель',
    icon: <PanelLeft size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-ws-accent-soft text-ws-accent',
  },
  select_tab: {
    label: 'Переключить вкладку',
    icon: <MousePointerClick size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-ws-accent-soft text-ws-accent',
  },
  close_tab: {
    label: 'Закрыть вкладку',
    icon: <X size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-ws-accent-soft text-ws-accent',
  },
  save: {
    label: 'Сохранить',
    icon: <Save size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-emerald-500/15 text-emerald-400',
  },
  notify: {
    label: 'Уведомление',
    icon: <Bell size={15} />,
    box: 'border-ws bg-ws-surface',
    iconBox: 'bg-ws-accent-soft text-ws-accent',
  },
  run_command: {
    label: 'Выполнить в терминале',
    icon: <Terminal size={15} />,
    box: 'border-amber-500/30 bg-amber-500/10',
    iconBox: 'bg-amber-500/20 text-amber-400',
  },
};

function describeAction(action: AIActionPayload): string {
  switch (action.action) {
    case 'create_table': {
      const cols = action.columns?.length || 0;
      const rows = action.rows?.length || 0;
      return `${action.name || 'Таблица.xlsx'}${cols || rows ? ` · ${cols} × ${rows}` : ''}`;
    }
    case 'create_document':
      return action.name || 'Документ.docx';
    case 'edit_table':
      return `${action.name || action.path || 'Активная таблица'} · ${
        action.mode === 'append' ? 'добавить' : 'заменить'
      }${action.rows?.length ? ` · ${action.rows.length} строк` : ''}`;
    case 'import_table':
      return `${action.source} → ${action.target || 'новая таблица'}`;
    case 'create_file':
      return [action.path, action.name].filter(Boolean).join('/') || 'Новый файл';
    case 'open_file':
      return action.path || 'Файл';
    case 'open_folder':
      return action.path || 'Выбор папки в диалоге';
    case 'create_folder':
      return action.path || action.name || 'Новая папка';
    case 'rename_path':
      return `${action.path} → ${action.newName || action.newPath}`;
    case 'delete_path':
      return action.path;
    case 'insert_into_editor':
    case 'replace_document':
      return `${action.text.length} символов`;
    case 'open_settings':
      return `Вкладка: ${action.tab || 'provider'}`;
    case 'toggle_panel':
      return `${action.panel} · ${action.state || 'toggle'}`;
    case 'select_tab':
      return action.path || action.name || `Вкладка #${(action.index ?? 0) + 1}`;
    case 'close_tab':
      return action.path || action.name || 'Текущая вкладка';
    case 'save':
      return action.all ? 'Все изменённые файлы' : 'Текущий файл';
    case 'notify':
      return action.message;
    case 'run_command':
      return action.command;
    default:
      return '';
  }
}

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

  if (
    action.action === 'create_document' ||
    (action.action === 'create_file' && /\.docx?$/i.test(action.name || action.path || ''))
  ) {
    const rawName =
      action.name ||
      ('title' in action && action.title ? `${action.title}.docx` : '') ||
      ('path' in action && action.path ? action.path : '') ||
      'Документ.docx';
    const fileName = rawName.toLowerCase().endsWith('.docx') ? rawName : `${rawName}.docx`;
    const contentLen = action.content?.length || 0;

    return (
      <div className="mt-2 rounded-lg border border-blue-500/30 bg-blue-500/10 p-2.5 text-xs">
        <div className="flex items-start justify-between gap-2">
          <div className="flex items-start gap-2 min-w-0">
            <div className="mt-0.5 rounded bg-blue-500/20 p-1 text-blue-400 shrink-0">
              <FileText size={15} />
            </div>
            <div className="min-w-0">
              <div className="font-semibold text-ws truncate">
                Документ Word: <span className="text-blue-400 font-mono">{fileName}</span>
              </div>
              <div className="text-[11px] text-ws-muted mt-0.5">
                {contentLen > 0 ? `${contentLen} символов` : 'Новый документ'} · Word (.docx)
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
                className="flex items-center gap-1 rounded bg-blue-600 px-2 py-1 text-[11px] font-medium text-white hover:bg-blue-500 shadow-sm transition-all"
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

  const meta = ACTION_META[action.action] ?? ACTION_META.notify;
  const target = describeAction(action);
  const needsConfirm = requiresConfirmation(action);
  const buttonLabel = needsConfirm
    ? action.action === 'delete_path'
      ? 'Удалить'
      : 'Выполнить'
    : action.action === 'create_folder'
      ? 'Создать'
      : 'Выполнить';

  return (
    <div className={cn('mt-2 rounded-lg border p-2.5 text-xs', meta.box)}>
      <div className="flex items-start justify-between gap-2">
        <div className="flex min-w-0 items-start gap-2">
          <div className={cn('mt-0.5 shrink-0 rounded p-1', meta.iconBox)}>{meta.icon}</div>
          <div className="min-w-0">
            <div className="truncate font-semibold text-ws">{meta.label}</div>
            {target && (
              <div className="mt-0.5 truncate font-mono text-[11px] text-ws-muted" title={target}>
                {target}
              </div>
            )}
            {status === 'done' && result && (
              <div className="mt-0.5 line-clamp-2 text-[10px] text-emerald-400">{result}</div>
            )}
            {status === 'error' && (
              <div className="mt-0.5 text-[10px] text-ws-danger" title={error}>
                {error || 'Не удалось выполнить'}
              </div>
            )}
            {status === 'idle' && needsConfirm && (
              <div className="mt-0.5 text-[10px] text-ws-warning">
                Требуется подтверждение — проверьте действие перед запуском.
              </div>
            )}
          </div>
        </div>

        <div className="shrink-0">
          {status === 'running' ? (
            <span className="flex items-center gap-1.5 text-[11px] text-ws-accent">
              <Spinner size={12} />
              Выполнение…
            </span>
          ) : status === 'done' ? (
            <div className="flex items-center gap-1.5">
              <span className="flex items-center gap-1 text-[11px] font-medium text-emerald-400">
                <Check size={12} />
                Готово
              </span>
              {onOpen && (
                <button
                  onClick={onOpen}
                  className="flex items-center gap-1 rounded border border-ws bg-ws-surface px-1.5 py-0.5 text-[10px] text-ws transition-colors hover:bg-ws-hover"
                >
                  <ExternalLink size={10} />
                  Открыть
                </button>
              )}
            </div>
          ) : status === 'error' ? (
            <button
              onClick={onExecute}
              className="rounded border border-ws bg-ws-surface px-2 py-0.5 text-[10px] text-ws hover:bg-ws-hover"
            >
              Повторить
            </button>
          ) : (
            <button
              onClick={onExecute}
              className={cn(
                'flex items-center gap-1 rounded px-2 py-1 text-[11px] font-medium text-white shadow-sm transition-all',
                needsConfirm ? 'bg-red-600 hover:bg-red-500' : 'bg-emerald-600 hover:bg-emerald-500'
              )}
            >
              <Play size={10} />
              {buttonLabel}
            </button>
          )}
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
  actionStates: Record<
    string,
    { status: 'idle' | 'running' | 'done' | 'error'; result?: string; error?: string; path?: string }
  >;
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
    const cleaned = stripEmoji(stripActionBlocks(message.content));
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
          {message.attachedFiles?.map((file, index) => (
            <div
              key={`${file.path}_${index}`}
              className="mb-1 flex items-center gap-1 rounded bg-white/15 px-1.5 py-0.5 text-[10px]"
            >
              <Paperclip size={9} />
              <span className="truncate">{file.name}</span>
              {file.truncated && <span className="opacity-80">(обрезано)</span>}
            </div>
          ))}
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
              state.path ||
              (act.action === 'create_table'
                ? act.name
                : act.action === 'create_document'
                  ? act.name || (act.title ? `${act.title}.docx` : 'Документ.docx')
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

            {codeBlocks.length === 0 && displayContent && hasActiveTab && (
              <button
                onClick={() => onInsert(displayContent)}
                title="Вставить текст ответа в позицию курсора активного документа"
                className="flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
              >
                <CornerDownLeft size={10} />
                Вставить в документ
              </button>
            )}

            {hasActiveTab && displayContent && (
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
