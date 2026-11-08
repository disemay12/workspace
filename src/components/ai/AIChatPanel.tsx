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
import { aiService, ToolsUnsupportedError } from '../../services/aiService';
import { buildDocumentSystemPrompt, buildDocumentTools, supportsNativeDocumentTools } from '../../services/aiTools';
import { AGENT_TAB_TYPES, DOCUMENT_TAB_TYPES, type DocumentToolResult } from '../../lib/documentModel';
import { SYSTEM_PRESETS, PROVIDER_PRESETS } from '../../state/settings';
import { cn, extractCodeBlocks, uid } from '../../lib/utils';
import { renderRichMarkdown, renderMermaidBlocks } from '../../lib/markdown';
import { extractDocumentContent, stripEmoji } from '../../lib/aiDocument';
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
const DOC_WRITE_ACTIONS = new Set<string>([
  'insert_into_editor',
  'replace_document',
  'document_apply',
]);

/** Просьба поработать с текстом документа, а не ответить на вопрос. */
const EDIT_INTENT =
  /(исправ|перепиш|перепис|оформ|сократ|приведи|сделай|структурир|отредакт|редактир|улучш|вычит|перевод|переведи|дополн|добавь|убери|замени|измени|обнови|сформат|выровняй|напиши|поменяй|переименуй|допиши)/i;

/** Создание или дописывание текста в документ. */
const CREATE_INTENT =
  /(напиши|написать|составь|создай|сгенерируй|добавь|вставь|допиши|запиши|сочини|придумай)/i;

/** Точечная правка: меняется конкретная строка/заголовок/фраза. */
const PRECISE_EDIT_INTENT =
  /(замени|поменяй|исправь|измени|поправь|переименуй|убери|удали)/i;

const EDIT_TARGET_INTENT =
  /(назван|заголов|подзаголов|строк|слов|фраз|дат|имя|имени|пункт|абзац|текст|номер|сумм|адрес|телефон)/i;

/** Полная переработка документа. */
const REWRITE_INTENT =
  /(перепиш|переписать|сократ|улучш|приведи|оформ|сделай.{0,30}документ|полностью|целиком|ats|atc|атс|фильтр)/i;

/** Резюме под ATS/ATC-фильтры. */
const ATS_INTENT = /(ats|atc|атс|проход\w*\s+фильтр|фильтр\w*\s+(ats|atc|атс))/i;

const PRECISE_REPLACE_ERROR =
  'Точечная правка: используй find_replace или replace_block. replace_document запрещён, пока меняется одна строка/заголовок.';

function isPreciseEditRequest(text: string): boolean {
  const hasVerb = PRECISE_EDIT_INTENT.test(text);
  const quoted = /[«"“][^»"”]+[»"”]/.test(text);
  return hasVerb && (EDIT_TARGET_INTENT.test(text) || quoted);
}

function isCreateRequest(text: string): boolean {
  return (
    CREATE_INTENT.test(text) &&
    !isPreciseEditRequest(text) &&
    !/(перепиш|переписать|сократ)/i.test(text)
  );
}

/** Действие целиком переписывает открытый документ. */
function actionReplacesWholeDocument(action: AIActionPayload): boolean {
  if (action.action === 'replace_document') return true;
  if (action.action === 'document_apply') {
    return action.operations.some((operation) => operation.tool === 'replace_document');
  }
  return false;
}

/** Есть действие документа, допустимое для текущего намерения. */
function hasAllowedDocumentAction(answer: string, precise: boolean): boolean {
  return parseAIActions(answer).some(
    (action) =>
      DOC_WRITE_ACTIONS.has(action.action) &&
      (!precise || !actionReplacesWholeDocument(action))
  );
}

/** Новый заголовок из просьбы «замени название на «…»». */
function extractRequestedTitle(text: string): string | null {
  const patterns = [
    /(?:название|заголовок|подзаголовок)[^«"“\n]{0,40}[«"“]([^»"”]+)[»"”]/i,
  ];
  for (const pattern of patterns) {
    const match = pattern.exec(text);
    const value = match?.[1]?.trim();
    if (value) return value;
  }
  return null;
}

/** Первый заголовок из структуры get_document. */
function firstHeading(structure: string): string | null {
  for (const line of structure.split('\n')) {
    const match = /^\s*b\d+\s+\[(h[1-6])\]\s*(.+)$/.exec(line);
    if (match) return match[2].trim();
  }
  return null;
}

/**
 * Мгновенные точные правки без обращения к модели: замена заголовка и
 * «Замени «X» на «Y»». Модель для этого не нужна, а ожидание ответа —
 * главная причина долгого «Думает».
 */
async function applyQuickDocumentEdit(
  text: string,
  tabId: string,
  runTool: (
    name: string,
    args: Record<string, unknown>,
    transactionId: string,
    tabId?: string
  ) => Promise<DocumentToolResult>
): Promise<string | null> {
  const transactionId = `quick_${Date.now()}`;

  const quotedReplace =
    /(?:замени|поменяй|исправь)\s+[«"“]([^»"”]+)[»"”]\s+на\s+[«"“]([^»"”]+)[»"”]/i.exec(
      text
    );
  if (quotedReplace) {
    const result = await runTool(
      'find_replace',
      { find: quotedReplace[1], replace: quotedReplace[2], all: false },
      transactionId,
      tabId
    );
    return result.ok ? `Заменил «${quotedReplace[1]}» на «${quotedReplace[2]}».` : null;
  }

  const title = extractRequestedTitle(text);
  if (!title || !/(заголов|назван|подзаголов)/i.test(text)) return null;

  const snapshot = await runTool('get_document', {}, transactionId, tabId);
  if (!snapshot.ok) return null;
  const current = firstHeading(snapshot.structure ?? '');
  if (!current || current === title) return null;

  const result = await runTool(
    'find_replace',
    { find: current, replace: title, all: false },
    transactionId,
    tabId
  );
  return result.ok ? `Заменил заголовок на «${title}».` : null;
}

/**
 * Убирает из ответа вступительную фразу-комментарий («Вот вариант текста…»),
 * служебные советы по оформлению и обёртку ```, чтобы в документ попал
 * только его текст.
 */
function buildDocumentBody(answer: string): string {
  return extractDocumentContent(answer);
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
  /** Исполнитель инструментов открытого документа (function calling). */
  onDocumentTool: (
    name: string,
    args: Record<string, unknown>,
    transactionId: string,
    tabId?: string
  ) => Promise<DocumentToolResult>;
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
  /** Разрешена ли запись в дополнительные файлы workspace. */
  extraWriteAccess: boolean;
  onToggleExtraWriteAccess: (value: boolean) => void;
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
  onDocumentTool,
  appContext,
  onResolveAttachment,
  onReadDroppedFiles,
  onResolveActiveDocument,
  docWriteAccess,
  onToggleDocWriteAccess,
  extraWriteAccess,
  onToggleExtraWriteAccess,
}) => {
  const [input, setInput] = useState('');
  const [isStreaming, setIsStreaming] = useState(false);
  /** Что делает агент прямо сейчас («Читаю документ…», «Правлю документ…»). */
  const [agentStatus, setAgentStatus] = useState('');
  const [toolTrace, setToolTrace] = useState<string[]>([]);
  const toolTraceRef = useRef<string[]>([]);
  const [contextFiles, setContextFiles] = useState<string[]>([]);
  const contextFilesRef = useRef<string[]>([]);
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

  /** Обёртка инструментов: собирает видимый след вызовов агента. */
  const callDocumentTool = useCallback(
    async (
      name: string,
      args: Record<string, unknown>,
      transactionId: string,
      tabId?: string
    ) => {
      if (!toolTraceRef.current.includes(name)) {
        toolTraceRef.current = [...toolTraceRef.current, name];
        setToolTrace(toolTraceRef.current);
      }
      if (name === 'read_document') {
        const rawPath = String(args.path ?? args.file ?? '').trim();
        const fileName = rawPath.split(/[\\/]/).filter(Boolean).pop() ?? '';
        if (fileName && !contextFilesRef.current.includes(fileName)) {
          contextFilesRef.current = [...contextFilesRef.current, fileName];
          setContextFiles(contextFilesRef.current);
        }
      }
      return onDocumentTool(name, args, transactionId, tabId);
    },
    [onDocumentTool]
  );

  const submit = useCallback(
    async (
      rawText: string,
      attached?: ChatMessage['attachedFile'],
      dropped?: ChatMessage['attachedFiles'],
      mode: 'chat' | 'inline' = 'chat'
    ) => {
      const text = rawText.trim();
      if (!text || isStreaming) return;

      toolTraceRef.current = [];
      setToolTrace([]);
      contextFilesRef.current = [];
      setContextFiles([]);

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

      // Точные замены заголовка/цитаты выполняем локально, без запроса к
      // модели: это мгновенно и гарантирует, что изменится только указанный
      // фрагмент.
      if (
        mode === 'chat' &&
        activeTab &&
        DOCUMENT_TAB_TYPES.has(activeTab.type) &&
        docWriteAccess
      ) {
        const quick = await applyQuickDocumentEdit(text, activeTab.id, callDocumentTool);
        if (quick) {
          onUpdateMessages(sessionId, [
            ...base,
            userMessage,
            {
              id: assistantId,
              role: 'assistant',
              content: quick,
              timestamp: Date.now(),
              toolsUsed: true,
              preciseEdit: true,
              toolTrace: [...toolTraceRef.current],
              contextFiles: [...contextFilesRef.current],
            },
          ]);
          setInput('');
          return;
        }
      }

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

      const useDocumentAgent =
        mode === 'chat' && Boolean(activeTab && AGENT_TAB_TYPES.has(activeTab.type));
      const transactionId = uid('doctx');

      // Переписывание целого документа не влезает в 4096 токенов:
      // ответ обрывается на середине, и в документ попадает половина текста.
      const requestSettings =
        EDIT_INTENT.test(text) || useDocumentAgent
          ? { ...settings, maxTokens: Math.max(settings.maxTokens, 8192) }
          : settings;

      const streamCallbacks = {
        onChunk: (partial: string) => {
          buffer = partial;
          paint();
        },
        onFinish: (full: string) => {
          buffer = full;
          paint(true);
        },
        onError: (error: Error) => {
          throw error;
        },
      };

      /** Документ уже правлен инструментами: текст отчёта вставлять нельзя. */
      let toolsUsed = false;
      /** Хотя бы один инструмент реально изменил документ. */
      let documentTouched = false;
      const preciseRequest = isPreciseEditRequest(text);
      const createRequest = isCreateRequest(text);
      const rewriteRequest =
        !preciseRequest && !createRequest && REWRITE_INTENT.test(text);
      const atsRequest = ATS_INTENT.test(text);
      const nativeToolsSupported = supportsNativeDocumentTools(settings.provider);
      let usedFallback = !nativeToolsSupported;

      try {
        if (useDocumentAgent && activeTab) {
          const tools = buildDocumentTools(docWriteAccess, activeTab.type);
          const fallbackPrompt = buildDocumentSystemPrompt({
            name: activeTab.name,
            type: activeTab.type,
            writeAccess: docWriteAccess,
            fallback: true,
          });

          /** В fallback модель не может вызвать get_document — документ прикладываем. */
          const fallbackMessage = async (): Promise<ChatMessage> => {
            if (apiUserMessage.attachedFile || !onResolveActiveDocument) return apiUserMessage;
            const document = await onResolveActiveDocument();
            return document ? { ...apiUserMessage, attachedFile: document } : apiUserMessage;
          };

          // Полная переработка и ATS не требуют отдельного get_document:
          // текст уже приложен, модель сразу переписывает документ. Это
          // убирает лишний круг ожидания.
          const agentMessage: ChatMessage = rewriteRequest
            ? await fallbackMessage()
            : { ...apiUserMessage, attachedFile: undefined };

          /**
           * Общая точка вызова инструментов: при точечной правке полная
           * замена документа отклоняется, чтобы модель перешла на
           * find_replace/replace_block, а не переписывала весь файл.
           */
          const callTool = async (
            name: string,
            args: Record<string, unknown>
          ): Promise<DocumentToolResult> => {
            if (preciseRequest && name === 'replace_document') {
              return { ok: false, message: PRECISE_REPLACE_ERROR };
            }
            const result = await callDocumentTool(name, args, transactionId, activeTab.id);
            if (result.ok && name !== 'get_document') documentTouched = true;
            return result;
          };

          setAgentStatus('Готовлю запрос…');

          if (nativeToolsSupported) {
            try {
              const result = await aiService.runDocumentAgent({
                messages: [...base, agentMessage],
                settings: requestSettings,
                appContext,
                systemPromptExtra: buildDocumentSystemPrompt({
                  name: activeTab.name,
                  type: activeTab.type,
                  writeAccess: docWriteAccess,
                }),
                tools,
                signal: controller.signal,
                onStatus: setAgentStatus,
                onToolCall: callTool,
              });

              buffer = result.text.trim();
              toolsUsed = documentTouched;
              if (!buffer) {
                buffer =
                  result.toolCalls > 0
                    ? 'Готово: документ обновлён.'
                    : 'Модель не вернула ответ.';
              }
              paint(true);
            } catch (error) {
              // Шлюз/модель не принимает tools — те же операции текстовым блоком.
              if (!(error instanceof ToolsUnsupportedError)) throw error;
              usedFallback = true;
              setAgentStatus('Модель без function calling — работаю через текстовые действия…');
              await aiService.streamChat(
                [...base, await fallbackMessage()],
                requestSettings,
                streamCallbacks,
                controller.signal,
                appContext,
                fallbackPrompt
              );
            }
          } else {
            setAgentStatus('Готовлю правки…');
            await aiService.streamChat(
              [...base, await fallbackMessage()],
              requestSettings,
              streamCallbacks,
              controller.signal,
              appContext,
              fallbackPrompt
            );
          }

          /* -------- Принудительное выполнение правки -------- */
          const hasAction = hasAllowedDocumentAction(buffer, preciseRequest);
          const firstAnswer = buffer;
          if (
            (createRequest || preciseRequest || rewriteRequest) &&
            !documentTouched &&
            !hasAction &&
            docWriteAccess
          ) {
            setAgentStatus('Повторяю правку документа…');

            const previous: ChatMessage = {
              id: uid('msg'),
              role: 'assistant',
              content: buffer,
              timestamp: Date.now(),
            };
            const nudge: ChatMessage = {
              id: uid('msg'),
              role: 'user',
              content: preciseRequest
                ? 'Ты не внёс изменения в открытый документ. Не отвечай текстом — вызови find_replace (find — точный старый фрагмент) или replace_block. replace_document при точечной правке запрещён.'
                : atsRequest
                  ? 'Ты не внёс изменения в открытый документ. Перепиши резюме под ATS/ATC: плоский одноколоночный текст, стандартные заголовки «Опыт работы», «Навыки», «Образование», ключевые слова, без таблиц и колонок. Не возвращай копию — вызови replace_document с изменённым текстом.'
                  : rewriteRequest
                    ? 'Ты не внёс изменения в открытый документ. Не отвечай текстом — вызови подходящий инструмент документа и выполни правку.'
                    : 'Ты не внёс текст в открытый документ. Не отвечай текстом — вызови replace_document для нового текста или insert_after для добавления.',
              timestamp: Date.now(),
            };

            const strictPrompt =
              buildDocumentSystemPrompt({
                name: activeTab.name,
                type: activeTab.type,
                writeAccess: docWriteAccess,
              }) +
              '\n\nВАЖНО: предыдущая попытка не изменила документ. В этом ответе обязательно вызови инструмент документа, а не пиши текст в чат.' +
              (atsRequest
                ? ' Перепиши резюме под ATS: плоский текст, стандартные заголовки, ключевые слова, без таблиц и колонок; не возвращай копию.'
                : '');

            try {
              if (!usedFallback) {
                const retry = await aiService.runDocumentAgent({
                  messages: [...base, agentMessage, previous, nudge],
                  settings: requestSettings,
                  appContext,
                  systemPromptExtra: strictPrompt,
                  tools,
                  signal: controller.signal,
                  onStatus: setAgentStatus,
                  onToolCall: callTool,
                });
                buffer = retry.text.trim() || buffer;
                toolsUsed = documentTouched;
              }
            } catch (error) {
              if ((error as Error)?.name === 'AbortError') throw error;
              if (error instanceof ToolsUnsupportedError) usedFallback = true;
            }

            if (usedFallback) {
              try {
                await aiService.streamChat(
                  [...base, await fallbackMessage(), previous, nudge],
                  requestSettings,
                  streamCallbacks,
                  controller.signal,
                  appContext,
                  fallbackPrompt +
                    (atsRequest
                      ? '\n\nВАЖНО: предыдущая попытка не изменила документ. Верни workspace-action document_apply с replace_document и текстом, переписанным под ATS (без таблиц и колонок), а не копию.'
                      : '\n\nВАЖНО: предыдущая попытка не изменила документ. Верни workspace-action document_apply с операциями find_replace/replace_block, а не текст в чат.')
                );
              } catch (error) {
                if ((error as Error)?.name === 'AbortError') throw error;
              }
            }

            const retryAction = hasAllowedDocumentAction(buffer, preciseRequest);
            if (!documentTouched && !retryAction) {
              if (preciseRequest) {
                // Детерминированный запасной путь для смены заголовка:
                // модель не вызвала инструмент — меняем только h1.
                const title = extractRequestedTitle(text);
                if (title) {
                  const snapshot = await callDocumentTool(
                    'get_document',
                    {},
                    transactionId,
                    activeTab.id
                  );
                  const currentTitle = snapshot.ok ? firstHeading(snapshot.structure ?? '') : null;
                  if (currentTitle && currentTitle !== title) {
                    const replaced = await callDocumentTool(
                      'find_replace',
                      { find: currentTitle, replace: title, all: false },
                      transactionId,
                      activeTab.id
                    );
                    if (replaced.ok) {
                      documentTouched = true;
                      buffer = `Заменил заголовок на «${title}».`;
                    }
                  }
                }
                if (!documentTouched) {
                  buffer = `${buffer.trim()}\n\n_Правка не применена: модель не вызвала инструмент. Повторите запрос._`;
                }
              } else if (createRequest) {
                const candidates = [buffer, firstAnswer]
                  .map((answer) => extractDocumentContent(answer))
                  .filter((body) => body.trim().length >= 60)
                  .sort((a, b) => b.length - a.length);
                const body = candidates[0];
                if (body) {
                  onInsertIntoEditor(body);
                  documentTouched = true;
                  toolsUsed = true;
                  buffer = 'Текст добавлен в документ.';
                }
              }
            }

            toolsUsed = toolsUsed || documentTouched || hasAllowedDocumentAction(buffer, preciseRequest);
          }
        } else {
          await aiService.streamChat(
            [...base, apiUserMessage],
            requestSettings,
            streamCallbacks,
            controller.signal,
            appContext
          );
        }

        onUpdateMessages(
          sessionId,
          nextMessages.map((m) =>
            m.id === assistantId
              ? {
                  ...m,
                  content: buffer,
                  streaming: false,
                  toolsUsed,
                  preciseEdit: preciseRequest,
                  toolTrace: [...toolTraceRef.current],
                  contextFiles: [...contextFilesRef.current],
                }
              : m
          )
        );

        // Сообщаем наружу готовый ответ — редактор таблицы предложит
        // применить сгенерированную формулу.
        if (buffer.trim() && !useDocumentAgent) onAnswerComplete?.(buffer);
      } catch (err) {
        if ((err as Error)?.name === 'AbortError') {
          // «Стоп» — не ошибка: сохраняем уже полученный текст.
          onUpdateMessages(
            sessionId,
            nextMessages.map((m) =>
              m.id === assistantId
                ? {
                    ...m,
                    content: buffer,
                    streaming: false,
                    toolsUsed,
                    preciseEdit: preciseRequest,
                    toolTrace: [...toolTraceRef.current],
                  }
                : m
            )
          );
        } else {
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
                    preciseEdit: preciseRequest,
                    toolTrace: [...toolTraceRef.current],
                    contextFiles: [...contextFilesRef.current],
                  }
                : m
            )
          );
        }
      } finally {
        setIsStreaming(false);
        setAgentStatus('');
        abortRef.current = null;
      }
    },
    [
      isStreaming,
      isConfigured,
      activeSessionId,
      messages,
      settings,
      activeTab,
      docWriteAccess,
      callDocumentTool,
      onInsertIntoEditor,
      onCreateSession,
      onUpdateMessages,
      onOpenSettings,
      onAnswerComplete,
      appContext,
      onResolveAttachment,
      onResolveActiveDocument,
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
        if (msg.preciseEdit && actionReplacesWholeDocument(act)) {
          executedActionsRef.current.add(key);
          setActionStates((prev) => ({
            ...prev,
            [key]: {
              status: 'error',
              error: 'Точечная правка: полная замена документа запрещена',
            },
          }));
          return;
        }
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
        : undefined,
      undefined,
      'inline'
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
            agentStatus={
              isStreaming && msg.id === messages[messages.length - 1]?.id ? agentStatus : ''
            }
            toolTrace={
              isStreaming && msg.id === messages[messages.length - 1]?.id
                ? toolTrace
                : msg.toolTrace
            }
            contextFiles={
              isStreaming && msg.id === messages[messages.length - 1]?.id
                ? contextFiles
                : msg.contextFiles
            }
            onCopy={() =>
              copyMessage(
                msg.id,
                msg.role === 'user' ? msg.content : stripEmoji(stripActionBlocks(msg.content))
              )
            }
            onInsert={(text) => {
              if (!docWriteAccess) onToggleDocWriteAccess(true);
              onInsertIntoEditor(text);
            }}
            onReplace={(text) => {
              if (!docWriteAccess) onToggleDocWriteAccess(true);
              onApplyToDocument(text);
            }}
            hasActiveTab={Boolean(activeTab)}
            documentTab={Boolean(activeTab && AGENT_TAB_TYPES.has(activeTab.type))}
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

            <button
              onClick={() => onToggleExtraWriteAccess(!extraWriteAccess)}
              title={
                extraWriteAccess
                  ? 'Ассистент может записывать в другие файлы workspace. Нажмите, чтобы вернуть режим только чтения.'
                  : 'Дополнительные файлы доступны только для чтения. Нажмите, чтобы разрешить запись.'
              }
              className={cn(
                'flex items-center gap-1 rounded border px-1.5 py-0.5 text-[10px] transition-colors',
                extraWriteAccess
                  ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-400'
                  : 'border-ws text-ws-faint hover:text-ws'
              )}
            >
              {extraWriteAccess ? <ShieldCheck size={10} /> : <ShieldAlert size={10} />}
              {extraWriteAccess ? 'Доп. файлы: запись' : 'Доп. файлы: чтение'}
            </button>
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
  document_apply: {
    label: 'Правка документа',
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
    case 'document_apply':
      return `${action.operations.length} операций с документом`;
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
  /** Текущий шаг агента, пока ответ ещё пуст. */
  agentStatus?: string;
  /** Инструменты, которые уже вызваны в этом ответе. */
  toolTrace?: string[];
  /** Файлы, прочитанные ассистентом ради ответа. */
  contextFiles?: string[];
  onCopy: () => void;
  onInsert: (text: string) => void;
  onReplace: (text: string) => void;
  hasActiveTab: boolean;
  /** Активен документ Word/Markdown: «Заменить» берёт только содержимое. */
  documentTab?: boolean;
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
  agentStatus,
  toolTrace,
  contextFiles,
  onCopy,
  onInsert,
  onReplace,
  hasActiveTab,
  documentTab,
  actionStates,
  onExecuteAction,
  onOpenFile,
  onCreateTableFromMarkdown,
}) => {
  const isUser = message.role === 'user';
  const codeBlocks = useMemo(
    () =>
      isUser
        ? []
        : extractCodeBlocks(message.content).filter(
            (block) =>
              block.lang !== 'workspace-action' &&
              !/^\s*\{\s*"(?:action|actions|tool)"/.test(block.code)
          ),
    [isUser, message.content]
  );

  const actions = useMemo(
    () => (isUser ? [] : parseAIActions(message.content)),
    [isUser, message.content]
  );

  /** Текст ответа без технических блоков действий. */
  const answerText = useMemo(
    () => (isUser ? '' : stripEmoji(stripActionBlocks(message.content))),
    [isUser, message.content]
  );

  const mdTables = useMemo(
    () =>
      !isUser && !documentTab && actions.length === 0
        ? detectMarkdownTables(message.content)
        : [],
    [isUser, documentTab, actions.length, message.content]
  );

  const displayContent = useMemo(() => {
    if (isUser) return message.content;
    if (!answerText && actions.length > 0) {
      return 'Выполняю команду по управлению рабочей областью...';
    }
    return answerText;
  }, [isUser, message.content, answerText, actions.length]);

  /** Чистый текст документа: без вступления, советов и обёртки ```. */
  const documentBody = useMemo(
    () => (isUser || !documentTab ? '' : buildDocumentBody(answerText)),
    [isUser, documentTab, answerText]
  );

  /** Документ уже изменён инструментами — отчёт в чате нельзя вставлять/заменять. */
  const editsDocument = useMemo(
    () =>
      !isUser &&
      (Boolean(message.toolsUsed) ||
        actions.some(
          (action) =>
            action.action === 'document_apply' ||
            action.action === 'replace_document' ||
            action.action === 'insert_into_editor'
        )),
    [isUser, message.toolsUsed, actions]
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
              {agentStatus || 'Думает…'}
            </span>
          ) : null}

          {!isUser && toolTrace && toolTrace.length > 0 && (
            <div className="mt-1.5 flex flex-wrap gap-1">
              {toolTrace.map((name) => (
                <span
                  key={name}
                  data-tool-trace={name}
                  className="rounded border border-ws bg-ws-elevated px-1 py-0.5 font-mono text-[9px] text-ws-faint"
                >
                  {name}
                </span>
              ))}
            </div>
          )}

          {!isUser && contextFiles && contextFiles.length > 0 && (
            <div className="mt-1.5 flex flex-wrap gap-1">
              {contextFiles.map((name) => (
                <span
                  key={name}
                  data-context-file={name}
                  className="rounded border border-ws bg-ws-elevated px-1 py-0.5 font-mono text-[9px] text-ws-muted"
                >
                  {name}
                </span>
              ))}
            </div>
          )}

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

            {codeBlocks.length > 0 && !documentTab && hasActiveTab && (
              <button
                onClick={() => onInsert(codeBlocks[0].code)}
                title="Вставить первый блок кода в позицию курсора"
                className="flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
              >
                <CornerDownLeft size={10} />
                Вставить код
              </button>
            )}

            {codeBlocks.length > 1 && !documentTab && hasActiveTab && (
              <button
                onClick={() => onInsert(codeBlocks.map((b) => b.code).join('\n\n'))}
                title={`Вставить все блоки кода (${codeBlocks.length})`}
                className="flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
              >
                <Code2 size={10} />
                Все блоки ({codeBlocks.length})
              </button>
            )}

            {documentTab && documentBody && !editsDocument && hasActiveTab && (
              <button
                onClick={() => onInsert(documentBody)}
                title="Вставить только текст документа в позицию курсора"
                className="flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
              >
                <CornerDownLeft size={10} />
                Вставить в документ
              </button>
            )}

            {codeBlocks.length === 0 && !documentTab && answerText && !editsDocument && hasActiveTab && (
              <button
                onClick={() => onInsert(answerText)}
                title="Вставить текст ответа в позицию курсора активного документа"
                className="flex items-center gap-1 rounded px-1.5 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
              >
                <CornerDownLeft size={10} />
                Вставить в документ
              </button>
            )}

            {hasActiveTab && !editsDocument && (documentTab ? documentBody : answerText) && (
              <button
                onClick={() => onReplace(documentTab ? documentBody : answerText)}
                title={
                  documentTab
                    ? 'Заменить содержимое документа только текстом ответа (без вступления и пояснений)'
                    : 'Заменить содержимое документа этим ответом'
                }
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
