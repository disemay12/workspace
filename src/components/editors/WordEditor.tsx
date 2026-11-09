import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { OpenTab, PageSetup } from '../../types';
import { bindDocxBlockSources, docxBase64ToHtml, docxPackageInfo, htmlToDocxBase64 } from '../../lib/docx';
import type { DocxBlockSource, HeaderFooterEdit } from '../../lib/docx';
import { offsetWithinBlock, applyDocumentTool, sanitizeSelectionHtml, selectionRangeHtml } from '../../lib/documentModel';
import { DocumentHistory } from '../../lib/documentHistory';
import type { DocumentTransaction } from '../../lib/documentHistory';
import type { DocumentSelection } from '../../services/docBridge';
import { markdownToDocumentHtml } from '../../lib/aiDocument';
import {
  FileText,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Heading1,
  Heading2,
  Heading3,
  List,
  ListOrdered,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  Sparkles,
  Save,
  AlertTriangle,
  Undo2,
  Redo2,
  History as HistoryIcon,
  ListTree,
  Search,
  Clock,
  Highlighter,
  Palette,
  Download,
  Baseline,
  Superscript,
  Subscript,
  RemoveFormatting,
  ChevronDown,
} from 'lucide-react';
import { cn } from '../../lib/utils';
import { IconButton, Spinner } from '../ui';
import { registerDocumentController } from '../../services/docBridge';

interface Props {
  tab: OpenTab;
  onChange: (newContent: string) => void;
  onSave: () => void;
  onAskAI?: (prompt: string, contextText: string) => void;
}

/** Шрифты, которые есть в типичной установке Windows и которые Word показывает первыми. */
const FONT_FAMILIES = [
  { label: 'Calibri', value: 'Calibri, Carlito, sans-serif' },
  { label: 'Arial', value: 'Arial, Helvetica, sans-serif' },
  { label: 'Times New Roman', value: '"Times New Roman", Times, serif' },
  { label: 'Segoe UI', value: '"Segoe UI", sans-serif' },
  { label: 'Georgia', value: 'Georgia, serif' },
  { label: 'Verdana', value: 'Verdana, Geneva, sans-serif' },
  { label: 'Tahoma', value: 'Tahoma, Geneva, sans-serif' },
  { label: 'Trebuchet MS', value: '"Trebuchet MS", sans-serif' },
  { label: 'Courier New', value: '"Courier New", Courier, monospace' },
  { label: 'Consolas', value: 'Consolas, "Courier New", monospace' },
  { label: 'Comic Sans MS', value: '"Comic Sans MS", cursive' },
  { label: 'Impact', value: 'Impact, Haettenschweiler, sans-serif' },
  { label: 'Cambria', value: 'Cambria, Georgia, serif' },
  { label: 'Candara', value: 'Candara, Calibri, sans-serif' },
  { label: 'Constantia', value: 'Constantia, Georgia, serif' },
  { label: 'Corbel', value: 'Corbel, "Segoe UI", sans-serif' },
  { label: 'Garamond', value: 'Garamond, "Times New Roman", serif' },
  { label: 'Palatino Linotype', value: '"Palatino Linotype", Palatino, serif' },
];

/** Кегли в пунктах — те же, что в выпадающем списке Word. */
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

const HIGHLIGHT_COLORS = [
  '#fef08a',
  '#fde68a',
  '#bbf7d0',
  '#86efac',
  '#bfdbfe',
  '#93c5fd',
  '#fecaca',
  '#fca5a5',
  '#e9d5ff',
  '#fbcfe8',
  '#fed7aa',
  '#e2e8f0',
  'transparent',
];

const TEXT_COLORS = [
  '#0f172a',
  '#334155',
  '#64748b',
  '#dc2626',
  '#ea580c',
  '#d97706',
  '#ca8a04',
  '#16a34a',
  '#059669',
  '#0891b2',
  '#2563eb',
  '#4f46e5',
  '#7c3aed',
  '#c026d3',
  '#db2777',
  '#ffffff',
];

interface FormatState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  superscript: boolean;
  subscript: boolean;
  fontName: string;
  fontSizePt: string;
  foreColor: string;
  hiliteColor: string;
  block: string;
}

const EMPTY_FORMAT: FormatState = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  superscript: false,
  subscript: false,
  fontName: '',
  fontSizePt: '',
  foreColor: '',
  hiliteColor: '',
  block: '',
};

interface DocumentVersion {
  id: string;
  timestamp: number;
  label: string;
  html: string;
}

/** Версии живут вне компонента: переключение вкладок не должно их терять. */
const versionStore = new Map<string, DocumentVersion[]>();

/** Диапазон по символьным смещениям внутри одного блока. */
function rangeForBlockOffsets(block: Element, start: number, end: number): Range | null {
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  let current = walker.nextNode();
  while (current) {
    nodes.push(current as Text);
    current = walker.nextNode();
  }
  if (nodes.length === 0) return null;

  let position = 0;
  let startNode: Text | null = null;
  let endNode: Text | null = null;
  let startOffset = 0;
  let endOffset = 0;
  for (const node of nodes) {
    const length = (node.textContent ?? '').length;
    if (!startNode && start < position + length) {
      startNode = node;
      startOffset = start - position;
    }
    if (!endNode && end <= position + length && end > position) {
      endNode = node;
      endOffset = end - position;
    }
    position += length;
  }
  if (!startNode) {
    startNode = nodes[nodes.length - 1];
    startOffset = (startNode.textContent ?? '').length;
  }
  if (!endNode) {
    endNode = nodes[nodes.length - 1];
    endOffset = (endNode.textContent ?? '').length;
  }
  const range = document.createRange();
  range.setStart(startNode, Math.max(0, startOffset));
  range.setEnd(endNode, Math.max(0, endOffset));
  return range;
}

export const WordEditor: React.FC<Props> = ({ tab, onChange, onSave, onAskAI }) => {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [showColors, setShowColors] = useState<'highlight' | 'text' | null>(null);
  const [showFonts, setShowFonts] = useState(false);
  const [showSizes, setShowSizes] = useState(false);
  const [wordCount, setWordCount] = useState(0);
  const [showAI, setShowAI] = useState(false);
  const [aiPrompt, setAiPrompt] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [aiError, setAiError] = useState('');
  const [format, setFormat] = useState<FormatState>(EMPTY_FORMAT);
  /** HTML до последней замены ассистентом — для кнопки «Вернуть прежний текст». */
  const [aiSnapshot, setAiSnapshot] = useState<string | null>(null);
  /** Серия правок одного ответа ассистента = один шаг отката. */
  const aiTxnRef = useRef('');
  /** Сколько изображений удалось сохранить при последней замене. */
  const [restoredImages, setRestoredImages] = useState(0);
  /** Поля, размер страницы и базовый шрифт: читаются из .docx и меняются AI. */
  const [pageSetup, setPageSetup] = useState<PageSetup>({});
  const pageSetupRef = useRef<PageSetup>({});

  const editorRef = useRef<HTMLDivElement>(null);
  /**
   * Разобранный HTML документа. Пока идёт загрузка, полотно не отрисовано и
   * editorRef.current равен null, поэтому записать разметку сразу после разбора
   * невозможно — она сохраняется здесь и применяется, когда полотно появится.
   */
  const [pendingHtml, setPendingHtml] = useState('');
  /** Флаг: контент установлен программно, не нужно считать это правкой пользователя. */
  const isSyncing = useRef(false);
  /** Порядковый номер сериализации, чтобы устаревший DOCX не затёр более новую правку. */
  const serializeSeq = useRef(0);
  /** Куда вставить ответ модели: выделение запоминается до того, как фокус уйдёт в чат. */
  const pendingAI = useRef<{ range: Range | null } | null>(null);
  /** Выделение, которое нужно вернуть после клика по панели (кнопка забирает фокус). */
  const savedRange = useRef<Range | null>(null);
  /** Точное выделение для инструментов: сохраняется, когда фокус уходит в чат. */
  const selectionRef = useRef<{ data: DocumentSelection; range: Range } | null>(null);
  /** Исходный пакет вкладки: источник истины для сохранения. */
  const sourceBase64Ref = useRef('');
  /** Исходные XML-блоки документа (в порядке HTML-блоков). */
  const blockSourcesRef = useRef<DocxBlockSource[]>([]);
  /** Карта data-docx-id → исходный XML для патча при сохранении. */
  const blockXmlRef = useRef<Map<string, string>>(new Map());
  /** Несохранённые правки колонтитулов: kind:type → edit. */
  const headerFooterEditsRef = useRef<Map<string, HeaderFooterEdit>>(new Map());
  /** История AI-транзакций: before/after, diff, undo/redo. */
  const historyRef = useRef(new DocumentHistory());
  const openTxnRef = useRef<{
    id: string;
    label: string;
    timestamp: number;
    before: string;
    timer?: ReturnType<typeof setTimeout>;
  } | null>(null);
  const [, setHistoryVersion] = useState(0);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [versionsOpen, setVersionsOpen] = useState(false);
  const versionsRef = useRef<DocumentVersion[]>(versionStore.get(tab.id) ?? []);
  const [, setVersionsTick] = useState(0);
  const [showOutline, setShowOutline] = useState(true);
  const [zoom, setZoom] = useState(1);
  const zoomScrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const element = zoomScrollRef.current;
    if (!element) return;
    const onWheel = (event: WheelEvent) => {
      if (!(event.ctrlKey || event.metaKey)) return;
      event.preventDefault();
      setZoom((previous) =>
        Math.min(2, Math.max(0.5, previous - event.deltaY * 0.0015))
      );
    };
    element.addEventListener('wheel', onWheel, { passive: false });
    return () => element.removeEventListener('wheel', onWheel);
  }, [loading]);
  const [findOpen, setFindOpen] = useState(false);
  const [findText, setFindText] = useState('');
  const [replaceWith, setReplaceWith] = useState('');
  const [findMessage, setFindMessage] = useState('');
  const [breadcrumb, setBreadcrumb] = useState('');
  const findRangeRef = useRef<Range | null>(null);
  const [outline, setOutline] = useState<Array<{ id: string; text: string; level: number }>>([]);

  /** Сериализация всегда идёт поверх исходного пакета вкладки. */
  const serializeDocx = useCallback(
    (html: string, page?: PageSetup) =>
      htmlToDocxBase64(html, page, {
        sourceBase64: sourceBase64Ref.current || undefined,
        blockXml: blockXmlRef.current,
        headerFooterEdits: [...headerFooterEditsRef.current.values()],
      }),
    []
  );

  /** Заголовки документа для панели структуры. */
  const refreshOutline = useCallback(() => {
    const root = editorRef.current;
    if (!root) {
      setOutline([]);
      return;
    }
    const items = [...root.querySelectorAll('h1, h2, h3, h4, h5, h6')]
      .map((element, index) => ({
        id: element.getAttribute('data-bid') ?? `outline-${index}`,
        text: (element.textContent ?? '').replace(/\s+/g, ' ').trim(),
        level: Number(element.tagName.slice(1)),
      }))
      .filter((item) => item.text);
    setOutline(items);
  }, []);
  /** Заголовок, в котором находится курсор — для breadcrumbs. */
  const refreshBreadcrumb = useCallback(() => {
    const root = editorRef.current;
    const saved = selectionRef.current;
    const block = saved && root ? root.querySelector(`[data-bid="${saved.data.blockId}"]`) : null;
    if (!root || !block) {
      setBreadcrumb('');
      return;
    }
    let current = '';
    for (const heading of root.querySelectorAll('h1, h2, h3, h4, h5, h6')) {
      if (heading.compareDocumentPosition(block) & Node.DOCUMENT_POSITION_FOLLOWING) {
        current = (heading.textContent ?? '').replace(/\s+/g, ' ').trim();
      }
    }
    setBreadcrumb(current);
  }, []);

  /* ---------------- Точное выделение для инструментов ---------------- */

  useEffect(() => {
    const onSelectionChange = () => {
      const root = editorRef.current;
      if (!root) return;
      const selection = window.getSelection();
      if (!selection || selection.rangeCount === 0) return;
      const range = selection.getRangeAt(0);
      if (!root.contains(range.commonAncestorContainer)) return;

      const startElement =
        range.startContainer.nodeType === Node.ELEMENT_NODE
          ? (range.startContainer as Element)
          : range.startContainer.parentElement;
      const block = startElement?.closest('[data-bid]') ?? null;
      if (!block || !root.contains(block)) return;

      const start = offsetWithinBlock(block, range.startContainer, range.startOffset);
      const end = offsetWithinBlock(block, range.endContainer, range.endOffset);
      if (start === null || end === null) return;

      const text = block.textContent ?? '';
      const from = Math.min(start, end);
      const to = Math.max(start, end);
      selectionRef.current = {
        range: range.cloneRange(),
        data: {
          documentId: tab.id,
          blockId: block.getAttribute('data-bid') ?? '',
          startOffset: from,
          endOffset: to,
          text: text.slice(from, to),
          html: selectionRangeHtml(range),
        },
      };
      refreshBreadcrumb();
    };

    document.addEventListener('selectionchange', onSelectionChange);
    return () => document.removeEventListener('selectionchange', onSelectionChange);
  }, [tab.id, refreshBreadcrumb]);

  /* ---------------- Разбор DOCX через mammoth ---------------- */

  useEffect(() => {
    let cancelled = false;

    const parseDocx = async () => {
      try {
        setLoading(true);
        setError('');
        sourceBase64Ref.current = '';
        blockSourcesRef.current = [];
        blockXmlRef.current = new Map();
        headerFooterEditsRef.current = new Map();
        versionsRef.current = versionStore.get(tab.id) ?? [];

        if (!tab.content) {
          if (!cancelled) {
            setPendingHtml('<p></p>');
            pageSetupRef.current = {};
            setPageSetup({});
            setLoading(false);
            setWordCount(0);
          }
          return;
        }

        // Шрифт, кегль и цвет поднимаются из OOXML в inline-стили.
        const parsed = await docxBase64ToHtml(tab.content);

        if (cancelled) return;

        // Разметка применяется отдельным эффектом: во время загрузки полотна
        // в DOM ещё нет, и прямая запись в editorRef терялась бы.
        setPendingHtml(parsed.html || '<p></p>');
        pageSetupRef.current = parsed.page ?? {};
        setPageSetup(parsed.page ?? {});
        sourceBase64Ref.current = tab.content;
        blockSourcesRef.current = parsed.blocks ?? [];
        setLoading(false);
      } catch (err) {
        if (cancelled) return;
        // Резервное восстановление: файл мог быть создан как Markdown/текст вместо бинарного ZIP
        const recoveredText = tryDecodeText(tab.content);
        if (recoveredText && recoveredText.trim()) {
          const html = markdownToDocumentHtml(recoveredText);
          setPendingHtml(html || '<p></p>');
          pageSetupRef.current = {};
          setPageSetup({});
          setLoading(false);
          void htmlToDocxBase64(html, pageSetupRef.current)
            .then((b64) => {
              if (!cancelled) onChange(b64);
            })
            .catch(() => {});
          return;
        }

        setError(`Не удалось разобрать документ: ${(err as Error).message}`);
        setLoading(false);
      }
    };

    void parseDocx();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  /* ---------------- Применение разобранной разметки ---------------- */

  useEffect(() => {
    const root = editorRef.current;
    if (!root || loading) return;

    // Программная установка HTML вызывает onInput. Флаг сбрасывается после
    // неё, иначе каждая правка пользователя отбрасывается и кнопка
    // «Сохранить» навсегда остаётся неактивной.
    isSyncing.current = true;
    root.innerHTML = pendingHtml || '<p></p>';
    blockXmlRef.current = bindDocxBlockSources(root, blockSourcesRef.current);
    setWordCount(countWords(root.innerText));
    refreshOutline();
    queueMicrotask(() => {
      isSyncing.current = false;
    });
  }, [pendingHtml, loading, refreshOutline]);


  /* ---------------- Отслеживание правок с debounce ---------------- */

  const debounceTimer = useRef<NodeJS.Timeout | null>(null);

  useEffect(() => {
    return () => {
      if (debounceTimer.current) clearTimeout(debounceTimer.current);
    };
  }, []);

  /** Переход к заголовку из панели структуры. */
  const scrollToOutline = useCallback((item: { id: string; text: string }) => {
    const root = editorRef.current;
    if (!root) return;
    const byId = root.querySelector(`[data-bid="${item.id.replace(/"/g, '')}"]`);
    const target =
      byId ??
      [...root.querySelectorAll('h1, h2, h3, h4, h5, h6')].find(
        (element) => (element.textContent ?? '').trim() === item.text
      ) ??
      null;
    if (target instanceof HTMLElement) {
      target.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
  }, []);

  const handleInput = useCallback(() => {
    if (!editorRef.current || isSyncing.current) return;
    const html = editorRef.current.innerHTML;
    setWordCount(countWords(editorRef.current.innerText));
    refreshOutline();
    refreshBreadcrumb();

    // Debounce упаковки пакета .docx (устраняет фризы при быстром наборе текста)
    if (debounceTimer.current) clearTimeout(debounceTimer.current);
    debounceTimer.current = setTimeout(() => {
      const seq = ++serializeSeq.current;
      void serializeDocx(html, pageSetupRef.current)
        .then((base64) => {
          if (serializeSeq.current === seq) onChange(base64);
        })
        .catch(() => {
          /* Ошибку сериализации не показываем на каждый символ: документ остаётся открытым. */
        });
    }, 400);
  }, [onChange, refreshOutline, refreshBreadcrumb]);

  /* ---------------- Поиск и замена в документе ---------------- */

  const collectTextNodes = (root: Node): Text[] => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const nodes: Text[] = [];
    let current = walker.nextNode();
    while (current) {
      nodes.push(current as Text);
      current = walker.nextNode();
    }
    return nodes;
  };

  const textOffsetOf = (root: Element, node: Node, offset: number): number => {
    const probe = document.createRange();
    probe.setStart(root, 0);
    probe.setEnd(node, offset);
    return (probe.cloneContents().textContent ?? '').length;
  };

  const rangeAtTextOffsets = (nodes: Text[], start: number, end: number): Range | null => {
    if (nodes.length === 0) return null;
    let position = 0;
    let startNode: Text | null = null;
    let endNode: Text | null = null;
    let startOffset = 0;
    let endOffset = 0;
    for (const node of nodes) {
      const length = (node.textContent ?? '').length;
      if (!startNode && start < position + length) {
        startNode = node;
        startOffset = start - position;
      }
      if (!endNode && end <= position + length && end > position) {
        endNode = node;
        endOffset = end - position;
      }
      position += length;
    }
    if (!startNode) {
      startNode = nodes[nodes.length - 1];
      startOffset = (startNode.textContent ?? '').length;
    }
    if (!endNode) {
      endNode = nodes[nodes.length - 1];
      endOffset = (endNode.textContent ?? '').length;
    }
    const range = document.createRange();
    range.setStart(startNode, startOffset);
    range.setEnd(endNode, endOffset);
    return range;
  };

  const highlightFindRange = (range: Range | null) => {
    const HighlightCtor = (window as unknown as { Highlight?: new () => { add: (range: Range) => void } })
      .Highlight;
    const registry = (CSS as unknown as { highlights?: Map<string, unknown> }).highlights;
    if (!registry) return;
    if (!range || !HighlightCtor) {
      registry.delete('ws-find');
      return;
    }
    const highlight = new HighlightCtor();
    highlight.add(range);
    registry.set('ws-find', highlight);
  };

  const findNext = useCallback(() => {
    const root = editorRef.current;
    const query = findText.trim();
    if (!root || !query) {
      setFindMessage('Введите текст для поиска');
      return;
    }

    const nodes = collectTextNodes(root);
    const full = nodes.map((node) => node.textContent ?? '').join('');
    const haystack = full.toLowerCase();
    const needle = query.toLowerCase();
    let from = 0;
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const current = selection.getRangeAt(0);
      if (root.contains(current.endContainer)) {
        from = textOffsetOf(root, current.endContainer, current.endOffset);
      }
    }

    let index = haystack.indexOf(needle, from);
    if (index < 0) index = haystack.indexOf(needle);
    if (index < 0) {
      findRangeRef.current = null;
      highlightFindRange(null);
      setFindMessage('Не найдено');
      return;
    }

    const range = rangeAtTextOffsets(nodes, index, index + query.length);
    if (!range) {
      setFindMessage('Не найдено');
      return;
    }
    findRangeRef.current = range;
    highlightFindRange(range);
    selection?.removeAllRanges();
    selection?.addRange(range);
    const anchor =
      range.startContainer.nodeType === Node.ELEMENT_NODE
        ? (range.startContainer as Element)
        : range.startContainer.parentElement;
    anchor?.scrollIntoView({ block: 'center', behavior: 'smooth' });
    setFindMessage(`Найдено: ${index + 1}`);
  }, [findText]);

  const replaceCurrent = useCallback(() => {
    const root = editorRef.current;
    const query = findText.trim();
    if (!root || !query) {
      setFindMessage('Введите текст для поиска');
      return;
    }

    let range = findRangeRef.current;
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const current = selection.getRangeAt(0);
      if (!current.collapsed && current.toString().toLowerCase() === query.toLowerCase()) {
        range = current;
      }
    }
    if (!range || range.toString().toLowerCase() !== query.toLowerCase()) {
      findNext();
      range = findRangeRef.current;
    }
    if (!range) return;

    range.deleteContents();
    const inserted = document.createTextNode(replaceWith);
    range.insertNode(inserted);
    const after = document.createRange();
    after.setStartAfter(inserted);
    after.collapse(true);
    selection?.removeAllRanges();
    selection?.addRange(after);
    findRangeRef.current = null;
    highlightFindRange(null);
    setFindMessage('Заменено');
    handleInput();
  }, [findNext, findText, handleInput, replaceWith]);

  const replaceAll = useCallback(() => {
    const root = editorRef.current;
    const query = findText.trim();
    if (!root || !query) {
      setFindMessage('Введите текст для поиска');
      return;
    }
    const result = applyDocumentTool(root.innerHTML, 'find_replace', {
      find: query,
      replace: replaceWith,
      all: true,
    });
    if (!result.ok) {
      setFindMessage(result.message);
      return;
    }
    if (result.html !== undefined) {
      isSyncing.current = true;
      root.innerHTML = result.html;
      isSyncing.current = false;
      setWordCount(countWords(root.innerText));
      refreshOutline();
      handleInput();
    }
    findRangeRef.current = null;
    highlightFindRange(null);
    setFindMessage(result.message);
  }, [findText, handleInput, refreshOutline, replaceWith]);

  /** Снимок документа для панели версий (не чаще одного на одинаковое состояние). */
  const captureVersion = useCallback((label: string) => {
    const root = editorRef.current;
    if (!root) return;
    const html = root.innerHTML;
    const list = versionsRef.current;
    if (list[0]?.html === html) return;
    list.unshift({ id: `v${Date.now()}`, timestamp: Date.now(), label, html });
    if (list.length > 20) list.length = 20;
    versionStore.set(tab.id, list);
    setVersionsTick((value) => value + 1);
  }, [tab.id]);

  const flushSave = useCallback(async () => {
    if (debounceTimer.current) {
      clearTimeout(debounceTimer.current);
      debounceTimer.current = null;
    }
    if (editorRef.current) {
      try {
        const html = editorRef.current.innerHTML;
        const base64 = await serializeDocx(html, pageSetupRef.current);
        onChange(base64);
      } catch {
        // ignore
      }
    }
    onSave();
    captureVersion('Сохранено');
  }, [onChange, onSave, serializeDocx, captureVersion]);

  /** Закрывает открытую транзакцию и пишет её в историю. */
  const finalizeTransaction = useCallback(() => {
    const open = openTxnRef.current;
    const root = editorRef.current;
    if (!open) return;
    if (open.timer) clearTimeout(open.timer);
    openTxnRef.current = null;
    if (!root || open.before === root.innerHTML) return;

    historyRef.current.commit({
      id: open.id,
      label: open.label,
      timestamp: open.timestamp,
      before: open.before,
      after: root.innerHTML,
    });

    // Подсвечиваем изменённые блоки через CSS Highlight API: DOM не меняется,
    // поэтому подсветка не попадает в HTML и .docx.
    const transaction = historyRef.current.current;
    const ranges: Range[] = [];
    for (const change of transaction?.diff.blocks ?? []) {
      const rawId = change.id.split('#')[0];
      if (!/^[a-zA-Z0-9_-]+$/.test(rawId)) continue;
      const node = root.querySelector(`[data-bid="${rawId}"], [data-docx-id="${rawId}"]`);
      if (node) {
        const range = document.createRange();
        range.selectNodeContents(node);
        ranges.push(range);
      }
    }
    const HighlightCtor = (window as unknown as { Highlight?: new () => { add: (range: Range) => void } })
      .Highlight;
    const registry = (CSS as unknown as { highlights?: Map<string, unknown> }).highlights;
    if (ranges.length > 0 && HighlightCtor && registry) {
      const highlight = new HighlightCtor();
      for (const range of ranges) highlight.add(range);
      registry.set('ws-ai-changed', highlight);
      setTimeout(() => registry.delete('ws-ai-changed'), 1600);
    }
    setHistoryVersion((value) => value + 1);
    captureVersion('После правки AI');
  }, [captureVersion]);

  const scheduleFinalize = useCallback(() => {
    const open = openTxnRef.current;
    if (!open) return;
    if (open.timer) clearTimeout(open.timer);
    open.timer = setTimeout(() => finalizeTransaction(), 2500);
  }, [finalizeTransaction]);

  /** Первая правка ответа AI запоминает документ целиком для отката. */
  const beginAiTransaction = useCallback(
    (root: HTMLElement, transactionId: string) => {
      if (aiTxnRef.current === transactionId) {
        // Таймер мог уже закрыть часть серии; продолжение — новая транзакция.
        if (!openTxnRef.current) {
          openTxnRef.current = {
            id: `${transactionId}-${Date.now()}`,
            label: 'Правка ассистента',
            timestamp: Date.now(),
            before: root.innerHTML,
          };
        }
        scheduleFinalize();
        return;
      }

      finalizeTransaction();
      captureVersion('До правки AI');
      aiTxnRef.current = transactionId;
      setAiSnapshot(root.innerHTML);
      setRestoredImages(0);
      setAiError('');
      openTxnRef.current = {
        id: transactionId,
        label: 'Правка ассистента',
        timestamp: Date.now(),
        before: root.innerHTML,
      };
      scheduleFinalize();
    },
    [finalizeTransaction, scheduleFinalize, captureVersion]
  );

  /** Вставляет санитизированный фрагмент в диапазон и схлопывает каретку. */
  const insertFragment = useCallback(
    (root: HTMLElement, range: Range, html: string, transactionId: string) => {
      const clean = sanitizeSelectionHtml(html);
      if (!clean) return false;

      beginAiTransaction(root, transactionId);
      const target = range.cloneRange();
      const template = document.createElement('template');
      template.innerHTML = clean;
      target.deleteContents();
      target.insertNode(template.content);
      target.collapse(false);

      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(target);
      selectionRef.current = null;

      setWordCount(countWords(root.innerText));
      handleInput();
      return true;
    },
    [beginAiTransaction, handleInput]
  );

  /* ---------------- Мост для инструментов ассистента ---------------- */

  useEffect(() => {
    return registerDocumentController({
      tabId: tab.id,
      readHtml: () => editorRef.current?.innerHTML ?? '',
      serialize: async () => {
        const root = editorRef.current;
        if (!root) return tab.content;
        return await serializeDocx(root.innerHTML, pageSetupRef.current);
      },
      writeHtml: (html, transactionId) => {
        const root = editorRef.current;
        if (!root) return;

        const savedSelection = selectionRef.current?.data ?? null;

        // Первая правка серии запоминает документ целиком: «Вернуть прежний
        // текст» откатывает весь ответ ассистента, а не последний блок.
        beginAiTransaction(root, transactionId);

        isSyncing.current = true;
        root.innerHTML = html || '<p></p>';
        if (savedSelection && savedSelection.startOffset < savedSelection.endOffset) {
          const block =
            root.querySelector(`[data-bid="${savedSelection.blockId}"]`) ??
            root.querySelector(`[data-docx-id="${savedSelection.blockId}"]`);
          if (block) {
            const length = (block.textContent ?? '').length;
            const start = Math.max(0, Math.min(savedSelection.startOffset, length));
            const end = Math.max(start, Math.min(savedSelection.endOffset, length));
            const range = start < end ? rangeForBlockOffsets(block, start, end) : null;
            if (range) {
              const selection = window.getSelection();
              selection?.removeAllRanges();
              selection?.addRange(range);
              selectionRef.current = {
                range: range.cloneRange(),
                data: {
                  ...savedSelection,
                  blockId: block.getAttribute('data-bid') ?? savedSelection.blockId,
                  startOffset: start,
                  endOffset: end,
                  text: (block.textContent ?? '').slice(start, end),
                  html: selectionRangeHtml(range),
                },
              };
            }
          }
        }
        setWordCount(countWords(root.innerText));
        isSyncing.current = false;
        handleInput();
      },
      getSelection: () => {
        const root = editorRef.current;
        const saved = selectionRef.current;
        if (!root || !saved || !root.contains(saved.range.startContainer)) return null;
        return saved.data;
      },
      restoreHtml: (html) => {
        const root = editorRef.current;
        if (!root) return;
        isSyncing.current = true;
        root.innerHTML = html || '<p></p>';
        selectionRef.current = null;
        setWordCount(countWords(root.innerText));
        refreshOutline();
        setAiSnapshot(null);
        aiTxnRef.current = '';
        isSyncing.current = false;
        handleInput();
      },
      replaceSelection: (html, transactionId) => {
        const root = editorRef.current;
        const saved = selectionRef.current;
        if (!root || !saved || !root.contains(saved.range.startContainer)) {
          return { ok: false, message: 'Нет выделенного фрагмента в документе' };
        }
        const ok = insertFragment(root, saved.range, html, transactionId);
        return ok
          ? { ok: true, message: 'Выделенный фрагмент заменён', changed: true }
          : { ok: false, message: 'Пустой фрагмент для замены' };
      },
      insertAtSelection: (html, transactionId) => {
        const root = editorRef.current;
        const saved = selectionRef.current;
        if (!root || !saved || !root.contains(saved.range.startContainer)) {
          return { ok: false, message: 'Нет каретки или выделения в документе' };
        }
        const range = saved.range.cloneRange();
        range.collapse(false);
        const ok = insertFragment(root, range, html, transactionId);
        return ok
          ? { ok: true, message: 'Фрагмент вставлен в позицию каретки', changed: true }
          : { ok: false, message: 'Пустой фрагмент для вставки' };
      },
      getDocumentInfo: async () => {
        const source = sourceBase64Ref.current || tab.content;
        if (!source) return null;
        try {
          const info = await docxPackageInfo(source);
          const edits = [...headerFooterEditsRef.current.values()];
          if (edits.length === 0) return info;

          const headers = [...info.headers];
          const footers = [...info.footers];
          for (const edit of edits) {
            const list = edit.kind === 'header' ? headers : footers;
            const type = edit.referenceType ?? 'default';
            const index = list.findIndex(
              (part) => (part.referenceType ?? 'default') === type
            );
            if (edit.remove) {
              if (index >= 0) list.splice(index, 1);
              continue;
            }
            if (index >= 0) {
              list[index] = { ...list[index], text: edit.text, referenceType: type };
            } else {
              list.push({ name: '(новый)', text: edit.text, referenceType: type });
            }
          }
          return { ...info, headers, footers };
        } catch {
          return null;
        }
      },
      setHeaderFooter: async (edit) => {
        const root = editorRef.current;
        if (!root) return { ok: false, message: 'Редактор документа недоступен' };

        const key = `${edit.kind}:${edit.referenceType ?? 'default'}`;
        if (edit.remove) headerFooterEditsRef.current.delete(key);
        else headerFooterEditsRef.current.set(key, edit);

        try {
          const seq = ++serializeSeq.current;
          const base64 = await serializeDocx(root.innerHTML, pageSetupRef.current);
          if (serializeSeq.current === seq) onChange(base64);
          return {
            ok: true,
            message: edit.remove ? 'Колонтитул удалён' : 'Колонтитул обновлён',
            changed: true,
          };
        } catch (error) {
          return {
            ok: false,
            message: `Не удалось обновить колонтитул: ${
              error instanceof Error ? error.message : String(error)
            }`,
          };
        }
      },
      applyPageSetup: (page) => {
        const merged = { ...pageSetupRef.current, ...page };
        pageSetupRef.current = merged;
        setPageSetup(merged);

        const root = editorRef.current;
        if (!root) return;
        // Поля и базовый шрифт живут в sectPr/styles.xml: без пересборки
        // пакета они не попадут в сохраняемый .docx.
        const seq = ++serializeSeq.current;
        void serializeDocx(root.innerHTML, merged)
          .then((base64) => {
            if (serializeSeq.current === seq) onChange(base64);
          })
          .catch(() => {
            /* Ошибку сериализации покажет ближайшее сохранение. */
          });
      },
    });
  }, [tab.id, handleInput, onChange, beginAiTransaction, insertFragment, serializeDocx]);

  /* ---------------- Ответ AI ---------------- */

  useEffect(() => {
    const handler = (event: Event) => {
      const detail = (event as CustomEvent).detail as { tabId?: string; answer?: string } | undefined;
      if (!detail?.answer || detail.tabId !== tab.id) return;

      const pending = pendingAI.current;
      if (!pending) return;
      pendingAI.current = null;

      const html = markdownToDocumentHtml(detail.answer);
      if (!html.trim()) {
        setAiBusy(false);
        setAiError('Модель вернула пустой ответ');
        return;
      }

      insertHtmlAt(editorRef.current, pending.range, html);
      setAiBusy(false);
      setAiPrompt('');
      setShowAI(false);
      setAiError('');
      handleInput();
    };

    window.addEventListener('workspace:ai-answer', handler);
    return () => window.removeEventListener('workspace:ai-answer', handler);
  }, [handleInput, tab.id]);

  /* ---------------- Вставка текста и замена документа из чата AI ---------------- */

  useEffect(() => {
    const handleInsert = (event: Event) => {
      const detail = (event as CustomEvent).detail as { tabId?: string; text?: string } | undefined;
      if (!detail?.text || detail.tabId !== tab.id) return;
      const html = markdownToDocumentHtml(detail.text);
      if (!html) return;
      // Кнопка в чате забирает фокус, поэтому текущее выделение уже не в
      // документе. Вставляем в последнюю позицию каретки, а если её нет —
      // в конец документа; insertHtmlAt это гарантирует.
      insertHtmlAt(editorRef.current, savedRange.current, html);
      rememberSelection();
      handleInput();
    };

    const handleReplace = (event: Event) => {
      const detail = (event as CustomEvent).detail as { tabId?: string; text?: string } | undefined;
      if (detail?.tabId !== tab.id) return;
      const root = editorRef.current;
      if (!root) return;
      const html = markdownToDocumentHtml(detail.text || '');
      beginAiTransaction(root, 'manual');
      isSyncing.current = true;
      // Картинки исходного документа не должны исчезать при переписывании:
      // модель их не воспроизводит, поэтому возвращаем на прежние места.
      const images = collectAnchoredImages(root);
      root.innerHTML = html || '<p></p>';
      setRestoredImages(restoreAnchoredImages(root, images));
      setWordCount(countWords(root.innerText));
      // Сбрасываем флаг до handleInput: иначе сериализация вкладки
      // пропускается и картинки не попадают в сохраняемый .docx.
      isSyncing.current = false;
      handleInput();
      scheduleFinalize();
    };

    window.addEventListener('workspace:insert-text', handleInsert);
    window.addEventListener('workspace:replace-document', handleReplace);
    return () => {
      window.removeEventListener('workspace:insert-text', handleInsert);
      window.removeEventListener('workspace:replace-document', handleReplace);
    };
  }, [beginAiTransaction, handleInput, scheduleFinalize, tab.id]);

  /** Применяет HTML из истории, не порождая новую транзакцию. */
  const applyHistoryHtml = useCallback(
    (html: string) => {
      const root = editorRef.current;
      if (!root) return;
      isSyncing.current = true;
      root.innerHTML = html || '<p></p>';
      selectionRef.current = null;
      setWordCount(countWords(root.innerText));
      setRestoredImages(0);
      setAiSnapshot(null);
      aiTxnRef.current = '';
      isSyncing.current = false;
      handleInput();
    },
    [handleInput]
  );

  const restoreVersion = useCallback(
    (html: string) => {
      applyHistoryHtml(html);
      setVersionsOpen(false);
      captureVersion('Восстановлено');
    },
    [applyHistoryHtml, captureVersion]
  );
  /** Возврат документа к состоянию до замены ассистентом. */
  const undoHistory = useCallback(() => {
    const snapshot = aiSnapshot;
    finalizeTransaction();
    const history = historyRef.current;
    const record = history.current;

    // Баннер хранит точный снимок; история нужна, только если она про него же.
    if (!record || (snapshot !== null && record.before !== snapshot)) {
      if (snapshot === null) return;
      applyHistoryHtml(snapshot);
      return;
    }

    const transaction = history.undo();
    if (!transaction) return;
    applyHistoryHtml(transaction.before);
    setHistoryVersion((value) => value + 1);
  }, [aiSnapshot, applyHistoryHtml, finalizeTransaction]);

  const restoreAiSnapshot = undoHistory;

  const redoHistory = useCallback(() => {
    finalizeTransaction();
    const transaction = historyRef.current.redo();
    if (!transaction) return;
    applyHistoryHtml(transaction.after);
    setHistoryVersion((value) => value + 1);
  }, [applyHistoryHtml, finalizeTransaction]);

  const gotoHistory = useCallback(
    (index: number) => {
      finalizeTransaction();
      const transaction = historyRef.current.goto(index);
      if (!transaction) return;
      applyHistoryHtml(transaction.after);
      setHistoryVersion((value) => value + 1);
    },
    [applyHistoryHtml, finalizeTransaction]
  );

  /** Снять текущее оформление с каретки — чтобы кнопки и списки показывали реальное состояние. */
  const refreshFormat = useCallback(() => {
    const root = editorRef.current;
    const sel = window.getSelection();
    if (!root || !sel || sel.rangeCount === 0) return;
    if (!root.contains(sel.anchorNode)) return;

    const query = (command: string) => {
      try {
        return document.queryCommandState(command);
      } catch {
        return false;
      }
    };
    const value = (command: string) => {
      try {
        return document.queryCommandValue(command) || '';
      } catch {
        return '';
      }
    };

    const rawFont = value('fontName').replace(/['"]/g, '');
    const matched = FONT_FAMILIES.find(
      (f) => f.value.split(',')[0].replace(/['"]/g, '').toLowerCase() === rawFont.toLowerCase()
    );

    setFormat({
      bold: query('bold'),
      italic: query('italic'),
      underline: query('underline'),
      strike: query('strikeThrough'),
      superscript: query('superscript'),
      subscript: query('subscript'),
      fontName: matched?.label || rawFont,
      fontSizePt: legacySizeToPt(value('fontSize')),
      foreColor: normalizeColor(value('foreColor')),
      hiliteColor: normalizeColor(value('hiliteColor') || value('backColor')),
      block: (value('formatBlock') || '').toLowerCase(),
    });
  }, []);

  const rememberSelection = () => {
    const root = editorRef.current;
    const sel = window.getSelection();
    if (!root || !sel || sel.rangeCount === 0) return;
    if (!root.contains(sel.anchorNode)) return;
    savedRange.current = sel.getRangeAt(0).cloneRange();
  };

  const restoreSelection = () => {
    const root = editorRef.current;
    const range = savedRange.current;
    root?.focus();
    if (!range) return;
    const sel = window.getSelection();
    if (!sel) return;
    sel.removeAllRanges();
    sel.addRange(range);
  };

  /* ---------------- Форматирование ---------------- */

  const exec = (command: string, value = '') => {
    restoreSelection();
    document.execCommand(command, false, value);
    rememberSelection();
    handleInput();
    refreshFormat();
  };

  const closeMenus = () => {
    setShowColors(null);
    setShowFonts(false);
    setShowSizes(false);
  };

  const applyFont = (family: string) => {
    exec('fontName', family);
    closeMenus();
  };

  const applySize = (pt: number) => {
    // fontSize в execCommand принимает только 1–7. Для точного кегля оборачиваем выделение в span.
    restoreSelection();
    const applied = applyExactFontSize(pt);
    if (!applied) {
      document.execCommand('fontSize', false, String(ptToLegacy(pt)));
    }
    rememberSelection();
    handleInput();
    refreshFormat();
    closeMenus();
  };

  const applyCustomColor = (kind: 'text' | 'highlight', color: string) => {
    exec(kind === 'text' ? 'foreColor' : 'hiliteColor', color);
    closeMenus();
  };

  const handleAIEdit = (promptText: string) => {
    if (!onAskAI) return;
    const selection = window.getSelection()?.toString().trim();
    const context = selection || editorRef.current?.innerText || '';
    if (!context.trim()) return;
    onAskAI(promptText, context);
  };

  /**
   * Просит модель написать оформленный фрагмент и вставляет его туда,
   * где стоит курсор. Выделение при этом заменяется.
   */
  const generateWithAI = () => {
    const task = aiPrompt.trim();
    if (!task || aiBusy) return;

    setAiBusy(true);
    setAiError('');
    pendingAI.current = { range: savedRange.current?.cloneRange() ?? null };

    const selection = savedRange.current?.toString().trim() ?? '';
    const instruction =
      'Напиши текст для документа Word по задаче: «' + task + '».\n' +
      'Ответь ТОЛЬКО готовым текстом в Markdown, без пояснений до и после и без ограждения из тройных кавычек.\n' +
      'Оформи так, чтобы документ выглядел цельно: заголовок через #, подзаголовки через ## и ###, ' +
      'абзацы разделяй пустой строкой, перечисления делай списками (- или 1.), важные слова выделяй **жирным**, ' +
      'цитату оформь через >, при сравнении используй таблицу Markdown. Не злоупотребляй заголовками.';

    onAskAI?.(
      instruction,
      selection ? 'Заменить этот фрагмент, сохранив его смысл, если задача этого требует:\n' + selection : ''
    );
  };

  /* ---------------- Экспорт в HTML ---------------- */

  const exportPdf = async () => {
    const api = window.electronAPI?.exportPdf;
    if (!api) {
      setError('Экспорт в PDF доступен только в настольном приложении');
      return;
    }
    const body = editorRef.current?.innerHTML ?? '';
    const html = `<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8"><title>${tab.name}</title>
<style>
@page { size: A4; margin: ${pageSetup.marginTopMm ?? 20}mm ${pageSetup.marginRightMm ?? 15}mm ${pageSetup.marginBottomMm ?? 20}mm ${pageSetup.marginLeftMm ?? 30}mm; }
body{font-family:${pageSetup.fontFamily ? `"${pageSetup.fontFamily}", serif` : "Calibri, 'Segoe UI', Georgia, serif"};font-size:${pageSetup.fontSizePt ?? 12}pt;line-height:${pageSetup.lineSpacing ?? 1.5};color:#1c2430;max-width:100%;}
table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccc;padding:6px 8px}
blockquote{border-left:3px solid #2563eb;margin:12px 0;padding-left:12px;color:#555}
img{max-width:100%}
</style></head><body>${body}</body></html>`;
    const result = await api(html, `${tab.name.replace(/\.[^.]+$/, '')}.pdf`);
    if (!result.success && !result.canceled) {
      setError(result.error ?? 'Не удалось экспортировать PDF');
    }
  };

  const exportHtml = () => {
    const body = editorRef.current?.innerHTML ?? '';
    const html = `<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8"><title>${tab.name}</title>
<style>body{font-family:Calibri,'Segoe UI',Georgia,serif;font-size:15px;line-height:1.7;max-width:820px;margin:40px auto;padding:0 24px;color:#1c2430}
table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccc;padding:6px 8px}
blockquote{border-left:3px solid #2563eb;margin:12px 0;padding-left:12px;color:#555}
img{max-width:100%}</style></head><body>${body}</body></html>`;
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${tab.name.replace(/\.[^.]+$/, '')}.html`;
    a.click();
    URL.revokeObjectURL(url);
  };

  useEffect(() => {
    const openFind = () => setFindOpen(true);
    const openHistory = () => setHistoryOpen(true);
    const openVersions = () => setVersionsOpen(true);
    window.addEventListener('workspace:word-find', openFind);
    window.addEventListener('workspace:word-history', openHistory);
    window.addEventListener('workspace:word-versions', openVersions);
    return () => {
      window.removeEventListener('workspace:word-find', openFind);
      window.removeEventListener('workspace:word-history', openHistory);
      window.removeEventListener('workspace:word-versions', openVersions);
    };
  }, []);

  if (loading) {
    return (
      <div className="flex h-full items-center justify-center bg-ws-editor">
        <div className="flex flex-col items-center gap-2 text-ws-muted">
          <FileText size={28} className="animate-pulse text-blue-500" />
          <span className="flex items-center gap-2 text-xs">
            <Spinner size={12} />
            Загрузка документа Word…
          </span>
        </div>
      </div>
    );
  }

  const fontLabel = format.fontName || 'Шрифт';
  const sizeLabel = format.fontSizePt || '15';

  const historyRecords = historyRef.current.records;
  const historyIndex = historyRef.current.index;
  const canUndoHistory = historyRef.current.canUndo();
  const canRedoHistory = historyRef.current.canRedo();
  const latestTransaction: DocumentTransaction | null = historyRef.current.current;
  const latestChanges = latestTransaction
    ? latestTransaction.diff.changed + latestTransaction.diff.added + latestTransaction.diff.removed
    : 0;

  // Полотно повторяет поля и базовый шрифт .docx, чтобы редактор совпадал с Word.
  const canvasStyle: React.CSSProperties = {
    fontFamily: pageSetup.fontFamily
      ? `"${pageSetup.fontFamily}", serif`
      : 'Calibri, "Segoe UI", Georgia, serif',
    fontSize: pageSetup.fontSizePt ? `${pageSetup.fontSizePt}pt` : '15px',
    lineHeight: pageSetup.lineSpacing ? String(pageSetup.lineSpacing) : 1.7,
  };
  const mmToCanvasPx = 820 / 210;
  if (pageSetup.marginTopMm) canvasStyle.paddingTop = `${Math.round(pageSetup.marginTopMm * mmToCanvasPx)}px`;
  if (pageSetup.marginRightMm) canvasStyle.paddingRight = `${Math.round(pageSetup.marginRightMm * mmToCanvasPx)}px`;
  if (pageSetup.marginBottomMm) canvasStyle.paddingBottom = `${Math.round(pageSetup.marginBottomMm * mmToCanvasPx)}px`;
  if (pageSetup.marginLeftMm) canvasStyle.paddingLeft = `${Math.round(pageSetup.marginLeftMm * mmToCanvasPx)}px`;

  return (
    <div className="flex h-full w-full flex-col bg-ws-editor text-ws">
      {/* Лента инструментов */}
      <div
        className="flex shrink-0 flex-wrap items-center justify-between gap-2 border-b border-ws bg-ws-surface px-3 py-1.5 text-xs"
        onMouseDown={(e) => {
          // Кнопки не должны сбрасывать выделение в документе.
          const target = e.target as HTMLElement;
          if (target.closest('input, select, textarea')) return;
          e.preventDefault();
          rememberSelection();
        }}
      >
        <div className="flex flex-wrap items-center gap-0.5">
          <span className="mr-1 flex items-center gap-1.5 rounded border border-blue-500/40 bg-blue-500/10 px-2 py-0.5 font-medium text-blue-500">
            <FileText size={13} />
            Word
          </span>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <IconButton title="Отменить (Ctrl+Z)" onClick={() => exec('undo')}>
            <Undo2 size={14} />
          </IconButton>
          <IconButton title="Повторить (Ctrl+Y)" onClick={() => exec('redo')}>
            <Redo2 size={14} />
          </IconButton>
          <IconButton
            title="История правок ассистента"
            onClick={() => setHistoryOpen((value) => !value)}
          >
            <HistoryIcon size={14} />
          </IconButton>
          <IconButton
            title="Структура документа"
            onClick={() => setShowOutline((value) => !value)}
          >
            <ListTree size={14} />
          </IconButton>
          <IconButton
            title="Найти и заменить"
            onClick={() => setFindOpen((value) => !value)}
          >
            <Search size={14} />
          </IconButton>
          <IconButton
            title="Версии документа"
            onClick={() => setVersionsOpen((value) => !value)}
          >
            <Clock size={14} />
          </IconButton>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <div className="relative">
            <button
              type="button"
              title="Шрифт"
              onClick={() => {
                setShowFonts((v) => !v);
                setShowSizes(false);
                setShowColors(null);
              }}
              className={cn(
                'flex h-7 w-[148px] items-center justify-between gap-1 rounded border border-ws bg-ws-elevated px-2 text-[11px] text-ws hover:border-ws-accent',
                showFonts && 'border-ws-accent'
              )}
            >
              <span className="truncate" style={{ fontFamily: fontFamilyOf(fontLabel) }}>
                {fontLabel}
              </span>
              <ChevronDown size={12} className="shrink-0 text-ws-faint" />
            </button>
            {showFonts && (
              <MenuPanel onClose={() => setShowFonts(false)} width="w-56">
                {FONT_FAMILIES.map((font) => (
                  <button
                    key={font.label}
                    type="button"
                    onClick={() => applyFont(font.value)}
                    className={cn(
                      'flex w-full items-center rounded px-2 py-1 text-left text-[13px] hover:bg-ws-hover',
                      format.fontName === font.label && 'bg-ws-accent-soft text-ws-accent'
                    )}
                    style={{ fontFamily: font.value }}
                  >
                    {font.label}
                  </button>
                ))}
              </MenuPanel>
            )}
          </div>

          <div className="relative">
            <div
              className={cn(
                'flex h-7 items-center rounded border border-ws bg-ws-elevated',
                showSizes && 'border-ws-accent'
              )}
            >
              <input
                title="Размер шрифта, пт"
                value={sizeLabel}
                onChange={(e) =>
                  setFormat((prev) => ({ ...prev, fontSizePt: e.target.value.replace(/[^0-9.]/g, '') }))
                }
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    const pt = Number(sizeLabel);
                    if (pt >= 1 && pt <= 200) applySize(pt);
                  }
                }}
                onBlur={() => {
                  const pt = Number(sizeLabel);
                  if (pt >= 1 && pt <= 200 && String(pt) !== format.fontSizePt) applySize(pt);
                }}
                className="w-9 bg-transparent px-1.5 text-center text-[11px] text-ws outline-none"
              />
              <button
                type="button"
                title="Размер шрифта"
                onClick={() => {
                  setShowSizes((v) => !v);
                  setShowFonts(false);
                  setShowColors(null);
                }}
                className="flex h-full items-center border-l border-ws px-1 text-ws-faint hover:text-ws"
              >
                <ChevronDown size={12} />
              </button>
            </div>
            {showSizes && (
              <MenuPanel onClose={() => setShowSizes(false)} width="w-20">
                {FONT_SIZES.map((pt) => (
                  <button
                    key={pt}
                    type="button"
                    onClick={() => applySize(pt)}
                    className={cn(
                      'flex w-full items-center rounded px-2 py-1 text-left text-[12px] hover:bg-ws-hover',
                      format.fontSizePt === String(pt) && 'bg-ws-accent-soft text-ws-accent'
                    )}
                  >
                    {pt}
                  </button>
                ))}
              </MenuPanel>
            )}
          </div>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <IconButton title="Жирный (Ctrl+B)" active={format.bold} onClick={() => exec('bold')}>
            <Bold size={14} />
          </IconButton>
          <IconButton title="Курсив (Ctrl+I)" active={format.italic} onClick={() => exec('italic')}>
            <Italic size={14} />
          </IconButton>
          <IconButton title="Подчёркнутый (Ctrl+U)" active={format.underline} onClick={() => exec('underline')}>
            <Underline size={14} />
          </IconButton>
          <IconButton title="Зачёркнутый" active={format.strike} onClick={() => exec('strikeThrough')}>
            <Strikethrough size={14} />
          </IconButton>
          <IconButton title="Надстрочный" active={format.superscript} onClick={() => exec('superscript')}>
            <Superscript size={14} />
          </IconButton>
          <IconButton title="Подстрочный" active={format.subscript} onClick={() => exec('subscript')}>
            <Subscript size={14} />
          </IconButton>

          <div className="relative">
            <IconButton
              title="Цвет текста"
              active={showColors === 'text'}
              onClick={() => {
                setShowColors((v) => (v === 'text' ? null : 'text'));
                setShowFonts(false);
                setShowSizes(false);
              }}
            >
              <span className="flex flex-col items-center leading-none">
                <Baseline size={13} />
                <span
                  className="mt-0.5 h-[3px] w-3.5 rounded-sm"
                  style={{ background: format.foreColor || '#2563eb' }}
                />
              </span>
            </IconButton>
            {showColors === 'text' && (
              <ColorPicker
                colors={TEXT_COLORS}
                current={format.foreColor}
                allowCustom
                onPick={(color) => applyCustomColor('text', color)}
                onClose={() => setShowColors(null)}
              />
            )}
          </div>

          <div className="relative">
            <IconButton
              title="Цвет выделения"
              active={showColors === 'highlight'}
              onClick={() => {
                setShowColors((v) => (v === 'highlight' ? null : 'highlight'));
                setShowFonts(false);
                setShowSizes(false);
              }}
            >
              <span className="flex flex-col items-center leading-none">
                <Highlighter size={13} />
                <span
                  className="mt-0.5 h-[3px] w-3.5 rounded-sm"
                  style={{ background: format.hiliteColor || '#fde047' }}
                />
              </span>
            </IconButton>
            {showColors === 'highlight' && (
              <ColorPicker
                colors={HIGHLIGHT_COLORS}
                current={format.hiliteColor}
                allowCustom
                onPick={(color) => applyCustomColor('highlight', color)}
                onClose={() => setShowColors(null)}
              />
            )}
          </div>

          <IconButton title="Очистить форматирование" onClick={() => exec('removeFormat')}>
            <RemoveFormatting size={14} />
          </IconButton>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <IconButton title="Заголовок 1" active={format.block === 'h1'} onClick={() => exec('formatBlock', '<h1>')}>
            <Heading1 size={14} />
          </IconButton>
          <IconButton title="Заголовок 2" active={format.block === 'h2'} onClick={() => exec('formatBlock', '<h2>')}>
            <Heading2 size={14} />
          </IconButton>
          <IconButton title="Заголовок 3" active={format.block === 'h3'} onClick={() => exec('formatBlock', '<h3>')}>
            <Heading3 size={14} />
          </IconButton>
          <IconButton title="Обычный текст" active={format.block === 'p' || format.block === 'div'} onClick={() => exec('formatBlock', '<p>')}>
            <span className="px-0.5 text-[10px] font-semibold">P</span>
          </IconButton>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <IconButton title="Маркированный список" onClick={() => exec('insertUnorderedList')}>
            <List size={14} />
          </IconButton>
          <IconButton title="Нумерованный список" onClick={() => exec('insertOrderedList')}>
            <ListOrdered size={14} />
          </IconButton>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <IconButton title="По левому краю" onClick={() => exec('justifyLeft')}>
            <AlignLeft size={14} />
          </IconButton>
          <IconButton title="По центру" onClick={() => exec('justifyCenter')}>
            <AlignCenter size={14} />
          </IconButton>
          <IconButton title="По правому краю" onClick={() => exec('justifyRight')}>
            <AlignRight size={14} />
          </IconButton>
          <IconButton title="По ширине" onClick={() => exec('justifyFull')}>
            <AlignJustify size={14} />
          </IconButton>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={exportHtml}
            title="Экспортировать документ в HTML"
            className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted transition-colors hover:border-ws-accent hover:text-ws-accent"
          >
            <Download size={11} />
            HTML
          </button>

          <div className="relative">
            <button
              onClick={() => {
                rememberSelection();
                setShowAI((open) => !open);
                setAiError('');
              }}
              title="Сгенерировать оформленный текст и вставить в документ"
              className={cn(
                'flex items-center gap-1.5 rounded border border-blue-500/40 bg-blue-500/10 px-2.5 py-1 text-blue-500 transition-colors hover:bg-blue-500/20',
                showAI && 'bg-blue-500/20'
              )}
            >
              <Sparkles size={12} />
              {aiBusy ? 'Пишу…' : 'AI-текст'}
            </button>
            {showAI && (
              <AIPrompt
                value={aiPrompt}
                busy={aiBusy}
                error={aiError}
                onChange={setAiPrompt}
                onSubmit={generateWithAI}
                onRewrite={() =>
                  handleAIEdit(
                    'Перепиши этот фрагмент в безупречном деловом стиле: устрани ошибки, улучши формулировки, сохрани смысл и структуру.'
                  )
                }
                onClose={() => setShowAI(false)}
              />
            )}
          </div>

          <div className="h-3.5 w-px border-l border-ws" />

          <button
            onClick={() => void exportPdf()}
            title="Экспорт в PDF"
            className="flex items-center gap-1 rounded border border-ws bg-ws-elevated px-2 py-1 text-ws hover:bg-ws-hover"
          >
            <Download size={12} />
            PDF
          </button>
          <button
            onClick={flushSave}
            disabled={!tab.isDirty}
            title="Сохранить (Ctrl+S)"
            className={cn(
              'flex items-center gap-1 rounded px-2.5 py-1 font-medium transition-colors',
              tab.isDirty ? 'bg-ws-accent text-white hover:opacity-90' : 'bg-ws-panel-alt text-ws-faint'
            )}
          >
            <Save size={12} />
            Сохранить
          </button>
        </div>
      </div>

      {error && (
        <div className="flex shrink-0 items-center gap-2 border-b border-amber-500/40 bg-amber-500/10 px-3 py-2 text-[11px] text-ws-warning">
          <AlertTriangle size={12} />
          {error}
        </div>
      )}

      <div
        data-breadcrumb="1"
        className="flex shrink-0 items-center gap-1 border-b border-ws bg-ws-surface px-3 py-1 text-[10px] text-ws-faint"
      >
        <span className="truncate">{tab.name}</span>
        {breadcrumb && (
          <>
            <span>›</span>
            <span className="truncate text-ws">{breadcrumb}</span>
          </>
        )}
      </div>

      {versionsOpen && (
        <div className="shrink-0 border-b border-ws bg-ws-surface px-3 py-2 text-[11px] text-ws">
          <div className="mb-1.5 flex items-center gap-2">
            <Clock size={12} className="text-blue-400" />
            <span className="flex-1 font-medium">Версии документа · {versionsRef.current.length}</span>
            <button
              onClick={() => {
                versionsRef.current = [];
                versionStore.set(tab.id, []);
                setVersionsTick((value) => value + 1);
              }}
              className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 text-[10px] hover:bg-ws-hover"
            >
              Очистить
            </button>
            <button
              onClick={() => setVersionsOpen(false)}
              title="Закрыть версии"
              className="rounded px-1 text-[11px] text-ws-faint hover:bg-ws-hover hover:text-ws"
            >
              ×
            </button>
          </div>
          {versionsRef.current.length === 0 ? (
            <div className="text-ws-faint">Версий пока нет — они создаются при сохранении и правках AI.</div>
          ) : (
            <div className="max-h-40 overflow-y-auto">
              {versionsRef.current.map((version) => (
                <div
                  key={version.id}
                  className="flex items-center gap-2 rounded px-1.5 py-1 hover:bg-ws-hover"
                >
                  <span className="font-mono text-ws-faint">
                    {new Date(version.timestamp).toLocaleTimeString('ru-RU', {
                      hour: '2-digit',
                      minute: '2-digit',
                      second: '2-digit',
                    })}
                  </span>
                  <span className="flex-1 truncate">{version.label}</span>
                  <button
                    onClick={() => restoreVersion(version.html)}
                    className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 text-[10px]"
                  >
                    Восстановить
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {findOpen && (
        <div className="flex shrink-0 flex-wrap items-center gap-1.5 border-b border-ws bg-ws-surface px-3 py-1.5 text-[11px]">
          <Search size={12} className="text-ws-faint" />
          <input
            data-find-input="1"
            value={findText}
            onChange={(event) => setFindText(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') findNext();
            }}
            placeholder="Найти"
            className="w-40 rounded border border-ws bg-ws-elevated px-1.5 py-0.5 text-ws placeholder:text-ws-faint focus:border-ws-accent focus:outline-none"
          />
          <input
            data-replace-input="1"
            value={replaceWith}
            onChange={(event) => setReplaceWith(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') replaceCurrent();
            }}
            placeholder="Заменить на"
            className="w-40 rounded border border-ws bg-ws-elevated px-1.5 py-0.5 text-ws placeholder:text-ws-faint focus:border-ws-accent focus:outline-none"
          />
          <button
            data-find-next="1"
            onClick={findNext}
            className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 hover:bg-ws-hover"
          >
            Далее
          </button>
          <button
            data-find-replace="1"
            onClick={replaceCurrent}
            className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 hover:bg-ws-hover"
          >
            Заменить
          </button>
          <button
            data-find-replace-all="1"
            onClick={replaceAll}
            className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 hover:bg-ws-hover"
          >
            Заменить все
          </button>
          <span className="text-ws-faint">{findMessage}</span>
          <button
            onClick={() => {
              setFindOpen(false);
              highlightFindRange(null);
            }}
            title="Закрыть поиск"
            className="ml-auto rounded px-1 text-ws-faint hover:bg-ws-hover hover:text-ws"
          >
            ×
          </button>
        </div>
      )}

      {historyOpen && (
        <div className="shrink-0 border-b border-ws bg-ws-surface px-3 py-2 text-[11px] text-ws">
          <div className="mb-1.5 flex items-center gap-2">
            <HistoryIcon size={12} className="text-blue-400" />
            <span className="flex-1 font-medium">История ассистента · {historyRecords.length}</span>
            <button
              disabled={!canUndoHistory}
              onClick={undoHistory}
              className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 text-[10px] disabled:opacity-40"
            >
              Отменить
            </button>
            <button
              disabled={!canRedoHistory}
              onClick={redoHistory}
              className="rounded border border-ws bg-ws-elevated px-1.5 py-0.5 text-[10px] disabled:opacity-40"
            >
              Повторить
            </button>
            <button
              onClick={() => setHistoryOpen(false)}
              className="rounded px-1 text-[11px] text-ws-faint hover:bg-ws-hover hover:text-ws"
              title="Закрыть историю"
            >
              ×
            </button>
          </div>
          {historyRecords.length === 0 ? (
            <div className="text-ws-faint">Пока нет правок ассистента.</div>
          ) : (
            <div className="max-h-40 overflow-y-auto">
              {historyRecords.map((transaction, index) => {
                const total =
                  transaction.diff.changed +
                  transaction.diff.added +
                  transaction.diff.removed;
                const preview = transaction.diff.blocks
                  .slice(0, 2)
                  .map((block) => block.after || block.before || '')
                  .filter(Boolean)
                  .join(' · ');
                return (
                  <button
                    key={`${transaction.id}-${index}`}
                    onClick={() => gotoHistory(index)}
                    className={cn(
                      'flex w-full items-center gap-2 rounded px-1.5 py-1 text-left text-[10px] hover:bg-ws-hover',
                      index === historyIndex && 'bg-ws-accent-soft text-ws-accent'
                    )}
                  >
                    <span className="font-mono text-ws-faint">
                      {new Date(transaction.timestamp).toLocaleTimeString('ru-RU', {
                        hour: '2-digit',
                        minute: '2-digit',
                      })}
                    </span>
                    <span className="shrink-0">{transaction.label}</span>
                    <span className="shrink-0 text-ws-faint">блоков: {total}</span>
                    <span className="truncate text-ws-faint">{preview}</span>
                  </button>
                );
              })}
            </div>
          )}
        </div>
      )}

      {tab.extension === 'doc' && (
        <div className="flex shrink-0 items-center gap-2 border-b border-amber-500/40 bg-amber-500/10 px-3 py-1.5 text-[11px] text-ws-warning">
          <AlertTriangle size={11} />
          <span className="flex-1">
            Формат .doc: текст извлечён без оформления. Сохранение создаст .docx рядом — исходный
            .doc останется без изменений.
          </span>
        </div>
      )}

      {aiSnapshot !== null && (
        <div className="flex shrink-0 items-center gap-2 border-b border-blue-500/40 bg-blue-500/10 px-3 py-1.5 text-[11px] text-ws">
          <Sparkles size={11} className="text-blue-400" />
          <span className="flex-1">
            Документ изменён ассистентом
            {latestChanges > 0 && ` · блоков: ${latestChanges}`}
            {restoredImages > 0 && ` · изображений сохранено: ${restoredImages}`}
          </span>
          <button
            onClick={() => setHistoryOpen(true)}
            className="flex items-center gap-1 rounded border border-ws bg-ws-surface px-1.5 py-0.5 text-[10px] text-ws hover:bg-ws-hover"
          >
            <HistoryIcon size={10} />
            История
          </button>
          <button
            onClick={restoreAiSnapshot}
            className="flex items-center gap-1 rounded border border-ws bg-ws-surface px-1.5 py-0.5 text-[10px] text-ws hover:bg-ws-hover"
          >
            <Undo2 size={10} />
            Вернуть прежний текст
          </button>
          <button
            onClick={() => setAiSnapshot(null)}
            title="Скрыть уведомление"
            className="rounded px-1 py-0.5 text-[10px] text-ws-faint hover:bg-ws-hover hover:text-ws"
          >
            Скрыть
          </button>
        </div>
      )}

      {/* Полотно документа и панель структуры */}
      <div className="flex min-h-0 flex-1">
        {showOutline && (
          <aside
            data-outline-panel="1"
            className="w-52 shrink-0 overflow-y-auto border-r border-ws bg-ws-surface p-2 text-[11px]"
          >
            <div className="mb-1 flex items-center gap-1 px-1 font-medium text-ws-muted">
              <ListTree size={11} />
              Структура
            </div>
            {outline.length === 0 ? (
              <div className="px-1 text-ws-faint">Заголовков нет</div>
            ) : (
              outline.map((item) => (
                <button
                  key={`${item.id}-${item.level}-${item.text}`}
                  data-outline-item="1"
                  onClick={() => scrollToOutline(item)}
                  className="block w-full truncate rounded px-1 py-0.5 text-left text-ws hover:bg-ws-hover"
                  style={{ paddingLeft: `${(item.level - 1) * 10 + 4}px` }}
                  title={item.text}
                >
                  {item.text}
                </button>
              ))
            )}
          </aside>
        )}
        <div ref={zoomScrollRef} className="min-h-0 flex-1 overflow-y-auto bg-ws-editor-alt p-6">
          <div
            ref={editorRef}
            contentEditable
            suppressContentEditableWarning
            onInput={handleInput}
            onMouseUp={() => {
              rememberSelection();
              refreshFormat();
            }}
            onKeyUp={() => {
              rememberSelection();
              refreshFormat();
            }}
            onBlur={handleInput}
            className="doc-canvas mx-auto min-h-[1000px] w-full max-w-[820px] rounded-sm border border-ws bg-ws-surface p-14 text-ws shadow-ws focus:outline-none select-text"
            style={{ ...canvasStyle, zoom }}
          />
        </div>
      </div>

      {/* Строка состояния */}
      <div className="flex shrink-0 items-center justify-between border-t border-ws bg-ws-surface px-3 py-1 text-[10px] text-ws-faint">
        <span>DOCX · {tab.name}</span>
        <div className="flex items-center gap-3">
          {format.fontName && <span>{format.fontName}</span>}
          {format.fontSizePt && <span>{format.fontSizePt} пт</span>}
          <span>Масштаб: {Math.round(zoom * 100)}%</span>
          <span>Слов: {wordCount.toLocaleString('ru-RU')}</span>
          {tab.isDirty && <span className="text-ws-warning">● Изменено</span>}
        </div>
      </div>
    </div>
  );
};

/* ============================================================
   Вспомогательные компоненты
   ============================================================ */

interface AnchoredImage {
  src: string;
  alt: string;
  /** Начало текста блока, в котором стояла картинка, — якорь для возврата. */
  anchor: string;
  /** Порядковый номер блока на случай, если текст переписан. */
  index: number;
}

/** Запоминает картинки документа вместе с местом, где они стояли. */
function collectAnchoredImages(root: HTMLElement): AnchoredImage[] {
  const blocks = [...root.children];
  return [...root.querySelectorAll('img')]
    .map((img) => {
      const host = blocks.find((block) => block.contains(img)) ?? null;
      return {
        src: img.getAttribute('src') ?? '',
        alt: img.getAttribute('alt') ?? '',
        anchor: (host?.textContent ?? '').replace(/\s+/g, ' ').trim().slice(0, 40).toLowerCase(),
        index: host ? blocks.indexOf(host) : -1,
      };
    })
    .filter((image) => image.src);
}

/** Возвращает картинки в переписанный документ и сообщает их число. */
function restoreAnchoredImages(root: HTMLElement, images: AnchoredImage[]): number {
  if (images.length === 0) return 0;
  const blocks = [...root.children];
  const normalize = (value: string) => value.replace(/\s+/g, ' ').trim().toLowerCase();
  let restored = 0;

  for (const image of images) {
    const element = document.createElement('img');
    element.src = image.src;
    if (image.alt) element.alt = image.alt;

    let target: Element | null = null;
    if (image.anchor.length >= 8 && blocks.length > 0) {
      target = blocks.find((block) => normalize(block.textContent ?? '').includes(image.anchor)) ?? null;
    }
    if (!target && image.index >= 0 && blocks.length > 0) {
      target = blocks[Math.min(image.index, blocks.length - 1)];
    }

    if (target) target.insertAdjacentElement('afterend', element);
    else root.appendChild(element);
    restored += 1;
  }
  return restored;
}


const AIPrompt: React.FC<{
  value: string;
  busy: boolean;
  error: string;
  onChange: (value: string) => void;
  onSubmit: () => void;
  onRewrite: () => void;
  onClose: () => void;
}> = ({ value, busy, error, onChange, onSubmit, onRewrite, onClose }) => (
  <>
    <div className="fixed inset-0 z-40" onMouseDown={onClose} />
    <div
      className="animate-scale-in absolute top-full right-0 z-50 mt-1 w-80 rounded-lg border border-ws bg-ws-surface p-2.5 shadow-ws"
      onMouseDown={(event) => event.stopPropagation()}
    >
      <div className="mb-1.5 flex items-center gap-1.5 text-[11px] font-semibold text-ws">
        <Sparkles size={12} className="text-blue-500" />
        Сгенерировать текст
      </div>
      <textarea
        autoFocus
        value={value}
        disabled={busy}
        placeholder="Например: деловое письмо клиенту о переносе поставки на 20 октября, вежливый тон"
        onChange={(event) => onChange(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
            event.preventDefault();
            onSubmit();
          }
        }}
        className="h-20 w-full resize-none rounded border border-ws bg-ws-elevated px-2 py-1.5 text-xs text-ws outline-none placeholder:text-ws-faint focus:border-ws-accent"
      />
      {error && <p className="mt-1 text-[11px] text-ws-warning">{error}</p>}
      <div className="mt-2 flex items-center justify-between gap-2">
        <button
          type="button"
          onClick={onRewrite}
          className="text-[11px] text-ws-muted underline-offset-2 hover:text-ws hover:underline"
        >
          Улучшить выделенное
        </button>
        <button
          type="button"
          disabled={busy || !value.trim()}
          onClick={onSubmit}
          className="flex items-center gap-1 rounded bg-ws-accent px-2.5 py-1 text-[11px] font-medium text-white disabled:opacity-45"
        >
          {busy ? <Spinner size={11} /> : <Sparkles size={11} />}
          Вставить
        </button>
      </div>
      <p className="mt-1.5 text-[10px] text-ws-faint">Ctrl+Enter — вставить в позицию курсора</p>
    </div>
  </>
);

const MenuPanel: React.FC<{
  children: React.ReactNode;
  onClose: () => void;
  width: string;
}> = ({ children, onClose, width }) => (
  <>
    <div className="fixed inset-0 z-40" onMouseDown={onClose} />
    <div
      className={cn(
        'animate-scale-in absolute top-full left-0 z-50 mt-1 max-h-72 overflow-y-auto rounded-lg border border-ws bg-ws-surface p-1 shadow-ws',
        width
      )}
    >
      {children}
    </div>
  </>
);

const ColorPicker: React.FC<{
  colors: string[];
  current?: string;
  allowCustom?: boolean;
  onPick: (color: string) => void;
  onClose: () => void;
}> = ({ colors, current, allowCustom, onPick, onClose }) => (
  <>
    <div className="fixed inset-0 z-40" onMouseDown={onClose} />
    <div className="animate-scale-in absolute top-full left-0 z-50 mt-1 w-[168px] rounded-lg border border-ws bg-ws-surface p-2 shadow-ws">
      <div className="grid grid-cols-8 gap-1">
        {colors.map((color) => (
          <button
            key={color}
            type="button"
            onClick={() => onPick(color)}
            title={color === 'transparent' ? 'Убрать выделение' : color}
            style={{ background: color === 'transparent' ? 'transparent' : color }}
            className={cn(
              'h-4 w-4 rounded border border-ws transition-transform hover:scale-110',
              current && colorsClose(current, color) && 'ring-2 ring-ws-accent ring-offset-1',
              color === 'transparent' &&
                'bg-[linear-gradient(45deg,#ccc_25%,transparent_25%,transparent_75%,#ccc_75%),linear-gradient(45deg,#ccc_25%,#fff_25%,#fff_75%,#ccc_75%)] bg-[length:8px_8px] bg-[position:0_0,4px_4px]'
            )}
          />
        ))}
      </div>
      {allowCustom && (
        <label className="mt-2 flex cursor-pointer items-center gap-2 border-t border-ws pt-2 text-[11px] text-ws-muted">
          <Palette size={12} />
          Другой цвет
          <input
            type="color"
            className="ml-auto h-5 w-8 cursor-pointer border-0 bg-transparent p-0"
            onMouseDown={(e) => e.stopPropagation()}
            onChange={(e) => onPick(e.target.value)}
          />
        </label>
      )}
    </div>
  </>
);

function fontFamilyOf(label: string): string | undefined {
  return FONT_FAMILIES.find((f) => f.label === label)?.value;
}

/** execCommand('fontSize') понимает только ступени 1–7. */
function ptToLegacy(pt: number): number {
  if (pt <= 9) return 1;
  if (pt <= 11) return 2;
  if (pt <= 13) return 3;
  if (pt <= 16) return 4;
  if (pt <= 22) return 5;
  if (pt <= 32) return 6;
  return 7;
}

function legacySizeToPt(raw: string): string {
  const map: Record<string, string> = { '1': '8', '2': '10', '3': '12', '4': '14', '5': '18', '6': '24', '7': '36' };
  if (map[raw]) return map[raw];
  const px = raw.match(/([\d.]+)px/);
  if (px) return String(Math.round(Number(px[1]) * 0.75));
  const pt = raw.match(/([\d.]+)pt/);
  if (pt) return String(Math.round(Number(pt[1])));
  return raw.replace(/[^0-9.]/g, '');
}

function normalizeColor(raw: string): string {
  if (!raw || raw === 'transparent') return '';
  const rgb = raw.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (!rgb) return raw;
  const hex = [rgb[1], rgb[2], rgb[3]]
    .map((n) => Number(n).toString(16).padStart(2, '0'))
    .join('');
  return `#${hex}`;
}

function colorsClose(a: string, b: string): boolean {
  return a.replace(/\s/g, '').toLowerCase() === b.replace(/\s/g, '').toLowerCase();
}

/**
 * Ставит точный размер в пунктах. execCommand умеет только 7 ступеней,
 * поэтому при выделении оборачиваем его в span, а без выделения — правим span каретки.
 */
function applyExactFontSize(pt: number): boolean {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return false;
  const range = sel.getRangeAt(0);

  if (!range.collapsed) {
    const span = document.createElement('span');
    span.style.fontSize = `${pt}pt`;
    try {
      range.surroundContents(span);
    } catch {
      const fragment = range.extractContents();
      span.appendChild(fragment);
      range.insertNode(span);
    }
    sel.removeAllRanges();
    const next = document.createRange();
    next.selectNodeContents(span);
    sel.addRange(next);
    return true;
  }

  const node = range.startContainer;
  const el = node instanceof Element ? node : node.parentElement;
  const span = el?.closest('span');
  if (span && span.style.fontSize) {
    span.style.fontSize = `${pt}pt`;
    return true;
  }
  return false;
}


/**
 * Вставляет готовый HTML в запомненное место. Выделение заменяется,
 * без выделения фрагмент добавляется в позицию курсора или в конец.
 *
 * execCommand возвращает true и при потере выделения ничего не делает,
 * поэтому результат проверяем по фактическому изменению DOM: иначе кнопка
 * «Вставить в документ» молча не срабатывала.
 */
function insertHtmlAt(root: HTMLDivElement | null, range: Range | null, html: string) {
  if (!root) return;
  root.focus();

  const selection = window.getSelection();
  // Диапазон, сохранённый на самом корне, после пересборки innerHTML
  // указывает на старый дочерний узел и вставляет текст в начало. Такой
  // диапазон считаем потерянным.
  const validRange =
    range && range.startContainer !== root && root.contains(range.startContainer)
      ? range
      : null;

  if (selection) {
    selection.removeAllRanges();
    if (validRange) {
      selection.addRange(validRange);
    } else {
      // Курсор потерян (клик по кнопке в чате) — вставляем в конец, а не
      // в начало документа, куда Chrome ставит каретку после focus().
      const end = document.createRange();
      end.selectNodeContents(root);
      end.collapse(false);
      selection.addRange(end);
    }
  }

  const before = root.innerHTML;
  const inserted = document.execCommand('insertHTML', false, html);
  if (!inserted || root.innerHTML === before) {
    if (validRange) {
      const fragment = validRange.createContextualFragment(html);
      const lastInserted = fragment.lastChild;
      validRange.deleteContents();
      validRange.insertNode(fragment);
      if (selection && lastInserted) {
        const caret = document.createRange();
        caret.setStartAfter(lastInserted);
        caret.collapse(true);
        selection.removeAllRanges();
        selection.addRange(caret);
      }
    } else {
      root.insertAdjacentHTML('beforeend', html);
      if (selection) {
        const caret = document.createRange();
        caret.selectNodeContents(root);
        caret.collapse(false);
        selection.removeAllRanges();
        selection.addRange(caret);
      }
    }
  }
}
function countWords(text: string): number {
  const trimmed = text.trim();
  if (!trimmed) return 0;
  return trimmed.split(/\s+/).length;
}

/**
 * Пытается декодировать текст из base64 или сырой строки.
 */
function tryDecodeText(raw: string): string | null {
  if (!raw) return null;
  try {
    const clean = raw.replace(/\s+/g, '');
    const binString = atob(clean);
    const bytes = Uint8Array.from(binString, (c) => c.charCodeAt(0));
    const decoded = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
    if (decoded && decoded.trim() && !decoded.includes('\uFFFD') && /[\w\u0400-\u04FF\s<>]/.test(decoded)) {
      return decoded;
    }
  } catch {
    // не base64
  }

  if (typeof raw === 'string' && /[\w\u0400-\u04FF\s#*<>]/.test(raw.slice(0, 200))) {
    return raw;
  }

  return null;
}