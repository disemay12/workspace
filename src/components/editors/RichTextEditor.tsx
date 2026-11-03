import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { OpenTab } from '../../types';
import {
  Heading1,
  Heading2,
  Heading3,
  Bold,
  Italic,
  Strikethrough,
  Code,
  List,
  ListOrdered,
  CheckSquare,
  Quote,
  Table as TableIcon,
  Sparkles,
  Save,
  Eye,
  FileEdit,
  Columns,
  Minus,
  Lightbulb,
  Sigma,
  Download,
  Command,
  ChevronRight,
} from 'lucide-react';
import { cn } from '../../lib/utils';
import { IconButton } from '../ui';
import { renderRichMarkdown, renderMermaidBlocks } from '../../lib/markdown';

interface Props {
  tab: OpenTab;
  onChange: (newContent: string) => void;
  onSave: () => void;
  onAskAI?: (prompt: string, contextText: string) => void;
  /** Вызов инлайн-копилота (Ctrl+K) — правка выделения прямо в документе. */
  onRequestInlineEdit?: () => void;
  /** Сохранение выбранного режима просмотра в состоянии вкладки. */
  onViewModeChange?: (mode: 'source' | 'split' | 'visual') => void;
}

type ViewMode = 'source' | 'split' | 'visual';

/** Слэш-команды блочного редактора (Этап 3.1 плана). */
const SLASH_COMMANDS = [
  { id: 'h1', label: 'Заголовок 1', icon: <Heading1 size={13} />, insert: '# ', block: true },
  { id: 'h2', label: 'Заголовок 2', icon: <Heading2 size={13} />, insert: '## ', block: true },
  { id: 'h3', label: 'Заголовок 3', icon: <Heading3 size={13} />, insert: '### ', block: true },
  { id: 'bullet', label: 'Маркированный список', icon: <List size={13} />, insert: '- ', block: true },
  { id: 'numbered', label: 'Нумерованный список', icon: <ListOrdered size={13} />, insert: '1. ', block: true },
  { id: 'todo', label: 'Чек-лист задач', icon: <CheckSquare size={13} />, insert: '- [ ] ', block: true },
  { id: 'quote', label: 'Цитата', icon: <Quote size={13} />, insert: '> ', block: true },
  { id: 'callout', label: 'Выноска', icon: <Lightbulb size={13} />, insert: '> [!NOTE]\n> ', block: true },
  { id: 'code', label: 'Блок кода', icon: <Code size={13} />, insert: '```\n', block: true },
  { id: 'table', label: 'Таблица', icon: <TableIcon size={13} />, insert: '\n| Колонка 1 | Колонка 2 |\n| --- | --- |\n| Значение | Значение |\n', block: true },
  { id: 'divider', label: 'Разделитель', icon: <Minus size={13} />, insert: '\n---\n', block: true },
  { id: 'math', label: 'Формула (LaTeX)', icon: <Sigma size={13} />, insert: '$$\nE = mc^2\n$$\n', block: true },
  { id: 'ai', label: 'Спросить AI', icon: <Sparkles size={13} />, insert: '', block: false, ai: true },
];

export const RichTextEditor: React.FC<Props> = ({
  tab,
  onChange,
  onSave,
  onAskAI,
  onRequestInlineEdit,
  onViewModeChange,
}) => {
  const [viewMode, setViewMode] = useState<ViewMode>(tab.viewMode || 'split');
  const [showSlash, setShowSlash] = useState(false);
  const [slashFilter, setSlashFilter] = useState('');
  const [slashIndex, setSlashIndex] = useState(0);
  const [slashAnchor, setSlashAnchor] = useState<{ top: number; left: number } | null>(null);
  const [scrollSync, setScrollSync] = useState(true);

  const changeViewMode = useCallback(
    (mode: ViewMode) => {
      setViewMode(mode);
      onViewModeChange?.(mode);
    },
    [onViewModeChange]
  );

  // Компонент переиспользуется между вкладками — режим берём из новой вкладки,
  // чтобы Markdown всегда открывался в предпросмотре.
  useEffect(() => {
    setViewMode(tab.viewMode || 'split');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  const syncingRef = useRef(false);

  /* ---------------- Вставка блока ---------------- */

  const insertText = useCallback(
    (before: string, after = '', placeholder = '') => {
      const textarea = textareaRef.current;
      if (!textarea) return;

      const start = textarea.selectionStart;
      const end = textarea.selectionEnd;
      const selected = tab.content.slice(start, end) || placeholder;
      const replacement = `${before}${selected}${after}`;

      const next = tab.content.slice(0, start) + replacement + tab.content.slice(end);
      onChange(next);

      window.setTimeout(() => {
        textarea.focus();
        textarea.setSelectionRange(start + before.length, start + before.length + selected.length);
      }, 10);
    },
    [tab.content, onChange]
  );

  /** Вставка блочной команды с учётом начала строки. */
  const insertBlock = useCallback(
    (template: string) => {
      const textarea = textareaRef.current;
      if (!textarea) return;

      const pos = textarea.selectionStart;
      const lineStart = tab.content.lastIndexOf('\n', pos - 1) + 1;
      const lineEnd = tab.content.indexOf('\n', pos);
      const end = lineEnd === -1 ? tab.content.length : lineEnd;
      const currentLine = tab.content.slice(lineStart, end);

      let next: string;
      let caret: number;

      if (currentLine.trim() === '') {
        // Пустая строка — заменяем её целиком
        next = tab.content.slice(0, lineStart) + template + tab.content.slice(end);
        caret = lineStart + template.length;
      } else {
        // Непустая строка — добавляем блок снизу
        next =
          tab.content.slice(0, end) + '\n' + template + tab.content.slice(end);
        caret = end + 1 + template.length;
      }

      onChange(next);
      window.setTimeout(() => {
        textarea.focus();
        textarea.setSelectionRange(caret, caret);
      }, 10);
    },
    [tab.content, onChange]
  );

  /* ---------------- Обработка слэш-меню ---------------- */

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (showSlash) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setSlashIndex((i) => (i + 1) % SLASH_COMMANDS.length);
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        setSlashIndex((i) => (i - 1 + SLASH_COMMANDS.length) % SLASH_COMMANDS.length);
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        runSlashCommand(SLASH_COMMANDS[slashIndex]);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        closeSlash();
        return;
      }
    }

    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
      e.preventDefault();
      onSave();
      return;
    }

    // Ctrl+K — инлайн-копилот по выделенному фрагменту
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
      e.preventDefault();
      onRequestInlineEdit?.();
    }
  };

  const handleChange = (value: string) => {
    onChange(value);

    // Отслеживаем ввод «/» в начале строки для открытия слэш-меню
    const textarea = textareaRef.current;
    if (!textarea) return;
    const pos = textarea.selectionStart;
    const lineStart = value.lastIndexOf('\n', pos - 1) + 1;
    const beforeCaret = value.slice(lineStart, pos);

    const slashMatch = beforeCaret.match(/^\s*\/(\S*)$/);
    if (slashMatch) {
      setSlashFilter(slashMatch[1]);
      setSlashIndex(0);
      // Позиционируем меню у курсора
      const coords = caretCoords(textarea, pos);
      setSlashAnchor(coords);
      setShowSlash(true);
    } else if (showSlash) {
      setShowSlash(false);
    }
  };

  const closeSlash = () => {
    setShowSlash(false);
    setSlashFilter('');
  };

  const runSlashCommand = (cmd: (typeof SLASH_COMMANDS)[number]) => {
    const textarea = textareaRef.current;
    closeSlash();

    if (cmd.ai) {
      handleAIAction('Помоги с этим фрагментом документа:');
      return;
    }

    if (!textarea) return;

    // Удаляем набранное «/команда» перед вставкой блока
    const pos = textarea.selectionStart;
    const lineStart = tab.content.lastIndexOf('\n', pos - 1) + 1;
    const stripped = tab.content.slice(0, lineStart) + tab.content.slice(pos);
    onChange(stripped);

    window.setTimeout(() => {
      const restore = textareaRef.current;
      if (!restore) return;
      restore.focus();
      restore.setSelectionRange(lineStart, lineStart);
      insertBlock(cmd.insert);
    }, 10);
  };

  const filteredCommands = useMemo(() => {
    if (!slashFilter) return SLASH_COMMANDS;
    const q = slashFilter.toLowerCase();
    return SLASH_COMMANDS.filter((c) => c.label.toLowerCase().includes(q) || c.id.includes(q));
  }, [slashFilter]);

  /* ---------------- AI-действия ---------------- */

  const handleAIAction = (prompt: string) => {
    if (!onAskAI) return;
    const textarea = textareaRef.current;
    let context = tab.content;
    if (textarea) {
      const selected = tab.content.slice(textarea.selectionStart, textarea.selectionEnd);
      if (selected.trim()) context = selected;
    }
    onAskAI(prompt, context);
  };

  /* ---------------- Рендеринг диаграмм Mermaid ---------------- */
  useEffect(() => {
    if (viewMode === 'visual' || viewMode === 'split') {
      void renderMermaidBlocks(previewRef.current);
    }
  }, [tab.content, viewMode]);

  /* ---------------- Вставка текста из чата AI ---------------- */
  useEffect(() => {
    const handleInsert = (event: Event) => {
      const detail = (event as CustomEvent).detail as { tabId?: string; text?: string } | undefined;
      if (!detail?.text || detail.tabId !== tab.id) return;
      const insertStr = detail.text;
      const textarea = textareaRef.current;
      if (textarea) {
        const start = textarea.selectionStart;
        const end = textarea.selectionEnd;
        const before = tab.content.slice(0, start);
        const after = tab.content.slice(end);
        const next = `${before}${insertStr}${after}`;
        onChange(next);
        window.setTimeout(() => {
          textarea.focus();
          textarea.setSelectionRange(start + insertStr.length, start + insertStr.length);
        }, 10);
      } else {
        const next = tab.content ? `${tab.content}\n\n${insertStr}` : insertStr;
        onChange(next);
      }
    };

    window.addEventListener('workspace:insert-text', handleInsert);
    return () => window.removeEventListener('workspace:insert-text', handleInsert);
  }, [tab.id, tab.content, onChange]);

  /* ---------------- Синхронизация прокрутки ---------------- */

  const handleEditorScroll = () => {
    if (!scrollSync || viewMode !== 'split') return;
    const textarea = textareaRef.current;
    const preview = previewRef.current;
    if (!textarea || !preview || syncingRef.current) return;

    syncingRef.current = true;
    const ratio = textarea.scrollTop / Math.max(1, textarea.scrollHeight - textarea.clientHeight);
    preview.scrollTop = ratio * (preview.scrollHeight - preview.clientHeight);
    window.setTimeout(() => {
      syncingRef.current = false;
    }, 30);
  };

  /* ---------------- Экспорт ---------------- */

  const exportAs = (format: 'md' | 'html') => {
    const baseName = tab.name.replace(/\.[^.]+$/, '');
    let content: string;
    let mime: string;

    if (format === 'md') {
      content = tab.content;
      mime = 'text/markdown;charset=utf-8';
    } else {
      const body = renderRichMarkdown(tab.content);
      content = `<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8"><title>${baseName}</title>
<style>body{font-family:'Segoe UI',system-ui,sans-serif;font-size:15px;line-height:1.7;max-width:860px;margin:40px auto;padding:0 24px;color:#1c2430}
pre{background:#f5f7fa;border:1px solid #dde3ec;border-radius:8px;padding:12px;overflow-x:auto}
code{font-family:Consolas,monospace;font-size:.9em;background:#f0f3f7;border-radius:4px;padding:1px 4px}
pre code{background:none;padding:0}
table{border-collapse:collapse;width:100%;margin:16px 0}td,th{border:1px solid #d5dce6;padding:7px 10px;text-align:left}
th{background:#f2f5f9}blockquote{border-left:3px solid #2563eb;margin:12px 0;padding-left:14px;color:#5a6a7d;font-style:italic}
img{max-width:100%}</style></head><body>${body}</body></html>`;
      mime = 'text/html;charset=utf-8';
    }

    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${baseName}.${format}`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const stats = useMemo(() => {
    const words = tab.content.trim() ? tab.content.trim().split(/\s+/).length : 0;
    return {
      words,
      chars: tab.content.length,
      lines: tab.content.split('\n').length,
      readingMin: Math.max(1, Math.round(words / 180)),
    };
  }, [tab.content]);

  return (
    <div className="flex h-full w-full flex-col bg-ws-editor text-ws">
      {/* Панель инструментов */}
      <div className="flex shrink-0 flex-wrap items-center justify-between gap-2 border-b border-ws bg-ws-surface px-3 py-1.5 text-xs select-none">
        <div className="flex flex-wrap items-center gap-0.5">
          <IconButton title="Заголовок 1" onClick={() => insertBlock('# ')}>
            <Heading1 size={14} />
          </IconButton>
          <IconButton title="Заголовок 2" onClick={() => insertBlock('## ')}>
            <Heading2 size={14} />
          </IconButton>
          <IconButton title="Заголовок 3" onClick={() => insertBlock('### ')}>
            <Heading3 size={14} />
          </IconButton>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <IconButton title="Жирный" onClick={() => insertText('**', '**', 'жирный текст')}>
            <Bold size={14} />
          </IconButton>
          <IconButton title="Курсив" onClick={() => insertText('*', '*', 'курсив')}>
            <Italic size={14} />
          </IconButton>
          <IconButton title="Зачёркнутый" onClick={() => insertText('~~', '~~', 'зачёркнутый')}>
            <Strikethrough size={14} />
          </IconButton>
          <IconButton title="Встроенный код" onClick={() => insertText('`', '`', 'код')}>
            <Code size={14} />
          </IconButton>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <IconButton title="Маркированный список" onClick={() => insertBlock('- ')}>
            <List size={14} />
          </IconButton>
          <IconButton title="Нумерованный список" onClick={() => insertBlock('1. ')}>
            <ListOrdered size={14} />
          </IconButton>
          <IconButton title="Чек-лист" onClick={() => insertBlock('- [ ] ')}>
            <CheckSquare size={14} />
          </IconButton>
          <IconButton title="Цитата" onClick={() => insertBlock('> ')}>
            <Quote size={14} />
          </IconButton>
          <IconButton
            title="Таблица"
            onClick={() =>
              insertBlock('| Колонка 1 | Колонка 2 |\n| --- | --- |\n| Значение | Значение |\n')
            }
          >
            <TableIcon size={14} />
          </IconButton>
          <IconButton title="Разделитель" onClick={() => insertBlock('\n---\n')}>
            <Minus size={14} />
          </IconButton>
        </div>

        <div className="flex items-center gap-1.5">
          <button
            onClick={() => onRequestInlineEdit?.()}
            title="AI-правка выделенного фрагмента (Ctrl+K)"
            className="flex items-center gap-1 rounded border border-ws-accent bg-ws-accent-soft px-2 py-1 text-[11px] text-ws-accent transition-colors"
          >
            <Sparkles size={11} />
            AI-правка
          </button>
          <button
            onClick={() =>
              handleAIAction(
                'Исправь грамматические и стилистические ошибки, улучши формулировки, сохранив исходный смысл. Верни готовый текст.'
              )
            }
            className="rounded border border-ws px-2 py-1 text-[11px] text-ws-muted transition-colors hover:border-ws-accent hover:text-ws-accent"
          >
            AI-редактор
          </button>
          <button
            onClick={() =>
              handleAIAction(
                'Сделай краткую структурированную выжимку ключевых мыслей документа в виде списка тезисов.'
              )
            }
            className="rounded border border-ws px-2 py-1 text-[11px] text-ws-muted transition-colors hover:border-ws-accent hover:text-ws-accent"
          >
            Выжимка
          </button>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          {/* Экспорт */}
          <div className="flex items-center gap-0.5">
            <button
              onClick={() => exportAs('md')}
              title="Экспортировать в Markdown"
              className="flex items-center gap-1 rounded border border-ws px-1.5 py-1 text-[10px] text-ws-faint transition-colors hover:border-ws-accent hover:text-ws-accent"
            >
              <Download size={10} />
              MD
            </button>
            <button
              onClick={() => exportAs('html')}
              title="Экспортировать в HTML"
              className="flex items-center gap-1 rounded border border-ws px-1.5 py-1 text-[10px] text-ws-faint transition-colors hover:border-ws-accent hover:text-ws-accent"
            >
              <Download size={10} />
              HTML
            </button>
          </div>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          {/* Режим просмотра */}
          <div className="flex items-center rounded border border-ws bg-ws-elevated p-0.5">
            <button
              onClick={() => changeViewMode('source')}
              title="Только исходный код Markdown"
              className={cn(
                'rounded p-1 transition-colors',
                viewMode === 'source' ? 'bg-ws-accent-soft text-ws-accent' : 'text-ws-faint hover:text-ws'
              )}
            >
              <FileEdit size={12} />
            </button>
            <button
              onClick={() => changeViewMode('split')}
              title="Разделённый режим (код + предпросмотр)"
              className={cn(
                'rounded p-1 transition-colors',
                viewMode === 'split' ? 'bg-ws-accent-soft text-ws-accent' : 'text-ws-faint hover:text-ws'
              )}
            >
              <Columns size={12} />
            </button>
            <button
              onClick={() => changeViewMode('visual')}
              title="Только предпросмотр"
              className={cn(
                'rounded p-1 transition-colors',
                viewMode === 'visual' ? 'bg-ws-accent-soft text-ws-accent' : 'text-ws-faint hover:text-ws'
              )}
            >
              <Eye size={12} />
            </button>
          </div>

          <button
            onClick={onSave}
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

      {/* Область редактирования */}
      <div className="relative flex min-h-0 flex-1">
        {(viewMode === 'source' || viewMode === 'split') && (
          <div
            className={cn(
              'flex h-full flex-col',
              viewMode === 'split' ? 'w-1/2 border-r border-ws' : 'w-full'
            )}
          >
            <textarea
              ref={textareaRef}
              value={tab.content}
              onChange={(e) => handleChange(e.target.value)}
              onKeyDown={handleKeyDown}
              onScroll={handleEditorScroll}
              onClick={closeSlash}
              placeholder={
                'Начните писать…\n\nПодсказки:\n# Заголовок    **жирный**    *курсив*\n- список    - [ ] задача    > цитата\n/ — меню вставки блоков'
              }
              spellCheck={false}
              className="h-full w-full resize-none bg-ws-editor p-5 font-mono text-[13px] leading-relaxed text-ws placeholder:text-ws-faint focus:outline-none select-text"
            />
          </div>
        )}

        {(viewMode === 'visual' || viewMode === 'split') && (
          <div
            ref={previewRef}
            className={cn(
              'h-full overflow-y-auto bg-ws-editor-alt',
              viewMode === 'split' ? 'w-1/2' : 'w-full'
            )}
          >
            <div
              className="ws-prose mx-auto max-w-3xl p-7 select-text"
              dangerouslySetInnerHTML={{ __html: renderRichMarkdown(tab.content) }}
            />
          </div>
        )}

        {/* Слэш-меню */}
        {showSlash && slashAnchor && (
          <div
            style={{ top: slashAnchor.top, left: slashAnchor.left }}
            className="animate-scale-in absolute z-30 max-h-72 w-64 overflow-y-auto rounded-lg border border-ws bg-ws-surface shadow-ws"
          >
            <div className="flex items-center gap-1.5 border-b border-ws px-2.5 py-1.5 text-[10px] font-semibold text-ws-muted">
              <Command size={10} />
              Вставить блок
              {slashFilter && <span className="text-ws-accent">/{slashFilter}</span>}
            </div>

            {filteredCommands.length === 0 && (
              <div className="px-3 py-3 text-[11px] text-ws-faint">Команды не найдены</div>
            )}

            {filteredCommands.map((cmd, idx) => (
              <button
                key={cmd.id}
                onClick={() => runSlashCommand(cmd)}
                onMouseEnter={() => setSlashIndex(idx)}
                className={cn(
                  'flex w-full items-center gap-2.5 px-2.5 py-1.5 text-left text-[11px] transition-colors',
                  idx === slashIndex ? 'bg-ws-accent text-white' : 'text-ws-muted hover:bg-ws-hover'
                )}
              >
                <span className={idx === slashIndex ? 'text-white' : 'text-ws-faint'}>{cmd.icon}</span>
                <span className="flex-1">{cmd.label}</span>
                {idx === slashIndex && <ChevronRight size={11} />}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Строка состояния */}
      <div className="flex shrink-0 items-center justify-between border-t border-ws bg-ws-surface px-3 py-1 text-[10px] text-ws-faint select-none">
        <div className="flex items-center gap-3">
          <span>Слов: {stats.words.toLocaleString('ru-RU')}</span>
          <span>Символов: {stats.chars.toLocaleString('ru-RU')}</span>
          <span>Строк: {stats.lines}</span>
          <span>Чтение: ~{stats.readingMin} мин</span>
        </div>
        <div className="flex items-center gap-3">
          {viewMode === 'split' && (
            <button
              onClick={() => setScrollSync((v) => !v)}
              title="Синхронизация прокрутки областей"
              className={cn(
                'rounded px-1.5 py-0.5 transition-colors',
                scrollSync ? 'text-ws-accent' : 'hover:text-ws'
              )}
            >
              Синхр. прокрутки
            </button>
          )}
          <span>Markdown</span>
          {tab.isDirty && <span className="font-medium text-ws-warning">● Не сохранено</span>}
        </div>
      </div>
    </div>
  );
};

/* ============================================================
   Позиция курсора в textarea (для позиционирования слэш-меню)
   ============================================================ */

function caretCoords(textarea: HTMLTextAreaElement, position: number): { top: number; left: number } {
  const textareaRect = textarea.getBoundingClientRect();
  const style = window.getComputedStyle(textarea);

  // Скрытый div-двойник для измерения координат
  const mirror = document.createElement('div');
  const props = [
    'boxSizing',
    'width',
    'height',
    'overflowX',
    'overflowY',
    'borderTopWidth',
    'borderRightWidth',
    'borderBottomWidth',
    'borderLeftWidth',
    'paddingTop',
    'paddingRight',
    'paddingBottom',
    'paddingLeft',
    'fontStyle',
    'fontVariant',
    'fontWeight',
    'fontStretch',
    'fontSize',
    'fontSizeAdjust',
    'lineHeight',
    'fontFamily',
    'textAlign',
    'textTransform',
    'textIndent',
    'textDecoration',
    'letterSpacing',
    'wordSpacing',
    'tabSize',
    'whiteSpace',
    'wordBreak',
    'wordWrap',
  ] as const;

  mirror.style.position = 'absolute';
  mirror.style.visibility = 'hidden';
  mirror.style.whiteSpace = 'pre-wrap';
  mirror.style.wordWrap = 'break-word';
  props.forEach((p) => {
    mirror.style.setProperty(
      p.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`),
      style.getPropertyValue(p.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`))
    );
  });

  mirror.textContent = textarea.value.substring(0, position);
  const span = document.createElement('span');
  span.textContent = textarea.value.substring(position) || '.';
  mirror.appendChild(span);

  document.body.appendChild(mirror);
  const spanRect = span.getBoundingClientRect();
  const mirrorRect = mirror.getBoundingClientRect();
  document.body.removeChild(mirror);

  return {
    top: textareaRect.top + (spanRect.top - mirrorRect.top) - textarea.scrollTop + 22,
    left: textareaRect.left + (spanRect.left - mirrorRect.left) - textarea.scrollLeft + 12,
  };
}
