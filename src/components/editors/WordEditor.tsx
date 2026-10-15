import React, { useCallback, useEffect, useRef, useState } from 'react';
import { OpenTab } from '../../types';
import { docxBase64ToHtml, htmlToDocxBase64 } from '../../lib/docx';
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

  /* ---------------- Разбор DOCX через mammoth ---------------- */

  useEffect(() => {
    let cancelled = false;

    const parseDocx = async () => {
      try {
        setLoading(true);
        setError('');

        if (!tab.content) {
          if (!cancelled) {
            setPendingHtml('<p></p>');
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
        setLoading(false);
      } catch (err) {
        if (cancelled) return;
        // Резервное восстановление: файл мог быть создан как Markdown/текст вместо бинарного ZIP
        const recoveredText = tryDecodeText(tab.content);
        if (recoveredText && recoveredText.trim()) {
          const html = markdownToDocumentHtml(recoveredText);
          setPendingHtml(html || '<p></p>');
          setLoading(false);
          void htmlToDocxBase64(html)
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
    setWordCount(countWords(root.innerText));
    queueMicrotask(() => {
      isSyncing.current = false;
    });
  }, [pendingHtml, loading]);


  /* ---------------- Отслеживание правок с debounce ---------------- */

  const debounceTimer = useRef<NodeJS.Timeout | null>(null);

  useEffect(() => {
    return () => {
      if (debounceTimer.current) clearTimeout(debounceTimer.current);
    };
  }, []);

  const handleInput = useCallback(() => {
    if (!editorRef.current || isSyncing.current) return;
    const html = editorRef.current.innerHTML;
    setWordCount(countWords(editorRef.current.innerText));

    // Debounce упаковки пакета .docx (устраняет фризы при быстром наборе текста)
    if (debounceTimer.current) clearTimeout(debounceTimer.current);
    debounceTimer.current = setTimeout(() => {
      const seq = ++serializeSeq.current;
      void htmlToDocxBase64(html)
        .then((base64) => {
          if (serializeSeq.current === seq) onChange(base64);
        })
        .catch(() => {
          /* Ошибку сериализации не показываем на каждый символ: документ остаётся открытым. */
        });
    }, 400);
  }, [onChange]);

  const flushSave = useCallback(async () => {
    if (debounceTimer.current) {
      clearTimeout(debounceTimer.current);
      debounceTimer.current = null;
    }
    if (editorRef.current) {
      try {
        const html = editorRef.current.innerHTML;
        const base64 = await htmlToDocxBase64(html);
        onChange(base64);
      } catch {
        // ignore
      }
    }
    onSave();
  }, [onChange, onSave]);

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
      insertHtmlAt(editorRef.current, null, html);
      handleInput();
    };

    const handleReplace = (event: Event) => {
      const detail = (event as CustomEvent).detail as { tabId?: string; text?: string } | undefined;
      if (detail?.tabId !== tab.id) return;
      const root = editorRef.current;
      if (!root) return;
      const html = markdownToDocumentHtml(detail.text || '');
      isSyncing.current = true;
      root.innerHTML = html || '<p></p>';
      setWordCount(countWords(root.innerText));
      queueMicrotask(() => {
        isSyncing.current = false;
      });
      handleInput();
    };

    window.addEventListener('workspace:insert-text', handleInsert);
    window.addEventListener('workspace:replace-document', handleReplace);
    return () => {
      window.removeEventListener('workspace:insert-text', handleInsert);
      window.removeEventListener('workspace:replace-document', handleReplace);
    };
  }, [handleInput, tab.id]);

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

      {/* Полотно документа */}
      <div className="min-h-0 flex-1 overflow-y-auto bg-ws-editor-alt p-6">
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
          style={{ fontFamily: 'Calibri, "Segoe UI", Georgia, serif', fontSize: '15px', lineHeight: 1.7 }}
        />
      </div>

      {/* Строка состояния */}
      <div className="flex shrink-0 items-center justify-between border-t border-ws bg-ws-surface px-3 py-1 text-[10px] text-ws-faint">
        <span>DOCX · {tab.name}</span>
        <div className="flex items-center gap-3">
          {format.fontName && <span>{format.fontName}</span>}
          {format.fontSizePt && <span>{format.fontSizePt} пт</span>}
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
 */
function insertHtmlAt(root: HTMLDivElement | null, range: Range | null, html: string) {
  if (!root) return;
  root.focus();

  const selection = window.getSelection();
  if (range && root.contains(range.startContainer) && selection) {
    selection.removeAllRanges();
    selection.addRange(range);
  }

  const inserted = document.execCommand('insertHTML', false, html);
  if (inserted) return;

  root.insertAdjacentHTML('beforeend', html);
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