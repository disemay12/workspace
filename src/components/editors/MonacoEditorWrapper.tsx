import React, { useCallback, useEffect, useRef, useState } from 'react';
import Editor, { OnMount } from '@monaco-editor/react';
import { getMonacoLanguage, SUPPORTED_ENCODINGS, fileSystemService } from '../../services/fileSystemService';
import { OpenTab } from '../../types';
import {
  WrapText,
  FileCode,
  MapPin,
  Sparkles,
  Save,
  Search,
  Sparkle,
  Info,
  Check,
} from 'lucide-react';
import { cn } from '../../lib/utils';
import { IconButton } from '../ui';

interface Props {
  tab: OpenTab;
  onChange: (newContent: string) => void;
  onSave: () => void;
  onAskAI?: (prompt: string, contextText: string) => void;
  /** Вызов инлайн-копилота (Ctrl+K) с текущим выделением. */
  onRequestInlineEdit?: () => void;
  /** Передаёт наружу доступ к контексту редактора (выделение/курсор). */
  onRegisterContext?: (getter: (() => { text: string; from: number; to: number }) | null) => void;
  /** Передаёт наружу функцию применения правки к живому редактору. */
  onRegisterApplier?: (
    applier: ((edit: { from: number; to: number; text: string }) => void) | null
  ) => void;
}

export const MonacoEditorWrapper: React.FC<Props> = ({
  tab,
  onChange,
  onSave,
  onAskAI,
  onRequestInlineEdit,
  onRegisterContext,
  onRegisterApplier,
}) => {
  const editorRef = useRef<any>(null);
  const monacoRef = useRef<any>(null);
  const [wordWrap, setWordWrap] = useState<'on' | 'off'>('on');
  const [minimap, setMinimap] = useState<boolean>(false);
  const [cursorPos, setCursorPos] = useState({ line: 1, col: 1 });
  const [selectionInfo, setSelectionInfo] = useState({ chars: 0, lines: 0 });
  const [encoding, setEncoding] = useState(tab.encoding);
  const [showEncodingMenu, setShowEncodingMenu] = useState(false);
  const [isReloading, setIsReloading] = useState(false);

  const language = getMonacoLanguage(tab.extension);
  const isLogFile =
    tab.extension === 'log' || /\.log(\.\d+)?$/i.test(tab.name) || tab.name.toLowerCase().includes('log');

  /* ---------------- Регистрация контекста для App ---------------- */

  useEffect(() => {
    if (!onRegisterContext) return;

    onRegisterContext(() => {
      const editor = editorRef.current;
      if (!editor) return { text: '', from: 0, to: 0 };
      const model = editor.getModel();
      if (!model) return { text: '', from: 0, to: 0 };

      const selection = editor.getSelection();
      if (selection && !selection.isEmpty()) {
        return {
          text: model.getValueInRange(selection),
          from: model.getOffsetAt(selection.getStartPosition()),
          to: model.getOffsetAt(selection.getEndPosition()),
        };
      }
      return { text: '', from: 0, to: 0 };
    });

    return () => onRegisterContext(null);
  }, [onRegisterContext]);

  /* ---------------- Регистрация applier правок ---------------- */

  useEffect(() => {
    if (!onRegisterApplier) return;

    onRegisterApplier((edit) => {
      const editor = editorRef.current;
      const monaco = monacoRef.current;
      if (!editor || !monaco) return;
      const model = editor.getModel();
      if (!model) return;

      // from === -1 означает «вставить в позицию курсора»
      if (edit.from === -1) {
        let position = editor.getPosition();
        if (!position) {
          const lineCount = model.getLineCount();
          position = { lineNumber: lineCount, column: model.getLineMaxColumn(lineCount) } as any;
        }
        editor.executeEdits('ai-insert', [
          { range: new monaco.Range(position.lineNumber, position.column, position.lineNumber, position.column), text: edit.text },
        ]);
      } else {
        const fullRange = model.getFullModelRange();
        const start = edit.from === 0 ? fullRange.getStartPosition() : model.getPositionAt(edit.from);
        const end = edit.to >= model.getValueLength() ? fullRange.getEndPosition() : model.getPositionAt(edit.to);
        editor.executeEdits('ai-replace', [
          { range: new monaco.Range(start.lineNumber, start.column, end.lineNumber, end.column), text: edit.text },
        ]);
      }
      editor.pushUndoStop();
      editor.focus();
    });

    return () => onRegisterApplier(null);
  }, [onRegisterApplier]);

  /* ---------------- Монтирование Monaco ---------------- */

  const handleEditorDidMount: OnMount = (editor, monaco) => {
    editorRef.current = editor;
    monacoRef.current = monaco;

    // Пока кнопка мыши зажата, обновление React пересоздаёт редактор
    // и обрывает протяжку выделения. Поэтому строку состояния обновляем
    // только когда кнопку уже отпустили.
    const publishCursor = () => {
      const position = editor.getPosition();
      const model = editor.getModel();
      const selection = editor.getSelection();
      if (position) setCursorPos({ line: position.lineNumber, col: position.column });
      if (!model || !selection || selection.isEmpty()) {
        setSelectionInfo({ chars: 0, lines: 0 });
        return;
      }
      const text = model.getValueInRange(selection);
      setSelectionInfo({ chars: text.length, lines: text ? text.split('\n').length : 0 });
    };
    let dragging = false;
    editor.onMouseDown(() => { dragging = true; });
    editor.onMouseUp(() => {
      dragging = false;
      publishCursor();
    });
    editor.onDidChangeCursorSelection(() => {
      if (!dragging) publishCursor();
    });

    // Ctrl+S — сохранить
    editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyS, () => onSave());

    // Ctrl+F — поиск
    editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyF, () => {
      editor.getAction('actions.find')?.run();
    });

    // Ctrl+K — инлайн-копилот
    editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyK, () => {
      onRequestInlineEdit?.();
    });

    editor.focus();
  };

  /* ---------------- Смена кодировки с перечитыванием файла ---------------- */

  const reloadWithEncoding = useCallback(
    async (nextEncoding: string) => {
      setShowEncodingMenu(false);
      if (nextEncoding === encoding) return;
      setIsReloading(true);
      try {
        const res = await fileSystemService.readFile(tab.path, nextEncoding);
        onChange(res.content);
        setEncoding(nextEncoding);
      } catch {
        // Не удалось перечитать — оставляем текущее содержимое
      } finally {
        setIsReloading(false);
      }
    },
    [encoding, tab.path, onChange]
  );

  /* ---------------- Открытие поиска ---------------- */

  const openFind = () => editorRef.current?.getAction('actions.find')?.run();

  const handleAnalyzeLogs = () => {
    if (!onAskAI) return;
    const model = editorRef.current?.getModel();
    const selection = model?.getValueInRange(editorRef.current.getSelection()) ?? '';
    const textToAnalyze =
      selection.trim() || tab.content.split('\n').slice(0, 400).join('\n').slice(0, 8000);
    onAskAI(
      'Проанализируй этот фрагмент лога: найди ошибки (ERROR, FATAL, Exception, timeout), объясни причину сбоя и предложи конкретные шаги по исправлению.',
      textToAnalyze
    );
  };

  // defaultValue задаётся один раз, поэтому правки извне (AI, вставка)
  // нужно перенести в уже живую модель, не трогая её при обычном наборе.
  useEffect(() => {
    const model = editorRef.current?.getModel();
    if (!model || model.getValue() === tab.content) return;
    model.setValue(tab.content);
  }, [tab.content]);

  const lineCount = tab.content.split('\n').length;
  const charCount = tab.content.length;

  return (
    <div className="flex h-full w-full flex-col bg-ws-editor text-ws">
      {/* Панель инструментов */}
      <div className="flex shrink-0 items-center justify-between border-b border-ws bg-ws-surface px-3 py-1.5 text-xs select-none">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1.5 font-medium text-ws-muted">
            <FileCode size={14} className="text-ws-accent" />
            <span className="font-mono text-ws">{tab.name}</span>
            <span className="rounded bg-ws-panel-alt px-1.5 py-0.5 font-mono text-[10px] text-ws-faint uppercase">
              {language}
            </span>
          </div>

          <div className="h-3.5 w-px bg-ws border-l border-ws" />

          <span className="text-ws-faint">
            Стр {cursorPos.line}, Кол {cursorPos.col}
          </span>

          {selectionInfo.chars > 0 && (
            <span className="text-ws-accent">
              Выделено: {selectionInfo.chars} симв. / {selectionInfo.lines} стр.
            </span>
          )}
        </div>

        <div className="flex items-center gap-1.5">
          {isLogFile && (
            <button
              onClick={handleAnalyzeLogs}
              title="Найти ошибки в логе с помощью AI"
              className="flex items-center gap-1 rounded border border-amber-500/40 bg-amber-500/10 px-2 py-1 text-[11px] text-ws-warning transition-colors hover:bg-amber-500/20"
            >
              <Sparkles size={11} />
              Анализ лога
            </button>
          )}

          <button
            onClick={() => onRequestInlineEdit?.()}
            title="AI-правка выделенного текста (Ctrl+K)"
            className="flex items-center gap-1 rounded border border-ws-accent bg-ws-accent-soft px-2 py-1 text-[11px] text-ws-accent transition-colors"
          >
            <Sparkle size={11} />
            AI-правка
          </button>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <IconButton
            title={`Перенос строк: ${wordWrap === 'on' ? 'включён' : 'выключен'}`}
            active={wordWrap === 'on'}
            onClick={() => setWordWrap((p) => (p === 'on' ? 'off' : 'on'))}
          >
            <WrapText size={14} />
          </IconButton>

          <IconButton
            title={`Миникарта: ${minimap ? 'включена' : 'выключена'}`}
            active={minimap}
            onClick={() => setMinimap((p) => !p)}
          >
            <MapPin size={14} />
          </IconButton>

          <IconButton title="Поиск и замена (Ctrl+F)" onClick={openFind}>
            <Search size={14} />
          </IconButton>

          {/* Переключатель кодировки */}
          <div className="relative">
            <button
              onClick={() => setShowEncodingMenu((v) => !v)}
              title="Кодировка файла"
              className={cn(
                'flex items-center gap-1 rounded border px-1.5 py-1 font-mono text-[10px] transition-colors',
                isReloading
                  ? 'border-ws text-ws-faint'
                  : 'border-ws text-ws-muted hover:border-ws-accent hover:text-ws-accent'
              )}
            >
              {encoding.toUpperCase()}
            </button>

            {showEncodingMenu && (
              <>
                <div className="fixed inset-0 z-40" onMouseDown={() => setShowEncodingMenu(false)} />
                <div className="animate-scale-in absolute right-0 top-full z-50 mt-1 w-52 overflow-hidden rounded-lg border border-ws bg-ws-surface shadow-ws">
                  <div className="border-b border-ws px-2.5 py-1.5 text-[10px] font-semibold text-ws-muted">
                    Открыть заново в кодировке
                  </div>
                  {SUPPORTED_ENCODINGS.map((enc) => (
                    <button
                      key={enc.id}
                      onClick={() => void reloadWithEncoding(enc.id)}
                      className={cn(
                        'flex w-full items-center justify-between px-2.5 py-1.5 text-left text-[11px] transition-colors hover:bg-ws-hover',
                        enc.id === encoding ? 'text-ws-accent' : 'text-ws-muted'
                      )}
                    >
                      <span>{enc.label}</span>
                      {enc.id === encoding && <Check size={11} />}
                    </button>
                  ))}
                </div>
              </>
            )}
          </div>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <button
            onClick={onSave}
            disabled={!tab.isDirty}
            title="Сохранить (Ctrl+S)"
            className={cn(
              'flex items-center gap-1 rounded px-2.5 py-1 text-xs font-medium transition-colors',
              tab.isDirty
                ? 'bg-ws-accent text-white hover:opacity-90'
                : 'bg-ws-panel-alt text-ws-faint'
            )}
          >
            <Save size={12} />
            Сохранить
          </button>
        </div>
      </div>

      {/* Monaco */}
      <div className="min-h-0 flex-1 overflow-hidden">
        <Editor
          key={tab.id}
          height="100%"
          language={language}
          defaultValue={tab.content}
          theme={document.documentElement.classList.contains('dark') ? 'vs-dark' : 'vs'}
          onMount={handleEditorDidMount}
          onChange={(val) => onChange(val || '')}
          options={{
            dragAndDrop: false,
            columnSelection: false,
            selectionHighlight: false,
            fontFamily: "'JetBrains Mono', 'Cascadia Mono', Consolas, monospace",
            fontSize: 13,
            lineHeight: 20,
            wordWrap,
            minimap: { enabled: minimap },
            scrollBeyondLastLine: false,
            automaticLayout: true,
            tabSize: 2,
            renderWhitespace: 'selection',
            cursorBlinking: 'smooth',
            smoothScrolling: true,
            bracketPairColorization: { enabled: true },
            guides: { bracketPairs: true, indentation: true },
            padding: { top: 8, bottom: 8 },
            // Большие файлы: отключаем дорогие эффекты автоматически
            ...(charCount > 1_500_000
              ? {
                  minimap: { enabled: false },
                  renderWhitespace: 'none' as const,
                  bracketPairColorization: { enabled: false },
                  wordWrap: 'off' as const,
                  folding: false,
                  occurrencesHighlight: 'off' as const,
                }
              : {}),
          }}
        />
      </div>

      {/* Строка состояния */}
      <div className="flex shrink-0 items-center justify-between border-t border-ws bg-ws-surface px-3 py-1 text-[10px] text-ws-faint select-none">
        <div className="flex items-center gap-3">
          <span>Символов: {charCount.toLocaleString('ru-RU')}</span>
          <span>Строк: {lineCount.toLocaleString('ru-RU')}</span>
          <span>Размер: {(new Blob([tab.content]).size / 1024).toFixed(1)} КБ</span>
        </div>
        <div className="flex items-center gap-3">
          {tab.isDirty && <span className="font-medium text-ws-warning">● Не сохранено</span>}
          <span className="uppercase">{encoding}</span>
          <span>LF</span>
          {charCount > 1_500_000 && (
            <span className="flex items-center gap-1 text-ws-accent" title="Режим оптимизации для больших файлов">
              <Info size={10} />
              Large file mode
            </span>
          )}
        </div>
      </div>
    </div>
  );
};
