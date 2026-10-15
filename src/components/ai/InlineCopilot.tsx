import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  Sparkles,
  Loader2,
  Check,
  X,
  RefreshCw,
  CornerDownLeft,
  AlertTriangle,
  Wand2,
  Languages,
  AlignLeft,
  HelpCircle,
  ArrowRight,
} from 'lucide-react';
import { AISettings } from '../../types';
import { aiService } from '../../services/aiService';
import { cn, stripCodeFence } from '../../lib/utils';
import { Button, Textarea } from '../ui';

export interface InlineEditTarget {
  tabId: string;
  /** Исходный текст (выделение либо весь документ). */
  original: string;
  /** Позиция вставки/замены в документе. */
  from: number;
  to: number;
  /** Имя файла для заголовка. */
  fileName: string;
  /** Экранные координаты для позиционирования окна. */
  anchor: { top: number; left: number };
  /** Выделен ли фрагмент (иначе — весь документ). */
  hasSelection: boolean;
}

interface Props {
  target: InlineEditTarget | null;
  settings: AISettings;
  isConfigured: boolean;
  onClose: () => void;
  onApply: (target: InlineEditTarget, newText: string) => void;
  onOpenSettings: () => void;
}

/** Быстрые команды инлайн-ассистента (Этап 5.3 плана). */
const QUICK_ACTIONS = [
  { id: 'fix', label: 'Исправить ошибки и стиль', icon: <Wand2 size={12} />, prompt: 'Исправь орфографические, пунктуационные и стилистические ошибки. Сохрани исходный смысл и структуру. Верни только исправленный текст.' },
  { id: 'shorten', label: 'Сократить', icon: <AlignLeft size={12} />, prompt: 'Сократи текст, сохранив ключевые смыслы. Сделай его лаконичнее. Верни только результат.' },
  { id: 'translate_en', label: 'Перевести на английский', icon: <Languages size={12} />, prompt: 'Переведи текст на английский язык, сохранив стиль и терминологию. Верни только перевод.' },
  { id: 'translate_ru', label: 'Перевести на русский', icon: <Languages size={12} />, prompt: 'Переведи текст на русский язык, сохранив стиль и терминологию. Верни только перевод.' },
  { id: 'explain', label: 'Объяснить', icon: <HelpCircle size={12} />, prompt: 'Объясни понятным языком, что делает этот фрагмент. Если это код или лог — разбери логику и укажи на возможные проблемы.' },
  { id: 'continue', label: 'Продолжить текст', icon: <ArrowRight size={12} />, prompt: 'Продолжи текст логически и стилистически, дописав 1–2 абзаца. Верни только продолжение.' },
];

export const InlineCopilot: React.FC<Props> = ({
  target,
  settings,
  isConfigured,
  onClose,
  onApply,
  onOpenSettings,
}) => {
  const [instruction, setInstruction] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [proposal, setProposal] = useState('');
  const [error, setError] = useState('');
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const abortRef = useRef<AbortController | null>(null);

  // Сброс состояния при открытии для новой цели
  useEffect(() => {
    if (!target) return;
    setInstruction('');
    setProposal('');
    setError('');
    setIsLoading(false);
    window.setTimeout(() => inputRef.current?.focus(), 60);
  }, [target?.tabId, target?.from, target?.to]);

  // Enter — запустить, Esc — закрыть
  useEffect(() => {
    if (!target) return;
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        abortRef.current?.abort();
        onClose();
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [target, onClose]);

  const run = async (promptText: string) => {
    if (!target || !promptText.trim()) return;
    if (!isConfigured) {
      onOpenSettings();
      return;
    }

    setInstruction(promptText);
    setIsLoading(true);
    setProposal('');
    setError('');

    const controller = new AbortController();
    abortRef.current = controller;

    let buffer = '';
    try {
      await aiService.streamChat(
        [
          {
            id: 'inline',
            role: 'user',
            content:
              `${promptText}\n\n` +
              'ВАЖНО: верни только итоговый текст без пояснений, без Markdown-обёртки ``` и без вступлений.\n\n' +
              'Текст:\n' +
              '---\n' +
              target.original +
              '\n---',
            timestamp: Date.now(),
          },
        ],
        settings,
        {
          onChunk: (partial) => {
            buffer = partial;
            setProposal(partial);
          },
          onFinish: (full) => {
            buffer = full;
            setProposal(stripCodeFence(full));
          },
          onError: (err) => {
            throw err;
          },
        },
        controller.signal
      );
      setProposal(stripCodeFence(buffer));
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        setError((err as Error).message);
      }
    } finally {
      setIsLoading(false);
      abortRef.current = null;
    }
  };

  /* ---------- Построчный diff ---------- */
  const diff = useMemo(() => {
    if (!target || !proposal) return null;
    return computeLineDiff(target.original, proposal);
  }, [target, proposal]);

  if (!target) return null;

  // Позиционирование: не даём окну выйти за пределы вьюпорта
  const width = 560;
  const left = Math.max(12, Math.min(target.anchor.left, window.innerWidth - width - 12));
  const top = Math.max(12, Math.min(target.anchor.top, window.innerHeight - 340));

  const canApply = Boolean(proposal.trim()) && !isLoading;

  return (
    <>
      {/* Перехват кликов вне окна */}
      <div className="fixed inset-0 z-40" onMouseDown={onClose} />

      <div
        style={{ left, top, width }}
        className="animate-scale-in fixed z-50 flex max-h-[80vh] flex-col overflow-hidden rounded-xl border border-ws bg-ws-surface shadow-ws"
      >
        {/* Заголовок */}
        <div className="flex items-center justify-between border-b border-ws px-3 py-2">
          <div className="flex min-w-0 items-center gap-2">
            <Sparkles size={13} className="shrink-0 text-ws-accent" />
            <span className="text-xs font-semibold text-ws">AI-правка</span>
            <span className="truncate text-[11px] text-ws-faint">
              {target.hasSelection
                ? `выделено ${target.original.length} симв.`
                : `весь документ · ${target.fileName}`}
            </span>
          </div>
          <button
            onClick={onClose}
            title="Закрыть (Esc)"
            className="rounded p-1 text-ws-faint transition-colors hover:bg-ws-hover hover:text-ws"
          >
            <X size={13} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto">
          {/* Ввод инструкции */}
          <div className="space-y-2 p-3">
            <Textarea
              ref={inputRef}
              value={instruction}
              onChange={(e) => setInstruction(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault();
                  void run(instruction);
                }
              }}
              rows={2}
              placeholder="Что сделать с текстом? Например: переписать официальным тоном…"
              className="text-xs"
            />

            {/* Быстрые действия */}
            <div className="flex flex-wrap gap-1">
              {QUICK_ACTIONS.map((a) => (
                <button
                  key={a.id}
                  onClick={() => void run(a.prompt)}
                  disabled={isLoading}
                  className="flex items-center gap-1 rounded border border-ws bg-ws-elevated px-1.5 py-1 text-[10px] text-ws-muted transition-colors hover:border-ws-accent hover:text-ws-accent disabled:opacity-50"
                >
                  {a.icon}
                  {a.label}
                </button>
              ))}
            </div>
          </div>

          {/* Результат / diff */}
          {(isLoading || proposal || error) && (
            <div className="border-t border-ws">
              <div className="flex items-center justify-between px-3 py-1.5">
                <span className="text-[10px] font-semibold tracking-wide text-ws-muted uppercase">
                  {error ? 'Ошибка' : isLoading ? 'Генерация…' : 'Предпросмотр изменений'}
                </span>
                {isLoading && (
                  <span className="flex items-center gap-1.5 text-[10px] text-ws-accent">
                    <Loader2 size={10} className="animate-spin" />
                    стриминг
                  </span>
                )}
              </div>

              {error && (
                <div className="mx-3 mb-3 flex items-start gap-2 rounded border border-red-500/40 bg-red-500/10 px-2.5 py-2 text-[11px] text-ws-danger">
                  <AlertTriangle size={12} className="mt-0.5 shrink-0" />
                  <span className="select-text">{error}</span>
                </div>
              )}

              {!error && diff && (
                <div className="max-h-64 overflow-y-auto px-3 pb-3">
                  <div className="overflow-hidden rounded-lg border border-ws font-mono text-[11px] leading-relaxed">
                    {diff.map((line, idx) => (
                      <div
                        key={idx}
                        className={cn(
                          'flex gap-2 px-2 py-0.5',
                          line.type === 'added' && 'bg-emerald-500/10 text-ws-success',
                          line.type === 'removed' && 'bg-red-500/10 text-ws-danger line-through opacity-70',
                          line.type === 'same' && 'text-ws-faint'
                        )}
                      >
                        <span className="w-4 shrink-0 text-center opacity-60 select-none">
                          {line.type === 'added' ? '+' : line.type === 'removed' ? '−' : ' '}
                        </span>
                        <span className="whitespace-pre-wrap select-text">{line.text || '\u00a0'}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Действия */}
        <div className="flex items-center justify-between gap-2 border-t border-ws px-3 py-2">
          <span className="text-[10px] text-ws-faint">
            {canApply ? 'Enter — сгенерировать · Esc — закрыть' : 'Опишите правку или выберите команду'}
          </span>
          <div className="flex items-center gap-1.5">
            <Button variant="ghost" onClick={onClose}>
              Отмена
            </Button>
            {proposal && !isLoading && (
              <Button variant="secondary" onClick={() => void run(instruction)}>
                <RefreshCw size={11} />
                Ещё раз
              </Button>
            )}
            <Button
              variant="secondary"
              disabled={!instruction.trim() || isLoading}
              onClick={() => void run(instruction)}
            >
              {isLoading ? <Loader2 size={11} className="animate-spin" /> : <Sparkles size={11} />}
              Сгенерировать
            </Button>
            <Button
              variant="primary"
              disabled={!canApply}
              onClick={() => {
                onApply(target, stripCodeFence(proposal));
                onClose();
              }}
            >
              <Check size={11} />
              Применить
              <CornerDownLeft size={10} />
            </Button>
          </div>
        </div>
      </div>
    </>
  );
};

/* ============================================================
   Диффинг строк (LCS)
   ============================================================ */

export interface DiffLine {
  type: 'same' | 'added' | 'removed';
  text: string;
}

/**
 * Построчный diff двух текстов на основе динамического программирования (LCS).
 * Для больших фрагментов таблица ограничивается, чтобы не блокировать UI.
 */
export function computeLineDiff(before: string, after: string): DiffLine[] {
  const a = before.split('\n');
  const b = after.split('\n');

  // Защита от квадратичной сложности на очень больших текстах
  const LIMIT = 600;
  if (a.length > LIMIT || b.length > LIMIT) {
    return [
      ...a.map<DiffLine>((text) => ({ type: 'removed', text })),
      ...b.map<DiffLine>((text) => ({ type: 'added', text })),
    ];
  }

  const n = a.length;
  const m = b.length;
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0));

  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }

  const result: DiffLine[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      result.push({ type: 'same', text: a[i] });
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      result.push({ type: 'removed', text: a[i] });
      i++;
    } else {
      result.push({ type: 'added', text: b[j] });
      j++;
    }
  }
  while (i < n) result.push({ type: 'removed', text: a[i++] });
  while (j < m) result.push({ type: 'added', text: b[j++] });

  return result;
}
