import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowDown,
  ArrowUp,
  Plus,
  Presentation,
  Save,
  Sparkles,
  Trash2,
  Type,
  ImageIcon,
  AlertTriangle,
} from 'lucide-react';
import { OpenTab } from '../../types';
import { Spinner } from '../ui';
import { cn } from '../../lib/utils';
import {
  addPptxImage,
  addPptxSlide,
  addPptxTextBox,
  deletePptxSlide,
  movePptxSlide,
  parsePptx,
  patchPptxText,
  setPptxLayout,
  setPptxNotes,
  type PptxModel,
  type PptxTextEdit,
} from '../../lib/pptx';
import { registerPptxController } from '../../services/pptBridge';

interface Props {
  tab: OpenTab;
  onChange: (content: string) => void;
  onSave?: () => void | Promise<unknown>;
  onAskAI?: (prompt: string, contextText: string) => void;
}

export const PowerPointEditor: React.FC<Props> = ({ tab, onChange, onSave }) => {
  const [model, setModel] = useState<PptxModel | null>(null);
  const [active, setActive] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [selectedShape, setSelectedShape] = useState<{ slide: number; shapeId: string } | null>(null);
  const [notesDraft, setNotesDraft] = useState('');
  const base64Ref = useRef(tab.content);
  const pendingRef = useRef<Map<string, PptxTextEdit>>(new Map());
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const imageInputRef = useRef<HTMLInputElement | null>(null);

  const applyPackage = useCallback(
    async (next: string) => {
      base64Ref.current = next;
      onChange(next);
      try {
        const parsed = await parsePptx(next);
        setModel(parsed);
        setActive((index) => Math.max(0, Math.min(index, parsed.slides.length - 1)));
      } catch {
        /* оставляем текущую модель: пакет уже обновлён */
      }
    },
    [onChange]
  );

  const flush = useCallback(async () => {
    if (timerRef.current) {
      clearTimeout(timerRef.current);
      timerRef.current = null;
    }
    if (pendingRef.current.size === 0) return;
    const edits = [...pendingRef.current.values()];
    pendingRef.current.clear();
    try {
      const next = await patchPptxText(base64Ref.current, edits);
      base64Ref.current = next;
      onChange(next);
    } catch {
      /* текст останется в локальной модели; повторная правка попробует снова */
    }
  }, [onChange]);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      try {
        setLoading(true);
        setError('');
        base64Ref.current = tab.content;
        const parsed = await parsePptx(tab.content);
        if (cancelled) return;
        setModel(parsed);
        setActive(0);
      } catch (err) {
        if (!cancelled) setError((err as Error).message);
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    void load();
    return () => {
      cancelled = true;
      if (timerRef.current) clearTimeout(timerRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  const updateText = useCallback(
    (slideIndex: number, shapeId: string, value: string) => {
      const paragraphs = value.split('\n');
      setModel((prev) => {
        if (!prev) return prev;
        return {
          ...prev,
          slides: prev.slides.map((slide) =>
            slide.index !== slideIndex
              ? slide
              : {
                  ...slide,
                  shapes: slide.shapes.map((shape) =>
                    shape.shapeId !== shapeId
                      ? shape
                      : { ...shape, paragraphs, text: paragraphs.join('\n') }
                  ),
                }
          ),
        };
      });

      pendingRef.current.set(`${slideIndex}:${shapeId}`, { slide: slideIndex, shapeId, paragraphs });
      if (timerRef.current) clearTimeout(timerRef.current);
      timerRef.current = setTimeout(() => void flush(), 500);
    },
    [flush]
  );

  const addSlide = useCallback(async () => {
    await flush();
    try {
      const next = await addPptxSlide(base64Ref.current, active + 1, 'Новый слайд');
      await applyPackage(next);
      setActive((index) => index + 1);
    } catch (err) {
      setError((err as Error).message);
    }
  }, [active, applyPackage, flush]);

  const removeSlide = useCallback(async () => {
    if (!model || model.slides.length <= 1) return;
    await flush();
    try {
      const next = await deletePptxSlide(base64Ref.current, active);
      await applyPackage(next);
    } catch (err) {
      setError((err as Error).message);
    }
  }, [active, applyPackage, flush, model]);

  const moveSlide = useCallback(
    async (delta: number) => {
      if (!model) return;
      const to = active + delta;
      if (to < 0 || to >= model.slides.length) return;
      await flush();
      try {
        const next = await movePptxSlide(base64Ref.current, active, to);
        await applyPackage(next);
        setActive(to);
      } catch (err) {
        setError((err as Error).message);
      }
    },
    [active, applyPackage, flush, model]
  );

  const save = useCallback(async () => {
    await flush();
    await onSave?.();
  }, [flush, onSave]);

  const addTextBox = useCallback(async () => {
    await flush();
    try {
      const next = await addPptxTextBox(base64Ref.current, active, 'Новый текст');
      await applyPackage(next);
    } catch (err) {
      setError((err as Error).message);
    }
  }, [active, applyPackage, flush]);

  useEffect(() => {
    return registerPptxController({
      tabId: tab.id,
      getContent: async () => {
        await flush();
        return base64Ref.current;
      },
      readModel: async () => {
        await flush();
        return parsePptx(base64Ref.current);
      },
      getSelection: () => {
        if (!model) return null;
        const slideIndex = selectedShape?.slide ?? active;
        const slide = model.slides.find((item) => item.index === slideIndex);
        const shape = selectedShape
          ? slide?.shapes.find((item) => item.shapeId === selectedShape.shapeId)
          : slide?.shapes[0];
        if (!shape) return null;
        return {
          slide: slideIndex,
          shapeId: shape.shapeId,
          text: shape.text,
          paragraphs: shape.paragraphs,
        };
      },
      setText: async (slide, shapeId, paragraphs) => {
        try {
          const next = await patchPptxText(base64Ref.current, [{ slide, shapeId, paragraphs }]);
          await applyPackage(next);
          return { ok: true, message: `Слайд ${slide + 1}: текст обновлён` };
        } catch (error) {
          return {
            ok: false,
            message: `Не удалось обновить текст: ${
              error instanceof Error ? error.message : String(error)
            }`,
          };
        }
      },
      addSlide: async (index, title) => {
        try {
          const next = await addPptxSlide(base64Ref.current, index, title ?? 'Новый слайд');
          await applyPackage(next);
          setActive(Math.max(0, Math.min(index, (model?.slides.length ?? 0))));
          return { ok: true, message: 'Слайд добавлен' };
        } catch (error) {
          return { ok: false, message: String(error) };
        }
      },
      deleteSlide: async (index) => {
        try {
          const next = await deletePptxSlide(base64Ref.current, index);
          await applyPackage(next);
          return { ok: true, message: 'Слайд удалён' };
        } catch (error) {
          return { ok: false, message: String(error) };
        }
      },
      moveSlide: async (from, to) => {
        try {
          const next = await movePptxSlide(base64Ref.current, from, to);
          await applyPackage(next);
          return { ok: true, message: 'Порядок слайдов обновлён' };
        } catch (error) {
          return { ok: false, message: String(error) };
        }
      },
      addTextBox: async (slideIndex, text) => {
        try {
          const next = await addPptxTextBox(base64Ref.current, slideIndex, text);
          await applyPackage(next);
          return { ok: true, message: 'Текстовая фигура добавлена' };
        } catch (error) {
          return { ok: false, message: String(error) };
        }
      },
      setNotes: async (slideIndex, text) => {
        try {
          const next = await setPptxNotes(base64Ref.current, slideIndex, text);
          await applyPackage(next);
          return { ok: true, message: `Слайд ${slideIndex + 1}: заметки обновлены` };
        } catch (error) {
          return {
            ok: false,
            message: `Не удалось обновить заметки: ${(error as Error).message}`,
          };
        }
      },
      addImage: async (slideIndex, dataUrl, options) => {
        try {
          const next = await addPptxImage(base64Ref.current, slideIndex, dataUrl, options);
          await applyPackage(next);
          return { ok: true, message: `Слайд ${slideIndex + 1}: изображение добавлено` };
        } catch (error) {
          return {
            ok: false,
            message: `Не удалось добавить изображение: ${(error as Error).message}`,
          };
        }
      },
      setLayout: async (slideIndex, layoutPart) => {
        try {
          const next = await setPptxLayout(base64Ref.current, slideIndex, layoutPart);
          await applyPackage(next);
          return { ok: true, message: `Слайд ${slideIndex + 1}: макет изменён` };
        } catch (error) {
          return {
            ok: false,
            message: `Не удалось изменить макет: ${(error as Error).message}`,
          };
        }
      },
    });
  }, [applyPackage, flush, model, tab.id, active, selectedShape]);

  const current = model?.slides[active] ?? null;
  const notesSlide = current?.index ?? -1;
  const notesValue = current?.notes ?? '';

  useEffect(() => {
    setNotesDraft(notesValue);
  }, [notesSlide, notesValue]);

  const saveNotes = useCallback(async () => {
    if (!current || notesDraft === current.notes) return;
    try {
      const next = await setPptxNotes(base64Ref.current, current.index, notesDraft);
      await applyPackage(next);
    } catch (err) {
      setError((err as Error).message);
    }
  }, [applyPackage, current, notesDraft]);

  const insertImage = useCallback(
    async (file: File) => {
      if (!current) return;
      try {
        await flush();
        const dataUrl = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(String(reader.result ?? ''));
          reader.onerror = () => reject(reader.error ?? new Error('не удалось прочитать файл'));
          reader.readAsDataURL(file);
        });
        const next = await addPptxImage(base64Ref.current, current.index, dataUrl, {
          alt: file.name,
        });
        await applyPackage(next);
      } catch (err) {
        setError((err as Error).message);
      }
    },
    [applyPackage, current, flush]
  );
  const changeLayout = useCallback(
    async (layoutPart: string) => {
      if (!current || !layoutPart) return;
      try {
        await flush();
        const next = await setPptxLayout(base64Ref.current, current.index, layoutPart);
        await applyPackage(next);
      } catch (err) {
        setError((err as Error).message);
      }
    },
    [applyPackage, current, flush]
  );

  const dark = useMemo(() => {
    const hex = (current?.background ?? '#ffffff').replace('#', '');
    if (hex.length !== 6) return false;
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    return 0.299 * r + 0.587 * g + 0.114 * b < 140;
  }, [current]);

  if (loading) {
    return (
      <div className="flex h-full items-center justify-center bg-ws-editor">
        <div className="flex items-center gap-2 text-xs text-ws-muted">
          <Spinner size={12} />
          Загрузка презентации…
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-full w-full flex-col bg-ws-editor text-ws">
      <div className="flex shrink-0 items-center justify-between gap-2 border-b border-ws bg-ws-surface px-3 py-1.5 text-xs">
        <div className="flex items-center gap-1">
          <span className="mr-1 flex items-center gap-1.5 rounded border border-orange-500/40 bg-orange-500/10 px-2 py-0.5 font-medium text-orange-500">
            <Presentation size={13} />
            PowerPoint
          </span>
          <button
            onClick={addSlide}
            title="Добавить слайд"
            className="flex items-center gap-1 rounded px-1.5 py-1 hover:bg-ws-hover"
          >
            <Plus size={14} />
            Слайд
          </button>
          <button
            onClick={() => void addTextBox()}
            title="Добавить текст"
            className="flex items-center gap-1 rounded px-1.5 py-1 hover:bg-ws-hover"
          >
            <Type size={14} />
            Текст
          </button>
          <button
            onClick={() => imageInputRef.current?.click()}
            title="Вставить изображение"
            className="flex items-center gap-1 rounded px-1.5 py-1 hover:bg-ws-hover"
          >
            <ImageIcon size={14} />
            Изображение
          </button>
          <input
            ref={imageInputRef}
            type="file"
            accept="image/*"
            className="hidden"
            onChange={(event) => {
              const file = event.target.files?.[0];
              event.target.value = '';
              if (file) void insertImage(file);
            }}
          />
          {model && model.layouts.length > 0 && (
            <select
              value={current?.layoutPart ?? ''}
              onChange={(event) => void changeLayout(event.target.value)}
              title="Макет слайда"
              className="rounded border border-ws bg-ws-elevated px-1.5 py-1 text-[11px] text-ws"
            >
              {model.layouts.map((layout) => (
                <option key={layout.part} value={layout.part}>
                  {layout.name}
                </option>
              ))}
            </select>
          )}
          <button
            onClick={() => void moveSlide(-1)}
            disabled={active <= 0}
            title="Выше"
            className="rounded p-1 hover:bg-ws-hover disabled:opacity-40"
          >
            <ArrowUp size={14} />
          </button>
          <button
            onClick={() => void moveSlide(1)}
            disabled={!model || active >= model.slides.length - 1}
            title="Ниже"
            className="rounded p-1 hover:bg-ws-hover disabled:opacity-40"
          >
            <ArrowDown size={14} />
          </button>
          <button
            onClick={removeSlide}
            disabled={!model || model.slides.length <= 1}
            title="Удалить слайд"
            className="rounded p-1 text-ws-danger hover:bg-red-500/10 disabled:opacity-40"
          >
            <Trash2 size={14} />
          </button>
        </div>
        <button
          onClick={() => void save()}
          className="flex items-center gap-1 rounded border border-ws bg-ws-elevated px-2 py-1 hover:bg-ws-hover"
        >
          <Save size={12} />
          Сохранить
        </button>
      </div>

      {error && (
        <div className="flex shrink-0 items-center gap-2 border-b border-red-500/40 bg-red-500/10 px-3 py-1 text-[11px] text-ws-danger">
          <AlertTriangle size={11} />
          {error}
        </div>
      )}

      <div className="flex min-h-0 flex-1">
        <aside className="w-40 shrink-0 overflow-y-auto border-r border-ws bg-ws-surface p-2">
          {(model?.slides ?? []).map((slide) => {
            const title = slide.shapes.find((shape) => shape.isTitle)?.text || slide.shapes[0]?.text;
            return (
              <button
                key={slide.index}
                onClick={() => setActive(slide.index)}
                className={cn(
                  'mb-1 flex w-full gap-1.5 rounded border p-1 text-left text-[10px]',
                  slide.index === active
                    ? 'border-ws-accent bg-ws-accent-soft text-ws-accent'
                    : 'border-ws bg-ws-elevated hover:bg-ws-hover'
                )}
              >
                <span className="w-4 shrink-0 text-center font-mono text-ws-faint">
                  {slide.index + 1}
                </span>
                <span className="truncate">{title || 'Пустой слайд'}</span>
              </button>
            );
          })}
        </aside>

        <div className="min-h-0 flex-1 overflow-y-auto bg-ws-editor-alt p-6">
          {current ? (
            <>
            <div
              className="relative mx-auto w-full max-w-[900px] overflow-hidden rounded border border-ws shadow-ws"
              style={{ aspectRatio: `${model?.width ?? 12192000} / ${model?.height ?? 6858000}`, background: current.background }}
            >
              {current.shapes.map((shape) => (
                <textarea
                  key={`${current.index}:${shape.shapeId}`}
                  value={shape.paragraphs.join('\n')}
                  onChange={(event) => updateText(current.index, shape.shapeId, event.target.value)}
                  onFocus={() => setSelectedShape({ slide: current.index, shapeId: shape.shapeId })}
                  onBlur={() => void flush()}
                  spellCheck={false}
                  className={cn(
                    'absolute resize-none overflow-hidden rounded border border-transparent bg-transparent p-1 leading-snug outline-none focus:border-blue-400 focus:bg-blue-500/5',
                    shape.isTitle && 'font-semibold'
                  )}
                  style={{
                    left: `${(shape.x / (model?.width ?? 12192000)) * 100}%`,
                    top: `${(shape.y / (model?.height ?? 6858000)) * 100}%`,
                    width: `${(shape.w / (model?.width ?? 12192000)) * 100}%`,
                    height: `${(shape.h / (model?.height ?? 6858000)) * 100}%`,
                    fontSize: `${Math.max(11, Math.round((shape.fontSize ?? (shape.isTitle ? 28 : 16)) * 0.85))}px`,
                    fontWeight: shape.bold || shape.isTitle ? 700 : 400,
                    color: shape.color ?? (dark ? '#F8FAFC' : '#111827'),
                  }}
                />
              ))}
              {current.shapes.length === 0 && (
                <div className="flex h-full items-center justify-center gap-2 text-xs text-ws-faint">
                  <Sparkles size={12} />
                  На слайде нет текста. Попросите AI добавить текст или создайте слайд заново.
                </div>
              )}
            </div>
            <div className="mx-auto mt-3 max-w-[900px] rounded border border-ws bg-ws-surface p-3 text-[11px] text-ws">
              <div className="mb-1 font-medium text-ws-muted">Заметки докладчика</div>
              <textarea
                value={notesDraft}
                onChange={(event) => setNotesDraft(event.target.value)}
                onBlur={() => void saveNotes()}
                placeholder="Заметки к слайду"
                spellCheck={false}
                className="h-20 w-full resize-y rounded border border-ws bg-ws-editor px-2 py-1 text-[11px] text-ws outline-none focus:border-ws-accent"
              />
            </div>
            </>
          ) : (
            <div className="text-xs text-ws-faint">Слайды не найдены</div>
          )}
        </div>
      </div>

      <div className="flex shrink-0 items-center justify-between border-t border-ws bg-ws-surface px-3 py-1 text-[10px] text-ws-faint">
        <span>PPTX · {tab.name}</span>
        <span>
          Слайд {Math.min(active + 1, model?.slides.length ?? 1)} из {model?.slides.length ?? 0}
        </span>
      </div>
    </div>
  );
};
