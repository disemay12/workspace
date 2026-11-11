import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  ArrowDown,
  ArrowUp,
  Bold,
  ChevronLeft,
  ChevronRight,
  Copy,
  Eye,
  EyeOff,
  FileDown,
  ImageIcon,
  Layers,
  Maximize2,
  Minus,
  Palette,
  Plus,
  Presentation,
  Redo2,
  Save,
  Search,
  Square,
  Table2,
  Trash2,
  Type,
  Undo2,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import { OpenTab } from '../../types';
import { Spinner } from '../ui';
import { cn, baseName } from '../../lib/utils';
import { SlideCanvas } from '../ppt/SlideCanvas';
import { ThemePanel } from '../ppt/ThemePanel';
import { Inspector, type TextFormatPatch } from '../ppt/Inspector';
import {
  openPptxDocument,
  type PptxDocModel,
  type PptxDocument,
  type PptxParagraph,
  type PptxShapeModel,
  type PptxSlideModel,
  type PptxTextBody,
  type PptxTransitionModel,
} from '../../lib/pptxDoc';
import {
  THEME_PRESETS,
  effectPresetById,
  normalizeHexColor,
  type ThemeColorSet,
  type ThemeFontPair,
} from '../../lib/pptxThemes';
import { registerPptxController } from '../../services/pptBridge';

interface Props {
  tab: OpenTab;
  onChange: (content: string) => void;
  onSave?: () => void | Promise<unknown>;
  onAskAI?: (prompt: string, contextText: string) => void;
}

const EMU_PER_INCH = 914400;
const EMU_PER_CM = 360000;
const SYNC_DELAY = 900;

/** Слайд 16:9 и 4:3 в EMU — быстрый переключатель размера. */
const SLIDE_SIZES: { id: string; label: string; width: number; height: number }[] = [
  { id: '16:9', label: '16:9', width: 12192000, height: 6858000 },
  { id: '4:3', label: '4:3', width: 9144000, height: 6858000 },
  { id: 'a4', label: 'A4 (альбом)', width: 10692000, height: 7560000 },
];

const SHAPES: { id: string; label: string }[] = [
  { id: 'rect', label: 'Прямоугольник' },
  { id: 'roundRect', label: 'Скруглённый' },
  { id: 'ellipse', label: 'Эллипс' },
  { id: 'triangle', label: 'Треугольник' },
  { id: 'line', label: 'Линия' },
  { id: 'arrow', label: 'Стрелка' },
  { id: 'star5', label: 'Звезда' },
];

type DragState = {
  mode: 'move' | 'resize';
  handle?: 'nw' | 'ne' | 'sw' | 'se';
  shapeId: string;
  startClientX: number;
  startClientY: number;
  start: { x: number; y: number; w: number; h: number };
  emuPerPx: number;
};

export const PowerPointEditor: React.FC<Props> = ({ tab, onChange, onSave }) => {
  const [model, setModel] = useState<PptxDocModel | null>(null);
  const [active, setActive] = useState(0);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [zoom, setZoom] = useState(1);
  const [present, setPresent] = useState(false);
  const [presentIndex, setPresentIndex] = useState(0);
  const [themeOpen, setThemeOpen] = useState(false);
  const [shadowDefault, setShadowDefault] = useState(true);
  const [editingText, setEditingText] = useState<{ shapeId: string; value: string } | null>(null);
  const [findOpen, setFindOpen] = useState(false);
  const [findQuery, setFindQuery] = useState('');
  const [replaceQuery, setReplaceQuery] = useState('');
  const [shapeMenu, setShapeMenu] = useState(false);
  const [override, setOverride] = useState<Record<string, { x: number; y: number; w: number; h: number }>>({});
  const [containerWidth, setContainerWidth] = useState(900);
  const [transitionIn, setTransitionIn] = useState(1);

  const docRef = useRef<PptxDocument | null>(null);
  const base64Ref = useRef(tab.content);
  const syncTimer = useRef<number | null>(null);
  const dragRef = useRef<DragState | null>(null);
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const imageInputRef = useRef<HTMLInputElement | null>(null);
  const noticeTimer = useRef<number | null>(null);

  const flash = useCallback((message: string) => {
    setNotice(message);
    if (noticeTimer.current) window.clearTimeout(noticeTimer.current);
    noticeTimer.current = window.setTimeout(() => setNotice(''), 2600);
  }, []);

  /* ---------------- Загрузка документа ---------------- */

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError('');
    setSelectedId(null);
    setActive(0);
    const load = async () => {
      try {
        const doc = await openPptxDocument(tab.content);
        if (cancelled) return;
        docRef.current = doc;
        base64Ref.current = tab.content;
        setModel({ ...doc.model });
      } catch (err) {
        if (!cancelled) setError(`Не удалось открыть презентацию: ${(err as Error).message}`);
      } finally {
        if (!cancelled) setLoading(false);
      }
      const doc = docRef.current;
      if (!doc || cancelled) return;
      // Изображения грузим после первого кадра: большой пакет не должен тормозить открытие.
      await doc.loadMedia().catch(() => undefined);
      if (!cancelled) setModel({ ...doc.model });
    };
    void load();
    return () => {
      cancelled = true;
      if (syncTimer.current) window.clearTimeout(syncTimer.current);
      if (noticeTimer.current) window.clearTimeout(noticeTimer.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  /* ---------------- Сохранение ---------------- */

  const syncNow = useCallback(async () => {
    const doc = docRef.current;
    if (!doc) return base64Ref.current;
    if (syncTimer.current) {
      window.clearTimeout(syncTimer.current);
      syncTimer.current = null;
    }
    const base64 = await doc.toBase64();
    base64Ref.current = base64;
    onChange(base64);
    return base64;
  }, [onChange]);

  const scheduleSync = useCallback(() => {
    if (syncTimer.current) window.clearTimeout(syncTimer.current);
    syncTimer.current = window.setTimeout(() => {
      syncTimer.current = null;
      void syncNow();
    }, SYNC_DELAY);
  }, [syncNow]);

  const apply = useCallback(
    (label: string, action: (doc: PptxDocument) => void) => {
      const doc = docRef.current;
      if (!doc) return;
      try {
        action(doc);
        setModel({ ...doc.model });
        setActive((index) => Math.min(index, Math.max(0, doc.model.slides.length - 1)));
        scheduleSync();
      } catch (err) {
        setError(`${label}: ${(err as Error).message}`);
      }
    },
    [scheduleSync]
  );

  const save = useCallback(async () => {
    await syncNow();
    await onSave?.();
  }, [onSave, syncNow]);

  /* ---------------- Размеры холста ---------------- */

  useEffect(() => {
    const element = wrapRef.current;
    if (!element) return;
    const update = () => setContainerWidth(element.clientWidth);
    update();
    const observer = new ResizeObserver(update);
    observer.observe(element);
    return () => observer.disconnect();
  }, [loading]);

  const canvasWidth = Math.max(280, Math.round((containerWidth - 56) * zoom));
  const slide = model?.slides[active] ?? null;
  const selectedShape = useMemo(
    () => slide?.shapes.find((shape) => shape.id === selectedId) ?? null,
    [slide, selectedId]
  );

  /* ---------------- Перетаскивание и ресайз ---------------- */

  const beginMove = useCallback(
    (shape: PptxShapeModel, event: React.PointerEvent) => {
      if (present || !model) return;
      setSelectedId(shape.id);
      const rect = (event.currentTarget as HTMLElement).closest('[data-slide-canvas]')?.getBoundingClientRect();
      if (!rect) return;
      dragRef.current = {
        mode: 'move',
        shapeId: shape.id,
        startClientX: event.clientX,
        startClientY: event.clientY,
        start: { x: shape.x, y: shape.y, w: shape.w, h: shape.h },
        emuPerPx: model.width / rect.width,
      };
      setOverride({ [shape.id]: { x: shape.x, y: shape.y, w: shape.w, h: shape.h } });
    },
    [model, present]
  );

  const beginResize = useCallback(
    (handle: 'nw' | 'ne' | 'sw' | 'se', event: React.PointerEvent) => {
      if (!model || !selectedShape) return;
      event.stopPropagation();
      dragRef.current = {
        mode: 'resize',
        handle,
        shapeId: selectedShape.id,
        startClientX: event.clientX,
        startClientY: event.clientY,
        start: {
          x: selectedShape.x,
          y: selectedShape.y,
          w: selectedShape.w,
          h: selectedShape.h,
        },
        // Масштаб берём по ширине холста, а не контейнера прокрутки.
        emuPerPx: model.width / Math.max(1, canvasWidth),
      };
      setOverride({
        [selectedShape.id]: {
          x: selectedShape.x,
          y: selectedShape.y,
          w: selectedShape.w,
          h: selectedShape.h,
        },
      });
    },
    [canvasWidth, model, selectedShape]
  );

  useEffect(() => {
    const onPointerMove = (event: PointerEvent) => {
      const drag = dragRef.current;
      if (!drag) return;
      const dx = (event.clientX - drag.startClientX) * drag.emuPerPx;
      const dy = (event.clientY - drag.startClientY) * drag.emuPerPx;
      let next = { ...drag.start };
      if (drag.mode === 'move') {
        next.x = drag.start.x + dx;
        next.y = drag.start.y + dy;
      } else {
        const minSize = 12700;
        if (drag.handle?.includes('e')) next.w = Math.max(minSize, drag.start.w + dx);
        if (drag.handle?.includes('s')) next.h = Math.max(minSize, drag.start.h + dy);
        if (drag.handle?.includes('w')) {
          next.w = Math.max(minSize, drag.start.w - dx);
          next.x = drag.start.x + (drag.start.w - next.w);
        }
        if (drag.handle?.includes('n')) {
          next.h = Math.max(minSize, drag.start.h - dy);
          next.y = drag.start.y + (drag.start.h - next.h);
        }
      }
      next = {
        x: Math.round(next.x),
        y: Math.round(next.y),
        w: Math.round(next.w),
        h: Math.round(next.h),
      };
      setOverride({ [drag.shapeId]: next });
    };

    const onPointerUp = () => {
      const drag = dragRef.current;
      if (!drag) return;
      dragRef.current = null;
      const geometry = override[drag.shapeId];
      setOverride({});
      if (!geometry) return;
      const unchanged =
        geometry.x === drag.start.x &&
        geometry.y === drag.start.y &&
        geometry.w === drag.start.w &&
        geometry.h === drag.start.h;
      if (unchanged) return;
      apply('Геометрия фигуры', (doc) => {
        doc.setShapeGeometry(active, drag.shapeId, geometry);
      });
    };

    window.addEventListener('pointermove', onPointerMove);
    window.addEventListener('pointerup', onPointerUp);
    window.addEventListener('pointercancel', onPointerUp);
    return () => {
      window.removeEventListener('pointermove', onPointerMove);
      window.removeEventListener('pointerup', onPointerUp);
      window.removeEventListener('pointercancel', onPointerUp);
    };
  }, [active, apply, override]);

  /* ---------------- Действия со слайдами ---------------- */

  const addSlide = useCallback(() => {
    apply('Новый слайд', (doc) => {
      const index = doc.addSlide(active + 1, slide?.layoutPart);
      setActive(index);
      setSelectedId(null);
    });
  }, [active, apply, slide?.layoutPart]);

  const duplicateSlide = useCallback(() => {
    apply('Дублирование слайда', (doc) => {
      const index = doc.duplicateSlide(active);
      setActive(index);
    });
  }, [active, apply]);

  const deleteSlide = useCallback(() => {
    if (!model || model.slides.length <= 1) return;
    apply('Удаление слайда', (doc) => {
      doc.deleteSlide(active);
      setActive(Math.max(0, Math.min(active, doc.model.slides.length - 1)));
      setSelectedId(null);
    });
  }, [active, apply, model]);

  const moveSlide = useCallback(
    (delta: number) => {
      if (!model) return;
      const target = active + delta;
      if (target < 0 || target >= model.slides.length) return;
      apply('Перестановка слайдов', (doc) => {
        const index = doc.moveSlide(active, target);
        setActive(index);
      });
    },
    [active, apply, model]
  );

  const toggleHidden = useCallback(() => {
    if (!slide) return;
    apply('Видимость слайда', (doc) => doc.setSlideHidden(active, !slide.hidden));
  }, [active, apply, slide]);

  const changeLayout = useCallback(
    (layoutPart: string) => {
      apply('Макет слайда', (doc) => doc.setSlideLayout(active, layoutPart));
    },
    [active, apply]
  );

  const changeSlideSize = useCallback(
    (id: string) => {
      const size = SLIDE_SIZES.find((item) => item.id === id);
      if (!size) return;
      apply('Размер слайда', (doc) => doc.setSlideSize(size.width, size.height));
    },
    [apply]
  );

  /* ---------------- Действия с фигурами ---------------- */

  const addTextBox = useCallback(() => {
    apply('Надпись', (doc) => {
      const id = doc.addTextBox(active, { text: 'Новый текст' });
      if (id) setSelectedId(id);
    });
  }, [active, apply]);

  const addShape = useCallback(
    (geom: string) => {
      setShapeMenu(false);
      apply('Фигура', (doc) => {
        const theme = doc.model.theme;
        const id = doc.addShape(active, {
          geom,
          fill: geom === 'line' ? null : theme.colors.accent1,
          outline: { color: theme.colors.dk2, width: 1 },
          shadow: shadowDefault,
        });
        if (id) setSelectedId(id);
      });
    },
    [active, apply, shadowDefault]
  );

  const addTable = useCallback(() => {
    apply('Таблица', (doc) => {
      const id = doc.addTable(active, { rows: 3, cols: 3 });
      if (id) setSelectedId(id);
    });
  }, [active, apply]);

  const insertImage = useCallback(
    async (file: File) => {
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result ?? ''));
        reader.onerror = () => reject(reader.error ?? new Error('не удалось прочитать файл'));
        reader.readAsDataURL(file);
      });
      apply('Изображение', (doc) => {
        const id = doc.addImage(active, dataUrl, { alt: file.name });
        if (id) setSelectedId(id);
      });
    },
    [active, apply]
  );

  const deleteShape = useCallback(() => {
    if (!selectedId) return;
    apply('Удаление фигуры', (doc) => doc.deleteShape(active, selectedId));
    setSelectedId(null);
  }, [active, apply, selectedId]);

  const duplicateShape = useCallback(() => {
    if (!selectedId) return;
    apply('Дублирование фигуры', (doc) => {
      const id = doc.duplicateShape(active, selectedId);
      if (id) setSelectedId(id);
    });
  }, [active, apply, selectedId]);

  const setGeometry = useCallback(
    (patch: { x?: number; y?: number; w?: number; h?: number; rotation?: number }) => {
      if (!selectedShape) return;
      apply('Геометрия фигуры', (doc) => {
        doc.setShapeGeometry(active, selectedShape.id, {
          x: patch.x ?? selectedShape.x,
          y: patch.y ?? selectedShape.y,
          w: patch.w ?? selectedShape.w,
          h: patch.h ?? selectedShape.h,
          rotation: patch.rotation ?? selectedShape.rotation ?? 0,
        });
      });
    },
    [active, apply, selectedShape]
  );

  const formatText = useCallback(
    (patch: TextFormatPatch) => {
      if (!selectedShape) return;
      const body = selectedShape.text;
      if (!body) {
        flash('У этой фигуры нет текста');
        return;
      }
      apply('Формат текста', (doc) => {
        const paragraphs = body.paragraphs.map((paragraph) => ({
          ...paragraph,
          runs:
            paragraph.runs.length > 0
              ? paragraph.runs.map((run) => ({
                  ...run,
                  ...('bold' in patch ? { bold: patch.bold } : {}),
                  ...('italic' in patch ? { italic: patch.italic } : {}),
                  ...('underline' in patch ? { underline: patch.underline } : {}),
                  ...('strike' in patch ? { strike: patch.strike } : {}),
                  ...(patch.size !== undefined ? { size: patch.size } : {}),
                  ...(patch.family !== undefined ? { family: patch.family } : {}),
                  ...(patch.color
                    ? { color: normalizeHexColor(patch.color), schemeColor: undefined }
                    : {}),
                }))
              : [
                  {
                    text: ' ',
                    size: patch.size ?? 18,
                    ...(patch.color ? { color: normalizeHexColor(patch.color) } : {}),
                    ...(patch.bold !== undefined ? { bold: patch.bold } : {}),
                  },
                ],
        }));
        doc.setShapeText(active, selectedShape.id, { ...body, paragraphs });
      });
    },
    [active, apply, flash, selectedShape]
  );

  const formatParagraph = useCallback(
    (patch: Partial<PptxParagraph>) => {
      if (!selectedShape?.text) return;
      const body = selectedShape.text;
      apply('Абзац', (doc) => {
        const paragraphs = body.paragraphs.map((paragraph) => ({ ...paragraph, ...patch }));
        doc.setShapeText(active, selectedShape.id, { ...body, paragraphs });
      });
    },
    [active, apply, selectedShape]
  );

  const setTextBody = useCallback(
    (patch: Partial<Pick<PptxTextBody, 'verticalAnchor' | 'autofit'>>) => {
      if (!selectedShape?.text) return;
      const body = selectedShape.text;
      apply('Текст фигуры', (doc) => {
        doc.setShapeText(active, selectedShape.id, { ...body, ...patch });
      });
    },
    [active, apply, selectedShape]
  );

  /* ---------------- Правка текста двойным кликом ---------------- */

  const beginTextEdit = useCallback(
    (shape: PptxShapeModel) => {
      if (!shape.text) {
        flash('Двойной клик по текстовой фигуре открывает правку текста');
        return;
      }
      const value = shape.text.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
      setEditingText({ shapeId: shape.id, value });
      setSelectedId(shape.id);
    },
    [flash]
  );

  const commitTextEdit = useCallback(() => {
    const draft = editingText;
    setEditingText(null);
    if (!draft) return;
    const shape = slide?.shapes.find((item) => item.id === draft.shapeId);
    if (!shape?.text) return;
    const lines = draft.value.split('\n');
    const sourceParagraphs = shape.text.paragraphs;
    const fallbackRun = sourceParagraphs.flatMap((p) => p.runs)[0] ?? { text: '', size: 18 };
    apply('Правка текста', (doc) => {
      const paragraphs = lines.map((line, index) => {
        const template = sourceParagraphs[Math.min(index, sourceParagraphs.length - 1)];
        const runTemplate = template?.runs[0] ?? fallbackRun;
        return {
          align: template?.align,
          level: template?.level,
          bullet: template?.bullet,
          lineSpacing: template?.lineSpacing,
          spaceBefore: template?.spaceBefore,
          spaceAfter: template?.spaceAfter,
          runs: line ? [{ ...runTemplate, text: line }] : [],
        };
      });
      doc.setShapeText(active, shape.id, { ...shape.text!, paragraphs });
    });
  }, [active, apply, editingText, slide]);

  /* ---------------- Фон, переход, тема ---------------- */

  const setBackground = useCallback(
    (color: string, all: boolean) => {
      apply('Фон слайда', (doc) => doc.setSlideBackground(active, { color, all }));
    },
    [active, apply]
  );

  const setTransition = useCallback(
    (transition: PptxTransitionModel | undefined) => {
      apply('Переход', (doc) =>
        doc.setSlideTransition(active, transition ?? { type: 'none' })
      );
    },
    [active, apply]
  );

  const applyThemePreset = useCallback(
    (presetId: string, label: string) => {
      apply('Тема', (doc) => doc.applyTheme({ presetId, label }));
      flash(`Тема: ${label}`);
    },
    [apply, flash]
  );

  const applyThemeColors = useCallback(
    (colors: Partial<ThemeColorSet>, label: string) => {
      apply('Цвета темы', (doc) => doc.applyTheme({ colors, label }));
      flash(`Цвета схемы обновлены (${label})`);
    },
    [apply, flash]
  );

  const applyThemeFonts = useCallback(
    (fonts: ThemeFontPair) => {
      apply('Шрифты темы', (doc) => doc.applyTheme({ fonts }));
      flash(`Шрифты темы: ${fonts.label}`);
    },
    [apply, flash]
  );

  const applyEffects = useCallback((effectId: string) => {
    setShadowDefault(effectPresetById(effectId).shadow);
  }, []);

  /* ---------------- Поиск и замена ---------------- */

  const runFindReplace = useCallback(() => {
    if (!findQuery.trim()) return;
    const query = findQuery;
    const replacement = replaceQuery;
    apply('Поиск и замена', (doc) => {
      const result = doc.findReplace(query, replacement, { includeNotes: true });
      flash(
        result.count > 0
          ? `Заменено вхождений: ${result.count} (слайдов: ${result.slides.length})`
          : 'Совпадений не найдено'
      );
    });
    setFindOpen(false);
  }, [apply, findQuery, flash, replaceQuery]);

  /* ---------------- Отмена/повтор ---------------- */

  const undo = useCallback(() => {
    const doc = docRef.current;
    if (!doc) return;
    if (doc.undo()) {
      setModel({ ...doc.model });
      setActive((index) => Math.min(index, Math.max(0, doc.model.slides.length - 1)));
      scheduleSync();
    }
  }, [scheduleSync]);

  const redo = useCallback(() => {
    const doc = docRef.current;
    if (!doc) return;
    if (doc.redo()) {
      setModel({ ...doc.model });
      setActive((index) => Math.min(index, Math.max(0, doc.model.slides.length - 1)));
      scheduleSync();
    }
  }, [scheduleSync]);

  /* ---------------- Экспорт PDF ---------------- */

  const exportPdf = useCallback(async () => {
    const doc = docRef.current;
    if (!doc) return;
    const current = doc.model;
    const html = buildPdfHtml(current);
    const pageSize = {
      width: Math.round((current.width / EMU_PER_INCH) * 25400),
      height: Math.round((current.height / EMU_PER_INCH) * 25400),
    };
    const result = await window.electronAPI?.exportPdf?.(
      html,
      baseName(tab.name).replace(/\.[^.]+$/, '') + '.pdf',
      undefined,
      pageSize
    );
    if (result?.success) flash('PDF экспортирован: ' + result.path);
    else if (!result?.canceled) flash('Не удалось экспортировать PDF');
  }, [flash, tab.name]);

  /* ---------------- Клавиатура ---------------- */

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (present) {
        if (event.key === 'ArrowRight' || event.key === ' ' || event.key === 'PageDown') {
          event.preventDefault();
          setPresentIndex((index) => Math.min((model?.slides.length ?? 1) - 1, index + 1));
          setTransitionIn(0);
        } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
          event.preventDefault();
          setPresentIndex((index) => Math.max(0, index - 1));
          setTransitionIn(0);
        } else if (event.key === 'Escape') {
          setPresent(false);
        }
        return;
      }

      const target = event.target as HTMLElement | null;
      const inField = target && ['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName);

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        void save();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') {
        event.preventDefault();
        if (event.shiftKey) redo();
        else undo();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y') {
        event.preventDefault();
        redo();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'd') {
        event.preventDefault();
        if (selectedId) duplicateShape();
        else duplicateSlide();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') {
        event.preventDefault();
        setFindOpen(true);
        return;
      }
      if (inField) return;

      if (event.key === 'Delete' && selectedId) {
        event.preventDefault();
        deleteShape();
        return;
      }
      if (event.key === 'Escape') {
        setSelectedId(null);
        return;
      }
      if (event.key.startsWith('Arrow') && selectedShape) {
        event.preventDefault();
        const step = event.shiftKey ? 9525 * 10 : 9525;
        const dx = event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0;
        const dy = event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0;
        setGeometry({ x: selectedShape.x + dx, y: selectedShape.y + dy });
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [
    deleteShape,
    duplicateShape,
    duplicateSlide,
    model?.slides.length,
    present,
    redo,
    save,
    selectedId,
    selectedShape,
    setGeometry,
    undo,
  ]);

  /* ---------------- Показ слайдов ---------------- */

  useEffect(() => {
    if (!present) return;
    setPresentIndex(active);
    const frame = window.requestAnimationFrame(() => setTransitionIn(1));
    return () => window.cancelAnimationFrame(frame);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [present]);

  /* ---------------- AI-мост ---------------- */

  useEffect(() => {
    return registerPptxController({
      tabId: tab.id,
      getContent: async () => await syncNow(),
      readModel: async () => {
        const doc = docRef.current;
        if (!doc) throw new Error('Документ не открыт');
        return doc.model;
      },
      readLegacyModel: async () => {
        const doc = docRef.current;
        if (!doc) throw new Error('Документ не открыт');
        return doc.toLegacyModel();
      },
      getSlide: async (index) => docRef.current?.model.slides[index] ?? null,
      listLayouts: async () => docRef.current?.model.layouts ?? [],
      listThemes: async () => THEME_PRESETS.map((preset) => ({ id: preset.id, label: preset.label })),
      getShape: async (slideIndex, shapeId) =>
        docRef.current?.model.slides[slideIndex]?.shapes.find((shape) => shape.id === shapeId) ?? null,
      getSelection: () => {
        const doc = docRef.current;
        if (!doc) return null;
        const current = doc.model.slides[active];
        const shape = selectedId
          ? current?.shapes.find((item) => item.id === selectedId)
          : current?.shapes.find((item) => item.text);
        if (!shape?.text) return null;
        const paragraphs = shape.text.paragraphs.map((paragraph) =>
          paragraph.runs.map((run) => run.text).join('')
        );
        return {
          slide: active,
          shapeId: shape.id,
          text: paragraphs.join('\n'),
          paragraphs,
        };
      },
      setText: async (slideIndex, shapeId, paragraphs) => {
        try {
          await applyAsync((doc) => {
            const current = doc.model.slides[slideIndex];
            const shape = current?.shapes.find((item) => item.id === shapeId);
            const fallback = shape?.text?.paragraphs[0]?.runs[0] ?? { text: '', size: 18 };
            doc.setShapeText(slideIndex, shapeId, {
              ...(shape?.text ?? {}),
              paragraphs: paragraphs.map((line) => ({
                runs: line ? [{ ...fallback, text: line }] : [],
              })),
            });
          });
          return { ok: true, message: `Слайд ${slideIndex + 1}: текст обновлён` };
        } catch (err) {
          return { ok: false, message: `Не удалось обновить текст: ${(err as Error).message}` };
        }
      },
      addSlide: async (index, title) => {
        try {
          await applyAsync((doc) => {
            const created = doc.addSlide(index, undefined, title ?? 'Новый слайд');
            setActive(created);
          });
          return { ok: true, message: 'Слайд добавлен' };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      deleteSlide: async (index) => {
        try {
          await applyAsync((doc) => doc.deleteSlide(index));
          return { ok: true, message: 'Слайд удалён' };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      moveSlide: async (from, to) => {
        try {
          await applyAsync((doc) => {
            const index = doc.moveSlide(from, to);
            setActive(index);
          });
          return { ok: true, message: 'Порядок слайдов обновлён' };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      addTextBox: async (slideIndex, text) => {
        try {
          await applyAsync((doc) => {
            doc.addTextBox(slideIndex, { text });
          });
          return { ok: true, message: 'Текстовая фигура добавлена' };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      setNotes: async (slideIndex, text) => {
        try {
          await applyAsync((doc) => doc.setNotes(slideIndex, text));
          return { ok: true, message: `Слайд ${slideIndex + 1}: заметки обновлены` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      addImage: async (slideIndex, dataUrl, options) => {
        try {
          await applyAsync((doc) => {
            doc.addImage(slideIndex, dataUrl, options);
          });
          return { ok: true, message: `Слайд ${slideIndex + 1}: изображение добавлено` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      setLayout: async (slideIndex, layoutPart) => {
        try {
          await applyAsync((doc) => doc.setSlideLayout(slideIndex, layoutPart));
          return { ok: true, message: `Слайд ${slideIndex + 1}: макет изменён` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },

      /* ---------- Слайды ---------- */
      duplicateSlide: async (index) => {
        try {
          await applyAsync((doc) => {
            const created = doc.duplicateSlide(index);
            setActive(created);
          });
          return { ok: true, message: `Слайд ${index + 1} продублирован` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      setSlideHidden: async (index, hidden) => {
        try {
          await applyAsync((doc) => doc.setSlideHidden(index, hidden));
          return { ok: true, message: `Слайд ${index + 1}: ${hidden ? 'скрыт' : 'показан'}` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      setSlideSize: async (width, height) => {
        try {
          await applyAsync((doc) => doc.setSlideSize(width, height));
          return {
            ok: true,
            message: `Размер слайда: ${Math.round(width / 36000)}×${Math.round(height / 36000)} см`,
          };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      setTransition: async (index, transition) => {
        try {
          await applyAsync((doc) => doc.setSlideTransition(index, transition ?? { type: 'none' }));
          return { ok: true, message: `Слайд ${index + 1}: переход ${transition?.type ?? 'none'}` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      setBackground: async (index, background) => {
        try {
          await applyAsync((doc) => doc.setSlideBackground(index, background));
          return {
            ok: true,
            message: background.all ? 'Фон применён ко всем слайдам' : `Слайд ${index + 1}: фон обновлён`,
          };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },

      /* ---------- Текст ---------- */
      setTextBody: async (slideIndex, shapeId, body) => {
        try {
          await applyAsync((doc) => doc.setShapeText(slideIndex, shapeId, body));
          return { ok: true, message: `Слайд ${slideIndex + 1}: текст фигуры обновлён` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      formatText: async (slideIndex, shapeId, format) => {
        try {
          const doc = docRef.current;
          const shape = doc?.model.slides[slideIndex]?.shapes.find((item) => item.id === shapeId);
          if (!doc || !shape?.text) return { ok: false, message: 'Фигура не найдена' };
          const body = {
            ...shape.text,
            paragraphs: shape.text.paragraphs.map((paragraph) => ({
              ...paragraph,
              align: format.align ?? paragraph.align,
              level: format.level ?? paragraph.level,
              bullet: format.bullet ?? paragraph.bullet,
              lineSpacing: format.lineSpacing ?? paragraph.lineSpacing,
              runs: paragraph.runs.map((run) => ({
                ...run,
                bold: format.bold ?? run.bold,
                italic: format.italic ?? run.italic,
                underline: format.underline ?? run.underline,
                strike: format.strike ?? run.strike,
                size: format.size ?? run.size,
                color: format.color ?? run.color,
                family: format.font ?? run.family,
              })),
            })),
          };
          await applyAsync((current) => current.setShapeText(slideIndex, shapeId, body));
          return { ok: true, message: `Слайд ${slideIndex + 1}: формат текста применён` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      findReplace: async (query, replacement, options) => {
        try {
          let count = 0;
          await applyAsync((doc) => {
            count = doc.findReplace(query, replacement, options ?? {}).count;
          });
          return { ok: true, message: `Заменено вхождений: ${count}`, count };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },

      /* ---------- Фигуры ---------- */
      addShape: async (slideIndex, options) => {
        try {
          let shapeId: string | null = null;
          await applyAsync((doc) => {
            const geometry = geometryInEmu(options.geometry);
            shapeId = doc.addShape(slideIndex, {
              geom: options.geom,
              x: geometry?.x,
              y: geometry?.y,
              w: geometry?.w,
              h: geometry?.h,
              fill: options.fill,
              outline: options.outline,
              shadow: options.shadow,
              text: options.text,
            });
            if (shapeId) setSelectedId(shapeId);
          });
          return { ok: true, message: 'Фигура добавлена', shapeId: shapeId ?? undefined };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      setShapeGeometry: async (slideIndex, shapeId, geometry) => {
        try {
          await applyAsync((doc) => {
            const shape = doc.model.slides[slideIndex]?.shapes.find((item) => item.id === shapeId);
            if (!shape) throw new Error('Фигура не найдена');
            const next = geometryInEmu(geometry) ?? {};
            doc.setShapeGeometry(slideIndex, shapeId, {
              x: next.x ?? shape.x,
              y: next.y ?? shape.y,
              w: next.w ?? shape.w,
              h: next.h ?? shape.h,
              rotation: geometry.rotation ?? shape.rotation,
            });
          });
          return { ok: true, message: `Фигура ${shapeId}: геометрия обновлена` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      deleteShape: async (slideIndex, shapeId) => {
        try {
          await applyAsync((doc) => doc.deleteShape(slideIndex, shapeId));
          return { ok: true, message: `Фигура ${shapeId} удалена` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      duplicateShape: async (slideIndex, shapeId) => {
        try {
          let created: string | null = null;
          await applyAsync((doc) => {
            created = doc.duplicateShape(slideIndex, shapeId);
            if (created) setSelectedId(created);
          });
          return { ok: true, message: `Фигура ${shapeId} продублирована`, shapeId: created ?? undefined };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      reorderShape: async (slideIndex, shapeId, mode) => {
        try {
          await applyAsync((doc) => doc.reorderShape(slideIndex, shapeId, mode));
          return { ok: true, message: `Фигура ${shapeId}: порядок ${mode}` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      setShapeStyle: async (slideIndex, shapeId, style) => {
        try {
          await applyAsync((doc) => {
            if (style.fill !== undefined) doc.setShapeFill(slideIndex, shapeId, style.fill);
            if (style.outline !== undefined) doc.setShapeOutline(slideIndex, shapeId, style.outline);
            if (style.shadow !== undefined) doc.setShapeShadow(slideIndex, shapeId, style.shadow);
          });
          return { ok: true, message: `Фигура ${shapeId}: стиль обновлён` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },

      /* ---------- Таблицы, тема, показ ---------- */
      addTable: async (slideIndex, options) => {
        try {
          let shapeId: string | null = null;
          await applyAsync((doc) => {
            const geometry = geometryInEmu(options.geometry);
            shapeId = doc.addTable(slideIndex, {
              rows: options.rows,
              cols: options.cols,
              data: options.data,
              x: geometry?.x,
              y: geometry?.y,
              w: geometry?.w,
              h: geometry?.h,
            });
            if (shapeId) setSelectedId(shapeId);
          });
          return {
            ok: true,
            message: `Таблица ${options.rows}×${options.cols} добавлена`,
            shapeId: shapeId ?? undefined,
          };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      setTableData: async (slideIndex, shapeId, data) => {
        try {
          await applyAsync((doc) => doc.setTableData(slideIndex, shapeId, data));
          return { ok: true, message: `Таблица ${shapeId}: данные обновлены` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      applyTheme: async (options) => {
        try {
          await applyAsync((doc) => doc.applyTheme(options));
          return { ok: true, message: `Тема применена${options.presetId ? `: ${options.presetId}` : ''}` };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      present: async (index) => {
        if (typeof index === 'number') setActive(Math.max(0, index));
        setPresent(true);
        return { ok: true, message: 'Показ слайдов запущен' };
      },
      exportPdf: async () => {
        await exportPdf();
        return { ok: true, message: 'Экспорт PDF запущен' };
      },
      exportSlideImages: async (folder) => {
        try {
          const doc = docRef.current;
          const api = window.electronAPI;
          if (!doc || !api?.exportSlideImages) {
            return { ok: false, message: 'Экспорт PNG не поддержан в этой сборке' };
          }
          const pages = doc.model.slides.map((_, index) => buildPdfHtml(doc.model, index));
          const width = Math.round(doc.model.width / 9525);
          const height = Math.round(doc.model.height / 9525);
          const result = await api.exportSlideImages(pages, folder ?? '', width, height);
          if (!result.success) return { ok: false, message: result.error || 'Не удалось сохранить PNG' };
          return {
            ok: true,
            message: `Слайдов сохранено: ${result.files?.length ?? 0}${result.dir ? ` → ${result.dir}` : ''}`,
            files: result.files,
          };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
    });

    /** Координаты инструментов: EMU по умолчанию, но можно передать сантиметры. */
    function geometryInEmu(geometry?: {
      x?: number;
      y?: number;
      w?: number;
      h?: number;
      unit?: 'emu' | 'cm';
    }): { x?: number; y?: number; w?: number; h?: number } | undefined {
      if (!geometry) return undefined;
      const scale = geometry.unit === 'cm' ? 360000 : 1;
      return {
        x: geometry.x !== undefined ? Math.round(geometry.x * scale) : undefined,
        y: geometry.y !== undefined ? Math.round(geometry.y * scale) : undefined,
        w: geometry.w !== undefined ? Math.round(geometry.w * scale) : undefined,
        h: geometry.h !== undefined ? Math.round(geometry.h * scale) : undefined,
      };
    }

    // Общий помощник для контроллера: изменение + синхронизация.
    function applyAsync(action: (doc: PptxDocument) => void): Promise<void> {
      return new Promise((resolve, reject) => {
        const doc = docRef.current;
        if (!doc) {
          reject(new Error('Документ не открыт'));
          return;
        }
        try {
          action(doc);
          setModel({ ...doc.model });
          setActive((index) => Math.min(index, Math.max(0, doc.model.slides.length - 1)));
          scheduleSync();
          resolve();
        } catch (err) {
          reject(err);
        }
      });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [active, scheduleSync, selectedId, syncNow, tab.id]);

  /* ---------------- Рендер ---------------- */

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

  if (!model) {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-2 bg-ws-editor text-xs text-ws-muted">
        <AlertTriangle size={18} className="text-ws-warning" />
        {error || 'Презентация не открыта'}
      </div>
    );
  }

  const currentSlide = model.slides[present ? presentIndex : active] ?? model.slides[0];
  const transition = currentSlide?.transition;

  if (present && currentSlide) {
    return (
      <div className="fixed inset-0 z-[100] flex flex-col bg-black" onClick={() => setPresent(false)}>
        <div className="flex flex-1 items-center justify-center p-6">
          <div
            className="w-full max-w-6xl overflow-hidden rounded shadow-2xl"
            style={{
              opacity: transitionIn ? 1 : 0,
              transform:
                transition?.type === 'push' && !transitionIn ? 'translateX(6%)' : 'translateX(0)',
              transition: 'opacity .3s ease, transform .3s ease',
            }}
            onClick={(event) => event.stopPropagation()}
          >
            <SlideCanvas slide={currentSlide} model={model} width={1280} />
          </div>
        </div>
        <div className="flex items-center justify-between px-5 py-3 text-xs text-white/70">
          <span className="flex items-center gap-2">
            <Presentation size={13} />
            {presentIndex + 1} / {model.slides.length}
            {transition && transition.type !== 'none' && <span className="text-white/40">· {transition.type}</span>}
          </span>
          <div className="flex items-center gap-2">
            <button
              className="rounded p-1.5 hover:bg-white/15 disabled:opacity-30"
              disabled={presentIndex === 0}
              onClick={(event) => {
                event.stopPropagation();
                setPresentIndex((index) => Math.max(0, index - 1));
                setTransitionIn(0);
              }}
            >
              <ChevronLeft size={16} />
            </button>
            <button
              className="rounded p-1.5 hover:bg-white/15 disabled:opacity-30"
              disabled={presentIndex >= model.slides.length - 1}
              onClick={(event) => {
                event.stopPropagation();
                setPresentIndex((index) => Math.min(model.slides.length - 1, index + 1));
                setTransitionIn(0);
              }}
            >
              <ChevronRight size={16} />
            </button>
            <button
              className="flex items-center gap-1.5 rounded bg-white/10 px-2.5 py-1.5 hover:bg-white/20"
              onClick={(event) => {
                event.stopPropagation();
                setPresent(false);
              }}
            >
              <X size={13} />
              Выйти (Esc)
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-full w-full flex-col bg-ws-editor text-ws">
      {/* Верхняя лента */}
      <div className="flex shrink-0 flex-wrap items-center gap-1 border-b border-ws bg-ws-surface px-2 py-1 text-[11px]">
        <span className="mr-1 flex items-center gap-1.5 rounded border border-orange-500/40 bg-orange-500/10 px-2 py-0.5 font-medium text-orange-500">
          <Presentation size={12} />
          PowerPoint
        </span>

        <ToolButton title="Новый слайд (Alt+Enter)" onClick={addSlide}>
          <Plus size={13} /> Слайд
        </ToolButton>
        <ToolButton title="Дублировать слайд" onClick={duplicateSlide}>
          <Copy size={12} />
        </ToolButton>
        <ToolButton title="Удалить слайд" onClick={deleteSlide} disabled={model.slides.length <= 1} danger>
          <Trash2 size={12} />
        </ToolButton>
        <ToolButton title="Выше" onClick={() => moveSlide(-1)} disabled={active <= 0}>
          <ArrowUp size={12} />
        </ToolButton>
        <ToolButton title="Ниже" onClick={() => moveSlide(1)} disabled={active >= model.slides.length - 1}>
          <ArrowDown size={12} />
        </ToolButton>
        <ToolButton title={slide?.hidden ? 'Показать слайд' : 'Скрыть слайд'} onClick={toggleHidden}>
          {slide?.hidden ? <EyeOff size={12} /> : <Eye size={12} />}
        </ToolButton>

        <Divider />

        <ToolButton title="Надпись" onClick={addTextBox}>
          <Type size={13} /> Текст
        </ToolButton>
        <div className="relative">
          <ToolButton title="Фигура" onClick={() => setShapeMenu((open) => !open)}>
            <Square size={12} /> Фигура
          </ToolButton>
          {shapeMenu && (
            <div className="absolute left-0 top-full z-40 mt-1 w-40 rounded border border-ws bg-ws-elevated py-1 shadow-lg">
              {SHAPES.map((shape) => (
                <button
                  key={shape.id}
                  onClick={() => addShape(shape.id)}
                  className="block w-full px-2.5 py-1 text-left text-[11px] text-ws hover:bg-ws-hover"
                >
                  {shape.label}
                </button>
              ))}
            </div>
          )}
        </div>
        <ToolButton title="Изображение" onClick={() => imageInputRef.current?.click()}>
          <ImageIcon size={12} />
        </ToolButton>
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
        <ToolButton title="Таблица 3×3" onClick={addTable}>
          <Table2 size={12} />
        </ToolButton>

        <Divider />

        <ToolButton title="Отменить (Ctrl+Z)" onClick={undo}>
          <Undo2 size={13} />
        </ToolButton>
        <ToolButton title="Повторить (Ctrl+Y)" onClick={redo}>
          <Redo2 size={13} />
        </ToolButton>

        <Divider />

        <ToolButton title="Тема оформления" onClick={() => setThemeOpen(true)}>
          <Palette size={13} /> Тема
        </ToolButton>
        <ToolButton title="Найти и заменить (Ctrl+F)" onClick={() => setFindOpen(true)}>
          <Search size={12} />
        </ToolButton>

        <select
          value={slide?.layoutPart ?? ''}
          onChange={(event) => changeLayout(event.target.value)}
          title="Макет слайда"
          className="ml-1 max-w-[9rem] rounded border border-ws bg-ws-elevated px-1.5 py-1 text-[11px] text-ws"
        >
          {model.layouts.map((layout) => (
            <option key={layout.part} value={layout.part}>
              {layout.name}
            </option>
          ))}
        </select>
        <select
          value={
            SLIDE_SIZES.find((size) => size.width === model.width && size.height === model.height)?.id ?? ''
          }
          onChange={(event) => changeSlideSize(event.target.value)}
          title="Размер слайда"
          className="rounded border border-ws bg-ws-elevated px-1.5 py-1 text-[11px] text-ws"
        >
          {SLIDE_SIZES.map((size) => (
            <option key={size.id} value={size.id}>
              {size.label}
            </option>
          ))}
        </select>

        <div className="ml-auto flex items-center gap-1">
          <ToolButton title="Экспорт в PDF" onClick={() => void exportPdf()}>
            <FileDown size={12} /> PDF
          </ToolButton>
          <ToolButton title="Показ слайдов (F5)" onClick={() => setPresent(true)} accent>
            <Maximize2 size={12} /> Показ
          </ToolButton>
          <ToolButton title="Сохранить (Ctrl+S)" onClick={() => void save()}>
            <Save size={12} /> Сохранить
          </ToolButton>
        </div>
      </div>

      {error && (
        <div className="flex shrink-0 items-center gap-2 border-b border-red-500/40 bg-red-500/10 px-3 py-1 text-[11px] text-ws-danger">
          <AlertTriangle size={11} />
          <span className="flex-1">{error}</span>
          <button className="rounded px-1 hover:bg-red-500/20" onClick={() => setError('')}>
            <X size={11} />
          </button>
        </div>
      )}
      {notice && (
        <div className="shrink-0 border-b border-ws-accent/30 bg-ws-accent-soft px-3 py-1 text-[11px] text-ws-accent">
          {notice}
        </div>
      )}

      <div className="flex min-h-0 flex-1">
        {/* Слайды */}
        <aside className="w-44 shrink-0 overflow-y-auto border-r border-ws bg-ws-surface p-2">
          <div className="mb-1.5 flex items-center gap-1.5 text-[10px] font-semibold uppercase tracking-wider text-ws-muted">
            <Layers size={11} /> Слайды ({model.slides.length})
          </div>
          {model.slides.map((item, index) => (
            <button
              key={item.part}
              onClick={() => {
                setActive(index);
                setSelectedId(null);
              }}
              onDoubleClick={() => {
                setActive(index);
                setPresent(true);
              }}
              className={cn(
                'mb-1.5 block w-full overflow-hidden rounded border text-left',
                index === active ? 'border-ws-accent ring-1 ring-ws-accent/40' : 'border-ws hover:border-ws-strong'
              )}
            >
              <div className="flex items-center justify-between bg-ws-elevated px-1.5 py-0.5 text-[9px] text-ws-faint">
                <span>{index + 1}</span>
                {item.hidden && <EyeOff size={9} />}
                {item.transition && item.transition.type !== 'none' && <span>▶</span>}
              </div>
              <div className="pointer-events-none">
                <SlideCanvas slide={item} model={model} width={150} />
              </div>
            </button>
          ))}
        </aside>

        {/* Канва */}
        <div ref={wrapRef} className="relative min-h-0 flex-1 overflow-auto bg-ws-editor-alt p-4">
          <div className="relative mx-auto" style={{ width: canvasWidth }} data-slide-canvas>
            {currentSlide ? (
              <div className="rounded border border-ws shadow-ws">
                <SlideCanvas
                  slide={currentSlide}
                  model={model}
                  width={canvasWidth}
                  selectedId={selectedId}
                  onShapeMouseDown={beginMove}
                  onShapeDoubleClick={beginTextEdit}
                  geometryOverride={override}
                />
              </div>
            ) : (
              <div className="flex h-48 items-center justify-center text-xs text-ws-faint">
                Слайдов нет — добавьте первый
              </div>
            )}

            {/* Рамка выделения с маркерами */}
            {selectedShape && currentSlide && (
              <div
                className="pointer-events-none absolute border-2 border-blue-500/80"
                style={{
                  left: `${(((override[selectedShape.id] ?? selectedShape).x) / model.width) * 100}%`,
                  top: `${(((override[selectedShape.id] ?? selectedShape).y) / model.height) * 100}%`,
                  width: `${(((override[selectedShape.id] ?? selectedShape).w) / model.width) * 100}%`,
                  height: `${(((override[selectedShape.id] ?? selectedShape).h) / model.height) * 100}%`,
                }}
              >
                {(['nw', 'ne', 'sw', 'se'] as const).map((handle) => (
                  <div
                    key={handle}
                    onPointerDown={(event) => beginResize(handle, event)}
                    className={cn(
                      'pointer-events-auto absolute h-2.5 w-2.5 rounded-sm border border-white bg-blue-500',
                      handle === 'nw' && '-left-1.5 -top-1.5 cursor-nwse-resize',
                      handle === 'ne' && '-right-1.5 -top-1.5 cursor-nesw-resize',
                      handle === 'sw' && '-left-1.5 -bottom-1.5 cursor-nesw-resize',
                      handle === 'se' && '-right-1.5 -bottom-1.5 cursor-nwse-resize'
                    )}
                  />
                ))}
              </div>
            )}

            {/* Инлайн-правка текста */}
            {editingText && selectedShape && currentSlide && (
              <textarea
                autoFocus
                value={editingText.value}
                onChange={(event) => setEditingText({ ...editingText, value: event.target.value })}
                onBlur={commitTextEdit}
                onKeyDown={(event) => {
                  if (event.key === 'Escape') setEditingText(null);
                  if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) commitTextEdit();
                }}
                spellCheck={false}
                className="absolute z-30 resize-none overflow-hidden rounded border-2 border-blue-500 bg-white/95 p-1 text-slate-900 outline-none"
                style={{
                  left: `${(selectedShape.x / model.width) * 100}%`,
                  top: `${(selectedShape.y / model.height) * 100}%`,
                  width: `${(selectedShape.w / model.width) * 100}%`,
                  height: `${(selectedShape.h / model.height) * 100}%`,
                  fontSize: `${Math.max(
                    11,
                    (selectedShape.text?.paragraphs[0]?.runs[0]?.size ?? 18) *
                      (canvasWidth / (model.width / 12700))
                  )}px`,
                }}
              />
            )}

            {/* Зум */}
            <div className="absolute -bottom-1 right-0 flex translate-y-full items-center gap-1 pb-1 pt-2 text-[10px] text-ws-muted">
              <button
                className="rounded border border-ws bg-ws-surface p-1 hover:bg-ws-hover"
                onClick={() => setZoom((value) => Math.max(0.4, Math.round((value - 0.1) * 10) / 10))}
                title="Отдалить"
              >
                <ZoomOut size={11} />
              </button>
              <span className="w-10 text-center font-mono">{Math.round(zoom * 100)}%</span>
              <button
                className="rounded border border-ws bg-ws-surface p-1 hover:bg-ws-hover"
                onClick={() => setZoom((value) => Math.min(2.5, Math.round((value + 0.1) * 10) / 10))}
                title="Приблизить"
              >
                <ZoomIn size={11} />
              </button>
              <button
                className="ml-1 rounded border border-ws bg-ws-surface px-1.5 py-1 hover:bg-ws-hover"
                onClick={() => setZoom(1)}
              >
                <Minus size={11} />
              </button>
            </div>
          </div>
        </div>

        <Inspector
          shape={selectedShape}
          slide={currentSlide ?? model.slides[0]}
          model={model}
          onGeometry={setGeometry}
          onFill={(color) => {
            if (!selectedShape) return;
            apply('Заливка', (doc) => doc.setShapeFill(active, selectedShape.id, color));
          }}
          onOutline={(outline) => {
            if (!selectedShape) return;
            apply('Контур', (doc) => doc.setShapeOutline(active, selectedShape.id, outline));
          }}
          onShadow={(enabled) => {
            if (!selectedShape) return;
            apply('Тень', (doc) => doc.setShapeShadow(active, selectedShape.id, enabled));
          }}
          onTextFormat={formatText}
          onParagraph={formatParagraph}
          onTextBody={setTextBody}
          onZOrder={(mode) => {
            if (!selectedShape) return;
            apply('Порядок фигур', (doc) => doc.reorderShape(active, selectedShape.id, mode));
          }}
          onDeleteShape={deleteShape}
          onDuplicateShape={duplicateShape}
          onBackground={setBackground}
          onTransition={setTransition}
        />
      </div>

      {/* Заметки и статус */}
      <div className="flex shrink-0 items-start gap-3 border-t border-ws bg-ws-surface px-3 py-2">
        <div className="min-w-0 flex-1">
          <div className="mb-1 flex items-center gap-1.5 text-[10px] font-semibold uppercase tracking-wider text-ws-muted">
            Заметки докладчика
            {currentSlide?.notes && <span className="font-normal lowercase text-ws-faint">есть</span>}
          </div>
          <textarea
            value={currentSlide?.notes ?? ''}
            onChange={(event) => {
              const value = event.target.value;
              apply('Заметки', (doc) => doc.setNotes(active, value));
            }}
            placeholder="Заметки к слайду"
            spellCheck={false}
            className="h-14 w-full resize-y rounded border border-ws bg-ws-editor px-2 py-1 text-[11px] text-ws outline-none focus:border-ws-accent"
          />
        </div>
        <div className="w-56 shrink-0 text-right text-[10px] text-ws-faint">
          <div>
            Слайд {model.slides.length ? active + 1 : 0} из {model.slides.length}
          </div>
          <div>Тема: {model.theme.name}</div>
          <div>
            {model.slides[active]?.shapes.length ?? 0} фигур · {EMU_PER_CM > 0 ? '16:9' : ''}{' '}
            {Math.round((model.width / 360000) * 10) / 10}×{Math.round((model.height / 360000) * 10) / 10} см
          </div>
          <div className="truncate">{tab.name}</div>
        </div>
      </div>

      <ThemePanel
        open={themeOpen}
        theme={model.theme}
        shadowDefault={shadowDefault}
        onClose={() => setThemeOpen(false)}
        onApplyPreset={applyThemePreset}
        onApplyColors={applyThemeColors}
        onApplyFonts={applyThemeFonts}
        onApplyEffects={applyEffects}
      />

      {findOpen && (
        <div className="fixed inset-0 z-[90] flex items-start justify-center bg-black/40 pt-24" onClick={() => setFindOpen(false)}>
          <div
            className="w-full max-w-md rounded-lg border border-ws bg-ws-surface p-3 shadow-xl"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="mb-2 text-xs font-medium text-ws">Найти и заменить</div>
            <input
              autoFocus
              value={findQuery}
              onChange={(event) => setFindQuery(event.target.value)}
              placeholder="Найти"
              className="mb-2 w-full rounded border border-ws bg-ws-editor px-2 py-1 text-xs text-ws outline-none focus:border-ws-accent"
            />
            <input
              value={replaceQuery}
              onChange={(event) => setReplaceQuery(event.target.value)}
              placeholder="Заменить на"
              className="mb-3 w-full rounded border border-ws bg-ws-editor px-2 py-1 text-xs text-ws outline-none focus:border-ws-accent"
            />
            <div className="flex justify-end gap-2">
              <button
                className="rounded border border-ws px-2 py-1 text-[11px] text-ws-muted hover:bg-ws-hover"
                onClick={() => setFindOpen(false)}
              >
                Отмена
              </button>
              <button
                className="rounded bg-ws-accent px-2 py-1 text-[11px] font-medium text-white hover:opacity-90"
                onClick={runFindReplace}
              >
                Заменить всё
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

/* ============================================================
   Мелкие компоненты
   ============================================================ */

const ToolButton: React.FC<{
  title: string;
  onClick: () => void;
  children: React.ReactNode;
  disabled?: boolean;
  danger?: boolean;
  accent?: boolean;
}> = ({ title, onClick, children, disabled, danger, accent }) => (
  <button
    title={title}
    onClick={onClick}
    disabled={disabled}
    className={cn(
      'flex items-center gap-1 rounded px-1.5 py-1 text-ws-muted hover:bg-ws-hover hover:text-ws disabled:opacity-40',
      danger && 'hover:text-ws-danger',
      accent && 'bg-ws-accent text-white hover:bg-ws-accent hover:text-white'
    )}
  >
    {children}
  </button>
);

const Divider: React.FC = () => <span className="mx-0.5 h-3.5 w-px bg-ws" />;

/* ============================================================
   Экспорт в PDF: слайды как DOM-страницы
   ============================================================ */

function buildPdfHtml(model: PptxDocModel, only?: number): string {
  const widthIn = model.width / EMU_PER_INCH;
  const heightIn = model.height / EMU_PER_INCH;
  const px = (emu: number) => `${(emu / EMU_PER_INCH).toFixed(4)}in`;

  const slides = model.slides
    .filter((_, index) => only === undefined || index === only)
    .map((slide) => {
      const shapes = slide.shapes
        .map((shape) => {
          const box = `position:absolute;left:${px(shape.x)};top:${px(shape.y)};width:${px(
            shape.w
          )};height:${px(shape.h)};overflow:hidden;`;
          const rotate = shape.rotation ? `transform:rotate(${shape.rotation}deg);` : '';
          if (shape.kind === 'image' && shape.image?.src) {
            return `<img src="${shape.image.src}" style="${box}${rotate}object-fit:contain">`;
          }
          if (shape.kind === 'table' && shape.table) {
            const rows = shape.table.rows
              .map(
                (row) =>
                  '<tr>' +
                  row
                    .map(
                      (cell) =>
                        `<td style="border:1px solid #cbd5e1;padding:2pt;vertical-align:middle;background:${
                          cell.fill ? '#' + cell.fill : '#ffffff'
                        }">${textBodyHtml(cell.text)}</td>`
                    )
                    .join('') +
                  '</tr>'
              )
              .join('');
            return `<div style="${box}${rotate}"><table style="width:100%;height:100%;border-collapse:collapse;table-layout:fixed">${rows}</table></div>`;
          }
          const background = shape.fill ? `background:#${shape.fill};` : '';
          const border = shape.outline?.color
            ? `border:${(shape.outline.width ?? 1).toFixed(2)}pt solid #${shape.outline.color};`
            : '';
          return `<div style="${box}${rotate}${background}${border}padding:2pt">${
            shape.text ? textBodyHtml(shape.text) : ''
          }</div>`;
        })
        .join('');

      const background = slide.backgroundImage
        ? `background:#000 url(${slide.backgroundImage}) center/cover no-repeat;`
        : `background:#${slide.background};`;

      return `<div class="slide" style="${background}">${shapes}</div>`;
    })
    .join('');

  return (
    `<!doctype html><html><head><meta charset="utf-8"><style>` +
    `@page{size:${widthIn.toFixed(4)}in ${heightIn.toFixed(4)}in;margin:0}` +
    `html,body{margin:0;padding:0}` +
    `.slide{position:relative;width:${widthIn.toFixed(4)}in;height:${heightIn.toFixed(
      4
    )}in;page-break-after:always;break-after:page;overflow:hidden}` +
    `.slide:last-child{page-break-after:auto;break-after:auto}` +
    `</style></head><body>` +
    slides +
    `</body></html>`
  );
}

function textBodyHtml(body: PptxTextBody): string {
  return body.paragraphs
    .map((paragraph) => {
      const align =
        paragraph.align === 'ctr' ? 'center' : paragraph.align === 'r' ? 'right' : paragraph.align === 'just' ? 'justify' : 'left';
      const runs = paragraph.runs
        .map((run) => {
          const style = [
            run.size ? `font-size:${run.size}pt` : '',
            run.color ? `color:#${run.color}` : '',
            run.bold ? 'font-weight:700' : '',
            run.italic ? 'font-style:italic' : '',
            run.underline ? 'text-decoration:underline' : '',
            run.family ? `font-family:'${run.family}','Segoe UI',sans-serif` : '',
          ]
            .filter(Boolean)
            .join(';');
          const text = run.text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
          return `<span style="${style}">${text}</span>`;
        })
        .join('');
      const bullet =
        paragraph.bullet === 'bullet' ? '• ' : paragraph.bullet === 'number' ? '1. ' : '';
      return `<div style="text-align:${align};margin:0;line-height:${
        paragraph.lineSpacing ? paragraph.lineSpacing / 100 : 1.2
      }">${bullet}${runs || '&nbsp;'}</div>`;
    })
    .join('');
}

