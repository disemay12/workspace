import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import JSZip from 'jszip';
import { OpenTab } from '../../types';
import {
  Presentation,
  ChevronLeft,
  ChevronRight,
  Maximize2,
  Minimize2,
  Sparkles,
  Layers,
  AlertTriangle,
  Grid3x3,
  Type,
  Image as ImageIcon,
  Save,
  Play,
} from 'lucide-react';
import { cn } from '../../lib/utils';
import { Spinner } from '../ui';

interface Props {
  tab: OpenTab;
  onChange?: (newContent: string) => void;
  onSave?: () => void;
  onAskAI?: (prompt: string, contextText: string) => void;
}

/** Текстовый блок слайда с приблизительной геометрией (EMU → проценты). */
interface SlideText {
  text: string;
  /** Позиция и размер в процентах от размера слайда. */
  x: number;
  y: number;
  w: number;
  h: number;
  /** Размер шрифта в pt, если удалось определить. */
  fontSize?: number;
  bold?: boolean;
  /** Является ли блок заголовком (по эвристике размера/позиции). */
  isTitle: boolean;
}

interface SlideShape {
  type: 'image' | 'table';
  x: number;
  y: number;
  w: number;
  h: number;
  src?: string;
}

interface ParsedSlide {
  index: number;
  texts: SlideText[];
  shapes: SlideShape[];
  notes: string;
}

/** Стандартный размер слайда PowerPoint 16:9 в EMU. */
const SLIDE_W = 12192000;
const SLIDE_H = 6858000;

const pct = (value: number, total: number) => (value / total) * 100;

export const PowerPointViewer: React.FC<Props> = ({ tab, onAskAI }) => {
  const [slides, setSlides] = useState<ParsedSlide[]>([]);
  const [current, setCurrent] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [showGrid, setShowGrid] = useState(false);
  const [slideSize, setSlideSize] = useState({ w: SLIDE_W, h: SLIDE_H });

  const containerRef = useRef<HTMLDivElement>(null);

  /* ---------------- Разбор PPTX ---------------- */

  useEffect(() => {
    let cancelled = false;

    const parse = async () => {
      try {
        setLoading(true);
        setError('');

        const binary = atob(tab.content);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

        const zip = await JSZip.loadAsync(bytes.buffer);
        const parser = new DOMParser();

        /* --- Размер слайда из presentation.xml --- */
        let size = { w: SLIDE_W, h: SLIDE_H };
        const presentationFile = zip.file('ppt/presentation.xml');
        if (presentationFile) {
          const xml = parser.parseFromString(await presentationFile.async('text'), 'text/xml');
          const sldSz = xml.getElementsByTagName('p:sldSz')[0];
          if (sldSz) {
            const cx = parseInt(sldSz.getAttribute('cx') ?? '', 10);
            const cy = parseInt(sldSz.getAttribute('cy') ?? '', 10);
            if (cx > 0 && cy > 0) size = { w: cx, h: cy };
          }
        }
        if (cancelled) return;
        setSlideSize(size);

        /* --- Собираем карту изображений rId → dataURL --- */
        const imageMap = new Map<string, string>();
        const slideFiles = Object.keys(zip.files).filter((p) =>
          /^ppt\/slides\/slide\d+\.xml$/.test(p)
        );

        // Карта связей для каждого слайда
        const relsMap = new Map<string, Map<string, string>>();
        for (const slidePath of slideFiles) {
          const relsPath = slidePath.replace('slides/', 'slides/_rels/') + '.rels';
          const relsFile = zip.file(relsPath);
          const map = new Map<string, string>();
          if (relsFile) {
            const relsXml = parser.parseFromString(await relsFile.async('text'), 'text/xml');
            const rels = relsXml.getElementsByTagName('Relationship');
            for (let i = 0; i < rels.length; i++) {
              const id = rels[i].getAttribute('Id');
              const target = rels[i].getAttribute('Target');
              if (id && target) map.set(id, target);
            }
          }
          relsMap.set(slidePath, map);
        }

        /* --- Загрузка изображений --- */
        for (const [, map] of relsMap) {
          for (const [, target] of map) {
            if (!/\.(png|jpe?g|gif|bmp|webp)$/i.test(target)) continue;
            const normalized = target.startsWith('/')
              ? target.slice(1)
              : `ppt/${target.replace(/^\.\.\//, '')}`;
            if (imageMap.has(normalized)) continue;
            const file = zip.file(normalized);
            if (!file) continue;
            const b64 = await file.async('base64');
            const ext = normalized.split('.').pop()!.toLowerCase().replace('jpg', 'jpeg');
            imageMap.set(normalized, `data:image/${ext};base64,${b64}`);
          }
        }

        /* --- Разбор каждого слайда --- */
        const sorted = slideFiles.sort((a, b) => {
          const na = parseInt(a.replace(/\D/g, ''), 10);
          const nb = parseInt(b.replace(/\D/g, ''), 10);
          return na - nb;
        });

        const parsed: ParsedSlide[] = [];
        let maxFontSize = 0;

        for (let i = 0; i < sorted.length; i++) {
          const path = sorted[i];
          const xml = parser.parseFromString(await zip.file(path)!.async('text'), 'text/xml');
          const rels = relsMap.get(path) ?? new Map<string, string>();

          const texts: SlideText[] = [];
          const shapes: SlideShape[] = [];
          const spNodes = xml.getElementsByTagName('p:sp');

          for (let s = 0; s < spNodes.length; s++) {
            const sp = spNodes[s];

            // Геометрия блока
            const off = sp.getElementsByTagName('a:off')[0];
            const ext = sp.getElementsByTagName('a:ext')[0];
            const x = off ? pct(parseInt(off.getAttribute('x') ?? '0', 10), size.w) : 6;
            const y = off ? pct(parseInt(off.getAttribute('y') ?? '0', 10), size.h) : 6;
            const w = ext ? pct(parseInt(ext.getAttribute('cx') ?? '0', 10), size.w) : 88;
            const h = ext ? pct(parseInt(ext.getAttribute('cy') ?? '0', 10), size.h) : 12;

            // Текстовые параграфы
            const paragraphs = sp.getElementsByTagName('a:p');
            const lines: string[] = [];
            let fontSize: number | undefined;
            let bold = false;

            for (let p = 0; p < paragraphs.length; p++) {
              const runs = paragraphs[p].getElementsByTagName('a:r');
              let line = '';
              for (let r = 0; r < runs.length; r++) {
                const t = runs[r].getElementsByTagName('a:t')[0];
                if (t?.textContent) line += t.textContent;

                // Размер шрифта — берём максимальный в блоке
                const rPr = runs[r].getElementsByTagName('a:rPr')[0];
                if (rPr) {
                  const sz = parseInt(rPr.getAttribute('sz') ?? '0', 10);
                  if (sz > 0) fontSize = Math.max(fontSize ?? 0, sz / 100);
                  if (rPr.getAttribute('b') === '1') bold = true;
                }
              }
              if (line.trim()) lines.push(line.trim());
            }

            const text = lines.join('\n');
            if (text) {
              maxFontSize = Math.max(maxFontSize, fontSize ?? 0);
              texts.push({ text, x, y, w, h, fontSize, bold, isTitle: false });
            }

            // Встроенные изображения
            const blips = sp.getElementsByTagName('a:blip');
            for (let b = 0; b < blips.length; b++) {
              const embed = blips[b].getAttribute('r:embed');
              if (!embed) continue;
              const target = rels.get(embed);
              if (!target) continue;
              const key = target.startsWith('/')
                ? target.slice(1)
                : `ppt/${target.replace(/^\.\.\//, '')}`;
              const src = imageMap.get(key);
              if (src) shapes.push({ type: 'image', x, y, w, h, src });
            }
          }

          // Отдельные картинки вне фигур (p:pic)
          const picNodes = xml.getElementsByTagName('p:pic');
          for (let p = 0; p < picNodes.length; p++) {
            const pic = picNodes[p];
            const off = pic.getElementsByTagName('a:off')[0];
            const ext = pic.getElementsByTagName('a:ext')[0];
            const blip = pic.getElementsByTagName('a:blip')[0];
            if (!blip) continue;
            const target = rels.get(blip.getAttribute('r:embed') ?? '');
            if (!target) continue;
            const key = target.startsWith('/')
              ? target.slice(1)
              : `ppt/${target.replace(/^\.\.\//, '')}`;
            const src = imageMap.get(key);
            if (!src) continue;
            shapes.push({
              type: 'image',
              x: off ? pct(parseInt(off.getAttribute('x') ?? '0', 10), size.w) : 10,
              y: off ? pct(parseInt(off.getAttribute('y') ?? '0', 10), size.h) : 10,
              w: ext ? pct(parseInt(ext.getAttribute('cx') ?? '0', 10), size.w) : 40,
              h: ext ? pct(parseInt(ext.getAttribute('cy') ?? '0', 10), size.h) : 40,
              src,
            });
          }

          // Заметки докладчика
          let notes = '';
          const notesFile = zip.file(`ppt/notesSlides/notesSlide${i + 1}.xml`);
          if (notesFile) {
            const notesXml = parser.parseFromString(await notesFile.async('text'), 'text/xml');
            const noteTexts = notesXml.getElementsByTagName('a:t');
            const buf: string[] = [];
            for (let n = 0; n < noteTexts.length; n++) {
              const t = noteTexts[n].textContent?.trim();
              if (t) buf.push(t);
            }
            notes = buf.join(' ');
          }

          parsed.push({ index: i + 1, texts, shapes, notes });
        }

        // Определяем заголовок: самый крупный шрифт в верхней части слайда
        for (const slide of parsed) {
          const candidates = slide.texts.filter((t) => t.y < 45);
          if (candidates.length > 0) {
            const top = candidates.reduce((best, t) =>
              (t.fontSize ?? 0) > (best.fontSize ?? 0) ? t : best
            );
            top.isTitle = true;
          } else if (slide.texts.length > 0) {
            slide.texts[0].isTitle = true;
          }
        }

        if (cancelled) return;

        if (parsed.length === 0) {
          parsed.push({
            index: 1,
            texts: [
              { text: 'Презентация пуста', x: 6, y: 40, w: 88, h: 15, isTitle: true },
              { text: 'Файл не содержит слайдов с текстом.', x: 6, y: 56, w: 88, h: 12, isTitle: false },
            ],
            shapes: [],
            notes: '',
          });
        }

        setSlides(parsed);
        setCurrent(0);
      } catch (err) {
        if (cancelled) return;
        setError(`Не удалось разобрать презентацию: ${(err as Error).message}`);
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    void parse();
    return () => {
      cancelled = true;
    };
  }, [tab.id, tab.content]);

  /* ---------------- Навигация ---------------- */

  const next = useCallback(() => setCurrent((c) => Math.min(slides.length - 1, c + 1)), [slides.length]);
  const prev = useCallback(() => setCurrent((c) => Math.max(0, c - 1)), []);

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (!isFullscreen) return;
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
        e.preventDefault();
        next();
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault();
        prev();
      } else if (e.key === 'Escape') {
        setIsFullscreen(false);
      } else if (e.key === 'Home') {
        setCurrent(0);
      } else if (e.key === 'End') {
        setCurrent(slides.length - 1);
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [isFullscreen, next, prev, slides.length]);

  const activeSlide = slides[current];

  const handleAISummarize = () => {
    if (!onAskAI) return;
    const allText = slides
      .map(
        (s) =>
          `[Слайд ${s.index}]\n` +
          s.texts.map((t) => t.text).join('\n') +
          (s.notes ? `\nЗаметки: ${s.notes}` : '')
      )
      .join('\n\n');
    onAskAI(
      'Проанализируй презентацию: сделай краткое резюме, выдели ключевые тезисы каждого слайда и подготовь заметки для докладчика.',
      allText
    );
  };

  const allText = useMemo(
    () => slides.flatMap((s) => s.texts.map((t) => t.text)).join(' '),
    [slides]
  );

  if (loading) {
    return (
      <div className="flex h-full items-center justify-center bg-ws-editor">
        <div className="flex flex-col items-center gap-2 text-ws-muted">
          <Presentation size={28} className="animate-pulse text-orange-500" />
          <span className="flex items-center gap-2 text-xs">
            <Spinner size={12} />
            Загрузка презентации…
          </span>
        </div>
      </div>
    );
  }

  /* ---------------- Полноэкранный показ ---------------- */
  if (isFullscreen) {
    return (
      <div
        className="fixed inset-0 z-[100] flex flex-col bg-black"
        onClick={(e) => {
          // Клик по правой половине — вперёд, по левой — назад
          const half = window.innerWidth / 2;
          if (e.clientX > half) next();
          else prev();
        }}
      >
        <div className="flex flex-1 items-center justify-center p-6">
          {activeSlide && (
            <div
              className="relative aspect-video w-full max-w-7xl overflow-hidden rounded bg-white"
              style={{ aspectRatio: `${slideSize.w} / ${slideSize.h}` }}
              onClick={(e) => e.stopPropagation()}
            >
              <SlideCanvas slide={activeSlide} fullscreen />
            </div>
          )}
        </div>

        <div className="flex items-center justify-between px-5 py-3 text-xs text-white/70">
          <span className="flex items-center gap-2">
            <Play size={12} />
            {current + 1} / {slides.length}
          </span>
          <div className="flex items-center gap-3">
            <button
              onClick={(e) => {
                e.stopPropagation();
                prev();
              }}
              disabled={current === 0}
              className="rounded p-1.5 transition-colors hover:bg-white/15 disabled:opacity-30"
            >
              <ChevronLeft size={16} />
            </button>
            <button
              onClick={(e) => {
                e.stopPropagation();
                next();
              }}
              disabled={current === slides.length - 1}
              className="rounded p-1.5 transition-colors hover:bg-white/15 disabled:opacity-30"
            >
              <ChevronRight size={16} />
            </button>
            <button
              onClick={(e) => {
                e.stopPropagation();
                setIsFullscreen(false);
              }}
              className="flex items-center gap-1.5 rounded bg-white/10 px-2.5 py-1.5 transition-colors hover:bg-white/20"
            >
              <Minimize2 size={13} />
              Выйти (Esc)
            </button>
          </div>
        </div>
      </div>
    );
  }

  /* ---------------- Обычный режим ---------------- */
  return (
    <div className="flex h-full w-full flex-col bg-ws-editor text-ws select-none">
      {/* Панель инструментов */}
      <div className="flex shrink-0 items-center justify-between border-b border-ws bg-ws-surface px-3 py-1.5 text-xs">
        <div className="flex items-center gap-2">
          <span className="flex items-center gap-1.5 rounded border border-orange-500/40 bg-orange-500/10 px-2 py-0.5 font-medium text-orange-500">
            <Presentation size={13} />
            Презентация
          </span>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <div className="flex items-center gap-1">
            <button
              onClick={prev}
              disabled={current === 0}
              title="Предыдущий слайд"
              className="rounded p-1 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws disabled:opacity-35"
            >
              <ChevronLeft size={15} />
            </button>
            <span className="px-1.5 font-mono text-ws">
              {current + 1} / {slides.length}
            </span>
            <button
              onClick={next}
              disabled={current === slides.length - 1}
              title="Следующий слайд"
              className="rounded p-1 text-ws-muted transition-colors hover:bg-ws-hover hover:text-ws disabled:opacity-35"
            >
              <ChevronRight size={15} />
            </button>
          </div>

          <div className="mx-1 h-3.5 w-px border-l border-ws" />

          <button
            onClick={() => setShowGrid((v) => !v)}
            title="Показать сетку слайдов"
            className={cn(
              'flex items-center gap-1 rounded border px-2 py-1 transition-colors',
              showGrid
                ? 'border-ws-accent bg-ws-accent-soft text-ws-accent'
                : 'border-ws text-ws-muted hover:text-ws'
            )}
          >
            <Grid3x3 size={11} />
            Сетка
          </button>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={handleAISummarize}
            disabled={!allText.trim()}
            className="flex items-center gap-1.5 rounded border border-orange-500/40 bg-orange-500/10 px-2.5 py-1 text-orange-500 transition-colors hover:bg-orange-500/20 disabled:opacity-50"
          >
            <Sparkles size={12} />
            Тезисы
          </button>

          <button
            onClick={() => setIsFullscreen(true)}
            className="flex items-center gap-1.5 rounded bg-ws-accent px-2.5 py-1 font-medium text-white transition-opacity hover:opacity-90"
          >
            <Maximize2 size={12} />
            Показ
          </button>
        </div>
      </div>

      {error && (
        <div className="flex shrink-0 items-center gap-2 border-b border-amber-500/40 bg-amber-500/10 px-3 py-2 text-[11px] text-ws-warning">
          <AlertTriangle size={12} />
          {error}
        </div>
      )}

      <div className="flex min-h-0 flex-1">
        {/* Миниатюры */}
        <div className="w-56 shrink-0 overflow-y-auto border-r border-ws bg-ws-panel p-3">
          <div className="mb-2 flex items-center gap-1.5 text-[10px] font-semibold tracking-wider text-ws-muted uppercase">
            <Layers size={12} />
            Слайды ({slides.length})
          </div>

          <div className="space-y-2">
            {slides.map((slide, idx) => (
              <div
                key={slide.index}
                onClick={() => setCurrent(idx)}
                className={cn(
                  'overflow-hidden rounded-lg border transition-all',
                  current === idx
                    ? 'border-orange-500 ring-1 ring-orange-500/50'
                    : 'border-ws hover:border-ws-strong'
                )}
              >
                <div className="flex items-center justify-between bg-ws-surface px-2 py-0.5 text-[9px] text-ws-faint">
                  <span>#{slide.index}</span>
                  {slide.notes && <span title="Есть заметки докладчика">📝</span>}
                </div>
                {/* Мини-превью слайда */}
                <div
                  className="relative bg-white"
                  style={{ aspectRatio: `${slideSize.w} / ${slideSize.h}` }}
                >
                  {slide.shapes
                    .filter((s) => s.type === 'image')
                    .slice(0, 3)
                    .map((shape, si) => (
                      <img
                        key={si}
                        src={shape.src}
                        alt=""
                        className="absolute object-cover"
                        style={{
                          left: `${shape.x}%`,
                          top: `${shape.y}%`,
                          width: `${shape.w}%`,
                          height: `${shape.h}%`,
                        }}
                      />
                    ))}
                  {slide.texts.slice(0, 3).map((t, ti) => (
                    <div
                      key={ti}
                      className="absolute overflow-hidden text-[5px] leading-[1.15] text-slate-900"
                      style={{
                        left: `${t.x}%`,
                        top: `${t.y}%`,
                        width: `${t.w}%`,
                        height: `${t.h}%`,
                        fontWeight: t.isTitle ? 700 : 400,
                      }}
                    >
                      {t.text}
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Основная область */}
        <div className="flex min-w-0 flex-1 flex-col">
          {showGrid ? (
            /* Режим сетки */
            <div className="min-h-0 flex-1 overflow-y-auto bg-ws-editor-alt p-5">
              <div className="grid grid-cols-2 gap-4 lg:grid-cols-3">
                {slides.map((slide, idx) => (
                  <div
                    key={slide.index}
                    onClick={() => {
                      setCurrent(idx);
                      setShowGrid(false);
                    }}
                    className={cn(
                      'cursor-pointer overflow-hidden rounded-lg border transition-all hover:border-orange-500',
                      current === idx ? 'border-orange-500 ring-2 ring-orange-500/40' : 'border-ws'
                    )}
                  >
                    <div
                      className="relative bg-white"
                      style={{ aspectRatio: `${slideSize.w} / ${slideSize.h}` }}
                    >
                      {slide.shapes
                        .filter((s) => s.type === 'image')
                        .slice(0, 4)
                        .map((shape, si) => (
                          <img
                            key={si}
                            src={shape.src}
                            alt=""
                            className="absolute object-cover"
                            style={{
                              left: `${shape.x}%`,
                              top: `${shape.y}%`,
                              width: `${shape.w}%`,
                              height: `${shape.h}%`,
                            }}
                          />
                        ))}
                      {slide.texts.slice(0, 6).map((t, ti) => (
                        <div
                          key={ti}
                          className="absolute overflow-hidden text-[7px] leading-[1.2] text-slate-900"
                          style={{
                            left: `${t.x}%`,
                            top: `${t.y}%`,
                            width: `${t.w}%`,
                            height: `${t.h}%`,
                            fontWeight: t.isTitle ? 700 : 400,
                          }}
                        >
                          {t.text}
                        </div>
                      ))}
                    </div>
                    <div className="bg-ws-surface px-2 py-1 text-[10px] text-ws-faint">
                      Слайд {slide.index}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ) : (
            /* Режим одного слайда */
            <div className="min-h-0 flex-1 overflow-auto bg-ws-editor-alt p-8">
              <div className="flex h-full items-center justify-center">
                {activeSlide && (
                  <div
                    ref={containerRef}
                    className="w-full max-w-5xl overflow-hidden rounded-lg border border-ws shadow-ws"
                    style={{ aspectRatio: `${slideSize.w} / ${slideSize.h}` }}
                  >
                    <SlideCanvas slide={activeSlide} />
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Заметки докладчика */}
          {activeSlide?.notes && (
            <div className="max-h-24 shrink-0 overflow-y-auto border-t border-ws bg-ws-surface px-4 py-2">
              <div className="mb-1 text-[10px] font-semibold tracking-wider text-ws-muted uppercase">
                Заметки докладчика
              </div>
              <p className="text-[11px] leading-relaxed text-ws-muted select-text">
                {activeSlide.notes}
              </p>
            </div>
          )}
        </div>
      </div>

      {/* Строка состояния */}
      <div className="flex shrink-0 items-center justify-between border-t border-ws bg-ws-surface px-3 py-1 text-[10px] text-ws-faint">
        <span className="truncate">{tab.name}</span>
        <div className="flex items-center gap-3">
          <span className="flex items-center gap-1">
            <Type size={10} />
            Текстовых блоков: {activeSlide?.texts.length ?? 0}
          </span>
          <span className="flex items-center gap-1">
            <ImageIcon size={10} />
            Изображений: {activeSlide?.shapes.length ?? 0}
          </span>
        </div>
      </div>
    </div>
  );
};

/* ============================================================
   Рендер слайда с реальным позиционированием элементов
   ============================================================ */

const SlideCanvas: React.FC<{ slide: ParsedSlide; fullscreen?: boolean }> = ({
  slide,
  fullscreen,
}) => (
  <div className="relative h-full w-full bg-white">
    {/* Изображения под текстом */}
    {slide.shapes.map((shape, idx) =>
      shape.type === 'image' && shape.src ? (
        <img
          key={idx}
          src={shape.src}
          alt=""
          className="absolute object-contain"
          style={{
            left: `${shape.x}%`,
            top: `${shape.y}%`,
            width: `${shape.w}%`,
            height: `${shape.h}%`,
          }}
        />
      ) : null
    )}

    {/* Текстовые блоки */}
    {slide.texts.map((t, idx) => (
      <div
        key={idx}
        className={cn(
          'absolute overflow-hidden text-slate-900',
          fullscreen ? 'leading-snug' : 'leading-[1.3]'
        )}
        style={{
          left: `${t.x}%`,
          top: `${t.y}%`,
          width: `${t.w}%`,
          height: `${t.h}%`,
          // Масштабируем размер шрифта от размера слайда
          fontSize: t.fontSize
            ? `calc(${t.fontSize}pt * var(--slide-scale, 1))`
            : t.isTitle
              ? fullscreen
                ? '3.2vw'
                : '1.6rem'
              : fullscreen
                ? '1.6vw'
                : '0.95rem',
          fontWeight: t.isTitle ? 700 : t.bold ? 600 : 400,
          display: 'flex',
          flexDirection: 'column',
          justifyContent: 'center',
        }}
      >
        {t.text.split('\n').map((line, li) => (
          <div key={li}>{line}</div>
        ))}
      </div>
    ))}

    {slide.texts.length === 0 && slide.shapes.length === 0 && (
      <div className="flex h-full items-center justify-center text-sm text-slate-400">
        Слайд не содержит текстовых элементов
      </div>
    )}
  </div>
);
