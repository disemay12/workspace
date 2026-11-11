/**
 * Просмотр и правка PDF: одна страница на весь экран, листание стрелками и
 * колесом мыши, поиск с подсветкой найденных фраз, аннотации (маркер, заметка).
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ChevronLeft,
  ChevronRight,
  Highlighter,
  RotateCcw,
  RotateCw,
  Scissors,
  StickyNote,
  Trash2,
} from 'lucide-react';
import { Button } from '../ui';
import type { OpenTab } from '../../types';
import { fileSystemService } from '../../services/fileSystemService';
import {
  base64ToBytes,
  deletePages,
  extractPages,
  inspectPdf,
  parsePageRange,
  rotatePages,
  type PdfStructure,
} from '../../lib/pdfOps';
import { applyAnnotations, normalizeClick, type AnnotationOp } from '../../lib/pdfAnnotate';
import {
  findAllMatches,
  loadPdfDocument,
  renderPageToDataUrl,
  type PdfJsDocument,
  type TextMatch,
} from '../../lib/pdfRender';

interface Props {
  tab: OpenTab;
  onChange: (content: string) => void;
  onSave: () => void;
  onNotify: (message: string) => void;
}

const READ_LIMIT = 64 * 1024 * 1024;
/** Ширина рендера страницы: больше окна, чтобы читался текст. */
const PAGE_WIDTH = 1400;

export const PdfEditor: React.FC<Props> = ({ tab, onChange, onSave, onNotify }) => {
  const [base64, setBase64] = useState('');
  const [structure, setStructure] = useState<PdfStructure | null>(null);
  const [page, setPage] = useState(0);
  const [pageImage, setPageImage] = useState('');
  const [docEpoch, setDocEpoch] = useState(0);
  const [rendering, setRendering] = useState(false);
  const [selected, setSelected] = useState<number[]>([]);
  const [range, setRange] = useState('');
  const [noteText, setNoteText] = useState('');
  const [point, setPoint] = useState<{ page: number; x: number; y: number } | null>(null);
  const [search, setSearch] = useState('');
  const [matches, setMatches] = useState<TextMatch[]>([]);
  const [matchIndex, setMatchIndex] = useState(0);
  const [fontBytes, setFontBytes] = useState<Uint8Array | null>(null);
  const [ops, setOps] = useState<AnnotationOp[]>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const docRef = useRef<PdfJsDocument | null>(null);
  const baseRef = useRef('');
  const opsRef = useRef<AnnotationOp[]>([]);
  const cacheRef = useRef<Map<number, string>>(new Map());
  const tokenRef = useRef(0);

  const pageCount = structure?.pageCount ?? 0;
  const pageInfo = structure?.pages[page];
  const pageRatio = pageInfo ? `${pageInfo.width} / ${pageInfo.height}` : '4 / 5';

  const renderPage = useCallback(
    async (index: number) => {
      const doc = docRef.current;
      if (!doc) return;
      const cached = cacheRef.current.get(index);
      if (cached) {
        setPageImage(cached);
        return;
      }
      const token = ++tokenRef.current;
      setRendering(true);
      try {
        const image = await renderPageToDataUrl(doc, index, PAGE_WIDTH);
        if (token !== tokenRef.current) return;
        cacheRef.current.set(index, image);
        setPageImage(image);
      } catch (err) {
        setPageImage('');
        onNotify(`Страница не отрисована: ${(err as Error).message}`);
      } finally {
        if (token === tokenRef.current) setRendering(false);
      }
    },
    [onNotify]
  );

  // Системный шрифт нужен для заметок с кириллицей.
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      for (const fontPath of ['C:/Windows/Fonts/arial.ttf', 'C:/Windows/Fonts/segoeui.ttf']) {
        try {
          const res = await fileSystemService.readPreview(fontPath, 8 * 1024 * 1024);
          if (!cancelled) setFontBytes(base64ToBytes(res.base64));
          return;
        } catch {
          // шрифта нет — пробуем следующий
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const reload = useCallback(async (nextBase64: string, keepPage = true) => {
    setBusy(true);
    setError('');
    try {
      const next = await inspectPdf(nextBase64);
      const doc = await loadPdfDocument(nextBase64);
      docRef.current = doc;
      cacheRef.current.clear();
      setBase64(nextBase64);
      setStructure(next);
      setPage((prev) => {
        const target = keepPage ? Math.min(prev, next.pageCount - 1) : 0;
        return Math.max(0, target);
      });
      setMatches([]);
      setMatchIndex(0);
      setDocEpoch((value) => value + 1);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setBusy(false);
    }
  }, []);

  useEffect(() => {
    let cancelled = false;
    setBusy(true);
    void fileSystemService
      .readPreview(tab.path, READ_LIMIT)
      .then(async (res) => {
        if (cancelled) return;
        baseRef.current = res.base64;
        opsRef.current = [];
        setOps([]);
        await reload(res.base64, false);
      })
      .catch((err) => {
        if (!cancelled) setError((err as Error).message);
      })
      .finally(() => {
        if (!cancelled) setBusy(false);
      });
    return () => {
      cancelled = true;
      tokenRef.current += 1;
      docRef.current = null;
      cacheRef.current.clear();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.path]);

  // Отрисовка текущей страницы и сброс выбора на неё.
  useEffect(() => {
    if (!docRef.current || !structure) return;
    setSelected([page]);
    setPoint(null);
    void renderPage(page);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page, docEpoch, renderPage]);

  const goTo = useCallback(
    (next: number) => {
      if (pageCount === 0) return;
      setPage((prev) => {
        const target = Math.min(pageCount - 1, Math.max(0, next));
        return target === prev ? prev : target;
      });
    },
    [pageCount]
  );

  const focusMatch = useCallback(
    (index: number) => {
      if (matches.length === 0) return;
      const wrapped = ((index % matches.length) + matches.length) % matches.length;
      setMatchIndex(wrapped);
      goTo(matches[wrapped].page);
    },
    [matches, goTo]
  );

  const clearMatches = () => {
    setMatches([]);
    setMatchIndex(0);
  };

  const apply = async (label: string, operation: () => Promise<string>) => {
    if (busy) return;
    setBusy(true);
    try {
      const next = await operation();
      baseRef.current = next;
      opsRef.current = [];
      setOps([]);
      await reload(next);
      onChange(next);
      const info = await inspectPdf(next);
      onNotify(`${label}: ${info.pageCount} стр. — сохраните файл (Ctrl+S)`);
    } catch (err) {
      onNotify(`${label}: ${(err as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  const targets = selected.length > 0 ? selected : [page];

  const handlePageClick = (event: React.MouseEvent<HTMLDivElement>) => {
    const rect = event.currentTarget.getBoundingClientRect();
    const normalized = normalizeClick(rect, event.clientX, event.clientY);
    setPoint({ page, x: normalized.x, y: normalized.y });
    // Клик в любом месте страницы снимает подсветку поиска.
    if (matches.length > 0) clearMatches();
  };

  /** Пересобирает документ из базового PDF и списка аннотаций. */
  const rebuild = async (nextOps: AnnotationOp[]) => {
    if (busy) return;
    setBusy(true);
    try {
      const next = await applyAnnotations(baseRef.current, nextOps, fontBytes ?? undefined);
      opsRef.current = nextOps;
      setOps(nextOps);
      await reload(next);
      onChange(next);
      onNotify(
        nextOps.length > 0
          ? `Аннотаций: ${nextOps.length} — сохраните файл (Ctrl+S)`
          : 'Аннотации удалены — сохраните файл (Ctrl+S)'
      );
    } catch (err) {
      onNotify(`Аннотации: ${(err as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  const addMarker = () => {
    if (!point) {
      onNotify('Кликните по странице, чтобы выбрать место');
      return;
    }
    void rebuild([...opsRef.current, { kind: 'marker', page: point.page, x: point.x, y: point.y }]);
  };

  const addTextNote = () => {
    if (!point) {
      onNotify('Кликните по странице, чтобы выбрать место');
      return;
    }
    if (!noteText.trim()) {
      onNotify('Введите текст заметки');
      return;
    }
    void rebuild([
      ...opsRef.current,
      { kind: 'note', page: point.page, x: point.x, y: point.y, text: noteText },
    ]);
  };

  const undoAnnotation = () => {
    if (opsRef.current.length === 0) {
      onNotify('Добавленных аннотаций нет');
      return;
    }
    void rebuild(opsRef.current.slice(0, -1));
  };

  const removeAnnotation = (index: number) => {
    void rebuild(opsRef.current.filter((_, position) => position !== index));
  };

  const runSearch = async () => {
    const query = search.trim();
    const doc = docRef.current;
    if (!query || !doc) {
      clearMatches();
      return;
    }
    setBusy(true);
    try {
      const found = await findAllMatches(doc, query);
      setMatches(found);
      setMatchIndex(0);
      if (found.length > 0) {
        goTo(found[0].page);
        onNotify(`Найдено совпадений: ${found.length}`);
      } else {
        onNotify('Не найдено');
      }
    } finally {
      setBusy(false);
    }
  };

  const pageMatches = useMemo(
    () => matches.filter((match) => match.page === page),
    [matches, page]
  );
  const activeMatch = matches[matchIndex];
  const hits = useMemo(() => [...new Set(matches.map((match) => match.page))], [matches]);

  const summary = useMemo(() => {
    if (!structure) return '';
    const first = structure.pages[0];
    return first
      ? `${structure.pageCount} стр. · ${first.width}×${first.height} pt${structure.title ? ` · ${structure.title}` : ''}`
      : 'Пустой документ';
  }, [structure]);

  if (error && !structure) {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-2 p-6 text-center text-xs text-ws-muted">
        <p>Не удалось прочитать PDF: {error}</p>
      </div>
    );
  }

  return (
    <div
      className="flex h-full min-h-0 flex-col outline-none"
      tabIndex={0}
      onKeyDown={(event) => {
        if (event.key === 'ArrowRight' || event.key === 'PageDown') {
          event.preventDefault();
          goTo(page + 1);
        } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
          event.preventDefault();
          goTo(page - 1);
        } else if (event.ctrlKey && (event.key === 'z' || event.key === 'Z')) {
          event.preventDefault();
          undoAnnotation();
        }
      }}
    >
      <div className="flex flex-wrap items-center gap-1.5 border-b border-ws bg-ws-elevated px-3 py-2 text-[11px]">
        <span className="text-ws-muted">{busy ? 'Обработка…' : summary}</span>
        <span className="text-ws-faint">· выбрано {selected.length || 1}</span>
        <div className="ml-auto flex flex-wrap items-center gap-1.5">
          <input
            value={range}
            onChange={(event) => setRange(event.target.value)}
            placeholder="1-3,5"
            className="w-20 rounded border border-ws bg-ws-editor px-2 py-1 font-mono text-[11px] text-ws"
          />
          <Button
            onClick={() => {
              if (!structure) return;
              const parsed = parsePageRange(range, structure.pageCount);
              if (parsed.length === 0) {
                onNotify('Диапазон не распознан: укажите, например, 1-3,5');
                return;
              }
              setSelected(parsed);
              goTo(parsed[0]);
              onNotify(`Выбрано страниц: ${parsed.length}`);
            }}
          >
            Выбрать
          </Button>
          <Button onClick={() => apply('Поворот ↺', () => rotatePages(base64, targets, -90))} disabled={busy}>
            <RotateCcw size={12} />
          </Button>
          <Button onClick={() => apply('Поворот ↻', () => rotatePages(base64, targets, 90))} disabled={busy}>
            <RotateCw size={12} />
          </Button>
          <Button
            onClick={() => apply('Извлечение', () => extractPages(base64, targets))}
            disabled={busy || selected.length === 0}
          >
            <Scissors size={12} /> Только выбранные
          </Button>
          <Button onClick={() => apply('Удаление', () => deletePages(base64, targets))} disabled={busy || pageCount === 0}>
            <Trash2 size={12} /> Удалить
          </Button>
          <Button onClick={onSave} disabled={busy}>
            Сохранить
          </Button>
        </div>
      </div>

      <div className="flex flex-wrap items-center gap-1.5 border-b border-ws bg-ws-surface px-3 py-1.5 text-[11px]">
        <input
          value={search}
          onChange={(event) => setSearch(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') void runSearch();
          }}
          placeholder="Поиск по тексту"
          className="w-40 rounded border border-ws bg-ws-editor px-2 py-1 text-[11px] text-ws"
        />
        <Button onClick={() => void runSearch()} disabled={busy || !structure}>
          Найти
        </Button>
        {matches.length > 0 && (
          <span className="flex items-center gap-1 text-ws-faint">
            <button
              onClick={() => focusMatch(matchIndex - 1)}
              className="rounded border border-ws px-1 text-[10px] text-ws-muted hover:bg-ws-hover"
              title="Предыдущее совпадение"
            >
              ‹
            </button>
            <span className="tabular-nums text-ws">
              {matchIndex + 1} / {matches.length}
            </span>
            <button
              onClick={() => focusMatch(matchIndex + 1)}
              className="rounded border border-ws px-1 text-[10px] text-ws-muted hover:bg-ws-hover"
              title="Следующее совпадение"
            >
              ›
            </button>
            {hits.length > 1 && (
              <span className="ml-1 flex items-center gap-1">
                стр.:
                {hits.slice(0, 8).map((hit) => (
                  <button
                    key={hit}
                    onClick={() => focusMatch(matches.findIndex((match) => match.page === hit))}
                    className={
                      hit === page
                        ? 'rounded border border-ws-accent px-1 text-[10px] text-ws-accent'
                        : 'rounded border border-ws px-1 text-[10px] text-ws-muted hover:bg-ws-hover'
                    }
                  >
                    {hit + 1}
                  </button>
                ))}
              </span>
            )}
            <button
              onClick={clearMatches}
              className="ml-1 rounded border border-ws px-1 text-[10px] text-ws-muted hover:bg-ws-hover"
              title="Снять выделение"
            >
              ✕
            </button>
          </span>
        )}
        <span className="mx-2 h-4 w-px bg-ws" />
        <input
          value={noteText}
          onChange={(event) => setNoteText(event.target.value)}
          placeholder="Текст заметки"
          className="w-48 rounded border border-ws bg-ws-editor px-2 py-1 text-[11px] text-ws"
        />
        <Button onClick={addTextNote} disabled={busy}>
          <StickyNote size={12} /> Заметка
        </Button>
        <Button onClick={addMarker} disabled={busy}>
          <Highlighter size={12} /> Маркер
        </Button>
        <Button onClick={undoAnnotation} disabled={busy || ops.length === 0} title="Отменить последнюю аннотацию (Ctrl+Z)">
          Отменить
        </Button>
        {ops.length > 0 && (
          <span className="flex flex-wrap items-center gap-1 text-[10px]">
            {ops.map((op, index) => (
              <span
                key={`op-${index}-${op.page}-${op.x}-${op.kind}`}
                className="flex items-center gap-1 rounded border border-ws bg-ws-elevated px-1 py-0.5 text-ws-muted"
                title={op.text ? op.text : 'маркер'}
              >
                {op.kind === 'marker' ? 'маркер' : 'заметка'} стр. {op.page + 1}
                <button
                  onClick={() => removeAnnotation(index)}
                  disabled={busy}
                  className="rounded px-0.5 text-ws-faint hover:bg-ws-hover hover:text-ws-danger"
                  title="Удалить аннотацию"
                >
                  ✕
                </button>
              </span>
            ))}
          </span>
        )}
        <span className="text-ws-faint">
          {point
            ? `точка: стр. ${point.page + 1}, ${Math.round(point.x * 100)}% / ${Math.round(point.y * 100)}%`
            : 'кликните по странице'}
        </span>
      </div>

      <div
        className="flex min-h-0 flex-1 flex-col overflow-hidden bg-ws-surface/40"
        onWheel={(event) => {
          if (Math.abs(event.deltaY) < 6) return;
          goTo(page + (event.deltaY > 0 ? 1 : -1));
        }}
      >
        <div className="flex min-h-0 flex-1 flex-col items-center gap-2 p-3">
          <div
            onClick={handlePageClick}
            className="relative flex min-h-0 w-full flex-1 cursor-crosshair items-center justify-center overflow-hidden rounded border border-ws bg-white shadow-sm"
            title="Клик — место для заметки или маркера; клик снимает подсветку поиска"
          >
            {pageImage ? (
              <div
                className="relative h-full max-h-full"
                style={{ lineHeight: 0, aspectRatio: pageRatio }}
              >
                <img
                  src={pageImage}
                  alt={`Страница ${page + 1}`}
                  className="h-full w-full select-none object-fill"
                  draggable={false}
                />
                <div className="pointer-events-none absolute inset-0">
                    {pageMatches.map((match, index) => {
                      const globalIndex = matches.indexOf(match);
                      const isActive = globalIndex === matchIndex;
                      return (
                        <div
                          key={`${match.page}-${index}-${match.x}-${match.y}`}
                          className={
                            isActive
                              ? 'absolute rounded-[1px] bg-orange-400/60 ring-1 ring-orange-500'
                              : 'absolute rounded-[1px] bg-yellow-300/60 ring-1 ring-yellow-400'
                          }
                          style={{
                            left: `${match.x * 100}%`,
                            top: `${match.y * 100}%`,
                            width: `${match.width * 100}%`,
                            height: `${match.height * 100}%`,
                          }}
                        />
                      );
                    })}
                </div>
              </div>
            ) : (
              <div className="flex h-full items-center justify-center text-xs text-ws-faint">
                {rendering || busy ? 'Рендерю страницу…' : 'Нет страницы'}
              </div>
            )}
          </div>

          <div className="flex shrink-0 items-center gap-3 text-xs text-ws-muted">
            <Button onClick={() => goTo(page - 1)} disabled={page === 0 || pageCount === 0} title="Предыдущая (←)">
              <ChevronLeft size={14} />
            </Button>
            <span className="min-w-24 text-center tabular-nums">
              {pageCount === 0 ? '—' : `${page + 1} / ${pageCount}`}
            </span>
            <Button
              onClick={() => goTo(page + 1)}
              disabled={page >= pageCount - 1 || pageCount === 0}
              title="Следующая (→)"
            >
              <ChevronRight size={14} />
            </Button>
            {activeMatch && (
              <span className="text-[10px] text-ws-faint">
                совпадение на стр. {activeMatch.page + 1}
              </span>
            )}
            <span className="text-[10px] text-ws-faint">колесо мыши и ← → листают страницы</span>
          </div>
        </div>
      </div>
    </div>
  );
};
