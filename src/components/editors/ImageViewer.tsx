import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  Download,
  FlipHorizontal,
  FlipVertical,
  RotateCcw,
  RotateCw,
  Save,
  ZoomIn,
  ZoomOut,
  Undo2,
  Crop,
  UploadCloud,
  Check,
  X,
  SlidersHorizontal,
} from 'lucide-react';
import { OpenTab } from '../../types';
import { fileSystemService, imageDataUrl, imageMime } from '../../services/fileSystemService';
import {
  SHARE_SERVICES,
  copyToClipboard,
  gdriveUpload,
  getImgurClientId,
  shareViaSite,
  uploadBase64,
} from '../../services/shareService';
import { readPreferences } from '../../state/preferences';
import { applyImageOps, normalizeImageOps, type ImageOp } from '../../lib/imageOps';
import { registerImageController } from '../../services/imageBridge';

interface Props {
  tab: OpenTab;
  onChange: (content: string) => void;
  onSave: () => void;
  onNotify: (message: string) => void;
}

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Формат не поддерживается встроенным просмотрщиком'));
    image.src = src;
  });
}

function encoded(canvas: HTMLCanvasElement, mime: string): string {
  const supported = ['image/png', 'image/jpeg', 'image/webp'];
  const type = supported.includes(mime) ? mime : 'image/png';
  return canvas.toDataURL(type, 0.92).split(',')[1] || '';
}

export const ImageViewer: React.FC<Props> = ({ tab, onChange, onSave, onNotify }) => {
  const [data, setData] = useState(tab.content);
  const [zoom, setZoom] = useState(1);
  const [fit, setFit] = useState(true);
  const [natural, setNatural] = useState({ width: 0, height: 0 });
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [cropMode, setCropMode] = useState(false);
  const [selection, setSelection] = useState<Rect | null>(null);
  const [exportOpen, setExportOpen] = useState(false);
  const imgRef = useRef<HTMLImageElement>(null);
  const dragRef = useRef<{ x: number; y: number } | null>(null);
  const originalMime = imageMime(tab.extension);
  const editable = ['image/png', 'image/jpeg', 'image/webp'].includes(originalMime);
  const src = useMemo(() => imageDataUrl(data, tab.extension), [data, tab.extension]);

  useEffect(() => {
    setData(tab.content);
    setError('');
    setCropMode(false);
    setSelection(null);
  }, [tab.id]);

  const commit = (next: string) => {
    setData(next);
    onChange(next);
    setCropMode(false);
    setSelection(null);
  };

  /* AI-мост: правки изображения инструментами идут через буфер редактора. */
  useEffect(() => {
    return registerImageController({
      tabId: tab.id,
      getImage: () => ({
        base64: data.split(',')[1] ?? '',
        mime: /^data:([^;]+);/.exec(data)?.[1] ?? 'image/png',
        width: natural.width,
        height: natural.height,
        path: tab.path,
      }),
      applyOps: async (ops) => {
        try {
          const mime = /^data:([^;]+);/.exec(data)?.[1] ?? 'image/png';
          const base64 = data.split(',')[1] ?? '';
          const result = await applyImageOps({ base64, mime }, normalizeImageOps(ops));
          commit(`data:${result.mime};base64,${result.base64}`);
          return {
            ok: true,
            message: `${result.width}×${result.height}, ${Math.round(result.bytes / 1024)} КБ`,
          };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
      save: async () => {
        try {
          await onSave();
          return { ok: true, message: 'Изображение сохранено' };
        } catch (err) {
          return { ok: false, message: (err as Error).message };
        }
      },
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, natural.width, natural.height, tab.id, tab.path]);

  const editCanvas = async (
    kind: 'left' | 'right' | 'flipH' | 'flipV' | 'crop',
    rect?: Rect
  ) => {
    try {
      setBusy(true);
      const image = await loadImage(src);
      const swap = kind === 'left' || kind === 'right';
      const width =
        kind === 'crop' && rect
          ? Math.max(1, Math.round(rect.w))
          : swap
            ? image.naturalHeight
            : image.naturalWidth;
      const height =
        kind === 'crop' && rect
          ? Math.max(1, Math.round(rect.h))
          : swap
            ? image.naturalWidth
            : image.naturalHeight;
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      if (!ctx) throw new Error('Canvas недоступен');
      ctx.imageSmoothingQuality = 'high';
      if (kind === 'left' || kind === 'right') {
        ctx.translate(width / 2, height / 2);
        ctx.rotate((kind === 'right' ? 1 : -1) * Math.PI / 2);
        ctx.drawImage(image, -image.naturalWidth / 2, -image.naturalHeight / 2);
      } else if (kind === 'flipH') {
        ctx.scale(-1, 1);
        ctx.drawImage(image, -width, 0);
      } else if (kind === 'flipV') {
        ctx.scale(1, -1);
        ctx.drawImage(image, 0, -height);
      } else if (rect) {
        ctx.drawImage(image, rect.x, rect.y, rect.w, rect.h, 0, 0, width, height);
      }
      commit(encoded(canvas, originalMime));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const pointerPos = (e: React.PointerEvent) => {
    const box = imgRef.current?.getBoundingClientRect();
    if (!box) return { x: 0, y: 0 };
    return {
      x: Math.min(box.width, Math.max(0, e.clientX - box.left)),
      y: Math.min(box.height, Math.max(0, e.clientY - box.top)),
    };
  };

  const applyCrop = () => {
    const img = imgRef.current;
    if (!img || !selection || selection.w < 4 || selection.h < 4) {
      onNotify('Выделите область обрезки мышью');
      return;
    }
    const box = img.getBoundingClientRect();
    void editCanvas('crop', {
      x: (selection.x / box.width) * img.naturalWidth,
      y: (selection.y / box.height) * img.naturalHeight,
      w: (selection.w / box.width) * img.naturalWidth,
      h: (selection.h / box.height) * img.naturalHeight,
    });
  };

  const saveAs = async () => {
    const target = await fileSystemService.saveFileDialog(tab.name);
    if (!target) return;
    try {
      await fileSystemService.writeFile(target, data, 'base64');
      onNotify('Сохранено: ' + target);
    } catch (e) {
      onNotify('Ошибка сохранения: ' + (e as Error).message);
    }
  };

  const exportTo = async (service: string) => {
    setExportOpen(false);
    setBusy(true);
    try {
      if (service === 'imgbb' || service === 'freeimage' || service === 'yapx') {
        const result = await shareViaSite(
          service,
          [{ path: tab.path, name: tab.name }],
          { authorized: readPreferences().yapxAccountMode }
        );
        const name =
          service === 'imgbb' ? 'ImgBB' : service === 'freeimage' ? 'Freeimage.host' : 'Япикс';
        if (!result.success) {
          throw new Error(result.error || result.skipped[0]?.reason || 'Не удалось открыть ' + name);
        }
        if (result.albums.length > 0) {
          await copyToClipboard(result.albums.join('\n'));
          onNotify(name + ': альбом создан, ссылка скопирована');
        } else {
          onNotify(name + ' открыт, фото добавлено в форму загрузки');
        }
        return;
      }
      if (service === 'gdrive') {
        if (tab.isDirty) {
          throw new Error('Сначала сохраните изменения изображения, затем загрузите его в Google Drive');
        }
        const result = await gdriveUpload([tab.path]);
        if (!result.success) throw new Error(result.error || result.failed.join(', ') || 'Не удалось загрузить');
        if (result.link) await copyToClipboard(result.link);
        onNotify('Загружено в Google Drive' + (result.link ? ' — ссылка скопирована' : ''));
        return;
      }
      if (service === 'imgur' && !getImgurClientId()) {
        throw new Error('Укажите Imgur Client-ID в Настройках → Сервисы');
      }
      const result = await uploadBase64(service, tab.name, originalMime, data);
      if (!result.success || !result.url) throw new Error(result.error || 'Не удалось получить ссылку');
      await copyToClipboard(result.url);
      onNotify('Ссылка скопирована: ' + result.url);
    } catch (e) {
      onNotify('Экспорт не выполнен: ' + (e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  /* Панель размера, формата и сжатия. */
  const [tuneOpen, setTuneOpen] = useState(false);
  const [tuneWidth, setTuneWidth] = useState(0);
  const [tuneFormat, setTuneFormat] = useState<'png' | 'jpeg' | 'webp'>('png');
  const [tuneQuality, setTuneQuality] = useState(0.9);
  const [tuneMaxKb, setTuneMaxKb] = useState(0);

  const applyTune = async () => {
    try {
      setBusy(true);
      const mime = /^data:([^;]+);/.exec(data)?.[1] ?? 'image/png';
      const base64 = data.split(',')[1] ?? '';
      const ops: ImageOp[] = [];
      if (tuneWidth > 0) ops.push({ op: 'resize', mode: 'fit', width: tuneWidth });
      ops.push({ op: 'convert', format: tuneFormat, quality: tuneQuality });
      if (tuneMaxKb > 0) {
        ops.push({ op: 'compress', maxKb: tuneMaxKb, format: tuneFormat === 'webp' ? 'webp' : 'jpeg' });
      }
      const result = await applyImageOps({ base64, mime }, ops);
      commit(`data:${result.mime};base64,${result.base64}`);
      onNotify(`Готово: ${result.width}×${result.height}, ${Math.round(result.bytes / 1024)} КБ`);
      setTuneOpen(false);
    } catch (err) {
      onNotify((err as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const changeZoom = (next: number) => {
    setFit(false);
    setZoom(Math.min(8, Math.max(0.1, next)));
  };

  return (
    <div className="flex h-full min-h-0 flex-col bg-ws-editor-alt">
      <div className="flex flex-wrap items-center gap-1 border-b border-ws bg-ws-surface px-2 py-1.5">
        <button title="Уменьшить" onClick={() => changeZoom(zoom / 1.25)} className="rounded p-1.5 text-ws-muted hover:bg-ws-hover"><ZoomOut size={14} /></button>
        <span className="min-w-14 text-center text-[11px] text-ws-muted">{fit ? 'Вписано' : Math.round(zoom * 100) + '%'}</span>
        <button title="Увеличить" onClick={() => changeZoom(zoom * 1.25)} className="rounded p-1.5 text-ws-muted hover:bg-ws-hover"><ZoomIn size={14} /></button>
        <button title="Вписать в окно" onClick={() => setFit(true)} className="rounded px-2 py-1 text-[11px] text-ws-muted hover:bg-ws-hover">Вписать</button>
        <button title="100%" onClick={() => changeZoom(1)} className="rounded px-2 py-1 text-[11px] text-ws-muted hover:bg-ws-hover">100%</button>
        <span className="mx-1 h-4 border-l border-ws" />
        <button title="Повернуть влево" disabled={busy || !editable} onClick={() => void editCanvas('left')} className="rounded p-1.5 text-ws-muted hover:bg-ws-hover disabled:opacity-40"><RotateCcw size={14} /></button>
        <button title="Повернуть вправо" disabled={busy || !editable} onClick={() => void editCanvas('right')} className="rounded p-1.5 text-ws-muted hover:bg-ws-hover disabled:opacity-40"><RotateCw size={14} /></button>
        <button title="Отразить горизонтально" disabled={busy || !editable} onClick={() => void editCanvas('flipH')} className="rounded p-1.5 text-ws-muted hover:bg-ws-hover disabled:opacity-40"><FlipHorizontal size={14} /></button>
        <button title="Отразить вертикально" disabled={busy || !editable} onClick={() => void editCanvas('flipV')} className="rounded p-1.5 text-ws-muted hover:bg-ws-hover disabled:opacity-40"><FlipVertical size={14} /></button>
        <button
          title="Обрезать"
          disabled={busy || !editable}
          onClick={() => { setCropMode((v) => !v); setSelection(null); }}
          className={'rounded p-1.5 hover:bg-ws-hover disabled:opacity-40 ' + (cropMode ? 'bg-ws-accent-soft text-ws-accent' : 'text-ws-muted')}
        >
          <Crop size={14} />
        </button>
        {cropMode && (
          <>
            <button title="Применить обрезку" onClick={applyCrop} className="rounded p-1.5 text-ws-success hover:bg-ws-hover"><Check size={14} /></button>
            <button title="Отменить обрезку" onClick={() => { setCropMode(false); setSelection(null); }} className="rounded p-1.5 text-ws-muted hover:bg-ws-hover"><X size={14} /></button>
          </>
        )}
        <span className="mx-1 h-4 border-l border-ws" />
        <button title="Вернуть исходное изображение" onClick={() => commit(tab.initialContent)} className="rounded p-1.5 text-ws-muted hover:bg-ws-hover"><Undo2 size={14} /></button>
        <button title="Сохранить изменения" disabled={!tab.isDirty} onClick={onSave} className="flex items-center gap-1 rounded bg-ws-accent px-2 py-1 text-[11px] text-white disabled:opacity-40"><Save size={13} />Сохранить</button>
        <button title="Сохранить копию" onClick={() => void saveAs()} className="rounded p-1.5 text-ws-muted hover:bg-ws-hover"><Download size={14} /></button>
        <button title="Размер, формат и сжатие" disabled={busy} onClick={() => setTuneOpen((v) => !v)} className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-[11px] text-ws-muted hover:bg-ws-hover">
          <SlidersHorizontal size={13} /> Размер и сжатие
        </button>
        {tuneOpen && (
          <div className="absolute top-full left-0 z-30 mt-1 w-72 rounded-lg border border-ws bg-ws-surface p-3 shadow-ws">
            <div className="grid grid-cols-2 gap-2 text-[11px]">
              <label className="flex flex-col gap-1">
                <span className="text-ws-muted">Ширина, px</span>
                <input type="number" min={0} value={tuneWidth} onChange={(e) => setTuneWidth(Number(e.target.value))} className="rounded border border-ws bg-ws-editor px-2 py-1 text-ws" />
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-ws-muted">Сжать до, КБ</span>
                <input type="number" min={0} value={tuneMaxKb} onChange={(e) => setTuneMaxKb(Number(e.target.value))} className="rounded border border-ws bg-ws-editor px-2 py-1 text-ws" />
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-ws-muted">Формат</span>
                <select value={tuneFormat} onChange={(e) => setTuneFormat(e.target.value as 'png' | 'jpeg' | 'webp')} className="rounded border border-ws bg-ws-editor px-2 py-1 text-ws">
                  <option value="png">PNG</option>
                  <option value="jpeg">JPEG</option>
                  <option value="webp">WebP</option>
                </select>
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-ws-muted">Качество</span>
                <input type="number" min={0.3} max={1} step={0.05} value={tuneQuality} onChange={(e) => setTuneQuality(Number(e.target.value))} className="rounded border border-ws bg-ws-editor px-2 py-1 text-ws" />
              </label>
            </div>
            <div className="mt-2 flex items-center justify-end gap-2 text-[11px]">
              <button onClick={() => setTuneOpen(false)} className="rounded px-2 py-1 text-ws-muted hover:bg-ws-hover">Отмена</button>
              <button onClick={() => void applyTune()} className="rounded bg-ws-accent px-2 py-1 text-white">Применить</button>
            </div>
          </div>
        )}
        <div className="relative ml-auto">
          <button title="Экспорт" disabled={busy} onClick={() => setExportOpen((v) => !v)} className="flex items-center gap-1 rounded border border-ws px-2 py-1 text-[11px] text-ws-muted hover:bg-ws-hover">
            <UploadCloud size={13} /> Экспорт
          </button>
          {exportOpen && (
            <div className="absolute top-full right-0 z-30 mt-1 w-56 overflow-hidden rounded-lg border border-ws bg-ws-surface shadow-ws">
              {SHARE_SERVICES.map((s) => (
                <button key={s.id} title={s.hint} onClick={() => void exportTo(s.id)} className="block w-full px-3 py-1.5 text-left text-[11px] text-ws hover:bg-ws-hover">
                  <span className="font-medium">{s.label}</span>
                  <span className="ml-2 text-ws-faint">{s.hint}</span>
                </button>
              ))}
              <button onClick={() => void exportTo('gdrive')} className="block w-full border-t border-ws px-3 py-1.5 text-left text-[11px] text-ws hover:bg-ws-hover">Google Drive</button>
            </div>
          )}
        </div>
      </div>
      <div className="min-h-0 flex-1 overflow-auto p-4">
        {error ? (
          <div className="m-auto max-w-md rounded border border-amber-500/40 bg-amber-500/10 p-4 text-center text-xs text-ws-warning">
            {error}
            <p className="mt-2 text-ws-muted">Этот формат распознаётся как изображение, но Chromium не умеет его показать. Файл можно экспортировать как есть.</p>
          </div>
        ) : (
          <div className="flex min-h-full min-w-full items-center justify-center">
            <div className="relative inline-block">
              <img
                ref={imgRef}
                src={src}
                alt={tab.name}
                draggable={false}
                onLoad={(e) => {
                  setNatural({ width: e.currentTarget.naturalWidth, height: e.currentTarget.naturalHeight });
                  setError('');
                }}
                onError={() => setError('Не удалось отобразить формат изображения')}
                onPointerDown={(e) => {
                  if (!cropMode) return;
                  const p = pointerPos(e);
                  dragRef.current = p;
                  setSelection({ x: p.x, y: p.y, w: 0, h: 0 });
                  e.currentTarget.setPointerCapture(e.pointerId);
                }}
                onPointerMove={(e) => {
                  const start = dragRef.current;
                  if (!cropMode || !start) return;
                  const p = pointerPos(e);
                  setSelection({
                    x: Math.min(start.x, p.x),
                    y: Math.min(start.y, p.y),
                    w: Math.abs(p.x - start.x),
                    h: Math.abs(p.y - start.y),
                  });
                }}
                onPointerUp={() => { dragRef.current = null; }}
                style={
                  fit
                    ? { maxWidth: '100%', maxHeight: '70vh', width: 'auto', height: 'auto' }
                    : { width: Math.max(1, natural.width * zoom), height: 'auto', maxWidth: 'none' }
                }
                className={'rounded border border-ws shadow-ws ' + (cropMode ? 'cursor-crosshair' : '')}
              />
              {selection && selection.w > 1 && selection.h > 1 && (
                <div
                  className="pointer-events-none absolute border-2 border-ws-accent bg-ws-accent/20"
                  style={{ left: selection.x, top: selection.y, width: selection.w, height: selection.h }}
                />
              )}
            </div>
          </div>
        )}
      </div>
      <div className="flex h-6 items-center justify-between border-t border-ws bg-ws-surface px-3 text-[10px] text-ws-faint">
        <span className="truncate">{tab.name}{cropMode ? ' — выделите область и нажмите галочку' : ''}</span>
        <span>{natural.width ? natural.width + ' × ' + natural.height + ' px' : ''}</span>
      </div>
    </div>
  );
};
