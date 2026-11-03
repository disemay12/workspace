import React, { useEffect, useMemo, useState } from 'react';
import { Clapperboard, Gauge, Scissors, Upload, VolumeX } from 'lucide-react';
import { OpenTab } from '../../types';
import { Button, Field, Input, Select, Textarea } from '../ui';
import { gdriveStatus } from '../../services/shareService';
import { fileSystemService } from '../../services/fileSystemService';

interface Props {
  tab: OpenTab;
  onNotify: (message: string) => void;
  onOpenResult: (filePath: string) => void;
}

interface Probe {
  duration: number;
  width: number;
  height: number;
  hasAudio: boolean;
  videoCodec: string;
  audioCodec: string;
  size: number;
}

const FALLBACK_CATEGORIES = [
  { id: '22', title: 'Люди и блоги' },
  { id: '1', title: 'Фильмы и анимация' },
  { id: '10', title: 'Музыка' },
  { id: '20', title: 'Игры' },
  { id: '24', title: 'Развлечения' },
  { id: '27', title: 'Образование' },
  { id: '28', title: 'Наука и техника' },
];

function clock(total: number): string {
  const s = Math.max(0, Math.floor(total));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = s % 60;
  const pad = (n: number) => String(n).padStart(2, '0');
  return (h > 0 ? pad(h) + ':' : '') + pad(m) + ':' + pad(sec);
}

function sizeLabel(bytes: number): string {
  if (!bytes) return '';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(0) + ' КБ';
  return (bytes / (1024 * 1024)).toFixed(1) + ' МБ';
}

export const VideoEditor: React.FC<Props> = ({ tab, onNotify, onOpenResult }) => {
  const src = useMemo(() => 'media://local/?path=' + encodeURIComponent(tab.path), [tab.path]);
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const [probe, setProbe] = useState<Probe | null>(null);
  const [ffmpeg, setFfmpeg] = useState<string | null>(null);
  const [start, setStart] = useState(0);
  const [end, setEnd] = useState(0);
  const [speed, setSpeed] = useState(1);
  const [previewRate, setPreviewRate] = useState(1);
  const [speedStart, setSpeedStart] = useState(0);
  const [speedEnd, setSpeedEnd] = useState(0);
  const [speedValue, setSpeedValue] = useState(2);
  const [speedSegments, setSpeedSegments] = useState<
    { start: number; end: number; speed: number }[]
  >([]);
  const [stripAudio, setStripAudio] = useState(false);
  const [compress, setCompress] = useState(false);
  const [busy, setBusy] = useState('');
  const [progress, setProgress] = useState(0);
  const [publishOpen, setPublishOpen] = useState(false);
  const [connected, setConnected] = useState(false);
  const [categories, setCategories] = useState(FALLBACK_CATEGORIES);
  const [playlists, setPlaylists] = useState<{ id: string; title: string }[]>([]);
  const [meta, setMeta] = useState({
    title: tab.name.replace(/\.[^.]+$/, ''),
    description: '',
    tags: '',
    categoryId: '22',
    privacyStatus: 'private' as 'public' | 'unlisted' | 'private',
    madeForKids: false,
    embeddable: true,
    license: 'youtube' as 'youtube' | 'creativeCommon',
    defaultLanguage: 'ru',
    publishAt: '',
    playlistId: '',
  });

  useEffect(() => {
    let stop = false;
    void window.electronAPI?.videoCapabilities?.().then((caps) => {
      if (!stop) setFfmpeg(caps.ffmpeg);
    });
    void window.electronAPI?.videoProbe?.(tab.path).then((info) => {
      if (stop || !info.success) return;
      setProbe({
        duration: info.duration || 0,
        width: info.width || 0,
        height: info.height || 0,
        hasAudio: Boolean(info.hasAudio),
        videoCodec: info.videoCodec || '',
        audioCodec: info.audioCodec || '',
        size: info.size || 0,
      });
      setEnd(info.duration || 0);
      setSpeedEnd(info.duration || 0);
    });
    return () => { stop = true; };
  }, [tab.path]);

  useEffect(() => {
    if (videoRef.current) videoRef.current.playbackRate = previewRate;
  }, [previewRate]);

  const addSpeedSegment = () => {
    const duration = probe?.duration || 0;
    const from = Math.max(0, Math.min(duration, speedStart));
    const to = Math.max(0, Math.min(duration, speedEnd));
    if (to <= from) {
      onNotify('Конец фрагмента должен быть позже начала');
      return;
    }
    if (speedSegments.some((segment) => from < segment.end && to > segment.start)) {
      onNotify('Фрагмент пересекается с уже добавленным');
      return;
    }
    setSpeedSegments((prev) =>
      [...prev, { start: from, end: to, speed: speedValue }].sort((a, b) => a.start - b.start)
    );
    onNotify(`Фрагмент ${clock(from)}–${clock(to)}: ${speedValue}×`);
  };

  const apply = async () => {
    const baseName = tab.name.replace(/\.[^.]+$/, '');
    const suffix =
      speed !== 1 || speedSegments.length > 0
        ? '.speed'
        : compress
          ? '.compressed'
          : stripAudio
            ? '.muted'
            : '.copy';
    const selected = await fileSystemService.saveFileDialog(`${baseName}${suffix}.mp4`, [
      { name: 'Видео MP4', extensions: ['mp4'] },
    ]);
    if (!selected) return;
    const outputPath = /\.mp4$/i.test(selected) ? selected : `${selected}.mp4`;
    if (outputPath.toLowerCase() === tab.path.toLowerCase()) {
      onNotify('Выберите другое имя: нельзя сохранить копию поверх исходного видео');
      return;
    }

    setBusy('ffmpeg');
    try {
      const result = await window.electronAPI!.videoEdit({
        filePath: tab.path,
        start,
        end: end > start ? end : undefined,
        stripAudio,
        compress,
        speed,
        speedSegments,
        duration: probe?.duration || 0,
        hasAudio: probe?.hasAudio ?? true,
        outputPath,
      });
      if (!result.success || !result.outputPath) throw new Error(result.error || 'Не удалось обработать видео');
      onNotify('Готово: ' + result.outputPath);
      onOpenResult(result.outputPath);
    } catch (err) {
      onNotify((err as Error).message);
    } finally {
      setBusy('');
    }
  };

  const openPublish = async () => {
    setPublishOpen(true);
    const status = await gdriveStatus();
    setConnected(status.connected);
    if (!status.connected) return;
    const [cats, lists] = await Promise.all([
      window.electronAPI!.youtubeCategories(),
      window.electronAPI!.youtubePlaylists(),
    ]);
    if (cats.success && cats.categories.length) setCategories(cats.categories);
    if (lists.success) setPlaylists(lists.playlists);
  };

  const publish = async () => {
    if (!meta.title.trim()) {
      onNotify('Укажите название, как на сайте YouTube');
      return;
    }
    setBusy('youtube');
    setProgress(0);
    const off = window.electronAPI?.onYoutubeProgress?.((data) => {
      if (data.filePath === tab.path && data.total > 0) setProgress(data.sent / data.total);
    });
    try {
      const result = await window.electronAPI!.youtubeUpload(tab.path, {
        ...meta,
        tags: meta.tags.split(',').map((tag) => tag.trim()).filter(Boolean),
        playlistId: meta.playlistId || undefined,
        publishAt: meta.publishAt ? new Date(meta.publishAt).toISOString() : undefined,
      });
      if (!result.success) throw new Error(result.error || 'YouTube отклонил загрузку');
      onNotify('Опубликовано: ' + (result.url || result.id));
      setPublishOpen(false);
    } catch (err) {
      onNotify((err as Error).message);
    } finally {
      off?.();
      setBusy('');
    }
  };

  return (
    <div className="flex h-full min-h-0 flex-col bg-black text-ws">
      <div className="min-h-0 flex-1 bg-black">
        <video ref={videoRef} src={src} controls className="h-full w-full bg-black" />
      </div>
      <div className="max-h-[46%] space-y-3 overflow-auto border-t border-ws bg-ws-surface p-3">
        <div className="flex flex-wrap items-center gap-3 text-[11px] text-ws-muted">
          <span className="flex items-center gap-1 text-ws"><Clapperboard size={13} /> {tab.name}</span>
          {probe && <span>{probe.width}×{probe.height}</span>}
          {probe && <span>{clock(probe.duration)}</span>}
          {probe && <span>{sizeLabel(probe.size)}</span>}
          {probe && <span>{probe.videoCodec}{probe.hasAudio ? ' + ' + probe.audioCodec : ' без звука'}</span>}
          <label className="ml-auto flex items-center gap-1">
            Просмотр:
            <select
              value={previewRate}
              onChange={(e) => setPreviewRate(Number(e.target.value))}
              className="rounded border border-ws bg-ws-editor px-1 py-0.5 text-[11px] text-ws"
            >
              {[0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4].map((rate) => (
                <option key={rate} value={rate}>{rate}×</option>
              ))}
            </select>
          </label>
          {!ffmpeg && <span className="text-ws-warning">ffmpeg не найден — просмотр есть, правка недоступна</span>}
        </div>

        <div className="grid gap-3 md:grid-cols-2">
          <Field label={'Начало: ' + clock(start)}>
            <input type="range" min={0} max={probe?.duration || 0} step={0.1} value={start} onChange={(e) => setStart(Number(e.target.value))} className="w-full accent-[var(--ws-accent)]" />
          </Field>
          <Field label={'Конец: ' + clock(end)}>
            <input type="range" min={0} max={probe?.duration || 0} step={0.1} value={end} onChange={(e) => setEnd(Number(e.target.value))} className="w-full accent-[var(--ws-accent)]" />
          </Field>
        </div>

        {/* Скорость: всё видео или отдельные фрагменты */}
        <div className="space-y-2 rounded-lg border border-ws bg-ws-elevated p-3">
          <div className="flex flex-wrap items-center gap-2 text-[11px] text-ws">
            <Gauge size={13} className="text-ws-accent" />
            <span className="font-medium">Скорость всего видео:</span>
            <select
              value={speed}
              onChange={(e) => setSpeed(Number(e.target.value))}
              className="rounded border border-ws bg-ws-editor px-1.5 py-0.5 text-[11px] text-ws"
            >
              {[0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4].map((rate) => (
                <option key={rate} value={rate}>{rate}×</option>
              ))}
            </select>
            <span className="text-ws-faint">(1× — без изменений)</span>
          </div>

          <div className="flex flex-wrap items-end gap-2">
            <Field label="Фрагмент: с (сек)">
              <Input
                type="number"
                min={0}
                max={probe?.duration || 0}
                step={0.1}
                value={speedStart}
                onChange={(e) => setSpeedStart(Number(e.target.value))}
              />
            </Field>
            <Field label="по (сек)">
              <Input
                type="number"
                min={0}
                max={probe?.duration || 0}
                step={0.1}
                value={speedEnd}
                onChange={(e) => setSpeedEnd(Number(e.target.value))}
              />
            </Field>
            <Field label="скорость">
              <select
                value={speedValue}
                onChange={(e) => setSpeedValue(Number(e.target.value))}
                className="rounded border border-ws bg-ws-editor px-1.5 py-1 text-xs text-ws"
              >
                {[0.5, 0.75, 1.25, 1.5, 2, 3, 4].map((rate) => (
                  <option key={rate} value={rate}>{rate}×</option>
                ))}
              </select>
            </Field>
            <Button variant="secondary" onClick={addSpeedSegment} className="gap-1">
              <Scissors size={12} />
              Добавить фрагмент
            </Button>
          </div>

          {speedSegments.length > 0 && (
            <div className="flex flex-wrap gap-1.5">
              {speedSegments.map((segment, index) => (
                <span
                  key={`${segment.start}_${segment.end}_${index}`}
                  className="flex items-center gap-1 rounded border border-ws-accent/40 bg-ws-accent-soft px-1.5 py-0.5 text-[10px] text-ws-accent"
                >
                  {clock(segment.start)}–{clock(segment.end)} · {segment.speed}×
                  <button
                    type="button"
                    title="Убрать фрагмент"
                    onClick={() =>
                      setSpeedSegments((prev) => prev.filter((_, item) => item !== index))
                    }
                    className="rounded px-0.5 hover:bg-ws-hover"
                  >
                    ×
                  </button>
                </span>
              ))}
            </div>
          )}
          {(speed !== 1 || speedSegments.length > 0) && (
            <p className="text-[10px] text-ws-muted">
              Итог: {clock(
                Math.max(
                  0,
                  (end > start ? end - start : probe?.duration || 0) /
                    speed
                )
              )}{' '}
              (без учёта ускоренных фрагментов)
            </p>
          )}
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button onClick={() => setStripAudio((v) => !v)} className={'flex items-center gap-1 rounded border px-2 py-1 text-[11px] ' + (stripAudio ? 'border-ws-accent bg-ws-accent-soft text-ws-accent' : 'border-ws text-ws-muted')}>
            <VolumeX size={12} /> Убрать звук
          </button>
          <button onClick={() => setCompress((v) => !v)} className={'flex items-center gap-1 rounded border px-2 py-1 text-[11px] ' + (compress ? 'border-ws-accent bg-ws-accent-soft text-ws-accent' : 'border-ws text-ws-muted')}>
            <Gauge size={12} /> Сжать (CRF 20)
          </button>
          <Button variant="primary" disabled={!ffmpeg || Boolean(busy)} onClick={() => void apply()}>
            <Scissors size={12} /> {busy === 'ffmpeg' ? 'Обработка…' : 'Сохранить копию'}
          </Button>
          <Button variant="secondary" disabled={Boolean(busy)} onClick={() => void openPublish()}>
            <Upload size={12} /> На YouTube
          </Button>
        </div>

        {publishOpen && (
          <div className="space-y-3 rounded-lg border border-ws bg-ws-elevated p-3">
            {!connected ? (
              <p className="text-xs text-ws-warning">Сначала подключите Google в проводнике. Для YouTube при следующем входе нужно согласие на загрузку видео — нажмите «Подключить» ещё раз.</p>
            ) : (
              <>
                <Field label="Название"><Input value={meta.title} maxLength={100} onChange={(e) => setMeta({ ...meta, title: e.target.value })} /></Field>
                <Field label="Описание"><Textarea rows={4} value={meta.description} onChange={(e) => setMeta({ ...meta, description: e.target.value })} /></Field>
                <div className="grid gap-3 md:grid-cols-2">
                  <Field label="Доступ">
                    <Select value={meta.privacyStatus} onChange={(e) => setMeta({ ...meta, privacyStatus: e.target.value as typeof meta.privacyStatus })}>
                      <option value="public">Открытый доступ</option>
                      <option value="unlisted">Доступ по ссылке</option>
                      <option value="private">Ограниченный доступ</option>
                    </Select>
                  </Field>
                  <Field label="Категория">
                    <Select value={meta.categoryId} onChange={(e) => setMeta({ ...meta, categoryId: e.target.value })}>
                      {categories.map((c) => <option key={c.id} value={c.id}>{c.title}</option>)}
                    </Select>
                  </Field>
                  <Field label="Теги через запятую"><Input value={meta.tags} onChange={(e) => setMeta({ ...meta, tags: e.target.value })} /></Field>
                  <Field label="Плейлист">
                    <Select value={meta.playlistId} onChange={(e) => setMeta({ ...meta, playlistId: e.target.value })}>
                      <option value="">Не добавлять</option>
                      {playlists.map((p) => <option key={p.id} value={p.id}>{p.title}</option>)}
                    </Select>
                  </Field>
                  <Field label="Язык">
                    <Select value={meta.defaultLanguage} onChange={(e) => setMeta({ ...meta, defaultLanguage: e.target.value })}>
                      <option value="ru">Русский</option>
                      <option value="en">English</option>
                      <option value="uk">Українська</option>
                    </Select>
                  </Field>
                  <Field label="Запланировать публикацию">
                    <Input type="datetime-local" value={meta.publishAt} onChange={(e) => setMeta({ ...meta, publishAt: e.target.value })} />
                  </Field>
                </div>
                <label className="flex items-center gap-2 text-xs text-ws"><input type="checkbox" checked={meta.madeForKids} onChange={(e) => setMeta({ ...meta, madeForKids: e.target.checked })} /> Это видео для детей</label>
                <label className="flex items-center gap-2 text-xs text-ws"><input type="checkbox" checked={meta.embeddable} onChange={(e) => setMeta({ ...meta, embeddable: e.target.checked })} /> Разрешить встраивание</label>
                <label className="flex items-center gap-2 text-xs text-ws">
                  <input type="checkbox" checked={meta.license === 'creativeCommon'} onChange={(e) => setMeta({ ...meta, license: e.target.checked ? 'creativeCommon' : 'youtube' })} />
                  Лицензия Creative Commons
                </label>
                {progress > 0 && busy === 'youtube' && (
                  <div className="h-1.5 overflow-hidden rounded bg-ws-panel"><div className="h-full bg-ws-accent" style={{ width: Math.round(progress * 100) + '%' }} /></div>
                )}
                <div className="flex justify-end gap-2">
                  <Button variant="ghost" onClick={() => setPublishOpen(false)}>Отмена</Button>
                  <Button variant="primary" disabled={busy === 'youtube'} onClick={() => void publish()}>{busy === 'youtube' ? 'Загрузка…' : 'Опубликовать'}</Button>
                </div>
              </>
            )}
          </div>
        )}
      </div>
    </div>
  );
};
