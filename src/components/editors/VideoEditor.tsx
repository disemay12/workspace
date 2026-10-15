import React, { useEffect, useMemo, useState } from 'react';
import { Clapperboard, Gauge, Scissors, Upload, VolumeX } from 'lucide-react';
import { OpenTab } from '../../types';
import { Button, Field, Input, Select, Textarea } from '../ui';
import { gdriveStatus } from '../../services/shareService';

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
  const [probe, setProbe] = useState<Probe | null>(null);
  const [ffmpeg, setFfmpeg] = useState<string | null>(null);
  const [start, setStart] = useState(0);
  const [end, setEnd] = useState(0);
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
    });
    return () => { stop = true; };
  }, [tab.path]);

  const apply = async () => {
    setBusy('ffmpeg');
    try {
      const result = await window.electronAPI!.videoEdit({
        filePath: tab.path,
        start,
        end: end > start ? end : undefined,
        stripAudio,
        compress,
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
        <video src={src} controls className="h-full w-full bg-black" />
      </div>
      <div className="max-h-[46%] space-y-3 overflow-auto border-t border-ws bg-ws-surface p-3">
        <div className="flex flex-wrap items-center gap-3 text-[11px] text-ws-muted">
          <span className="flex items-center gap-1 text-ws"><Clapperboard size={13} /> {tab.name}</span>
          {probe && <span>{probe.width}×{probe.height}</span>}
          {probe && <span>{clock(probe.duration)}</span>}
          {probe && <span>{sizeLabel(probe.size)}</span>}
          {probe && <span>{probe.videoCodec}{probe.hasAudio ? ' + ' + probe.audioCodec : ' без звука'}</span>}
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
