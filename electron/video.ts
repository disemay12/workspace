/**
 * Просмотр и правка видео через ffmpeg.
 *
 * Файл не читается в память целиком: плеер берёт его по media://,
 * а обрезка, снятие звука и сжатие запускают ffmpeg как отдельный процесс.
 * Если ffmpeg не установлен, операции сообщают об этом, а просмотр остаётся.
 */
import { app, ipcMain, protocol } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import { spawn } from 'child_process';
import { Readable } from 'stream';

export const VIDEO_EXTENSIONS = [
  'mp4', 'm4v', 'mov', 'webm', 'mkv', 'avi', 'wmv', 'flv', 'f4v', 'mpeg', 'mpg',
  'mpe', 'mp2', 'm2v', 'ts', 'mts', 'm2ts', 'vob', 'ogv', 'ogg', '3gp', '3g2',
  'asf', 'rm', 'rmvb', 'divx', 'qt',
];

const MIME: Record<string, string> = {
  mp4: 'video/mp4',
  m4v: 'video/mp4',
  mov: 'video/quicktime',
  webm: 'video/webm',
  mkv: 'video/x-matroska',
  avi: 'video/x-msvideo',
  wmv: 'video/x-ms-wmv',
  mpeg: 'video/mpeg',
  mpg: 'video/mpeg',
  ogv: 'video/ogg',
  ogg: 'video/ogg',
  '3gp': 'video/3gpp',
  ts: 'video/mp2t',
  pdf: 'application/pdf',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  bmp: 'image/bmp',
};

export function videoMime(ext: string): string {
  return MIME[ext.toLowerCase().replace(/^\./, '')] || 'application/octet-stream';
}

function which(name: string): string | null {
  const exe = process.platform === 'win32' ? name + '.exe' : name;
  const roots = [
    process.env.FFMPEG_PATH || '',
    path.join(process.cwd(), 'tools', 'ffmpeg', 'bin', exe),
    path.join(process.cwd(), 'ffmpeg', 'bin', exe),
    'C:\\ffmpeg\\bin\\' + exe,
    'C:\\Program Files\\ffmpeg\\bin\\' + exe,
  ];
  for (const candidate of roots) {
    if (candidate && fs.existsSync(candidate) && fs.statSync(candidate).isFile()) return candidate;
  }
  const folders = (process.env.PATH || '').split(path.delimiter);
  for (const folder of folders) {
    const full = path.join(folder, exe);
    if (fs.existsSync(full)) return full;
  }
  return null;
}

function run(bin: string, args: string[]): Promise<{ code: number; stderr: string }> {
  return new Promise((resolve) => {
    const child = spawn(bin, args, { windowsHide: true });
    let stderr = '';
    child.stderr.on('data', (chunk) => {
      stderr += String(chunk);
      if (stderr.length > 8000) stderr = stderr.slice(-8000);
    });
    child.on('error', (err) => resolve({ code: 1, stderr: err.message }));
    child.on('close', (code) => resolve({ code: code ?? 1, stderr }));
  });
}

/** Запущенные кодирования: по ним работает «Отменить». */
const running = new Map<string, import('child_process').ChildProcess>();

/**
 * ffmpeg с разбором `-progress pipe:1`: строка `out_time_us` даёт текущую
 * позицию в микросекундах, по ней считается процент выполнения.
 */
function runWithProgress(
  bin: string,
  args: string[],
  durationSec: number,
  key: string,
  onProgress?: (percent: number) => void
): Promise<{ code: number; stderr: string }> {
  return new Promise((resolve) => {
    const child = spawn(bin, args, { windowsHide: true });
    running.set(key, child);
    let stderr = '';
    let stdout = '';
    child.stderr.on('data', (chunk) => {
      stderr += String(chunk);
      if (stderr.length > 8000) stderr = stderr.slice(-8000);
    });
    child.stdout.on('data', (chunk) => {
      stdout += String(chunk);
      const lines = stdout.split(/\r?\n/);
      stdout = lines.pop() ?? '';
      for (const line of lines) {
        const match = /^out_time_(?:us|ms)=(\d+)/.exec(line);
        if (!match) continue;
        // Оба поля ffmpeg пишет в микросекундах.
        const seconds = Number(match[1]) / 1_000_000;
        if (durationSec > 0 && onProgress) {
          onProgress(Math.max(0, Math.min(99, Math.round((seconds / durationSec) * 100))));
        }
      }
    });
    const finish = (code: number) => {
      running.delete(key);
      resolve({ code, stderr });
    };
    child.on('error', (err) => finish(1));
    child.on('close', (code) => finish(code ?? 1));
  });
}

/**
 * Фильтр масштаба кадра: приближение при экспорте.
 * `scale` > 1 увеличивает кадр и обрезает его по центру (сдвиг x/y в долях).
 */
export function buildZoomFilter(zoom: { scale?: number; x?: number; y?: number } | undefined): string {
  const scale = Math.min(4, Math.max(1, Number(zoom?.scale) || 1));
  if (scale <= 1.0001) return '';
  const ox = Math.max(-1, Math.min(1, Number(zoom?.x) || 0)).toFixed(3);
  const oy = Math.max(-1, Math.min(1, Number(zoom?.y) || 0)).toFixed(3);
  const factor = scale.toFixed(4);
  return (
    `scale=trunc(iw*${factor}/2)*2:trunc(ih*${factor}/2)*2,` +
    `crop=trunc(iw/${factor}/2)*2:trunc(ih/${factor}/2)*2:` +
    `(iw-ow)/2*(1+${ox}):(ih-oh)/2*(1+${oy})`
  );
}

function parseDuration(stderr: string): number {
  const match = stderr.match(/Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)/);
  if (!match) return 0;
  return Number(match[1]) * 3600 + Number(match[2]) * 60 + Number(match[3]);
}

async function probe(filePath: string): Promise<{
  duration: number;
  width: number;
  height: number;
  hasAudio: boolean;
  videoCodec: string;
  audioCodec: string;
  size: number;
}> {
  const stat = await fs.promises.stat(filePath);
  const ffprobe = which('ffprobe');
  const empty = { duration: 0, width: 0, height: 0, hasAudio: false, videoCodec: '', audioCodec: '', size: stat.size };
  if (!ffprobe) return empty;
  const result = await run(ffprobe, [
    '-v', 'error',
    '-show_entries', 'format=duration:stream=codec_type,codec_name,width,height',
    '-of', 'json',
    filePath,
  ]);
  try {
    const data = JSON.parse(result.stderr && result.stderr.trim().startsWith('{') ? result.stderr : await fs.promises.readFile(filePath, 'utf8').then(() => '{}'));
    void data;
  } catch {
    /* ffprobe пишет JSON в stdout, а run собирает только stderr — ниже отдельный запуск */
  }
  const json = await capture(ffprobe, [
    '-v', 'error',
    '-show_entries', 'format=duration:stream=codec_type,codec_name,width,height',
    '-of', 'json',
    filePath,
  ]);
  try {
    const parsed = JSON.parse(json.stdout || '{}') as {
      format?: { duration?: string };
      streams?: { codec_type?: string; codec_name?: string; width?: number; height?: number }[];
    };
    const video = (parsed.streams || []).find((s) => s.codec_type === 'video');
    const audio = (parsed.streams || []).find((s) => s.codec_type === 'audio');
    return {
      duration: Number(parsed.format?.duration || 0),
      width: video?.width || 0,
      height: video?.height || 0,
      hasAudio: Boolean(audio),
      videoCodec: video?.codec_name || '',
      audioCodec: audio?.codec_name || '',
      size: stat.size,
    };
  } catch {
    return { ...empty, duration: parseDuration(json.stderr) };
  }
}

function capture(bin: string, args: string[]): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((resolve) => {
    const child = spawn(bin, args, { windowsHide: true });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk) => { stdout += String(chunk); });
    child.stderr.on('data', (chunk) => { stderr += String(chunk); });
    child.on('error', (err) => resolve({ code: 1, stdout, stderr: err.message }));
    child.on('close', (code) => resolve({ code: code ?? 1, stdout, stderr }));
  });
}

export interface VideoEditRequest {
  filePath: string;
  start?: number;
  end?: number;
  stripAudio?: boolean;
  compress?: boolean;
  outputPath?: string;
  /** Ускорение всего видео (1 — без изменений). */
  speed?: number;
  /** Ускорение отдельных фрагментов (в секундах исходника). */
  speedSegments?: { start: number; end: number; speed: number }[];
  /** Длительность и наличие звука нужны для построения фильтров. */
  duration?: number;
  hasAudio?: boolean;
  /** Масштаб кадра при экспорте: 1 — без изменений. */
  zoom?: { scale?: number; x?: number; y?: number };
}

function atempoChain(speed: number): string {
  let value = Math.max(0.25, Math.min(4, speed));
  const filters: string[] = [];
  while (value > 2) {
    filters.push('atempo=2');
    value /= 2;
  }
  while (value < 0.5) {
    filters.push('atempo=0.5');
    value /= 0.5;
  }
  filters.push(`atempo=${value.toFixed(4)}`);
  return filters.join(',');
}

function speedFilterGraph(
  req: VideoEditRequest,
  trimStart: number,
  trimEnd: number,
  speed: number,
  segments: { start: number; end: number; speed: number }[],
  withAudio: boolean,
  zoomFilter = ''
): string {
  const pieces: { start: number; end: number; speed: number }[] = [];
  if (segments.length === 0) {
    pieces.push({ start: 0, end: trimEnd - trimStart, speed });
  } else {
    let cursor = trimStart;
    for (const segment of segments) {
      const segStart = Math.max(trimStart, Math.min(trimEnd, segment.start));
      const segEnd = Math.max(segStart, Math.min(trimEnd, segment.end));
      if (segEnd <= cursor) continue;
      if (segStart > cursor) {
        pieces.push({ start: cursor - trimStart, end: segStart - trimStart, speed });
      }
      pieces.push({ start: Math.max(cursor, segStart) - trimStart, end: segEnd - trimStart, speed: segment.speed });
      cursor = segEnd;
    }
    if (cursor < trimEnd) {
      pieces.push({ start: cursor - trimStart, end: trimEnd - trimStart, speed });
    }
  }
  if (pieces.length === 0) pieces.push({ start: 0, end: trimEnd - trimStart, speed });

  const filters: string[] = [];
  const concatInputs: string[] = [];
  pieces.forEach((piece, index) => {
    const vLabel = `v${index}`;
    const aLabel = `a${index}`;
    filters.push(
      `[0:v]trim=start=${piece.start.toFixed(3)}:end=${piece.end.toFixed(
        3
      )},setpts=(PTS-STARTPTS)/${piece.speed.toFixed(4)},setsar=1${
        zoomFilter ? ',' + zoomFilter : ''
      }[${vLabel}]`
    );
    if (withAudio) {
      filters.push(
        `[0:a]atrim=start=${piece.start.toFixed(3)}:end=${piece.end.toFixed(
          3
        )},asetpts=PTS-STARTPTS,${atempoChain(piece.speed)}[${aLabel}]`
      );
      concatInputs.push(`[${vLabel}][${aLabel}]`);
    } else {
      concatInputs.push(`[${vLabel}]`);
    }
  });
  filters.push(
    `${concatInputs.join('')}concat=n=${pieces.length}:v=1:a=${withAudio ? 1 : 0}[vout]${
      withAudio ? '[aout]' : ''
    }`
  );
  return filters.join(';');
}

async function editVideo(
  req: VideoEditRequest,
  onProgress?: (percent: number) => void
): Promise<{ success: boolean; outputPath?: string; error?: string }> {
  const ffmpeg = which('ffmpeg');
  if (!ffmpeg) {
    return {
      success: false,
      error: 'ffmpeg не найден. Установите его и добавьте в PATH либо положите в C:\\ffmpeg\\bin.',
    };
  }
  const start = Math.max(0, Number(req.start) || 0);
  const end = Number(req.end) || 0;
  if (end > 0 && end <= start) return { success: false, error: 'Конец обрезки должен быть позже начала' };

  const speed = Math.max(0.25, Math.min(4, Number(req.speed) || 1));
  const duration = Number(req.duration) || 0;
  const trimEnd = end > start ? end : duration > 0 ? duration : 0;
  const segments = (req.speedSegments || [])
    .map((segment) => ({
      start: Math.max(0, Number(segment.start) || 0),
      end: Math.max(0, Number(segment.end) || 0),
      speed: Math.max(0.25, Math.min(4, Number(segment.speed) || 1)),
    }))
    .filter((segment) => segment.end > segment.start)
    .sort((a, b) => a.start - b.start);
  const hasSpeedChange = speed !== 1 || segments.length > 0;
  const withAudio = !req.stripAudio && req.hasAudio !== false;
  const zoomFilter = buildZoomFilter(req.zoom);
  const hasZoom = zoomFilter.length > 0;

  const ext = path.extname(req.filePath).toLowerCase() || '.mp4';
  const base = path.basename(req.filePath, ext);
  const dir = path.dirname(req.filePath);
  const suffix = hasZoom
    ? '.zoom'
    : hasSpeedChange
      ? '.speed'
      : req.compress
        ? '.compressed'
        : req.stripAudio
          ? '.muted'
          : '.cut';
  const output =
    req.outputPath && req.outputPath !== req.filePath
      ? req.outputPath
      : path.join(dir, base + suffix + (hasSpeedChange || hasZoom ? '.mp4' : ext === '.webm' ? '.webm' : '.mp4'));

  const args = ['-y'];
  if (start > 0) args.push('-ss', start.toFixed(3));
  args.push('-i', req.filePath);
  if (trimEnd > start) args.push('-t', (trimEnd - start).toFixed(3));

  if (hasSpeedChange) {
    const graph = speedFilterGraph(
      req,
      start,
      trimEnd > start ? trimEnd : start + (duration || 0),
      speed,
      segments,
      withAudio,
      zoomFilter
    );
    args.push('-filter_complex', graph, '-map', '[vout]');
    if (withAudio) args.push('-map', '[aout]');
    if (withAudio) args.push('-c:a', 'aac', '-b:a', '192k');
    else args.push('-an');
    args.push(
      '-c:v',
      'libx264',
      '-crf',
      '20',
      '-preset',
      'medium',
      '-pix_fmt',
      'yuv420p',
      '-movflags',
      '+faststart'
    );
  } else if (req.compress || hasZoom) {
    if (hasZoom) args.push('-vf', zoomFilter);
    args.push('-c:v', 'libx264', '-crf', '20', '-preset', 'slow');
    if (!withAudio) args.push('-an');
    else args.push('-c:a', 'aac', '-b:a', '160k');
    args.push('-movflags', '+faststart');
  } else {
    args.push('-c:v', 'copy');
    if (!withAudio) args.push('-an');
    else args.push('-c:a', 'copy');
  }
  args.push('-progress', 'pipe:1', '-nostats', output);
  const expected = trimEnd > start ? trimEnd - start : duration;
  const result = await runWithProgress(ffmpeg, args, expected, req.filePath, onProgress);
  if (result.code !== 0 || !fs.existsSync(output)) {
    const tail = result.stderr.split(/\r?\n/).slice(-6).join(' ').trim();
    return { success: false, error: tail || 'ffmpeg завершился с ошибкой' };
  }
  onProgress?.(100);
  return { success: true, outputPath: output };
}

export function registerVideo(): void {
  protocol.handle('media', (request) => {
    try {
      const url = new URL(request.url);
      const filePath = decodeURIComponent(url.pathname.replace(/^\//, ''));
      const normalized = process.platform === 'win32' && /^[A-Za-z]:/.test(filePath.slice(0, 2))
        ? filePath
        : filePath;
      // media://local/C:/path или media://local/C%3A/... — pathname после host
      const raw = decodeURIComponent(url.searchParams.get('path') || normalized);
      if (!raw || !fs.existsSync(raw)) {
        return new Response('not found', { status: 404 });
      }

      // Range-запросы обязательны для перемотки видео: без 206 Chromium
      // не может прыгнуть на произвольную позицию и перемотка «залипает».
      const stat = fs.statSync(raw);
      const mime = videoMime(path.extname(raw));
      const rangeHeader = request.headers.get('range');
      if (rangeHeader) {
        const match = /^bytes=(\d*)-(\d*)$/.exec(rangeHeader.trim());
        if (match) {
          const start = match[1] ? Number(match[1]) : 0;
          const end = match[2]
            ? Math.min(Number(match[2]), stat.size - 1)
            : stat.size - 1;
          if (start <= end && start < stat.size) {
            const stream = fs.createReadStream(raw, { start, end });
            return new Response(Readable.toWeb(stream) as ReadableStream, {
              status: 206,
              headers: {
                'Content-Type': mime,
                'Accept-Ranges': 'bytes',
                'Content-Length': String(end - start + 1),
                'Content-Range': `bytes ${start}-${end}/${stat.size}`,
              },
            });
          }
          return new Response('range not satisfiable', {
            status: 416,
            headers: { 'Content-Range': `bytes */${stat.size}` },
          });
        }
      }

      const stream = fs.createReadStream(raw);
      return new Response(Readable.toWeb(stream) as ReadableStream, {
        status: 200,
        headers: {
          'Content-Type': mime,
          'Accept-Ranges': 'bytes',
          'Content-Length': String(stat.size),
        },
      });
    } catch (err) {
      return new Response((err as Error).message, { status: 500 });
    }
  });

  ipcMain.handle('video:capabilities', () => ({
    ffmpeg: which('ffmpeg'),
    ffprobe: which('ffprobe'),
  }));

  ipcMain.handle('video:probe', async (_event, filePath: string) => {
    try {
      const info = await probe(filePath);
      return { success: true, ...info, mime: videoMime(path.extname(filePath)) };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  ipcMain.handle('video:edit', async (event, req: VideoEditRequest) => {
    try {
      return await editVideo(req, (percent) => {
        event.sender.send('video:progress', { filePath: req.filePath, percent });
      });
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  /** Отмена кодирования: убиваем процесс ffmpeg этого файла. */
  ipcMain.handle('video:cancel', (_event, filePath: string) => {
    const child = running.get(filePath);
    if (!child) return { success: false, error: 'Нет активной обработки' };
    try {
      child.kill();
      running.delete(filePath);
      return { success: true };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });
  ipcMain.handle('video:extractAudio', async (_event, filePath: string, req?: VideoAudioRequest) => {
    try {
      const ffmpeg = requireFfmpeg();
      const format = req?.format ?? 'mp3';
      const outputPath = req?.outputPath || uniqueOutput(filePath, format === 'mp3' ? '.mp3' : format === 'aac' ? '.m4a' : '.m4a', '-audio');
      const args = ['-y', '-hide_banner', '-loglevel', 'error', ...buildExtractAudioArgs(filePath, outputPath, format)];
      const result = await run(ffmpeg, args);
      if (result.code !== 0) return { success: false, error: result.stderr.trim().split('\n').pop() || 'ffmpeg error' };
      return { success: true, outputPath };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  ipcMain.handle('video:frames', async (_event, filePath: string, req?: VideoFramesRequest) => {
    try {
      const ffmpeg = requireFfmpeg();
      const info = await probe(filePath);
      const count = Math.max(1, Math.min(24, Number(req?.count) || 12));
      const at =
        Array.isArray(req?.at) && req.at.length > 0
          ? req.at.map((value) => Math.max(0, Number(value) || 0)).slice(0, 24)
          : Array.from({ length: count }, (_, index) =>
              Math.max(0, Math.min(info.duration || 0, ((info.duration || 0) * (index + 0.5)) / count))
            );
      const dir = req?.dir || path.dirname(filePath);
      fs.mkdirSync(dir, { recursive: true });
      const stem = path.basename(filePath, path.extname(filePath));
      const files: string[] = [];
      for (let index = 0; index < at.length; index += 1) {
        const out = path.join(dir, `${stem}-frame-${String(index + 1).padStart(2, '0')}.jpg`);
        const args = ['-y', '-hide_banner', '-loglevel', 'error', ...buildFramesArgs(filePath, out, at[index], req?.maxWidth)];
        const result = await run(ffmpeg, args);
        if (result.code !== 0) return { success: false, files, error: result.stderr.trim() || 'ffmpeg error' };
        files.push(out);
      }
      return { success: true, files };
    } catch (err) {
      return { success: false, files: [], error: (err as Error).message };
    }
  });

  ipcMain.handle('video:convert', async (event, filePath: string, req?: VideoConvertRequest) => {
    try {
      const ffmpeg = requireFfmpeg();
      const format = req?.format ?? 'mp4';
      const ext = format === 'webm' ? '.webm' : format === 'gif' ? '.gif' : '.mp4';
      const outputPath = req?.outputPath || uniqueOutput(filePath, ext, `-${format}`);
      const args = ['-y', '-hide_banner', '-nostats', ...buildConvertArgs(filePath, outputPath, format, req?.quality)];
      const result = await run(ffmpeg, args);
      if (result.code !== 0) return { success: false, error: result.stderr.trim().split('\n').pop() || 'ffmpeg error' };
      event.sender.send('video:progress', { filePath, percent: 100 });
      return { success: true, outputPath };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  ipcMain.handle('video:gif', async (_event, filePath: string, req?: VideoGifRequest) => {
    try {
      const ffmpeg = requireFfmpeg();
      const outputPath = req?.outputPath || uniqueOutput(filePath, '.gif', '-clip');
      const start = Math.max(0, Number(req?.start) || 0);
      const duration = Math.max(0.2, Number(req?.duration) || 3);
      const args = ['-y', '-hide_banner', '-loglevel', 'error', ...buildGifArgs(filePath, outputPath, start, duration, req?.fps, req?.width)];
      const result = await run(ffmpeg, args);
      if (result.code !== 0) return { success: false, error: result.stderr.trim().split('\n').pop() || 'ffmpeg error' };
      return { success: true, outputPath };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });
}

/* ---------- Чистые построители аргументов (проверяются в v11-check) ---------- */

export type VideoAudioFormat = 'mp3' | 'aac' | 'copy';
export type VideoConvertFormat = 'mp4' | 'webm' | 'gif';

export interface VideoAudioRequest {
  format?: VideoAudioFormat;
  outputPath?: string;
}

export interface VideoFramesRequest {
  count?: number;
  at?: number[];
  maxWidth?: number;
  dir?: string;
}

export interface VideoConvertRequest {
  format?: VideoConvertFormat;
  quality?: 'low' | 'medium' | 'high';
  outputPath?: string;
}

export interface VideoGifRequest {
  start?: number;
  duration?: number;
  fps?: number;
  width?: number;
  outputPath?: string;
}

export function buildExtractAudioArgs(
  source: string,
  output: string,
  format: VideoAudioFormat = 'mp3'
): string[] {
  if (format === 'copy') return ['-i', source, '-vn', '-c:a', 'copy', output];
  if (format === 'aac') return ['-i', source, '-vn', '-c:a', 'aac', '-b:a', '192k', output];
  return ['-i', source, '-vn', '-c:a', 'libmp3lame', '-q:a', '2', output];
}

export function buildFramesArgs(source: string, output: string, at: number, maxWidth?: number): string[] {
  const width = Math.max(64, Math.min(1920, Number(maxWidth) || 480));
  return [
    '-ss',
    String(Math.max(0, at)),
    '-i',
    source,
    '-frames:v',
    '1',
    '-vf',
    `scale=${width}:-2:flags=lanczos`,
    '-q:v',
    '3',
    output,
  ];
}

export function buildConvertArgs(
  source: string,
  output: string,
  format: VideoConvertFormat = 'mp4',
  quality: 'low' | 'medium' | 'high' = 'medium'
): string[] {
  const crf = quality === 'high' ? '20' : quality === 'low' ? '30' : '24';
  if (format === 'webm') {
    return ['-i', source, '-c:v', 'libvpx-vp9', '-crf', crf, '-b:v', '0', '-c:a', 'libopus', output];
  }
  if (format === 'gif') {
    return [
      '-i',
      source,
      '-vf',
      'fps=12,scale=480:-1:flags=lanczos,split[a][b];[a]palettegen[p];[b][p]paletteuse',
      '-loop',
      '0',
      output,
    ];
  }
  return ['-i', source, '-c:v', 'libx264', '-crf', crf, '-preset', 'medium', '-c:a', 'aac', '-b:a', '160k', output];
}

export function buildGifArgs(
  source: string,
  output: string,
  start: number,
  duration: number,
  fps = 12,
  width = 480
): string[] {
  const safeFps = Math.max(5, Math.min(30, Math.round(Number(fps) || 12)));
  const safeWidth = Math.max(120, Math.min(1280, Math.round(Number(width) || 480)));
  return [
    '-ss',
    String(Math.max(0, start)),
    '-t',
    String(Math.max(0.2, duration)),
    '-i',
    source,
    '-vf',
    `fps=${safeFps},scale=${safeWidth}:-1:flags=lanczos,split[a][b];[a]palettegen[p];[b][p]paletteuse`,
    '-loop',
    '0',
    output,
  ];
}

export function mediaUrl(filePath: string): string {
  return 'media://local/?path=' + encodeURIComponent(filePath);
}

/** ffmpeg обязателен для новых операций: без него — понятная ошибка. */
function requireFfmpeg(): string {
  const ffmpeg = which('ffmpeg');
  if (!ffmpeg) throw new Error('ffmpeg не найден в PATH. Установите ffmpeg или добавьте его в PATH.');
  return ffmpeg;
}

/** Свободное имя рядом с исходником: name-edit.mp4, name-edit 2.mp4… */
function uniqueOutput(filePath: string, ext: string, suffix: string): string {
  const dir = path.dirname(filePath);
  const stem = path.basename(filePath, path.extname(filePath));
  let candidate = path.join(dir, `${stem}${suffix}${ext}`);
  let index = 2;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${stem}${suffix} ${index}${ext}`);
    index += 1;
  }
  return candidate;
}
