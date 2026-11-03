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
  withAudio: boolean
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
      )},setpts=(PTS-STARTPTS)/${piece.speed.toFixed(4)},setsar=1[${vLabel}]`
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

async function editVideo(req: VideoEditRequest): Promise<{ success: boolean; outputPath?: string; error?: string }> {
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

  const ext = path.extname(req.filePath).toLowerCase() || '.mp4';
  const base = path.basename(req.filePath, ext);
  const dir = path.dirname(req.filePath);
  const suffix = hasSpeedChange ? '.speed' : req.compress ? '.compressed' : req.stripAudio ? '.muted' : '.cut';
  const output =
    req.outputPath && req.outputPath !== req.filePath
      ? req.outputPath
      : path.join(dir, base + suffix + (hasSpeedChange ? '.mp4' : ext === '.webm' ? '.webm' : '.mp4'));

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
      withAudio
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
  } else if (req.compress) {
    args.push('-c:v', 'libx264', '-crf', '20', '-preset', 'slow');
    if (!withAudio) args.push('-an');
    else args.push('-c:a', 'aac', '-b:a', '160k');
    args.push('-movflags', '+faststart');
  } else {
    args.push('-c:v', 'copy');
    if (!withAudio) args.push('-an');
    else args.push('-c:a', 'copy');
  }
  args.push(output);
  const result = await run(ffmpeg, args);
  if (result.code !== 0 || !fs.existsSync(output)) {
    const tail = result.stderr.split(/\r?\n/).slice(-6).join(' ').trim();
    return { success: false, error: tail || 'ffmpeg завершился с ошибкой' };
  }
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

  ipcMain.handle('video:edit', async (_event, req: VideoEditRequest) => {
    try {
      return await editVideo(req);
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });
}

export function mediaUrl(filePath: string): string {
  return 'media://local/?path=' + encodeURIComponent(filePath);
}
