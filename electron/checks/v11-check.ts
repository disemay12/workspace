/**
 * Проверки логики версии 1.1: сайтовые фото-хостинги (Япикс), зум видео,
 * OAuth Google. Все проверки — чистые функции: сеть, окна и app.whenReady
 * не нужны, поэтому скрипт выполняется синхронно и сразу завершает процесс.
 */
const { app } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const ROOT = (() => {
  let dir = __dirname;
  for (let i = 0; i < 4; i++) {
    if (fs.existsSync(path.join(dir, 'package.json'))) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return path.join(__dirname, '..');
})();

const out: Record<string, unknown> = {};

function check(name: string, value: boolean, detail?: string) {
  out[name] = value;
  if (!value && detail) out[name + 'Detail'] = detail;
}

function run(): void {
  try {
    const hosts = require('./imageHost');
    const video = require('./video');
    const auth = require('./googleAuth');

    /* ---------- Фото-хостинги ---------- */
    const batches = hosts.batchFiles(Array.from({ length: 25 }, (_, i: number) => i), 10);
    check(
      'hostsBatchesSizes',
      batches.length === 3 &&
        batches[0].length === 10 &&
        batches[1].length === 10 &&
        batches[2].length === 5,
      JSON.stringify(batches.map((batch: number[]) => batch.length))
    );
    check('hostsBatchesEmpty', hosts.batchFiles([], 10).length === 0);
    check(
      'hostsBatchesSingle',
      hosts.batchFiles([1], 10).length === 1 && hosts.batchFiles([1], 10)[0].length === 1
    );

    const yapx = hosts.IMAGE_HOSTS.yapx;
    check('yapxConfigured', Boolean(yapx) && yapx.url === 'https://yapx.ru/upload');
    check(
      'yapxLimits',
      yapx.perAlbum === 10 && yapx.perAlbumAuthorized === 20 && yapx.maxFileMb === 3,
      JSON.stringify({
        perAlbum: yapx.perAlbum,
        authorized: yapx.perAlbumAuthorized,
        mb: yapx.maxFileMb,
      })
    );
    check(
      'yapxFormats',
      ['apng', 'bmp', 'gif', 'jpg', 'jpeg', 'png'].every((ext: string) => yapx.formats.includes(ext)) &&
        !yapx.formats.includes('webp')
    );
    check('yapxSelectors', yapx.selectors[0] === '#upload-input' && yapx.supportsPaste === true);
    check(
      'yapxResultPattern',
      new RegExp(yapx.resultPattern).test('https://yapx.ru/album/AbC123') &&
        !new RegExp(yapx.resultPattern).test('https://yapx.ru/upload')
    );

    // Системный TEMP на этой машине может быть закрыт для записи — берём папку проекта.
    const tempDir = path.join(ROOT, '.v11-check-tmp-' + process.pid);
    fs.mkdirSync(tempDir, { recursive: true });
    try {
      const pngPath = path.join(tempDir, 'ok.png');
      const webpPath = path.join(tempDir, 'no.webp');
      const bigPath = path.join(tempDir, 'big.png');
      fs.writeFileSync(pngPath, Buffer.alloc(1024));
      fs.writeFileSync(webpPath, Buffer.alloc(1024));
      fs.writeFileSync(bigPath, Buffer.alloc(4 * 1024 * 1024));
      check('hostsPngAccepted', hosts.hostFileProblem(pngPath, yapx) === null);
      check('hostsWebpRejected', typeof hosts.hostFileProblem(webpPath, yapx) === 'string');
      check('hostsBigRejected', typeof hosts.hostFileProblem(bigPath, yapx) === 'string');
    } finally {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }

    /* ---------- Зум видео ---------- */
    check(
      'videoZoomOff',
      video.buildZoomFilter({ scale: 1 }) === '' && video.buildZoomFilter(undefined) === ''
    );
    const zoom2 = video.buildZoomFilter({ scale: 2 });
    check(
      'videoZoomTwo',
      zoom2.indexOf('scale=trunc(iw*2.0000') === 0 &&
        zoom2.indexOf('crop=') > 0 &&
        zoom2.indexOf('(iw-ow)/2') > 0,
      zoom2
    );
    const zoomClamped = video.buildZoomFilter({ scale: 10 });
    check('videoZoomClamped', zoomClamped.indexOf('4.0000') > 0, zoomClamped);
    const offset = video.buildZoomFilter({ scale: 1.5, x: 5, y: -5 });
    check(
      'videoZoomOffsetClamped',
      offset.indexOf('(1+1.000)') > 0 && offset.indexOf('(1+-1.000)') > 0,
      offset
    );

    /* ---------- Операции над видео (1.2) ---------- */
    const audioMp3 = video.buildExtractAudioArgs('in.mp4', 'out.mp3', 'mp3');
    check(
      'videoAudioMp3',
      audioMp3.join(' ').includes('-vn') && audioMp3.join(' ').includes('libmp3lame'),
      audioMp3.join(' ')
    );
    const audioCopy = video.buildExtractAudioArgs('in.mp4', 'out.m4a', 'copy');
    check(
      'videoAudioCopy',
      audioCopy.join(' ').includes('-c:a copy'),
      audioCopy.join(' ')
    );
    const frameArgs = video.buildFramesArgs('in.mp4', 'out.jpg', 4.5, 320);
    check(
      'videoFramesArgs',
      frameArgs.join(' ') === '-ss 4.5 -i in.mp4 -frames:v 1 -vf scale=320:-2:flags=lanczos -q:v 3 out.jpg',
      frameArgs.join(' ')
    );
    const convertArgs = video.buildConvertArgs('in.mp4', 'out.webm', 'webm', 'high');
    check(
      'videoConvertArgs',
      convertArgs.join(' ').includes('libvpx-vp9') && convertArgs.join(' ').includes('-crf 20'),
      convertArgs.join(' ')
    );
    const gifArgs = video.buildGifArgs('in.mp4', 'out.gif', 2, 5, 15, 640);
    check(
      'videoGifArgs',
      gifArgs.join(' ').includes('-ss 2') &&
        gifArgs.join(' ').includes('-t 5') &&
        gifArgs.join(' ').includes('fps=15') &&
        gifArgs.join(' ').includes('scale=640:-1'),
      gifArgs.join(' ')
    );

    /* ---------- Google OAuth ---------- */
    const verifier = auth.createVerifier();
    const challenge = auth.challengeFor(verifier);
    check(
      'pkceVerifierLength',
      verifier.length >= 43 && verifier.length <= 128,
      String(verifier.length)
    );
    check('pkceChallengeShape', /^[A-Za-z0-9_-]{43}$/.test(challenge), challenge);

    const driveUrl = auth.buildAuthUrl({
      clientId: 'test.apps.googleusercontent.com',
      service: 'drive',
      redirectUri: 'http://127.0.0.1:42813',
      state: 'abc123',
      codeChallenge: challenge,
    });
    check(
      'authUrlDrive',
      driveUrl.indexOf('https://accounts.google.com/o/oauth2/v2/auth?') === 0 &&
        driveUrl.indexOf('drive.file') > 0 &&
        driveUrl.indexOf('code_challenge_method=S256') > 0 &&
        driveUrl.indexOf('state=abc123') > 0 &&
        driveUrl.indexOf('access_type=offline') > 0 &&
        driveUrl.indexOf('youtube.upload') < 0
    );

    const youtubeUrl = auth.buildAuthUrl({
      clientId: 'test.apps.googleusercontent.com',
      service: 'youtube',
      redirectUri: 'http://127.0.0.1:42813',
      state: 'st',
      codeChallenge: challenge,
    });
    check(
      'authUrlYoutube',
      youtubeUrl.indexOf('youtube.upload') > 0 &&
        youtubeUrl.indexOf('drive.file') < 0 &&
        youtubeUrl.indexOf('include_granted_scopes=true') > 0
    );

    const mismatch = String(auth.mapGoogleError('redirect_uri_mismatch'));
    check('errorRedirectMismatch', mismatch.indexOf('Desktop') > 0, mismatch);
    check('errorTestingMode', String(auth.mapGoogleError('invalid_grant')).indexOf('7 дней') > 0);
    check('errorQuota', String(auth.mapGoogleError('quotaExceeded')).indexOf('квот') > 0);
    check('errorChannel', String(auth.mapGoogleError('youtubeSignupRequired')).indexOf('канал') > 0);
    check('errorScope', String(auth.mapGoogleError('insufficientPermissions')).indexOf('согласия') > 0);
    check(
      'errorUnknownKept',
      String(auth.mapGoogleError('совсем новая ошибка')).indexOf('совсем новая ошибка') === 0
    );

    out.errors = [];
  } catch (err) {
    out.fatal = (err as Error).stack || (err as Error).message;
  }

  fs.writeFileSync(path.join(ROOT, 'v11-check.json'), JSON.stringify(out, null, 2), 'utf-8');
  app.exit(0);
}

run();
