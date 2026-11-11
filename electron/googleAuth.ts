/**
 * OAuth 2.0 Google: отдельные согласия для Диска и YouTube.
 *
 * Почему раздельно: `youtube.upload` — чувствительный scope, и при общем
 * согласии неверифицированное приложение ломает даже обычную выгрузку на Диск.
 * Пользователь подключает ровно то, что собирается использовать; Drive живёт
 * на `drive.file` (доступ только к файлам, созданным приложением) и не требует
 * верификации, YouTube включается отдельной кнопкой.
 *
 * Токены шифруются `safeStorage` (DPAPI в Windows) и хранятся только в главном
 * процессе; в браузер уходит лишь статус.
 */
import { app, safeStorage, shell } from 'electron';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import * as http from 'http';

export type GoogleService = 'drive' | 'youtube';

const PREFERRED_PORT = 42813;

const SCOPES: Record<GoogleService, string[]> = {
  drive: [
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid',
    'email',
  ],
  youtube: [
    'https://www.googleapis.com/auth/youtube.upload',
    'https://www.googleapis.com/auth/youtube.readonly',
  ],
};

interface TokenSet {
  refreshToken?: string;
  accessToken?: string;
  expiry?: number;
  scope?: string;
}

interface GoogleStore {
  clientId?: string;
  clientSecret?: string;
  email?: string;
  tokens: Partial<Record<GoogleService, TokenSet>>;
}

function storePath(): string {
  return path.join(app.getPath('userData'), 'google.json');
}

/** Прозрачное шифрование значений токенов. */
function encrypt(value: string): string {
  try {
    if (safeStorage.isEncryptionAvailable()) {
      return 'enc:' + safeStorage.encryptString(value).toString('base64');
    }
  } catch {
    // Нет доступа к DPAPI — сохраняем как есть, но помечаем.
  }
  return 'raw:' + value;
}

function decrypt(value: string): string {
  if (value.startsWith('enc:')) {
    try {
      return safeStorage.decryptString(Buffer.from(value.slice(4), 'base64'));
    } catch {
      return '';
    }
  }
  return value.startsWith('raw:') ? value.slice(4) : value;
}

export function readGoogleStore(): GoogleStore {
  try {
    const raw = JSON.parse(fs.readFileSync(storePath(), 'utf-8')) as GoogleStore & {
      tokens?: Partial<Record<GoogleService, Record<string, string>>>;
    };
    const tokens: Partial<Record<GoogleService, TokenSet>> = {};
    for (const service of ['drive', 'youtube'] as GoogleService[]) {
      const stored = raw.tokens?.[service];
      if (!stored) continue;
      const next: TokenSet = {};
      if (typeof stored.refreshToken === 'string' && stored.refreshToken) {
        next.refreshToken = decrypt(stored.refreshToken);
      }
      if (typeof stored.accessToken === 'string' && stored.accessToken) {
        next.accessToken = decrypt(stored.accessToken);
      }
      if (typeof stored.expiry === 'number') next.expiry = stored.expiry;
      if (typeof stored.scope === 'string') next.scope = stored.scope;
      tokens[service] = next;
    }
    return {
      clientId: raw.clientId,
      clientSecret: raw.clientSecret,
      email: typeof raw.email === 'string' ? raw.email : undefined,
      tokens,
    };
  } catch {
    return { tokens: {} };
  }
}

function writeGoogleStore(patch: Partial<GoogleStore>): GoogleStore {
  const current = readGoogleStore();
  const next: GoogleStore = {
    ...current,
    ...patch,
    tokens: { ...current.tokens, ...(patch.tokens ?? {}) },
  };
  const serialized = {
    clientId: next.clientId,
    clientSecret: next.clientSecret,
    email: next.email,
    tokens: Object.fromEntries(
      Object.entries(next.tokens).map(([service, token]) => [
        service,
        {
          ...(token?.refreshToken ? { refreshToken: encrypt(token.refreshToken) } : {}),
          ...(token?.accessToken ? { accessToken: encrypt(token.accessToken) } : {}),
          expiry: token?.expiry,
          scope: token?.scope,
        },
      ])
    ),
  };
  try {
    fs.mkdirSync(path.dirname(storePath()), { recursive: true });
    fs.writeFileSync(storePath(), JSON.stringify(serialized, null, 2), 'utf-8');
  } catch {
    // Каталог недоступен — операцию не срываем, но токен не переживёт перезапуск.
  }
  return next;
}

export function setGoogleConfig(clientId: string, clientSecret: string): void {
  writeGoogleStore({ clientId: clientId.trim(), clientSecret: clientSecret.trim() });
}

export function googleConfig(): { clientId: string; clientSecret: string } {
  const store = readGoogleStore();
  return { clientId: store.clientId ?? '', clientSecret: store.clientSecret ?? '' };
}

export function isConfigured(): boolean {
  const { clientId, clientSecret } = googleConfig();
  return Boolean(clientId && clientSecret);
}

/* ============================================================
   Ошибки Google → понятный текст
   ============================================================ */

const ERROR_MAP: { test: RegExp; message: string }[] = [
  {
    test: /redirect_uri_mismatch/i,
    message:
      'redirect_uri_mismatch: тип OAuth-клиента должен быть «Desktop app». Если клиент создан как «Web application», добавьте http://127.0.0.1:42813 в Authorized redirect URIs.',
  },
  {
    test: /invalid_client|unauthorized_client/i,
    message:
      'invalid_client: Client ID или Client Secret не подходят. Скопируйте значения заново из Google Cloud Console (Credentials → OAuth client).',
  },
  {
    test: /org_internal/i,
    message:
      'org_internal: приложение доступно только внутри организации Google Workspace. Тип пользователя (User type) должен быть External.',
  },
  {
    test: /access_denied|admin_policy_enforced/i,
    message:
      'access_denied: аккаунт не добавлен в Test users проекта или администратор запретил доступ. OAuth consent screen → Test users → добавьте свой адрес.',
  },
  {
    test: /invalid_grant/i,
    message:
      'invalid_grant: доступ отозван или истёк. В статусе «Testing» refresh-token живёт 7 дней — подключите заново.',
  },
  {
    test: /youtubeSignupRequired/i,
    message:
      'youtubeSignupRequired: у аккаунта нет канала YouTube. Создайте канал на youtube.com и повторите подключение.',
  },
  {
    test: /quotaExceeded|rateLimitExceeded/i,
    message:
      'quotaExceeded: исчерпана квота YouTube Data API на сегодня. Подождите до полуночи по тихоокеанскому времени или запросите увеличение квоты.',
  },
  {
    test: /uploadLimitExceeded/i,
    message: 'uploadLimitExceeded: аккаунт исчерпал дневной лимит загрузок YouTube (обычно 6–15 видео).',
  },
  {
    test: /audience|madeForKids/i,
    message:
      'YouTube требует явно указать аудиторию видео (для детей или нет) — включите или выключите этот флаг и повторите.',
  },
  {
    test: /API has not been used in project|accessNotConfigured/i,
    message:
      'API не включён в проекте Google Cloud: откройте «APIs & Services → Library» и включите Google Drive API и YouTube Data API v3.',
  },
  {
    test: /insufficientPermissions|insufficient_scope/i,
    message:
      'Недостаточно прав: подключите нужный сервис отдельной кнопкой (у Диска и YouTube отдельные согласия).',
  },
  {
    test: /ECONNREFUSED|ENOTFOUND|ETIMEDOUT|fetch failed/i,
    message: 'Нет связи с Google. Проверьте интернет, VPN или прокси и повторите.',
  },
];

/** Текст ошибки Google для пользователя. */
export function mapGoogleError(input: unknown): string {
  const raw =
    typeof input === 'string'
      ? input
      : input instanceof Error
        ? input.message
        : JSON.stringify(input);
  for (const entry of ERROR_MAP) {
    if (entry.test.test(raw)) return entry.message;
  }
  return raw.slice(0, 400);
}

/* ============================================================
   PKCE и адрес авторизации
   ============================================================ */

export function createVerifier(): string {
  return crypto.randomBytes(48).toString('base64url');
}

export function challengeFor(verifier: string): string {
  return crypto.createHash('sha256').update(verifier).digest('base64url');
}

export function buildAuthUrl(options: {
  clientId: string;
  service: GoogleService;
  redirectUri: string;
  state: string;
  codeChallenge: string;
}): string {
  return (
    'https://accounts.google.com/o/oauth2/v2/auth?' +
    new URLSearchParams({
      client_id: options.clientId,
      redirect_uri: options.redirectUri,
      response_type: 'code',
      scope: SCOPES[options.service].join(' '),
      access_type: 'offline',
      prompt: 'consent',
      include_granted_scopes: 'true',
      state: options.state,
      code_challenge: options.codeChallenge,
      code_challenge_method: 'S256',
    }).toString()
  );
}

/* ============================================================
   Loopback-сервер
   ============================================================ */

const HTML_OK =
  '<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>WorkSpace</title></head>' +
  '<body style="font-family:Segoe UI,sans-serif;background:#0b0f17;color:#e2e8f0;display:flex;' +
  'align-items:center;justify-content:center;height:100vh;margin:0"><div style="text-align:center">' +
  '<h2>Готово</h2><p style="color:#94a3b8">Авторизация Google завершена. Вернитесь в WorkSpace — ' +
  'это окно можно закрыть.</p></div></body></html>';

function htmlError(message: string): string {
  return (
    '<!doctype html><html lang="ru"><head><meta charset="utf-8"></head>' +
    '<body style="font-family:Segoe UI,sans-serif;background:#0b0f17;color:#f87171;display:flex;' +
    'align-items:center;justify-content:center;height:100vh;margin:0"><div style="text-align:center">' +
    '<h2>Ошибка авторизации</h2><p>' +
    message.slice(0, 500) +
    '</p></div></body></html>'
  );
}

interface LoopbackResult {
  code: string;
  redirectUri: string;
}

/**
 * Поднимает loopback-сервер и открывает браузер.
 * Пути кроме `/` игнорируются (favicon и прочее), поэтому посторонний запрос
 * браузера не прерывает ожидание кода.
 */
function awaitAuthorizationCode(
  authUrlBuilder: (redirectUri: string) => string
): Promise<LoopbackResult> {
  return new Promise<LoopbackResult>((resolve, reject) => {
    let settled = false;
    let server: http.Server | null = null;
    let timer: NodeJS.Timeout | undefined;

    const finish = (fn: () => void) => {
      if (settled) return;
      settled = true;
      if (timer) clearTimeout(timer);
      try {
        server?.close();
      } catch {
        // сервер уже закрыт
      }
      fn();
    };

    const handle = (req: http.IncomingMessage, res: http.ServerResponse) => {
      const url = new URL(req.url || '/', 'http://127.0.0.1');
      if (url.pathname !== '/' && url.pathname !== '/oauth2callback') {
        res.writeHead(204);
        res.end();
        return;
      }
      const code = url.searchParams.get('code');
      const error = url.searchParams.get('error');
      const state = url.searchParams.get('state') ?? '';

      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      if (code) {
        res.end(HTML_OK);
        finish(() => resolve({ code, redirectUri: redirectUriOf(req) }));
        return;
      }
      if (error) {
        res.end(htmlError(mapGoogleError(error + ' ' + (url.searchParams.get('error_description') ?? ''))));
        finish(() => reject(new Error(mapGoogleError(error))));
        return;
      }
      // Посторонний запрос без кода — отвечаем и ждём дальше.
      res.end(htmlError('Код авторизации не получен'));
      void state;
    };

    let boundPort = 0;
    const redirectUriOf = (req?: http.IncomingMessage): string => {
      void req;
      return `http://127.0.0.1:${boundPort}`;
    };

    server = http.createServer(handle);

    const start = (port: number, fallback: boolean) => {
      const current = server!;
      current.once('error', (err: NodeJS.ErrnoException) => {
        if (fallback && (err.code === 'EADDRINUSE' || err.code === 'EACCES')) {
          // Порт занят (например, прошлой авторизацией) — берём свободный.
          try {
            current.close();
          } catch {
            // уже закрыт
          }
          server = http.createServer(handle);
          start(0, false);
          return;
        }
        finish(() => reject(err));
      });
      current.listen(port, '127.0.0.1', () => {
        boundPort = (current.address() as { port: number }).port;
        const authUrl = authUrlBuilder(`http://127.0.0.1:${boundPort}`);
        void shell.openExternal(authUrl);
      });
    };

    timer = setTimeout(
      () => finish(() => reject(new Error('Истекло время ожидания авторизации (5 минут)'))),
      5 * 60 * 1000
    );

    start(PREFERRED_PORT, true);
  });
}

/* ============================================================
   Обмен кода и обновление токена
   ============================================================ */

async function exchangeCode(
  service: GoogleService,
  code: string,
  redirectUri: string,
  verifier: string
): Promise<void> {
  const { clientId, clientSecret } = googleConfig();
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      code,
      client_id: clientId,
      client_secret: clientSecret,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
      code_verifier: verifier,
    }),
  });
  if (!res.ok) throw new Error(mapGoogleError(await res.text()));

  const data = (await res.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in?: number;
    scope?: string;
  };

  const store = readGoogleStore();
  const tokens = { ...store.tokens };
  tokens[service] = {
    accessToken: data.access_token,
    refreshToken: data.refresh_token ?? tokens[service]?.refreshToken,
    expiry: Date.now() + (data.expires_in ?? 3600) * 1000,
    scope: data.scope,
  };
  writeGoogleStore({ tokens });

  try {
    const info = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: 'Bearer ' + data.access_token },
    });
    if (info.ok) {
      const profile = (await info.json()) as { email?: string };
      if (profile.email) writeGoogleStore({ email: profile.email });
    }
  } catch {
    // Профиль не обязателен
  }
}

export interface GoogleConnectResult {
  success: boolean;
  email?: string;
  scopes?: string[];
  error?: string;
}

/** Подключение сервиса: браузер, loopback, PKCE. */
export async function connectService(service: GoogleService): Promise<GoogleConnectResult> {
  if (!isConfigured()) {
    return { success: false, error: 'Укажите OAuth Client ID и Client Secret приложения Google' };
  }
  const verifier = createVerifier();
  const challenge = challengeFor(verifier);
  const state = crypto.randomBytes(16).toString('hex');
  const { clientId } = googleConfig();

  try {
    const result = await awaitAuthorizationCode((redirectUri) =>
      buildAuthUrl({ clientId, service, redirectUri, state, codeChallenge: challenge })
    );
    await exchangeCode(service, result.code, result.redirectUri, verifier);
    const store = readGoogleStore();
    return {
      success: true,
      email: store.email,
      scopes: SCOPES[service],
    };
  } catch (err) {
    return { success: false, error: mapGoogleError(err) };
  }
}

/** Действующий access-token сервиса, с обновлением по refresh-token. */
export async function getAccessToken(service: GoogleService): Promise<string> {
  const store = readGoogleStore();
  const token = store.tokens[service];
  if (!token?.refreshToken) {
    throw new Error(
      service === 'youtube'
        ? 'YouTube не подключён: нажмите «Подключить YouTube» в настройках'
        : 'Google Drive не подключён: нажмите «Подключить Google»'
    );
  }
  if (token.accessToken && token.expiry && Date.now() < token.expiry - 60_000) {
    return token.accessToken;
  }

  const { clientId, clientSecret } = googleConfig();
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      refresh_token: token.refreshToken,
      grant_type: 'refresh_token',
    }),
  });
  if (!res.ok) {
    const text = await res.text();
    if (/invalid_grant/i.test(text)) {
      // Доступ отозван: чистим сервис, чтобы интерфейс честно показал «не подключён».
      const tokens = { ...store.tokens };
      delete tokens[service];
      writeGoogleStore({ tokens });
    }
    throw new Error(mapGoogleError(text));
  }
  const data = (await res.json()) as { access_token: string; expires_in?: number };
  const tokens = { ...readGoogleStore().tokens };
  tokens[service] = {
    ...tokens[service],
    accessToken: data.access_token,
    expiry: Date.now() + (data.expires_in ?? 3600) * 1000,
  };
  writeGoogleStore({ tokens });
  return data.access_token;
}

export interface GoogleServiceStatus {
  connected: boolean;
  email?: string;
  scopes?: string;
}

export interface GoogleStatus {
  configured: boolean;
  drive: GoogleServiceStatus;
  youtube: GoogleServiceStatus;
  email?: string;
}

export function googleStatus(): GoogleStatus {
  const store = readGoogleStore();
  const build = (service: GoogleService): GoogleServiceStatus => ({
    connected: Boolean(store.tokens[service]?.refreshToken),
    email: store.email,
    scopes: store.tokens[service]?.scope,
  });
  return {
    configured: isConfigured(),
    drive: build('drive'),
    youtube: build('youtube'),
    email: store.email,
  };
}

/** Отзыв доступа: и в Google, и локально. */
export async function disconnectService(service: GoogleService): Promise<void> {
  const store = readGoogleStore();
  const token = store.tokens[service]?.refreshToken;
  if (token) {
    try {
      await fetch('https://oauth2.googleapis.com/revoke?token=' + encodeURIComponent(token), {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      });
    } catch {
      // Отзыв не критичен для локальной очистки
    }
  }
  const tokens = { ...store.tokens };
  delete tokens[service];
  writeGoogleStore({ tokens, ...(service === 'drive' ? { email: undefined } : {}) });
}

export interface GoogleDiagnose {
  configured: boolean;
  tokenValid: boolean;
  driveApi: 'ok' | 'no-access' | 'error' | 'not-connected';
  youtubeApi: 'ok' | 'no-access' | 'error' | 'not-connected';
  quotaBytes?: number;
  channelTitle?: string;
  error?: string;
}

/** Проверка «Проверить подключение»: дергает реальные API. */
export async function diagnose(): Promise<GoogleDiagnose> {
  const result: GoogleDiagnose = {
    configured: isConfigured(),
    tokenValid: false,
    driveApi: 'not-connected',
    youtubeApi: 'not-connected',
  };
  const status = googleStatus();

  if (status.drive.connected) {
    try {
      const token = await getAccessToken('drive');
      result.tokenValid = true;
      const res = await fetch(
        'https://www.googleapis.com/drive/v3/about?fields=storageQuota(limit)&prettyPrint=false',
        { headers: { Authorization: 'Bearer ' + token } }
      );
      if (res.ok) {
        const data = (await res.json()) as { storageQuota?: { limit?: string } };
        result.driveApi = 'ok';
        if (data.storageQuota?.limit) result.quotaBytes = Number(data.storageQuota.limit);
      } else {
        result.driveApi = res.status === 403 ? 'no-access' : 'error';
        result.error = mapGoogleError(await res.text());
      }
    } catch (err) {
      result.error = mapGoogleError(err);
      result.driveApi = 'error';
    }
  }

  if (status.youtube.connected) {
    try {
      const token = await getAccessToken('youtube');
      const res = await fetch(
        'https://www.googleapis.com/youtube/v3/channels?part=snippet&mine=true',
        { headers: { Authorization: 'Bearer ' + token } }
      );
      if (res.ok) {
        const data = (await res.json()) as { items?: { snippet?: { title?: string } }[] };
        const title = data.items?.[0]?.snippet?.title;
        if (title) {
          result.youtubeApi = 'ok';
          result.channelTitle = title;
        } else {
          // Токен принят, но канала нет — это отдельный случай в гайде.
          result.youtubeApi = 'no-access';
          result.error = mapGoogleError('youtubeSignupRequired');
        }
      } else {
        result.youtubeApi = res.status === 403 ? 'no-access' : 'error';
        result.error = mapGoogleError(await res.text());
      }
    } catch (err) {
      result.youtubeApi = 'error';
      result.error = mapGoogleError(err);
    }
  }

  return result;
}
