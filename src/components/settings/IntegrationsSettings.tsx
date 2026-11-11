import React, { useEffect, useState } from 'react';
import {
  AlertTriangle,
  Check,
  CheckCircle2,
  ChevronDown,
  Cloud,
  Copy,
  ExternalLink,
  Loader2,
  RefreshCw,
  ShieldCheck,
  Unplug,
  Clapperboard,
} from 'lucide-react';
import { Button, Field, Input } from '../ui';
import { cn } from '../../lib/utils';
import {
  GDriveStatus,
  GoogleDiagnoseResult,
  GoogleFullStatus,
  getImgurClientId,
  googleConnect,
  googleDiagnose,
  googleDisconnectService,
  googleFullStatus,
  gdriveGetConfig,
  gdriveStatus,
  gdriveSetConfig,
  setImgurClientId,
} from '../../services/shareService';
import { usePreferences } from '../../state/preferences';

/**
 * Сервисы: Google Drive, YouTube и фото-хостинги.
 *
 * Мини-гайд повторяет путь в Google Cloud Console по шагам, а кнопка
 * «Проверить подключение» дергает реальные API и показывает, что именно
 * работает: токен, Drive API, YouTube Data API.
 */

const REDIRECT_URI = 'http://127.0.0.1:42813';

const FAQ: { question: string; answer: string }[] = [
  {
    question: 'redirect_uri_mismatch',
    answer:
      'OAuth-клиент должен быть типа «Desktop app» — тогда адрес возврата не нужно регистрировать. Если клиент создан как «Web application», добавьте http://127.0.0.1:42813 в Authorized redirect URIs.',
  },
  {
    question: '«Google не проверил это приложение»',
    answer:
      'Это нормально для собственного клиента. На экране согласия нажмите «Дополнительно» → «Перейти на (небезопасно)». Аккаунт должен быть добавлен в Test users проекта.',
  },
  {
    question: 'Доступ отвалился через неделю',
    answer:
      'В статусе OAuth consent screen «Testing» refresh-token живёт 7 дней. Для постоянной работы переведите приложение в статус «In production» (для Drive это не требует проверки).',
  },
  {
    question: 'org_internal / admin_policy_enforced',
    answer:
      'Проект ограничен организацией Google Workspace. В OAuth consent screen смените User type на External либо используйте личный аккаунт Google.',
  },
  {
    question: 'YouTube: «access_denied» или нет доступа к загрузке',
    answer:
      'Скоупы YouTube — чувствительные: пока приложение не верифицировано, работать смогут только аккаунты из Test users. Подключите YouTube отдельной кнопкой и используйте один из этих аккаунтов.',
  },
  {
    question: 'YouTube: видеозагрузка отклонена с quotaExceeded',
    answer:
      'Дневная квота YouTube Data API (10 000 единиц) тратится на загрузки и метаданные. Квота обновляется в полночь по тихоокеанскому времени; её можно увеличить запросом в Google Cloud.',
  },
  {
    question: 'youtubeSignupRequired',
    answer:
      'У аккаунта Google нет канала YouTube. Создайте канал на youtube.com, затем нажмите «Подключить YouTube» заново.',
  },
  {
    question: 'API не включён в проекте',
    answer:
      'В Google Cloud Console → APIs & Services → Library включите «Google Drive API» и «YouTube Data API v3». Без этого загрузки вернут 403 accessNotConfigured.',
  },
];

export const IntegrationsSettings: React.FC<{ onNotify: (message: string) => void }> = ({
  onNotify,
}) => {
  const [clientId, setClientId] = useState('');
  const [clientSecret, setClientSecret] = useState('');
  const [imgurId, setImgurId] = useState('');
  const [status, setStatus] = useState<GoogleFullStatus>({
    configured: false,
    drive: { connected: false },
    youtube: { connected: false },
  });
  const [legacyStatus, setLegacyStatus] = useState<GDriveStatus>({
    configured: false,
    connected: false,
  });
  const [busy, setBusy] = useState<'drive' | 'youtube' | 'diagnose' | 'save' | ''>('');
  const [error, setError] = useState('');
  const [diagnose, setDiagnose] = useState<GoogleDiagnoseResult | null>(null);
  const [copied, setCopied] = useState(false);
  const [faqOpen, setFaqOpen] = useState(false);
  const [preferences, updatePreferences] = usePreferences();

  const reload = async () => {
    const [cfg, full, legacy] = await Promise.all([
      gdriveGetConfig(),
      googleFullStatus(),
      gdriveStatus(),
    ]);
    setClientId(cfg.clientId);
    setClientSecret(cfg.clientSecret);
    setStatus(full);
    setLegacyStatus(legacy);
    setImgurId(getImgurClientId());
  };

  useEffect(() => {
    void reload();
  }, []);

  const saveConfig = async () => {
    setBusy('save');
    try {
      await gdriveSetConfig(clientId.trim(), clientSecret.trim());
      await reload();
      onNotify('Ключи Google сохранены');
    } finally {
      setBusy('');
    }
  };

  const connect = async (service: 'drive' | 'youtube') => {
    setBusy(service);
    setError('');
    setDiagnose(null);
    try {
      await gdriveSetConfig(clientId.trim(), clientSecret.trim());
      const result = await googleConnect(service);
      if (!result.success) throw new Error(result.error || 'Не удалось подключить Google');
      await reload();
      onNotify(
        (service === 'drive' ? 'Google Drive подключён' : 'YouTube подключён') +
          (result.email ? ': ' + result.email : '')
      );
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy('');
    }
  };

  const runDiagnose = async () => {
    setBusy('diagnose');
    setError('');
    try {
      const result = await googleDiagnose();
      setDiagnose(result);
      if (result?.error) setError(result.error);
    } finally {
      setBusy('');
    }
  };

  const copyRedirect = async () => {
    try {
      await navigator.clipboard.writeText(REDIRECT_URI);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1600);
    } catch {
      // буфер может быть недоступен
    }
  };

  const chip = (
    label: string,
    state: 'ok' | 'no-access' | 'error' | 'not-connected' | undefined
  ) => {
    const map = {
      ok: { text: 'работает', className: 'border-emerald-500/40 bg-emerald-500/10 text-ws-success' },
      'no-access': {
        text: 'нет доступа',
        className: 'border-amber-500/40 bg-amber-500/10 text-ws-warning',
      },
      error: { text: 'ошибка', className: 'border-red-500/40 bg-red-500/10 text-ws-danger' },
      'not-connected': { text: 'не подключено', className: 'border-ws text-ws-faint' },
    } as const;
    const value = map[state ?? 'not-connected'];
    return (
      <span className={cn('rounded border px-1.5 py-0.5 text-[10px]', value.className)}>
        {label}: {value.text}
      </span>
    );
  };

  return (
    <div className="space-y-4">
      {/* Шапка гайда */}
      <div className="rounded-lg border border-ws bg-ws-elevated p-3 text-[11px] leading-relaxed">
        <div className="mb-1.5 flex items-center gap-1.5 font-medium text-ws">
          <ShieldCheck size={13} className="text-ws-success" />
          Мини-гайд: свой OAuth-клиент Google
        </div>
        <ol className="list-inside list-decimal space-y-0.5 text-ws-muted">
          <li>
            <a
              className="text-ws-accent hover:underline"
              href="https://console.cloud.google.com/projectcreate"
              target="_blank"
              rel="noreferrer"
            >
              Создайте проект
            </a>{' '}
            в Google Cloud (или выберите существующий).
          </li>
          <li>
            Включите{' '}
            <a
              className="text-ws-accent hover:underline"
              href="https://console.cloud.google.com/apis/library/drive.googleapis.com"
              target="_blank"
              rel="noreferrer"
            >
              Google Drive API
            </a>{' '}
            и{' '}
            <a
              className="text-ws-accent hover:underline"
              href="https://console.cloud.google.com/apis/library/youtube.googleapis.com"
              target="_blank"
              rel="noreferrer"
            >
              YouTube Data API v3
            </a>
            .
          </li>
          <li>
            Credentials → Create credentials → <b>OAuth client ID</b> → тип{' '}
            <b>Desktop app</b>.
          </li>
          <li>
            OAuth consent screen: User type <b>External</b>, себя добавьте в <b>Test users</b>.
          </li>
          <li>Скопируйте Client ID и Client Secret в поля ниже и сохраните.</li>
          <li>Подключите Диск и/или YouTube отдельными кнопками — согласия независимы.</li>
        </ol>
        <div className="mt-2 flex flex-wrap items-center gap-2 text-[10px] text-ws-faint">
          <span>Адрес возврата (нужен только для клиента типа Web):</span>
          <code className="select-text rounded bg-ws-editor px-1.5 py-0.5">{REDIRECT_URI}</code>
          <button
            onClick={() => void copyRedirect()}
            className="flex items-center gap-1 rounded border border-ws px-1.5 py-0.5 hover:bg-ws-hover"
          >
            {copied ? <Check size={10} /> : <Copy size={10} />}
            {copied ? 'Скопировано' : 'Копировать'}
          </button>
          <a
            className="ml-auto inline-flex items-center gap-1 text-ws-accent hover:underline"
            href="https://console.cloud.google.com/apis/credentials"
            target="_blank"
            rel="noreferrer"
          >
            Credentials <ExternalLink size={10} />
          </a>
        </div>
      </div>

      <div className="grid gap-3 md:grid-cols-2">
        <Field label="Google OAuth Client ID" hint="Один клиент обслуживает и Диск, и YouTube.">
          <Input
            value={clientId}
            onChange={(e) => setClientId(e.target.value)}
            placeholder="....apps.googleusercontent.com"
            spellCheck={false}
          />
        </Field>
        <Field label="Google Client Secret" hint="Хранится локально в каталоге данных приложения.">
          <Input
            type="password"
            value={clientSecret}
            onChange={(e) => setClientSecret(e.target.value)}
            placeholder="GOCSPX-…"
            spellCheck={false}
          />
        </Field>
      </div>

      <div className="flex flex-wrap gap-2">
        <Button
          variant="primary"
          disabled={busy !== '' || !clientId.trim() || !clientSecret.trim()}
          onClick={() => void connect('drive')}
        >
          {busy === 'drive' ? <Loader2 size={13} className="animate-spin" /> : <Cloud size={13} />}
          Подключить Google Drive
        </Button>
        <Button
          variant="secondary"
          disabled={busy !== '' || !clientId.trim() || !clientSecret.trim()}
          onClick={() => void connect('youtube')}
        >
          {busy === 'youtube' ? (
            <Loader2 size={13} className="animate-spin" />
          ) : (
            <Clapperboard size={13} />
          )}
          Подключить YouTube
        </Button>
        <Button variant="secondary" disabled={busy === 'save'} onClick={() => void saveConfig()}>
          Сохранить ключи
        </Button>
        <Button
          variant="secondary"
          disabled={busy !== '' || !status.configured}
          onClick={() => void runDiagnose()}
        >
          {busy === 'diagnose' ? (
            <Loader2 size={13} className="animate-spin" />
          ) : (
            <RefreshCw size={13} />
          )}
          Проверить подключение
        </Button>
      </div>

      <div className="flex flex-wrap items-center gap-2 text-[11px]">
        {chip('Drive', diagnose?.driveApi ?? (status.drive.connected ? 'ok' : 'not-connected'))}
        {chip(
          'YouTube',
          diagnose?.youtubeApi ?? (status.youtube.connected ? 'ok' : 'not-connected')
        )}
        {status.email && <span className="text-ws-muted">Аккаунт: {status.email}</span>}
        {diagnose?.channelTitle && (
          <span className="text-ws-muted">Канал: {diagnose.channelTitle}</span>
        )}
        {typeof diagnose?.quotaBytes === 'number' && diagnose.quotaBytes > 0 && (
          <span className="text-ws-faint">
            Квота Диска: {(diagnose.quotaBytes / (1024 * 1024 * 1024)).toFixed(0)} ГБ
          </span>
        )}
      </div>

      {(status.drive.connected || status.youtube.connected) && (
        <div className="flex flex-wrap gap-2">
          {status.drive.connected && (
            <Button
              variant="danger"
              disabled={busy !== ''}
              onClick={async () => {
                await googleDisconnectService('drive');
                await reload();
                onNotify('Google Drive отключён');
              }}
            >
              <Unplug size={13} /> Отключить Drive
            </Button>
          )}
          {status.youtube.connected && (
            <Button
              variant="danger"
              disabled={busy !== ''}
              onClick={async () => {
                await googleDisconnectService('youtube');
                await reload();
                onNotify('YouTube отключён');
              }}
            >
              <Unplug size={13} /> Отключить YouTube
            </Button>
          )}
        </div>
      )}

      {legacyStatus.connected && !status.drive.connected && (
        <div className="rounded border border-ws bg-ws-elevated p-2 text-[11px] text-ws-muted">
          Обнаружено подключение старого формата. Нажмите «Подключить Google Drive», чтобы
          обновить согласия — файлы и ссылки останутся прежними.
        </div>
      )}

      {error && (
        <div className="flex items-start gap-2 rounded border border-red-500/30 bg-red-500/10 p-2 text-[11px] text-ws-danger">
          <AlertTriangle size={13} className="mt-0.5 shrink-0" />
          <span>{error}</span>
        </div>
      )}

      <button
        onClick={() => setFaqOpen((open) => !open)}
        className="flex w-full items-center justify-between rounded border border-ws bg-ws-elevated px-2.5 py-1.5 text-[11px] text-ws"
      >
        <span className="flex items-center gap-1.5">
          <CheckCircle2 size={12} className="text-ws-accent" />
          Частые ошибки и что делать
        </span>
        <ChevronDown
          size={12}
          className={cn('transition-transform', faqOpen && 'rotate-180')}
        />
      </button>
      {faqOpen && (
        <div className="space-y-1.5">
          {FAQ.map((item) => (
            <details key={item.question} className="rounded border border-ws bg-ws-surface p-2">
              <summary className="cursor-pointer text-[11px] text-ws">{item.question}</summary>
              <p className="mt-1 text-[11px] leading-relaxed text-ws-muted">{item.answer}</p>
            </details>
          ))}
        </div>
      )}

      <Field
        label="Imgur Client-ID"
        hint="Нужен для загрузки одного фото и альбома. Client Secret Imgur не используется."
      >
        <Input
          value={imgurId}
          onChange={(e) => setImgurId(e.target.value)}
          onBlur={() => setImgurClientId(imgurId)}
          placeholder="Client-ID из настроек приложения Imgur"
          spellCheck={false}
          className="font-mono"
        />
      </Field>

      <div className="rounded-lg border border-ws bg-ws-elevated p-3 text-[11px] leading-relaxed">
        <div className="mb-1 font-medium text-ws">Япикс (yapx.ru)</div>
        <p className="mb-2 text-ws-muted">
          Бесплатный фотохостинг без ключа: до 10 изображений в альбоме, 20 загрузок в час,
          3 МБ на файл. Форматы APNG, BMP, GIF, JPG, PNG — остальные пропускаются с отчётом.
          Большие наборы уходят несколькими альбомами, ссылки копируются в буфер.
        </p>
        <label className="flex items-center gap-2 text-ws">
          <input
            type="checkbox"
            checked={preferences.yapxAccountMode}
            onChange={(event) => updatePreferences({ yapxAccountMode: event.target.checked })}
          />
          Я вошёл в аккаунт Япикса (лимиты: 20 в альбоме, 30 в час)
        </label>
      </div>
    </div>
  );
};
