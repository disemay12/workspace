import React, { useEffect, useState } from 'react';
import { Cloud, ExternalLink, Loader2, ShieldCheck, Unplug } from 'lucide-react';
import { Button, Field, Input } from '../ui';
import {
  GDriveStatus,
  gdriveConnect,
  gdriveDisconnect,
  gdriveGetConfig,
  gdriveSetConfig,
  gdriveStatus,
  getImgurClientId,
  setImgurClientId,
} from '../../services/shareService';

/** Общие ключи Google Cloud и Imgur. YouTube использует тот же вход, что и Drive. */
export const IntegrationsSettings: React.FC<{ onNotify: (message: string) => void }> = ({ onNotify }) => {
  const [clientId, setClientId] = useState('');
  const [clientSecret, setClientSecret] = useState('');
  const [imgurId, setImgurId] = useState('');
  const [status, setStatus] = useState<GDriveStatus>({ configured: false, connected: false });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    setImgurId(getImgurClientId());
    void Promise.all([gdriveGetConfig(), gdriveStatus()]).then(([cfg, current]) => {
      setClientId(cfg.clientId);
      setClientSecret(cfg.clientSecret);
      setStatus(current);
    });
  }, []);

  const save = async () => {
    setImgurClientId(imgurId);
    await gdriveSetConfig(clientId.trim(), clientSecret.trim());
    setStatus(await gdriveStatus());
    onNotify('Настройки Imgur и Google сохранены');
  };

  const connect = async () => {
    setBusy(true);
    setError('');
    try {
      await save();
      const result = await gdriveConnect();
      if (!result.success) throw new Error(result.error || 'Не удалось подключить Google');
      setStatus(await gdriveStatus());
      onNotify('Google подключён' + (result.email ? ': ' + result.email : ''));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="space-y-4">
      <div className="flex items-start gap-2 rounded-lg border border-ws bg-ws-elevated p-3 text-[11px] leading-relaxed text-ws-muted">
        <ShieldCheck size={14} className="mt-0.5 shrink-0 text-ws-success" />
        <div>
          Один OAuth-клиент Google используется и для Диска, и для YouTube. В Google Cloud включите
          Google Drive API и YouTube Data API v3, создайте OAuth Client ID типа Desktop и добавьте
          callback <code className="select-text">http://127.0.0.1:42813</code>.
        </div>
      </div>
      <div className="flex flex-wrap gap-3 text-[11px]">
        <a className="inline-flex items-center gap-1 text-ws-accent hover:underline" href="https://console.cloud.google.com/apis/credentials" target="_blank" rel="noreferrer">Google Cloud Credentials <ExternalLink size={11} /></a>
        <a className="inline-flex items-center gap-1 text-ws-accent hover:underline" href="https://console.cloud.google.com/apis/library/youtube.googleapis.com" target="_blank" rel="noreferrer">YouTube Data API <ExternalLink size={11} /></a>
        <a className="inline-flex items-center gap-1 text-ws-accent hover:underline" href="https://api.imgur.com/oauth2/addclient" target="_blank" rel="noreferrer">Создать Imgur Client-ID <ExternalLink size={11} /></a>
      </div>

      <Field label="Google OAuth Client ID" hint="Один идентификатор для Google Drive и YouTube.">
        <Input value={clientId} onChange={(e) => setClientId(e.target.value)} placeholder="....apps.googleusercontent.com" spellCheck={false} />
      </Field>
      <Field label="Google Client Secret">
        <Input type="password" value={clientSecret} onChange={(e) => setClientSecret(e.target.value)} placeholder="Client Secret" spellCheck={false} />
      </Field>
      <Field label="Imgur Client-ID" hint="Нужен для загрузки одного фото и альбома. Client Secret Imgur не используется.">
        <Input value={imgurId} onChange={(e) => setImgurId(e.target.value)} placeholder="Client-ID из настроек приложения Imgur" spellCheck={false} className="font-mono" />
      </Field>

      {status.connected && <div className="rounded border border-emerald-500/30 bg-emerald-500/10 p-2 text-xs text-ws-success">Google подключён{status.email ? ': ' + status.email : ''}. Доступ действует для Диска и YouTube.</div>}
      {error && <div className="rounded border border-red-500/30 bg-red-500/10 p-2 text-xs text-ws-danger">{error}</div>}

      <div className="flex flex-wrap gap-2">
        <Button variant="primary" disabled={busy || !clientId.trim() || !clientSecret.trim()} onClick={() => void connect()}>
          {busy ? <Loader2 size={13} className="animate-spin" /> : <Cloud size={13} />} Подключить Google
        </Button>
        <Button variant="secondary" onClick={() => void save()}>Сохранить</Button>
        {status.connected && <Button variant="danger" onClick={async () => { await gdriveDisconnect(); setStatus(await gdriveStatus()); onNotify('Google отключён'); }}><Unplug size={13} /> Отключить</Button>}
      </div>
    </div>
  );
};
