import React, { useEffect, useState } from 'react';
import { Cloud, ExternalLink, Loader2, ShieldCheck, Unplug } from 'lucide-react';
import { Button, Field, Input, Modal } from '../ui';
import { gdriveConnect, gdriveGetConfig, gdriveSetConfig, gdriveStatus, gdriveDisconnect, GDriveStatus } from '../../services/shareService';

export const GoogleDriveModal: React.FC<{ open: boolean; onClose: () => void; onNotify: (message: string) => void }> = ({ open, onClose, onNotify }) => {
  const [clientId, setClientId] = useState('');
  const [clientSecret, setClientSecret] = useState('');
  const [status, setStatus] = useState<GDriveStatus>({ configured: false, connected: false });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  const reload = async () => {
    const [cfg, current] = await Promise.all([gdriveGetConfig(), gdriveStatus()]);
    setClientId(cfg.clientId); setClientSecret(cfg.clientSecret); setStatus(current);
  };
  useEffect(() => { if (open) { setError(''); void reload(); } }, [open]);

  const save = async () => {
    await gdriveSetConfig(clientId.trim(), clientSecret.trim());
    await reload();
    onNotify('Настройки Google Drive сохранены');
  };
  const connect = async () => {
    setBusy(true); setError('');
    try {
      await save();
      const result = await gdriveConnect();
      if (!result.success) throw new Error(result.error || 'Не удалось подключить Google Drive');
      await reload(); onNotify('Google Drive подключён' + (result.email ? ': ' + result.email : ''));
    } catch (e) { setError((e as Error).message); }
    finally { setBusy(false); }
  };

  return <Modal open={open} onClose={onClose} title="Google Drive" icon={<Cloud size={15} className="text-ws-accent" />} width="max-w-xl" footer={<Button variant="ghost" onClick={onClose}>Закрыть</Button>}>
    <div className="space-y-4">
      <div className="flex items-start gap-2 rounded-lg border border-ws bg-ws-elevated p-3 text-[11px] leading-relaxed text-ws-muted">
        <ShieldCheck size={14} className="mt-0.5 shrink-0 text-ws-success" />
        <div>OAuth выполняется через браузер Google. Создайте OAuth Client ID типа Desktop app в Google Cloud Console, включите Google Drive API. Callback: <code className="select-text">http://127.0.0.1:42813</code>. Client Secret сохраняется локально в каталоге данных приложения.</div>
      </div>
      <a className="inline-flex items-center gap-1 text-[11px] text-ws-accent hover:underline" href="https://console.cloud.google.com/apis/credentials" target="_blank" rel="noreferrer">Google Cloud Credentials <ExternalLink size={11}/></a>
      <Field label="OAuth Client ID"><Input value={clientId} onChange={e=>setClientId(e.target.value)} placeholder="...apps.googleusercontent.com" spellCheck={false}/></Field>
      <Field label="Client Secret"><Input type="password" value={clientSecret} onChange={e=>setClientSecret(e.target.value)} placeholder="Client Secret" spellCheck={false}/></Field>
      {status.connected && <div className="rounded border border-emerald-500/30 bg-emerald-500/10 p-2 text-xs text-ws-success">Подключено{status.email ? ': ' + status.email : ''}</div>}
      {error && <div className="rounded border border-red-500/30 bg-red-500/10 p-2 text-xs text-ws-danger">{error}</div>}
      <div className="flex flex-wrap gap-2">
        <Button variant="primary" disabled={busy || !clientId.trim() || !clientSecret.trim()} onClick={() => void connect()}>{busy ? <Loader2 size={13} className="animate-spin"/> : <Cloud size={13}/>}Подключить Google</Button>
        <Button variant="secondary" onClick={() => void save()}>Сохранить настройки</Button>
        {status.connected && <Button variant="danger" onClick={async()=>{await gdriveDisconnect(); await reload(); onNotify('Google Drive отключён');}}><Unplug size={13}/>Отключить</Button>}
      </div>
    </div>
  </Modal>;
};
