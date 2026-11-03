import React, { useEffect, useState } from 'react';
import {
  Check,
  FolderOpen,
  Info,
  Loader2,
  Monitor,
  Plus,
  RefreshCw,
  ShieldCheck,
  Trash2,
  FileCode2,
  ExternalLink,
  Pin,
} from 'lucide-react';
import { Button } from '../ui';

interface Props {
  onNotify?: (message: string) => void;
}

export const SystemSettings: React.FC<Props> = ({ onNotify }) => {
  const [installed, setInstalled] = useState(false);
  const [exePath, setExePath] = useState('');
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [alwaysOnTop, setAlwaysOnTop] = useState(false);
  const [alwaysBusy, setAlwaysBusy] = useState(false);
  /** Настройка доступна только в настольной версии Electron. */
  const alwaysSupported = Boolean(window.electronAPI?.setAlwaysOnTop);
  const [shortcutInstalled, setShortcutInstalled] = useState(false);
  const [shortcutPath, setShortcutPath] = useState('');
  const [shortcutBusy, setShortcutBusy] = useState(false);

  const refreshStatus = async () => {
    if (!window.electronAPI?.getContextMenuStatus) {
      setLoading(false);
      return;
    }
    try {
      const res = await window.electronAPI.getContextMenuStatus();
      setInstalled(res.installed);
      setExePath(res.exePath);
      setError(res.error || '');
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void refreshStatus();
  }, []);

  useEffect(() => {
    if (!window.electronAPI?.getAlwaysOnTop) return;
    void window.electronAPI
      .getAlwaysOnTop()
      .then(setAlwaysOnTop)
      .catch(() => setAlwaysOnTop(false));
  }, []);

  useEffect(() => {
    if (!window.electronAPI?.getDesktopShortcutStatus) return;
    void window.electronAPI
      .getDesktopShortcutStatus()
      .then((res) => {
        setShortcutInstalled(res.installed);
        setShortcutPath(res.shortcutPath);
      })
      .catch(() => setShortcutInstalled(false));
  }, []);

  const createShortcut = async () => {
    if (!window.electronAPI?.createDesktopShortcut) return;
    setShortcutBusy(true);
    try {
      const res = await window.electronAPI.createDesktopShortcut();
      if (res.success) {
        setShortcutInstalled(true);
        if (res.shortcutPath) setShortcutPath(res.shortcutPath);
        onNotify?.('Ярлык WorkSpace создан на рабочем столе');
      } else {
        throw new Error(res.error || 'Не удалось создать ярлык');
      }
    } catch (e) {
      onNotify?.(`Ошибка: ${(e as Error).message}`);
    } finally {
      setShortcutBusy(false);
    }
  };

  const toggleAlwaysOnTop = async () => {
    if (!window.electronAPI?.setAlwaysOnTop) return;
    setAlwaysBusy(true);
    try {
      const res = await window.electronAPI.setAlwaysOnTop(!alwaysOnTop);
      setAlwaysOnTop(res.alwaysOnTop);
      onNotify?.(
        res.alwaysOnTop
          ? 'Окно закреплено поверх остальных приложений'
          : 'Закрепление окна отключено'
      );
    } catch (e) {
      onNotify?.(`Ошибка: ${(e as Error).message}`);
    } finally {
      setAlwaysBusy(false);
    }
  };

  const handleInstall = async () => {
    if (!window.electronAPI?.installContextMenu) return;
    setBusy(true);
    setError('');
    try {
      const res = await window.electronAPI.installContextMenu();
      if (res.success) {
        setInstalled(true);
        onNotify?.('Пункт «Редактировать в WorkSpace» добавлен в контекстное меню Windows');
      } else {
        throw new Error(res.error || 'Не удалось зарегистрировать пункт');
      }
    } catch (e) {
      const msg = (e as Error).message;
      setError(msg);
      onNotify?.(`Ошибка: ${msg}`);
    } finally {
      setBusy(false);
      void refreshStatus();
    }
  };

  const handleRemove = async () => {
    if (!window.electronAPI?.removeContextMenu) return;
    setBusy(true);
    setError('');
    try {
      const res = await window.electronAPI.removeContextMenu();
      if (res.success) {
        setInstalled(false);
        onNotify?.('Пункт WorkSpace удалён из контекстного меню Windows');
      } else {
        throw new Error(res.error || 'Не удалось удалить пункт');
      }
    } catch (e) {
      const msg = (e as Error).message;
      setError(msg);
      onNotify?.(`Ошибка: ${msg}`);
    } finally {
      setBusy(false);
      void refreshStatus();
    }
  };

  if (!window.electronAPI?.getContextMenuStatus) {
    return (
      <div className="rounded-lg border border-ws bg-ws-elevated p-4 text-xs text-ws-muted">
        Интеграция с контекстным меню Windows доступна только в настольной версии приложения WorkSpace.
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* Поверх всех окон */}
      <div className="flex items-start justify-between gap-3 rounded-lg border border-ws bg-ws-elevated p-3.5">
        <div className="flex items-start gap-2.5">
          <Pin size={16} className="mt-0.5 shrink-0 text-ws-accent" />
          <div className="space-y-1">
            <div className="text-xs font-semibold text-ws">Поверх всех окон</div>
            <p className="text-[11px] leading-relaxed text-ws-muted">
              Окно WorkSpace остаётся видимым над другими приложениями. Настройка
              сохраняется и применяется при следующем запуске.
            </p>
          </div>
        </div>
        <div className="flex shrink-0 flex-col items-end gap-1">
          <span
            className={`text-[11px] font-medium ${
              alwaysOnTop ? 'text-emerald-500' : 'text-ws-muted'
            }`}
          >
            {alwaysOnTop ? 'Включено' : 'Выключено'}
          </span>
          <button
            type="button"
            role="switch"
            aria-checked={alwaysOnTop}
            aria-label="Поверх всех окон"
            disabled={!alwaysSupported || alwaysBusy}
            onClick={() => void toggleAlwaysOnTop()}
            className={`relative h-6 w-11 rounded-full transition-colors disabled:opacity-50 ${
              alwaysOnTop ? 'bg-emerald-500' : 'bg-slate-400'
            }`}
          >
            <span
              className={`absolute top-1 h-4 w-4 rounded-full bg-white shadow transition-transform ${
                alwaysOnTop ? 'translate-x-[24px]' : 'translate-x-[4px]'
              }`}
            />
          </button>
          {!alwaysSupported && (
            <span className="text-[10px] text-ws-danger">Недоступно</span>
          )}
        </div>
      </div>

      {/* Ярлык на рабочем столе */}
      <div className="flex items-start justify-between gap-3 rounded-lg border border-ws bg-ws-elevated p-3.5">
        <div className="flex items-start gap-2.5">
          <Monitor size={16} className="mt-0.5 shrink-0 text-ws-accent" />
          <div className="space-y-1">
            <div className="text-xs font-semibold text-ws">Ярлык на рабочем столе</div>
            <p className="text-[11px] leading-relaxed text-ws-muted">
              Создаёт ярлык WorkSpace с иконкой приложения. Ярлык можно
              пересоздать после перемещения папки с программой.
            </p>
            {shortcutInstalled && shortcutPath && (
              <p className="break-all font-mono text-[10px] text-ws-faint">{shortcutPath}</p>
            )}
          </div>
        </div>
        <Button
          variant={shortcutInstalled ? 'secondary' : 'primary'}
          disabled={shortcutBusy || !window.electronAPI?.createDesktopShortcut}
          onClick={() => void createShortcut()}
          className="shrink-0 gap-1.5"
        >
          {shortcutBusy ? <Loader2 size={13} className="animate-spin" /> : <Plus size={13} />}
          {shortcutInstalled ? 'Обновить ярлык' : 'Создать ярлык'}
        </Button>
      </div>

      {/* Информационный баннер */}
      <div className="flex items-start gap-2.5 rounded-lg border border-ws bg-ws-elevated p-3 text-xs leading-relaxed text-ws-muted">
        <Monitor size={16} className="mt-0.5 shrink-0 text-ws-accent" />
        <div className="space-y-1">
          <div className="font-semibold text-ws">Контекстное меню Проводника Windows</div>
          <p className="text-[11px] text-ws-muted">
            Позволяет открывать любые файлы и проекты прямо из Проводника Windows кликом правой кнопки
            мыши (ПКМ). Регистрация выполняется в пользовательской ветке реестра и не требует прав администратора (UAC).
          </p>
        </div>
      </div>

      {/* Текущий статус */}
      <div
        className={`rounded-lg border p-3.5 transition-colors ${
          installed
            ? 'border-emerald-500/40 bg-emerald-500/10'
            : 'border-ws bg-ws-surface'
        }`}
      >
        <div className="flex items-start justify-between gap-3">
          <div className="space-y-1">
            <div className="flex items-center gap-2">
              <span
                className={`flex h-2 w-2 rounded-full ${
                  installed ? 'bg-emerald-400' : 'bg-ws-faint'
                }`}
              />
              <span className="text-xs font-semibold text-ws">
                {installed
                  ? 'Контекстное меню включено'
                  : 'Контекстное меню отключено'}
              </span>
            </div>

            <p className="text-[11px] text-ws-muted">
              {installed
                ? 'Пункт «Редактировать в WorkSpace» активен в контекстном меню Проводника Windows.'
                : 'Клик ПКМ по файлу или папке в Windows сейчас не содержит пункта WorkSpace.'}
            </p>
          </div>

          <div className="shrink-0">
            {installed ? (
              <span className="flex items-center gap-1 rounded bg-emerald-500/20 px-2 py-0.5 text-[11px] font-medium text-emerald-300">
                <Check size={11} /> Активно
              </span>
            ) : (
              <span className="rounded bg-ws-elevated px-2 py-0.5 text-[11px] text-ws-faint">
                Не добавлено
              </span>
            )}
          </div>
        </div>

        {/* Что добавляется */}
        <div className="mt-3 grid grid-cols-1 gap-2 border-t border-ws/40 pt-3 text-[11px] sm:grid-cols-2">
          <div className="flex items-start gap-2 rounded bg-ws-elevated/60 p-2">
            <FileCode2 size={13} className="mt-0.5 shrink-0 text-blue-400" />
            <div>
              <span className="font-medium text-ws">Клик ПКМ по файлу:</span>
              <p className="text-[10px] text-ws-muted mt-0.5">
                «Редактировать в WorkSpace» — открывает файл в соответствующем редакторе (код, Word, Excel, Markdown).
              </p>
            </div>
          </div>

          <div className="flex items-start gap-2 rounded bg-ws-elevated/60 p-2">
            <FolderOpen size={13} className="mt-0.5 shrink-0 text-amber-400" />
            <div>
              <span className="font-medium text-ws">Клик ПКМ по папке:</span>
              <p className="text-[10px] text-ws-muted mt-0.5">
                «Открыть в WorkSpace» — открывает папку как проект со всем деревом файлов.
              </p>
            </div>
          </div>
        </div>

        {/* Путь к исполняемому файлу */}
        {exePath && (
          <div className="mt-3 border-t border-ws/40 pt-2 text-[10px] text-ws-faint">
            Исполняемый файл:{' '}
            <span className="font-mono text-ws select-text break-all">{exePath}</span>
          </div>
        )}
      </div>

      {error && (
        <div className="rounded border border-red-500/30 bg-red-500/10 p-2 text-xs text-ws-danger">
          {error}
        </div>
      )}

      {/* Кнопки управления */}
      <div className="flex flex-wrap items-center gap-2 pt-1">
        {!installed ? (
          <Button
            variant="primary"
            disabled={busy || loading}
            onClick={() => void handleInstall()}
            className="gap-1.5"
          >
            {busy ? <Loader2 size={13} className="animate-spin" /> : <Plus size={13} />}
            Включить контекстное меню
          </Button>
        ) : (
          <>
            <Button
              variant="danger"
              disabled={busy || loading}
              onClick={() => void handleRemove()}
              className="gap-1.5"
            >
              {busy ? <Loader2 size={13} className="animate-spin" /> : <Trash2 size={13} />}
              Убрать контекстное меню
            </Button>
            <Button
              variant="secondary"
              disabled={busy || loading}
              onClick={() => void handleInstall()}
              title="Перезаписать путь к приложению в реестре, если папка была перемещена"
              className="gap-1.5"
            >
              <RefreshCw size={12} />
              Обновить путь
            </Button>
          </>
        )}
      </div>

      <div className="flex items-center gap-1.5 text-[10px] text-ws-faint pt-1">
        <ShieldCheck size={12} className="text-ws-success shrink-0" />
        <span>
          Запись осуществляется исключительно в раздел реестра HKEY_CURRENT_USER и удаляется в один клик.
        </span>
      </div>
    </div>
  );
};
