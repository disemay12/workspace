import React, { useEffect, useMemo, useState } from 'react';
import {
  Binary,
  Copy,
  ExternalLink,
  FolderOpen,
  Loader2,
  RefreshCw,
} from 'lucide-react';
import { OpenTab } from '../../types';
import { fileSystemService } from '../../services/fileSystemService';
import { Button } from '../ui';

interface Props {
  tab: OpenTab;
  onNotify: (message: string) => void;
}

const EXECUTABLE = /\.(exe|msi|com|scr|bat|cmd|ps1|vbs|js|jar)$/i;
const INITIAL_LIMIT = 16 * 1024;
const MAX_LIMIT = 512 * 1024;

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} Б`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} КБ`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} МБ`;
}

function toHex(value: number): string {
  return value.toString(16).padStart(2, '0').toUpperCase();
}

/** Просмотр любого файла: hex-дамп первых байт + действия Windows. */
export const BinaryViewer: React.FC<Props> = ({ tab, onNotify }) => {
  const [bytes, setBytes] = useState<Uint8Array>(new Uint8Array());
  const [size, setSize] = useState(0);
  const [limit, setLimit] = useState(INITIAL_LIMIT);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError('');
    void fileSystemService
      .readPreview(tab.path, limit)
      .then((res) => {
        if (cancelled) return;
        const binary = atob(res.base64);
        const next = new Uint8Array(binary.length);
        for (let index = 0; index < binary.length; index++) {
          next[index] = binary.charCodeAt(index);
        }
        setBytes(next);
        setSize(res.size);
      })
      .catch((err) => {
        if (!cancelled) setError((err as Error).message);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [tab.path, limit]);

  const rows = useMemo(() => {
    const result: { offset: number; hex: string[]; ascii: string }[] = [];
    for (let offset = 0; offset < bytes.length; offset += 16) {
      const slice = bytes.subarray(offset, offset + 16);
      const hex = Array.from(slice, toHex);
      while (hex.length < 16) hex.push('  ');
      const ascii = Array.from(slice, (value) =>
        value >= 32 && value < 127 ? String.fromCharCode(value) : '.'
      ).join('');
      result.push({ offset, hex, ascii });
    }
    return result;
  }, [bytes]);

  const openExternally = async () => {
    const name = tab.name;
    if (EXECUTABLE.test(tab.path)) {
      const confirmed = window.confirm(
        `Запустить исполняемый файл?\n\n${tab.path}\n\nWorkSpace откроет его средствами Windows.`
      );
      if (!confirmed) return;
    }
    try {
      await fileSystemService.openExternally(tab.path);
      onNotify(`Открыто в системе: ${name}`);
    } catch (err) {
      onNotify(`Не удалось открыть: ${(err as Error).message}`);
    }
  };

  const reveal = async () => {
    await fileSystemService.showInExplorer(tab.path);
  };

  const copyPath = async () => {
    try {
      if (window.electronAPI?.writeClipboard) {
        await window.electronAPI.writeClipboard(tab.path);
      } else {
        await navigator.clipboard.writeText(tab.path);
      }
      onNotify('Путь скопирован');
    } catch {
      onNotify('Не удалось скопировать путь');
    }
  };

  return (
    <div className="flex h-full flex-col bg-ws-editor text-ws">
      <div className="flex shrink-0 flex-wrap items-center justify-between gap-2 border-b border-ws bg-ws-surface px-3 py-2">
        <div className="flex min-w-0 items-center gap-2">
          <Binary size={16} className="shrink-0 text-amber-500" />
          <div className="min-w-0">
            <div className="truncate text-sm font-semibold">{tab.name}</div>
            <div className="truncate font-mono text-[10px] text-ws-faint" title={tab.path}>
              {tab.path}
            </div>
          </div>
          <span className="ml-2 shrink-0 rounded bg-ws-elevated px-1.5 py-0.5 text-[10px] text-ws-muted">
            {size > 0 ? formatSize(size) : '…'}
          </span>
        </div>
        <div className="flex flex-wrap items-center gap-1.5">
          <Button variant="secondary" onClick={() => void openExternally()} className="gap-1.5">
            <ExternalLink size={12} />
            Открыть в системе
          </Button>
          <Button variant="secondary" onClick={() => void reveal()} className="gap-1.5">
            <FolderOpen size={12} />
            Проводник
          </Button>
          <Button variant="secondary" onClick={() => void copyPath()} className="gap-1.5">
            <Copy size={12} />
            Путь
          </Button>
          {bytes.length < size && (
            <Button
              variant="secondary"
              onClick={() => setLimit((value) => Math.min(MAX_LIMIT, value * 4))}
              className="gap-1.5"
              disabled={loading || limit >= MAX_LIMIT}
            >
              <RefreshCw size={12} className={loading ? 'animate-spin' : ''} />
              Показать больше
            </Button>
          )}
        </div>
      </div>

      <div className="min-h-0 flex-1 overflow-auto p-3">
        {loading && bytes.length === 0 ? (
          <div className="flex items-center gap-2 text-xs text-ws-muted">
            <Loader2 size={14} className="animate-spin" />
            Чтение файла…
          </div>
        ) : error ? (
          <div className="rounded border border-red-500/30 bg-red-500/10 p-3 text-xs text-ws-danger">
            {error}
          </div>
        ) : (
          <div className="rounded border border-ws bg-ws-surface">
            <table className="w-full border-collapse font-mono text-[11px]">
              <thead>
                <tr className="border-b border-ws bg-ws-elevated text-[10px] text-ws-faint">
                  <th className="px-2 py-1 text-right">Offset</th>
                  <th className="px-2 py-1 text-left" colSpan={16}>
                    Hex
                  </th>
                  <th className="px-2 py-1 text-left">ASCII</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row) => (
                  <tr key={row.offset} className="hover:bg-ws-hover/50">
                    <td className="border-r border-ws px-2 py-0.5 text-right text-ws-faint">
                      {row.offset.toString(16).padStart(8, '0').toUpperCase()}
                    </td>
                    {row.hex.map((part, index) => (
                      <td
                        key={index}
                        className="px-1 py-0.5 text-center text-ws-muted"
                      >
                        {part}
                      </td>
                    ))}
                    <td className="border-l border-ws px-2 py-0.5 text-ws">
                      {row.ascii}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {size > bytes.length && (
              <div className="border-t border-ws bg-ws-elevated px-3 py-1.5 text-[10px] text-ws-muted">
                Показаны первые {formatSize(bytes.length)} из {formatSize(size)}.
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};
