import React, { useEffect, useMemo, useState } from 'react';
import { AlertTriangle, Download, FolderOpen } from 'lucide-react';
import type { OpenTab } from '../../types';
import { baseName, dirName, joinPath } from '../../lib/utils';
import { fileSystemService } from '../../services/fileSystemService';
import {
  type ConversionFormat,
  defaultTargetName,
  formatLabel,
  formatsFor,
} from '../../services/documentConversion';
import { Button, Input, Modal, Select } from '../ui';

interface Props {
  open: boolean;
  tab: OpenTab | null;
  busy?: boolean;
  onClose: () => void;
  onSave: (targetPath: string, format: ConversionFormat) => Promise<boolean>;
  onNotify?: (message: string) => void;
}

export const SaveAsDialog: React.FC<Props> = ({
  open,
  tab,
  busy = false,
  onClose,
  onSave,
  onNotify,
}) => {
  const formats = useMemo(
    () => (tab ? formatsFor({ extension: tab.extension, type: tab.type }) : []),
    [tab]
  );
  const [format, setFormat] = useState<ConversionFormat>('docx');
  const [name, setName] = useState('');
  const [dir, setDir] = useState('');
  const [exists, setExists] = useState(false);
  const [confirmOverwrite, setConfirmOverwrite] = useState(false);
  const [error, setError] = useState('');
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (!open || !tab || formats.length === 0) return;
    const current = (formats.find((item) => item === tab.extension) ?? formats[0]) as ConversionFormat;
    setFormat(current);
    setName(defaultTargetName(tab.name, current));
    setDir(dirName(tab.path));
    setExists(false);
    setConfirmOverwrite(false);
    setError('');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, tab?.id]);

  const targetPath = useMemo(
    () => (dir && name.trim() ? joinPath(dir, name.trim()) : ''),
    [dir, name]
  );

  const changeFormat = (next: ConversionFormat) => {
    setFormat(next);
    setConfirmOverwrite(false);
    setError('');
    setName((value) => {
      const base = value.replace(/\.[^./\\]+$/, '') || (tab ? tab.name.replace(/\.[^.]+$/, '') : 'Документ');
      return `${base}.${next}`;
    });
  };

  const chooseFolder = async () => {
    const picked = await fileSystemService.pickDirectory(dir || undefined);
    if (picked) {
      setDir(picked);
      setConfirmOverwrite(false);
    }
  };

  useEffect(() => {
    if (!open || !targetPath) {
      setExists(false);
      return;
    }
    let cancelled = false;
    const timer = window.setTimeout(async () => {
      const value = await window.electronAPI?.pathExists?.(targetPath);
      if (!cancelled) {
        setExists(Boolean(value));
        if (!value) setConfirmOverwrite(false);
      }
    }, 250);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [open, targetPath]);

  const handleSave = async () => {
    if (!targetPath || formats.length === 0) return;
    setError('');
    if (exists && !confirmOverwrite) {
      setConfirmOverwrite(true);
      setError('Файл с таким именем уже существует. Нажмите «Перезаписать», чтобы заменить его.');
      return;
    }
    setSaving(true);
    try {
      const ok = await onSave(targetPath, format);
      if (ok) {
        onNotify?.(`Сохранено: ${baseName(targetPath)}`);
        onClose();
      }
    } finally {
      setSaving(false);
    }
  };

  const disabled = busy || saving || !name.trim() || !dir || formats.length === 0;

  return (
    <Modal
      open={open}
      onClose={onClose}
      title="Сохранить как"
      icon={<Download size={15} className="text-ws-accent" />}
      width="max-w-lg"
      footer={
        <>
          <Button variant="ghost" onClick={onClose} disabled={saving}>
            Отмена
          </Button>
          <Button
            variant="primary"
            onClick={() => void handleSave()}
            disabled={disabled}
          >
            {saving ? 'Сохранение…' : confirmOverwrite ? 'Перезаписать' : 'Сохранить'}
          </Button>
        </>
      }
    >
      {!tab || formats.length === 0 ? (
        <div className="flex items-start gap-2 rounded border border-amber-500/30 bg-amber-500/10 p-3 text-[11px] text-ws-warning">
          <AlertTriangle size={14} className="mt-0.5 shrink-0" />
          Для этого типа файла конвертация не поддерживается. Откройте документ, таблицу,
          презентацию, Markdown или текст.
        </div>
      ) : (
        <div className="space-y-3">
          <div className="grid grid-cols-[92px_1fr] items-center gap-2">
            <span className="text-[11px] text-ws-muted">Формат</span>
            <Select
              value={format}
              onChange={(event) => changeFormat(event.target.value as ConversionFormat)}
            >
              {formats.map((item) => (
                <option key={item} value={item}>
                  {formatLabel(item)}
                </option>
              ))}
            </Select>

            <span className="text-[11px] text-ws-muted">Имя файла</span>
            <Input
              autoFocus
              value={name}
              onChange={(event) => {
                setName(event.target.value);
                setConfirmOverwrite(false);
                setError('');
              }}
              onKeyDown={(event) => {
                if (event.key === 'Enter') void handleSave();
              }}
              placeholder="Документ.docx"
            />

            <span className="text-[11px] text-ws-muted">Папка</span>
            <div className="flex min-w-0 items-center gap-2">
              <span className="min-w-0 flex-1 truncate rounded border border-ws bg-ws-elevated px-2 py-1 font-mono text-[10px] text-ws-muted" title={dir}>
                {dir || 'Не выбрана'}
              </span>
              <Button variant="secondary" onClick={() => void chooseFolder()} className="shrink-0 gap-1">
                <FolderOpen size={12} />
                Выбрать
              </Button>
            </div>
          </div>

          <p className="text-[10px] text-ws-faint">
            {format === 'pdf'
              ? 'PDF создаётся из текущего вида документа и открывается системным просмотрщиком.'
              : format === 'xlsx'
                ? 'Excel сохраняется с автоподбором ширины колонок, чтобы текст не обрезался.'
                : format === 'csv' || format === 'tsv'
                  ? 'CSV/TSV сохраняется в UTF-8 с BOM — русский Excel откроет его без «кракозябр».'
                  : 'Исходный файл не изменяется: создаётся копия в выбранном формате.'}
          </p>

          {exists && !error && (
            <p className="text-[10px] text-ws-warning">Файл уже существует и будет заменён.</p>
          )}
          {error && (
            <p className="rounded border border-amber-500/30 bg-amber-500/10 px-2 py-1.5 text-[10px] text-ws-warning">
              {error}
            </p>
          )}
        </div>
      )}
    </Modal>
  );
};
