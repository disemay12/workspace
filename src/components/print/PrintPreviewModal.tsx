/**
 * Предпросмотр печати: страницы показываются изображениями из того же PDF,
 * который уйдёт на принтер.
 *
 * `<iframe src="data:application/pdf">` не подходит — CSP приложения запрещает
 * фреймы и объекты (`frame-src 'none'`, `object-src 'none'`), поэтому страницы
 * рендерит pdf.js в картинки.
 */
import React, { useEffect, useRef, useState } from 'react';
import { Printer, Save, X } from 'lucide-react';
import { Spinner } from '../ui';
import { loadPdfDocument } from '../../lib/pdfRender';
import type { PrintDocument } from '../../services/printBridge';
import { fileSystemService } from '../../services/fileSystemService';

interface Props {
  open: boolean;
  document: PrintDocument | null;
  pdfBase64: string | null;
  onClose: () => void;
  onPrint: () => Promise<void> | void;
  onError: (message: string) => void;
}

/** Ширина миниатюры: печатная страница целиком, читаемая на экране. */
const THUMB_WIDTH = 560;

export const PrintPreviewModal: React.FC<Props> = ({
  open,
  document,
  pdfBase64,
  onClose,
  onPrint,
  onError,
}) => {
  const [pages, setPages] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const runId = useRef(0);

  useEffect(() => {
    if (!open || !pdfBase64) {
      setPages([]);
      return;
    }
    const current = ++runId.current;
    setLoading(true);
    void (async () => {
      try {
        const pdf = await loadPdfDocument(pdfBase64);
        const result: string[] = [];
        for (let index = 0; index < pdf.numPages; index += 1) {
          const page = await pdf.getPage(index + 1);
          const base = page.getViewport({ scale: 1 });
          const scale = Math.min(2, THUMB_WIDTH / base.width);
          const viewport = page.getViewport({ scale });
          const canvas = window.document.createElement('canvas');
          canvas.width = Math.ceil(viewport.width);
          canvas.height = Math.ceil(viewport.height);
          const context = canvas.getContext('2d');
          if (!context) break;
          await page.render({ canvasContext: context, viewport, canvas }).promise;
          result.push(canvas.toDataURL('image/png'));
          if (runId.current !== current) return;
        }
        setPages(result);
      } catch (err) {
        onError(`Не удалось показать предпросмотр: ${(err as Error).message}`);
      } finally {
        if (runId.current === current) setLoading(false);
      }
    })();
  }, [open, pdfBase64, onError]);

  if (!open) return null;

  const savePdf = async () => {
    if (!pdfBase64 || !document) return;
    setSaving(true);
    try {
      const target = await fileSystemService.saveFileDialog(`${document.title || 'document'}.pdf`, [
        { name: 'PDF', extensions: ['pdf'] },
      ]);
      if (!target) return;
      await fileSystemService.writeFile(target, pdfBase64, 'base64');
    } catch (err) {
      onError(`Не удалось сохранить PDF: ${(err as Error).message}`);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-6"
      onClick={onClose}
      role="presentation"
    >
      <div
        className="flex h-full max-h-[92vh] w-full max-w-5xl flex-col overflow-hidden rounded-lg border border-ws bg-ws-surface shadow-ws"
        onClick={(event) => event.stopPropagation()}
        role="dialog"
        aria-label="Предпросмотр печати"
      >
        <div className="flex shrink-0 items-center gap-2 border-b border-ws px-4 py-2 text-xs">
          <Printer size={14} className="text-blue-500" />
          <span className="font-medium text-ws">Предпросмотр печати</span>
          <span className="truncate text-ws-muted">{document?.title}</span>
          <span className="ml-2 text-ws-faint" data-print-pages="1">
            {loading ? 'Готовим страницы…' : `Страниц: ${pages.length}`}
          </span>
          <div className="ml-auto flex items-center gap-2">
            <button
              type="button"
              onClick={() => void onPrint()}
              className="flex items-center gap-1 rounded bg-ws-accent px-2.5 py-1 text-[11px] font-medium text-white hover:opacity-90"
            >
              <Printer size={12} />
              Печать
            </button>
            <button
              type="button"
              onClick={() => void savePdf()}
              disabled={saving || !pdfBase64}
              className="flex items-center gap-1 rounded border border-ws px-2.5 py-1 text-[11px] text-ws hover:bg-ws-hover disabled:opacity-40"
            >
              <Save size={12} />
              {saving ? 'Сохраняем…' : 'Сохранить PDF'}
            </button>
            <button
              type="button"
              onClick={onClose}
              className="rounded p-1 text-ws-muted hover:bg-ws-hover hover:text-ws"
              title="Закрыть"
            >
              <X size={14} />
            </button>
          </div>
        </div>
        <div className="min-h-0 flex-1 overflow-auto bg-ws-editor-alt p-4">
          {loading && pages.length === 0 ? (
            <div className="flex h-full items-center justify-center gap-2 text-xs text-ws-muted">
              <Spinner size={14} />
              Рендерим страницы…
            </div>
          ) : (
            <div className="mx-auto flex w-fit flex-col items-center gap-4">
              {pages.map((src, index) => (
                <div key={index} className="flex flex-col items-center gap-1">
                  <img
                    src={src}
                    alt={`Страница ${index + 1}`}
                    className="max-w-full rounded-sm bg-white shadow-ws"
                    style={{ width: THUMB_WIDTH }}
                  />
                  <span className="text-[10px] text-ws-faint">Страница {index + 1}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
