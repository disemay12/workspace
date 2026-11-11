/**
 * Источник данных таблицы для редактора.
 *
 * Маленькие файлы открываются как раньше — целиком в памяти рендерера.
 * Большие (см. `state/preferences.ts` → пороги) разбираются в
 * `sheetWorker.ts`: редактор получает строки окнами и не держит книгу целиком.
 */
import { fileSystemService } from '../../services/fileSystemService';
import { readPreferences, shouldUseWindowedSheet } from '../../state/preferences';
import type { CellEdit, SheetMeta, WorkerRequest, WorkerResponse } from './sheetWorker';

export type { CellEdit, SheetMeta } from './sheetWorker';

export type SheetRow = (string | number | boolean | null)[];

export interface SheetSource {
  /** Список листов и их размеры. */
  meta(): Promise<SheetMeta[]>;
  /** Окно строк указанного листа. */
  getWindow(sheet: string, start: number, count: number): Promise<SheetRow[]>;
  /** Правки ячеек: копятся в воркере. */
  setCells(sheet: string, edits: CellEdit[]): Promise<void>;
  /** Сериализация книги в байты (null — файл только для чтения). */
  save(): Promise<ArrayBuffer | null>;
  /** Прогресс чтения файла, проценты 0..100. */
  onProgress(handler: (percent: number) => void): () => void;
  close(): void;
}

/**
 * Решение об оконном режиме для конкретного файла.
 * Для CSV размер файла даёт хорошую оценку объёма разбора.
 */
export async function shouldOpenWindowed(filePath: string, ext: string): Promise<boolean> {
  const preferences = readPreferences();
  if (preferences.largeSheet.mode === 'always') return true;
  if (preferences.largeSheet.mode === 'never') return false;
  try {
    const stat = await fileSystemService.statFile(filePath);
    // Для CSV порог ниже: строки дешевле, но их обычно намного больше.
    const mb = stat.size / (1024 * 1024);
    return shouldUseWindowedSheet(preferences, stat.size) || (ext !== 'csv' && mb > 64);
  } catch {
    return false;
  }
}

/** Источник на основе Web Worker. */
export class WorkerSheetSource implements SheetSource {
  private worker: Worker;
  private sequence = 1;
  private pending = new Map<
    number,
    { resolve: (value: unknown) => void; reject: (error: Error) => void }
  >();
  private progressHandlers = new Set<(percent: number) => void>();
  private closed = false;

  constructor(private readonly filePath: string) {
    this.worker = new Worker(new URL('./sheetWorker.ts', import.meta.url), { type: 'module' });
    this.worker.onmessage = (event: MessageEvent<WorkerResponse | {
      id: number;
      kind: 'read';
      path: string;
      offset: number;
      length: number;
    }>) => {
      void this.handleMessage(event.data);
    };
    this.worker.onerror = (event) => {
      const error = new Error(event.message || 'Ошибка воркера разбора таблиц');
      for (const [, handler] of this.pending) handler.reject(error);
      this.pending.clear();
    };
  }

  private async handleMessage(
    message:
      | WorkerResponse
      | { id: number; kind: 'read'; path: string; offset: number; length: number }
  ): Promise<void> {
    if ('kind' in message && message.kind === 'read') {
      // Воркер просит кусок файла: читает главный процесс.
      try {
        const slice = await fileSystemService.readBinarySlice(
          message.path,
          message.offset,
          message.length
        );
        this.worker.postMessage(
          {
            id: message.id,
            kind: 'readResult',
            bytes: slice.bytes,
            eof: slice.eof,
            total: slice.total,
          },
          [slice.bytes]
        );
      } catch (err) {
        this.worker.postMessage({
          id: message.id,
          kind: 'readResult',
          bytes: new ArrayBuffer(0),
          eof: true,
          total: 0,
        });
        this.failAll(err instanceof Error ? err : new Error(String(err)));
      }
      return;
    }

    if ('kind' in message && message.kind === 'progress') {
      const percent =
        message.total > 0 ? Math.min(99, Math.round((message.done / message.total) * 100)) : 0;
      for (const handler of this.progressHandlers) handler(percent);
      return;
    }

    const response = message as WorkerResponse;
    if (!('id' in response)) return;
    const entry = this.pending.get(response.id);
    if (!entry) return;
    this.pending.delete(response.id);
    if ('ok' in response && response.ok) entry.resolve(response.result);
    else entry.reject(new Error(('error' in response && response.error) || 'Неизвестная ошибка разбора'));
  }

  private failAll(error: Error): void {
    for (const [, handler] of this.pending) handler.reject(error);
    this.pending.clear();
  }

  private send<T>(payload: Omit<WorkerRequest, 'id'>, transfer?: Transferable[]): Promise<T> {
    if (this.closed) return Promise.reject(new Error('Источник таблицы закрыт'));
    const id = this.sequence++;
    return new Promise<T>((resolve, reject) => {
      this.pending.set(id, { resolve: resolve as (value: unknown) => void, reject });
      this.worker.postMessage({ ...payload, id }, transfer ?? []);
    });
  }

  onProgress(handler: (percent: number) => void): () => void {
    this.progressHandlers.add(handler);
    return () => this.progressHandlers.delete(handler);
  }

  async openCsv(name: string, encoding: string): Promise<SheetMeta[]> {
    return this.send<SheetMeta[]>({
      kind: 'openCsv',
      path: this.filePath,
      name,
      encoding,
    } as Omit<WorkerRequest, 'id'>);
  }

  async openXlsx(bytes: ArrayBuffer, name: string, ext: string): Promise<SheetMeta[]> {
    return this.send<SheetMeta[]>(
      { kind: 'openXlsx', bytes, name, ext } as Omit<WorkerRequest, 'id'>,
      [bytes]
    );
  }

  async meta(): Promise<SheetMeta[]> {
    return this.send<SheetMeta[]>({ kind: 'meta' } as Omit<WorkerRequest, 'id'>);
  }

  async getWindow(sheet: string, start: number, count: number): Promise<SheetRow[]> {
    return this.send<SheetRow[]>({ kind: 'getWindow', sheet, start, count } as Omit<
      WorkerRequest,
      'id'
    >);
  }

  async setCells(sheet: string, edits: CellEdit[]): Promise<void> {
    await this.send<void>({ kind: 'setCells', sheet, edits } as Omit<WorkerRequest, 'id'>);
  }

  async save(): Promise<ArrayBuffer | null> {
    return this.send<ArrayBuffer>({ kind: 'save' } as Omit<WorkerRequest, 'id'>);
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    try {
      this.worker.postMessage({ id: this.sequence++, kind: 'close' });
    } catch {
      // воркер уже мог быть завершён
    }
    this.worker.terminate();
    this.failAll(new Error('Источник таблицы закрыт'));
  }
}

/**
 * Открывает большой файл в воркере.
 * CSV читается потоково, XLSX — одним блоком (разбор всё равно в воркере).
 */
export async function openWindowedSheet(
  filePath: string,
  options: { name: string; ext: string; encoding?: string }
): Promise<{ source: WorkerSheetSource; meta: SheetMeta[] }> {
  const source = new WorkerSheetSource(filePath);
  try {
    const ext = options.ext.toLowerCase();
    if (ext === 'csv' || ext === 'tsv') {
      const encoding =
        options.encoding && options.encoding !== 'utf-8-bom' ? options.encoding : 'utf-8';
      const meta = await source.openCsv(options.name, encoding);
      return { source, meta };
    }
    const bytes = await fileSystemService.readBinary(filePath);
    const meta = await source.openXlsx(bytes, options.name, ext);
    return { source, meta };
  } catch (err) {
    source.close();
    throw err;
  }
}
