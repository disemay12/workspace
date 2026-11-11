/**
 * Диагностика на стороне интерфейса: перехват ошибок и сборка отчёта.
 *
 * Отчёт всегда проходит через `redact`: ключи, токены, почта и домашние пути
 * не должны попадать в буфер обмена или в переписку с поддержкой.
 */
import { redact } from '../lib/redact';

/** Подписка на ошибки окна: они уходят в кольцевой буфер главного процесса. */
export function installDiagnosticsHooks(): () => void {
  const onError = (event: ErrorEvent) => {
    const where = event.filename ? ` @ ${event.filename}:${event.lineno}` : '';
    void Promise.resolve(window.electronAPI?.diagnosticsLog?.('renderer', `${event.message}${where}`)).catch(
      () => undefined
    );
  };
  const onRejection = (event: PromiseRejectionEvent) => {
    const reason = event.reason as { message?: string } | string | undefined;
    const message = typeof reason === 'string' ? reason : (reason?.message ?? 'unhandledrejection');
    void Promise.resolve(window.electronAPI?.diagnosticsLog?.('unhandledrejection', message)).catch(
      () => undefined
    );
  };
  window.addEventListener('error', onError);
  window.addEventListener('unhandledrejection', onRejection);
  return () => {
    window.removeEventListener('error', onError);
    window.removeEventListener('unhandledrejection', onRejection);
  };
}

/** Текст отчёта для буфера обмена: JSON с отредактированными путями. */
export async function buildDiagnosticsReport(): Promise<string> {
  const api = window.electronAPI;
  const result = await api?.diagnosticsReport?.();
  if (!result?.success || !result.report) {
    return `WorkSpace: диагностика недоступна (${result?.error ?? 'нет IPC'})`;
  }
  const report = result.report as {
    system?: { paths?: { home?: string; appData?: string } };
  };
  const raw = JSON.stringify(result.report, null, 2);
  return redact(raw, {
    homeDir: report.system?.paths?.home ?? '',
    profile: report.system?.paths?.appData ?? '',
  });
}

/** Копирует отчёт в буфер обмена и возвращает его размер. */
export async function copyDiagnosticsReport(): Promise<number> {
  const text = await buildDiagnosticsReport();
  await navigator.clipboard.writeText(text);
  return text.length;
}
