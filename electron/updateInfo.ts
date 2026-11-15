/**
 * Контракт ответа `update:check` для неподдерживаемой сборки.
 *
 * Раньше эту ветку собирали трижды: в updater.ts, в заглушке проверочного
 * стенда (common-ipc.ts) и в ожиданиях integration-check.ts. Проверка сверяла
 * заглушку с собственной копией литералов, поэтому расхождение с боевым
 * ответом проходило молча. Теперь определение одно, и его используют все.
 */
export const UNSUPPORTED_UPDATE_ERROR =
  'Проверка обновлений доступна только в собранной portable-версии WorkSpace';

export interface UnsupportedUpdateInfo {
  success: true;
  supported: false;
  currentVersion: string;
  error: string;
}

/** Ответ `update:check`, когда приложение запущено не из portable-сборки. */
export function unsupportedUpdateInfo(currentVersion: string): UnsupportedUpdateInfo {
  return {
    success: true,
    supported: false,
    currentVersion,
    error: UNSUPPORTED_UPDATE_ERROR,
  };
}
