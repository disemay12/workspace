import { ipcMain } from 'electron';
import { execFileSync } from 'child_process';
import * as path from 'path';
import * as fs from 'fs';

/**
 * Определяет оптимальный путь к исполняемому файлу WorkSpace.
 * Если существует внешний оконный загрузчик WorkSpace.exe (в корне портативной папки),
 * используется именно он, чтобы запускаться с иконкой и нужными параметрами.
 */
export function getAppExecutablePath(): string {
  try {
    const parentDir = path.dirname(process.execPath);
    const candidateLauncher = path.resolve(parentDir, '..', 'WorkSpace.exe');
    if (fs.existsSync(candidateLauncher)) {
      return candidateLauncher;
    }

    const localLauncher = path.resolve(parentDir, 'WorkSpace.exe');
    if (fs.existsSync(localLauncher)) {
      return localLauncher;
    }
  } catch {
    // fallback
  }

  return process.execPath;
}

/**
 * Проверяет, зарегистрирован ли пункт WorkSpace в контекстном меню Windows Explorer.
 */
export function isContextMenuInstalled(): boolean {
  if (process.platform !== 'win32') return false;
  try {
    const out = execFileSync('reg', ['query', 'HKCU\\Software\\Classes\\*\\shell\\WorkSpace'], {
      encoding: 'utf8',
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'ignore'],
    });
    return Boolean(out && out.includes('WorkSpace'));
  } catch {
    return false;
  }
}

/**
 * Регистрирует пункт «Редактировать в WorkSpace» для файлов и папок в HKCU.
 * Не требует прав администратора (UAC) и действует для текущего пользователя Windows.
 */
export function installContextMenu(): { success: boolean; error?: string } {
  if (process.platform !== 'win32') {
    return { success: false, error: 'Контекстное меню доступно только в ОС Windows' };
  }

  try {
    const exePath = getAppExecutablePath();

    // 1. Для всех файлов (*)
    // Пункт: "Редактировать в WorkSpace"
    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\*\\shell\\WorkSpace',
      '/ve',
      '/d',
      'Редактировать в WorkSpace',
      '/f',
    ], { windowsHide: true });

    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\*\\shell\\WorkSpace',
      '/v',
      'Icon',
      '/d',
      `"${exePath}"`,
      '/f',
    ], { windowsHide: true });

    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\*\\shell\\WorkSpace\\command',
      '/ve',
      '/d',
      `"${exePath}" "%1"`,
      '/f',
    ], { windowsHide: true });

    // 2. Для папок (Directory)
    // Пункт: "Открыть в WorkSpace"
    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\Directory\\shell\\WorkSpace',
      '/ve',
      '/d',
      'Открыть в WorkSpace',
      '/f',
    ], { windowsHide: true });

    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\Directory\\shell\\WorkSpace',
      '/v',
      'Icon',
      '/d',
      `"${exePath}"`,
      '/f',
    ], { windowsHide: true });

    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\Directory\\shell\\WorkSpace\\command',
      '/ve',
      '/d',
      `"${exePath}" "%1"`,
      '/f',
    ], { windowsHide: true });

    // 3. Для фона открытой папки (Directory\Background)
    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\Directory\\Background\\shell\\WorkSpace',
      '/ve',
      '/d',
      'Открыть в WorkSpace',
      '/f',
    ], { windowsHide: true });

    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\Directory\\Background\\shell\\WorkSpace',
      '/v',
      'Icon',
      '/d',
      `"${exePath}"`,
      '/f',
    ], { windowsHide: true });

    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\Directory\\Background\\shell\\WorkSpace\\command',
      '/ve',
      '/d',
      `"${exePath}" "%V"`,
      '/f',
    ], { windowsHide: true });

    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
}

/**
 * Удаляет записи WorkSpace из контекстного меню Windows Explorer.
 */
export function removeContextMenu(): { success: boolean; error?: string } {
  if (process.platform !== 'win32') {
    return { success: false, error: 'Контекстное меню доступно только в ОС Windows' };
  }

  try {
    // Удаляем из файлов
    try {
      execFileSync('reg', ['delete', 'HKCU\\Software\\Classes\\*\\shell\\WorkSpace', '/f'], {
        windowsHide: true,
        stdio: ['ignore', 'ignore', 'ignore'],
      });
    } catch {}

    // Удаляем из папок
    try {
      execFileSync('reg', ['delete', 'HKCU\\Software\\Classes\\Directory\\shell\\WorkSpace', '/f'], {
        windowsHide: true,
        stdio: ['ignore', 'ignore', 'ignore'],
      });
    } catch {}

    // Удаляем из фона папок
    try {
      execFileSync('reg', ['delete', 'HKCU\\Software\\Classes\\Directory\\Background\\shell\\WorkSpace', '/f'], {
        windowsHide: true,
        stdio: ['ignore', 'ignore', 'ignore'],
      });
    } catch {}

    return { success: true };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
}

/**
 * Регистрирует IPC-обработчики контекстного меню для взаимодействия с рендерером.
 */
export function registerContextMenuIpc(): void {
  ipcMain.handle('system:getContextMenuStatus', async () => {
    return {
      installed: isContextMenuInstalled(),
      exePath: getAppExecutablePath(),
    };
  });

  ipcMain.handle('system:installContextMenu', async () => {
    return installContextMenu();
  });

  ipcMain.handle('system:removeContextMenu', async () => {
    return removeContextMenu();
  });
}
