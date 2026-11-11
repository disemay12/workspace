import { app, ipcMain, shell } from 'electron';
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

/** Значение параметра ветки реестра (пустая строка, если его нет). */
function regValue(key: string, name: string): string {
  try {
    const out = execFileSync('reg', ['query', key, '/v', name], {
      encoding: 'utf8',
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'ignore'],
    });
    const match = new RegExp(`${name}\\s+REG_\\w+\\s+(.+)`).exec(out);
    return match?.[1]?.trim().replace(/^"|"$/g, '') ?? '';
  } catch {
    return '';
  }
}

/** Подпись пункта контекстного меню ((default) ветки). */
function regDefault(key: string): string {
  try {
    const out = execFileSync('reg', ['query', key, '/ve'], {
      encoding: 'utf8',
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'ignore'],
    });
    const match = /REG_SZ\s+(.+)/.exec(out);
    return match?.[1]?.trim() ?? '';
  } catch {
    return '';
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
    if (path.basename(exePath).toLowerCase() === 'electron.exe') {
      return {
        success: false,
        error: 'Контекстное меню настраивается в собранной версии WorkSpace, а не в режиме разработки',
      };
    }

    // Если путь уже актуален, не перезаписываем реестр: это важно в средах,
    // где процессу запрещена запись в HKCU, хотя меню уже зарегистрировано.
    try {
      const out = execFileSync(
        'reg',
        ['query', 'HKCU\\Software\\Classes\\*\\shell\\WorkSpace', '/v', 'Icon'],
        { encoding: 'utf8', windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'] }
      );
      const match = /Icon\s+REG_SZ\s+(.+)/.exec(out);
      const registered = match?.[1]?.trim().replace(/^"|"$/g, '');
      if (
        registered &&
        path.resolve(registered).toLowerCase() === path.resolve(exePath).toLowerCase() &&
        isContextMenuInstalled()
      ) {
        return { success: true };
      }
    } catch {
      // значение не прочитано — выполняем полную регистрацию
    }

    // 1. Для всех файлов (*)
    // Подпись одна для всех типов: «Открыть в WorkSpace» — пункт открывает
    // файл в приложении, а не «правит» его.
    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\*\\shell\\WorkSpace',
      '/ve',
      '/d',
      'Открыть в WorkSpace',
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

    // 2. Для исполняемых файлов (exefile) — пункт должен быть виден и для .exe
    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\exefile\\shell\\WorkSpace',
      '/ve',
      '/d',
      'Открыть в WorkSpace',
      '/f',
    ], { windowsHide: true });

    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\exefile\\shell\\WorkSpace',
      '/v',
      'Icon',
      '/d',
      `"${exePath}"`,
      '/f',
    ], { windowsHide: true });

    execFileSync('reg', [
      'add',
      'HKCU\\Software\\Classes\\exefile\\shell\\WorkSpace\\command',
      '/ve',
      '/d',
      `"${exePath}" "%1"`,
      '/f',
    ], { windowsHide: true });

    // 3. Для папок (Directory)
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

    // 4. Для фона открытой папки (Directory\Background)
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

    // reg.exe на части систем срезает кавычки в значении команды.
    // Перезаписываем через PowerShell, чтобы пути с пробелами работали.
    const quotedExe = `"${exePath}"`;
    const quoteScript = [
      "$ErrorActionPreference = 'Stop'",
      `Set-ItemProperty -LiteralPath 'HKCU:\\Software\\Classes\\*\\shell\\WorkSpace\\command' -Name '(default)' -Value '${quotedExe.replace(/'/g, "''")} "%1"'`,
      `Set-ItemProperty -LiteralPath 'HKCU:\\Software\\Classes\\exefile\\shell\\WorkSpace\\command' -Name '(default)' -Value '${quotedExe.replace(/'/g, "''")} "%1"'`,
      `Set-ItemProperty -LiteralPath 'HKCU:\\Software\\Classes\\Directory\\shell\\WorkSpace\\command' -Name '(default)' -Value '${quotedExe.replace(/'/g, "''")} "%1"'`,
      `Set-ItemProperty -LiteralPath 'HKCU:\\Software\\Classes\\Directory\\Background\\shell\\WorkSpace\\command' -Name '(default)' -Value '${quotedExe.replace(/'/g, "''")} "%V"'`,
    ].join('; ');
    try {
      execFileSync('powershell.exe', [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-Command',
        quoteScript,
      ], { windowsHide: true });
    } catch {
      // Основные значения уже записаны через reg.exe
    }

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

    // Удаляем из exefile
    try {
      execFileSync('reg', ['delete', 'HKCU\\Software\\Classes\\exefile\\shell\\WorkSpace', '/f'], {
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
 * Создаёт или обновляет ярлык WorkSpace на рабочем столе текущего пользователя.
 * Иконка берётся из исполняемого файла (встроена в launcher при сборке).
 */
export function createDesktopShortcut(): {
  success: boolean;
  shortcutPath?: string;
  error?: string;
} {
  if (process.platform !== 'win32') {
    return { success: false, error: 'Ярлыки доступны только в ОС Windows' };
  }
  try {
    const exePath = getAppExecutablePath();
    const shortcutPath = path.join(app.getPath('desktop'), 'WorkSpace.lnk');
    let created = shell.writeShortcutLink(shortcutPath, 'replace', {
      target: exePath,
      cwd: path.dirname(exePath),
      icon: exePath,
      iconIndex: 0,
      description: 'WorkSpace — документы, офис и AI',
    });

    // shell.writeShortcutLink может вернуть false в среде с ограниченным
    // токеном — повторяем через WScript.Shell, который умеет то же самое.
    if (!created) {
      const quote = (value: string) => value.replace(/'/g, "''");
      const script = [
        '$ErrorActionPreference = "Stop"',
        '$ws = New-Object -ComObject WScript.Shell',
        `$s = $ws.CreateShortcut('${quote(shortcutPath)}')`,
        `$s.TargetPath = '${quote(exePath)}'`,
        `$s.WorkingDirectory = '${quote(path.dirname(exePath))}'`,
        `$s.IconLocation = '${quote(exePath)},0'`,
        '$s.Description = "WorkSpace - documents, office and AI"',
        '$s.Save()',
      ].join('; ');
      execFileSync('powershell.exe', [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-Command',
        script,
      ], { windowsHide: true });
      created = fs.existsSync(shortcutPath);
    }

    return created
      ? { success: true, shortcutPath }
      : { success: false, error: 'Windows не создал ярлык' };
  } catch (err) {
    return { success: false, error: (err as Error).message };
  }
}

export function desktopShortcutPath(): string {
  return path.join(app.getPath('desktop'), 'WorkSpace.lnk');
}

export function isDesktopShortcutPresent(): boolean {
  try {
    return fs.existsSync(desktopShortcutPath());
  } catch {
    return false;
  }
}

/**
 * Регистрирует IPC-обработчики контекстного меню для взаимодействия с рендерером.
 */
export function registerContextMenuIpc(): void {
  ipcMain.handle('system:getContextMenuStatus', async () => {
    const exePath = getAppExecutablePath();
    const keys = {
      files: 'HKCU\\Software\\Classes\\*\\shell\\WorkSpace',
      exefile: 'HKCU\\Software\\Classes\\exefile\\shell\\WorkSpace',
      directory: 'HKCU\\Software\\Classes\\Directory\\shell\\WorkSpace',
      background: 'HKCU\\Software\\Classes\\Directory\\Background\\shell\\WorkSpace',
    };
    const labels: Record<string, string> = {};
    for (const [id, key] of Object.entries(keys)) labels[id] = regDefault(key);
    const icon = regValue(keys.files, 'Icon');
    const command = regDefault(keys.files + '\\command');
    return {
      installed: isContextMenuInstalled(),
      exePath,
      labels,
      command,
      icon,
      /** Ярлык указывает на текущий исполняемый файл. */
      iconMatches: Boolean(
        icon && path.resolve(icon).toLowerCase() === path.resolve(exePath).toLowerCase()
      ),
    };
  });

  ipcMain.handle('system:installContextMenu', async () => {
    return installContextMenu();
  });

  ipcMain.handle('system:removeContextMenu', async () => {
    return removeContextMenu();
  });

  ipcMain.handle('system:createDesktopShortcut', async () => createDesktopShortcut());

  ipcMain.handle('system:getDesktopShortcutStatus', async () => ({
    installed: isDesktopShortcutPresent(),
    shortcutPath: desktopShortcutPath(),
  }));
}
