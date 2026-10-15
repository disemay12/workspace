/**
 * Встроенный терминал WorkSpace.
 *
 * Управляет сессиями интерактивного терминала (PowerShell / cmd на Windows, bash на Unix).
 * Общается с рендерером через IPC: передача ввода пользователя и стриминг вывода в xterm.js.
 */
import { ipcMain, BrowserWindow } from 'electron';
import { spawn, ChildProcess } from 'child_process';
import * as os from 'os';

interface TerminalSession {
  id: string;
  process: ChildProcess;
}

const sessions = new Map<string, TerminalSession>();

export function registerTerminal(windowGetter: () => BrowserWindow | null): void {
  ipcMain.handle('terminal:create', (_event, req?: { id?: string; cwd?: string; shell?: string }) => {
    const id = req?.id || 'term_' + Math.random().toString(36).substring(2, 9);
    const cwd = req?.cwd && os.platform() === 'win32' ? req.cwd : process.cwd();
    const isWin = os.platform() === 'win32';
    const shellExe = req?.shell || (isWin ? 'powershell.exe' : 'bash');

    // Если сессия с таким id уже существует, завершаем её
    const existing = sessions.get(id);
    if (existing) {
      try {
        existing.process.kill();
      } catch {
        // ignore
      }
      sessions.delete(id);
    }

    try {
      const args = isWin ? ['-NoLogo'] : [];
      const child = spawn(shellExe, args, {
        cwd,
        env: {
          ...process.env,
          TERM: 'xterm-256color',
          COLORTERM: 'truecolor',
        },
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true,
      });

      child.stdout?.on('data', (chunk: Buffer) => {
        const win = windowGetter();
        if (win && !win.isDestroyed()) {
          win.webContents.send('terminal:data', { id, data: chunk.toString('utf-8') });
        }
      });

      child.stderr?.on('data', (chunk: Buffer) => {
        const win = windowGetter();
        if (win && !win.isDestroyed()) {
          win.webContents.send('terminal:data', { id, data: chunk.toString('utf-8') });
        }
      });

      child.on('close', (code: number | null) => {
        sessions.delete(id);
        const win = windowGetter();
        if (win && !win.isDestroyed()) {
          win.webContents.send('terminal:exit', { id, code: code ?? 0 });
        }
      });

      child.on('error', (err: Error) => {
        const win = windowGetter();
        if (win && !win.isDestroyed()) {
          win.webContents.send('terminal:data', {
            id,
            data: `\r\nОшибка запуска терминала: ${err.message}\r\n`,
          });
        }
      });

      sessions.set(id, { id, process: child });
      return { success: true, id };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  ipcMain.handle('terminal:input', (_event, req: { id: string; data: string }) => {
    const session = sessions.get(req.id);
    if (!session || !session.process.stdin) {
      return { success: false, error: 'Сессия терминала не найдена' };
    }
    try {
      session.process.stdin.write(req.data);
      return { success: true };
    } catch (err) {
      return { success: false, error: (err as Error).message };
    }
  });

  ipcMain.handle('terminal:kill', (_event, id: string) => {
    const session = sessions.get(id);
    if (session) {
      try {
        session.process.kill();
      } catch {
        // ignore
      }
      sessions.delete(id);
    }
    return { success: true };
  });
}
