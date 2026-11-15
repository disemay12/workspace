/**
 * Встроенный терминал WorkSpace.
 *
 * Управляет сессиями интерактивного терминала (PowerShell / cmd на Windows, bash на Unix).
 * Общается с рендерером через IPC: передача ввода пользователя и стриминг вывода в xterm.js.
 */
import { ipcMain, BrowserWindow, dialog } from 'electron';
import { spawn, ChildProcess } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { getWorkspaceRoot, isInsidePath } from './fsGuard';

interface TerminalSession {
  id: string;
  process: ChildProcess;
}

const sessions = new Map<string, TerminalSession>();

/**
 * Разрешённые оболочки. Рендерер — граница доверия (в документах и ответах
 * модели встречается чужой текст), поэтому произвольный путь к exe здесь
 * запускать нельзя: это готовый запуск любого файла с полным окружением.
 * Ищутся они по абсолютным путям: `spawn('powershell.exe')` разрешает имя по
 * порядку поиска Windows, где каталог приложения идёт раньше System32.
 */
const ALLOWED_SHELLS = new Set([
  'powershell.exe',
  'pwsh.exe',
  'cmd.exe',
  'bash',
  'sh',
  'zsh',
  'wsl.exe',
]);

/** Первое существующее имя из кандидатов: без неявного поиска по PATH. */
function firstExisting(candidates: string[]): string | null {
  for (const candidate of candidates) {
    try {
      if (candidate && fs.existsSync(candidate)) return candidate;
    } catch {
      // Проверка доступа: кандидат не подходит
    }
  }
  return null;
}

function resolveShell(requested?: string): string | null {
  const isWin = os.platform() === 'win32';
  const name = path.basename(String(requested || (isWin ? 'powershell.exe' : 'bash'))).toLowerCase();
  if (!ALLOWED_SHELLS.has(name)) return null;
  if (!isWin) return firstExisting(['/bin/bash', '/bin/sh', '/bin/zsh', '/usr/bin/' + name]);
  const system32 = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32');
  const table: Record<string, string[]> = {
    'powershell.exe': [path.join(system32, 'WindowsPowerShell', 'v1.0', 'powershell.exe')],
    'pwsh.exe': [
      path.join(process.env.ProgramFiles || 'C:\\Program Files', 'PowerShell', '7', 'pwsh.exe'),
      path.join(process.env.ProgramFiles || 'C:\\Program Files', 'PowerShell', '7-preview', 'pwsh.exe'),
    ],
    'cmd.exe': [process.env.ComSpec || path.join(system32, 'cmd.exe')],
    'wsl.exe': [path.join(system32, 'wsl.exe')],
    bash: [
      path.join(process.env.ProgramFiles || 'C:\\Program Files', 'Git', 'bin', 'bash.exe'),
      path.join(system32, 'bash.exe'),
    ],
    sh: [path.join(process.env.ProgramFiles || 'C:\\Program Files', 'Git', 'bin', 'sh.exe')],
    zsh: [path.join(process.env.ProgramFiles || 'C:\\Program Files', 'Git', 'bin', 'sh.exe')],
  };
  return firstExisting(table[name] ?? []);
}

/**
 * Терминал даёт рендереру полноценную оболочку: при XSS это исполнение любых
 * команд без диалога. Поэтому ответ спрашивается на каждое окно и живёт
 * ограниченное время: разрешение, выданное при открытии терминала, не должно
 * доставаться инъекции в документ через полчаса без нового вопроса.
 */
const TERMINAL_GRANT_TTL_MS = 15 * 60 * 1000;
let terminalGrant: { senderId: number; allowed: boolean; expiresAt: number } | null = null;

async function confirmTerminal(
  windowGetter: () => BrowserWindow | null,
  senderId: number
): Promise<boolean> {
  const now = Date.now();
  if (terminalGrant && terminalGrant.senderId === senderId && terminalGrant.expiresAt > now) {
    return terminalGrant.allowed;
  }
  const win = windowGetter();
  const options = {
    type: 'question' as const,
    title: 'Интерактивный терминал',
    message: 'Разрешить запуск оболочки из WorkSpace?',
    detail:
      'Интерфейс просит открыть терминал. Команды выполняются с правами вашей учётной записи. ' +
      'Разрешайте только если вы сами открыли терминал.',
    buttons: ['Запретить', 'Разрешить'],
    defaultId: 0,
    cancelId: 0,
  };
  const answer = win ? await dialog.showMessageBox(win, options) : await dialog.showMessageBox(options);
  const allowed = answer.response === 1;
  // Отказ тоже запоминается, но на то же время: иначе инъекция в документ
  // заваливает пользователя диалогами, а он не сможет разрешить терминал,
  // не перезапустив приложение.
  terminalGrant = { senderId, allowed, expiresAt: now + TERMINAL_GRANT_TTL_MS };
  return allowed;
}

export function registerTerminal(windowGetter: () => BrowserWindow | null): void {
  ipcMain.handle('terminal:create', async (event, req?: { id?: string; cwd?: string; shell?: string }) => {
    const id = req?.id || 'term_' + Math.random().toString(36).substring(2, 9);
    const requestedCwd = req?.cwd;
    // Рабочий каталог запрашивает интерфейс: вне открытой папки оболочка
    // стартовала бы в произвольном месте, поэтому такой запрос подменяем
    // корнем рабочей папки.
    const root = getWorkspaceRoot();
    const cwd =
      requestedCwd && fs.existsSync(requestedCwd) && (!root || isInsidePath(requestedCwd, root))
        ? requestedCwd
        : root || process.cwd();
    const isWin = os.platform() === 'win32';
    const shellExe = resolveShell(req?.shell);
    if (!shellExe) {
      return { success: false, error: 'Недопустимая оболочка. Доступны: powershell, pwsh, cmd, bash, sh, zsh, wsl' };
    }
    if (!(await confirmTerminal(windowGetter, event.sender.id))) {
      return { success: false, error: 'Запуск терминала запрещён пользователем' };
    }

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
