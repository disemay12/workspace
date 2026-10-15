/**
 * Прокси для запросов к AI-провайдерам.
 *
 * Зачем это нужно: браузерный `fetch` из рендерера подчиняется политике CORS.
 * Многие OpenAI-совместимые шлюзы (в т.ч. Starimg AI) не отдают заголовок
 * `Access-Control-Allow-Origin`, поэтому Chromium блокирует ответ с
 * `net::ERR_FAILED` — хотя сетевое соединение устанавливается успешно.
 *
 * Запросы выполняются в главном процессе через модуль `net` Electron,
 * который не подчиняется CORS. Ответ стримится в рендерер через IPC-события,
 * поэтому потоковый вывод сохраняется полностью.
 *
 * Ключ API передаётся из рендерера при каждом запросе и нигде не сохраняется
 * в главном процессе.
 */
import { ipcMain, net, BrowserWindow } from 'electron';

interface StreamRequest {
  /** Уникальный идентификатор запроса для сопоставления событий. */
  requestId: string;
  url: string;
  method: string;
  headers: Record<string, string>;
  body: string;
}

/** Активные запросы — позволяют прервать генерацию по кнопке «Стоп». */
const activeRequests = new Map<string, Electron.ClientRequest>();

function sendToRenderer(channel: string, payload: unknown) {
  for (const win of BrowserWindow.getAllWindows()) {
    if (!win.isDestroyed()) {
      win.webContents.send(channel, payload);
    }
  }
}

export function registerAiProxy() {
  /**
   * Запуск стримингового запроса к провайдеру.
   * Ответ приходит в рендерер событиями `ai:chunk` / `ai:end` / `ai:error`.
   */
  ipcMain.handle('ai:stream', async (_event, req: StreamRequest) => {
    const { requestId, url, method, headers, body } = req;

    if (!url) {
      sendToRenderer('ai:error', { requestId, message: 'Не указан URL запроса' });
      return { started: false };
    }

    try {
      const request = net.request({ method: method || 'POST', url, redirect: 'follow' });

      for (const [key, value] of Object.entries(headers || {})) {
        try {
          request.setHeader(key, value);
        } catch {
          // Некоторые заголовки Chromium может запрещать — пропускаем
        }
      }

      activeRequests.set(requestId, request);

      request.on('response', (response) => {
        const status = response.statusCode || 0;
        const responseHeaders: Record<string, string> = {};
        for (const [k, v] of Object.entries(response.headers)) {
          responseHeaders[k] = Array.isArray(v) ? v.join(', ') : String(v);
        }

        sendToRenderer('ai:response', { requestId, status, headers: responseHeaders });

        // Не-200 ответы читаем целиком, чтобы показать текст ошибки
        if (status < 200 || status >= 300) {
          let errBody = '';
          response.on('data', (chunk: Buffer) => {
            errBody += chunk.toString('utf-8');
          });
          response.on('end', () => {
            activeRequests.delete(requestId);
            sendToRenderer('ai:error', {
              requestId,
              status,
              message: errBody.slice(0, 2000),
            });
          });
          return;
        }

        response.on('data', (chunk: Buffer) => {
          sendToRenderer('ai:chunk', { requestId, chunk: chunk.toString('utf-8') });
        });

        response.on('end', () => {
          activeRequests.delete(requestId);
          sendToRenderer('ai:end', { requestId });
        });

        response.on('error', (err: Error) => {
          activeRequests.delete(requestId);
          sendToRenderer('ai:error', { requestId, message: err.message });
        });
      });

      request.on('error', (err: Error) => {
        activeRequests.delete(requestId);
        sendToRenderer('ai:error', {
          requestId,
          message: err.message,
          // Различаем сетевые сбои и ошибки API
          isNetworkError: true,
        });
      });

      if (body) request.write(body);
      request.end();

      return { started: true };
    } catch (err) {
      activeRequests.delete(requestId);
      sendToRenderer('ai:error', { requestId, message: (err as Error).message });
      return { started: false };
    }
  });

  /** Прерывание активного запроса (кнопка «Стоп»). */
  ipcMain.handle('ai:abort', (_event, requestId: string) => {
    const request = activeRequests.get(requestId);
    if (!request) return { aborted: false };
    try {
      request.abort();
    } catch {
      // уже завершён
    }
    activeRequests.delete(requestId);
    return { aborted: true };
  });

  /** Прогрев соединения — снижает задержку первого запроса. */
  ipcMain.handle('ai:probe', async (_event, url: string) => {
    return new Promise((resolve) => {
      try {
        const request = net.request({ method: 'HEAD', url });
        const timer = setTimeout(() => {
          try {
            request.abort();
          } catch {}
          resolve({ reachable: false, error: 'timeout' });
        }, 8000);

        request.on('response', (response) => {
          clearTimeout(timer);
          resolve({ reachable: true, status: response.statusCode });
        });
        request.on('error', (err: Error) => {
          clearTimeout(timer);
          resolve({ reachable: false, error: err.message });
        });
        request.end();
      } catch (err) {
        resolve({ reachable: false, error: (err as Error).message });
      }
    });
  });

  /**
   * Запрос списка доступных моделей от API провайдера.
   * Выполняется через модуль net главного процесса для обхода ограничений CORS.
   */
  ipcMain.handle(
    'ai:listModels',
    async (
      _event,
      req: { provider: string; baseUrl?: string; apiKey?: string }
    ): Promise<{ success: boolean; models: string[]; error?: string }> => {
      const { provider, baseUrl, apiKey } = req;

      function requestJson(url: string, headers: Record<string, string>): Promise<any> {
        return new Promise((resolve, reject) => {
          try {
            const request = net.request({ method: 'GET', url, redirect: 'follow' });
            for (const [k, v] of Object.entries(headers)) {
              try {
                request.setHeader(k, v);
              } catch {}
            }

            const timer = setTimeout(() => {
              try {
                request.abort();
              } catch {}
              reject(new Error('Превышен таймаут ответа API (12 сек)'));
            }, 12000);

            request.on('response', (response) => {
              let body = '';
              response.on('data', (chunk: Buffer) => {
                body += chunk.toString('utf-8');
              });
              response.on('end', () => {
                clearTimeout(timer);
                const status = response.statusCode || 0;
                if (status < 200 || status >= 300) {
                  let msg = `HTTP ${status}`;
                  try {
                    const parsed = JSON.parse(body);
                    msg = parsed?.error?.message || parsed?.message || msg;
                  } catch {}
                  return reject(new Error(msg));
                }
                try {
                  resolve(JSON.parse(body));
                } catch {
                  reject(new Error('Некорректный JSON в ответе API'));
                }
              });
              response.on('error', (err: Error) => {
                clearTimeout(timer);
                reject(err);
              });
            });

            request.on('error', (err: Error) => {
              clearTimeout(timer);
              reject(err);
            });

            request.end();
          } catch (err) {
            reject(err);
          }
        });
      }

      try {
        let models: string[] = [];

        if (provider === 'gemini') {
          const base = (baseUrl || 'https://generativelanguage.googleapis.com/v1beta').replace(/\/+$/, '');
          const keyParam = apiKey ? `?key=${encodeURIComponent(apiKey.trim())}` : '';
          const url = `${base}/models${keyParam}`;
          const headers: Record<string, string> = apiKey ? { 'x-goog-api-key': apiKey.trim() } : {};
          const data = await requestJson(url, headers);
          if (Array.isArray(data?.models)) {
            models = data.models
              .filter((m: any) => !m.supportedGenerationMethods || m.supportedGenerationMethods.includes('generateContent'))
              .map((m: any) => String(m.name || '').replace(/^models\//, ''))
              .filter(Boolean);
          }
        } else if (provider === 'anthropic') {
          const base = (baseUrl || 'https://api.anthropic.com/v1').replace(/\/+$/, '');
          const url = `${base}/models`;
          const headers: Record<string, string> = {
            'x-api-key': apiKey ? apiKey.trim() : '',
            'anthropic-version': '2023-06-01',
          };
          const data = await requestJson(url, headers);
          if (Array.isArray(data?.data)) {
            models = data.data.map((m: any) => m.id).filter(Boolean);
          }
        } else if (provider === 'ollama') {
          const base = (baseUrl || 'http://localhost:11434/v1').replace(/\/+$/, '');
          try {
            const data = await requestJson(`${base}/models`, {});
            if (Array.isArray(data?.data)) {
              models = data.data.map((m: any) => m.id || m.name).filter(Boolean);
            }
          } catch {}
          if (models.length === 0) {
            try {
              const rootUrl = base.replace(/\/v1\/?$/, '');
              const data = await requestJson(`${rootUrl}/api/tags`, {});
              if (Array.isArray(data?.models)) {
                models = data.models.map((m: any) => m.name || m.model).filter(Boolean);
              }
            } catch {}
          }
        } else {
          // OpenAI, DeepSeek, OpenRouter, Starimg, Custom и любые OpenAI-совместимые шлюзы
          const cleanBase = (baseUrl || 'https://api.openai.com/v1')
            .replace(/\/+$/, '')
            .replace(/\/chat\/completions$/, '');
          const url = `${cleanBase}/models`;
          const headers: Record<string, string> = { Accept: 'application/json' };
          if (apiKey && apiKey.trim()) {
            headers['Authorization'] = `Bearer ${apiKey.trim()}`;
          }
          if (provider === 'openrouter') {
            headers['HTTP-Referer'] = 'https://workspace.local';
            headers['X-Title'] = 'WorkSpace';
          }
          const data = await requestJson(url, headers);
          let raw: any[] = [];
          if (Array.isArray(data?.data)) raw = data.data;
          else if (Array.isArray(data?.models)) raw = data.models;
          else if (Array.isArray(data)) raw = data;

          models = raw
            .map((m: any) => (typeof m === 'string' ? m : m.id || m.name))
            .filter(Boolean);
        }

        const unique = Array.from(new Set(models)).sort((a, b) => a.localeCompare(b));
        return { success: true, models: unique };
      } catch (err) {
        return { success: false, models: [], error: (err as Error).message };
      }
    }
  );
}
