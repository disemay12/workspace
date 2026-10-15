/**
 * Подключение Monaco Editor из локального бандла.
 *
 * По умолчанию @monaco-editor/react грузит редактор с CDN (jsdelivr), что
 * неприемлемо для офлайн-приложения и блокируется Content-Security-Policy.
 * Здесь мы передаём загрузчику локальный экземпляр `monaco-editor` из
 * node_modules, поэтому Vite включает его в сборку и сеть не требуется.
 *
 * Дополнительно настраиваем web-worker'ы (подсветка, JSON, TS) — без них
 * Monaco выводит предупреждение и работает в урезанном режиме.
 */
import { loader } from '@monaco-editor/react';
import * as monaco from 'monaco-editor';

// Воркеры подключаются как модули — Vite собирает их в отдельные чанки.
// ВАЖНО: карта экспорта monaco-editor задаёт "./*" -> "./esm/vs/*.js",
// поэтому путь указывается БЕЗ префикса esm/vs — иначе он удвоится.
import editorWorker from 'monaco-editor/editor/editor.worker.js?worker';
import jsonWorker from 'monaco-editor/language/json/json.worker.js?worker';
import cssWorker from 'monaco-editor/language/css/css.worker.js?worker';
import htmlWorker from 'monaco-editor/language/html/html.worker.js?worker';
import tsWorker from 'monaco-editor/language/typescript/ts.worker.js?worker';

// Регистрируем фабрику воркеров под именами, которые ожидает Monaco
self.MonacoEnvironment = {
  getWorker(_moduleId: string, label: string) {
    switch (label) {
      case 'json':
        return new jsonWorker();
      case 'css':
      case 'scss':
      case 'less':
        return new cssWorker();
      case 'html':
      case 'handlebars':
      case 'razor':
        return new htmlWorker();
      case 'typescript':
      case 'javascript':
        return new tsWorker();
      default:
        return new editorWorker();
    }
  },
};

// Заменяем CDN-загрузчик на локальный экземпляр
loader.config({ monaco });

export { monaco };
