import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
// Monaco подключается из локального бандла (без CDN) — до первого рендера
import './monacoSetup';
import './index.css';

// Тестовый мост для electron/checks/model-check: включается только явно.
if (new URLSearchParams(window.location.search).has('modelTest')) {
  void Promise.all([
    import('./lib/documentModel'),
    import('./lib/workspacePath'),
    import('./lib/documentHistory'),
    import('./lib/docx'),
    import('./lib/documentSession'),
    import('./lib/markdownOps'),
    import('./lib/customStyles'),
    import('./lib/aiDocument'),
    import('./lib/pptx'),
    import('./lib/pptxTemplate'),
    import('./lib/pptxDoc'),
    import('./lib/pptxBuilder'),
    import('./services/aiTools'),
    import('./services/aiActions'),
    import('./lib/imageOps'),
    import('./lib/explorerOps'),
    import('./lib/appearance'),
    import('./lib/notifyRules'),
    import('./lib/snapshot'),
    import('./lib/redact'),
    import('./lib/layout'),
    import('./lib/aiDiff'),
    import('./lib/pdfOps'),
    import('./lib/pdfAnnotate'),
    import('./lib/pdfRender'),
  ]).then(([model, workspacePath, history, docx, session, markdownOps, customStyles, aiDocument, pptx, pptxTemplate, pptxDoc, pptxBuilder, aiTools, aiActions, imageOps, explorerOps, appearance, notifyRules, snapshot, redact, layout, aiDiff, pdfOps, pdfAnnotate, pdfRender]) => {
    const target = window as unknown as {
      __documentModel?: typeof model;
      __workspacePath?: typeof workspacePath;
      __documentHistory?: typeof history;
      __docx?: typeof docx;
      __documentSession?: typeof session;
      __markdownOps?: typeof markdownOps;
      __customStyles?: typeof customStyles;
      __aiDocument?: typeof aiDocument;
      __pptx?: typeof pptx;
      __pptxTemplate?: typeof pptxTemplate;
      __pptxDoc?: typeof pptxDoc;
      __pptxBuilder?: typeof pptxBuilder;
      __aiTools?: typeof aiTools;
      __aiActions?: typeof aiActions;
      __imageOps?: typeof imageOps;
      __explorerOps?: typeof explorerOps;
      __appearance?: typeof appearance;
      __notifyRules?: typeof notifyRules;
      __snapshot?: typeof snapshot;
      __redact?: typeof redact;
      __layout?: typeof layout;
      __aiDiff?: typeof aiDiff;
      __pdfOps?: typeof pdfOps;
      __pdfAnnotate?: typeof pdfAnnotate;
      __pdfRender?: typeof pdfRender;
    };
    target.__documentModel = model;
    target.__workspacePath = workspacePath;
    target.__documentHistory = history;
    target.__docx = docx;
    target.__documentSession = session;
    target.__markdownOps = markdownOps;
    target.__customStyles = customStyles;
    target.__aiDocument = aiDocument;
    target.__pptx = pptx;
    target.__pptxTemplate = pptxTemplate;
    target.__pptxDoc = pptxDoc;
    target.__pptxBuilder = pptxBuilder;
    target.__aiTools = aiTools;
    target.__aiActions = aiActions;
    target.__imageOps = imageOps;
    target.__explorerOps = explorerOps;
    target.__appearance = appearance;
    target.__notifyRules = notifyRules;
    target.__snapshot = snapshot;
    target.__redact = redact;
    target.__layout = layout;
    target.__aiDiff = aiDiff;
    target.__pdfOps = pdfOps;
    target.__pdfAnnotate = pdfAnnotate;
    target.__pdfRender = pdfRender;
  });
}

interface AppErrorBoundaryState {
  error: Error | null;
}

/**
 * Ошибка рендера внутри App (парсеры xlsx/pdf/pptx, Monaco) размонтирует всё
 * дерево и оставляет белый экран. Граница ловит её и показывает причину с
 * возможностью перезагрузиться.
 */
class AppErrorBoundary extends React.Component<
  { children: React.ReactNode },
  AppErrorBoundaryState
> {
  state: AppErrorBoundaryState = { error: null };

  static getDerivedStateFromError(error: Error): AppErrorBoundaryState {
    return { error };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo): void {
    // Единственный след падения: в консоли он переживает перезагрузку.
    console.error('Необработанная ошибка интерфейса WorkSpace', error, info.componentStack);
  }

  render(): React.ReactNode {
    const { error } = this.state;
    if (!error) return this.props.children;

    return (
      <div className="flex h-screen w-screen flex-col items-center justify-center gap-4 bg-ws-panel p-8 text-center text-ws">
        <div className="w-full max-w-xl space-y-3">
          <h1 className="text-base font-semibold">Приложение остановилось из-за ошибки</h1>
          <p className="text-xs leading-relaxed text-ws-muted">
            Подробности уже записаны в консоль. Несохранённые изменения вкладок при
            перезагрузке потеряются.
          </p>
          <pre className="max-h-48 overflow-auto rounded border border-ws bg-ws-surface p-2 text-left font-mono text-[11px] whitespace-pre-wrap text-ws-danger">
            {error.message || String(error)}
          </pre>
          <button
            onClick={() => window.location.reload()}
            className="rounded bg-ws-accent px-3 py-1.5 text-xs font-medium text-white transition-opacity hover:opacity-90"
          >
            Перезагрузить приложение
          </button>
        </div>
      </div>
    );
  }
}

const container = document.getElementById('root');
if (!container) {
  throw new Error('Корневой элемент #root не найден в index.html');
}

createRoot(container).render(
  <React.StrictMode>
    <AppErrorBoundary>
      <App />
    </AppErrorBoundary>
  </React.StrictMode>
);
