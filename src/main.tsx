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
  ]).then(([model, workspacePath, history, docx, session, markdownOps, customStyles, aiDocument]) => {
    const target = window as unknown as {
      __documentModel?: typeof model;
      __workspacePath?: typeof workspacePath;
      __documentHistory?: typeof history;
      __docx?: typeof docx;
      __documentSession?: typeof session;
      __markdownOps?: typeof markdownOps;
      __customStyles?: typeof customStyles;
      __aiDocument?: typeof aiDocument;
    };
    target.__documentModel = model;
    target.__workspacePath = workspacePath;
    target.__documentHistory = history;
    target.__docx = docx;
    target.__documentSession = session;
    target.__markdownOps = markdownOps;
    target.__customStyles = customStyles;
    target.__aiDocument = aiDocument;
  });
}

const container = document.getElementById('root');
if (!container) {
  throw new Error('Корневой элемент #root не найден в index.html');
}

createRoot(container).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
