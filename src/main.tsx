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

const container = document.getElementById('root');
if (!container) {
  throw new Error('Корневой элемент #root не найден в index.html');
}

createRoot(container).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
