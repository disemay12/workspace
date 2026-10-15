import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
// Monaco подключается из локального бандла (без CDN) — до первого рендера
import './monacoSetup';
import './index.css';

const container = document.getElementById('root');
if (!container) {
  throw new Error('Корневой элемент #root не найден в index.html');
}

createRoot(container).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
