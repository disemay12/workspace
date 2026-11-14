import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import crypto from 'crypto';
import path from 'path';

/**
 * CSP без 'unsafe-inline' для собранной страницы.
 *
 * В index.html есть два намеренных инлайн-скрипта (тема до первой отрисовки и
 * перехват ошибок загрузки), поэтому 'unsafe-inline' в dev остаётся: HMR и
 * React-refresh вставляют свои инлайн-скрипты. В сборке каждый инлайн-скрипт
 * заменяется своим sha256, и 'unsafe-inline' из script-src убирается — тогда
 * `javascript:`-ссылки и инъекция скрипта в документ не исполняются вовсе.
 */
function cspInlineHashes(): Plugin {
  return {
    name: 'workspace-csp-inline-hashes',
    apply: 'build',
    enforce: 'post',
    transformIndexHtml(html) {
      const hashes: string[] = [];
      const inline = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g;
      for (const match of html.matchAll(inline)) {
        // Парсер HTML нормализует переводы строк до разбора, поэтому хеш
        // считается по тексту с LF: с CRLF (Windows-исходник) он не совпадал
        // и Chromium блокировал скрипт темы.
        const body = match[1].replace(/\r\n?/g, '\n');
        if (!body.trim()) continue;
        hashes.push(`'sha256-${crypto.createHash('sha256').update(body, 'utf8').digest('base64')}'`);
      }
      if (hashes.length === 0) return html;

      const replaced = html.replace(
        /(script-src[^;"]*)'unsafe-inline'/,
        (_all, prefix: string) => `${prefix}${hashes.join(' ')}`
      );
      if (replaced === html) {
        // CSP изменился — значит directive не найден; тихо оставлять
        // 'unsafe-inline' нельзя, иначе защита исчезает незаметно.
        throw new Error('CSP: не удалось заменить unsafe-inline в script-src');
      }
      return replaced;
    },
  };
}

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss(), cspInlineHashes()],
  resolve: {
    alias: {
      '@': path.resolve(import.meta.dirname, './src'),
    },
  },
  base: './',
  server: {
    host: '127.0.0.1',
    port: 5173,
    strictPort: true,
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    // Основной бандл тяжёлый из-за Monaco и SheetJS — разделяем на чанки.
    // Rolldown (Vite 8) требует manualChunks в виде функции.
    chunkSizeWarningLimit: 1800,
    rollupOptions: {
      output: {
        manualChunks(id: string) {
          if (id.includes('node_modules')) {
            if (id.includes('monaco-editor')) return 'monaco';
            if (id.includes('/xlsx/')) return 'xlsx';
            if (id.includes('/mammoth/') || id.includes('/jszip/')) return 'office';
            if (id.includes('/react-dom/') || id.includes('/react/')) return 'react';
            if (id.includes('lucide-react')) return 'icons';
            return 'vendor';
          }
          return undefined;
        },
      },
    },
  },
});
