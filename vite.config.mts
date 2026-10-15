import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import path from 'path';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
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
