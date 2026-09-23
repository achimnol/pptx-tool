/// <reference types="vitest/config" />
import react from '@vitejs/plugin-react';
import {defineConfig} from 'vite';

const backend = 'http://127.0.0.1:8000';

// The backend checks the Host and Origin headers, so let the dev proxy present itself as the backend.
const proxyOptions = {target: backend, changeOrigin: true, headers: {origin: backend}};

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: '../pptx_tool/web/static',
    emptyOutDir: true,
  },
  server: {
    proxy: {
      '/api': proxyOptions,
      '/schema': proxyOptions,
    },
  },
  test: {
    environment: 'jsdom',
    setupFiles: ['src/test/setup.ts'],
    css: false,
  },
});
