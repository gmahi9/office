import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  base: './',
  plugins: [react()],
  resolve: {
    alias: {
      'exceljs-browser': fileURLToPath(new URL('./vendor/exceljs.min.js', import.meta.url)),
    },
  },
  server: {
    host: '127.0.0.1',
    port: 43123,
  },
  preview: {
    host: '127.0.0.1',
    port: 43123,
  },
  build: {
    outDir: 'dist',
    sourcemap: false,
    chunkSizeWarningLimit: 4000,
  },
  test: {
    environment: 'node',
    include: ['src/**/*.test.ts'],
    setupFiles: ['src/test/setup.ts'],
  },
})
