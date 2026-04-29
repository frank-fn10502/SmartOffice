import { fileURLToPath, URL } from 'node:url'

import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'
import ElementPlus from 'unplugin-element-plus/vite'

export default defineConfig({
  base: '/dist/',
  plugins: [
    vue(),
    ElementPlus(),
  ],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  server: {
    port: 5173,
    proxy: {
      '/api': 'http://localhost:2805',
      '/hub': {
        target: 'http://localhost:2805',
        ws: true,
      },
    },
  },
  build: {
    outDir: '../wwwroot/dist',
    emptyOutDir: true,
  },
})
