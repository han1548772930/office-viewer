import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'
import { fileURLToPath } from 'node:url'

const librarySource = fileURLToPath(new URL('../../packages/office-studio/src/index.ts', import.meta.url))
const libraryStyles = fileURLToPath(new URL('../../packages/office-studio/src/style.css', import.meta.url))

export default defineConfig({
  plugins: [vue()],
  resolve: {
    dedupe: ['vue', 'pinia'],
    alias: [
      { find: '@office-studio/office-studio/style.css', replacement: libraryStyles },
      { find: '@office-studio/office-studio', replacement: librarySource },
    ],
  },
  server: { port: 5174 },
})
