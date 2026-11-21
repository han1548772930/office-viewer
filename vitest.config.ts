import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  plugins: [vue()],
  test: {
    include: ['packages/office-studio/tests/**/*.test.mjs'],
    exclude: ['**/*.browser.test.mjs'],
    environment: 'node',
    pool: 'forks',
    maxWorkers: 4,
    testTimeout: 30_000,
    hookTimeout: 30_000,
  },
})
