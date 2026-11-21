import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  plugins: [vue()],
  test: {
    include: ['packages/office-studio/tests/**/*.browser.test.mjs'],
    fileParallelism: false,
    maxWorkers: 1,
    pool: 'forks',
    environment: 'node',
    testTimeout: 60_000,
    hookTimeout: 60_000,
  },
})
