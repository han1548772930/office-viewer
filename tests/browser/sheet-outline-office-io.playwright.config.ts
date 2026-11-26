import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

export default defineConfig({
  testDir: '.', testMatch: 'sheet-outline-office-io.spec.ts', outputDir: '../../test-results/sheet-outline-office-io',
  timeout: 60_000, workers: 1, reporter: 'list',
  use: { baseURL: 'http://127.0.0.1:4199', browserName: 'chromium', channel: 'msedge', headless: true },
  // Deferred only: rebuild parser/writer WASM before invoking this config.
  webServer: { command: 'pnpm exec vite --host 127.0.0.1 --port 4199 --strictPort',
    url: 'http://127.0.0.1:4199/', reuseExistingServer: !process.env.CI,
    timeout: 30_000, cwd: fileURLToPath(new URL('../..', import.meta.url)) },
})
