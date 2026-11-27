import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

export default defineConfig({
  testDir: '.', testMatch: 'sheet-view-toolbar.spec.ts', outputDir: '../../test-results/sheet-view-toolbar',
  workers: 1, timeout: 45_000, expect: { timeout: 8_000 }, reporter: 'list',
  use: { baseURL: 'http://127.0.0.1:4207', browserName: 'chromium', channel: 'msedge', headless: true,
    viewport: { width: 2200, height: 1000 }, screenshot: 'off', video: 'off', trace: 'off' },
  webServer: { command: 'pnpm --filter @office-studio/workbench dev --host 127.0.0.1 --port 4207 --strictPort',
    url: 'http://127.0.0.1:4207/', reuseExistingServer: !process.env.CI, timeout: 30_000,
    cwd: fileURLToPath(new URL('../..', import.meta.url)) },
})
