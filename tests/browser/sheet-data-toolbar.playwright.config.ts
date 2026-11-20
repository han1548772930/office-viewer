import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

export default defineConfig({
  testDir: '.', testMatch: 'sheet-data-toolbar.spec.ts', outputDir: '../../test-results/sheet-data-toolbar',
  workers: 1, timeout: 45_000, expect: { timeout: 8_000 }, reporter: 'list',
  use: { baseURL: 'http://127.0.0.1:4201', browserName: 'chromium', channel: 'msedge', headless: true, viewport: { width: 1440, height: 1000 }, screenshot: 'off', video: 'off' },
  webServer: { command: 'pnpm --filter @office-studio/workbench dev --host 127.0.0.1 --port 4201 --strictPort', url: 'http://127.0.0.1:4201/', reuseExistingServer: !process.env.CI, timeout: 30_000, cwd: fileURLToPath(new URL('../..', import.meta.url)) },
})
