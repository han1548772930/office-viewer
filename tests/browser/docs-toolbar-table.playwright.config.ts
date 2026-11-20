import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

export default defineConfig({
  testDir: '.', testMatch: 'docs-toolbar-table.spec.ts', outputDir: '../../test-results/docs-toolbar-table',
  fullyParallel: false, workers: 1, timeout: 45_000, expect: { timeout: 8_000 }, reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4199', browserName: 'chromium', channel: 'msedge', headless: true,
    viewport: { width: 1440, height: 1000 }, permissions: ['clipboard-read', 'clipboard-write'],
    screenshot: 'off', video: 'off', trace: { mode: 'retain-on-failure', screenshots: false, snapshots: true, sources: false },
  },
  webServer: {
    command: 'pnpm --filter @office-studio/workbench dev --host 127.0.0.1 --port 4199 --strictPort',
    url: 'http://127.0.0.1:4199/', reuseExistingServer: !process.env.CI, timeout: 30_000,
    cwd: fileURLToPath(new URL('../..', import.meta.url)),
  },
})
