import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

/** One server and isolated output for the complete toolbar integration suite. */
export default defineConfig({
  testDir: '.',
  testMatch: ['sheet-toolbar.spec.ts', 'sheet-data-toolbar.spec.ts', 'docs-toolbar-table.spec.ts', 'feature-panels.spec.ts'],
  outputDir: '../../test-results/toolbar-parity', workers: 1, fullyParallel: false,
  timeout: 45_000, expect: { timeout: 8_000 }, reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4203', browserName: 'chromium', channel: 'msedge', headless: true,
    viewport: { width: 1440, height: 1000 }, screenshot: 'off', video: 'off',
    permissions: ['clipboard-read', 'clipboard-write'],
    trace: { mode: 'retain-on-failure', screenshots: false, snapshots: true, sources: false },
  },
  webServer: {
    command: 'pnpm --filter @office-studio/workbench dev --host 127.0.0.1 --port 4203 --strictPort',
    url: 'http://127.0.0.1:4203/', reuseExistingServer: !process.env.CI, timeout: 30_000,
    cwd: fileURLToPath(new URL('../..', import.meta.url)),
  },
})
