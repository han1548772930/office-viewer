import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

const workspaceRoot = fileURLToPath(new URL('../..', import.meta.url))

/** Run after pnpm build; this server only serves production artifacts. */
export default defineConfig({
  testDir: '.',
  testMatch: ['production-office-io.spec.ts', 'office-worker-smoke.spec.ts'],
  outputDir: '../../test-results/production-office-io',
  fullyParallel: false,
  workers: 1,
  timeout: 90_000,
  expect: { timeout: 15_000 },
  reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4195',
    browserName: 'chromium',
    channel: 'msedge',
    headless: true,
    viewport: { width: 1440, height: 1000 },
    screenshot: 'off',
    video: 'off',
    trace: { mode: 'retain-on-failure', screenshots: false, snapshots: true, sources: false },
  },
  webServer: {
    command: 'pnpm --filter @office-studio/workbench preview --host 127.0.0.1 --port 4195 --strictPort',
    url: 'http://127.0.0.1:4195/',
    reuseExistingServer: false,
    timeout: 30_000,
    cwd: workspaceRoot,
  },
})
