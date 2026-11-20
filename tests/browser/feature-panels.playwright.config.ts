import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

const workspaceRoot = fileURLToPath(new URL('../..', import.meta.url))

export default defineConfig({
  testDir: '.',
  testMatch: 'feature-panels.spec.ts',
  outputDir: '../../test-results/feature-panels',
  fullyParallel: false,
  workers: 1,
  timeout: 45_000,
  expect: { timeout: 8_000 },
  reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4197',
    browserName: 'chromium',
    channel: 'msedge',
    headless: true,
    viewport: { width: 1440, height: 1000 },
    screenshot: 'off',
    video: 'off',
    trace: { mode: 'retain-on-failure', screenshots: false, snapshots: true, sources: false },
  },
  webServer: {
    command: 'pnpm --filter @office-studio/workbench dev --host 127.0.0.1 --port 4197 --strictPort',
    url: 'http://127.0.0.1:4197/',
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
    cwd: workspaceRoot,
  },
})
