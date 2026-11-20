import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

const workspaceRoot = fileURLToPath(new URL('../..', import.meta.url))

export default defineConfig({
  testDir: '.',
  testMatch: 'chart-workbench.spec.ts',
  outputDir: '../../test-results/chart-workbench',
  timeout: 60_000,
  expect: { timeout: 15_000 },
  reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4179',
    browserName: 'chromium',
    channel: 'msedge',
    headless: true,
  },
  webServer: {
    command: 'pnpm exec vite examples/workbench --config examples/workbench/vite.config.ts --host 127.0.0.1 --port 4179 --strictPort',
    url: 'http://127.0.0.1:4179/',
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
    cwd: workspaceRoot,
  },
})
