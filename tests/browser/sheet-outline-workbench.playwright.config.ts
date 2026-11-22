import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

const workspaceRoot = fileURLToPath(new URL('../..', import.meta.url))

export default defineConfig({
  testDir: '.',
  testMatch: 'sheet-outline-workbench.spec.ts',
  outputDir: '../../test-results/sheet-outline-workbench',
  timeout: 60_000,
  expect: { timeout: 15_000 },
  reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4180',
    browserName: 'chromium',
    channel: 'msedge',
    headless: true,
  },
  webServer: {
    command: 'pnpm exec vite examples/workbench --config examples/workbench/vite.config.ts --host 127.0.0.1 --port 4180 --strictPort',
    url: 'http://127.0.0.1:4180/',
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
    cwd: workspaceRoot,
  },
})
