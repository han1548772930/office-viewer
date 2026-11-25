import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

const workspaceRoot = fileURLToPath(new URL('../..', import.meta.url))

export default defineConfig({
  testDir: '.',
  testMatch: ['compute-runtime.spec.ts', 'chart-chord-worker.spec.ts'],
  outputDir: '../../test-results/compute-runtime',
  timeout: 45_000,
  expect: { timeout: 10_000 },
  reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4178',
    browserName: 'chromium',
    channel: 'chrome',
    headless: true,
  },
  webServer: {
    command: 'pnpm exec vite --config tests/browser/vite.config.ts --host 127.0.0.1 --port 4178 --strictPort',
    url: 'http://127.0.0.1:4178/tests/browser/compute-runtime.html',
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
    cwd: workspaceRoot,
  },
})
