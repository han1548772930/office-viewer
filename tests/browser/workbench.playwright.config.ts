import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

const workspaceRoot = fileURLToPath(new URL('../..', import.meta.url))

export default defineConfig({
  testDir: '.',
  testMatch: 'document-workbench.spec.ts',
  timeout: 60_000,
  expect: { timeout: 10_000 },
  reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4191',
    browserName: 'chromium',
    channel: 'msedge',
    headless: true,
    viewport: { width: 1440, height: 1000 },
  },
  webServer: {
    command: 'pnpm --filter @office-studio/workbench dev -- --host 127.0.0.1 --port 4191 --strictPort',
    url: 'http://127.0.0.1:4191/',
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
    cwd: workspaceRoot,
  },
})
