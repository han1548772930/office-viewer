import { existsSync } from 'node:fs'
import { defineConfig } from '@playwright/test'

const systemBrowser = [
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  'C:/Program Files/Microsoft/Edge/Application/msedge.exe',
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
].find((path) => existsSync(path))

export default defineConfig({
  testDir: './packages/office-studio/tests/browser',
  testMatch: '**/*.spec.mjs',
  outputDir: './test-results/canvas-contracts',
  reporter: 'list',
  use: {
    baseURL: 'http://localhost:5177',
    browserName: 'chromium',
    viewport: { width: 1440, height: 900 },
    deviceScaleFactor: 1,
    launchOptions: {
      ...(systemBrowser ? { executablePath: systemBrowser } : {}),
    },
  },
  webServer: {
    command: 'pnpm --filter @office-studio/workbench exec vite --host localhost --port 5177 --strictPort',
    url: 'http://localhost:5177',
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
  },
})
