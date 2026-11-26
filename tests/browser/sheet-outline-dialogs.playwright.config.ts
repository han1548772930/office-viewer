import { defineConfig } from '@playwright/test'
import workbench from './sheet-outline-workbench.playwright.config'

export default defineConfig({ ...workbench, testMatch: 'sheet-outline-dialogs.spec.ts', outputDir: '../../test-results/sheet-outline-dialogs' })
