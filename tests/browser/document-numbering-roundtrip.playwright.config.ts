import { defineConfig } from '@playwright/test'
import productionOfficeIO from './production-office-io.playwright.config'

/** Uses the same production preview and real WASM artifacts as the Office IO suite. */
export default defineConfig({
  ...productionOfficeIO,
  testMatch: 'document-numbering-roundtrip.spec.ts',
  outputDir: '../../test-results/document-numbering-roundtrip',
})
