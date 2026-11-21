import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'vitest'

const panel = readFileSync(new URL('../src/components/spreadsheet/SheetDataValidationPanel.vue', import.meta.url), 'utf8')
const workbench = readFileSync(new URL('../src/components/ExcelWorkbench.vue', import.meta.url), 'utf8')

test('data validation uses an in-workbench sidebar with the fixed-version width and spacing', () => {
  assert.match(panel, /<aside v-if="open" class="validation-sidebar"/)
  assert.match(panel, /\.validation-sidebar\s*\{[^}]*flex:\s*0 0 312px/)
  assert.match(panel, /\.validation-header\s*\{[^}]*padding:\s*16px 16px 8px/)
  assert.match(panel, /\.validation-body\s*\{[^}]*padding:\s*0 16px 16px/)
  assert.doesNotMatch(panel, /<Drawer\b/)
  assert.match(workbench, /<div class="sheet-canvas-layout">[\s\S]*?<div class="sheet-canvas-stack">[\s\S]*?<SheetDataValidationPanel\b/)
})

test('validation details edit ranges and persist rule settings through the workbook command', () => {
  assert.match(panel, /v-model:value="rangeValue"/)
  assert.match(panel, /function parseRanges\(value: string\): CellRange\[\]/)
  assert.match(panel, /const updatedRule = createSheetDataValidationRule\([\s\S]*?ranges\)/)
  assert.match(panel, /executeCommandSync\('sheet\.command\.replace-data-validation'/)
  assert.match(panel, /class="validation-options"/)
  assert.match(panel, /v-model:value="ruleRenderMode"/)
  assert.match(panel, /executeCommandSync\('sheet\.command\.remove-data-validation',[\s\S]*?ruleId: current\.id/)
})
