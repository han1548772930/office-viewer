import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createSheetRibbon } from '../src/components/spreadsheet/sheet-ribbon.ts'
import { ribbonInvocation, visibleRibbonTabs } from '../src/platform/ribbon/ribbon-schema.ts'
import { SHEET_FUNCTION_CATALOG } from '../src/features/sheet-workbench/sheet-function-catalog.ts'

const state = {
  canUndo: false, canRedo: true, hasSheet: true, canMerge: false, canUnmerge: false,
  canDeleteRow: true, canDeleteColumn: true, chartAvailable: true, dataValidationAvailable: true, canSort: true,
  fontFamily: 'Aptos', fontSize: 11, numberFormat: 'general', bold: false, italic: true, underline: false,
  wrap: false, align: 'center', gridlines: true, frozen: false, zoom: 100,
}

function flatten(tabs) {
  return tabs.flatMap(tab => tab.groups.flatMap(group => group.items.flatMap(item => [item, ...(item.children ?? [])])))
}

test('the sheet Ribbon exposes only populated semantic categories and stable editor intents', () => {
  const tabs = visibleRibbonTabs(createSheetRibbon(state))
  assert.deepEqual(tabs.map(tab => tab.title), ['开始', '插入', '公式', '数据', '视图'])
  assert.deepEqual(tabs[0].groups.map(group => group.title), ['剪贴板', '字体', '布局', '数字', '工具'])
  const items = flatten(tabs)
  assert.equal(new Set(items.map(item => item.id)).size, items.length)
  const leafItems = items.filter(item => item.type !== 'menu')
  assert.ok(leafItems.every(item => /^sheet\.(command|operation)\./.test(item.commandId)))
  assert.ok(!items.some(item => item.title === '剪切' || item.title === '复制'))
  assert.ok(!items.some(item => item.title === '打印'))
})

test('format activation and mixed selectors reflect supplied selection state', () => {
  const items = flatten(createSheetRibbon({ ...state, fontFamily: null, fontSize: null, numberFormat: null, bold: undefined, align: undefined }))
  assert.equal(items.find(item => item.title === '字体').value, null)
  assert.equal(items.find(item => item.title === '字号').placeholder, '混合')
  assert.equal(items.find(item => item.title === '加粗').activated, false)
  assert.equal(items.find(item => item.title === '斜体').activated, true)
  assert.equal(items.find(item => item.id === 'sheet.command.align-center').activated, false)
})

test('font changes carry their own scalar without overwriting the other format controls', () => {
  const items = flatten(createSheetRibbon({ ...state, fontFamily: 'Imported Face', fontSize: 15 }))
  const font = items.find(item => item.title === '字体')
  const size = items.find(item => item.title === '字号')
  assert.ok(font.options.some(option => option.value === 'Imported Face'))
  assert.ok(size.options.some(option => option.value === 15))
  assert.deepEqual(ribbonInvocation(font, 'Arial'), { commandId: 'sheet.command.set-font-family', params: 'Arial' })
  assert.deepEqual(ribbonInvocation(size, 24), { commandId: 'sheet.command.set-font-size', params: 24 })
})

test('history, single-cell merges and unavailable chart contexts cannot dispatch', () => {
  const items = flatten(createSheetRibbon({ ...state, chartAvailable: false }))
  const undo = items.find(item => item.title === '撤销')
  const merge = items.find(item => item.id === 'sheet.command.merge')
  const chart = items.find(item => item.title === '图表')
  assert.equal(ribbonInvocation(undo), undefined)
  assert.equal(ribbonInvocation(merge), undefined)
  assert.equal(ribbonInvocation(chart), undefined)
  assert.match(chart.disabledReason, /工作簿/)
  assert.deepEqual(ribbonInvocation(items.find(item => item.title === '重做')), { commandId: 'sheet.command.redo', params: undefined })
})

test('nested fill, merge and freeze controls share the same command contract as their main actions', () => {
  const items = flatten(createSheetRibbon({ ...state, canMerge: true, canUnmerge: true, frozen: true }))
  assert.deepEqual(ribbonInvocation(items.find(item => item.id === 'sheet.menu.fill.#fff4cc')), { commandId: 'sheet.command.set-background', params: '#fff4cc' })
  assert.deepEqual(ribbonInvocation(items.find(item => item.id === 'sheet.command.merge-all')), { commandId: 'sheet.command.merge', params: undefined })
  assert.equal(items.find(item => item.title === '取消冻结').disabled, false)
  assert.equal(items.find(item => item.title === '网格线').activated, true)
})

test('formula categories route to the same supported-function picker operation', () => {
  const items = flatten(createSheetRibbon(state))
  const picker = items.find(item => item.id === 'sheet.operation.open-function')
  const categories = items.filter(item => item.id.startsWith('sheet.operation.open-function.category-'))
  const supportedCategories = [...new Set(SHEET_FUNCTION_CATALOG.map(entry => entry.category))]

  assert.equal(picker.type, 'split')
  assert.deepEqual(ribbonInvocation(picker), { commandId: 'sheet.operation.open-function', params: undefined })
  assert.deepEqual(categories.map(item => item.title), supportedCategories)
  assert.deepEqual(ribbonInvocation(categories.find(item => item.title === '逻辑')), {
    commandId: 'sheet.operation.open-function', params: '逻辑',
  })
  assert.deepEqual(ribbonInvocation(items.find(item => item.id === 'sheet.operation.open-function.all')),
    { commandId: 'sheet.operation.open-function', params: undefined })
})

test('data validation ribbon entry opens the worksheet rule panel operation', () => {
  const items = flatten(createSheetRibbon(state))
  const validation = items.find(item => item.id === 'sheet.operation.open-data-validation')
  assert.equal(validation.disabled, false)
  assert.deepEqual(ribbonInvocation(validation), { commandId: 'sheet.operation.open-data-validation', params: undefined })
})
