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
  const paste = tabs[0].groups[0].items.find(item => item.commandId === 'sheet.command.paste')
  assert.deepEqual(paste.children.map(item => item.commandId), [
    'sheet.command.paste', 'sheet.command.paste-values', 'sheet.command.paste-formats',
    'sheet.command.paste-column-widths', 'sheet.command.paste-except-borders',
  ])
  assert.ok(items.some(item => item.commandId === 'sheet.operation.open-range-protection'))
  const leafItems = items.filter(item => item.type !== 'menu')
  assert.ok(leafItems.every(item => /^sheet\.(command|operation)\./.test(item.commandId)))
  assert.ok(!items.some(item => item.title === '剪切' || item.title === '复制'))
  assert.ok(!items.some(item => item.title === '打印'))
})

test('the View tab groups display and visibility controls and binds zoom to the sheet view', () => {
  const view = createSheetRibbon(state).find(tab => tab.id === 'view')
  assert.deepEqual(view.groups.map(group => group.title), ['显示', '可见性'])
  assert.deepEqual(view.groups[0].items.map(item => item.title), ['冻结窗格', '缩放'])
  assert.deepEqual(view.groups[1].items.map(item => item.title), ['网格线', '全屏'])
  const zoom = view.groups[0].items.find(item => item.id === 'sheet.operation.set-zoom')
  assert.equal(zoom.value, 100)
  assert.deepEqual(zoom.options.map(option => option.value), [50, 75, 100, 125, 150, 175, 200, 400])
  assert.deepEqual(ribbonInvocation(zoom, 150), { commandId: 'sheet.operation.set-zoom', params: 150 })
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

test('formula categories expose function dropdowns and insert functions directly', () => {
  const items = flatten(createSheetRibbon(state))
  const picker = items.find(item => item.id === 'sheet.operation.open-function')
  const categories = items.filter(item => item.id.startsWith('sheet.operation.open-function.category-'))
  const supportedCategories = [...new Set(SHEET_FUNCTION_CATALOG.map(entry => entry.category))]

  assert.equal(picker.type, 'button')
  assert.deepEqual(ribbonInvocation(picker), { commandId: 'sheet.operation.open-function', params: undefined })
  assert.deepEqual(categories.map(item => item.title), supportedCategories)
  assert.ok(categories.every(item => item.type === 'menu' && item.children.length > 0))
  for (const category of categories) {
    const functions = SHEET_FUNCTION_CATALOG.filter(entry => entry.category === category.title)
    assert.deepEqual(category.children.map(item => item.title), functions.map(entry => entry.name))
    assert.ok(category.children.every(item => item.commandId === 'sheet.operation.insert-function'))
  }
})

test('data validation ribbon entry opens the worksheet rule panel operation', () => {
  const items = flatten(createSheetRibbon(state))
  const validation = items.find(item => item.id === 'sheet.operation.open-data-validation')
  assert.equal(validation.disabled, false)
  assert.deepEqual(ribbonInvocation(validation), { commandId: 'sheet.operation.open-data-validation', params: undefined })
})

test('pivot chart is reachable from data analysis only when pivot tables are available', () => {
  const items = flatten(createSheetRibbon({ ...state, pivotChartAvailable: true }))
  const pivotChart = items.find(item => item.id === 'sheet.operation.open-pivot-chart')
  assert.equal(pivotChart.disabled, false)
  assert.deepEqual(ribbonInvocation(pivotChart), { commandId: 'sheet.operation.open-pivot-chart', params: undefined })
})
