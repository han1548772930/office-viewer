import assert from 'node:assert/strict'
import { test } from 'vitest'
import { CalendarDays, Info, TextCursor } from 'lucide-vue-next'
import { createSheetRibbon } from '../src/components/spreadsheet/sheet-ribbon.ts'
import { placeRibbonGridItems, ribbonInvocation, visibleRibbonTabs } from '../src/platform/ribbon/ribbon-schema.ts'
import { SHEET_COMMON_FUNCTION_CATALOG, SHEET_FUNCTION_CATALOG } from '../src/features/sheet-workbench/sheet-function-catalog.ts'
import { QUICK_CONDITIONAL_FORMAT_COMMANDS as conditional } from '../src/features/conditional-formatting/quick-conditional-formatting.ts'
import { frozenPaneCountsForSelection } from '../src/features/sheet-workbench/sheet-workbench.ts'
import { createSheetOutlineState } from '../src/features/sheet-outline/sheet-outline.ts'

const state = {
  canUndo: false, canRedo: true, hasSheet: true, canMerge: false, canUnmerge: false, mergeActive: false,
  canDeleteRow: true, canDeleteColumn: true, chartAvailable: true, dataValidationAvailable: true, canSort: true,
  filterEnabled: false, filterCanToggle: true,
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
    'sheet.command.paste', 'sheet.command.paste-values', 'sheet.command.paste-formulas', 'sheet.command.paste-formats',
    'sheet.command.paste-column-widths', 'sheet.command.paste-except-borders',
  ])
  assert.ok(items.some(item => item.commandId === 'sheet.command.paste-formulas'))
  assert.ok(items.some(item => item.commandId === 'sheet.operation.open-range-protection'))
  const leafItems = items.filter(item => item.type !== 'menu')
  assert.ok(leafItems.every(item => /^sheet\.(command|operation)\./.test(item.commandId)))
  assert.ok(!items.some(item => item.title === '剪切' || item.title === '复制'))
  assert.ok(!items.some(item => item.title === '打印'))
})

test('Home clear-format uses its range capability independently of text-format controls and retains its source position', () => {
  for (const context of [
    { clearFormatEditable: undefined, fontEditable: true, hasSheet: true, disabled: false },
    { clearFormatEditable: true, fontEditable: false, hasSheet: true, disabled: false },
    { clearFormatEditable: false, fontEditable: true, hasSheet: true, disabled: true },
    { clearFormatEditable: true, fontEditable: true, hasSheet: false, disabled: true },
  ]) {
    const clipboard = createSheetRibbon({ ...state, ...context }).find(tab => tab.id === 'home').groups.find(group => group.id === 'clipboard')
    const clear = clipboard.items[3]
    assert.equal(clear.id, 'sheet.command.clear-format'); assert.equal(clear.disabled, context.disabled)
    assert.deepEqual(clear.gridLayout, { row: 2, column: 2, showLabel: false })
    assert.deepEqual(ribbonInvocation(clear), context.disabled ? undefined : { commandId: 'sheet.command.clear-format', params: undefined })
  }
})

test('Home content clear keeps its extension position and uses value-edit capability independently of text formatting', () => {
  for (const context of [
    { clearContentEditable: undefined, fontEditable: true, hasSheet: true, disabled: false },
    { clearContentEditable: true, fontEditable: false, hasSheet: true, disabled: false },
    { clearContentEditable: false, fontEditable: true, hasSheet: true, disabled: true },
    { clearContentEditable: true, fontEditable: true, hasSheet: false, disabled: true },
  ]) {
    const tools = createSheetRibbon({ ...state, ...context }).find(tab => tab.id === 'home').groups.find(group => group.id === 'tools')
    assert.deepEqual(tools.items.map(item => item.id), [
      'sheet.operation.open-range-protection', 'base-ui.operation.toggle-shortcut-panel',
      'ui.operation.open-feature-search', 'sheet.command.clear-content', 'sheet.operation.open-range-theme',
    ])
    const clear = tools.items[3]
    assert.equal(clear.disabled, context.disabled)
    assert.deepEqual(clear.gridLayout, { row: 1, column: 3, rowSpan: 2, showLabel: true })
    assert.deepEqual(ribbonInvocation(clear), context.disabled ? undefined : { commandId: 'sheet.command.clear-content', params: undefined })
    assert.equal(flatten(createSheetRibbon({ ...state, ...context })).find(item => item.id === 'sheet.command.toggle-bold').disabled, !context.hasSheet || !context.fontEditable)
  }
})

test('Home range protection requests a new rule and retains the fixed toolbar position', () => {
  const tools = createSheetRibbon(state).find(tab => tab.id === 'home').groups.find(group => group.id === 'tools')
  const protection = tools.items[0]
  assert.equal(protection.id, 'sheet.operation.open-range-protection')
  assert.deepEqual(protection.gridLayout, { row: 1, column: 1, rowSpan: 2, showLabel: true })
  assert.deepEqual(ribbonInvocation(protection), { commandId: 'sheet.operation.open-range-protection', params: { mode: 'add' } })
  const unavailable = createSheetRibbon({ ...state, hasSheet: false }).find(tab => tab.id === 'home').groups.find(group => group.id === 'tools').items[0]
  assert.equal(unavailable.disabled, true); assert.equal(ribbonInvocation(unavailable), undefined)
})

test('Home protection input capability blocks new-rule invocations and preserves the optional-state and existing toolbar contracts', () => {
  for (const context of [
    { rangeProtectionEditable: undefined, hasSheet: true, disabled: false },
    { rangeProtectionEditable: true, hasSheet: true, disabled: false },
    { rangeProtectionEditable: false, hasSheet: true, disabled: true },
    { rangeProtectionEditable: true, hasSheet: false, disabled: true },
  ]) {
    const protection = createSheetRibbon({ ...state, ...context }).find(tab => tab.id === 'home').groups.find(group => group.id === 'tools').items[0]
    assert.equal(protection.disabled, context.disabled)
    assert.deepEqual(protection.gridLayout, { row: 1, column: 1, rowSpan: 2, showLabel: true })
    assert.deepEqual(ribbonInvocation(protection), context.disabled ? undefined : { commandId: 'sheet.operation.open-range-protection', params: { mode: 'add' } })
    assert.equal(flatten(createSheetRibbon({ ...state, ...context })).find(item => item.id === 'sheet.command.toggle-bold').disabled, !context.hasSheet)
  }
})

test('Home range theme uses its input capability and keeps the existing extension position and invocation', () => {
  for (const context of [
    { rangeThemeEditable: undefined, hasSheet: true, disabled: false },
    { rangeThemeEditable: true, hasSheet: true, disabled: false },
    { rangeThemeEditable: false, hasSheet: true, disabled: true },
    { rangeThemeEditable: true, hasSheet: false, disabled: true },
  ]) {
    const tools = createSheetRibbon({ ...state, ...context }).find(tab => tab.id === 'home').groups.find(group => group.id === 'tools')
    const rangeTheme = tools.items[4]
    assert.equal(rangeTheme.id, 'sheet.operation.open-range-theme'); assert.equal(rangeTheme.disabled, context.disabled)
    assert.deepEqual(rangeTheme.gridLayout, { row: 1, column: 4, rowSpan: 2, showLabel: true })
    assert.deepEqual(ribbonInvocation(rangeTheme), context.disabled ? undefined : { commandId: 'sheet.operation.open-range-theme', params: undefined })
    assert.equal(tools.items[0].disabled, !context.hasSheet)
  }
})

test('number-format input capability disables all five roots and their intents without changing source positions or font state', () => {
  for (const context of [
    { numberFormatEditable: undefined, disabled: false },
    { numberFormatEditable: true, disabled: false },
    { numberFormatEditable: false, disabled: true },
    { numberFormatEditable: true, fontEditable: false, disabled: true },
    { numberFormatEditable: true, hasSheet: false, disabled: true },
  ]) {
    const tabs = createSheetRibbon({ ...state, ...context })
    const number = tabs.find(tab => tab.id === 'home').groups.find(group => group.id === 'number')
    assert.deepEqual(number.items.map(item => item.id), [
      'sheet.command.set-number-format', 'sheet.number-format.percent', 'sheet.number-format.currency',
      'sheet.command.increase-decimals', 'sheet.command.decrease-decimals',
    ])
    assert.deepEqual(number.items.map(item => item.gridLayout), [
      { row: 1, column: 1, columnSpan: 4, width: 160 },
      { row: 2, column: 1 }, { row: 2, column: 2 }, { row: 2, column: 3 }, { row: 2, column: 4 },
    ])
    assert.ok(number.items.every(item => item.disabled === context.disabled))
    assert.deepEqual(ribbonInvocation(number.items[0], 'percent'), context.disabled ? undefined
      : { commandId: 'sheet.command.set-number-format', params: { pattern: '0.00%' } })
    const more = number.items[0].options.find(option => option.commandId === 'sheet.operation.open-custom-number-format')
    assert.deepEqual(ribbonInvocation(number.items[0], more.value), context.disabled ? undefined
      : { commandId: 'sheet.operation.open-custom-number-format', params: more.value })
    for (const [index, commandId] of [
      'sheet.command.set-percent-format', 'sheet.command.set-currency-format', 'sheet.command.increase-decimals', 'sheet.command.decrease-decimals',
    ].entries()) {
      assert.deepEqual(ribbonInvocation(number.items[index + 1]), context.disabled ? undefined : { commandId, params: undefined })
    }
    const bold = flatten(tabs).find(item => item.id === 'sheet.command.toggle-bold')
    assert.equal(bold.disabled, context.hasSheet === false || context.fontEditable === false)
  }
})

test('border input capability disables the split root and all fallback choices while retaining font editing and source visibility', () => {
  for (const context of [
    { borderEditable: undefined, disabled: false },
    { borderEditable: true, disabled: false },
    { borderEditable: false, disabled: true },
    { borderEditable: true, fontEditable: false, disabled: true },
    { borderEditable: true, hasSheet: false, disabled: true },
  ]) {
    const font = createSheetRibbon({ ...state, ...context, borderPreset: 'outside' }).find(tab => tab.id === 'home').groups.find(group => group.id === 'font')
    const root = font.items[10]
    assert.equal(root.id, 'sheet.menu.borders'); assert.equal(root.type, 'split')
    assert.deepEqual(root.gridLayout, { row: 2, column: 5 })
    assert.equal(root.popupCommandId, 'sheet.command.set-borders'); assert.equal(root.popupProps.preset, 'outside')
    assert.equal(root.children.length, 17)
    assert.equal(root.children[15].children.length, 12); assert.equal(root.children[16].children.length, 45)
    const choices = root.children.flatMap(item => [item, ...(item.children ?? [])])
    assert.ok([root, ...choices].every(item => item.disabled === context.disabled))
    assert.deepEqual(ribbonInvocation(root), context.disabled ? undefined : { commandId: 'sheet.command.set-borders', params: 'outside' })
    assert.ok(choices.filter(item => item.type !== 'menu').every(item => Boolean(ribbonInvocation(item)) === !context.disabled))
    assert.equal(font.items[4].disabled, context.hasSheet === false || context.fontEditable === false)
  }
  for (const drawingFocused of [false, true]) {
    const tabs = createSheetRibbon({ ...state, drawingFocused })
    const font = tabs.find(tab => tab.id === 'home').groups.find(group => group.id === 'font')
    assert.equal(font.items[10].hidden, drawingFocused)
    assert.equal(visibleRibbonTabs(tabs).find(tab => tab.id === 'home').groups.find(group => group.id === 'font').items.some(item => item.id === 'sheet.menu.borders'), !drawingFocused)
  }
})

test('the View tab groups display and visibility controls and binds zoom to the sheet view', () => {
  const view = createSheetRibbon(state).find(tab => tab.id === 'view')
  assert.deepEqual(view.groups.map(group => group.title), ['显示', '可见性'])
  assert.deepEqual(view.groups[0].items.map(item => item.title), ['冻结窗格', '缩放', '十字高亮'])
  assert.deepEqual(view.groups[1].items.map(item => item.title), ['网格线', '全屏'])
  const crosshair = view.groups[0].items.find(item => item.id === 'sheet.operation.toggle-crosshair-highlight')
  assert.deepEqual(ribbonInvocation(crosshair), { commandId: 'sheet.operation.toggle-crosshair-highlight', params: undefined })
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

test('merge activation follows the primary cell independently of merge capability', () => {
  for (const context of [
    { canMerge: true, canUnmerge: true, mergeActive: false, disabled: false },
    { canMerge: true, canUnmerge: true, mergeActive: true, disabled: false },
    { canMerge: false, canUnmerge: false, mergeActive: true, disabled: true },
  ]) {
    const merge = flatten(createSheetRibbon({ ...state, ...context })).find(item => item.id === 'sheet.menu.merge')
    assert.equal(merge.activated, context.mergeActive)
    assert.equal(merge.disabled, context.disabled)
  }
})

test('merge input capability disables the parent and four commands while retaining layout, activation and range rules', () => {
  for (const context of [
    { mergeEditable: undefined, disabled: [false, false, false, false, false] },
    { mergeEditable: true, disabled: [false, false, false, false, false] },
    { mergeEditable: false, disabled: [true, true, true, true, true] },
    { mergeEditable: true, fontEditable: false, disabled: [true, true, true, true, true] },
    { mergeEditable: true, hasSheet: false, disabled: [true, true, true, true, true] },
    { mergeEditable: true, canMerge: false, disabled: [false, true, true, true, false] },
    { mergeEditable: true, canUnmerge: false, disabled: [false, false, false, false, true] },
    { mergeEditable: true, canMerge: false, canUnmerge: false, disabled: [true, true, true, true, true] },
  ]) {
    const tabs = createSheetRibbon({ ...state, canMerge: true, canUnmerge: true, mergeActive: true, ...context })
    const layout = tabs.find(tab => tab.id === 'home').groups.find(group => group.id === 'layout')
    const merge = layout.items[5]
    assert.equal(merge.id, 'sheet.menu.merge'); assert.equal(merge.activated, true)
    assert.deepEqual(merge.gridLayout, { row: 2, column: 3 })
    assert.deepEqual([merge, ...merge.children].map(item => item.disabled), context.disabled)
    assert.deepEqual(merge.children.map(item => item.id), [
      'sheet.command.merge-all', 'sheet.command.merge-vertical', 'sheet.command.merge-horizontal', 'sheet.command.unmerge',
    ])
    for (const [index, invocation] of [
      { commandId: 'sheet.command.merge', params: 'all' },
      { commandId: 'sheet.command.merge', params: 'vertical' },
      { commandId: 'sheet.command.merge', params: 'horizontal' },
      { commandId: 'sheet.command.unmerge', params: undefined },
    ].entries()) {
      assert.deepEqual(ribbonInvocation(merge.children[index]), context.disabled[index + 1] ? undefined : invocation)
    }
  }
})

test('shrink-to-fit leaves the toggle value to the shared command for uniform and mixed selections', () => {
  for (const shrinkToFit of [true, false, undefined]) {
    const shrink = flatten(createSheetRibbon({ ...state, shrinkToFit })).find(item => item.id === 'sheet.command.set-shrink-to-fit')
    assert.deepEqual(ribbonInvocation(shrink), { commandId: 'sheet.command.set-shrink-to-fit', params: undefined })
    assert.equal(shrink.activated, shrinkToFit === true)
  }
})

test('drawing focus hides the five source layout roots while retaining the shrink-to-fit action', () => {
  const roots = drawingFocused => visibleRibbonTabs(createSheetRibbon({ ...state, drawingFocused }))
    .find(tab => tab.id === 'home').groups.find(group => group.id === 'layout').items
  assert.deepEqual(roots(false).map(item => item.id), [
    'sheet.menu.horizontal-align', 'sheet.menu.vertical-align', 'sheet.menu.text-wrap',
    'sheet.command.set-shrink-to-fit', 'sheet.menu.text-rotation', 'sheet.menu.merge',
  ])
  assert.deepEqual(roots(true).map(item => item.id), ['sheet.command.set-shrink-to-fit'])
  assert.deepEqual(ribbonInvocation(roots(true)[0]), { commandId: 'sheet.command.set-shrink-to-fit', params: undefined })
})

test('shrink activation and permission follow the primary wrap state independently of its stored shrink value', () => {
  for (const context of [
    { wrap: true, shrinkToFit: true, fontEditable: true, disabled: true, activated: false },
    { wrap: false, shrinkToFit: true, fontEditable: true, disabled: false, activated: true },
    { wrap: false, shrinkToFit: true, fontEditable: false, disabled: true, activated: true },
  ]) {
    const shrink = flatten(createSheetRibbon({ ...state, ...context })).find(item => item.id === 'sheet.command.set-shrink-to-fit')
    assert.equal(shrink.disabled, context.disabled)
    assert.equal(shrink.activated, context.activated)
  }
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

test('font-size stepper hints and permissions retain the shared command and grid positions', () => {
  for (const [commandId, shortcut, column] of [
    ['sheet.command.increase-font-size', 'Ctrl+Shift+.', 6],
    ['sheet.command.decrease-font-size', 'Ctrl+Shift+,', 7],
  ]) {
    const item = flatten(createSheetRibbon(state)).find(item => item.id === commandId)
    assert.equal(item.shortcut, shortcut)
    assert.deepEqual(item.gridLayout, { row: 1, column })
    assert.deepEqual(ribbonInvocation(item), { commandId, params: undefined })
    for (const context of [{ hasSheet: false }, { fontEditable: false }]) {
      assert.equal(ribbonInvocation(flatten(createSheetRibbon({ ...state, ...context })).find(item => item.id === commandId)), undefined)
    }
  }
})

test('history, single-cell merges and unavailable chart contexts cannot dispatch', () => {
  const items = flatten(createSheetRibbon({ ...state, chartAvailable: false }))
  const undo = items.find(item => item.title === '撤销')
  const merge = items.find(item => item.id === 'sheet.menu.merge')
  const mergeAll = items.find(item => item.id === 'sheet.command.merge-all')
  const chart = items.find(item => item.title === '图表')
  assert.equal(ribbonInvocation(undo), undefined)
  assert.equal(ribbonInvocation(merge), undefined)
  assert.equal(ribbonInvocation(mergeAll), undefined)
  assert.equal(merge.disabled, true)
  assert.equal(ribbonInvocation(chart), undefined)
  assert.match(chart.disabledReason, /工作簿/)
  assert.deepEqual(ribbonInvocation(items.find(item => item.title === '重做')), { commandId: 'sheet.command.redo', params: undefined })
})

test('nested fill, merge and freeze controls share the same command contract as their main actions', () => {
  const items = flatten(createSheetRibbon({ ...state, canMerge: true, canUnmerge: true, frozen: true }))
  assert.deepEqual(ribbonInvocation(items.find(item => item.id === 'sheet.menu.fill.#fff4cc')), { commandId: 'sheet.command.set-background', params: '#fff4cc' })
  assert.deepEqual(ribbonInvocation(items.find(item => item.id === 'sheet.command.merge-all')), { commandId: 'sheet.command.merge', params: 'all' })
  assert.equal(items.find(item => item.title === '取消冻结').disabled, false)
  assert.equal(items.find(item => item.title === '网格线').activated, true)
})

test('freezing a selected range keeps both the rows above and columns to its left frozen', () => {
  assert.deepEqual(frozenPaneCountsForSelection({ startRow: 4, startColumn: 3, endRow: 8, endColumn: 7 }), { rows: 4, columns: 3 })
  assert.deepEqual(frozenPaneCountsForSelection({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 }), { rows: 0, columns: 0 })
})

test('formula categories expose populated dropdowns and dispatch the selected function name', () => {
  const items = flatten(createSheetRibbon(state))
  const picker = items.find(item => item.id === 'sheet.operation.open-function')
  const categories = items.filter(item => item.id.startsWith('sheet.operation.open-function.category-'))
  const supportedCategories = [
    '常用与统计', '财务', '逻辑', '文本', '日期与时间', '查找与引用', '数学', '统计', '工程', '信息', '数据库',
  ].filter(category => category === '常用与统计' || SHEET_FUNCTION_CATALOG.some(entry => entry.category === category))

  assert.equal(picker.type, 'button')
  assert.deepEqual(ribbonInvocation(picker), { commandId: 'sheet.operation.open-function', params: undefined })
  assert.deepEqual(categories.map(item => item.title), supportedCategories)
  assert.ok(categories.every(item => item.type === 'menu' && item.children.length > 0))
  assert.equal(categories.find(item => item.title === '文本').icon, TextCursor)
  assert.equal(categories.find(item => item.title === '日期与时间').icon, CalendarDays)
  for (const [category, name] of [['文本', 'LEFT'], ['文本', 'MID'], ['文本', 'RIGHT'], ['文本', 'VALUE'], ['日期与时间', 'WEEKDAY']]) {
    const entry = categories.find(item => item.title === category).children.find(item => item.title === name)
    assert.deepEqual(ribbonInvocation(entry), { commandId: 'formula-ui.operation.insert-function', params: { value: name } })
  }
  for (const name of [
    'ACOS', 'ACOSH', 'ACOT', 'ACOTH', 'ASIN', 'ASINH', 'ATAN', 'ATANH', 'DEGREES', 'PI', 'RADIANS',
    'SIN', 'COS', 'TAN', 'COT', 'SEC', 'CSC', 'SINH', 'COSH', 'TANH', 'COTH', 'SECH', 'CSCH',
    'EXP', 'LN', 'LOG', 'LOG10', 'SIGN', 'ATAN2', 'EVEN', 'ODD', 'QUOTIENT', 'SQRTPI',
    'FACT', 'FACTDOUBLE', 'COMBIN', 'COMBINA', 'ROUND', 'ROUNDBANK', 'ROUNDDOWN', 'ROUNDUP', 'TRUNC',
    'CEILING', 'CEILING.MATH', 'CEILING.PRECISE', 'ISO.CEILING', 'FLOOR', 'FLOOR.MATH', 'FLOOR.PRECISE',
    'GCD', 'LCM', 'MULTINOMIAL', 'MROUND',
    'ABS', 'SQRT', 'INT', 'MOD', 'POWER', 'SUMSQ',
    'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'SUMPRODUCT', 'SERIESSUM',
    'ARABIC', 'ROMAN', 'BASE', 'DECIMAL',
    'RAND', 'RANDBETWEEN',
    'MDETERM',
  ]) {
    const entry = categories.find(item => item.title === '数学').children.find(item => item.title === name)
    assert.deepEqual(ribbonInvocation(entry), { commandId: 'formula-ui.operation.insert-function', params: { value: name } })
  }
  assert.equal(categories.at(-1).title, '信息')
  assert.equal(categories.at(-1).icon, Info)
  assert.deepEqual(categories.at(-1).children.map(item => item.title), [
    'ERROR.TYPE', 'ISBETWEEN', 'ISBLANK', 'ISDATE', 'ISEMAIL', 'ISERR', 'ISERROR', 'ISEVEN', 'ISFORMULA', 'ISLOGICAL', 'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISODD', 'ISREF', 'ISTEXT', 'ISURL', 'N', 'NA', 'SHEET', 'SHEETS', 'TYPE',
  ])
  assert.equal(categories[0].children.at(-1), picker)
  assert.deepEqual(categories[0].children.slice(0, -1).map(item => item.title), [
    'SUMIF', 'SUM', 'AVERAGE', 'IF', 'COUNT', 'SIN', 'MAX',
  ])
  assert.ok(categories.find(item => item.title === '数学').children.some(item => item.title === 'SUM'))
  assert.ok(['AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX'].every(name =>
    categories.find(item => item.title === '统计').children.some(item => item.title === name)))
  for (const name of ['AVEDEV', 'DEVSQ', 'GEOMEAN', 'HARMEAN', 'MODE.SNGL']) {
    const entry = categories.find(item => item.title === '统计').children.find(item => item.title === name)
    assert.deepEqual(ribbonInvocation(entry), { commandId: 'formula-ui.operation.insert-function', params: { value: name } })
  }
  const functionItems = categories.flatMap(category => category.children.filter(item => item.commandId === 'formula-ui.operation.insert-function'))
  assert.equal(functionItems.length, SHEET_FUNCTION_CATALOG.length + SHEET_COMMON_FUNCTION_CATALOG.length)
  assert.equal(new Set(functionItems.map(item => item.params.value)).size, SHEET_FUNCTION_CATALOG.length)
  for (const category of categories) {
    const functions = category === categories[0] ? SHEET_COMMON_FUNCTION_CATALOG
      : SHEET_FUNCTION_CATALOG.filter(entry => entry.category === category.title)
    const entries = category === categories[0] ? category.children.slice(0, -1) : category.children
    assert.deepEqual(entries.map(item => item.title), functions.map(entry => entry.name))
    assert.deepEqual(entries.map(item => ribbonInvocation(item)), functions.map(entry =>
      ({ commandId: 'formula-ui.operation.insert-function', params: { value: entry.name } })))
  }
  const unavailable = createSheetRibbon({ ...state, hasSheet: false }).find(tab => tab.id === 'formulas')
  assert.ok(unavailable.groups.flatMap(group => group.items).every(item => item.disabled
    && (item.children ?? []).every(child => child.disabled)))
})

test('calculation mode preserves its value and disables the selector without a worksheet', () => {
  for (const overrides of [{}, { calculationMode: 'manual' }, { hasSheet: false }]) {
    const item = flatten(createSheetRibbon({ ...state, ...overrides })).find(entry => entry.id === 'sheet.command.set-calculation-mode')
    assert.equal(item.value, overrides.calculationMode ?? 'automatic')
    assert.equal(item.disabled, overrides.hasSheet === false)
    assert.deepEqual(item.options, [{ label: '自动', value: 'automatic' }, { label: '手动', value: 'manual' }])
    assert.deepEqual(ribbonInvocation(item, 'manual'), item.disabled ? undefined
      : { commandId: 'sheet.command.set-calculation-mode', params: 'manual' })
  }
})

test('formula capability and drawing focus disable insertion without disabling workbook calculation', () => {
  for (const overrides of [{ formulaEditable: false }, { drawingFocused: true }]) {
    const formulas = createSheetRibbon({ ...state, ...overrides }).find(tab => tab.id === 'formulas')
    const categories = formulas.groups.find(group => group.id === 'formula-edit').items
    assert.ok(categories.every(item => item.disabled && item.children.every(child => child.disabled)))
    assert.equal(ribbonInvocation(categories[0].children.at(-1)), undefined)
    const calculation = formulas.groups.find(group => group.id === 'calculation-options').items
    assert.deepEqual(calculation.map(item => item.disabled), [false, true, false])
  }
  const enabled = createSheetRibbon({ ...state, formulaEditable: true, fontEditable: false }).find(tab => tab.id === 'formulas')
  assert.ok(enabled.groups.find(group => group.id === 'formula-edit').items.every(item => !item.disabled))
})

test('Insert cell tools fill two rows in source order and keep their quick-rule commands', () => {
  const insert = visibleRibbonTabs(createSheetRibbon(state)).find(tab => tab.id === 'insert')
  const cells = insert.groups.find(group => group.id === 'cells')
  assert.deepEqual(placeRibbonGridItems(cells.items).map(({ item, row, column, rowSpan }) =>
    [item.type, row, column, rowSpan]), [
    ['split', 1, 1, 1], ['split', 1, 2, 1], ['split', 1, 3, 1], ['menu', 2, 1, 1], ['split', 2, 2, 1],
  ])
  const [checkbox, dropdown, date, dataBar, iconSet] = cells.items
  assert.deepEqual([checkbox, date].map(item => ribbonInvocation(item).params), ['checkbox', 'date'])
  assert.deepEqual(dropdown.children.map(item => ribbonInvocation(item).params), ['dropdown-yes-no', 'dropdown-status', 'dropdown-abc', 'edit-dropdown', 'clear-dropdown'])
  assert.equal(dropdown.popupCommandId, 'sheet.operation.quick-data-validation')
  assert.deepEqual([dataBar.order, iconSet.order], [3, 4])
  assert.deepEqual(ribbonInvocation(dataBar), { commandId: conditional.dataBar, params: {} })
  assert.deepEqual(ribbonInvocation(iconSet), { commandId: conditional.iconSet, params: { iconSet: '3Stars' } })
  assert.deepEqual(dataBar.children.map(item => item.commandId), [conditional.dataBar, 'sheet.operation.open-conditional-formatting', conditional.clear])
  assert.deepEqual(iconSet.children.map(item => item.commandId),
    [conditional.iconSet, conditional.iconSet, conditional.iconSet, 'sheet.operation.open-conditional-formatting', conditional.clear])
  assert.equal(dataBar.children.at(-1).disabled, true)
  assert.equal(iconSet.children.at(-1).disabled, true)
})

test('quick conditional menus reflect active rule and type-specific clear state', () => {
  const quickConditional = { hasDataBar: true, hasIconSet: true, canClearDataBar: true, canClearIconSet: true, canClearRange: true, canClearWorksheet: true, activeIconSet: '3TrafficLights1' }
  const roots = value => createSheetRibbon({ ...state, ...value, quickConditional }).find(tab => tab.id === 'insert')
    .groups.find(group => group.id === 'cells').items.slice(-2)
  const [dataBar, iconSet] = roots({})
  assert.equal(dataBar.activated, true)
  assert.equal(iconSet.activated, true)
  assert.deepEqual(iconSet.children.slice(0, 3).map(item => item.activated), [false, true, false])
  assert.deepEqual(ribbonInvocation(iconSet.children.at(-1)), { commandId: conditional.clear, params: { types: ['iconSet'] } })
  assert.ok(roots({ fontEditable: false }).every(item => item.disabled && item.children.every(child => child.disabled)))
})

test('data validation ribbon entry opens the worksheet rule panel operation', () => {
  const items = flatten(createSheetRibbon(state))
  const validation = items.find(item => item.id === 'sheet.operation.open-data-validation')
  assert.equal(validation.disabled, false)
  assert.deepEqual(ribbonInvocation(validation), { commandId: 'sheet.operation.open-data-validation', params: undefined })
})

test('conditional formatting management exposes the ten source actions with shared commands and scope permissions', () => {
  const root = value => flatten(createSheetRibbon({ ...state, ...value })).find(item => item.id === 'sheet.operation.open-conditional-formatting')
  const management = root({ quickConditional: { canClearRange: true, canClearWorksheet: false } })
  assert.equal(management.type, 'menu')
  assert.deepEqual(management.gridLayout, { row: 2, column: 1, showLabel: true })
  assert.deepEqual(management.children.map(item => ribbonInvocation(item)), [
    ...['highlightCell', 'top10', 'expression', 'colorScale', 'dataBar', 'iconSet', undefined, 'create'].map(params =>
      ({ commandId: 'sheet.operation.open-conditional-formatting', params })),
    { commandId: conditional.clear, params: {} },
    { commandId: conditional.clearWorksheet, params: {} },
  ])
  assert.equal(management.children.at(-2).disabled, false)
  assert.equal(management.children.at(-1).disabled, true)
  assert.ok(root({ fontEditable: false }).children.every(item => item.disabled))
})

test('filter split menu toggles auto-filter and exposes clear, reapply, and panel commands', () => {
  const filter = flatten(createSheetRibbon(state)).find(item => item.id === 'sheet.data.filter')
  assert.equal(filter.type, 'split')
  assert.deepEqual(ribbonInvocation(filter), { commandId: 'sheet.command.toggle-auto-filter', params: undefined })
  assert.deepEqual(filter.children.map(item => item.commandId), [
    'sheet.command.clear-filter', 'sheet.command.reapply-filter', 'sheet.operation.open-data-panel',
  ])
  assert.ok(filter.children.every(item => item.disabled))

  const enabled = flatten(createSheetRibbon({ ...state, filterEnabled: true, filterCanToggle: true })).find(item => item.id === 'sheet.data.filter')
  assert.equal(enabled.activated, true)
  assert.deepEqual(enabled.children.map(item => item.disabled), [true, true, false])
  const filtered = flatten(createSheetRibbon({ ...state, filterEnabled: true, filterHasCriteria: true, filterCanToggle: true })).find(item => item.id === 'sheet.data.filter')
  assert.ok(filtered.children.every(item => !item.disabled))
})

test('data tools follow the source group order and row or column structure stays in header menus', () => {
  const tabs = createSheetRibbon(state)
  const insert = tabs.find(tab => tab.id === 'insert')
  const data = tabs.find(tab => tab.id === 'data')
  assert.deepEqual(data.groups.map(group => group.id), ['organization', 'rules', 'pivot', 'data-tools'])
  assert.deepEqual(data.groups.map(group => group.title), ['数据整理', '规则', '分析', '工具'])
  assert.deepEqual(data.groups.find(group => group.id === 'organization').items.map(item => item.title), ['排序', '筛选', '组合', '分组显示', '大纲设置'])
  assert.ok(data.groups.find(group => group.id === 'rules').items.some(item => item.id === 'sheet.operation.open-conditional-formatting'))
  assert.ok(!tabs.some(tab => tab.groups.some(group => group.items.some(item => /^sheet\.command\.(insert|delete)-(row|column)$/.test(item.commandId)))))
  assert.ok(insert.groups.some(group => group.items.some(item => item.commandId === 'sheet.operation.insert-chart')))
})

test('outline details follow grouping candidates, collapse no-ops, and selection scope', () => {
  const context = { unitId: 'outline-book', sheetId: 'outline-sheet', userId: 'visitor',
    selections: [{ startRow: 2, endRow: 4, startCol: 3, endCol: 5 }],
    sheet: { rowCount: 20, columnCount: 10, outline: createSheetOutlineState() } }
  const menu = outlineContext => flatten(createSheetRibbon({ ...state, outlineContext })).find(item => item.id === 'sheet.menu.outline-details')
  assert.equal(menu(undefined).disabled, true)
  const empty = menu(context)
  for (const suffix of ['rows', 'columns']) {
    assert.equal(empty.children.find(item => item.id === `sheet.command.outline-selected-${suffix}`).disabled, false)
    assert.equal(empty.children.find(item => item.id === `sheet.command.clear-selected-${suffix}-outlines`).disabled, true)
    for (const operation of ['toggle', 'collapse', 'expand']) {
      assert.equal(empty.children.find(item => item.id === `sheet.operation.outline-${operation}-${suffix}`).disabled, true)
    }
  }
  const grouped = { ...context, sheet: { ...context.sheet, outline: { ...context.sheet.outline, groups: [
    { id: 'outer-row', axis: 'row', start: 1, end: 8, collapsed: false },
    { id: 'selected-column', axis: 'column', start: 3, end: 5, collapsed: false },
  ] } } }
  const expanded = menu(grouped)
  assert.equal(expanded.children.find(item => item.id === 'sheet.command.clear-selected-rows-outlines').disabled, true)
  assert.equal(expanded.children.find(item => item.id === 'sheet.command.clear-selected-columns-outlines').disabled, false)
  assert.deepEqual(ribbonInvocation(expanded.children.find(item => item.id === 'sheet.command.outline-selected-rows')),
    { commandId: 'sheet.command.outline-selected-rows', params: { value: 'row' } })
  const collapsed = menu({ ...grouped, sheet: { ...grouped.sheet, outline: { ...grouped.sheet.outline,
    groups: grouped.sheet.outline.groups.map(group => ({ ...group, collapsed: true })) } } })
  for (const suffix of ['rows', 'columns']) {
    assert.equal(expanded.children.find(item => item.id === `sheet.operation.outline-toggle-${suffix}`).disabled, false)
    assert.equal(expanded.children.find(item => item.id === `sheet.operation.outline-collapse-${suffix}`).disabled, false)
    assert.equal(expanded.children.find(item => item.id === `sheet.operation.outline-expand-${suffix}`).disabled, true)
    assert.equal(collapsed.children.find(item => item.id === `sheet.operation.outline-collapse-${suffix}`).disabled, true)
    assert.equal(collapsed.children.find(item => item.id === `sheet.operation.outline-expand-${suffix}`).disabled, false)
  }
  const multiple = menu({ ...grouped, selections: [...grouped.selections, { startRow: 10, endRow: 12, startCol: 6, endCol: 8 }] })
  assert.ok(multiple.children.filter(item => item.type === 'button').every(item => item.disabled))
  assert.equal(multiple.children.find(item => item.id === 'sheet.menu.outline-summary-rows').children
    .find(item => item.id === 'sheet.menu.outline-summary-above').disabled, false)
})

test('outline menu permissions include the changed group footprint and retain an editable inner group', () => {
  const context = { unitId: 'outline-book', sheetId: 'outline-sheet', userId: 'visitor',
    selections: [{ startRow: 2, endRow: 4, startCol: 3, endCol: 5 }],
    sheet: { rowCount: 20, columnCount: 10, outline: { ...createSheetOutlineState(), groups: [
      { id: 'outer-row', axis: 'row', start: 1, end: 8, collapsed: false },
      { id: 'selected-column', axis: 'column', start: 3, end: 5, collapsed: false },
    ] }, rangeProtectionRules: [{ id: 'range-guard', permissionId: 'range-guard', ownerUserId: 'owner', editState: 'onlyMe',
      viewState: 'othersCanView', allowedUserIds: [], ranges: [{ startRow: 7, endRow: 7, startCol: 8, endCol: 8 }] }] } }
  const menu = outlineContext => flatten(createSheetRibbon({ ...state, outlineContext })).find(item => item.id === 'sheet.menu.outline-details')
  const guarded = menu(context)
  const rowToggle = guarded.children.find(item => item.id === 'sheet.operation.outline-toggle-rows')
  assert.equal(guarded.disabled, false)
  assert.equal(rowToggle.disabled, true); assert.match(rowToggle.disabledReason, /受保护/)
  assert.equal(guarded.children.find(item => item.id === 'sheet.operation.outline-collapse-rows').disabled, true)
  assert.equal(guarded.children.find(item => item.id === 'sheet.operation.outline-collapse-columns').disabled, false)
  assert.equal(guarded.children.find(item => item.id === 'sheet.menu.outline-summary-rows').children
    .find(item => item.id === 'sheet.menu.outline-summary-above').disabled, true)
  assert.equal(guarded.children.find(item => item.id === 'sheet.menu.outline-summary-columns').children
    .find(item => item.id === 'sheet.menu.outline-summary-left').disabled, false)
  const nested = menu({ ...context, sheet: { ...context.sheet, outline: { ...context.sheet.outline,
    groups: [...context.sheet.outline.groups, { id: 'inner-row', axis: 'row', start: 2, end: 4, collapsed: false }] } } })
  assert.equal(nested.children.find(item => item.id === 'sheet.operation.outline-toggle-rows').disabled, false)
  assert.equal(nested.children.find(item => item.id === 'sheet.menu.outline-summary-rows').children
    .find(item => item.id === 'sheet.menu.outline-summary-above').disabled, true)
  const protectedSheet = menu({ ...context, sheet: { ...context.sheet,
    worksheetProtection: { id: 'worksheet-guard', unitType: 'worksheet', permissionId: 'worksheet-guard', ownerUserId: 'owner',
      editState: 'onlyMe', viewState: 'othersCanView', allowedUserIds: [] } } })
  assert.equal(protectedSheet.disabled, true)
  assert.ok(protectedSheet.children.flatMap(item => item.children ?? [item]).every(item => item.disabled))
})

test('pivot chart is reachable from data analysis only when pivot tables are available', () => {
  const tabs = visibleRibbonTabs(createSheetRibbon({ ...state, pivotChartAvailable: true }))
  const items = flatten(tabs)
  const analysis = tabs.find(tab => tab.id === 'data').groups.find(group => group.id === 'pivot')
  assert.deepEqual(placeRibbonGridItems(analysis.items).map(p => [p.item.commandId, p.row, p.column, p.rowSpan]), [
    ['sheet.operation.open-pivot', 1, 1, 1], ['sheet.operation.open-slicer', 1, 2, 1],
    ['sheet.operation.open-pivot-chart', 2, 1, 1], ['sheet.operation.open-timeline', 2, 2, 1],
  ])
  const media = tabs.find(tab => tab.id === 'insert').groups.find(group => group.id === 'media')
  assert.ok(media.items.some(item => item.commandId === 'sheet.operation.open-sparkline'))
  const pivotChart = items.find(item => item.id === 'sheet.operation.open-pivot-chart')
  assert.equal(pivotChart.disabled, false)
  assert.deepEqual(ribbonInvocation(pivotChart), { commandId: 'sheet.operation.open-pivot-chart', params: undefined })
})
