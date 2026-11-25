import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { applyExcelWorkbookView, sheetViewCellAt, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { initializeWorkbookCellStyles, refreshWorkbookCellStyles, resolveSheetCellStyle } from '../src/features/sheet-workbench/sheet-style-inheritance.ts'
import { RangeThemeStyle, normalizeSheetRangeThemeStyle } from '../src/features/sheet-workbench/sheet-range-theme-style.ts'
import { SHEET_BUILTIN_RANGE_THEME_NAMES, SHEET_BUILTIN_RANGE_THEMES } from '../src/features/sheet-workbench/sheet-range-theme-presets.ts'
import { captureSheetRangeThemeRegion, resolveSheetRangeThemeStyle, sheetRangeThemeByName, sheetRangeThemeHiddenRowSpans,
  sheetRangeThemeRegionStyle, sheetUsedRangeThemeRule } from '../src/features/sheet-workbench/sheet-range-theme-resolver.ts'
import { assertSheetRangeThemeRegionSnapshot, importSheetRangeThemeResources, normalizeWorkbookRangeThemes,
  removeSheetRangeThemeRules, setSheetRangeThemeRules, transformSheetRangeThemeRules } from '../src/features/sheet-workbench/sheet-range-theme-model.ts'
import { planSheetRangeThemeCommand, sheetRangeThemeCommandInput, SHEET_RANGE_THEME_COMMANDS } from '../src/features/sheet-workbench/sheet-range-theme-command.ts'
import { captureSheetStyleRegion, sheetStyleRegionCellStyle } from '../src/features/sheet-workbench/sheet-style-region.ts'
import { sheetStyleSelectionProbes } from '../src/features/sheet-workbench/sheet-style-selection.ts'
import { planWorkbookAxisStructureCommand } from '../src/features/sheet-workbench/sheet-axis-structure-command.ts'
import { planWorksheetCopy } from '../src/features/sheet-workbench/sheet-worksheet-copy.ts'

// Drafts only. No test, type, build, browser, Rust/WASM or benchmark execution in this implementation phase.
const range = (startRow = 0, endRow = startRow, startCol = 0, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
const theme = () => normalizeSheetRangeThemeStyle({ name: 'custom', wholeStyle: { fs: 20, cl: { rgb: '#101010' } },
  headerRowStyle: { bg: { rgb: '#112233' }, bl: 1 }, firstRowStyle: { bg: { rgb: '#224466' } },
  secondRowStyle: { bg: { rgb: '#ddeeff' } }, lastRowStyle: { bg: { rgb: '#445566' }, bl: 0 } })
const sample = () => {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'range-theme-contract'; model.calculationMode = 'manual'; model.rangeThemes = { custom: theme() }
  sheet.id = 'data'; model.activeSheetId = sheet.id; sheet.rowCount = 100; sheet.columnCount = 20; sheet.defaultRowAutoHeight = false
  sheet.rangeThemeRules = [{ id: 'custom-rule', themeName: 'custom', range: range(0, 8, 0, 3), metadata: { keep: true } }]
  sheet.cells['0:0'] = createCell(0, 0, 'source'); sheet.cells['0:0'].authoredStyle = {}
  initializeWorkbookCellStyles(model)
  return model
}

test('the fixed built-in catalog exposes 32 immutable configurations and the default footer border', () => {
  assert.equal(SHEET_BUILTIN_RANGE_THEME_NAMES.length, 32); assert.equal(new Set(SHEET_BUILTIN_RANGE_THEME_NAMES).size, 32)
  assert.deepEqual(SHEET_BUILTIN_RANGE_THEME_NAMES.slice(0, 2), ['default', 'default-last-row'])
  assert.equal(SHEET_BUILTIN_RANGE_THEMES['dark-blue'].headerRowStyle.fill, 'rgb(30, 66, 159)')
  assert.equal(SHEET_BUILTIN_RANGE_THEMES['default-last-row'].lastRowStyle.border.top.style, 'thin')
  assert.throws(() => { SHEET_BUILTIN_RANGE_THEMES.default.headerRowStyle.font.bold = false }, TypeError)
})

test('built-in names keep fixed priority when a custom catalog contains the same name', () => {
  const custom = normalizeSheetRangeThemeStyle({ name: 'default', headerRowStyle: { bg: { rgb: '#000000' } } })
  assert.equal(sheetRangeThemeByName('default', { rangeThemes: { default: custom } }), SHEET_BUILTIN_RANGE_THEMES.default)
  assert.equal(sheetRangeThemeByName('missing', { rangeThemes: {} }), undefined)
})

test('all nine builder sections read and serialize sparse public styles without adding defaults', () => {
  const style = new RangeThemeStyle('nine')
  for (const section of ['Whole', 'FirstRow', 'SecondRow', 'LastRow', 'FirstColumn', 'SecondColumn', 'LastColumn', 'HeaderRow', 'HeaderColumn']) {
    style[`set${section}Style`]({ bl: 0, bg: { rgb: '#123456' } })
    assert.deepEqual(style[`get${section}Style`](), { bl: 0, bg: { rgb: '#123456' } })
  }
  assert.equal(Object.keys(style.toJson()).length, 10); assert.equal(style.getName(), 'nine')
  assert.equal(style.toJson().wholeStyle.ff, undefined)
})

test('section precedence merges border edges while header, footer and explicit false override stripes', () => {
  const style = new RangeThemeStyle('corners', { wholeStyle: { bl: 1, bd: { t: { s: 1, cl: { rgb: '#111111' } }, b: { s: 1 } } },
    secondColumnStyle: { bg: { rgb: '#222222' } }, secondRowStyle: { bg: { rgb: '#333333' } },
    headerColumnStyle: { cl: { rgb: '#444444' } }, lastColumnStyle: { bd: { r: { s: 6 } } },
    headerRowStyle: { bg: { rgb: '#555555' } }, lastRowStyle: { bl: 0, bg: { rgb: '#666666' }, bd: { t: { s: 6 } } } })
  const result = style.getStyle(0, 0, true, true)
  assert.equal(result.bg.rgb, '#666666'); assert.equal(result.cl.rgb, '#444444'); assert.equal(result.bl, 0)
  assert.equal(result.bd.t.s, 6); assert.equal(result.bd.b.s, 1); assert.equal(result.bd.r.s, 6)
})

test('a valid empty section mask yields an empty style while absent masks yield null and malformed flags reject', () => {
  const style = new RangeThemeStyle('empty')
  assert.deepEqual(style.getStyle(0, 0, false, false), {}); assert.equal(style.getStyle(-1, -1, false, false), null)
  assert.deepEqual(style.getStyle(-2, -1, false, false), {}); assert.deepEqual(style.getStyle(-1, -1, true, false), {})
  assert.throws(() => style.getStyle(0.5, 0, false, false)); assert.throws(() => style.getStyle(0, 0, 1, false))
})

test('builder input and returned DTOs stay detached and invalid setters or fromJson leave the builder intact', () => {
  const input = { bg: { rgb: '#123456' }, ol: { vendor: { keep: true } } }, style = new RangeThemeStyle('detached', { wholeStyle: input })
  input.bg.rgb = '#000000'; const read = style.getWholeStyle(); read.bg.rgb = '#ffffff'
  assert.equal(style.getWholeStyle().bg.rgb, '#123456'); assert.deepEqual(style.toJson().wholeStyle.ol, { vendor: { keep: true } })
  const before = style.toJson(); assert.throws(() => style.setWholeStyle('catalog-reference')); assert.throws(() => style.fromJson({ name: '', wholeStyle: {} }))
  assert.throws(() => style.fromJson(null)); assert.throws(() => new RangeThemeStyle('bad', [])); assert.deepEqual(style.toJson(), before)
})

test('range styles follow axis defaults and precede cell theme and direct formatting on populated and blank cells', () => {
  const model = sample(), sheet = model.sheets[0], cell = sheet.cells['0:0']
  sheet.rowStyles = { 0: { font: { family: 'Georgia', size: 14 }, fill: '#abcdef' } }
  cell.themeStyle = { font: { italic: true } }; cell.authoredStyle = { font: { bold: false }, fill: null }
  refreshWorkbookCellStyles(model, undefined, false)
  assert.equal(cell.style.font.family, 'Georgia'); assert.equal(cell.style.font.size, 20); assert.equal(cell.style.font.italic, true)
  assert.equal(cell.style.font.bold, false); assert.equal(cell.style.fill, null)
  assert.equal(resolveSheetCellStyle(sheet, 0, 1, undefined, model).fill, '#112233'); assert.equal(Object.keys(sheet.cells).length, 1)
})

test('overlapping ranges use the later author rule and replacing an exact range keeps its identity', () => {
  const model = sample(), sheet = model.sheets[0]
  setSheetRangeThemeRules(sheet, [range(0, 1, 0, 1)], 'default', model)
  assert.equal(sheetUsedRangeThemeRule(sheet, range()).themeName, 'default'); assert.equal(resolveSheetRangeThemeStyle(sheet, 2, 0, model).fill, '#ddeeff')
  const id = sheet.rangeThemeRules.at(-1).id
  setSheetRangeThemeRules(sheet, [range(0, 1, 0, 1)], 'light-blue', model)
  assert.equal(sheet.rangeThemeRules.length, 2); assert.equal(sheet.rangeThemeRules.at(-1).id, id)
})

test('manual hiding, filtering, zero height and outline collapse combine once when advancing row stripes', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.hiddenRows = [1]; sheet.autoFilter = { filteredRows: [1, 2] }; sheet.rowHeights = { 2: 0 }
  sheet.outline = { groups: [{ id: 'g', axis: 'row', start: 3, end: 3, collapsed: true }] }
  assert.deepEqual(sheetRangeThemeHiddenRowSpans(sheet, model), [{ start: 1, end: 3, value: true }])
  assert.equal(resolveSheetRangeThemeStyle(sheet, 4, 0, model).fill, '#224466')
  assert.equal(resolveSheetRangeThemeStyle(sheet, 5, 0, model).fill, '#ddeeff')
})

test('the footer follows the last visible row while a hidden header never moves to the next row', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.rowHiddenSpans = [{ start: 7, end: 8, value: true }]; sheet.hiddenRows = [0]
  assert.equal(resolveSheetRangeThemeStyle(sheet, 6, 0, model).fill, '#445566')
  assert.equal(resolveSheetRangeThemeStyle(sheet, 1, 0, model).fill, '#ddeeff')
  assert.equal(resolveSheetRangeThemeStyle(sheet, 0, 0, model).fill, '#112233')
})

test('table filtering affects row stripes and hidden columns retain physical column phase', () => {
  const model = sample(), sheet = model.sheets[0]
  model.tables = [{ id: 't', sheetId: sheet.id, autoFilter: { filteredRows: [1] } }]
  model.rangeThemes.custom.firstColumnStyle = { fill: '#111111' }; model.rangeThemes.custom.secondColumnStyle = { fill: '#222222' }
  delete model.rangeThemes.custom.firstRowStyle; delete model.rangeThemes.custom.secondRowStyle
  sheet.hiddenColumns = [1]
  assert.equal(resolveSheetRangeThemeStyle(sheet, 2, 2, model).fill, '#222222')
  assert.deepEqual(sheetRangeThemeHiddenRowSpans(sheet, model), [{ start: 1, end: 1, value: true }])
})

test('a hidden million-row span stays compact and its footer lookup allocates no cells', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.rangeThemeRules[0].range.endRow = 1_048_575; sheet.rowHiddenSpans = [{ start: 1, end: 1_048_575, value: true }]
  assert.equal(resolveSheetRangeThemeStyle(sheet, 0, 0, model).fill, '#445566')
  assert.equal(sheetRangeThemeHiddenRowSpans(sheet, model).length, 1); assert.equal(Object.keys(sheet.cells).length, 1)
})

test('resolved region snapshots preserve source header, footer and filtered stripe phase after source closure', () => {
  const model = sample(), sheet = model.sheets[0]; sheet.rangeThemeRules[0].range.endRow = 5; sheet.filteredRows = [1, 3]
  const snapshot = captureSheetRangeThemeRegion(sheet, [0, 2, 5], [0, 2], model)
  assertSheetRangeThemeRegionSnapshot(snapshot, 3, 2)
  delete model.rangeThemes.custom; sheet.rangeThemeRules = []
  assert.equal(sheetRangeThemeRegionStyle(snapshot, 0, 0).fill, '#112233')
  assert.equal(sheetRangeThemeRegionStyle(snapshot, 1, 1).fill, '#224466')
  assert.equal(sheetRangeThemeRegionStyle(snapshot, 2, 0).fill, '#445566')
})

test('an unresolved upper rule blocks older themes in both live and detached region reads', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.rangeThemeRules.push({ id: 'missing', themeName: 'unregistered', range: range(0, 1, 0, 1) })
  const snapshot = captureSheetRangeThemeRegion(sheet, { start: 0, end: 2 }, { start: 0, end: 1 }, model)
  assert.equal(resolveSheetRangeThemeStyle(sheet, 0, 0, model), undefined)
  assert.equal(sheetRangeThemeRegionStyle(snapshot, 0, 0), undefined)
  assert.equal(sheetRangeThemeRegionStyle(snapshot, 2, 0).fill, '#ddeeff'); assertSheetRangeThemeRegionSnapshot(snapshot, 3, 2)
})

test('malformed region coordinates, themes and noncanonical hidden spans are rejected before use', () => {
  const model = sample(), snapshot = captureSheetRangeThemeRegion(model.sheets[0], [0, 1, 2], [0, 1], model)
  for (const change of [value => { value.rows = [0, 0, 2] }, value => { value.columns = { start: -1, end: 0 } },
    value => { value.rules[0].theme.name = '' }, value => { value.hiddenRowSpans = [{ start: 2, end: 2, value: true }, { start: 1, end: 1, value: true }] },
    value => { value.hiddenRowSpans = [{ start: 0, end: 1_048_576, value: true }] }]) {
    const invalid = cloneModel(snapshot); change(invalid); assert.throws(() => assertSheetRangeThemeRegionSnapshot(invalid, 3, 2))
  }
})

test('style-region and blank view projection carry effective themes without expanding cell maps', () => {
  const model = sample(), sheet = model.sheets[0], view = toExcelWorkbookView(model), region = captureSheetStyleRegion(sheet, { start: 0, end: 8 }, { start: 0, end: 3 }, model)
  assert.equal(sheetStyleRegionCellStyle(region, 1, 2).fill, '#224466'); assert.equal(sheetViewCellAt(view.sheets[0], 8, 2).background, '#445566')
  const reopened = workbookFromInput(applyExcelWorkbookView(model, view))
  assert.deepEqual(reopened.sheets[0].rangeThemeRules, sheet.rangeThemeRules); assert.equal(Object.keys(reopened.sheets[0].cells).length, 1)
})

test('mixed-format probes sample theme stripes, corners and hidden row transitions', () => {
  const model = sample(), view = toExcelWorkbookView(model).sheets[0]
  view.hiddenRows = [1, 2]
  const probes = sheetStyleSelectionProbes(view, [range(0, 8, 0, 3)])
  assert.equal(probes.complete, true); assert.ok(new Set(probes.cells.map(cell => cell.background)).size > 1)
  assert.ok(probes.cells.some(cell => cell.background === '#112233')); assert.ok(probes.cells.some(cell => cell.background === '#445566'))
})

test('an invalid later range or missing theme leaves the original author rules unchanged', () => {
  const model = sample(), sheet = model.sheets[0], before = cloneModel(sheet)
  assert.throws(() => setSheetRangeThemeRules(sheet, [range(10, 11), range(-1)], 'default', model))
  assert.throws(() => setSheetRangeThemeRules(sheet, [range()], 'unknown', model)); assert.deepEqual(sheet, before)
  assert.throws(() => removeSheetRangeThemeRules(sheet, [range(), range(1_048_576)], 'custom')); assert.deepEqual(sheet, before)
})

test('remove and null use whole intersecting rules and expose their full permission scope in the plan', () => {
  const model = sample(), sheet = model.sheets[0], before = cloneModel(sheet)
  const plan = planSheetRangeThemeCommand(sheet, [range(2, 2, 1, 1)], 'set', null, model)
  assert.equal(plan.changed, true); assert.deepEqual(plan.rangeThemeRules, []); assert.ok(plan.affected.some(item => item.endRow === 8 && item.endCol === 3))
  assert.deepEqual(sheet, before); removeSheetRangeThemeRules(sheet, [range(2, 2, 1, 1)], 'custom'); assert.deepEqual(sheet.rangeThemeRules, [])
})

test('same-range theme reuse and removal misses are content no-ops and command DTOs validate identities and revision', () => {
  const model = sample(), sheet = model.sheets[0]
  assert.equal(planSheetRangeThemeCommand(sheet, [range(0, 8, 0, 3)], 'set', 'custom', model).changed, false)
  assert.equal(planSheetRangeThemeCommand(sheet, [range(20)], 'remove', 'custom', model).changed, false)
  assert.deepEqual(sheetRangeThemeCommandInput({ range: { startRow: 0, endRow: 1, startColumn: 2, endColumn: 3 }, themeName: null }).ranges, [range(0, 1, 2, 3)])
  for (const input of [{ unitId: '' }, { themeName: 1 }, { ranges: 'A1' }, { expectedRevision: -1 }]) assert.throws(() => sheetRangeThemeCommandInput(input))
})

test('workbook normalization prepares every catalog and rule before publishing any normalized field', () => {
  const model = sample(), other = createBlankSheet('Other'); other.rangeThemeRules = [{ id: 'bad', themeName: '', range: range() }]
  model.sheets.push(other); model.rangeThemes.custom = { name: 'custom', wholeStyle: { bg: { rgb: '#112233' } } }
  const before = cloneModel(model); assert.throws(() => normalizeWorkbookRangeThemes(model)); assert.deepEqual(model, before)
})

test('fixed snapshot resources import short styles and column coordinates while retaining original rule extensions', () => {
  const resource = { rangeThemeStyleMapJson: { custom: { name: 'custom', headerRowStyle: { bl: 0, bg: { rgb: '#abcdef' } } } },
    rangeThemeStyleRuleMap: { publicRule: { themeName: 'custom', rangeInfo: { subUnitId: 'data', range: { startRow: 0, endRow: 3, startColumn: 1, endColumn: 2 } }, vendor: { keep: 1 } } } }
  const model = workbookFromInput({ id: 'resource', calculationMode: 'manual', sheets: { data: { name: 'Data', rowCount: 20, columnCount: 10 } },
    resources: [{ name: 'SHEET_RANGE_THEME_MODEL_PLUGIN', data: JSON.stringify(resource) }] })
  assert.deepEqual(model.sheets[0].rangeThemeRules[0].range, range(0, 3, 1, 2))
  assert.deepEqual(model.sheets[0].rangeThemeRules[0].metadata.sourceRule.vendor, { keep: 1 })
  assert.equal(resolveSheetCellStyle(model.sheets[0], 0, 1, undefined, model).font.bold, false)
  assert.equal(resolveSheetCellStyle(model.sheets[0], 0, 1, undefined, model).fill, '#abcdef')
})

test('malformed imported catalogs, sheet targets and duplicate rule IDs reject atomically', () => {
  const model = sample(), before = cloneModel(model)
  const resource = (styles, rules) => [{ name: 'SHEET_RANGE_THEME_MODEL_PLUGIN', data: { rangeThemeStyleMapJson: styles, rangeThemeStyleRuleMap: rules } }]
  assert.throws(() => importSheetRangeThemeResources(model, resource({ mismatch: { name: 'other' } }, {})))
  const rule = sheetId => ({ themeName: 'custom', rangeInfo: { subUnitId: sheetId, range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } } })
  assert.throws(() => importSheetRangeThemeResources(model, resource({}, { missing: rule('absent') })))
  assert.throws(() => importSheetRangeThemeResources(model, resource({}, { 'custom-rule': rule('data') })))
  assert.deepEqual(model, before)
})

test('axis insertion expands author ranges, deletion removes exhausted ranges and copying creates unique rule IDs', () => {
  const model = sample(), sheet = model.sheets[0]; sheet.rangeThemeRules[0].range = range(0, 1, 0, 3)
  const copied = planWorkbookAxisStructureCommand(model, sheet.id, [{ axis: 'row', kind: 'copy', start: 0, count: 2, destination: 5 }]).workbook.sheets[0]
  assert.equal(copied.rangeThemeRules.length, 2); assert.equal(new Set(copied.rangeThemeRules.map(rule => rule.id)).size, 2)
  assert.ok(copied.rangeThemeRules.some(rule => rule.range.startRow === 5 && rule.range.endRow === 6))
  transformSheetRangeThemeRules(sheet, { axis: 'row', kind: 'insert', start: 1, count: 2 })
  assert.deepEqual(sheet.rangeThemeRules[0].range, range(0, 3, 0, 3))
  transformSheetRangeThemeRules(sheet, { axis: 'row', kind: 'delete', start: 0, count: 4 }); assert.deepEqual(sheet.rangeThemeRules, [])
})

test('axis movement splits disjoint theme ranges and a worksheet copy detaches configuration and regenerates identities', () => {
  const model = sample(), sheet = model.sheets[0]; sheet.rangeThemeRules[0].range = range(0, 4, 0, 3)
  transformSheetRangeThemeRules(sheet, { axis: 'row', kind: 'move', start: 0, count: 2, destination: 8 })
  assert.equal(sheet.rangeThemeRules.length, 2); assert.equal(new Set(sheet.rangeThemeRules.map(rule => rule.id)).size, 2)
  const plan = planWorksheetCopy(model, sheet.id, { userId: 'local-user', targetSheetId: 'copied' }), target = plan.workbook.sheets.find(item => item.id === 'copied')
  assert.deepEqual(target.rangeThemeRules.map(rule => rule.range), sheet.rangeThemeRules.map(rule => rule.range))
  assert.ok(target.rangeThemeRules.every(rule => !sheet.rangeThemeRules.some(source => source.id === rule.id)))
  target.rangeThemeRules[0].metadata.keep = false; assert.equal(sheet.rangeThemeRules[0].metadata.keep, true)
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const withStudio = run => {
  const studio = new OfficeStudio({ workbook: sample() }), api = FOffice.newAPI(studio), book = api.getActiveWorkbook()
  try { return run(studio, api, book) } finally { studio.dispose() }
}

test('Facade exposes fixed and custom catalogs and detached builders require explicit registration', () => withStudio((studio, _api, book) => {
  assert.equal(book.getRegisteredRangeThemes().length, 32); assert.deepEqual(book.getCustomRangeThemes(), ['custom'])
  const builder = book.createRangeThemeStyle('new', { headerRowStyle: { bg: { rgb: '#010203' } } })
  assert.equal(book.getRangeThemeStyle('new'), null); book.registerRangeTheme(builder); builder.setHeaderRowStyle({ bg: { rgb: '#999999' } })
  book.getActiveSheet().getRange('E1:E3').useThemeStyle('new')
  assert.equal(book.getActiveSheet().getRange('E1').getCellStyleData().bg.rgb, '#010203')
  assert.equal(studio.getWorkbookModel(book.getId()).rangeThemes.new.headerRowStyle.fill, '#010203')
}))

test('Facade application has one content history step and undo, redo and reopen preserve sparse theme rules', () => withStudio((studio, api, book) => {
  const before = cloneModel(book.save().sheets), sheet = book.getActiveSheet(), history = studio.editor.history.units[book.getId()]?.undo.length ?? 0
  sheet.getRange('E1:F4').useThemeStyle('default-last-row')
  assert.equal(studio.editor.history.units[book.getId()].undo.length, history + 1)
  assert.equal(sheet.getRange('E1').getUsedThemeStyle(), 'default-last-row'); assert.equal(sheet.getRange('E4').getCellStyleData().bl, 1)
  book.undo(); assert.deepEqual(book.save().sheets, before); book.redo()
  const reopened = api.createWorkbook({ ...book.save(), id: 'range-theme-reopened' })
  assert.equal(reopened.getActiveSheet().getRange('E1').getUsedThemeStyle(), 'default-last-row')
  assert.equal(Object.keys(reopened.save().sheets[0].cells).length, 1)
}))

test('custom re-registration updates live formatting while keeping cell direct overrides and author rule identity', () => withStudio((_studio, _api, book) => {
  const sheet = book.getActiveSheet(); sheet.getRange('B1').setBackground('#ff0000')
  const rules = cloneModel(book.save().sheets[0].rangeThemeRules)
  book.registerRangeTheme(book.createRangeThemeStyle('custom', { headerRowStyle: { bg: { rgb: '#abcdef' } } }))
  assert.equal(sheet.getRange('A1').getCellStyleData().bg.rgb, '#abcdef'); assert.equal(sheet.getRange('B1').getCellStyleData().bg.rgb, '#ff0000')
  assert.deepEqual(book.save().sheets[0].rangeThemeRules, rules); book.undo()
  assert.equal(sheet.getRange('A1').getCellStyleData().bg.rgb, '#112233')
}))

test('Facade named removal and null remove entire intersecting rules and leave unrelated themes intact', () => withStudio((_studio, _api, book) => {
  const sheet = book.getActiveSheet(); sheet.getRange('A1:B2').useThemeStyle('default')
  sheet.getRange('A1').removeThemeStyle('custom')
  assert.equal(sheet.getRange('A1').getUsedThemeStyle(), 'default'); assert.equal(sheet.getRange('A4').getUsedThemeStyle(), undefined)
  sheet.getRange('A1').useThemeStyle(null); assert.equal(sheet.getRange('B2').getUsedThemeStyle(), undefined)
}))

test('unregister retains unresolved author rules, suppresses their rendering and re-register restores them', () => withStudio((_studio, _api, book) => {
  const sheet = book.getActiveSheet(), rules = cloneModel(book.save().sheets[0].rangeThemeRules)
  book.unregisterRangeTheme('custom'); assert.equal(sheet.getRange('A1').getUsedThemeStyle(), 'custom')
  assert.equal(sheet.getRange('A1').getCellStyleData().bg.rgb, null); assert.deepEqual(book.save().sheets[0].rangeThemeRules, rules)
  book.registerRangeTheme(book.createRangeThemeStyle('custom', { headerRowStyle: { bg: { rgb: '#778899' } } }))
  assert.equal(sheet.getRange('A1').getCellStyleData().bg.rgb, '#778899')
}))

test('removing from an editable subrange rejects if another part of the same theme rule is protected', () => withStudio((studio, _api, book) => {
  studio.editor.workbook.model.sheets[0].rangeProtectionRules = [{ id: 'locked', ranges: [range(8, 8, 3, 3)], allowedUserIds: ['owner'] }]
  const before = cloneModel(book.save()), revision = studio.getWorkbookRevision(book.getId()), count = studio.editor.history.units[book.getId()]?.undo.length ?? 0
  assert.throws(() => book.getActiveSheet().getRange('A2').useThemeStyle(null), /保护/)
  assert.deepEqual(book.save(), before); assert.equal(studio.getWorkbookRevision(book.getId()), revision)
  assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, count)
}))

test('registered commands reject stale revisions and missing themes without history, and same-range reuse is a no-op', () => withStudio((studio, _api, book) => {
  const target = { unitId: book.getId(), sheetId: 'data', ranges: [range(0, 8, 0, 3)], themeName: 'custom' }, before = cloneModel(book.save())
  assert.equal(studio.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.set, target), false)
  assert.throws(() => studio.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.set, { ...target, expectedRevision: studio.getWorkbookRevision(book.getId()) + 1 }), /已更改/)
  assert.throws(() => studio.executeCommandSync(SHEET_RANGE_THEME_COMMANDS.set, { ...target, themeName: 'missing' }))
  assert.deepEqual(book.save(), before); assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
}))

test('registration stores special string names as own catalog properties and rejects a reentrant builder result', () => withStudio((studio, _api, book) => {
  book.registerRangeTheme(book.createRangeThemeStyle('__proto__', { wholeStyle: { bg: { rgb: '#112233' } } }))
  assert.ok(Object.hasOwn(studio.getWorkbookModel(book.getId()).rangeThemes, '__proto__'))
  const builder = book.createRangeThemeStyle('reentrant'), serialize = builder.toJson.bind(builder)
  builder.toJson = () => { book.getActiveSheet().getRange('F1').setValue('observer'); return serialize() }
  assert.throws(() => book.registerRangeTheme(builder), /已更改/)
  assert.equal(book.getRangeThemeStyle('reentrant'), null); assert.equal(book.getActiveSheet().getRange('F1').getValue(), 'observer')
}))
