import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { applyCellValue, workbookFromInput } from '../src/api/model-utils.ts'
import { applyExcelWorkbookView, sheetViewCellAt, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { applySheetFontFormatCommand, applySheetViewFontFormatCommand, fontFormatActionForCommand } from '../src/features/sheet-workbench/sheet-font-format-command.ts'
import { applySheetClearCommand, applySheetViewClearCommand } from '../src/features/sheet-workbench/sheet-clear-command.ts'
import { applySheetLayoutFormatCommand } from '../src/features/sheet-workbench/sheet-layout-format-command.ts'
import { initializeWorkbookCellStyles, normalizeSheetStyleInput, refreshWorkbookCellStyles, resolveSheetCellStyle,
  setSheetCellStyleProperties, transformSheetDefaultStyles } from '../src/features/sheet-workbench/sheet-style-inheritance.ts'
import { sheetStyleSelectionProbes } from '../src/features/sheet-workbench/sheet-style-selection.ts'
import { planSheetCellInput } from '../src/features/sheet-workbench/sheet-cell-input-command.ts'
import { assertSheetStyleRegion, captureSheetStyleRegion, sheetStyleRegionCellStyle } from '../src/features/sheet-workbench/sheet-style-region.ts'
import { captureSheetClipboard, createSheetClipboardPastePlan, sheetClipboardCellAt, sheetClipboardContent,
  sheetClipboardSourceFingerprint } from '../src/features/sheet-workbench/sheet-clipboard.ts'
import { applySheetFormatPainterCommand, applySheetViewFormatPainterCommand, captureSheetFormatPainterSource,
  captureSheetViewFormatPainterSource } from '../src/features/sheet-workbench/sheet-format-painter.ts'
import { applySheetBorderFormatCommand, applySheetViewBorderFormatCommand } from '../src/features/sheet-workbench/sheet-border-format-command.ts'
import { createSheetAutoFillPlan } from '../src/features/sheet-workbench/sheet-auto-fill.ts'

// Drafts only. The user deferred test, browser, type, build and WASM execution.
const range = (startRow = 0, endRow = startRow, startCol = 0, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
const sample = () => {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'default-style-contract'; model.calculationMode = 'manual'
  model.defaultStyle = { font: { family: 'Arial', size: 12 }, vertical: 'bottom' }
  model.styles = { row: { font: { family: 'Georgia', bold: true } }, theme: { font: { size: 18 }, fill: '#ffeecc' } }
  sheet.rowCount = 100; sheet.columnCount = 26
  sheet.defaultStyle = { font: { size: 14 }, horizontal: 'right' }
  sheet.columnStyles = { 0: { font: { family: 'Courier New', size: 16 }, fill: '#ddeeff' } }
  sheet.rowStyles = { 0: 'row' }
  sheet.cells['0:0'] = createCell(0, 0, 'abc')
  initializeWorkbookCellStyles(model)
  return model
}

test('workbook, worksheet, column, row, theme and direct formatting compose by property', () => {
  const model = sample(), sheet = model.sheets[0], cell = sheet.cells['0:0']
  cell.themeStyle = 'theme'; cell.authoredStyle = { font: { italic: true, color: null }, fill: null }
  refreshWorkbookCellStyles(model, undefined, false)
  assert.deepEqual(cell.style.font, { family: 'Georgia', size: 18, bold: true, italic: true, underline: false, strike: false, color: null })
  assert.equal(cell.style.fill, null); assert.equal(cell.style.vertical, 'bottom'); assert.equal(cell.style.horizontal, 'right')
})

test('column priority changes the intersection while retaining independent row properties', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.rowStylePriority = false
  const style = resolveSheetCellStyle(sheet, 0, 0, sheet.cells['0:0'], model)
  assert.equal(style.font.family, 'Courier New'); assert.equal(style.font.bold, true); assert.equal(style.font.size, 16)
})

test('direct boolean wrapping overrides inherited strategy while later strategy restores coherent wrapping', () => {
  const model = sample(), sheet = model.sheets[0], cell = sheet.cells['0:0']
  sheet.defaultStyle.wrapStrategy = 'wrap'; cell.authoredStyle = { wrap: false }
  refreshWorkbookCellStyles(model, undefined, false)
  assert.equal(cell.style.wrap, false); assert.equal(cell.style.wrapStrategy, 'overflow')
  cell.authoredStyle = { wrapStrategy: 'clip' }; sheet.defaultStyle.wrap = true
  refreshWorkbookCellStyles(model, undefined, false)
  assert.equal(cell.style.wrap, false); assert.equal(cell.style.wrapStrategy, 'clip')
})

test('explicit boolean wrap edits replace an earlier direct strategy and survive re-materialization', () => {
  const model = sample(), sheet = model.sheets[0], cell = sheet.cells['0:0']
  setSheetCellStyleProperties(cell, { wrapStrategy: 'wrap' })
  setSheetCellStyleProperties(cell, { wrap: false })
  refreshWorkbookCellStyles(model, undefined, false)
  assert.equal(cell.style.wrap, false); assert.equal(cell.style.wrapStrategy, 'overflow')
  assert.equal(cell.authoredStyle.wrap, false); assert.equal(cell.authoredStyle.wrapStrategy, 'overflow')
})

test('blank-cell view reads use inherited formatting without allocating document cells', () => {
  const model = sample(), view = toExcelWorkbookView(model).sheets[0], before = cloneModel(view.cells)
  const blank = sheetViewCellAt(view, 9, 0)
  assert.equal(blank.fontFamily, 'Courier New'); assert.equal(blank.fontSize, 16); assert.equal(blank.background, '#ddeeff')
  assert.deepEqual(view.cells, before); assert.equal(blank.rawValue, null)
})

test('family and size null reset clear direct and rich-run properties then inherit', () => {
  const model = sample(), sheet = model.sheets[0], cell = sheet.cells['0:0']
  cell.authoredStyle = { font: { family: 'Aptos', size: 22, italic: true } }
  cell.richText = [{ id: 'first', text: 'a', font: { family: 'Times New Roman', size: 30, underline: true } },
    { id: 'second', text: 'bc', hyperlink: 'https://example.test', vendor: { keep: 1 } }]
  refreshWorkbookCellStyles(model, undefined, false)
  applySheetFontFormatCommand(sheet, [range()], { kind: 'family', value: null }, 'local-user', [], undefined, model)
  applySheetFontFormatCommand(sheet, [range()], { kind: 'size', value: null }, 'local-user', [], undefined, model)
  assert.equal(cell.style.font.family, 'Georgia'); assert.equal(cell.style.font.size, 16)
  assert.equal(cell.authoredStyle.font.family, undefined); assert.equal(cell.authoredStyle.font.size, undefined)
  assert.equal(cell.richText[0].font.family, undefined); assert.equal(cell.richText[0].font.size, undefined)
  assert.equal(cell.richText[0].font.underline, true); assert.equal(cell.richText[1].vendor.keep, 1)
})

test('font-weight null clears direct false and restores inherited bold', () => {
  const model = sample(), sheet = model.sheets[0], cell = sheet.cells['0:0']
  cell.authoredStyle = { font: { bold: false } }; refreshWorkbookCellStyles(model, undefined, false)
  const action = fontFormatActionForCommand('sheet.command.set-font-weight', null)
  assert.deepEqual(action, { kind: 'mark', field: 'bold', value: null })
  applySheetFontFormatCommand(sheet, [range()], action, 'local-user', [], undefined, model)
  assert.equal(cell.style.font.bold, true); assert.equal(cell.authoredStyle.font, undefined)
})

test('legacy snapshots retain full direct styles until a property is explicitly cleared', () => {
  const model = sample(), sheet = model.sheets[0], legacy = createCell(0, 0, '00123')
  delete legacy.authoredStyle; legacy.style.font.family = 'Aptos'; legacy.style.font.size = 11
  legacy.style.numberFormat = '@'; legacy.dataType = 'text'; sheet.cells['0:0'] = legacy
  initializeWorkbookCellStyles(model)
  assert.equal(legacy.style.font.family, 'Aptos'); assert.equal(legacy.style.font.size, 11)
  applySheetFontFormatCommand(sheet, [range()], { kind: 'family', value: null }, 'local-user', [], undefined, model)
  assert.equal(legacy.style.font.family, 'Georgia'); assert.equal(legacy.style.font.size, 11)
  assert.equal(legacy.value, '00123'); assert.equal(legacy.style.numberFormat, '@')
})

test('explicit default-valued formatting under an inheritance layer survives later default changes', () => {
  const model = sample(), sheet = model.sheets[0]
  applySheetFontFormatCommand(sheet, [range(2, 2, 1)], { kind: 'size', value: 14 }, 'local-user', [], undefined, model)
  sheet.defaultStyle.font.size = 20; refreshWorkbookCellStyles(model, undefined, false)
  assert.equal(sheet.cells['2:1'].style.font.size, 14)
  assert.equal(resolveSheetCellStyle(sheet, 3, 1, undefined, model).font.size, 20)
})

test('bare default and null no-ops retain an empty sparse sheet', () => {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  for (const action of [{ kind: 'family', value: 'Aptos' }, { kind: 'size', value: 11 }, { kind: 'family', value: null }, { kind: 'size', value: null }]) {
    assert.equal(applySheetFontFormatCommand(sheet, [range()], action), false)
  }
  assert.deepEqual(sheet.cells, {})
})

test('model and detached view resets agree and skip filtered rows while retaining manual hidden rows', () => {
  const model = sample(), sheet = model.sheets[0]
  for (const row of [1, 2]) {
    sheet.cells[`${row}:0`] = createCell(row, 0, 'abc')
    sheet.cells[`${row}:0`].authoredStyle = { font: { size: 30 } }
  }
  sheet.hiddenRows = [2]; initializeWorkbookCellStyles(model)
  const view = toExcelWorkbookView(model).sheets[0]; view.filteredRows = [1]
  applySheetFontFormatCommand(sheet, [range(0, 2)], { kind: 'size', value: null }, 'local-user', [1], undefined, model)
  applySheetViewFontFormatCommand(view, [range(0, 2)], { kind: 'size', value: null })
  assert.equal(view.cells.A2.fontSize, 30); assert.equal(view.cells.A3.fontSize, 16)
  assert.equal(sheet.cells['1:0'].style.font.size, 30); assert.equal(sheet.cells['2:0'].style.font.size, 16)
})

test('clear formatting restores inherited layout and number format without changing raw values', () => {
  const model = sample(), sheet = model.sheets[0], cell = sheet.cells['0:0']
  cell.authoredStyle = { font: { family: 'Aptos', italic: true }, horizontal: 'left', numberFormat: '@' }
  refreshWorkbookCellStyles(model, undefined, false)
  applySheetClearCommand(sheet, [range()], 'format', 'local-user', [], model)
  const restored = sheet.cells['0:0']
  assert.equal(restored.value, 'abc'); assert.deepEqual(restored.authoredStyle, {})
  assert.equal(restored.style.font.family, 'Georgia'); assert.equal(restored.style.horizontal, 'right')
})

test('new layout changes author only the requested fields and keep the inherited font', () => {
  const model = sample(), sheet = model.sheets[0]
  applySheetLayoutFormatCommand(sheet, [range(2, 2, 0)], { kind: 'vertical', value: 'top' }, 'local-user', [], model)
  assert.equal(sheet.cells['2:0'].style.font.family, 'Courier New')
  assert.equal(sheet.cells['2:0'].authoredStyle.font, undefined)
})

test('snapshot and view round trips do not turn inherited caches into direct formatting', () => {
  const model = sample(), before = cloneModel(model), reopened = cloneModel(model)
  initializeWorkbookCellStyles(reopened)
  assert.deepEqual(reopened.sheets[0].cells['0:0'].authoredStyle, {})
  const mapped = applyExcelWorkbookView(model, toExcelWorkbookView(model))
  assert.deepEqual(mapped.sheets[0].cells['0:0'].authoredStyle, {})
  mapped.sheets[0].rowStyles[0] = { font: { family: 'Inter' } }; refreshWorkbookCellStyles(mapped, undefined, false)
  assert.equal(mapped.sheets[0].cells['0:0'].style.font.family, 'Inter'); assert.deepEqual(model, before)
})

test('inserting half a million inherited style rows stores one span', () => {
  const model = sample(), sheet = model.sheets[0]
  transformSheetDefaultStyles(sheet, { axis: 'row', kind: 'insert', start: 1, count: 500_000 }, 0)
  assert.equal(sheet.rowStyleSpans.length, 1); assert.equal(Object.keys(sheet.rowStyles).length, 1)
  assert.equal(resolveSheetCellStyle(sheet, 300_000, 1, undefined, model).font.family, 'Georgia')
})

test('row styles follow move/copy/delete coordinates independently from workbook defaults', () => {
  const model = sample(), sheet = model.sheets[0]
  transformSheetDefaultStyles(sheet, { axis: 'row', kind: 'copy', start: 0, count: 1, destination: 4 })
  assert.equal(sheet.rowStyles[4], 'row')
  transformSheetDefaultStyles(sheet, { axis: 'row', kind: 'delete', start: 0, count: 1 })
  assert.equal(sheet.rowStyles[3], 'row'); assert.equal(model.defaultStyle.font.family, 'Arial')
})

test('fixed short style DTOs normalize decorations and layout while retaining source extensions', () => {
  const style = normalizeSheetStyleInput({ ff: 'Inter', fs: 13, bl: 0, ul: { s: 1, t: 7 }, st: { s: 0 }, ht: 2, vt: 1, tb: 3, stf: 1,
    bg: { rgb: '#123456' }, n: { pattern: '0.00%' }, eastAsiaFontFamily: 'SimSun' })
  assert.equal(style.font.family, 'Inter'); assert.equal(style.font.underline, true); assert.equal(style.font.strike, false)
  assert.equal(style.horizontal, 'center'); assert.equal(style.vertical, 'top'); assert.equal(style.wrapStrategy, 'wrap')
  assert.equal(style.metadata.sourceStyle.ul.t, 7); assert.equal(style.metadata.sourceStyle.eastAsiaFontFamily, 'SimSun')
  for (const input of [{ ht: 9 }, { vt: '1' }, { fs: 0 }, { font: { bold: 'yes' } }]) assert.throws(() => normalizeSheetStyleInput(input))
})

test('fixed default border DTOs map every physical and diagonal key to drawable local edges', () => {
  const edge = { s: 7, cl: { rgb: '#112233' }, vendor: 'retained' }
  const style = normalizeSheetStyleInput({ bd: { t: edge, r: null, b: edge, l: edge, tl_br: edge, tl_bc: edge,
    tl_mr: edge, bl_tr: edge, ml_tr: edge, bc_tr: edge } })
  assert.equal(style.border.top.style, 'double'); assert.equal(style.border.top.color, '#112233'); assert.equal(style.border.right, null)
  for (const key of ['bottom', 'left', 'diagonalDown', 'topLeftBottomCenter', 'topLeftMiddleRight', 'diagonalUp', 'middleLeftTopRight', 'bottomCenterTopRight']) {
    assert.equal(style.border[key].style, 'double'); assert.equal(style.border[key].vendor, 'retained')
  }
  assert.deepEqual(style.metadata.sourceStyle.bd.t, edge)
  assert.throws(() => normalizeSheetStyleInput({ bd: { t: { s: 99 } } }))
})

test('selection probes include inherited blanks, direct cells and axis intersections without allocating cells', () => {
  const model = sample(), view = toExcelWorkbookView(model).sheets[0], before = cloneModel(view.cells)
  const probes = sheetStyleSelectionProbes(view, [range(0, 1_048_575, 0, 16_383)])
  assert.equal(probes.complete, true); assert.ok(probes.cells.length < 100)
  assert.ok(probes.cells.some(cell => cell.fontFamily === 'Georgia'))
  assert.ok(probes.cells.some(cell => cell.fontFamily === 'Arial')); assert.deepEqual(view.cells, before)
  assert.equal(sheetStyleSelectionProbes(view, [range(0, 10, 0, 10)], 1).complete, false)
})

test('blank-cell input interprets the inherited number format and retains sparse direct formatting', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.defaultStyle.numberFormat = '0.0%'
  const plan = planSheetCellInput(model, { unitId: model.id, sheetId: sheet.id, row: 3, column: 1, input: '25%' },
    { region: 'en-US', currentYear: 2026, userId: 'local-user' })
  assert.equal(plan.result.accepted, true); assert.equal(plan.cell.value, 0.25)
  assert.equal(plan.cell.style.font.family, 'Arial'); assert.equal(plan.cell.authoredStyle.font, undefined)
})

test('a million blank source rows retain one resolved style span and a null point exception', () => {
  const model = sample(), sheet = model.sheets[0], before = cloneModel(sheet.cells)
  sheet.rowStyleSpans = [{ start: 100, end: 900_000, value: 'row' }]; sheet.rowStyles[150] = null
  const source = captureSheetStyleRegion(sheet, { start: 0, end: 1_048_575 }, { start: 0, end: 0 }, model)
  assert.equal(source.rowStyleSpans.length, 1); assert.equal(source.rowStyles[150], null)
  assert.equal(sheetStyleRegionCellStyle(source, 200, 0).font.family, 'Georgia')
  assert.equal(sheetStyleRegionCellStyle(source, 150, 0).font.family, 'Courier New')
  model.styles.row.font.family = 'Inter'
  assert.equal(sheetStyleRegionCellStyle(source, 200, 0).font.family, 'Georgia'); assert.deepEqual(sheet.cells, before)
})

test('region snapshots reject live style IDs, out-of-bounds points and overlapping axis spans', () => {
  const model = sample(), sheet = model.sheets[0], source = captureSheetStyleRegion(sheet, [0, 2], [0], model)
  assert.doesNotThrow(() => assertSheetStyleRegion(source, 2, 1))
  for (const patch of [{ rowStyles: { 0: 'row' } }, { rowStyles: { 2: {} } },
    { rowStyleSpans: [{ start: 0, end: 1, value: {} }, { start: 1, end: 1, value: {} }] }]) {
    assert.throws(() => assertSheetStyleRegion({ ...source, ...patch }, 2, 1))
  }
})

test('copying filtered sparse blanks remaps inherited rows and writes their effective HTML styles', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.rowStyles[2] = { font: { size: 23, italic: true } }; sheet.autoFilter = { range: range(0, 2), columns: {}, filteredRows: [1] }
  const before = cloneModel(sheet.cells), snapshot = captureSheetClipboard(model, { unitId: model.id, sheetId: sheet.id, range: range(0, 2) })
  assert.deepEqual(snapshot.rows, [0, 2]); assert.equal(Object.keys(snapshot.cells).length, 1)
  const blank = sheetClipboardCellAt(snapshot, 1, 0)
  assert.equal(blank.style.font.size, 23); assert.equal(blank.style.font.family, 'Courier New'); assert.equal(blank.value, null)
  assert.match(sheetClipboardContent(snapshot).html, /font-size:23pt/); assert.deepEqual(sheet.cells, before)
  assert.equal(Object.keys(snapshot.cells).length, 1)
})

test('blank source copy applies captured formatting under different destination defaults', () => {
  const model = sample(), sheet = model.sheets[0], destination = createBlankSheet('Target')
  destination.defaultStyle = { font: { family: 'Inter', size: 40 }, border: { top: { style: 'double', color: '#000000' } } }
  model.sheets.push(destination)
  const snapshot = captureSheetClipboard(model, { unitId: model.id, sheetId: sheet.id, range: range(2, 3) })
  sheet.columnStyles[0].font.family = 'Changed after copy'
  const plan = createSheetClipboardPastePlan(model, { unitId: model.id, sheetId: destination.id, range: range(5, 5, 1) }, snapshot, 'all', { userId: 'local-user' })
  assert.equal(plan.result.accepted, true)
  const changed = plan.sheetChanges.find(item => item.id === destination.id)
  assert.equal(changed.cells['5:1'].style.font.family, 'Courier New'); assert.equal(changed.cells['6:1'].style.font.size, 16)
  assert.equal(changed.cells['5:1'].authoredStyle.font.family, 'Courier New'); assert.deepEqual(changed.cells['5:1'].style.border, {})
  assert.equal(Object.keys(sheet.cells).length, 1)
})

test('copying bare blank cells between bare sheets keeps an empty sparse map', () => {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'bare-blank-copy'
  const snapshot = captureSheetClipboard(model, { unitId: model.id, sheetId: sheet.id, range: range(2) })
  const plan = createSheetClipboardPastePlan(model, { unitId: model.id, sheetId: sheet.id, range: range(5) }, snapshot, 'all', { userId: 'local-user' })
  assert.equal(plan.result.accepted, true); assert.equal(plan.result.changed, false); assert.deepEqual(sheet.cells, {})
})

test('values-only blank paste clears content while preserving inherited destination formatting', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.cells['5:0'] = createCell(5, 0, 'target'); sheet.cells['5:0'].authoredStyle = { font: { italic: true } }; initializeWorkbookCellStyles(model)
  const snapshot = captureSheetClipboard(model, { unitId: model.id, sheetId: sheet.id, range: range(2) })
  const plan = createSheetClipboardPastePlan(model, { unitId: model.id, sheetId: sheet.id, range: range(5) }, snapshot, 'values', { userId: 'local-user' })
  const cell = plan.sheetChanges[0].cells['5:0']
  assert.equal(cell.value, null); assert.equal(cell.style.font.family, 'Courier New'); assert.equal(cell.style.font.italic, true)
  assert.deepEqual(cell.authoredStyle, { font: { italic: true } })
})

test('pending cuts become stale when an inherited blank-source style changes', () => {
  const model = sample(), sheet = model.sheets[0], sourceRange = range(2)
  const snapshot = captureSheetClipboard(model, { unitId: model.id, sheetId: sheet.id, range: sourceRange }, { copyType: 'cut' })
  sheet.defaultStyle.font.size = 21
  assert.notEqual(sheetClipboardSourceFingerprint(sheet, sourceRange, model), snapshot.source.fingerprint)
  const plan = createSheetClipboardPastePlan(model, { unitId: model.id, sheetId: sheet.id, range: range(5) }, snapshot, 'all', { userId: 'local-user' })
  assert.equal(plan.result.reason, 'stale-source'); assert.deepEqual(plan.sheetChanges, [])
})

test('format painter tiles inherited blank styles consistently through model and detached view paths', () => {
  const model = sample(), sheet = model.sheets[0], destination = createBlankSheet('Painter target')
  destination.defaultStyle = { font: { family: 'Inter', size: 40 } }; destination.cells['4:1'] = createCell(4, 1, 'keep')
  model.sheets.push(destination); initializeWorkbookCellStyles(model)
  const source = captureSheetFormatPainterSource(sheet, range(0, 1), 'local-user', model)
  const viewSource = captureSheetViewFormatPainterSource(toExcelWorkbookView(model).sheets[0], range(0, 1))
  const view = toExcelWorkbookView(model).sheets[1]
  assert.deepEqual(source, viewSource); assert.deepEqual(source.styles, {})
  sheet.rowStyles[0] = { font: { family: 'Changed after capture' } }
  applySheetFormatPainterCommand(destination, source, [range(4, 7, 1)], { styleContext: model })
  applySheetViewFormatPainterCommand(view, viewSource, [range(4, 7, 1)])
  assert.equal(destination.cells['4:1'].value, 'keep'); assert.equal(destination.cells['4:1'].style.font.family, 'Georgia')
  assert.equal(destination.cells['5:1'].style.font.family, 'Courier New'); assert.equal(destination.cells['6:1'].style.font.family, 'Georgia')
  assert.equal(view.cells.B5.fontFamily, 'Georgia'); assert.equal(view.cells.B6.fontFamily, 'Courier New')
  assert.equal(view.cells.B5.authoredStyle.font.family, 'Georgia'); assert.equal(destination.cells['5:1'].authoredStyle.font.size, 16)
})

const borderAction = type => ({ apply: true, state: { active: true, type, style: 'thin', color: '#008800' } })
test('border edits read blank inherited edges and author only the changed border group', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.defaultStyle.border = { bottom: { style: 'double', color: '#333333' } }
  applySheetBorderFormatCommand(sheet, [range(3, 3, 3)], borderAction('top'), 'local-user', model)
  const cell = sheet.cells['3:3']
  assert.equal(cell.style.font.family, 'Arial'); assert.equal(cell.authoredStyle.font, undefined)
  assert.equal(cell.style.border.bottom.style, 'double'); assert.equal(cell.style.border.top.color, '#008800')
  sheet.defaultStyle.font.size = 30; refreshWorkbookCellStyles(model, undefined, false)
  assert.equal(cell.style.font.size, 30)
})

test('no-border clears inherited blank edges and adjacent segments through model and view paths', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.defaultStyle.border = { top: { style: 'thin', color: '#333333' }, bottom: { style: 'thin', color: '#333333' } }
  const view = toExcelWorkbookView(model).sheets[0]
  applySheetBorderFormatCommand(sheet, [range(3, 3, 3)], borderAction('none'), 'local-user', model)
  applySheetViewBorderFormatCommand(view, [range(3, 3, 3)], borderAction('none'))
  assert.deepEqual(sheet.cells['3:3'].style.border, {}); assert.deepEqual(view.cells.D4.border, {})
  assert.equal(sheet.cells['2:3'].style.border.bottom, undefined); assert.equal(sheet.cells['2:3'].style.border.top.style, 'thin')
  assert.equal(view.cells.D3.border.bottom, undefined); assert.equal(sheet.cells['3:3'].authoredStyle.font, undefined)
})

test('protected inherited neighboring edges reject a clear before mutating sparse cells', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.defaultStyle.border = { top: { style: 'thin', color: '#333333' }, bottom: { style: 'thin', color: '#333333' } }
  sheet.rangeProtectionRules = [{ id: 'neighbor', ranges: [range(2, 2, 3)], allowedUserIds: ['owner'] }]
  const before = cloneModel(sheet)
  assert.throws(() => applySheetBorderFormatCommand(sheet, [range(3, 3, 3)], borderAction('none'), 'local-user', model), /保护/)
  assert.deepEqual(sheet, before)
})

const publicStyleInput = () => ({ id: 'public-style-input', calculationMode: 'manual', defaultStyle: { ff: 'Arial', fs: 12 },
  styles: { row: { ff: 'Georgia', bl: 1 }, column: { ff: 'Courier New', fs: 16 }, direct: { fs: 28, it: 1, bg: { rgb: '#ffeedd' } } },
  sheetOrder: ['sheet1'], sheets: { sheet1: { id: 'sheet1', name: 'Sheet1', rowCount: 100, columnCount: 20,
    defaultStyle: { fs: 14 }, rowData: { 0: { s: 'row', h: 40, vendor: { keep: true } } },
    columnData: { 0: { s: 'column', w: 120 } }, cellData: { 0: { 0: { v: 'abc', s: 'direct' } } } } } })

test('public rowData, columnData and cell.s style references compose without discarding axis source fields', () => {
  const model = workbookFromInput(publicStyleInput()), sheet = model.sheets[0], cell = sheet.cells['0:0']
  assert.equal(sheet.rowStyles[0], 'row'); assert.equal(sheet.columnStyles[0], 'column'); assert.equal(cell.authoredStyleReference, 'direct')
  assert.deepEqual(cell.authoredStyle, {}); assert.equal(cell.style.font.family, 'Georgia'); assert.equal(cell.style.font.size, 28)
  assert.equal(cell.style.font.bold, true); assert.equal(cell.style.font.italic, true)
  assert.deepEqual(sheet.metadata.publicAxisData.rowData[0].vendor, { keep: true })
  assert.equal(resolveSheetCellStyle(sheet, 1, 0, undefined, model).font.family, 'Courier New')
})

test('catalog updates remain live for unrelated referenced properties while direct overrides stay explicit', () => {
  const model = workbookFromInput(publicStyleInput()), sheet = model.sheets[0], cell = sheet.cells['0:0']
  setSheetCellStyleProperties(cell, { font: { size: 22, italic: false } })
  const before = cloneModel(model)
  model.styles.direct.font.size = 31; model.styles.direct.font.italic = true; model.styles.direct.fill = '#ddeeff'
  refreshWorkbookCellStyles(model, before)
  assert.equal(cell.style.font.size, 22); assert.equal(cell.style.font.italic, false); assert.equal(cell.style.fill, '#ddeeff')
  assert.equal(cell.authoredStyleReference, 'direct'); assert.equal(cell.authoredStyle.fill, undefined)
})

test('font resets suppress referenced direct properties without changing the shared catalog or sibling cells', () => {
  const model = workbookFromInput(publicStyleInput()), sheet = model.sheets[0], cell = sheet.cells['0:0']
  model.styles.direct.font.family = 'Times New Roman'; model.styles.direct.font.bold = false
  sheet.cells['1:0'] = applyCellValue(undefined, 1, 0, { v: 'sibling', s: 'direct' }, model, sheet)
  refreshWorkbookCellStyles(model, undefined, false)
  const source = cloneModel(model.styles.direct)
  applySheetFontFormatCommand(sheet, [range()], { kind: 'family', value: null }, 'local-user', [], undefined, model)
  applySheetFontFormatCommand(sheet, [range()], { kind: 'mark', field: 'bold', value: null }, 'local-user', [], undefined, model)
  assert.equal(cell.style.font.family, 'Georgia'); assert.equal(cell.style.font.bold, true)
  assert.deepEqual(cell.clearedStyleProperties, ['font.family', 'font.bold']); assert.equal(cell.authoredStyleReference, 'direct')
  assert.equal(sheet.cells['1:0'].style.font.family, 'Times New Roman'); assert.deepEqual(model.styles.direct, source)
})

test('referenced styles and reset masks survive view round trips and later catalog changes', () => {
  const model = workbookFromInput(publicStyleInput()), sheet = model.sheets[0]
  applySheetFontFormatCommand(sheet, [range()], { kind: 'size', value: null }, 'local-user', [], undefined, model)
  const view = toExcelWorkbookView(model)
  assert.equal(view.sheets[0].cells.A1.authoredStyleReference, 'direct')
  const mapped = applyExcelWorkbookView(model, view)
  assert.deepEqual(mapped.sheets[0].cells['0:0'].clearedStyleProperties, ['font.size'])
  mapped.styles.direct.font.size = 41; refreshWorkbookCellStyles(mapped, undefined, false)
  assert.equal(mapped.sheets[0].cells['0:0'].style.font.size, 16)
})

test('cell.s null removes the referenced style and explicit overrides while preserving the new value', () => {
  const model = workbookFromInput(publicStyleInput()), sheet = model.sheets[0]
  const cell = applyCellValue(sheet.cells['0:0'], 0, 0, { v: 'replacement', s: null }, model, sheet)
  assert.equal(cell.authoredStyleReference, undefined); assert.deepEqual(cell.authoredStyle, {})
  assert.equal(cell.value, 'replacement'); assert.equal(cell.style.font.family, 'Georgia'); assert.equal(cell.style.font.size, 16)
  const before = cloneModel(sheet)
  assert.throws(() => applyCellValue(sheet.cells['0:0'], 0, 0, { v: 'bad', s: 'missing' }, model, sheet))
  assert.deepEqual(sheet, before)
})

test('clear formatting removes references and modeled rich-run fonts while retaining source text and extensions', () => {
  const model = workbookFromInput(publicStyleInput()), sheet = model.sheets[0], cell = sheet.cells['0:0']
  cell.richText = [{ id: 'a', text: 'abc', fontFamily: 'Times New Roman', font: { size: 32, bold: false, vendorSlot: 'keep' },
    hyperlink: 'https://example.test', vendor: { keep: true } }]
  const view = toExcelWorkbookView(model).sheets[0]
  applySheetClearCommand(sheet, [range()], 'format', 'local-user', [], model)
  applySheetViewClearCommand(view, [range()], 'format')
  const cleared = sheet.cells['0:0'], run = cleared.richText[0]
  assert.equal(cleared.authoredStyleReference, undefined); assert.equal(cleared.style.font.family, 'Georgia')
  assert.equal(run.fontFamily, undefined); assert.deepEqual(run.font, { vendorSlot: 'keep' })
  assert.equal(run.id, 'a'); assert.equal(run.text, 'abc'); assert.equal(run.hyperlink, 'https://example.test')
  assert.deepEqual(view.cells.A1.richText, cleared.richText); assert.equal(view.cells.A1.authoredStyleReference, null)
})

test('public axis style indices and malformed referenced reset masks reject before importing a workbook', () => {
  const invalid = publicStyleInput(); invalid.sheets.sheet1.rowData[1_048_576] = { s: 'row' }
  assert.throws(() => workbookFromInput(invalid))
  const model = workbookFromInput(publicStyleInput()); model.sheets[0].cells['0:0'].clearedStyleProperties = ['font.unknown']
  assert.throws(() => initializeWorkbookCellStyles(model))
})

test('reopening a saved style cache uses changed defaults and catalog values without authoring its old presentation', () => {
  const saved = cloneModel(sample()), sheet = saved.sheets[0]
  saved.styles.row.font.family = 'Inter'; saved.defaultStyle.font.size = 19
  sheet.columnStyles[0].font.size = 21
  initializeWorkbookCellStyles(saved)
  assert.equal(sheet.cells['0:0'].style.font.family, 'Inter'); assert.equal(sheet.cells['0:0'].style.font.size, 21)
  assert.deepEqual(sheet.cells['0:0'].authoredStyle, {}); assert.equal(sheet.cells['0:0'].styleMaterialized, true)
})

test('a saved referenced reset remains sparse when reopening after its catalog and axis defaults change', () => {
  const model = workbookFromInput(publicStyleInput()), sheet = model.sheets[0]
  applySheetFontFormatCommand(sheet, [range()], { kind: 'size', value: null }, 'local-user', [], undefined, model)
  const saved = cloneModel(model)
  saved.styles.direct.font.size = 60; saved.styles.direct.fill = '#123456'; saved.styles.column.font.size = 20
  initializeWorkbookCellStyles(saved)
  const cell = saved.sheets[0].cells['0:0']
  assert.equal(cell.style.font.size, 20); assert.equal(cell.style.fill, '#123456')
  assert.deepEqual(cell.authoredStyle, {}); assert.deepEqual(cell.clearedStyleProperties, ['font.size'])
})

test('fresh imperative cell styling is adopted once while malformed saved cache markers reject', () => {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.defaultStyle = { font: { family: 'Arial', size: 14 } }
  const cell = createCell(0, 0, 'fresh'); cell.style.font.family = 'Georgia'; cell.style.font.bold = true
  sheet.cells[cell.id] = cell; initializeWorkbookCellStyles(model)
  assert.deepEqual(cell.authoredStyle.font, { family: 'Georgia', bold: true }); assert.equal(cell.style.font.size, 14)
  const invalid = cloneModel(model); invalid.sheets[0].cells['0:0'].styleMaterialized = false
  assert.throws(() => initializeWorkbookCellStyles(invalid), /缓存标记/)
})

test('copy fill samples inherited blank source formatting and explicitly replaces destination defaults', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.rowStyles[3] = { font: { family: 'Inter', size: 30 }, fill: '#111111' }
  const plan = createSheetAutoFillPlan(model, sheet.id, range(1, 1, 0), range(1, 3, 0), { applyType: 'COPY' })
  assert.equal(plan.result.accepted, true)
  const candidate = { ...model, sheets: [plan.sheet] }; refreshWorkbookCellStyles(candidate, undefined, false)
  for (const row of [2, 3]) {
    const cell = plan.sheet.cells[`${row}:0`]
    assert.equal(cell.value, null); assert.equal(cell.style.font.family, 'Courier New'); assert.equal(cell.style.font.size, 16)
    assert.equal(cell.authoredStyle.font.family, 'Courier New'); assert.equal(cell.authoredStyleReference, undefined)
  }
  assert.equal(sheet.cells['3:0'], undefined)
})

test('format-only fill retains destination content and clears old direct reference and reset masks', () => {
  const model = workbookFromInput(publicStyleInput()), sheet = model.sheets[0]
  sheet.cells['2:0'] = applyCellValue(undefined, 2, 0, { v: 'destination', s: 'direct' }, model, sheet)
  sheet.cells['2:0'].clearedStyleProperties = ['font.size']; sheet.cells['2:0'].themeStyle = { font: { size: 44 } }
  refreshWorkbookCellStyles(model, undefined, false)
  const plan = createSheetAutoFillPlan(model, sheet.id, range(1, 1, 0), range(1, 2, 0), { applyType: 'ONLY_FORMAT' })
  const cell = plan.sheet.cells['2:0']
  assert.equal(cell.value, 'destination'); assert.equal(cell.style.font.size, 16)
  assert.equal(cell.authoredStyleReference, undefined); assert.equal(cell.clearedStyleProperties, undefined); assert.equal(cell.themeStyle, undefined)
})

test('no-format series fill preserves destination sparse overrides, references, resets and theme sources after refresh', () => {
  const model = workbookFromInput(publicStyleInput()), sheet = model.sheets[0]
  sheet.cells['0:0'] = applyCellValue(undefined, 0, 0, { v: 1, s: { fs: 36 } }, model, sheet)
  sheet.cells['1:0'] = applyCellValue(undefined, 1, 0, { v: 2 }, model, sheet)
  const target = applyCellValue(undefined, 2, 0, { v: 100, s: 'direct' }, model, sheet)
  target.clearedStyleProperties = ['font.size']; target.themeStyle = { font: { size: 18 } }
  setSheetCellStyleProperties(target, { font: { bold: false }, numberFormat: '0.00' }); sheet.cells['2:0'] = target
  refreshWorkbookCellStyles(model, undefined, false)
  const expected = cloneModel(target), plan = createSheetAutoFillPlan(model, sheet.id, range(0, 1, 0), range(0, 3, 0), { applyType: 'NO_FORMAT' })
  const candidate = { ...model, sheets: [plan.sheet] }; refreshWorkbookCellStyles(candidate, model, false)
  const cell = plan.sheet.cells['2:0']
  assert.equal(cell.value, 3); assert.equal(cell.display, '3.00'); assert.deepEqual(cell.style, expected.style)
  for (const key of ['authoredStyle', 'authoredStyleReference', 'clearedStyleProperties', 'themeStyle']) assert.deepEqual(cell[key], expected[key])
  candidate.styles.direct.fill = '#445566'; refreshWorkbookCellStyles(candidate, undefined, false)
  assert.equal(cell.style.fill, '#445566'); assert.equal(cell.style.font.size, 18); assert.equal(cell.style.font.bold, false)
})

test('blank no-format fill clears existing contents and keeps previously absent destination cells sparse', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.cells['2:0'] = createCell(2, 0, 'clear me'); sheet.rowStyles[2] = { font: { family: 'Inter', size: 24 } }
  refreshWorkbookCellStyles(model, undefined, false)
  const plan = createSheetAutoFillPlan(model, sheet.id, range(1, 1, 0), range(1, 3, 0), { applyType: 'NO_FORMAT' })
  assert.equal(plan.sheet.cells['2:0'].value, null); assert.equal(plan.sheet.cells['2:0'].style.font.family, 'Inter')
  assert.deepEqual(plan.sheet.cells['2:0'].authoredStyle, {}); assert.equal(plan.sheet.cells['3:0'], undefined)
  const bare = createBlankWorkbook(), empty = bare.sheets[0]
  const copied = createSheetAutoFillPlan(bare, empty.id, range(), range(0, 4), { applyType: 'COPY' })
  assert.deepEqual(copied.sheet.cells, {})
})

test('filtered fill samples each physical source row and retains untouched filtered destination formatting', () => {
  const model = sample(), sheet = model.sheets[0]
  sheet.autoFilter = { range: range(0, 6, 0), columns: [], filteredRows: [2, 5] }
  sheet.rowStyles[1] = { font: { family: 'Inter' } }; sheet.rowStyles[3] = { font: { family: 'Georgia' } }
  sheet.cells['5:0'] = createCell(5, 0, 'filtered target')
  const before = cloneModel(sheet.cells['5:0'])
  const plan = createSheetAutoFillPlan(model, sheet.id, range(1, 3, 0), range(1, 6, 0), { applyType: 'COPY' })
  assert.equal(plan.sheet.cells['4:0'].style.font.family, 'Inter'); assert.equal(plan.sheet.cells['6:0'].style.font.family, 'Georgia')
  assert.deepEqual(plan.sheet.cells['5:0'], before)
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const withStudio = run => {
  const studio = new OfficeStudio({ workbook: sample(), currentUserId: 'local-user' })
  try { return run(studio, FOffice.newAPI(studio)) } finally { studio.dispose() }
}

test('Facade defaults share commands, raw/resolved references, undo/redo and snapshot reopen', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet(), commands = []
  studio.events.addEvent('CommandExecuted', event => commands.push(event.id))
  sheet.setRowDefaultStyle(0, { ff: 'Inter' })
  assert.equal(sheet.getRange('A1').getFontFamily(), 'Inter')
  assert.equal(commands.at(-1), 'sheet.command.set-row-default-style')
  book.undo(); assert.equal(sheet.getRowDefaultStyle(0, true), 'row'); assert.equal(sheet.getRange('A1').getFontFamily(), 'Georgia')
  book.redo(); const saved = book.save()
  const reopened = api.createWorkbook({ ...saved, id: 'default-style-reopened' })
  assert.equal(reopened.getActiveSheet().getRange('A1').getFontFamily(), 'Inter')
  assert.equal(reopened.getActiveSheet().getRange('A1').getCellStyleData('cell').font, undefined)
}))

test('Facade row/column read precedence and explicit font reset retain default references', () => withStudio((_studio, api) => {
  const sheet = api.getActiveWorkbook().getActiveSheet(), cell = sheet.getRange('A1')
  assert.equal(sheet.getRowDefaultStyle(0, true), 'row'); assert.equal(sheet.getRowDefaultStyle(0).font.family, 'Georgia')
  assert.equal(cell.getFontFamily('col'), 'Courier New'); assert.equal(cell.getFontFamily('row'), 'Georgia')
  cell.setFontFamily('Aptos'); cell.setFontFamily(null)
  assert.equal(cell.getFontFamily(), 'Georgia'); assert.equal(cell.getFontFamily('cell'), null)
}))

test('invalid defaults and missing style IDs reject before adding any content history', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet(), before = cloneModel(studio.getWorkbookModel(book.id))
  for (const run of [() => sheet.setDefaultStyle('missing'), () => sheet.setRowDefaultStyle(-1, {}), () => sheet.setDefaultStyle({ fs: NaN })]) assert.throws(run)
  assert.deepEqual(studio.getWorkbookModel(book.id), before)
  assert.equal(studio.editor.history.units[book.id]?.undo.length ?? 0, 0)
}))

test('canceling a default command leaves authored formatting, caches and savepoint unchanged', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), before = cloneModel(studio.getWorkbookModel(book.id))
  studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'sheet.command.set-worksheet-default-style') event.cancel = true })
  book.getActiveSheet().setDefaultStyle({ ff: 'Inter' })
  assert.deepEqual(studio.getWorkbookModel(book.id), before)
  assert.equal(studio.editor.history.units[book.id]?.undo.length ?? 0, 0)
}))

test('range protection rejects a row default that would change protected cells outside the active selection', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet()
  studio.editor.workbook.model.sheets[0].rangeProtectionRules = [{ id: 'protected', ranges: [range(0, 0, 20)], allowedUserIds: ['owner'] }]
  const before = cloneModel(studio.getWorkbookModel(book.id))
  assert.throws(() => sheet.setRowDefaultStyle(0, { fs: 20 }), /保护/)
  assert.deepEqual(studio.getWorkbookModel(book.id), before)
}))

test('Facade public style IDs, selective resets and later value edits share history and reopen provenance', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet(), range = sheet.getRange('A1')
  range.setValue({ v: 'referenced', s: 'row' })
  assert.equal(book.save().sheets[0].cells['0:0'].authoredStyleReference, 'row')
  range.setFontFamily(null); range.setValue('latest')
  const saved = book.save(), cell = saved.sheets[0].cells['0:0']
  assert.equal(cell.value, 'latest'); assert.deepEqual(cell.clearedStyleProperties, ['font.family'])
  assert.equal(range.getCellStyleData('cell').font.family, undefined)
  book.undo(); assert.equal(book.save().sheets[0].cells['0:0'].value, 'referenced')
  book.redo(); const reopened = api.createWorkbook({ ...book.save(), id: 'reference-reopened' }).getActiveSheet()
  assert.equal(reopened.getRange('A1').getFontFamily(), 'Georgia')
  assert.deepEqual(studio.getWorkbookModel('reference-reopened').sheets[0].cells['0:0'].clearedStyleProperties, ['font.family'])
}))

test('Facade copy/refill keeps one content history entry and reopens no-format destination inheritance', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet(), unitId = book.getId(), subUnitId = sheet.getSheetId()
  sheet.getRange('A1').setValue(1); sheet.getRange('A2').setValue(2)
  sheet.setRowDefaultStyle(2, { ff: 'Inter', fs: 24 })
  const depth = studio.editor.history.units[unitId].undo.length
  assert.equal(api.executeCommandSync('sheet.command.auto-fill', { unitId, subUnitId, sourceRange: range(0, 1, 0), targetRange: range(0, 3, 0), applyType: 'COPY' }), true)
  const sessionId = studio.sheetFill.getView(unitId).id
  assert.equal(api.executeCommandSync('sheet.command.refill', { unitId, subUnitId, sessionId, applyType: 'NO_FORMAT' }), true)
  assert.equal(studio.editor.history.units[unitId].undo.length, depth + 1)
  assert.equal(sheet.getRange('A3').getValue(), 3); assert.equal(sheet.getRange('A3').getFontFamily(), 'Inter')
  assert.deepEqual(book.save().sheets[0].cells['2:0'].authoredStyle, {})
  book.undo(); assert.equal(sheet.getRange('A3').getValue(), null)
  book.redo(); const reopened = api.createWorkbook({ ...book.save(), id: 'default-fill-reopened' }).getActiveSheet()
  assert.equal(reopened.getRange('A3').getValue(), 3); assert.equal(reopened.getRange('A3').getFontFamily(), 'Inter')
  reopened.setRowDefaultStyle(2, { ff: 'Georgia', fs: 20 })
  assert.equal(reopened.getRange('A3').getFontFamily(), 'Georgia'); assert.equal(reopened.getRange('A3').getFontSize(), 20)
}))
