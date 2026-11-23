import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createBlankSheet, createCell } from '../src/domain/model.ts'
import { captureSheetClipboard } from '../src/features/sheet-workbench/sheet-clipboard-data.ts'
import { parseSheetClipboardText } from '../src/features/sheet-workbench/sheet-clipboard-content.ts'
import { sheetClipboardCapabilities, sheetClipboardCommandCapability } from '../src/features/sheet-workbench/sheet-clipboard-capabilities.ts'

// Draft only. These assertions describe intended contracts and have not been executed.
const range = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })
const copy = value => JSON.parse(JSON.stringify(value))
function workbook() {
  const book = createBlankWorkbook(); book.id = 'capability-book'
  book.sheets = ['Source', 'Target'].map((name, index) => ({ ...createBlankSheet(name), id: 'capability-sheet-' + index, rowCount: 40, columnCount: 12 }))
  book.activeSheetId = book.sheets[1].id
  book.sheets[0].cells['0:0'] = createCell(0, 0, 'text')
  book.sheets[0].cells['0:1'] = { ...createCell(0, 1, 5), formula: '=A2', dataType: 'formula' }
  book.sheets[0].cells['1:0'] = createCell(1, 0, 5)
  return book
}
function clipboard(book, selection = range(0, 0, 1, 1), options = {}) {
  return captureSheetClipboard(book, { unitId: book.id, sheetId: book.sheets[0].id, range: selection }, options)
}
function input(book, snapshot, extra = {}) {
  return { workbook: book, sheetId: book.sheets.at(-1).id, selection: { ranges: [range(4, 4)] }, userId: 'local-user',
    clipboard: snapshot, systemClipboardAvailable: false, ...extra }
}
function protection(selected, extra = {}) {
  return { id: 'protected', permissionId: 'permission', ownerUserId: 'remote', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView',
    ranges: [selected], ...extra }
}

test('entry availability distinguishes absent content, an unknown callable reader, missing sheets and active editing', () => {
  const book = workbook(), absent = sheetClipboardCapabilities(input(book))
  assert.equal(absent.copy.disabled, false); assert.ok(Object.values(absent.paste).every(entry => entry.disabled))
  assert.equal(absent.source.kind, 'none'); assert.equal(absent.source.systemClipboardAvailable, false)
  const reader = sheetClipboardCapabilities(input(book, undefined, { systemClipboardAvailable: true }))
  assert.ok(Object.values(reader.paste).every(entry => !entry.disabled)); assert.equal(reader.source.kind, 'none')
  const missing = sheetClipboardCapabilities(input(book, undefined, { sheetId: 'removed', systemClipboardAvailable: true }))
  assert.equal(missing.copy.hidden, true); assert.equal(missing.ribbonPaste.hidden, true)
  const editing = sheetClipboardCapabilities(input(book, clipboard(book), { editing: true }))
  assert.equal(editing.copy.disabled, true); assert.equal(editing.cut.disabled, true); assert.equal(editing.paste.all.disabled, true)
  assert.equal(editing.copy.hidden, false)
})

test('blank internal ranges retain clear-value and default-format paste contracts without requiring a live source', () => {
  const book = workbook(), snapshot = clipboard(book, range(8, 0)); delete snapshot.source
  const result = sheetClipboardCapabilities(input(book, snapshot))
  assert.equal(result.source.hasValue, false); assert.equal(result.source.hasFormula, false)
  assert.equal(result.source.sourcePresent, false); assert.equal(result.source.hasFormat, true)
  assert.equal(result.source.hasColumnWidths, true)
  assert.ok(Object.values(result.paste).every(entry => !entry.disabled))
})

test('text content cannot supply formats or widths while formula-only internal defaults use value permissions', () => {
  const book = workbook(), text = parseSheetClipboardText('external\t0')
  const plain = sheetClipboardCapabilities(input(book, text))
  assert.equal(plain.source.kind, 'text'); assert.equal(plain.source.hasFormat, false)
  assert.equal(plain.paste.formats.disabled, true); assert.equal(plain.paste.columnWidths.disabled, true)
  assert.equal(plain.paste.all.disabled, false); assert.equal(plain.paste.formulas.disabled, false)
  const formulas = sheetClipboardCapabilities(input(book, clipboard(book, range(0, 0, 0, 1), { formulaOnly: true }), {
    permission: request => request.point !== 'setCellStyle',
  }))
  assert.equal(formulas.source.formulaOnly, true); assert.equal(formulas.source.hasFormula, true)
  assert.equal(formulas.paste.all.effectiveMode, 'formulas'); assert.equal(formulas.paste.all.disabled, false)
  assert.equal(formulas.paste.formats.disabled, true)
})

test('worksheet copy, view and edit protection stay separate and optional operation points cannot override them', () => {
  const book = workbook(), snapshot = clipboard(book)
  book.sheets[1].worksheetProtection = { ...protection(range(0, 0)), unitType: 'worksheet' }
  const readOnly = sheetClipboardCapabilities(input(book, snapshot, { permission: () => true }))
  assert.equal(readOnly.copy.disabled, false); assert.equal(readOnly.cut.disabled, true)
  assert.ok(Object.values(readOnly.paste).every(entry => entry.disabled))
  book.sheets[1].worksheetProtection.viewState = 'noOneElseCanView'
  assert.equal(sheetClipboardCapabilities(input(book, snapshot, { permission: () => true })).copy.disabled, true)
  delete book.sheets[1].worksheetProtection
  const noCopy = sheetClipboardCapabilities(input(book, snapshot, { permission: request => request.point !== 'copy' }))
  assert.equal(noCopy.copy.disabled, true); assert.equal(noCopy.cut.disabled, true); assert.equal(noCopy.paste.values.disabled, false)
})

test('value, style and whole-column operation points independently govern all six paste modes', () => {
  const book = workbook(), snapshot = clipboard(book)
  const noValue = sheetClipboardCapabilities(input(book, snapshot, { permission: request => request.point !== 'setCellValue' }))
  for (const mode of ['all', 'values', 'formulas', 'exceptBorders']) assert.equal(noValue.paste[mode].disabled, true)
  assert.equal(noValue.paste.formats.disabled, false); assert.equal(noValue.paste.columnWidths.disabled, false)
  const noStyle = sheetClipboardCapabilities(input(book, snapshot, { permission: request => request.point !== 'setCellStyle' }))
  for (const mode of ['all', 'formats', 'exceptBorders']) assert.equal(noStyle.paste[mode].disabled, true)
  assert.equal(noStyle.paste.values.disabled, false); assert.equal(noStyle.paste.formulas.disabled, false)
  assert.equal(noStyle.paste.columnWidths.disabled, false)
  const noWidth = sheetClipboardCapabilities(input(book, snapshot, { permission: request => request.point !== 'setColumnWidth' }))
  assert.equal(noWidth.paste.columnWidths.disabled, true); assert.equal(noWidth.paste.all.disabled, false)
})

test('known source dimensions check expanded writes beyond the selected anchor and skip filtered target rows', () => {
  const book = workbook(), snapshot = clipboard(book)
  book.sheets[1].rangeProtectionRules = [protection(range(5, 5))]
  assert.equal(sheetClipboardCapabilities(input(book, snapshot)).paste.values.disabled, true)
  assert.equal(sheetClipboardCapabilities(input(book, snapshot, { systemClipboardAvailable: true })).paste.values.disabled, false)
  book.sheets[1].filteredRows = [5]; book.sheets[1].rangeProtectionRules = [protection(range(6, 4))]
  const result = sheetClipboardCapabilities(input(book, snapshot, { selection: { ranges: [range(4, 4, 5, 4)] } }))
  assert.equal(result.paste.values.disabled, true)
  book.sheets[1].rangeProtectionRules = [protection(range(5, 4))]
  assert.equal(sheetClipboardCapabilities(input(book, snapshot, { selection: { ranges: [range(4, 4, 5, 4)] } })).paste.values.disabled, false)
})

test('multiple selections use the explicit active range and do not imply an exclusive plugin range', () => {
  const book = workbook(), snapshot = clipboard(book, range(0, 0))
  book.sheets[1].rangeProtectionRules = [protection(range(10, 8))]
  const selection = { ranges: [range(4, 4), range(10, 8)], activeRangeIndex: 0 }
  const result = sheetClipboardCapabilities(input(book, snapshot, { selection }))
  assert.equal(result.paste.values.disabled, false); assert.equal(result.pasteSpecial.hidden, false)
  assert.equal(result.copySpecial.hidden, false); assert.equal(result.ribbonPaste.hidden, false)
  assert.equal(sheetClipboardCapabilities(input(book, snapshot, { selection: { ranges: selection.ranges } })).paste.values.disabled, true)
  const exclusive = sheetClipboardCapabilities(input(book, snapshot, { selection, exclusiveRange: true }))
  assert.equal(exclusive.copySpecial.hidden, true); assert.equal(exclusive.pasteSpecial.hidden, true); assert.equal(exclusive.ribbonPaste.hidden, true)
  assert.equal(exclusive.copy.hidden, false); assert.equal(exclusive.cut.hidden, false); assert.equal(exclusive.paste.all.hidden, false)
})

test('column widths check the full target column even without a source origin or an explicit row count', () => {
  const book = workbook(), snapshot = clipboard(book, range(0, 0)); delete snapshot.source
  const calls = []
  book.sheets[1].rangeProtectionRules = [protection(range(10, 4))]
  const guarded = sheetClipboardCapabilities(input(book, snapshot))
  assert.equal(guarded.paste.values.disabled, false); assert.equal(guarded.paste.columnWidths.disabled, true)
  delete book.sheets[1].rowCount; delete book.sheets[1].rangeProtectionRules
  const result = sheetClipboardCapabilities(input(book, snapshot, { permission: request => {
    if (request.point === 'setColumnWidth') calls.push(request)
    return true
  } }))
  assert.equal(result.paste.columnWidths.disabled, false)
  assert.deepEqual(calls[0].ranges, [range(0, 4, 1_048_575, 4)]); assert.equal(calls[0].phase, 'target')
  for (const width of [NaN, -1, 2001]) {
    const invalid = { ...snapshot, columnWidths: [width] }
    const guarded = sheetClipboardCapabilities(input(book, invalid))
    assert.equal(guarded.source.hasColumnWidths, false); assert.equal(guarded.paste.columnWidths.disabled, true)
    assert.equal(guarded.paste.values.disabled, false)
  }
})

test('cut snapshots require a live editable same-workbook source for every mode and ordinary copies survive source closure', () => {
  const book = workbook(), moved = clipboard(book, range(0, 0), { copyType: 'cut' }), copied = clipboard(book, range(0, 0))
  book.sheets[0].rangeProtectionRules = [protection(range(0, 0))]
  assert.ok(Object.values(sheetClipboardCapabilities(input(book, moved)).paste).every(entry => entry.disabled))
  moved.source.unitId = 'other-workbook'
  assert.ok(Object.values(sheetClipboardCapabilities(input(book, moved)).paste).every(entry => entry.disabled))
  const reader = sheetClipboardCapabilities(input(book, moved, { systemClipboardAvailable: true }))
  assert.equal(reader.source.sameWorkbook, false); assert.equal(reader.paste.values.disabled, false)
  book.sheets.shift()
  const closed = sheetClipboardCapabilities({ ...input(book, copied), sheetId: book.sheets[0].id })
  assert.ok(Object.values(closed.paste).every(entry => !entry.disabled))
})

test('merge expansion protects the complete target and shape text disables cell commands without inventing a multi-selection rule', () => {
  const book = workbook(), snapshot = clipboard(book, range(0, 0))
  book.sheets[1].mergedRanges = [range(3, 3, 5, 5)]
  book.sheets[1].rangeProtectionRules = [protection(range(3, 3))]
  assert.equal(sheetClipboardCapabilities(input(book, snapshot)).paste.values.disabled, true)
  delete book.sheets[1].rangeProtectionRules
  book.sheets[1].filteredRows = [4]
  assert.equal(sheetClipboardCapabilities(input(book, snapshot)).paste.values.disabled, true)
  const shape = sheetClipboardCapabilities(input(book, snapshot, { shapeTextEditing: true }))
  assert.equal(shape.ribbonPaste.hidden, true); assert.equal(shape.copy.disabled, true); assert.equal(shape.paste.values.disabled, true)
})

test('single-cell column-width paste into a complete merge checks only the written anchor column and still its full height', () => {
  const book = workbook(), snapshot = clipboard(book, range(0, 0)), calls = []
  book.sheets[1].mergedRanges = [range(3, 3, 5, 5)]
  book.sheets[1].rangeProtectionRules = [protection(range(10, 5))]
  const result = sheetClipboardCapabilities(input(book, snapshot, { permission: request => {
    if (request.point === 'setColumnWidth') calls.push(request)
    return true
  } }))
  assert.equal(result.paste.columnWidths.disabled, false); assert.equal(result.paste.values.disabled, false)
  assert.deepEqual(calls[0].ranges, [range(0, 3, 39, 3)])
  book.sheets[1].rangeProtectionRules = [protection(range(10, 3))]
  assert.equal(sheetClipboardCapabilities(input(book, snapshot)).paste.columnWidths.disabled, true)
})

test('recent-paste mode prediction checks prior restoration permissions while ordinary new pastes keep their independent modes', () => {
  const book = workbook(), snapshot = clipboard(book), permission = request => request.point !== 'setColumnWidth'
  const ordinary = sheetClipboardCapabilities(input(book, snapshot, { permission }))
  assert.equal(ordinary.paste.values.disabled, false); assert.equal(ordinary.paste.formats.disabled, false)
  const recent = sheetClipboardCapabilities(input(book, snapshot, { permission, previousPasteMode: 'columnWidths' }))
  assert.ok(Object.values(recent.paste).every(entry => entry.disabled)); assert.equal(recent.pasteSpecial.disabled, true)
  assert.match(recent.paste.values.reason, /列宽/); assert.equal(recent.ribbonPaste.disabled, true)
  const noValue = request => request.point !== 'setCellValue'
  const ordinaryFormats = sheetClipboardCapabilities(input(book, snapshot, { permission: noValue }))
  assert.equal(ordinaryFormats.paste.formats.disabled, false); assert.equal(ordinaryFormats.paste.columnWidths.disabled, false)
  const restoreValues = sheetClipboardCapabilities(input(book, snapshot, { permission: noValue, previousPasteMode: 'values' }))
  assert.equal(restoreValues.paste.formats.disabled, true); assert.equal(restoreValues.paste.columnWidths.disabled, true)
})

test('fixed command aliases share capability objects and capability evaluation leaves the input model and clipboard untouched', () => {
  const book = workbook(), snapshot = clipboard(book), before = copy({ book, snapshot })
  const result = sheetClipboardCapabilities(input(book, snapshot))
  for (const [alias, canonical] of [
    ['sheet.command.copy-formula-only', 'sheet.command.copy-formulas'], ['sheet.command.paste-value', 'sheet.command.paste-values'],
    ['sheet.command.paste-formula', 'sheet.command.paste-formulas'], ['sheet.command.paste-format', 'sheet.command.paste-formats'],
    ['sheet.command.paste-col-width', 'sheet.command.paste-column-widths'], ['sheet.command.paste-besides-border', 'sheet.command.paste-except-borders'],
  ]) assert.equal(sheetClipboardCommandCapability(result, alias), sheetClipboardCommandCapability(result, canonical))
  assert.equal(sheetClipboardCommandCapability(result, 'sheet.command.optional-paste', 'formats'), result.paste.formats)
  assert.equal(sheetClipboardCommandCapability(result, 'sheet.command.unrelated'), undefined)
  assert.deepEqual(copy({ book, snapshot }), before)
})
