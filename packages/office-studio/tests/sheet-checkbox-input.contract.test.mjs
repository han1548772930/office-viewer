import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { sheetModelCellViewAt } from '../src/services/excelViewAdapter.ts'
import { createSheetCellEditSession } from '../src/features/sheet-workbench/sheet-cell-edit-session.ts'
import { normalizeSheetCheckboxValue, sheetCheckboxRuleValues, sheetCheckboxState } from '../src/features/sheet-validation/validation-checkbox.ts'
import { COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, sheetCheckboxInput } from '../src/features/sheet-validation/validation-checkbox-input.ts'

// Unexecuted implementation-phase drafts. The SSR harness, browser bindings and Office roundtrips remain deferred.
let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = (row = 0, column = 0, endRow = row, endCol = column) => ({ startRow: row, startCol: column, endRow, endCol })
function fixture(configure = () => {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'checkbox-input'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'
  sheet.rowCount = 50; sheet.columnCount = 15
  sheet.dataValidations = [{ id: 'checkbox-rule', type: 'checkbox', ranges: [area(0, 0, 2, 0)] }]
  sheet.cells['0:0'] = createCell(0, 0, 0); sheet.cells['1:0'] = createCell(1, 0, 0); sheet.cells['0:1'] = createCell(0, 1, 'other')
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other); configure(model, sheet)
  const studio = new OfficeStudio({ workbook: model, currentUserId: 'local-user' }), book = FOffice.newAPI(studio).getActiveWorkbook()
  const current = () => studio.getWorkbookModel(model.id), targetSheet = () => current().sheets.find(item => item.id === 'data')
  const target = (row = 0, column = 0) => {
    const workbook = current(), sheet = targetSheet(), merge = sheet.mergedRanges.find(range => row >= range.startRow && row <= range.endRow && column >= range.startCol && column <= range.endCol)
    const main = { row: merge?.startRow ?? row, column: merge?.startCol ?? column }
    const session = createSheetCellEditSession({ documentId: workbook.id, sheetId: sheet.id,
      key: `${String.fromCharCode(65 + main.column)}${main.row + 1}`, date1904: workbook.date1904,
      cell: sheetModelCellViewAt(workbook, sheet, main.row, main.column) }, { region: 'en-US', currentYear: 2026 })
    return { unitId: workbook.id, sheetId: sheet.id, row, column, ruleId: 'checkbox-rule', session,
      expectedRevision: studio.getWorkbookRevision(workbook.id), expectedHistoryGeneration: studio.editor.history.captureSavepoint(workbook.id).generation,
      expectedUserId: studio.currentUserId, expectedRuleSignature: JSON.stringify(sheet.dataValidations.find(rule => rule.id === 'checkbox-rule')) }
  }
  return { studio, book, current, sheet: targetSheet, target, cell: () => targetSheet().cells['0:0'],
    count: () => studio.editor.history.units[model.id]?.undo.length ?? 0, dispose: () => studio.dispose() }
}

test.each([true, 'TRUE', 'true', 1, '1'])('checkbox value %s shares the complete boolean alias without normalizing custom text', value => {
  assert.equal(normalizeSheetCheckboxValue(value), '1')
  assert.equal(normalizeSheetCheckboxValue(' true '), ' true ')
  assert.equal(normalizeSheetCheckboxValue('001'), '001')
})

test('default checkbox values and empty/legal/invalid rendering share one pure rule contract', () => {
  const rule = { id: 'r', type: 'checkbox', ranges: [area()] }, values = sheetCheckboxRuleValues(rule)
  assert.deepEqual(values, { checked: 1, unchecked: 0 })
  assert.deepEqual(sheetCheckboxState(values, null), { checked: false, renderable: true, valid: false })
  assert.deepEqual(sheetCheckboxState(values, true), { checked: true, renderable: true, valid: true })
  assert.deepEqual(sheetCheckboxState(values, 'FALSE'), { checked: false, renderable: true, valid: true })
  assert.equal(sheetCheckboxState(values, 'invalid').renderable, false)
  assert.equal(sheetCheckboxState(values, null, true).renderable, false)
  assert.equal(sheetCheckboxState(values, 0, false, true).renderable, false)
  assert.equal(sheetCheckboxRuleValues({ ...rule, formula1: 'TRUE', formula2: '1' }), undefined)
})

test('a rule formula is evaluated at the rule anchor and preserves its scalar result', () => {
  const calls = [], values = sheetCheckboxRuleValues({ id: 'r', type: 'checkbox', ranges: [area(10, 3)],
    formulaAnchor: { row: 2, col: 1 }, formula1: '=B3', formula2: 'TODO' }, (formula, anchor) => { calls.push([formula, anchor]); return true })
  assert.deepEqual(calls, [['=B3', { row: 2, column: 1 }]])
  assert.deepEqual(values, { checked: true, unchecked: 'TODO' })
})

test('the checkbox DTO keeps a frozen session and accepts only a boolean checked flag', () => {
  const h = fixture()
  try {
    const source = Object.freeze({ ...h.target(), checked: true }), parsed = sheetCheckboxInput(source)
    assert.notEqual(parsed, source); assert.equal(parsed.checked, true)
    assert.throws(() => sheetCheckboxInput({ ...source, checked: 1 }), /布尔/)
    assert.throws(() => sheetCheckboxInput({ ...source, editTrace: [] }), /文字输入/)
  } finally { h.dispose() }
})

test('one checkbox gesture commits only its main cell with one history and one value event', () => {
  const h = fixture(), events = [], subscription = h.studio.events.addEvent('SheetValueChanged', event => events.push(cloneModel(event)))
  try {
    const before = cloneModel(h.current()), result = h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, h.target())
    assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(h.cell().value, 1); assert.equal(h.cell().dataType, 'number')
    assert.deepEqual(h.sheet().cells['1:0'], before.sheets[0].cells['1:0']); assert.deepEqual(h.sheet().cells['0:1'], before.sheets[0].cells['0:1'])
    assert.equal(h.count(), 1); assert.equal(events.length, 1); assert.deepEqual(events[0].ranges[0], { unitId: 'checkbox-input', sheetId: 'data', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })
    const saved = cloneModel(h.current()); h.book.undo(); assert.deepEqual(h.sheet().cells, before.sheets[0].cells)
    h.book.redo(); assert.deepEqual(h.sheet().cells, saved.sheets[0].cells)
    const reopened = new OfficeStudio({ workbook: saved }); try { assert.deepEqual(reopened.getWorkbookModel(saved.id).sheets[0].cells, saved.sheets[0].cells) } finally { reopened.dispose() }
  } finally { subscription.dispose(); h.dispose() }
})

test.each([{ checked: 'DONE', unchecked: 'TODO' }, { checked: '001', unchecked: 'literal =text' }, { checked: 'TRUE', unchecked: 'FALSE' }])('custom checkbox $checked/$unchecked writes exact text rather than parsing labels', item => {
  const h = fixture((_model, sheet) => {
    sheet.dataValidations[0].formula1 = item.checked; sheet.dataValidations[0].formula2 = item.unchecked
    sheet.cells['0:0'].value = sheet.dataValidations[0].formula2; sheet.cells['0:0'].metadata = { vendor: { retained: true } }
  })
  try {
    const metadata = cloneModel(h.cell().metadata), result = h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, h.target())
    assert.equal(result.accepted, true); assert.equal(h.cell().value, item.checked); assert.equal(h.cell().dataType, 'text'); assert.equal(h.cell().formula, null)
    assert.deepEqual(h.cell().metadata.vendor, metadata.vendor)
    assert.equal(h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, h.target()).accepted, true)
    assert.equal(h.cell().value, item.unchecked)
  } finally { h.dispose() }
})

test.each([true, 'TRUE', 1])('an explicit unchanged checked state preserves the original %s representation without history', original => {
  const h = fixture((_model, sheet) => { sheet.cells['0:0'].value = original; sheet.cells['0:0'].dataType = typeof original === 'boolean' ? 'boolean' : typeof original === 'string' ? 'text' : 'number' })
  try {
    const before = cloneModel(h.current()), result = h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, { ...h.target(), checked: true })
    assert.deepEqual(result, { accepted: true, changed: false, issues: [] }); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('an explicit unchecked state resolves a blank that disallows blanks, rather than accepting an invalid no-op', () => {
  const h = fixture((_model, sheet) => { sheet.cells['0:0'].value = null; sheet.dataValidations[0].allowBlank = false })
  try {
    const result = h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, { ...h.target(), checked: false })
    assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(h.cell().value, 0); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test.each(['value', 'richText', 'image'])('existing incompatible %s is not overwritten by a checkbox command', kind => {
  const h = fixture((_model, sheet) => {
    if (kind === 'value') sheet.cells['0:0'].value = 'invalid'
    else if (kind === 'richText') sheet.cells['0:0'].richText = [{ text: 'retained text' }]
    else sheet.cells['0:0'].image = { id: 'cell-image', altText: 'retained image', title: 'Image', resource: {
      id: 'image-resource', name: 'pixel.png', mimeType: 'image/png', width: 1, height: 1, byteLength: 68,
      source: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==' } }
  })
  try {
    const before = cloneModel(h.current()), result = h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, h.target())
    assert.equal(result.accepted, false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('a formula-backed checkbox keeps its formula on no-op and clears it only on a real toggle', () => {
  const h = fixture((_model, sheet) => { sheet.cells['0:0'].formula = '=TRUE'; sheet.cells['0:0'].value = true; sheet.cells['0:0'].dataType = 'formula' })
  try {
    assert.equal(h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, { ...h.target(), checked: true }).changed, false)
    assert.equal(h.cell().formula, '=TRUE'); assert.equal(h.count(), 0)
    const result = h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, h.target())
    assert.equal(result.accepted, true); assert.equal(h.cell().value, 0); assert.equal(h.cell().formula, null); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test.each(['view', 'edit'])('a merged covered-cell toggle requires complete merge %s permission', kind => {
  const h = fixture((_model, sheet) => {
    sheet.mergedRanges = [area(0, 0, 1, 1)]; sheet.dataValidations[0].ranges = [area(0, 0)]
    sheet.rangeProtectionRules = [{ id: 'restricted', permissionId: 'restricted', ownerUserId: 'peer', allowedUserIds: [], ranges: [area(1, 1)],
      editState: 'onlyMe', viewState: kind === 'view' ? 'noOneElseCanView' : 'othersCanView' }]
  })
  try {
    const before = cloneModel(h.current()); assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, h.target(1, 1)), /权限|保护/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('a merged covered-cell toggle changes the primary value and preserves other existing covered cells', () => {
  const h = fixture((_model, sheet) => { sheet.mergedRanges = [area(0, 0, 1, 1)]; sheet.dataValidations[0].ranges = [area(0, 0)] })
  try {
    const previous = cloneModel(h.sheet().cells['1:0']), result = h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, h.target(1, 1))
    assert.equal(result.accepted, true); assert.equal(h.cell().value, 1); assert.deepEqual(h.sheet().cells['1:0'], previous); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test('a checkbox formula cannot disclose or commit a hidden direct or transitive source value', () => {
  const h = fixture((_model, sheet) => {
    sheet.cells['0:1'] = createCell(0, 1, 'secret'); sheet.cells['0:2'] = createCell(0, 2, 'secret'); sheet.cells['0:2'].formula = '=B1'
    sheet.dataValidations[0].formula1 = '=C1'; sheet.dataValidations[0].formula2 = 'TODO'; sheet.cells['0:0'].value = 'TODO'
    sheet.rangeProtectionRules = [{ id: 'hidden', permissionId: 'hidden', ownerUserId: 'peer', allowedUserIds: [], ranges: [area(0, 1)], editState: 'onlyMe', viewState: 'noOneElseCanView' }]
  })
  try {
    const before = cloneModel(h.current()), result = h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, h.target())
    assert.equal(result.accepted, false); assert.equal(JSON.stringify(result).includes('secret'), false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('stale revision and target session, cancellation and prepare-view reentry never publish an extra checkbox history', () => {
  const h = fixture()
  try {
    const old = h.target(), canceled = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === COMMIT_SHEET_CHECKBOX_INPUT_COMMAND) event.cancel = true })
    const before = cloneModel(h.current()); assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, old), /canceled/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0); canceled.dispose()
    h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, old); const saved = cloneModel(h.current())
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, old), /更改|版本/)
    assert.deepEqual(h.current(), saved); assert.equal(h.count(), 1)
    const commit = h.studio.commitWorkbook.bind(h.studio)
    h.studio.commitWorkbook = (id, label, mutate, options) => commit(id, label, mutate, { ...options, prepareView: candidate => {
      h.studio.currentUserId = 'peer'; options.prepareView(candidate)
    } })
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, h.target()), /用户|更改|权限/)
    assert.deepEqual(h.current(), saved); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})
