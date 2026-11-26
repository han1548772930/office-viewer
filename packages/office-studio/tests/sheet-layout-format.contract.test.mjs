import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { applySheetLayoutFormatCommand, applySheetViewLayoutFormatCommand, getSheetViewLayoutFormatState,
  layoutFormatActionForCommand, SHEET_LAYOUT_FORMAT_COMMANDS, sheetCellTextRotation, sheetCellWrapStrategy } from '../src/features/sheet-workbench/sheet-layout-format-command.ts'
import { clearSheetCellFormat, setSheetTextFit } from '../src/features/sheet-workbench/sheet-toolbar-format.ts'
import { createSheetPastePlan } from '../src/features/sheet-workbench/sheet-paste.ts'
import { createSheetFillPlan } from '../src/features/sheet-workbench/sheet-fill.ts'
import { setSheetCellStyleProperties } from '../src/features/sheet-workbench/sheet-style-inheritance.ts'

const range = (startCol, endCol = startCol, startRow = 0, endRow = startRow) => ({ startRow, endRow, startCol, endCol })
const copy = value => structuredClone(value)

test('local and fixed-version commands normalize primitive, enum and explicit range inputs', () => {
  assert.deepEqual(layoutFormatActionForCommand(SHEET_LAYOUT_FORMAT_COMMANDS.horizontal, 'center'), { kind: 'horizontal', value: 'center' })
  assert.deepEqual(layoutFormatActionForCommand('sheet.command.set-horizontal-text-align', { value: 3 }), { kind: 'horizontal', value: 'right' })
  assert.deepEqual(layoutFormatActionForCommand('sheet.command.set-vertical-text-align', { value: 1 }), { kind: 'vertical', value: 'top' })
  for (const [value, strategy] of [[1, 'overflow'], [2, 'clip'], [3, 'wrap']]) {
    assert.deepEqual(layoutFormatActionForCommand('sheet.command.set-text-wrap', { value }), { kind: 'wrap', value: strategy })
  }
  assert.deepEqual(layoutFormatActionForCommand('sheet.command.set-wrap', false), { kind: 'wrap', value: 'overflow' })
  assert.deepEqual(layoutFormatActionForCommand(SHEET_LAYOUT_FORMAT_COMMANDS.shrink), { kind: 'shrink' })
  assert.deepEqual(layoutFormatActionForCommand(SHEET_LAYOUT_FORMAT_COMMANDS.rotation, { value: 'v' }), { kind: 'rotation', value: 'v' })
  assert.deepEqual(layoutFormatActionForCommand(SHEET_LAYOUT_FORMAT_COMMANDS.rotation, -37), { kind: 'rotation', value: -37 })
  assert.equal(layoutFormatActionForCommand('unrelated.command', 'wrap'), undefined)
  assert.equal(layoutFormatActionForCommand('sheet.command.set-text-wrap'), undefined)
  for (const value of [NaN, Infinity, -91, 91, 'diagonal']) assert.throws(() => layoutFormatActionForCommand(SHEET_LAYOUT_FORMAT_COMMANDS.rotation, value), RangeError)
  assert.throws(() => layoutFormatActionForCommand(SHEET_LAYOUT_FORMAT_COMMANDS.vertical, 'left'), TypeError)
  assert.throws(() => layoutFormatActionForCommand('sheet.command.set-text-wrap', 99), TypeError)
})

test('enabling wrap normalizes a previous clip strategy and clears shrink in both model and standalone hosts', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  const formula = { ...createCell(0, 0, null), formula: '=UNSUPPORTED()', dataType: 'formula', cachedDataType: null,
    display: 'source cache', richText: [{ text: 'kept rich source' }], metadata: { vendor: 'kept' } }
  formula.style.wrapStrategy = 'clip'; formula.style.shrinkToFit = true; formula.style.font.bold = true
  formula.style.metadata = { inherited: 'kept' }
  sheet.cells['0:0'] = formula
  const view = toExcelWorkbookView(workbook).sheets[0]
  const before = copy(formula)
  const action = layoutFormatActionForCommand('sheet.command.set-wrap', true)
  assert.equal(applySheetLayoutFormatCommand(sheet, [range(0)], action), true)
  assert.equal(applySheetViewLayoutFormatCommand(view, [range(0)], action), true)
  assert.equal(sheet.cells['0:0'].style.wrap, true)
  assert.equal(sheet.cells['0:0'].style.wrapStrategy, 'wrap')
  assert.equal(sheet.cells['0:0'].style.shrinkToFit, false)
  for (const key of ['value', 'formula', 'display', 'dataType', 'cachedDataType', 'richText', 'metadata']) assert.deepEqual(sheet.cells['0:0'][key], before[key])
  assert.deepEqual(sheet.cells['0:0'].style.font, before.style.font)
  assert.deepEqual(sheet.cells['0:0'].style.metadata, before.style.metadata)
  const restored = applyExcelWorkbookView(workbook, { ...toExcelWorkbookView(workbook), sheets: [view] })
  const restoredStyle = restored.sheets[0].cells['0:0'].style
  for (const key of ['wrap', 'wrapStrategy', 'shrinkToFit', 'horizontal', 'vertical', 'font', 'metadata']) {
    assert.deepEqual(restoredStyle[key], sheet.cells['0:0'].style[key])
  }
  assert.equal(applySheetLayoutFormatCommand(sheet, [range(0)], action), false)
})

test('shrink disables wrapped targets while preserving overflow or clip and uses the primary when toggled', () => {
  const sheet = createBlankSheet()
  for (const [col, strategy] of ['overflow', 'clip', 'wrap'].entries()) {
    const cell = createCell(0, col, `cell ${col}`)
    cell.style.wrapStrategy = strategy; cell.style.wrap = strategy === 'wrap'
    sheet.cells[`0:${col}`] = cell
  }
  assert.equal(applySheetLayoutFormatCommand(sheet, [range(0, 2)], { kind: 'shrink', value: true }), true)
  assert.deepEqual([0, 1, 2].map(col => sheetCellWrapStrategy(sheet.cells[`0:${col}`].style)), ['overflow', 'clip', 'overflow'])
  assert.ok(Object.values(sheet.cells).every(cell => cell.style.shrinkToFit))
  assert.equal(applySheetLayoutFormatCommand(sheet, [range(0, 2)], { kind: 'shrink' }), true)
  assert.ok(Object.values(sheet.cells).every(cell => cell.style.shrinkToFit === false))
  assert.deepEqual([0, 1, 2].map(col => sheetCellWrapStrategy(sheet.cells[`0:${col}`].style)), ['overflow', 'clip', 'overflow'])
})

test('shrink toggles from the supplied primary or the last range once across mixed targets', () => {
  for (const [primary, targets, expected] of [
    [{ row: 0, col: 2 }, [range(0, 2), range(1, 2)], [false, false, false]],
    [undefined, [range(0), range(2)], [false, true, false]],
  ]) {
    const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
    for (const [col, shrinkToFit] of [false, true, true].entries()) {
      sheet.cells[`0:${col}`] = createCell(0, col, `cell ${col}`)
      sheet.cells[`0:${col}`].style.shrinkToFit = shrinkToFit
    }
    const view = toExcelWorkbookView(workbook).sheets[0]
    assert.equal(applySheetLayoutFormatCommand(sheet, targets, { kind: 'shrink' }, 'local-user', [], workbook, primary), true)
    assert.equal(applySheetViewLayoutFormatCommand(view, targets, { kind: 'shrink' }, 'local-user', primary), true)
    assert.deepEqual([0, 1, 2].map(col => sheet.cells[`0:${col}`].style.shrinkToFit), expected)
    assert.deepEqual(['A1', 'B1', 'C1'].map(key => view.cells[key].shrinkToFit), expected)
    assert.deepEqual([0, 1, 2].map(col => sheet.cells[`0:${col}`].value), ['cell 0', 'cell 1', 'cell 2'])
    assert.equal(applySheetLayoutFormatCommand(sheet, targets, { kind: 'shrink' }, 'local-user', [], workbook, primary), true)
    assert.equal(applySheetViewLayoutFormatCommand(view, targets, { kind: 'shrink' }, 'local-user', primary), true)
    assert.ok(Object.values(sheet.cells).every(cell => cell.style.shrinkToFit))
    assert.ok(Object.values(view.cells).every(cell => cell.shrinkToFit))
  }
})

test('shrink reads the composed primary style without populating unaffected empty cells', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.columnStyles = { 2: { shrinkToFit: true } }
  const view = toExcelWorkbookView(workbook).sheets[0], targets = [range(0), range(2)]
  assert.equal(applySheetLayoutFormatCommand(sheet, targets, { kind: 'shrink' }, 'local-user', [], workbook), true)
  assert.equal(applySheetViewLayoutFormatCommand(view, targets, { kind: 'shrink' }), true)
  assert.deepEqual(Object.keys(sheet.cells), ['0:2'])
  assert.deepEqual(Object.keys(view.cells), ['C1'])
  assert.equal(sheet.cells['0:2'].style.shrinkToFit, false)
  assert.equal(view.cells.C1.shrinkToFit, false)
  assert.deepEqual(sheet.columnStyles, { 2: { shrinkToFit: true } })
})

test('invalid shrink primary positions reject both hosts before writing any target', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, 'kept')
  const view = toExcelWorkbookView(workbook).sheets[0], initialModel = copy(sheet), initialView = copy(view)
  for (const primary of [{ row: 4, col: 4 }, { row: -.5, col: 0 }, { row: 0, col: 16_384 }, { row: NaN, col: 0 }, {}, null, 1]) {
    assert.throws(() => applySheetLayoutFormatCommand(sheet, [range(0, 2)], { kind: 'shrink' }, 'local-user', [], workbook, primary))
    assert.throws(() => applySheetViewLayoutFormatCommand(view, [range(0, 2)], { kind: 'shrink' }, 'local-user', primary))
    assert.deepEqual(sheet, initialModel)
    assert.deepEqual(view, initialView)
  }
})

test('mixed selection toggles once across overlapping ranges and skips filtered rows without changing manual hiding', () => {
  const sheet = createBlankSheet()
  for (let row = 0; row < 3; row++) {
    const cell = createCell(row, 0, `row ${row}`)
    cell.style.wrap = row === 0
    sheet.cells[`${row}:0`] = cell
  }
  sheet.hiddenRows = [2]
  const filtered = copy(sheet.cells['1:0'])
  const scopes = [range(0, 0, 0, 2), range(0, 0, 0, 0)]
  applySheetLayoutFormatCommand(sheet, scopes, { kind: 'toggle-wrap' }, 'local-user', [1])
  assert.equal(sheetCellWrapStrategy(sheet.cells['0:0'].style), 'wrap')
  assert.equal(sheetCellWrapStrategy(sheet.cells['2:0'].style), 'wrap')
  assert.deepEqual(sheet.cells['1:0'], filtered)
  assert.deepEqual(sheet.hiddenRows, [2])
  applySheetLayoutFormatCommand(sheet, scopes, { kind: 'toggle-wrap' }, 'local-user', [1])
  assert.equal(sheetCellWrapStrategy(sheet.cells['0:0'].style), 'overflow')
  assert.equal(sheetCellWrapStrategy(sheet.cells['2:0'].style), 'overflow')
  assert.deepEqual(sheet.cells['1:0'], filtered)
})

test('layout menus read the supplied primary or last range rather than combining mixed cell states', () => {
  const sheet = { id: 'state', name: 'State', rowCount: 1_048_576, columnCount: 16_384,
    cells: {
      A1: { value: 'one', align: 'right', vertical: 'top', wrap: true, shrinkToFit: true, rotationVertical: true },
      B1: { value: 'two', align: 'center', vertical: 'bottom', wrapStrategy: 'clip', shrinkToFit: false, rotation: 45 },
    } }
  assert.deepEqual(getSheetViewLayoutFormatState(sheet, [range(0, 1)]), {
    align: 'right', vertical: 'top', wrapStrategy: 'wrap', wrap: true, shrinkToFit: true, rotation: 0, rotationVertical: true,
  })
  const second = { align: 'center', vertical: 'bottom', wrapStrategy: 'clip', wrap: false, shrinkToFit: false, rotation: 45, rotationVertical: false }
  assert.deepEqual(getSheetViewLayoutFormatState(sheet, [range(0, 1)], { row: 0, col: 1 }), second)
  assert.deepEqual(getSheetViewLayoutFormatState(sheet, [range(0), range(1)]), second)
  sheet.filteredRows = [0]
  assert.deepEqual(getSheetViewLayoutFormatState(sheet, [range(0, 16_383, 0, 1_048_575)], { row: 0, col: 1 }), second)
  assert.deepEqual(Object.keys(sheet.cells), ['A1', 'B1'])
  assert.equal(sheetCellTextRotation({ rotation: 255 }), 'v')
  assert.equal(sheetCellWrapStrategy({ wrap: true, wrapStrategy: 'clip' }), 'clip')
})

test('layout menus resolve a merged primary to its content owner and retain composed sparse styles', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.defaultStyle = { horizontal: 'right', vertical: 'bottom', wrapStrategy: 'clip' }
  sheet.columnStyles = { 1: { horizontal: 'center', wrapStrategy: 'wrap', shrinkToFit: true, rotation: -45 } }
  sheet.cells['0:0'] = createCell(0, 0, 'master')
  setSheetCellStyleProperties(sheet.cells['0:0'], { vertical: 'top', rotationVertical: true })
  sheet.mergedRanges = [range(0, 1, 0, 1)]
  const view = toExcelWorkbookView(workbook).sheets[0], initial = copy(view)
  const merged = getSheetViewLayoutFormatState(view, [range(0, 1, 0, 1)], { row: 1, col: 1 })
  assert.equal(merged.vertical, 'top')
  assert.equal(merged.rotationVertical, true)
  assert.equal(merged.wrapStrategy, 'clip')
  assert.deepEqual(getSheetViewLayoutFormatState(view, [range(1, 1, 4, 4)], { row: 4, col: 1 }), {
    align: 'center', vertical: 'bottom', wrapStrategy: 'wrap', wrap: true, shrinkToFit: true, rotation: -45, rotationVertical: false,
  })
  assert.deepEqual(getSheetViewLayoutFormatState(view, [range(3, 3, 99, 99)]), {
    align: 'right', vertical: 'bottom', wrapStrategy: 'clip', wrap: false, shrinkToFit: false, rotation: 0, rotationVertical: false,
  })
  assert.deepEqual(view, initial)
})

test('invalid or protected ranges reject all writes and default layout operations do not create empty cells', () => {
  const sheet = createBlankSheet(), initial = copy(sheet)
  assert.equal(applySheetLayoutFormatCommand(sheet, [range(0, 4)], { kind: 'horizontal', value: 'general' }), false)
  assert.deepEqual(sheet, initial)
  assert.throws(() => applySheetLayoutFormatCommand(sheet, [range(0), range(16_384)], { kind: 'rotation', value: 45 }), RangeError)
  assert.deepEqual(sheet, initial)
  sheet.rangeProtectionRules = [{ id: 'denied', permissionId: 'denied-p', ranges: [range(1)],
    ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }]
  const locked = copy(sheet)
  assert.throws(() => applySheetLayoutFormatCommand(sheet, [range(0, 1)], { kind: 'wrap', value: 'wrap' }, 'guest'), /保护/)
  assert.deepEqual(sheet, locked)
  const workbook = createBlankWorkbook(); workbook.sheets = [sheet]
  const view = toExcelWorkbookView(workbook).sheets[0], lockedView = copy(view)
  assert.throws(() => applySheetViewLayoutFormatCommand(view, [range(0, 1)], { kind: 'rotation', value: 'v' }, 'guest'), /保护/)
  assert.deepEqual(view, lockedView)
})

test('clear, special paste and format-only fill preserve effective wrap semantics without changing content', () => {
  const source = { value: 'source', wrap: true, wrapStrategy: 'wrap', shrinkToFit: false }
  const target = { value: '', formula: '=A1', rawValue: null, cachedValue: 'keep cache', wrapStrategy: 'clip' }
  const cleared = clearSheetCellFormat(source)
  assert.equal(sheetCellWrapStrategy(cleared), 'overflow')
  assert.equal(cleared.value, 'source')
  assert.equal(sheetCellWrapStrategy(setSheetTextFit(source, 'shrinkToFit', true)), 'overflow')
  const pasted = createSheetPastePlan({ mode: 'formats', sourceCells: [[source]], targetCells: [[target]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 1, targetStartColumn: 0 }).cellPatches[0].cell
  assert.equal(sheetCellWrapStrategy(pasted), 'wrap')
  assert.equal(pasted.formula, '=A1')
  assert.equal(pasted.cachedValue, 'keep cache')
  const filled = createSheetFillPlan({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, { row: 1, column: 0 }, [[source]],
    { applyType: 'ONLY_FORMAT', targetCells: [[source], [target]] }).cellPatches[0].cell
  assert.equal(sheetCellWrapStrategy(filled), 'wrap')
  assert.equal(filled.formula, '=A1')
  assert.equal(filled.cachedValue, 'keep cache')
  const legacySource = { value: 'legacy', wrap: true }
  const legacyPaste = createSheetPastePlan({ mode: 'formats', sourceCells: [[legacySource]], targetCells: [[target]],
    sourceStartRow: 0, sourceStartColumn: 0, targetStartRow: 1, targetStartColumn: 0 }).cellPatches[0].cell
  assert.equal(legacyPaste.wrapStrategy, 'wrap')
  assert.equal(legacyPaste.wrap, true)
  const legacyFill = createSheetFillPlan({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, { row: 1, column: 0 }, [[legacySource]],
    { applyType: 'ONLY_FORMAT', targetCells: [[legacySource], [target]] }).cellPatches[0].cell
  assert.equal(sheetCellWrapStrategy(legacyFill), 'wrap')
})

test('Facade, explicit upstream command ranges, cancellation, history and internal reopen use one layout path', async () => {
  const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  try {
    const { FOffice, OfficeStudio } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts')
    const workbook = createBlankWorkbook(); workbook.id = 'layout-commands'; workbook.calculationMode = 'manual'
    workbook.sheets[0].cells['0:0'] = { ...createCell(0, 0, null), formula: '=UNSUPPORTED()', display: 'kept', dataType: 'formula', cachedDataType: null }
    for (const [col, shrinkToFit] of [false, true, true].entries()) {
      workbook.sheets[0].cells[`0:${col + 1}`] = createCell(0, col + 1, `primary ${col}`)
      workbook.sheets[0].cells[`0:${col + 1}`].style.shrinkToFit = shrinkToFit
    }
    const studio = new OfficeStudio({ workbook })
    try {
      const api = FOffice.newAPI(studio), book = api.getActiveWorkbook(), sheet = book.getActiveSheet(), target = sheet.getRange('A1')
      const commands = []
      studio.events.addEvent('CommandExecuted', event => commands.push(event.id))
      target.setWrapStrategy('clip').setWrap(true).setHorizontalAlignment('right').setVerticalAlignment('bottom').setTextRotation('v')
      assert.equal(target.getWrap(), true)
      assert.equal(target.getWrapStrategy(), 'wrap')
      assert.equal(target.getHorizontalAlignment(), 'right')
      assert.equal(target.getVerticalAlignment(), 'bottom')
      assert.equal(target.getTextRotation(), 'v')
      assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['0:0'].display, 'kept')
      assert.ok(commands.includes(SHEET_LAYOUT_FORMAT_COMMANDS.wrap))
      const undoCount = studio.editor.history.units[book.getId()].undo.length
      target.setTextRotation('v')
      assert.equal(studio.editor.history.units[book.getId()].undo.length, undoCount)
      const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === SHEET_LAYOUT_FORMAT_COMMANDS.shrink) event.cancel = true })
      const previous = copy(studio.getWorkbookModel(book.getId()))
      assert.throws(() => target.setShrinkToFit(true), /canceled/)
      assert.deepEqual(studio.getWorkbookModel(book.getId()), previous)
      assert.equal(studio.editor.history.units[book.getId()].undo.length, undoCount)
      cancel.dispose()
      studio.executeCommandSync('sheet.command.set-text-wrap', { unitId: book.getId(), subUnitId: sheet.getSheetId(),
        range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, value: 2 })
      assert.equal(target.getWrapStrategy(), 'clip')
      book.undo(); assert.equal(target.getWrapStrategy(), 'wrap')
      book.redo(); assert.equal(target.getWrapStrategy(), 'clip')
      const serialized = book.save()
      const reopened = api.createWorkbook({ ...serialized, id: 'layout-reopened' }).getActiveSheet().getRange('A1')
      assert.equal(reopened.getWrapStrategy(), 'clip')
      assert.equal(reopened.getTextRotation(), 'v')
      target.setWrap(false)
      assert.equal(target.getWrapStrategy(), 'overflow')
      assert.equal(reopened.getWrapStrategy(), 'clip')
      studio.setSelections(book.getId(), sheet.getSheetId(), [{ startRow: 0, endRow: 0, startColumn: 1, endColumn: 2 }],
        { primary: { row: 0, column: 2 } })
      assert.equal(studio.executeCommandSync(SHEET_LAYOUT_FORMAT_COMMANDS.shrink), true)
      assert.deepEqual([1, 2, 3].map(col => studio.getWorkbookModel(book.getId()).sheets[0].cells[`0:${col}`].style.shrinkToFit), [false, false, true])
      book.undo()
      assert.deepEqual([1, 2, 3].map(col => studio.getWorkbookModel(book.getId()).sheets[0].cells[`0:${col}`].style.shrinkToFit), [false, true, true])
      book.redo()
      const scopes = [range(1), range(3)], input = { unitId: book.getId(), sheetId: sheet.getSheetId(), ranges: scopes }
      assert.equal(studio.executeCommandSync(SHEET_LAYOUT_FORMAT_COMMANDS.shrink, input), true)
      assert.deepEqual([1, 2, 3].map(col => studio.getWorkbookModel(book.getId()).sheets[0].cells[`0:${col}`].style.shrinkToFit), [false, false, false])
      assert.equal(studio.executeCommandSync(SHEET_LAYOUT_FORMAT_COMMANDS.shrink, { ...input, primary: { row: 0, col: 1 } }), true)
      assert.deepEqual([1, 2, 3].map(col => studio.getWorkbookModel(book.getId()).sheets[0].cells[`0:${col}`].style.shrinkToFit), [true, false, true])
      const beforeInvalid = copy(studio.getWorkbookModel(book.getId())), beforeInvalidHistory = studio.editor.history.units[book.getId()].undo.length
      for (const primary of [{ row: 4, col: 4 }, null]) {
        assert.throws(() => studio.executeCommandSync(SHEET_LAYOUT_FORMAT_COMMANDS.shrink, { ...input, primary }))
        assert.deepEqual(studio.getWorkbookModel(book.getId()), beforeInvalid)
        assert.equal(studio.editor.history.units[book.getId()].undo.length, beforeInvalidHistory)
      }
      const primaryReopened = api.createWorkbook({ ...book.save(), id: 'layout-primary-reopened' })
      assert.deepEqual([1, 2, 3].map(col => studio.getWorkbookModel(primaryReopened.getId()).sheets[0].cells[`0:${col}`].style.shrinkToFit), [true, false, true])
    } finally { studio.dispose() }
  } finally { await vite.close() }
})
