import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// These drafts are reserved for the user's unified verification pass.
let server, sdk, visibility, structure, dimensions, visibilityUi, structureUi, dimensionUi
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[sdk, visibility, structure, dimensions, visibilityUi, structureUi, dimensionUi] = await Promise.all([
    '/packages/office-studio/src/index.ts',
    '/packages/office-studio/src/features/sheet-workbench/sheet-axis-visibility.ts',
    '/packages/office-studio/src/features/sheet-workbench/sheet-axis-structure-command.ts',
    '/packages/office-studio/src/domain/sheet-axis.ts',
    '/packages/office-studio/src/components/spreadsheet/use-sheet-axis-visibility.ts',
    '/packages/office-studio/src/components/spreadsheet/use-sheet-axis-structure.ts',
    '/packages/office-studio/src/components/spreadsheet/use-sheet-axis-dimensions.ts',
  ].map(path => server.ssrLoadModule(path)))
})
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const range = (startRow, endRow = startRow, startCol = 0, endCol = 7) => ({ startRow, endRow, startCol, endCol })
const rowAction = visible => ({ axis: 'row', visible, selectedOnly: false })
function workbook() {
  const model = sdk.createBlankWorkbook(); model.id = 'axis-edit-book'; model.calculationMode = 'manual'
  model.sheets = ['Target', 'Report'].map((name, index) => ({ ...sdk.createBlankSheet(name), id: 'axis-edit-sheet-' + index, rowCount: 12, columnCount: 8 }))
  model.activeSheetId = model.sheets[0].id
  return model
}
function runtime(t, source = workbook()) {
  const studio = new sdk.OfficeStudio({ workbook: source, sheetTextMeasure: text => [...text].length * 8 })
  t.onTestFinished(() => studio.dispose())
  const api = sdk.FOffice.newAPI(studio), book = api.getActiveWorkbook(), [first, second] = book.getSheets()
  return { studio, api, book, first, second }
}
const history = (studio, id = studio.activeWorkbookId) => {
  const unit = studio.editor.history.units[id]
  return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, revision: studio.getWorkbookRevision(id) }
}
function host(studio, source = workbook()) {
  let view = sdk.toExcelWorkbookView(studio ? studio.getWorkbookModel(studio.activeWorkbookId) : source)
  if (!studio) delete view.id
  let sheetId = view.activeSheetId, active = true
  const calls = { applied: 0, selections: 0, changed: 0, focused: 0, options: undefined, projection: undefined }
  const context = {
    studio: () => studio, workbook: () => view, sheetId: () => sheetId, ranges: () => [range(1)],
    userId: () => 'local-user', region: () => 'en-US', active: () => active, finishEdit: () => true,
    prepareView: next => next,
    applyView: next => { view = next; sheetId = next.activeSheetId; calls.applied++; calls.projection = studio?.getSelectionState() },
    selection: () => { calls.selections++ },
    changed: options => { calls.changed++; calls.options = options },
    focus: () => { calls.focused++ },
  }
  return { context, calls, view: () => view, replace: next => { view = next }, deactivate: () => { active = false } }
}
const uiCommands = () => [
  { id: 'sheet.command.set-rows-hidden', event: 'SheetVisibilityChanged', selects: true,
    create: context => visibilityUi.useSheetAxisVisibility(context).runAxisVisibilityCommand },
  { id: 'sheet.command.insert-row-before', event: 'SheetStructureChanged', selects: true,
    create: context => structureUi.useSheetAxisStructure(context).runAxisStructureCommand },
  { id: 'sheet.command.set-row-height', event: 'SheetDimensionsChanged', selects: false, params: { value: 38 }, create: context => {
    const runner = dimensionUi.useSheetAxisDimensions(context).runAxisDimensionCommand
    return (id, params) => runner(id, { value: 38, ...params })
  } },
]

test('a million-row hide retains compact spans and selects the remaining visible row without mutating its source', () => {
  const sheet = workbook().sheets[0]; sheet.rowCount = 1_048_576
  const before = copy(sheet), plan = visibility.planSheetAxisVisibilityCommand(sheet, [range(1, 1_048_575)], rowAction(false), {}, 'local-user')
  assert.equal(plan.changed, true); assert.deepEqual(plan.changes.hiddenRows, [])
  assert.deepEqual(plan.changes.rowHiddenSpans, [{ start: 1, end: 1_048_575, value: true }])
  assert.deepEqual(plan.selections, [range(0)]); assert.deepEqual(sheet, before)
  const axis = new dimensions.SparseAxisIndex(sheet.rowCount, 26, {}, [], [], plan.changes.rowHiddenSpans)
  assert.equal(axis.nextVisible(1, 1), undefined); assert.equal(axis.nextVisible(1_048_575, -1), 0)
})

test('show removes manual reasons while retaining filters, collapsed groups and the upstream two-stage zero-size contract', () => {
  const sheet = workbook().sheets[0]
  sheet.hiddenRows = [8]; sheet.rowHiddenSpans = [{ start: 1, end: 4, value: true }]
  sheet.rowHeights = { 1: 0 }; sheet.autoFilter = { range: range(0, 11), columns: {}, filteredRows: [2] }
  sheet.outline = { groups: [{ id: 'collapsed', axis: 'row', start: 3, end: 4, collapsed: true }], summaryBelow: true, summaryRight: true }
  const before = copy(sheet), first = visibility.planSheetAxisVisibilityCommand(sheet, [range(1, 4)], rowAction(true), {}, 'local-user')
  assert.deepEqual(first.changes.rowHiddenSpans, [{ start: 8, end: 8, value: true }])
  assert.equal(first.changes.rowHeights, undefined); assert.deepEqual(sheet, before)
  const shown = { ...sheet, ...first.changes }
  assert.deepEqual(shown.autoFilter.filteredRows, [2]); assert.deepEqual(shown.outline, before.outline)
  const second = visibility.planSheetAxisVisibilityCommand(shown, [range(1)], rowAction(true), {}, 'local-user')
  assert.deepEqual(second.changes.rowHeights, {})
  assert.deepEqual(second.changes.rowHeightSpans, [{ start: 1, end: 1, value: 26 }])
  assert.deepEqual(second.changes.rowAutoHeightSpans, [{ start: 1, end: 1, value: false }])
})

test('hide-all rejects the effective filter/outline/zero-size combination and explicit scope parameters cannot be ignored', () => {
  const sheet = workbook().sheets[0]
  sheet.autoFilter = { range: range(0, 11), columns: {}, filteredRows: [1, 2, 3, 4] }
  sheet.rowHiddenSpans = [{ start: 5, end: 9, value: true }]; sheet.rowHeights = { 10: 0, 11: 0 }
  const before = copy(sheet)
  assert.throws(() => visibility.planSheetAxisVisibilityCommand(sheet, [range(0)], rowAction(false), {}, 'local-user'), error => error.code === 'SHEET_AXIS_HIDE_ALL')
  assert.deepEqual(sheet, before)
  assert.throws(() => visibility.assertSheetAxisHasVisiblePosition({ ...sheet, rowHiddenSpans: [{ start: 0, end: 11, value: true }] }, 'row'), /至少保留/)
  assert.throws(() => visibility.sheetAxisVisibilityCommandInput({ count: 2 }), /起始索引/)
  assert.throws(() => visibility.sheetAxisVisibilityCommandInput({ all: true, index: 1 }), /同时/)
  assert.throws(() => visibility.sheetAxisVisibilityTargetRanges(sheet, rowAction(true), [], { index: 11, count: 2 }), /超出/)
  assert.deepEqual(visibility.sheetAxisVisibilityTargetRanges(sheet, { ...rowAction(true), selectedOnly: true }, [range(1, 2, 1, 2)]), [])
  assert.deepEqual(visibility.sheetAxisVisibilityTargetRanges(sheet, rowAction(true), [], { index: 2, count: 3 }), [range(2, 4)])
})

test('insert plans shift cells, cross-sheet references, merges and hidden dimensions together while inheriting preceding cell styles', () => {
  const source = workbook(), [sheet, report] = source.sheets
  sheet.cells['0:0'] = sdk.createCell(0, 0, 'style source'); sheet.cells['0:0'].style.font.bold = true
  sheet.cells['1:0'] = sdk.createCell(1, 0, 'shifted content')
  sheet.rowHeights = { 1: 40 }; sheet.hiddenRows = [2]; sheet.rowHiddenSpans = [{ start: 3, end: 4, value: true }]
  sheet.mergedRanges = [range(3, 4, 0, 1)]
  report.cells['0:0'] = { ...sdk.createCell(0, 0, null), formula: "='Target'!A2" }
  const before = copy(source), plan = structure.planWorkbookAxisStructureCommand(source, sheet.id,
    [{ axis: 'row', kind: 'insert', start: 1, count: 2 }], 'local-user', 'before')
  const result = plan.workbook.sheets[0]
  assert.equal(result.rowCount, 14); assert.equal(result.cells['3:0'].value, 'shifted content')
  for (const row of [1, 2]) { assert.equal(result.cells[row + ':0'].value, null); assert.equal(result.cells[row + ':0'].style.font.bold, true) }
  assert.equal(dimensions.sheetRowHeight(result, 1), 40); assert.equal(dimensions.sheetRowHeight(result, 2), 40)
  assert.deepEqual(result.hiddenRows, [4]); assert.deepEqual(result.rowHiddenSpans, [{ start: 5, end: 6, value: true }])
  assert.deepEqual(result.mergedRanges, [range(5, 6, 0, 1)])
  assert.equal(plan.workbook.sheets[1].cells['0:0'].formula, "='Target'!A4")
  assert.deepEqual(plan.selection, range(1, 2)); assert.deepEqual(source, before)
})

test('overlapping row deletions skip filtered axes and execute from the bottom in one detached plan', () => {
  const source = workbook(), sheet = source.sheets[0]
  for (let row = 0; row < 12; row++) sheet.cells[row + ':0'] = sdk.createCell(row, 0, 'row ' + row)
  const selected = [range(2, 4), range(4, 7)], before = copy(source)
  const edits = structure.sheetAxisStructureEdits(sheet, selected, { axis: 'row', kind: 'delete' }, {}, [3, 5])
  assert.deepEqual(edits, [{ axis: 'row', kind: 'delete', start: 6, count: 2 }, { axis: 'row', kind: 'delete', start: 4, count: 1 }, { axis: 'row', kind: 'delete', start: 2, count: 1 }])
  const result = structure.planWorkbookAxisStructureCommand(source, sheet.id, edits).workbook.sheets[0]
  assert.equal(result.rowCount, 8); assert.equal(result.cells['2:0'].value, 'row 3'); assert.equal(result.cells['3:0'].value, 'row 5')
  assert.equal(result.cells['4:0'].value, 'row 8'); assert.deepEqual(source, before)
  assert.deepEqual(structure.sheetAxisStructureEdits(sheet, selected, { axis: 'row', kind: 'delete' }, { includeFiltered: true }, [3, 5]),
    [{ axis: 'row', kind: 'delete', start: 2, count: 6 }])
  assert.throws(() => structure.sheetAxisStructureEdits(sheet, [range(0, 11)], { axis: 'row', kind: 'delete' }, {}), /至少/)
})

test('row move and copy retain dimension ownership and copy relative formulas without changing their absolute references', () => {
  const source = workbook(), sheet = source.sheets[0]
  sheet.rowHiddenSpans = [{ start: 1, end: 2, value: true }]; sheet.rowHeights = { 1: 40 }
  sheet.cells['1:1'] = { ...sdk.createCell(1, 1, null), formula: '=A1+$A$1' }
  source.sheets[1].cells['0:0'] = { ...sdk.createCell(0, 0, null), formula: "='Target'!B2" }
  const before = copy(source), moved = structure.planWorkbookAxisStructureCommand(source, sheet.id,
    [{ axis: 'row', kind: 'move', start: 1, count: 2, destination: 5 }]).workbook
  assert.deepEqual(moved.sheets[0].rowHiddenSpans, [{ start: 3, end: 4, value: true }]); assert.equal(moved.sheets[0].rowHeights[3], 40)
  assert.equal(moved.sheets[1].cells['0:0'].formula, "='Target'!B4"); assert.equal(moved.sheets[0].rowCount, 12)
  const copied = structure.planWorkbookAxisStructureCommand(source, sheet.id,
    [{ axis: 'row', kind: 'copy', start: 1, count: 2, destination: 5 }]).workbook
  assert.deepEqual(copied.sheets[0].rowHiddenSpans, [{ start: 1, end: 2, value: true }, { start: 5, end: 6, value: true }])
  assert.equal(copied.sheets[0].cells['5:1'].formula, '=A5+$A$1'); assert.equal(copied.sheets[0].rowHeights[5], 40)
  assert.equal(copied.sheets[0].rowCount, 14); assert.deepEqual(source, before)
})

test('partial merge movement, Office overflow and protected axes reject without mutating the source', () => {
  const source = workbook(), sheet = source.sheets[0]
  sheet.mergedRanges = [range(1, 2, 0, 1)]
  const before = copy(source)
  assert.throws(() => structure.planWorkbookAxisStructureCommand(source, sheet.id, [{ axis: 'row', kind: 'move', start: 1, count: 1, destination: 5 }]), /部分合并/)
  assert.throws(() => structure.planWorkbookAxisStructureCommand(source, sheet.id, [{ axis: 'row', kind: 'move', start: 1, count: 2, destination: 2 }]), /内部/)
  assert.deepEqual(source, before)
  sheet.rowCount = 1_048_576; sheet.cells['1048575:0'] = sdk.createCell(1_048_575, 0, 'must not disappear')
  assert.throws(() => structure.planWorkbookAxisStructureCommand(source, sheet.id, [{ axis: 'row', kind: 'insert', start: 0, count: 1 }]), /移出/)
  sheet.worksheetProtection = { id: 'locked', permissionId: 'locked-permission', unitType: 'worksheet', ownerUserId: 'another-user', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
  assert.throws(() => structure.planWorkbookAxisStructureCommand(source, sheet.id, [{ axis: 'row', kind: 'delete', start: 0, count: 1 }]), /保护|权限|不可/)
})

test('Facade visibility and structure aliases commit once to inactive sheets and isolate observer payloads and failures', t => {
  const { studio, api, book, first, second } = runtime(t), received = [], failure = new Error('axis observer failed')
  const originalSelection = studio.getSelectionState()
  api.addEvent(api.Event.SheetVisibilityChanged, event => { event.ranges[0].startRow = 10; throw failure })
  api.addEvent(api.Event.SheetVisibilityChanged, event => received.push(['visible', event.sheetId, event.ranges[0].startRow]))
  api.addEvent(api.Event.SheetStructureChanged, event => { event.edits[0].start = 999; throw failure })
  api.addEvent(api.Event.SheetStructureChanged, event => received.push(['structure', event.edits[0].start]))
  api.addEvent(api.Event.SheetValueChanged, () => { throw failure })
  api.addEvent(api.Event.SheetValueChanged, event => received.push(['value', event.sheetId]))
  second.hideRows(2, 3)
  assert.deepEqual(studio.getWorkbookModel(book.getId()).sheets[1].rowHiddenSpans, [{ start: 2, end: 4, value: true }])
  assert.equal(history(studio).undo, 1); assert.deepEqual(studio.getSelectionState(), originalSelection)
  first.insertRowsBefore(1, 1)
  assert.equal(first.getMaxRows(), 13); assert.equal(history(studio).undo, 2)
  assert.deepEqual(received, [['visible', second.getSheetId(), 2], ['structure', 1], ['value', first.getSheetId()]])
  assert.equal(studio.getEventListenerErrors().filter(event => event.error === failure).length, 3)
  book.undo(); assert.equal(first.getMaxRows(), 12)
  book.undo(); assert.deepEqual(studio.getWorkbookModel(book.getId()).sheets[1].rowHiddenSpans ?? [], [])
  book.redo(); book.redo(); assert.equal(first.getMaxRows(), 13)
  const saved = book.save(), reopened = api.createWorkbook({ ...saved, id: 'axis-edit-reopen' })
  assert.equal(reopened.getSheets()[0].getMaxRows(), 13)
  assert.deepEqual(reopened.save().sheets[1].rowHiddenSpans, [{ start: 2, end: 4, value: true }])
})

test('structure history preserves complete multiple selections and restores the planned row selection on redo', t => {
  const { studio, book, first } = runtime(t)
  first.setActiveRangeList([first.getRange('A1:C2'), first.getRange('F6:G7')], { activeRangeIndex: 1, primary: { row: 6, column: 5 } })
  const before = studio.getSelectionState()
  first.insertRowsBefore(1, 2)
  const after = studio.getSelectionState()
  assert.equal(after.ranges.length, 1); assert.equal(after.ranges[0].rangeType, 1)
  assert.equal(after.ranges[0].startRow, 1); assert.equal(after.ranges[0].endRow, 2)
  book.undo(); assert.deepEqual(studio.getSelectionState(), before)
  book.redo(); assert.deepEqual(studio.getSelectionState(), after); assert.equal(history(studio).undo, 1)
})

test('axis view preparation refusal and cancellation preserve model identity, full selection, Tab origin, revision and history', t => {
  const { studio, api, book, first } = runtime(t)
  first.getRange('B2').activate()
  api.executeCommandSync(sdk.MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND, { direction: api.Enum.Direction.RIGHT, keycode: 9 })
  const model = studio.getWorkbookModel(book.getId()), before = copy(model), selection = studio.getSelectionState(), depth = history(studio)
  const origin = studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()), commit = studio.commitWorkbook.bind(studio), events = []
  api.addEvent(api.Event.SheetVisibilityChanged, event => events.push(event)); api.addEvent(api.Event.SheetStructureChanged, event => events.push(event))
  studio.commitWorkbook = (id, label, update, options) => {
    assert.equal(typeof options.prepareView, 'function')
    return commit(id, label, update, { ...options, prepareView: candidate => {
      options.prepareView(candidate); candidate.sheets[0].name = 'detached only'; throw new Error('axis preparation refused')
    } })
  }
  for (const command of uiCommands()) assert.throws(() => api.executeCommandSync(command.id,
    { ...command.params, unitId: book.getId(), sheetId: first.getSheetId(), ranges: [range(1)] }), /axis preparation refused/)
  assert.equal(studio.getWorkbookModel(book.getId()), model); assert.deepEqual(copy(model), before)
  assert.deepEqual(studio.getSelectionState(), selection); assert.deepEqual(history(studio), depth)
  assert.deepEqual(studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()), origin); assert.deepEqual(events, [])
  studio.commitWorkbook = commit
  const veto = api.addEvent(api.Event.BeforeCommandExecute, event => { if (event.id === 'sheet.command.set-rows-hidden') event.cancel = true })
  assert.throws(() => first.hideRows(1), /canceled/); veto.dispose()
  assert.deepEqual(copy(model), before); assert.deepEqual(history(studio), depth)
  first.hideRows(1)
  assert.throws(() => api.executeCommandSync('sheet.command.insert-row-before', { unitId: book.getId(), sheetId: first.getSheetId(), ranges: [range(1)], expectedRevision: depth.revision }), /已更改/)
})

test('a subscriber workbook switch prevents every axis controller from applying a stale view or selection', t => {
  for (const command of uiCommands()) {
    const { studio, api, book } = runtime(t), target = host(studio)
    const neighbor = api.createWorkbook({ ...workbook(), id: 'axis-edit-neighbor' }, { makeCurrent: false })
    const eventName = api.Event[command.event]
    api.addEvent(eventName, () => neighbor.getActiveSheet().getRange('D4').activate())
    assert.equal(command.create(target.context)(command.id), true)
    assert.equal(studio.activeWorkbookId, neighbor.getId()); assert.deepEqual(studio.getSelectionState().primary, { row: 3, column: 3 })
    assert.equal(target.calls.applied, 0); assert.equal(target.calls.selections, 0); assert.equal(target.calls.changed, 0)
    assert.equal(history(studio, book.getId()).undo, 1); assert.equal(history(studio, neighbor.getId()).undo, 0)
  }
})

test('same-sheet observer multiple selections are projected once and a nested content edit stops stale host notifications', t => {
  const { studio, api, first } = runtime(t), target = host(studio)
  api.addEvent(api.Event.SheetStructureChanged, () => first.setActiveRangeList([first.getRange('E5:F6'), first.getRange('B8')],
    { activeRangeIndex: 0, primary: { row: 4, column: 4 } }))
  assert.equal(structureUi.useSheetAxisStructure(target.context).runAxisStructureCommand('sheet.command.insert-row-before'), true)
  assert.deepEqual(target.calls.projection, studio.getSelectionState()); assert.equal(target.calls.projection.ranges.length, 2)
  assert.equal(target.calls.selections, 0); assert.equal(target.calls.changed, 1); assert.equal(target.calls.options.syncSelection, false)
  assert.equal(target.calls.options.current(), true)
  const nested = host(studio), apply = nested.context.applyView
  nested.context.applyView = next => { apply(next); first.getRange('H12').setInputText('nested content') }
  assert.equal(visibilityUi.useSheetAxisVisibility(nested.context).runAxisVisibilityCommand('sheet.command.set-rows-hidden', { index: 1, count: 1 }), true)
  assert.equal(first.getRange('H12').getValue(), 'nested content'); assert.equal(nested.calls.changed, 0)
  assert.equal(target.calls.options.current(), false)
})

test('standalone preparation uses detached baselines and refuses replaced or edited sources without publishing a partial command', () => {
  for (const command of uiCommands()) {
    const target = host(undefined), before = copy(target.view())
    target.context.prepareView = (_next, previous) => { previous.name = 'changed detached baseline'; throw new Error('standalone preparation refused') }
    assert.throws(() => command.create(target.context)(command.id), /standalone preparation refused/)
    assert.deepEqual(target.view(), before); assert.equal(target.calls.applied, 0); assert.equal(target.calls.changed, 0)
    target.context.prepareView = next => { target.replace({ ...copy(target.view()), name: 'replacement view' }); return next }
    assert.equal(command.create(target.context)(command.id), false); assert.equal(target.view().name, 'replacement view')
    assert.equal(target.calls.applied, 0); assert.equal(target.calls.selections, 0)
    target.context.prepareView = next => { target.view().name = 'in-place newer view'; return next }
    assert.equal(command.create(target.context)(command.id), false); assert.equal(target.view().name, 'in-place newer view')
    assert.equal(target.calls.applied, 0); assert.equal(target.calls.changed, 0)
    target.context.prepareView = next => next
    assert.equal(command.create(target.context)(command.id), true); assert.equal(target.view().id, undefined)
    assert.equal(target.calls.applied, 1); assert.equal(target.calls.selections, command.selects ? 1 : 0); assert.equal(target.calls.changed, 1)
  }
})

test('all standalone axis controllers stop old selection and publication after applyView performs a newer in-place edit', () => {
  for (const command of uiCommands()) {
    const target = host(undefined), apply = target.context.applyView
    target.context.applyView = next => { apply(next); target.view().name = 'newer edit during projection' }
    assert.equal(command.create(target.context)(command.id), true)
    assert.equal(target.calls.applied, 1); assert.equal(target.calls.selections, 0); assert.equal(target.calls.changed, 0)
    assert.equal(target.view().name, 'newer edit during projection')
  }
})

test('standalone current guards detect in-place content changes after successful axis publication', () => {
  for (const command of uiCommands()) {
    const target = host(undefined)
    assert.equal(command.create(target.context)(command.id), true)
    assert.equal(target.calls.options.current(), true)
    target.view().sheets[0].name = 'later same-view content edit'
    assert.equal(target.calls.options.current(), false)
  }
})

test('dimension observers receive independent ranges and cannot invalidate a committed height transaction', t => {
  const { studio, api, book, first } = runtime(t), failures = new Error('dimension observer failed'), seen = []
  api.addEvent(api.Event.SheetDimensionsChanged, event => { event.ranges[0].startRow = 10; throw failures })
  api.addEvent(api.Event.SheetDimensionsChanged, event => seen.push([event.axis, event.automatic, event.ranges[0].startRow]))
  api.executeCommandSync('sheet.command.set-row-height', { unitId: book.getId(), sheetId: first.getSheetId(), ranges: [range(1)], value: 38 })
  assert.equal(dimensions.sheetRowHeight(studio.getWorkbookModel(book.getId()).sheets[0], 1), 38)
  assert.deepEqual(seen, [['row', false, 1]])
  assert.equal(history(studio).undo, 1)
  assert.equal(studio.getEventListenerErrors().filter(event => event.error === failures).length, 1)
  book.undo(); assert.equal(dimensions.sheetRowHeight(studio.getWorkbookModel(book.getId()).sheets[0], 1), 26)
  book.redo(); assert.equal(dimensions.sheetRowHeight(studio.getWorkbookModel(book.getId()).sheets[0], 1), 38)
})

test('dimension projection retains subscriber multi-selection and refuses a stale size dialog', t => {
  const { studio, api, book, first } = runtime(t), target = host(studio), ui = dimensionUi.useSheetAxisDimensions(target.context)
  api.addEvent(api.Event.SheetDimensionsChanged, () => first.setActiveRangeList([first.getRange('C4:D5'), first.getRange('F7')],
    { activeRangeIndex: 1, primary: { row: 6, column: 5 } }))
  assert.equal(ui.runAxisDimensionCommand('sheet.command.set-row-height', { value: 38 }), true)
  assert.deepEqual(target.calls.projection, studio.getSelectionState())
  assert.equal(target.calls.selections, 0); assert.equal(target.calls.options.syncSelection, false)
  ui.openAxisSizeDialog('row')
  first.getRange('H12').setInputText('newer content')
  const depth = history(studio)
  ui.confirmAxisSize(50)
  assert.match(ui.axisSizeError.value, /已更改/); assert.ok(ui.pending.value)
  assert.deepEqual(history(studio), depth)
  assert.equal(dimensions.sheetRowHeight(studio.getWorkbookModel(book.getId()).sheets[0], 1), 38)
})

test('a studio host with no workbook ID uses standalone identity and still guards pending size edits', t => {
  const { studio } = runtime(t), target = host(undefined), context = { ...target.context, studio: () => studio }
  const ui = dimensionUi.useSheetAxisDimensions(context)
  ui.openAxisSizeDialog('column')
  target.view().name = 'newer detached host'
  ui.confirmAxisSize(80)
  assert.match(ui.axisSizeError.value, /已更改/); assert.equal(target.calls.applied, 0)
  ui.cancelAxisSize()
  assert.equal(ui.runAxisDimensionCommand('sheet.command.set-row-height', { value: 40 }), true)
  assert.equal(target.view().id, undefined); assert.equal(target.calls.changed, 1)
})
