import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetContextMenuItems, sheetContextMenuFeatureSearchContext, sheetContextMenuSelectionHit,
  translateSheetContextMenuItems } from '../src/components/spreadsheet/sheet-context-menu.ts'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { sheetClipboardCapabilities } from '../src/features/sheet-workbench/sheet-clipboard-capabilities.ts'
import { collectFeatureSearchResults } from '../src/platform/ribbon/feature-search.ts'

const options = { sheetId: 'sheet-a', canDeleteRow: true, canDeleteColumn: true, frozen: true, hasHiddenRows: true, hasHiddenColumns: true, canConvertText: true }
const flattenItems = (items) => items.flatMap(item => [item, ...flattenItems(item.children ?? [])])
const commands = (hit, state = options) => flattenItems(createSheetContextMenuItems(hit, state)).filter(item => item.commandId).map(item => item.commandId)

test('cell context menu exposes clipboard, editing, row/column, freeze and protection commands', () => {
  const items = createSheetContextMenuItems({ kind: 'cell', row: 2, column: 1, key: 'B3' }, options)
  const actualCommands = commands({ kind: 'cell', row: 2, column: 1, key: 'B3' })
  for (const command of [
    'sheet.command.copy', 'sheet.command.cut', 'sheet.command.paste', 'sheet.command.copy-formulas', 'sheet.command.paste-values',
    'sheet.command.paste-formats', 'sheet.command.paste-column-widths', 'sheet.command.paste-except-borders', 'sheet.command.paste-formulas',
    'sheet.command.clear-content', 'sheet.command.clear-format', 'sheet.command.clear-all', 'sheet.command.text-to-number',
    'sheet.command.insert-row-before', 'sheet.command.insert-col-before', 'sheet.command.insert-range-move-right-confirm', 'sheet.command.insert-range-move-down-confirm',
    'sheet.command.delete-column', 'sheet.command.delete-row', 'sheet.command.delete-range-move-left-confirm', 'sheet.command.delete-range-move-up-confirm', 'sheet.command.freeze-selection',
    'sheet.command.freeze-first-row', 'sheet.command.freeze-first-column', 'sheet.command.unfreeze',
    'sheet.operation.open-range-protection', 'sheet.operation.open-range-protection',
  ]) assert.ok(actualCommands.includes(command), command)
  assert.deepEqual(items.map(item => item.id), ['copy', 'cut', 'paste', 'copy-special', 'paste-special', 'clear-selection', 'borders', 'text-to-number', 'insert-cells', 'delete-cells', 'freeze', 'protect-range'])
  assert.equal(items.find(item => item.id === 'insert-cells').separatorBefore, true)
  assert.deepEqual(items.find(item => item.id === 'insert-cells').children.map(item => item.commandId), [
    'sheet.command.insert-row-before', 'sheet.command.insert-col-before', 'sheet.command.insert-range-move-right-confirm', 'sheet.command.insert-range-move-down-confirm',
  ])
  assert.deepEqual(items.find(item => item.id === 'delete-cells').children.map(item => item.commandId), [
    'sheet.command.delete-column', 'sheet.command.delete-row', 'sheet.command.delete-range-move-left-confirm', 'sheet.command.delete-range-move-up-confirm',
  ])
  assert.equal(items.find(item => item.id === 'copy').quick, true)
  assert.deepEqual(items.find(item => item.id === 'paste-special')?.children?.map(item => item.commandId), [
    'sheet.command.paste-values', 'sheet.command.paste-formats', 'sheet.command.paste-column-widths', 'sheet.command.paste-except-borders', 'sheet.command.paste-formulas',
  ])
})

test('row and column headers get axis-specific operations and preserve freeze/protection actions', () => {
  const rowItems = createSheetContextMenuItems({ kind: 'row', row: 4 }, options)
  const columnItems = createSheetContextMenuItems({ kind: 'column', column: 3 }, options)
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.command.insert-row-before'))
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.command.set-row-height'))
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.command.delete-row'))
  assert.ok(!commands({ kind: 'row', row: 4 }).includes('sheet.command.insert-col-before'))
  assert.ok(!commands({ kind: 'row', row: 4 }).includes('sheet.command.delete-column'))
  assert.ok(commands({ kind: 'column', column: 3 }).includes('sheet.command.insert-col-before'))
  assert.ok(commands({ kind: 'column', column: 3 }).includes('sheet.command.set-worksheet-col-width'))
  assert.ok(commands({ kind: 'column', column: 3 }).includes('sheet.command.delete-column'))
  assert.ok(!commands({ kind: 'column', column: 3 }).includes('sheet.command.insert-row-before'))
  assert.ok(!commands({ kind: 'column', column: 3 }).includes('sheet.command.delete-row'))
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.operation.open-range-protection'))
  const rowTarget = { sheetId: 'sheet-a', ranges: [{ startRow: 4, endRow: 4, startCol: 0, endCol: 16_383 }] }
  const columnTarget = { sheetId: 'sheet-a', ranges: [{ startRow: 0, endRow: 1_048_575, startCol: 3, endCol: 3 }] }
  for (const id of ['insert-row-above', 'insert-row-below', 'delete-row']) assert.deepEqual(rowItems.find(item => item.id === id)?.params, rowTarget)
  for (const id of ['insert-column-left', 'insert-column-right', 'delete-column']) assert.deepEqual(columnItems.find(item => item.id === id)?.params, columnTarget)
  assert.deepEqual(rowItems.find(item => item.id === 'row-height')?.params, { sheetId: 'sheet-a', index: 4 })
  assert.deepEqual(columnItems.find(item => item.id === 'column-width')?.params, { sheetId: 'sheet-a', index: 3 })
  assert.ok(rowItems.some(item => item.commandId === 'sheet.command.hide-row-confirm'))
  assert.ok(columnItems.some(item => item.commandId === 'sheet.command.auto-fit-column'))
})

test('delete, unhide and unfreeze controls reflect model constraints', () => {
  const items = createSheetContextMenuItems({ kind: 'cell', row: 0, column: 0, key: 'A1' }, { ...options, canDeleteRow: false, canDeleteColumn: false, frozen: false, hasHiddenRows: false, hasHiddenColumns: false })
  const flatItems = flattenItems(items)
  assert.equal(flatItems.find(item => item.id === 'delete-row').disabled, true)
  assert.equal(flatItems.find(item => item.id === 'delete-column').disabled, true)
  assert.equal(flatItems.find(item => item.id === 'unfreeze').disabled, true)
  const rowItems = createSheetContextMenuItems({ kind: 'row', row: 0 }, { ...options, hasHiddenRows: false })
  const columnItems = createSheetContextMenuItems({ kind: 'column', column: 0 }, { ...options, hasHiddenColumns: false })
  assert.equal(flattenItems(rowItems).find(item => item.id === 'show-hidden-rows').disabled, true)
  assert.equal(flattenItems(columnItems).find(item => item.id === 'show-hidden-columns').disabled, true)
})

test('cell shift commands are disabled when the shifted tail intersects protected ranges', () => {
  const items = flattenItems(createSheetContextMenuItems({ kind: 'cell', row: 2, column: 1, key: 'B3' }, {
    ...options, canShiftHorizontally: false, canShiftVertically: true,
  }))
  assert.equal(items.find(item => item.id === 'insert-range-right').disabled, true)
  assert.equal(items.find(item => item.id === 'delete-range-left').disabled, true)
  assert.equal(items.find(item => item.id === 'insert-range-down').disabled, false)
  assert.equal(items.find(item => item.id === 'delete-range-up').disabled, false)
})

test('whole-sheet context menu does not offer destructive row or column actions', () => {
  const items = createSheetContextMenuItems({ kind: 'all' }, options)
  assert.ok(commands({ kind: 'all' }).includes('sheet.command.copy'))
  assert.ok(commands({ kind: 'all' }).includes('sheet.command.unfreeze'))
  assert.ok(!flattenItems(items).some(item => /delete-(row|column)/.test(item.id)))
})

test('menu metadata preserves the fixed-version icon IDs and quick-action order', () => {
  const items = createSheetContextMenuItems({ kind: 'cell', row: 2, column: 1, key: 'B3' }, options)
  assert.deepEqual(items.filter(item => item.quick).map(item => [item.id, item.icon]), [
    ['copy', 'CopyDoubleIcon'], ['cut', 'CutIcon'], ['paste', 'PasteSpecialDoubleIcon'],
  ])
  assert.deepEqual(items.find(item => item.id === 'freeze')?.children?.map(item => item.icon), [
    'FreezeToSelectedIcon', 'FreezeRowIcon', 'FreezeColumnIcon', 'CancelFreezeIcon',
  ])
})

test('resize boundaries and unsupported hits do not open a menu', () => {
  assert.deepEqual(createSheetContextMenuItems({ kind: 'resize-row', row: 3, size: 24 }, options), [])
  assert.deepEqual(createSheetContextMenuItems({ kind: 'resize-column', column: 2, size: 96 }, options), [])
})

// Deferred drafts: feature-search selection, capabilities and command parameters have not been executed.
test('feature-search context uses the primary cell and preserves explicit selection types', () => {
  const primary = { kind: 'cell', row: 2, column: 1, key: 'B3' }
  const dimensions = { rowCount: 20, columnCount: 8 }
  const rows = { startRow: 1, endRow: 3, startColumn: 0, endColumn: 7 }
  const columns = { startRow: 0, endRow: 19, startColumn: 1, endColumn: 3 }
  assert.equal(sheetContextMenuSelectionHit(primary, undefined, dimensions), primary)
  assert.deepEqual(sheetContextMenuSelectionHit(primary, rows, dimensions), { kind: 'row', row: 2 })
  assert.deepEqual(sheetContextMenuSelectionHit(primary, columns, dimensions), { kind: 'column', column: 1 })
  assert.deepEqual(sheetContextMenuSelectionHit(primary, { ...rows, startRow: 0, endRow: 19 }, dimensions), { kind: 'all' })
  assert.equal(sheetContextMenuSelectionHit(primary, { ...rows, rangeType: 0 }, dimensions), primary)
  assert.deepEqual(sheetContextMenuSelectionHit(primary, { ...columns, rangeType: 1 }, dimensions), { kind: 'row', row: 2 })
  assert.deepEqual(sheetContextMenuSelectionHit(primary, { ...rows, rangeType: 2 }, dimensions), { kind: 'column', column: 1 })
})

test('feature-search row actions share translated counts and complete selected-axis parameters', () => {
  const state = { ...options, rowCount: 20, columnCount: 8,
    selectedRanges: [{ startRow: 1, endRow: 3, startColumn: 0, endColumn: 7, rangeType: 1 }] }
  const translate = (key, fallback) => key === 'sheet.context-menu.insert-row-above' ? 'Insert {count} rows above' : fallback
  const items = flattenItems(sheetContextMenuFeatureSearchContext({ kind: 'row', row: 2 }, state, translate)[0].items)
  const action = items.find(item => item.id === 'insert-row-above')
  assert.equal(action.title, 'Insert 3 rows above')
  assert.equal(action.commandId, 'sheet.command.insert-row-before')
  assert.deepEqual(action.params, { sheetId: 'sheet-a', ranges: [{ startRow: 1, endRow: 3, startCol: 0, endCol: 7 }] })
  assert.equal(items.some(item => item.id.endsWith('-separator')), false)
  const menu = translateSheetContextMenuItems(createSheetContextMenuItems({ kind: 'row', row: 2 }, state), translate)
  assert.equal(menu.find(item => item.id === 'insert-row-above').label, action.title)
  assert.equal(menu.find(item => item.id === 'insert-row-above-separator').type, 'divider')
})

test('feature-search filters hidden clipboard groups and disabled axis or formatting commands', () => {
  const workbook = createBlankWorkbook()
  workbook.sheets[0].id = 'sheet-a'
  const clipboardCapabilities = sheetClipboardCapabilities({ workbook, sheetId: 'sheet-a', userId: 'local-user',
    selection: { ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }] }, exclusiveRange: true, systemClipboardAvailable: false })
  const groups = sheetContextMenuFeatureSearchContext({ kind: 'row', row: 0 }, { ...options, clipboardCapabilities,
    canFormat: false, canStructureRows: false, canVisibilityRows: false }, (_key, fallback) => fallback)
  const items = flattenItems(groups[0].items)
  assert.equal(items.some(item => item.id === 'copy-special' || item.id === 'paste-special'), false)
  assert.equal(items.find(item => item.id === 'insert-row-above').disabled, true)
  const results = collectFeatureSearchResults([], '功能区', { title: '上下文', groups })
  assert.ok(results.some(item => item.commandId === 'sheet.command.copy'))
  for (const id of ['sheet.command.paste', 'sheet.command.insert-row-before', 'sheet.command.hide-row-confirm', 'sheet.command.set-border-position']) {
    assert.equal(results.some(item => item.commandId === id), false, id)
  }
  assert.equal(results.some(item => !item.commandId), false)
})

test('feature-search keeps image, hyperlink and sparkline ownership parameters without changing the context', () => {
  const state = { ...options, hasHyperlink: true, hasCellImages: true, canSaveCellImages: false,
    contextImage: { id: 'cell-picture', row: 1, column: 1 },
    sparkline: { primary: 'B3', ranges: [{ sheetId: 'sheet-a', startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 }],
      canEdit: false, canDelete: true, expectedRevision: 7 } }
  const before = JSON.stringify(state)
  const groups = sheetContextMenuFeatureSearchContext({ kind: 'cell', row: 2, column: 1, key: 'B3' }, state, (_key, fallback) => fallback)
  const items = flattenItems(groups[0].items)
  assert.deepEqual(items.find(item => item.id === 'open-hyperlink').params, { key: 'B3' })
  assert.deepEqual(items.find(item => item.id === 'edit-cell-image').params, { sheetId: 'sheet-a', imageId: 'cell-picture', row: 1, column: 1 })
  assert.equal(items.find(item => item.id === 'save-cell-images').disabled, true)
  assert.equal(items.find(item => item.id === 'edit-sparkline').disabled, true)
  assert.match(items.find(item => item.id === 'edit-sparkline').disabledReason, /权限/)
  assert.deepEqual(items.find(item => item.id === 'remove-sparkline').params, { sheetId: 'sheet-a', ranges: state.sparkline.ranges,
    expectedRevision: 7, isSingle: true })
  assert.equal(JSON.stringify(state), before)
})

test('feature-search outline commands retain revision, history generation, user and whole-axis selection', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.id = 'sheet-a'; sheet.rowCount = 20; sheet.columnCount = 8
  const ranges = [{ startRow: 1, endRow: 3, startCol: 0, endCol: 7, rangeType: 1 }]
  const groups = sheetContextMenuFeatureSearchContext({ kind: 'row', row: 2 }, { ...options,
    rowCount: 20, columnCount: 8, selectedRanges: [{ startRow: 1, endRow: 3, startColumn: 0, endColumn: 7, rangeType: 1 }],
    outline: { context: { unitId: 'context-book', sheetId: sheet.id, sheet, userId: 'local-user', selections: ranges },
      owner: { expectedRevision: 7, expectedHistoryGeneration: 2, expectedUserId: 'local-user' } } }, (_key, fallback) => fallback)
  const action = flattenItems(groups[0].items).find(item => item.commandId === 'sheet.command.outline-selected-rows')
  assert.equal(action.disabled, false)
  assert.deepEqual(action.params, { value: 'row', unitId: 'context-book', subUnitId: 'sheet-a', ranges, host: 'workbench',
    expectedRevision: 7, expectedHistoryGeneration: 2, expectedUserId: 'local-user' })
})
