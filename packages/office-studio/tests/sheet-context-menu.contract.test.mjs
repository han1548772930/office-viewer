import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetContextMenuItems } from '../src/components/spreadsheet/sheet-context-menu.ts'

const options = { sheetId: 'sheet-a', canDeleteRow: true, canDeleteColumn: true, frozen: true, hasHiddenRows: true, hasHiddenColumns: true, canConvertText: true }
const flattenItems = (items) => items.flatMap(item => [item, ...flattenItems(item.children ?? [])])
const commands = (hit, state = options) => flattenItems(createSheetContextMenuItems(hit, state)).filter(item => item.commandId).map(item => item.commandId)

test('cell context menu exposes clipboard, editing, row/column, freeze and protection commands', () => {
  const items = createSheetContextMenuItems({ kind: 'cell', row: 2, column: 1, key: 'B3' }, options)
  assert.deepEqual(commands({ kind: 'cell', row: 2, column: 1, key: 'B3' }), [
    'sheet.command.copy', 'sheet.command.cut', 'sheet.command.paste', 'sheet.command.copy-formulas', 'sheet.command.paste-values',
    'sheet.command.paste-formats', 'sheet.command.paste-column-widths', 'sheet.command.paste-except-borders', 'sheet.command.paste-formulas',
    'sheet.command.clear-content', 'sheet.command.clear-format', 'sheet.command.clear-all', 'sheet.command.text-to-number',
    'sheet.command.insert-range-move-right', 'sheet.command.insert-range-move-down', 'sheet.command.insert-row', 'sheet.command.insert-column',
    'sheet.command.delete-range-move-left', 'sheet.command.delete-range-move-up', 'sheet.command.delete-column', 'sheet.command.delete-row', 'sheet.command.freeze-selection',
    'sheet.command.freeze-first-row', 'sheet.command.freeze-first-column', 'sheet.command.unfreeze',
    'sheet.operation.open-range-protection', 'sheet.operation.open-range-protection',
  ])
  assert.deepEqual(items.map(item => item.id), ['copy', 'cut', 'paste', 'copy-special', 'paste-special', 'clear-selection', 'text-to-number', 'insert-cells', 'delete-cells', 'freeze', 'protect-range'])
  assert.equal(items.find(item => item.id === 'insert-cells').separatorBefore, true)
  assert.deepEqual(items.find(item => item.id === 'insert-cells').children.slice(0, 2).map(item => item.commandId), [
    'sheet.command.insert-range-move-right', 'sheet.command.insert-range-move-down',
  ])
  assert.deepEqual(items.find(item => item.id === 'delete-cells').children.slice(0, 2).map(item => item.commandId), [
    'sheet.command.delete-range-move-left', 'sheet.command.delete-range-move-up',
  ])
  assert.equal(items.find(item => item.id === 'copy').quick, true)
  assert.deepEqual(items.find(item => item.id === 'paste-special')?.children?.map(item => item.commandId), [
    'sheet.command.paste-values', 'sheet.command.paste-formats', 'sheet.command.paste-column-widths', 'sheet.command.paste-except-borders', 'sheet.command.paste-formulas',
  ])
})

test('row and column headers get axis-specific operations and preserve freeze/protection actions', () => {
  const rowItems = createSheetContextMenuItems({ kind: 'row', row: 4 }, options)
  const columnItems = createSheetContextMenuItems({ kind: 'column', column: 3 }, options)
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.command.insert-row'))
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.command.set-row-height'))
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.command.delete-row'))
  assert.ok(!commands({ kind: 'row', row: 4 }).includes('sheet.command.insert-column'))
  assert.ok(!commands({ kind: 'row', row: 4 }).includes('sheet.command.delete-column'))
  assert.ok(commands({ kind: 'column', column: 3 }).includes('sheet.command.insert-column'))
  assert.ok(commands({ kind: 'column', column: 3 }).includes('sheet.command.set-worksheet-col-width'))
  assert.ok(commands({ kind: 'column', column: 3 }).includes('sheet.command.delete-column'))
  assert.ok(!commands({ kind: 'column', column: 3 }).includes('sheet.command.insert-row'))
  assert.ok(!commands({ kind: 'column', column: 3 }).includes('sheet.command.delete-row'))
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.operation.open-range-protection'))
  assert.deepEqual(rowItems.find(item => item.id === 'insert-row-above')?.params, { sheetId: 'sheet-a', index: 4 })
  assert.deepEqual(rowItems.find(item => item.id === 'insert-row-below')?.params, { sheetId: 'sheet-a', index: 5 })
  assert.deepEqual(rowItems.find(item => item.id === 'delete-row')?.params, { sheetId: 'sheet-a', index: 4 })
  assert.deepEqual(columnItems.find(item => item.id === 'insert-column-left')?.params, { sheetId: 'sheet-a', index: 3 })
  assert.deepEqual(columnItems.find(item => item.id === 'insert-column-right')?.params, { sheetId: 'sheet-a', index: 4 })
  assert.deepEqual(columnItems.find(item => item.id === 'delete-column')?.params, { sheetId: 'sheet-a', index: 3 })
  assert.deepEqual(rowItems.find(item => item.id === 'row-height')?.params, { sheetId: 'sheet-a', index: 4 })
  assert.deepEqual(columnItems.find(item => item.id === 'column-width')?.params, { sheetId: 'sheet-a', index: 3 })
  assert.ok(rowItems.some(item => item.commandId === 'sheet.command.hide-row'))
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
