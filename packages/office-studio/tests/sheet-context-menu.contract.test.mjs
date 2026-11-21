import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetContextMenuItems } from '../src/components/spreadsheet/sheet-context-menu.ts'

const options = { canDeleteRow: true, canDeleteColumn: true, frozen: true }
const commands = (hit, state = options) => createSheetContextMenuItems(hit, state).map(item => item.commandId)

test('cell context menu exposes clipboard, editing, row/column, freeze and protection commands', () => {
  const items = createSheetContextMenuItems({ kind: 'cell', row: 2, column: 1, key: 'B3' }, options)
  assert.deepEqual(items.map(item => item.commandId), [
    'sheet.command.cut', 'sheet.command.copy', 'sheet.command.paste', 'sheet.command.paste-values',
    'sheet.command.paste-formats', 'sheet.command.paste-column-widths', 'sheet.command.paste-except-borders',
    'sheet.command.clear-content', 'sheet.command.insert-row', 'sheet.command.insert-column',
    'sheet.command.delete-row', 'sheet.command.delete-column', 'sheet.command.freeze-selection',
    'sheet.command.freeze-first-row', 'sheet.command.freeze-first-column', 'sheet.command.unfreeze',
    'sheet.operation.open-range-protection',
  ])
  assert.equal(items.find(item => item.id === 'clear-content').separatorBefore, true)
})

test('row and column headers get axis-specific operations and preserve freeze/protection actions', () => {
  const rowItems = createSheetContextMenuItems({ kind: 'row', row: 4 }, options)
  const columnItems = createSheetContextMenuItems({ kind: 'column', column: 3 }, options)
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.command.insert-row'))
  assert.ok(commands({ kind: 'row', row: 4 }).includes('sheet.command.delete-row'))
  assert.ok(!rowItems.some(item => ['sheet.command.insert-column', 'sheet.command.delete-column'].includes(item.commandId)))
  assert.ok(commands({ kind: 'column', column: 3 }).includes('sheet.command.insert-column'))
  assert.ok(commands({ kind: 'column', column: 3 }).includes('sheet.command.delete-column'))
  assert.ok(!columnItems.some(item => ['sheet.command.insert-row', 'sheet.command.delete-row'].includes(item.commandId)))
  assert.ok(rowItems.some(item => item.commandId === 'sheet.operation.open-range-protection'))
})

test('delete and unfreeze controls reflect model constraints', () => {
  const items = createSheetContextMenuItems({ kind: 'cell', row: 0, column: 0, key: 'A1' }, { canDeleteRow: false, canDeleteColumn: false, frozen: false })
  assert.equal(items.find(item => item.id === 'delete-row').disabled, true)
  assert.equal(items.find(item => item.id === 'delete-column').disabled, true)
  assert.equal(items.find(item => item.id === 'unfreeze').disabled, true)
})

test('whole-sheet context menu does not offer destructive row or column actions', () => {
  const items = createSheetContextMenuItems({ kind: 'all' }, options)
  assert.ok(items.some(item => item.commandId === 'sheet.command.copy'))
  assert.ok(items.some(item => item.commandId === 'sheet.command.unfreeze'))
  assert.ok(!items.some(item => /delete-(row|column)/.test(item.id)))
})

test('resize boundaries and unsupported hits do not open a menu', () => {
  assert.deepEqual(createSheetContextMenuItems({ kind: 'resize-row', row: 3, size: 24 }, options), [])
  assert.deepEqual(createSheetContextMenuItems({ kind: 'resize-column', column: 2, size: 96 }, options), [])
})
