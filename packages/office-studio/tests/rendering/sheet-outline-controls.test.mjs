import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'vitest'
import { SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../../src/rendering/spreadsheet/scene.ts'
import { createSheetOutlineState, addSheetOutlineGroup, getCollapsedOutlineIndexes, setSheetOutlineCollapsed, setSheetOutlineLevel } from '../../src/features/sheet-outline/sheet-outline.ts'

function geometryFor(outline) {
  return new SpreadsheetGeometry({
    id: 'outline-sheet', name: 'Outline', cells: {}, rowCount: 20, columnCount: 10,
    defaultRowHeight: 24, defaultColumnWidth: 80,
    outline,
  })
}

const viewport = { width: 900, height: 500, scrollLeft: 0, scrollTop: 0, zoom: 1 }

test('row and column group controls have stable hit rectangles on summary axes', () => {
  let outline = createSheetOutlineState()
  outline = addSheetOutlineGroup(outline, 'row', 2, 4, 20, 'rows')
  outline = addSheetOutlineGroup(outline, 'column', 1, 3, 10, 'columns')
  const geometry = geometryFor(outline)
  const controls = geometry.outlineToggleRects(viewport)
  const row = controls.find(control => control.groupId === 'rows')
  const column = controls.find(control => control.groupId === 'columns')
  assert.deepEqual(geometry.hitTest(row.x + 6, row.y + 6, viewport), { kind: 'outline-toggle', axis: 'row', groupId: 'rows', collapsed: false })
  assert.deepEqual(geometry.hitTest(column.x + 6, column.y + 6, viewport), { kind: 'outline-toggle', axis: 'column', groupId: 'columns', collapsed: false })
})

test('outline level buttons and bracket drawing commands are present for each grouped axis', () => {
  let outline = createSheetOutlineState()
  outline = addSheetOutlineGroup(outline, 'row', 2, 8, 20, 'outer')
  outline = addSheetOutlineGroup(outline, 'row', 4, 6, 20, 'inner')
  outline = addSheetOutlineGroup(outline, 'column', 1, 4, 10, 'cols')
  const geometry = geometryFor(outline)
  assert.deepEqual(geometry.hitTest(2, 3, viewport), { kind: 'outline-level', axis: 'row', level: 1 })
  assert.deepEqual(geometry.hitTest(2, 15, viewport), { kind: 'outline-level', axis: 'column', level: 1 })
  const scene = createSpreadsheetScene(geometry, viewport, { activeCell: 'A1', selectedCells: [] })
  assert.equal(geometry.outlineBracketLines(viewport).filter(line => line.axis === 'row').length, 6)
  assert.equal(geometry.outlineBracketLines(viewport).filter(line => line.axis === 'column').length, 3)
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.color === '#8593a5'))
  assert.ok(scene.commands.some(command => command.kind === 'rect' && command.width === 12 && command.height === 12))
  assert.ok(scene.commands.some(command => command.kind === 'text' && command.text === '2'))
})

test('level and plus/minus operations update the same persisted group state', () => {
  let outline = createSheetOutlineState()
  outline = addSheetOutlineGroup(outline, 'row', 2, 8, 20, 'outer')
  outline = addSheetOutlineGroup(outline, 'row', 4, 6, 20, 'inner')
  outline = setSheetOutlineLevel(outline, 'row', 1)
  assert.ok(outline.groups.every(group => group.collapsed))
  outline = setSheetOutlineCollapsed(outline, 'outer', false)
  assert.equal(outline.groups.find(group => group.id === 'outer').collapsed, false)
  assert.equal(outline.groups.find(group => group.id === 'inner').collapsed, true)
  outline = setSheetOutlineLevel(outline, 'row', 3)
  assert.ok(outline.groups.every(group => !group.collapsed))
  assert.deepEqual([...getCollapsedOutlineIndexes({ count: 20, groups: outline.groups })], [])
})

test('Canvas hit results dispatch the group id and axis to the persisted workbench operations', () => {
  const canvas = readFileSync(new URL('../../src/components/spreadsheet/SpreadsheetCanvas.vue', import.meta.url), 'utf8')
  const workbench = readFileSync(new URL('../../src/components/ExcelWorkbench.vue', import.meta.url), 'utf8')
  assert.match(canvas, /'sheet\.operation\.outline-toggle-group', \{ axis: hit\.axis, groupId: hit\.groupId \}/)
  assert.match(canvas, /'sheet\.operation\.outline-set-level', \{ axis: hit\.axis, level: hit\.level \}/)
  assert.match(workbench, /setSheetOutlineCollapsed\(sheet\.outline, groupId, !group\.collapsed\)/)
  assert.match(workbench, /setSheetOutlineLevel\(sheet\.outline, axis, level\)/)
})
