import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSpreadsheetScene } from '../../src/rendering/spreadsheet/scene.ts'
import { GRID_HEADER_HEIGHT, SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'

test('column labels are vertically centered in the worksheet header at every render scale', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'header-alignment', name: 'Header alignment', cells: {}, rowCount: 4, columnCount: 4,
  })

  for (const zoom of [0.75, 1, 1.5, 2]) {
    const scene = createSpreadsheetScene(
      geometry,
      { width: 500, height: 300, scrollLeft: 0, scrollTop: 0, zoom },
      { activeCell: 'A1', selectedCells: [] },
    )
    const label = scene.commands.find(command => command.kind === 'text' && command.text === 'A')
    assert.ok(label)
    assert.equal(label.y, GRID_HEADER_HEIGHT / 2)
    assert.equal(label.baseline, 'middle')
  }
})
