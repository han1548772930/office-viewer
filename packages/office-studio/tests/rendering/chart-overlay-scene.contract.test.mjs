import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet } from '../../src/domain/model.ts'
import { createSpreadsheetScene } from '../../src/rendering/spreadsheet/scene.ts'
import { SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'

test('chart scene commands use the same worksheet zoom and scroll transform as cells', () => {
  const sheet = createBlankSheet()
  sheet.rowCount = 30
  sheet.columnCount = 12
  const viewport = { width: 600, height: 400, scrollLeft: 30, scrollTop: 40, zoom: 1.5 }
  const scene = createSpreadsheetScene(new SpreadsheetGeometry({
    ...sheet,
    cells: {},
    merges: [],
    freeze: { rows: 0, columns: 0 },
  }), viewport, {
    activeCell: 'A1',
    selectedCells: [],
    overlayCommands: [
      { kind: 'rect', x: 100, y: 80, width: 200, height: 120, fill: '#fff', stroke: '#333' },
      { kind: 'text', x: 120, y: 100, text: 'Revenue', font: '12px Arial', color: '#222' },
    ],
  })

  const chartRect = scene.commands.find(command => command.kind === 'rect' && command.fill === '#fff')
  const chartTitle = scene.commands.find(command => command.kind === 'text' && command.text === 'Revenue')
  assert.deepEqual({ x: chartRect.x, y: chartRect.y, width: chartRect.width, height: chartRect.height }, { x: 120, y: 80, width: 300, height: 180 })
  assert.equal(chartTitle.x, 150)
  assert.equal(chartTitle.y, 110)
  assert.equal(chartTitle.font, '18px Arial')
  assert.ok(scene.commands.indexOf(chartRect) < scene.commands.findIndex(command => command.kind === 'rect' && command.x === 0 && command.y === 0 && command.height === 26))
})
