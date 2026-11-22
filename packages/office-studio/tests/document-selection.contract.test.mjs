import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createTable } from '../src/features/document/model/document-model.ts'
import { documentSelectionContainsPoint, documentSelectionForContextMenu } from '../src/features/document/model/document-selection.ts'

const point = (blockId, offset, coordinates = {}) => ({ blockId, offset, ...coordinates })
const range = (anchor, focus) => ({ anchor, focus })
function documentWith(blocks) { return { ...createBlankDocument(), blocks } }

test('right-clicking selected text preserves forward and reverse ranges and their half-open boundaries', () => {
  const paragraph = { ...createParagraph('甲乙丙丁戊'), id: 'paragraph' }
  const model = documentWith([paragraph])
  for (const selection of [range(point('paragraph', 1), point('paragraph', 4)), range(point('paragraph', 4), point('paragraph', 1))]) {
    assert.equal(documentSelectionForContextMenu(model, selection, point('paragraph', 1)), selection)
    assert.equal(documentSelectionForContextMenu(model, selection, point('paragraph', 3)), selection)
    assert.deepEqual(documentSelectionForContextMenu(model, selection, point('paragraph', 0)), range(point('paragraph', 0), point('paragraph', 0)))
    assert.deepEqual(documentSelectionForContextMenu(model, selection, point('paragraph', 4)), range(point('paragraph', 4), point('paragraph', 4)))
  }
})

test('paragraph order, including intervening paragraphs, determines whether a context click is inside', () => {
  const model = documentWith(['start', 'middle', 'end', 'outside'].map(id => ({ ...createParagraph('abcde'), id })))
  const selection = range(point('end', 2), point('start', 3))
  assert.equal(documentSelectionForContextMenu(model, selection, point('middle', 0)), selection)
  assert.equal(documentSelectionForContextMenu(model, selection, point('end', 1)), selection)
  assert.equal(documentSelectionContainsPoint(model, selection, point('start', 2)), false)
  assert.equal(documentSelectionContainsPoint(model, selection, point('end', 2)), false)
  assert.equal(documentSelectionContainsPoint(model, selection, point('outside', 0)), false)
})

test('table context clicks compare the row, column and cell paragraph rather than only the table id', () => {
  const table = { ...createParagraph(), id: 'table', type: 'table', runs: [], metadata: { table: createTable(2, 2) } }
  const rows = table.metadata.table.rows
  for (let row = 0; row < 2; row++) for (let col = 0; col < 2; col++) {
    rows[row].cells[col].blocks = [0, 1].map(index => ({ ...createParagraph('ABCDE'), id: `r${row}c${col}p${index}` }))
  }
  const model = documentWith([table])
  const cellPoint = (row, col, paragraph, offset) => point('table', offset, { row, col, cellBlockId: `r${row}c${col}p${paragraph}` })
  const selection = range(cellPoint(1, 0, 0, 2), cellPoint(0, 1, 0, 1))
  assert.equal(documentSelectionForContextMenu(model, selection, cellPoint(0, 1, 1, 4)), selection)
  assert.equal(documentSelectionForContextMenu(model, selection, cellPoint(1, 0, 0, 1)), selection)
  assert.equal(documentSelectionContainsPoint(model, selection, cellPoint(0, 0, 1, 4)), false)
  assert.equal(documentSelectionContainsPoint(model, selection, cellPoint(1, 0, 1, 0)), false)
  assert.equal(documentSelectionContainsPoint(model, selection, cellPoint(1, 1, 0, 0)), false)
})

test('unknown targets and invalid offsets cannot be mistaken for a selected paragraph', () => {
  const model = documentWith([{ ...createParagraph('text'), id: 'one' }])
  const selection = range(point('one', 0), point('one', 4))
  for (const candidate of [point('missing', 1), point('one', -1), point('one', 5), point('one', 1.5)]) {
    assert.equal(documentSelectionContainsPoint(model, selection, candidate), false)
  }
  const caret = range(point('one', 2), point('one', 2))
  assert.equal(documentSelectionForContextMenu(model, caret, point('one', 2)), caret)
  assert.deepEqual(documentSelectionForContextMenu(model, caret, point('one', 1)), range(point('one', 1), point('one', 1)))
})
