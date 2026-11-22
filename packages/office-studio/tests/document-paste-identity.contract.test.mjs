import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { collapsedSelection, createParagraph, createRun, createTable, runText, tableForBlock } from '../src/features/document/model/document-model.ts'

test('repeated inline paste keeps existing runs stable and gives inserted runs independent identities', () => {
  const block = createParagraph('A')
  block.runs.push(createRun('B', { bold: true }))
  const source = structuredClone(block)
  const sourceSnapshot = JSON.stringify(source)
  const originalIds = block.runs.map(run => run.id)
  const model = { ...createBlankDocument(), blocks: [block] }
  const command = { id: 'doc.command.paste-blocks', blocks: [source] }
  let selection = collapsedSelection({ blockId: block.id, offset: 2 })
  selection = executeDocumentCommand(model, selection, command)
  executeDocumentCommand(model, selection, command)
  assert.equal(runText(block.runs), 'ABABAB')
  assert.deepEqual(block.runs.slice(0, 2).map(run => run.id), originalIds)
  assert.equal(new Set(block.runs.map(run => run.id)).size, block.runs.length)
  block.runs[2].color = '#123456'
  assert.equal(block.runs[0].color, '#1f2937')
  assert.equal(JSON.stringify(source), sourceSnapshot)
})

test('repeated block and table paste keeps independent block cell and run identities without mutating its source', () => {
  const table = createTable(1, 1)
  table.rows[0].cells[0].blocks[0].runs = [createRun('Cell value')]
  const tableBlock = { ...createParagraph(), type: 'table', metadata: { table } }
  const list = { ...createParagraph('List value'), type: 'numbered', list: { id: 'shared-list', preset: 'ORDER_LIST' } }
  const command = { id: 'doc.command.paste-blocks', blocks: [list, tableBlock] }
  const sourceSnapshot = JSON.stringify(command)
  const paragraph = createParagraph('BeforeAfter')
  const model = { ...createBlankDocument(), blocks: [paragraph] }
  let selection = collapsedSelection({ blockId: paragraph.id, offset: 6 })
  selection = executeDocumentCommand(model, selection, command)
  executeDocumentCommand(model, selection, command)
  const ids = []
  const visit = blocks => {
    for (const block of blocks) {
      ids.push(block.id, ...block.runs.map(run => run.id))
      const table = tableForBlock(block)
      if (table) for (const row of table.rows) {
        ids.push(row.id)
        for (const cell of row.cells) { ids.push(cell.id); visit(cell.blocks) }
      }
    }
  }
  visit(model.blocks)
  assert.equal(new Set(ids).size, ids.length)
  assert.equal(model.blocks.filter(block => block.type === 'numbered').length, 2)
  assert.equal(model.blocks.filter(block => tableForBlock(block)).length, 2)
  assert.equal(runText(model.blocks[0].runs), 'Before')
  assert.equal(runText(model.blocks.at(-1).runs), 'After')
  const pasted = model.blocks.filter(block => tableForBlock(block))
  tableForBlock(pasted[0]).rows[0].cells[0].blocks[0].runs[0].text = 'Edited first cell'
  assert.equal(tableForBlock(pasted[1]).rows[0].cells[0].blocks[0].runs[0].text, 'Cell value')
  assert.equal(JSON.stringify(command), sourceSnapshot)
})

test('pasting an empty inline paragraph at a caret preserves the existing persistent model', () => {
  const paragraph = createParagraph()
  const model = { ...createBlankDocument(), blocks: [paragraph] }
  const selection = collapsedSelection({ blockId: paragraph.id, offset: 0 })
  const before = JSON.stringify(model)
  executeDocumentCommand(model, selection, { id: 'doc.command.paste-blocks', blocks: [createParagraph()] })
  assert.equal(JSON.stringify(model), before)
})

test('inline formatted paste across separate cells preserves the tail in its original cell', () => {
  const table = createTable(1, 2)
  const first = table.rows[0].cells[0].blocks[0], last = table.rows[0].cells[1].blocks[0]
  first.runs = [createRun('First')]
  last.runs = [createRun('Last')]
  const block = { ...createParagraph(), type: 'table', metadata: { table } }
  const model = { ...createBlankDocument(), blocks: [block] }
  const source = createParagraph('X', { bold: true })
  const next = executeDocumentCommand(model, {
    anchor: { blockId: block.id, row: 0, col: 0, cellBlockId: first.id, offset: 2 },
    focus: { blockId: block.id, row: 0, col: 1, cellBlockId: last.id, offset: 2 },
  }, { id: 'doc.command.paste-blocks', blocks: [source] })
  assert.equal(runText(first.runs), 'FiX')
  assert.equal(runText(last.runs), 'st')
  assert.equal(first.runs.at(-1).bold, true)
  assert.notEqual(first.runs.at(-1).id, source.runs[0].id)
  assert.equal(next.focus.cellBlockId, first.id)
  assert.equal(next.focus.offset, 3)
})
