import assert from 'node:assert/strict'
import { afterAll, test } from 'vitest'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
afterAll(() => vite.close())
const { createParagraph, createTable, runText, runsForTarget, textTargets } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
const { findDocumentMatches, replaceDocumentMatches } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/services/document-search.ts')

function createModel() {
  const paragraph = createParagraph('Alpha alpha alphabet')
  paragraph.id = 'paragraph'
  const table = createParagraph()
  table.id = 'table'
  table.type = 'table'
  table.metadata.table = createTable(1, 1)
  table.metadata.table.rows[0].cells[0].blocks[0].runs[0].text = 'alpha in cell'
  return { kind: 'document', id: 'doc', name: 'Search', blocks: [paragraph, table], page: {}, modified: false, wordCount: 0 }
}

test('finds all non-overlapping matches with case and Unicode whole-word options', () => {
  const model = createModel()
  assert.equal(findDocumentMatches(model, 'alpha').length, 4)
  assert.equal(findDocumentMatches(model, 'alpha', { caseSensitive: true }).length, 3)
  assert.equal(findDocumentMatches(model, 'alpha', { wholeWord: true }).length, 3)
  assert.equal(findDocumentMatches(model, 'ALPHA', { caseSensitive: true }).length, 0)
  assert.equal(findDocumentMatches(model, 'a.*a').length, 0)
})

test('replaces all matches from the end while preserving surrounding text and table cell targets', () => {
  const model = createModel()
  const matches = findDocumentMatches(model, 'alpha', { wholeWord: true })
  const selection = replaceDocumentMatches(model, matches, 'X')
  const targets = textTargets(model)
  assert.equal(runText(runsForTarget(model, targets[0])), 'X X alphabet')
  assert.equal(runText(runsForTarget(model, targets[1])), 'X in cell')
  assert.equal(selection?.focus.offset, 1)
})
