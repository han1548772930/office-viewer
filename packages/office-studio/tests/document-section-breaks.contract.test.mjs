import assert from 'node:assert/strict'
import { afterAll, test } from 'vitest'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
afterAll(() => vite.close())
const { normalizeDocument } = await vite.ssrLoadModule('/packages/office-studio/src/parser/documentNormalizer.ts')

function envelope(blocks) {
  return {
    schemaVersion: 1, parserVersion: 'contract', format: 'docx', fileName: 'breaks.docx', size: 1,
    model: 'document', payload: { modelType: 'document', data: { documentId: 'breaks-doc', blocks } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: ['pageBreaks', 'columnBreaks', 'sectionBreaks'] },
  }
}

test('parser break details normalize to first-class DOCX marker blocks', () => {
  const model = normalizeDocument(envelope([
    { id: 'page-break', type: 'text', content: '', details: { pageBreak: true } },
    { id: 'column-break', type: 'text', content: '', details: { columnBreak: true } },
    { id: 'section-break', type: 'text', content: '', details: { sectionBreak: { type: 'evenPage' } } },
  ]))

  assert.deepEqual(model.blocks.map(block => block.type), ['pageBreak', 'columnBreak', 'sectionBreak'])
  assert.deepEqual(model.blocks.map(block => block.runs.length), [0, 0, 0])
  assert.deepEqual(model.blocks[2].sectionBreak, { type: 'evenPage' })
})
