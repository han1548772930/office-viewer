import assert from 'node:assert/strict'
import { afterAll, test } from 'vitest'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
const { createBlankDocument } = await vite.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
const { createParagraph, createRun } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
const { normalizeDocument } = await vite.ssrLoadModule('/packages/office-studio/src/parser/documentNormalizer.ts')
const { executeDocumentCommand } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')

afterAll(() => vite.close())

test('parser normalization recognizes vertical alignment from inline style DTOs', () => {
  const imported = normalizeDocument({
    schemaVersion: 1,
    parserVersion: 'contract',
    format: 'docx',
    fileName: 'formula.docx',
    size: 1,
    model: 'document',
    payload: { modelType: 'document', data: { documentId: 'formula', blocks: [{
      id: 'formula-paragraph',
      type: 'paragraph',
      details: { inlines: [
        { type: 'text', text: 'H', style: { verticalAlign: 'baseline' } },
        { type: 'text', text: '2', style: { vertAlign: 'sub' } },
        { type: 'text', text: 'O', verticalAlign: 'sup' },
      ] },
    }] } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] },
  })

  assert.deepEqual(imported.blocks[0].runs.map(run => run.verticalAlign), [undefined, 'subscript', 'superscript'])
})

test('document format command applies superscript and baseline to only the selected range', () => {
  const model = createBlankDocument()
  const paragraph = createParagraph('H2O')
  model.blocks = [paragraph]
  const selection = { anchor: { blockId: paragraph.id, offset: 1 }, focus: { blockId: paragraph.id, offset: 2 } }

  executeDocumentCommand(model, selection, { id: 'doc.command.format-text', patch: { verticalAlign: 'subscript' } })
  assert.deepEqual(model.blocks[0].runs.map(run => [run.text, run.verticalAlign]), [['H', undefined], ['2', 'subscript'], ['O', undefined]])
  executeDocumentCommand(model, selection, { id: 'doc.command.format-text', patch: { verticalAlign: 'baseline' } })
  assert.deepEqual(model.blocks[0].runs.map(run => [run.text, run.verticalAlign]), [['H2O', undefined]])
})

test('baseline patches clear legacy vertical alignment and compact with adjacent normal text', () => {
  const model = createBlankDocument()
  const paragraph = createParagraph()
  paragraph.runs = [
    createRun('a'),
    createRun('b', { verticalAlign: 'superscript', metadata: { verticalAlign: 'superscript' } }),
  ]
  model.blocks = [paragraph]

  executeDocumentCommand(model, { anchor: { blockId: paragraph.id, offset: 1 }, focus: { blockId: paragraph.id, offset: 2 } }, {
    id: 'doc.command.format-text', patch: { verticalAlign: 'baseline' },
  })

  assert.equal(model.blocks[0].runs.length, 1)
  assert.equal(model.blocks[0].runs[0].text, 'ab')
  assert.equal(model.blocks[0].runs[0].verticalAlign, undefined)
  assert.equal(model.blocks[0].runs[0].metadata, undefined)
})
