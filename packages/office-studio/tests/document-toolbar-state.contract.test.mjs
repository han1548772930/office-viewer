import assert from 'node:assert/strict'
import { after, test } from 'node:test'
import { createServer } from 'vite'
const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
after(() => vite.close())
const { createBlankDocument } = await vite.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
const { createParagraph, createRun, collapsedSelection } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
const { clampDocumentZoom, documentCharacterCount, documentSelectionCharacterCount, documentToolbarState } = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-toolbar-state.ts')
const { createDocumentRibbon, documentCommandFromRibbon } = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-ribbon.ts')

test('toolbar reports uniform and mixed formats from the actual selected run slices', () => {
  const model = createBlankDocument()
  const first = createParagraph()
  first.runs = [createRun('ab', { bold: true, fontFamily: 'Arial' }), createRun('cd', { fontFamily: 'Georgia' })]
  model.blocks = [first]
  const mixed = documentToolbarState(model, { anchor: { blockId: first.id, offset: 1 }, focus: { blockId: first.id, offset: 3 } })
  assert.equal(mixed.marks.bold, undefined)
  assert.equal(mixed.marks.fontFamily, undefined)
  assert.ok(mixed.mixedMarks.includes('bold'))
  assert.ok(mixed.mixedMarks.includes('fontFamily'))
  const uniform = documentToolbarState(model, { anchor: { blockId: first.id, offset: 0 }, focus: { blockId: first.id, offset: 2 } })
  assert.equal(uniform.marks.bold, true)
  assert.equal(uniform.marks.fontFamily, 'Arial')
})

test('collapsed pending marks affect only the caret and backwards ranges resolve the same state', () => {
  const model = createBlankDocument()
  const paragraph = createParagraph('abcd', { fontSize: 14 })
  model.blocks = [paragraph]
  const caret = documentToolbarState(model, collapsedSelection({ blockId: paragraph.id, offset: 2 }), { bold: true })
  assert.equal(caret.marks.bold, true)
  assert.equal(caret.marks.fontSize, 14)
  const range = documentToolbarState(model, { anchor: { blockId: paragraph.id, offset: 3 }, focus: { blockId: paragraph.id, offset: 1 } }, { fontSize: 28 })
  assert.equal(range.marks.fontSize, 14)
  assert.equal(range.hasSelection, true)
})

test('a range ending at the next paragraph offset zero excludes its toolbar styles', () => {
  const model = createBlankDocument()
  const first = createParagraph('first', { bold: true })
  first.type = 'heading'; first.level = 2
  const second = createParagraph('second')
  second.style.align = 'right'
  model.blocks = [first, second]
  const state = documentToolbarState(model, { anchor: { blockId: first.id, offset: 0 }, focus: { blockId: second.id, offset: 0 } })
  assert.equal(state.marks.bold, true)
  assert.equal(state.paragraph.align, 'left')
  assert.equal(state.blockStyle, 'H2')
})

test('character statistics count selected graphemes and ignore whitespace', () => {
  assert.equal(documentCharacterCount('中文 e\u0301 👨‍👩‍👧‍👦\n'), 4)
  const model = createBlankDocument()
  const block = createParagraph('甲 乙丙')
  model.blocks = [block]
  assert.equal(documentSelectionCharacterCount(model, { anchor: { blockId: block.id, offset: 0 }, focus: { blockId: block.id, offset: 3 } }), 2)
  assert.equal(documentSelectionCharacterCount(model, collapsedSelection({ blockId: block.id, offset: 0 })), 0)
})

test('zoom clamps invalid, fractional and out of range values at the shared boundary', () => {
  assert.equal(clampDocumentZoom(-10), 10)
  assert.equal(clampDocumentZoom(1000), 400)
  assert.equal(clampDocumentZoom(125.4), 125)
  assert.equal(clampDocumentZoom(Number.NaN), 100)
})

test('ribbon schemas derive current values, mixed toggle behavior and dispatch validated Canvas commands', () => {
  const state = { marks: { fontSize: 18 }, mixedMarks: ['bold', 'fontFamily'], paragraph: { align: 'right', indent: 0 }, blockStyle: 'H2', hasSelection: true }
  const tabs = createDocumentRibbon({ state, canUndo: true, canRedo: false, outline: true, focus: false, paper: 'A4', orientation: 'portrait', zoom: 125, contentControlsAvailable: true })
  const items = tabs.flatMap((tab) => tab.groups.flatMap((group) => group.items))
  assert.deepEqual(tabs.map((tab) => tab.id), ['home', 'insert', 'layout', 'view'])
  const font = items.find((item) => item.id === 'doc-font-family')
  assert.equal(font.value, null)
  assert.equal(font.placeholder, '混合字体')
  const bold = items.find((item) => item.id === 'doc-bold')
  assert.equal(bold.activated, false)
  assert.deepEqual(documentCommandFromRibbon(bold.commandId, bold.params), { id: 'doc.command.format-text', patch: { bold: true } })
  assert.deepEqual(documentCommandFromRibbon(font.commandId, font.valueParams('Georgia')), { id: 'doc.command.format-text', patch: { fontFamily: 'Georgia' } })
  assert.equal(items.find((item) => item.id === 'doc-redo').disabled, true)
  assert.equal(documentCommandFromRibbon('doc.command.format-text', { patch: { fontSize: NaN } }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.format-text', { patch: { bold: 'false' } }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.format-paragraph', { patch: { align: 'diagonal' } }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.set-block-type', { type: 'heading', level: 100 }), undefined)
})
