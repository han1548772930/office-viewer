import assert from 'node:assert/strict'
import { afterAll, test } from 'vitest'
import { createServer } from 'vite'
const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
afterAll(() => vite.close())
const { createBlankDocument } = await vite.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
const { createParagraph, createRun, collapsedSelection } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
const { clampDocumentZoom, documentCharacterCount, documentSelectionCharacterCount, documentToolbarState } = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-toolbar-state.ts')
const { createDocumentRibbon, documentCommandFromRibbon } = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-ribbon.ts')
const { documentOrientedPage, documentPageSettingsError, documentPaperSize } = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-page-settings.ts')

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

test('subscript and superscript expose uniform and mixed selection state', () => {
  const model = createBlankDocument()
  const paragraph = createParagraph()
  paragraph.runs = [createRun('H', { verticalAlign: 'baseline' }), createRun('2', { verticalAlign: 'subscript' }), createRun('O', { verticalAlign: 'superscript' })]
  model.blocks = [paragraph]

  const subscript = documentToolbarState(model, { anchor: { blockId: paragraph.id, offset: 1 }, focus: { blockId: paragraph.id, offset: 2 } })
  assert.equal(subscript.marks.verticalAlign, 'subscript')
  assert.deepEqual(subscript.mixedMarks, [])
  const mixed = documentToolbarState(model, { anchor: { blockId: paragraph.id, offset: 0 }, focus: { blockId: paragraph.id, offset: 2 } })
  assert.equal(mixed.marks.verticalAlign, undefined)
  assert.ok(mixed.mixedMarks.includes('verticalAlign'))
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
  assert.deepEqual(tabs.map((tab) => tab.id), ['home', 'insert', 'view'])
  const breakMenu = items.find((item) => item.id === 'doc-breaks')
  assert.ok(breakMenu)
  const columnBreak = breakMenu.children.find((item) => item.id === 'doc-column-break')
  assert.deepEqual(documentCommandFromRibbon(columnBreak.commandId, columnBreak.params), { id: 'doc.command.insert-column-break' })
  const sectionMenu = breakMenu.children.find((item) => item.id === 'doc-section-breaks')
  assert.deepEqual(sectionMenu.children.map((item) => item.params.type), ['nextPage', 'continuous', 'nextColumn', 'evenPage', 'oddPage'])
  for (const item of sectionMenu.children) {
    assert.deepEqual(documentCommandFromRibbon(item.commandId, item.params), { id: 'doc.command.insert-section-break', type: item.params.type })
  }
  assert.equal(documentCommandFromRibbon('doc.command.insert-section-break', { type: 'unexpected' }), undefined)
  const font = items.find((item) => item.id === 'doc-font-family')
  assert.equal(font.value, null)
  assert.equal(font.placeholder, '混合字体')
  const bold = items.find((item) => item.id === 'doc-bold')
  assert.equal(bold.activated, false)
  assert.deepEqual(documentCommandFromRibbon(bold.commandId, bold.params), { id: 'doc.command.format-text', patch: { bold: true } })
  const superscript = items.find((item) => item.id === 'doc-superscript')
  assert.equal(superscript.activated, false)
  assert.deepEqual(documentCommandFromRibbon(superscript.commandId, superscript.params), { id: 'doc.command.format-text', patch: { verticalAlign: 'superscript' } })
  const activeSubscript = createDocumentRibbon({ state: { ...state, marks: { verticalAlign: 'subscript' } }, canUndo: true, canRedo: false, outline: true, focus: false, fullscreen: false, paper: 'A4', orientation: 'portrait', zoom: 125, contentControlsAvailable: true })
    .flatMap((tab) => tab.groups.flatMap((group) => group.items)).find((item) => item.id === 'doc-subscript')
  assert.equal(activeSubscript.activated, true)
  assert.deepEqual(documentCommandFromRibbon(activeSubscript.commandId, activeSubscript.params), { id: 'doc.command.format-text', patch: { verticalAlign: 'baseline' } })
  assert.deepEqual(documentCommandFromRibbon(font.commandId, font.valueParams('Georgia')), { id: 'doc.command.format-text', patch: { fontFamily: 'Georgia' } })
  assert.equal(items.find((item) => item.id === 'doc-redo').disabled, true)
  const viewItems = tabs.find((tab) => tab.id === 'view').groups.flatMap((group) => group.items)
  assert.ok(viewItems.find((item) => item.id === 'doc-page-settings'))
  assert.ok(viewItems.find((item) => item.id === 'doc-fullscreen'))
  assert.equal(tabs.some((tab) => tab.id === 'layout'), false)
  assert.equal(documentCommandFromRibbon('doc.command.format-text', { patch: { fontSize: NaN } }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.format-text', { patch: { bold: 'false' } }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.format-paragraph', { patch: { align: 'diagonal' } }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.set-block-type', { type: 'heading', level: 100 }), undefined)
  const quote = items.find((item) => item.id === 'doc-block-style').options.find((option) => option.value === 'quote')
  assert.deepEqual(documentCommandFromRibbon('doc.command.set-block-type', quote.params), { id: 'doc.command.set-block-type', type: 'quote', level: undefined })
  const checklist = items.find((item) => item.id === 'doc-checklist')
  assert.equal(checklist.commandId, 'doc.command.set-block-type')
  assert.deepEqual(documentCommandFromRibbon(checklist.commandId, checklist.params), { id: 'doc.command.set-block-type', type: 'checklist', level: undefined })
})

test('custom imported page geometry is preserved by orientation and identified accurately', () => {
  const page = { width: 960, height: 1200, marginTop: 45, marginRight: 55, marginBottom: 65, marginLeft: 75 }
  assert.equal(documentPaperSize(page), 'Custom')
  assert.deepEqual(documentOrientedPage(page, 'landscape', 'Custom'), { ...page, width: 1200, height: 960 })
  assert.equal(documentPaperSize({ ...page, width: 1123, height: 794 }), 'A4')
  assert.equal(documentPaperSize({ ...page, width: 1056, height: 816 }), 'Letter')
  assert.equal(documentPageSettingsError(page), undefined)
})

test('page settings reject invalid dimensions and margins before a document mutation', () => {
  const page = createBlankDocument().page
  assert.match(documentPageSettingsError({ ...page, width: NaN }), /有效数字/)
  assert.match(documentPageSettingsError({ ...page, width: 50 }), /宽度/)
  assert.match(documentPageSettingsError({ ...page, marginTop: -1 }), /负数/)
  assert.match(documentPageSettingsError({ ...page, marginLeft: page.width }), /之和/)
  assert.match(documentPageSettingsError({ ...page, marginLeft: page.width * 0.45 }), /42%/)
})
