import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred source/menu contracts. This batch did not start this server or runner.
let server, domain, model, latex, popup, menu, commands
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-latex-popup', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;[domain, model, latex, popup, menu, commands] = await Promise.all(['domain/model', 'features/document/model/document-model',
    'features/document/latex/latex-formula', 'features/document/controllers/document-latex-popup', 'components/document/document-latex-menu',
    'features/document/commands/document-latex-commands'].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterAll(async () => { await server?.close() })
const paragraph = (id, text = '') => ({ ...model.createParagraph(text), id })
const selection = (id, start = 0, end = start) => ({ anchor: { blockId: id, offset: start }, focus: { blockId: id, offset: end } })
const source = blocks => ({ ...domain.createBlankDocument(), id: 'popup-source', blocks })
const t = key => key

test('the five effective mounts retain fixed upstream host, group and order without counting the floating registration twice', () => {
  assert.deepEqual(menu.DOCUMENT_LATEX_MENU_MOUNTS, [
    { id: 'empty-layout', host: 'doc.menu.empty-paragraph', group: 'layout', order: 8 },
    { id: 'insert-quick-bottom', host: 'doc.menu.paragraph-t.insert', group: 'quickBottom', order: 4 },
    { id: 'insert-menu', host: 'doc.menu.paragraph-t.insert', group: 'insert', order: 3 },
    { id: 'edit-quick-bottom', host: 'doc.menu.paragraph-t.edit', group: 'quickBottom', order: 5 },
    { id: 'float-text-style', host: 'doc.menu.float-toolbar', group: 'doc.menu.float-text-style', order: 13 },
  ])
  assert.equal(new Set(menu.DOCUMENT_LATEX_MENU_MOUNTS.map(item => item.id)).size, 5)
})

test.each(['empty-layout', 'insert-quick-bottom', 'insert-menu', 'edit-quick-bottom', 'float-text-style'])('%s routes its actual enabled/disabled action to the same popup operation', mount => {
  const enabled = menu.documentLatexMountedMenuItem(mount, { available: true }, t)
  const blocked = menu.documentLatexMountedMenuItem(mount, { available: false, reason: 'locked source' }, t)
  assert.equal(enabled.commandId, popup.DOCUMENT_LATEX_POPUP_OPERATIONS.open); assert.equal(enabled.disabled, false)
  assert.equal(blocked.commandId, enabled.commandId); assert.equal(blocked.disabled, true); assert.equal(blocked.disabledReason, 'locked source')
})

test('mobile editing shares the public operation while remove and replace retain checked content commands', () => {
  const document = source([{ ...paragraph('p'), runs: [latex.createLatexFormulaRun(' x_2 ')] }])
  const item = latex.getLatexFormulas(document)[0], actions = menu.documentLatexContextItems(item, { available: true }, t)
  assert.deepEqual(actions.map(item => item.commandId), [popup.DOCUMENT_LATEX_POPUP_OPERATIONS.open,
    'doc.command.remove-latex-formula', 'doc.command.replace-latex-formula-with-text'])
  assert.deepEqual(actions[0].params, { rangeId: item.rangeId })
  assert.deepEqual(menu.documentLatexCommandFromMenu(actions[1].commandId, actions[1].params), { id: actions[1].commandId, rangeId: item.rangeId })
  assert.deepEqual(menu.documentLatexCommandFromMenu(actions[2].commandId, actions[2].params), { id: actions[2].commandId, rangeId: item.rangeId, text: ' x_2 ' })
  assert.equal(menu.documentLatexCommandFromMenu(actions[0].commandId, actions[0].params), undefined)
  assert.equal(menu.documentLatexCommandFromMenu(actions[1].commandId, { rangeId: '' }), undefined)
})

test('ordinary selected text retains whitespace and opening a seed leaves every run unchanged', () => {
  const document = source([paragraph('p', '  x + y  ')]), before = structuredClone(document)
  assert.equal(latex.documentLatexSourceForSelection(document, selection('p', 0, 9)), '  x + y  ')
  assert.equal(latex.documentLatexSourceForSelection(document, selection('p', 3)), '')
  assert.deepEqual(document, before)
})

test('mixed text and complete formula atoms seed their exact original source rather than the object replacement character', () => {
  const document = source([{ ...paragraph('p'), runs: [model.createRun('a+'), latex.createLatexFormulaRun('  x ^2  '), model.createRun('=b')] }])
  assert.equal(latex.documentLatexSourceForSelection(document, selection('p', 0, 5)), 'a+  x ^2  =b')
  assert.equal(latex.documentLatexSourceForSelection(document, selection('p', 2, 3)), '  x ^2  ')
  assert.equal(latex.documentLatexSourceForSelection(document, selection('p', 5, 0)), 'a+  x ^2  =b')
})

test.each([[1, 3, '😀'], [3, 5, 'e\u0301'], [0, 6, 'A😀e\u0301中']])('Unicode seed %i..%i retains complete UTF-16 and combining content', (start, end, expected) => {
  assert.equal(latex.documentLatexSourceForSelection(source([paragraph('p', 'A😀e\u0301中')]), selection('p', start, end)), expected)
})

test.each([[1, 2], [2, 3]])('a source boundary %i..%i inside a surrogate pair is rejected by the same availability gate', (start, end) => {
  const document = source([paragraph('p', 'A😀B')]), range = selection('p', start, end)
  assert.throws(() => latex.documentLatexSourceForSelection(document, range), /Unicode/)
  assert.equal(commands.documentLatexSelectionState(document, range, { documentId: document.id, userId: 'editor' }).available, false)
})

test.each([-1, 0.5, NaN, Infinity, 20])('invalid raw offset %s is rejected before generic selection clamping can redirect the seed', offset => {
  assert.throws(() => latex.documentLatexSourceForSelection(source([paragraph('p', 'xy')]), selection('p', offset)), /越界/)
})

test('deleted targets, cross-paragraph ranges and table structure selections never produce a misleading empty draft', () => {
  const document = source([paragraph('p', 'xy'), paragraph('tail', 'tail')])
  assert.throws(() => latex.documentLatexSourceForSelection(document, selection('deleted')), /目标已不存在/)
  assert.throws(() => latex.documentLatexSourceForSelection(document, { anchor: { blockId: 'p', offset: 0 }, focus: { blockId: 'tail', offset: 1 } }), /同一段落/)
  assert.throws(() => latex.documentLatexSourceForSelection(document, { ...selection('p'), tableSelection: { tableId: 'table' } }), /文字选区/)
})

test('expanded source, including formula metadata, observes the editor source budget', () => {
  const document = source([paragraph('p', 'x'.repeat(65_537))])
  assert.throws(() => latex.documentLatexSourceForSelection(document, selection('p', 0, 65_537)), /预算/)
  const atom = latex.createLatexFormulaRun('x'); atom.metadata.latexFormula.latex = 'x'.repeat(65_537)
  assert.throws(() => latex.documentLatexSourceForSelection(source([{ ...paragraph('p'), runs: [atom] }]), selection('p', 0, 1)), /预算/)
})

test.each([{}, { startIndex: 2, endIndex: 4 }, { selection: { startOffset: 3, endOffset: 3, collapsed: true, direction: 'none' } },
  { unitId: 'doc', segmentId: 'header', startIndex: 2, endIndex: 4, selection: { startOffset: 2, endOffset: 5, collapsed: false, direction: 'backward', segmentId: 'header', segmentPage: 0, isActive: true, rangeType: 'TEXT' }, segmentPage: 0 }])('valid public popup request %# accepts the fixed text range direction and explicit range convention', request => {
  assert.doesNotThrow(() => popup.assertDocumentLatexPopupRequest(request))
})

test.each([null, [], { unknown: true }, { unitId: ' ' }, { rangeId: '' }, { startIndex: 0 }, { endIndex: 2 },
  { startIndex: 3, endIndex: 2 }, { startIndex: 0, endIndex: Number.MAX_SAFE_INTEGER }, { segmentPage: -1 },
  { expectedInputGeneration: 0.5 }, { initialLatex: 'x'.repeat(65_537) }, { selection: { startOffset: 0, endOffset: 1, direction: 'sideways' } },
  { selection: { startOffset: 1, endOffset: 2, collapsed: true } }, { selection: { startOffset: 2, endOffset: 1 } },
  { startIndex: 0, endIndex: 2, selection: { startOffset: 0, endOffset: 2 } },
  { segmentId: 'body', selection: { startOffset: 0, endOffset: 0, segmentId: 'header' } }])('malformed popup request %# is rejected before reaching a UI owner', request => {
  assert.throws(() => popup.assertDocumentLatexPopupRequest(request))
})
