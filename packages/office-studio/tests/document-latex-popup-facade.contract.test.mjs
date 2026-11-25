import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred API/owner/history contracts; nothing in this file was executed here.
let server, OfficeStudio, FDocument, domain, model, latex, popup, editors, resolver, codec, controls
const studios = new Set(), copy = value => structuredClone(value)
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-latex-popup-facade', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, model, latex, popup, editors, resolver, codec, controls] = await Promise.all(['domain/model', 'features/document/model/document-model',
    'features/document/latex/latex-formula', 'features/document/controllers/document-latex-popup', 'api/document-latex-popup-editors',
    'api/document-latex-popup', 'api/document-body-codec', 'features/document/services/content-controls']
    .map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const paragraph = (id, text = '') => ({ ...model.createParagraph(text), id })
const selection = (id, start = 0, end = start) => ({ anchor: { blockId: id, offset: start }, focus: { blockId: id, offset: end } })
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })
function open(input) {
  const source = input ?? { ...domain.createBlankDocument(), id: 'popup-api', blocks: [paragraph('p', 'ABCD'), paragraph('tail', 'Tail')], modified: false }
  const studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(source.id)
  const document = new FDocument(studio, source.id); document.setInputSelection(selection('p'))
  return { studio, document }
}
const named = formulas => ({ name: 'DOC_LATEX_PLUGIN', data: JSON.stringify({ formulas }) })

test('an unbound editor declines open and accepts close without changing snapshot, revision, input or history', () => {
  const { studio, document } = open(), before = copy(document.getModel()), state = copy(document.getInputState()), revision = document.getModelRevision()
  assert.equal(document.openLatexPopup(), false); assert.equal(document.closeLatexPopup(), true)
  assert.deepEqual(document.getModel(), before); assert.deepEqual(document.getInputState(), state)
  assert.equal(document.getModelRevision(), revision); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 })
})

test('Facade, public and local operations share the live model identity and the same source-only UI owner', () => {
  const { studio, document } = open(), calls = [], closes = []
  document.setInputSelection(selection('p', 1, 3))
  const before = copy(document.getModel()), input = copy(document.getInputState()), generation = document.getHistoryGeneration(), revision = document.getModelRevision()
  const handle = document.bindLatexPopupEditor({ open: target => { calls.push(target); return true }, close: force => { closes.push(force); return true } })
  assert.equal(document.openLatexPopup(), true)
  assert.equal(studio.executeCommandSync(popup.DOCUMENT_LATEX_POPUP_OPERATIONS.open, { unitId: document.id }), true)
  assert.equal(studio.executeCommandSync(popup.DOCUMENT_LATEX_POPUP_OPERATIONS.localOpen, { unitId: document.id, initialLatex: 'draft_2' }), true)
  assert.equal(calls.length, 3); assert.deepEqual(calls.map(target => target.initialLatex), ['BC', 'BC', 'draft_2'])
  for (const target of calls) { assert.equal(target.source, document.getModel()); assert.equal(target.source, studio.getDocumentModel(document.id)); assert.deepEqual(target.selection, selection('p', 1, 3)) }
  assert.equal(document.closeLatexPopup(), true)
  assert.equal(studio.executeCommandSync(popup.DOCUMENT_LATEX_POPUP_OPERATIONS.localClose, { unitId: document.id }), true)
  assert.deepEqual(closes, [false, false]); assert.deepEqual(document.getModel(), before); assert.deepEqual(document.getInputState(), input)
  assert.equal(document.getModelRevision(), revision); assert.equal(document.getHistoryGeneration(), generation)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); handle.dispose()
})

test.each(['none', 'forward', 'backward'])('inclusive public indices and exclusive %s selections resolve to the same source with preserved direction', direction => {
  const { document } = open(), calls = [], before = copy(document.getModel())
  const handle = document.bindLatexPopupEditor({ open: target => { calls.push(target); return true }, close: () => true })
  assert.equal(document.openLatexPopup({ startIndex: 1, endIndex: 2, selection: { startOffset: 1, endOffset: 3, collapsed: false, direction } }), true)
  assert.equal(calls[0].initialLatex, 'BC')
  assert.deepEqual(calls[0].selection, direction === 'backward' ? selection('p', 3, 1) : selection('p', 1, 3))
  assert.deepEqual(document.getModel(), before); assert.deepEqual(document.getInputState().location.selection, selection('p')); handle.dispose()
})

test('a range ID resolves its complete original source while a mixed selection expands the same formula atom', () => {
  const { studio, document } = open(), formula = document.insertLatexAtOffset(1, '  x ^2  '), range = formula.getRange(), calls = []
  const before = copy(document.getModel()), previous = history(studio, document.id)
  const handle = document.bindLatexPopupEditor({ open: target => { calls.push(target); return true }, close: () => true })
  assert.equal(document.openLatexPopup({ rangeId: formula.getId() }), true)
  assert.equal(calls[0].rangeId, formula.getId()); assert.equal(calls[0].initialLatex, '  x ^2  '); assert.deepEqual(calls[0].selection, selection('p', 1, 2))
  assert.equal(document.openLatexPopup({ startIndex: 0, endIndex: range.endOffset - 1 }), true)
  assert.equal(calls[1].rangeId, undefined); assert.equal(calls[1].initialLatex, 'A  x ^2  ')
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous); handle.dispose()
})

test('partial formula source, mismatched range identity, unknown segment and cross-paragraph ranges fail before a UI callback', () => {
  const { studio, document } = open(), formula = document.insertLatexAtOffset(1, 'x_2'), range = formula.getRange(), before = copy(document.getModel()), previous = history(studio, document.id)
  let count = 0
  const handle = document.bindLatexPopupEditor({ open: () => { count++; return true }, close: () => true })
  const tailOffset = document.getBody().dataStream.indexOf('Tail')
  for (const input of [{ startIndex: range.startOffset + 1, endIndex: range.endOffset - 1 },
    { rangeId: formula.getId(), startIndex: 0, endIndex: 0 }, { rangeId: 'deleted' }, { segmentId: 'missing', startIndex: 0, endIndex: 0 },
    { startIndex: 0, endIndex: tailOffset + 1 }, { selection: { startOffset: 0, endOffset: document.getBody().dataStream.length + 1 } },
    { selection: { startOffset: 0, endOffset: 1, rangeType: 'RECT' } }]) {
    assert.throws(() => document.openLatexPopup(input))
  }
  assert.equal(count, 0); assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), previous); handle.dispose()
})

test.each(['expectedRevision', 'expectedHistoryGeneration', 'expectedInputGeneration'])('%s guards reject a popup from a different content/history/input generation', key => {
  const { studio, document } = open(), before = copy(document.getModel())
  const handle = document.bindLatexPopupEditor({ open: () => { throw new Error('Stale target reached UI') }, close: () => true })
  const current = key === 'expectedRevision' ? document.getModelRevision() : key === 'expectedHistoryGeneration' ? document.getHistoryGeneration() : document.getInputState().generation
  assert.throws(() => document.openLatexPopup({ [key]: current + 1 }), /变化/)
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); handle.dispose()
})

test('an explicit empty draft overrides selected source without authoring it', () => {
  const { studio, document } = open(), calls = []
  document.setInputSelection(selection('p', 0, 4))
  const handle = document.bindLatexPopupEditor({ open: target => { calls.push(target.initialLatex); return true }, close: () => true })
  assert.equal(document.openLatexPopup({ initialLatex: '' }), true); assert.deepEqual(calls, [''])
  assert.equal(document.getBody().dataStream.startsWith('ABCD'), true); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); handle.dispose()
})

test('replacement force-closes the old owner and its old disposable cannot remove the new one', () => {
  const { document } = open(), calls = []
  const old = document.bindLatexPopupEditor({ open: () => { calls.push('old open'); return true }, close: force => { calls.push('old close:' + force); return true } })
  const current = document.bindLatexPopupEditor({ open: () => { calls.push('new open'); return true }, close: force => { calls.push('new close:' + force); return true } })
  assert.equal(old.isCurrent(), false); assert.equal(current.isCurrent(), true)
  old.dispose(); old.dispose(); assert.equal(document.openLatexPopup(), true)
  assert.deepEqual(calls, ['old close:true', 'new open']); current.dispose(); current.dispose(); assert.equal(current.isCurrent(), false)
  assert.equal(document.openLatexPopup(), false); assert.deepEqual(calls, ['old close:true', 'new open', 'new close:true'])
})

test('popup owners are isolated between units and studios, including the same unit ID', () => {
  const first = open(), second = open(), otherModel = first.studio.createDocument({ ...domain.createBlankDocument(), id: 'other-popup', blocks: [paragraph('p')] }, false)
  const other = new FDocument(first.studio, otherModel.id), calls = []
  other.setInputSelection(selection('p'))
  const handles = [first.document.bindLatexPopupEditor({ open: () => { calls.push('first'); return true }, close: () => true }),
    other.bindLatexPopupEditor({ open: () => { calls.push('other'); return true }, close: () => true }),
    second.document.bindLatexPopupEditor({ open: () => { calls.push('second'); return true }, close: () => true })]
  assert.equal(other.openLatexPopup(), true); assert.equal(second.document.openLatexPopup(), true); assert.deepEqual(calls, ['other', 'second'])
  assert.equal(handles[0].isCurrent(), true); handles.forEach(handle => handle.dispose())
})

test('ordinary close can reject composition, while unit disposal force-closes and clears ownership before reusing an ID', () => {
  const { studio, document } = open(), closes = []
  const handle = document.bindLatexPopupEditor({ open: () => true, close: force => { closes.push(force); return !!force } })
  assert.equal(document.closeLatexPopup(), false); assert.equal(handle.isCurrent(), true)
  assert.equal(studio.disposeUnit(document.id), true); assert.equal(handle.isCurrent(), false); assert.deepEqual(closes, [false, true])
  studio.createDocument({ ...domain.createBlankDocument(), id: document.id, blocks: [paragraph('p')] }, false)
  const reopened = new FDocument(studio, document.id); reopened.setInputSelection(selection('p')); assert.equal(reopened.openLatexPopup(), false)
})

test('studio disposal force-closes every owner and rejects a captured target after disposal', () => {
  const { studio, document } = open(), target = resolver.resolveDocumentLatexPopupTarget(document, {}), closes = []
  const handle = document.bindLatexPopupEditor({ open: () => true, close: force => { closes.push(force); return true } })
  studio.dispose(); assert.deepEqual(closes, [true]); assert.equal(handle.isCurrent(), false)
  assert.equal(editors.requestDocumentLatexPopupOpen(studio, target), false)
})

test('a document model replacement makes an already captured UI target unusable', () => {
  const { studio, document } = open(), target = resolver.resolveDocumentLatexPopupTarget(document, {}), calls = []
  const handle = document.bindLatexPopupEditor({ open: () => { calls.push('open'); return true }, close: () => true })
  document.executeCommand({ id: 'doc.command.insert-text', text: 'Z' }, selection('p'))
  assert.notEqual(document.getModel(), target.source); assert.equal(editors.requestDocumentLatexPopupOpen(studio, target), false)
  assert.deepEqual(calls, []); handle.dispose()
})

test('command cancellation and a rejected owner retain source, pending marks and history', () => {
  const { studio, document } = open(), calls = []
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, selection('p'))
  const before = copy(document.getModel()), input = copy(document.getInputState())
  assert.equal(input.marks.bold, true)
  const handle = document.bindLatexPopupEditor({ open: target => { calls.push(target.initialLatex); return false }, close: () => true })
  assert.equal(document.openLatexPopup({ initialLatex: 'draft' }), false)
  const stop = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === popup.DOCUMENT_LATEX_POPUP_OPERATIONS.open) event.cancel = true })
  try { assert.throws(() => document.openLatexPopup(), /canceled/) } finally { stop.dispose() }
  assert.deepEqual(calls, ['draft']); assert.deepEqual(document.getModel(), before); assert.deepEqual(document.getInputState(), input)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); handle.dispose()
})

test('a content-locked source fails the shared availability check before an editor can open', () => {
  const source = { ...domain.createBlankDocument(), id: 'locked-popup', blocks: [paragraph('p', 'ABCD')] }
  controls.createDocumentContentControl(source, { id: 'lock', type: 'richText', blockId: 'p', startOffset: 0, endOffset: 3, lock: 'contentLocked' })
  const { studio, document } = open(source), before = copy(document.getModel())
  document.setInputSelection(selection('p', 1, 2))
  const handle = document.bindLatexPopupEditor({ open: () => { throw new Error('Locked source reached UI') }, close: () => true })
  assert.throws(() => document.openLatexPopup(), /锁定|编辑/); assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); handle.dispose()
})

test('explicit header and note segments preserve formula source, page affinity and host without changing the active body selection', () => {
  const saved = open().document.save(), header = codec.encodeDocumentBody({ blocks: [{ ...paragraph('header-p'), runs: [latex.createLatexFormulaRun('h_1')] }], runs: [], paragraphs: [], metadata: {} })
  saved.documentStyle.defaultHeaderId = 'popup-header'; saved.headers = { 'popup-header': { ...header, headerId: 'popup-header', resources: [named(header.latexFormulas)] } }
  const { studio, document } = open(saved), note = codec.encodeDocumentBody({ blocks: [{ ...paragraph('note-p'), runs: [latex.createLatexFormulaRun('n_1')] }], runs: [], paragraphs: [], metadata: {} })
  const noteId = document.insertFootnote(1, { ...note, resources: [named(note.latexFormulas)] }); document.setInputSelection(selection('p'))
  const calls = [], input = copy(document.getInputState()), previous = history(studio, document.id)
  const handle = document.bindLatexPopupEditor({ open: target => { calls.push(target); return true }, close: () => true })
  for (const segmentId of ['popup-header', noteId]) {
    const formula = document.getLatexFormulas(segmentId)[0], range = formula.getRange()
    assert.equal(document.openLatexPopup({ segmentId, rangeId: formula.getId(), selection: { startOffset: range.startOffset, endOffset: range.endOffset,
      collapsed: false, direction: 'forward', segmentId, segmentPage: 0, isActive: true, rangeType: 'TEXT' } }), true)
    const target = calls.at(-1); assert.equal(target.initialLatex, formula.getLatex()); assert.equal(target.rangeId, formula.getId())
    assert.deepEqual(target.host, document.getSegmentHost(segmentId)); assert.equal(target.selection.anchor.pageAffinity, 0); assert.equal(target.selection.focus.pageAffinity, 0)
  }
  assert.deepEqual(document.getInputState(), input); assert.deepEqual(history(studio, document.id), previous); handle.dispose()
})

test('a popup draft enters content only on canonical apply and survives undo/redo and source-preserving save/reopen', () => {
  const { studio, document } = open(), targets = []; document.setInputSelection(selection('p', 1, 3))
  const handle = document.bindLatexPopupEditor({ open: target => { targets.push(target); return true }, close: () => true })
  assert.equal(document.openLatexPopup(), true); assert.equal(targets[0].initialLatex, 'BC'); assert.equal(history(studio, document.id).undo, 0)
  document.executeCommand({ id: 'doc.command.insert-latex-formula', latex: 'x_2', mode: 'inline' }, targets[0].selection)
  const formula = document.getLatexFormulas()[0], formulaId = formula.getId()
  assert.equal(history(studio, document.id).undo, 1); assert.equal(document.undo(), true); assert.equal(document.getLatexFormulas().length, 0)
  assert.equal(document.redo(), true); assert.equal(document.getLatexFormula(formulaId).getLatex(), 'x_2')
  const reopened = open(copy(document.save())).document
  assert.equal(reopened.getLatexFormula(formulaId).getLatex(), 'x_2'); assert.equal(reopened.getBody().dataStream.startsWith('Ax_2D'), true); handle.dispose()
})
