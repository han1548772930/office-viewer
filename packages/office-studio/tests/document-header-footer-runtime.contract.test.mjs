import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred real Studio/Facade/owner and scoped-read contracts. No tests were run in this implementation phase.
let server, OfficeStudio, FDocument, domain, documentModel, controller, editors, sections
const studios = new Set(), copy = value => structuredClone(value)
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-header-footer-runtime', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, documentModel, controller, editors, sections] = await Promise.all(['domain/model', 'features/document/model/document-model',
    'features/document/controllers/document-header-footer-editor', 'api/document-header-footer-editors',
    'features/document/services/section-header-footer'].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const point = (id = 'p', offset = 0) => ({ blockId: id, offset })
const selection = (offset = 0) => ({ anchor: point('p', offset), focus: point('p', offset) })
const paragraph = (id, text = '') => ({ ...documentModel.createParagraph(text), id })
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })
function fixture(id = 'header-footer-api') {
  const source = { ...domain.createBlankDocument(), id, modified: false, blocks: [paragraph('p', 'Body'), paragraph('tail', 'Tail')],
    sectionHeaderFooters: { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: 'Header', footer: 'Footer' } } }
  const studio = new OfficeStudio({ document: source, currentUserId: 'viewer' }); studios.add(studio); studio.activateDocument(id)
  const document = new FDocument(studio, id); document.setInputSelection(selection())
  return { studio, document, source }
}
function bind(document, calls = []) {
  const handle = document.bindHeaderFooterEditor({ execute: target => { calls.push(target); return true }, release: () => calls.push('released') })
  return { handle, calls }
}

test('three fixed IDs and both legacy aliases preserve the string edit-area contract', () => {
  assert.deepEqual(controller.DOCUMENT_HEADER_FOOTER_COMMANDS, { open: 'doc.command.open-header-footer-panel',
    close: 'doc.command.close-header-footer', ribbon: 'sidebar.operation.doc-header-footer-panel' })
  assert.deepEqual(controller.resolveDocumentHeaderFooterRequest(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), { action: 'open', editArea: 'HEADER' })
  assert.deepEqual(controller.resolveDocumentHeaderFooterRequest(controller.DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS.open, { editArea: 'FOOTER' }), { action: 'open', editArea: 'FOOTER' })
  assert.deepEqual(controller.resolveDocumentHeaderFooterRequest(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.ribbon, { value: 'close', expectedRevision: 0 }),
    { action: 'ribbon', visible: false, expectedRevision: 0 })
  assert.equal(controller.isDocumentHeaderFooterCommand('doc.operation.header-footer-control'), false)
})

test('an unbound UI declines open, close and ribbon operations without modifying content, input or history', () => {
  const { studio, document } = fixture(), before = copy(document.getModel()), state = copy(document.getInputState())
  const revision = document.getModelRevision(), generation = document.getHistoryGeneration()
  for (const id of [...Object.values(controller.DOCUMENT_HEADER_FOOTER_COMMANDS), ...Object.values(controller.DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS)]) {
    assert.equal(studio.executeCommandSync(id, id === controller.DOCUMENT_HEADER_FOOTER_COMMANDS.ribbon ? { value: 'open' } : undefined), false)
  }
  assert.deepEqual(document.getModel(), before); assert.deepEqual(document.getInputState(), state)
  assert.equal(document.getModelRevision(), revision); assert.equal(document.getHistoryGeneration(), generation)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 })
})

test('public commands, legacy aliases and bound handles share captured model identity and the registered command events', () => {
  const { studio, document } = fixture(), before = copy(document.getModel()), state = copy(document.getInputState()), events = []
  const { handle, calls } = bind(document)
  const listener = studio.events.addEvent('CommandExecuted', event => { if (controller.isDocumentHeaderFooterCommand(event.id)) events.push(event.id) })
  assert.equal(studio.executeCommandSync(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open, undefined), true)
  assert.equal(studio.executeCommandSync(controller.DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS.open, { unitId: document.id, editArea: 'FOOTER' }), true)
  assert.equal(handle.execute(controller.DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS.close), true)
  assert.equal(handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.ribbon, { value: 'close' }), true)
  assert.deepEqual(calls.map(target => target.action), ['open', 'open', 'close', 'ribbon'])
  assert.deepEqual(calls.filter(target => target.action === 'open').map(target => target.editArea), ['HEADER', 'FOOTER'])
  for (const target of calls) {
    assert.equal(target.source, studio.getDocumentModel(document.id)); assert.equal(target.unitId, document.id)
    assert.equal(target.revision, document.getModelRevision()); assert.equal(target.historyGeneration, document.getHistoryGeneration())
    assert.equal(target.inputGeneration, document.getInputState().generation)
  }
  assert.deepEqual(events, [controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open, controller.DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS.open,
    controller.DOCUMENT_HEADER_FOOTER_COMMANDS.close, controller.DOCUMENT_HEADER_FOOTER_COMMANDS.ribbon])
  assert.deepEqual(document.getModel(), before); assert.deepEqual(document.getInputState(), state)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0 }); listener.dispose(); handle.dispose()
})

test('malformed areas, ribbon values, IDs, unknown fields and non-natural guards fail before an owner callback', () => {
  const { studio, document } = fixture(), { handle, calls } = bind(document), commands = controller.DOCUMENT_HEADER_FOOTER_COMMANDS
  for (const raw of [null, [], 0, 'HEADER', { editArea: 1 }, { editArea: 'header' }, { editArea: 'BODY' }, { unitId: '' },
    { expectedRevision: -1 }, { expectedInputGeneration: 0.5 }, { expectedHistoryGeneration: Infinity }, { unexpected: true }]) {
    assert.throws(() => studio.executeCommandSync(commands.open, raw))
  }
  for (const raw of [undefined, {}, { value: '' }, { value: 'invalid' }, { value: true }, { value: 'open', editArea: 'HEADER' }]) {
    assert.throws(() => studio.executeCommandSync(commands.ribbon, raw))
  }
  assert.throws(() => handle.execute(commands.close, { unitId: 'other' }), /不一致/)
  assert.throws(() => studio.executeCommandSync(commands.open, { unitId: 'missing' }), /不存在/)
  assert.equal(calls.length, 0); handle.dispose()
})

test.each(['expectedRevision', 'expectedHistoryGeneration', 'expectedInputGeneration'])('%s rejects a stale request before opening or closing an owner', key => {
  const { studio, document } = fixture(), { handle, calls } = bind(document)
  const value = key === 'expectedRevision' ? document.getModelRevision() : key === 'expectedHistoryGeneration' ? document.getHistoryGeneration() : document.getInputState().generation
  assert.throws(() => handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open, { [key]: value + 1 }), /变化/)
  assert.throws(() => studio.executeCommandSync(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.close, { [key]: value + 1 }), /变化/)
  assert.equal(calls.length, 0); handle.dispose()
})

test('owner rejection and BeforeCommand cancellation preserve input marks and content history', () => {
  const { studio, document } = fixture(), calls = []
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, selection())
  const before = copy(document.getModel()), input = copy(document.getInputState()), previous = history(studio, document.id)
  const handle = document.bindHeaderFooterEditor({ execute: target => { calls.push(target.action); return false }, release() {} })
  assert.equal(handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), false)
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open) event.cancel = true })
  assert.throws(() => handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), /canceled/)
  assert.deepEqual(calls, ['open']); assert.deepEqual(document.getModel(), before); assert.deepEqual(document.getInputState(), input)
  assert.deepEqual(history(studio, document.id), previous); cancel.dispose(); handle.dispose()
})

test('replacement and idempotent disposal release old owners once without removing their replacements', () => {
  const { document } = fixture(), first = bind(document), second = bind(document)
  assert.equal(first.handle.isCurrent(), false); assert.equal(second.handle.isCurrent(), true)
  assert.deepEqual(first.calls, ['released']); first.handle.dispose(); first.handle.dispose()
  assert.equal(first.handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), false)
  assert.equal(second.handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), true)
  assert.equal(second.calls.length, 1); second.handle.dispose(); second.handle.dispose()
  assert.equal(second.handle.isCurrent(), false); assert.deepEqual(second.calls.slice(1), ['released'])
})

test('a reentrant old-owner release cannot overwrite a third owner or dispose its handle', () => {
  const { document } = fixture(), calls = []; let newest
  const first = document.bindHeaderFooterEditor({ execute: () => false, release: () => { newest = document.bindHeaderFooterEditor({
    execute: () => { calls.push('newest'); return true }, release: () => calls.push('newest released') }) } })
  const second = document.bindHeaderFooterEditor({ execute: () => { calls.push('second'); return true }, release: () => calls.push('second released') })
  assert.equal(second.isCurrent(), false); assert.equal(newest.isCurrent(), true); second.dispose(); first.dispose()
  assert.equal(newest.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), true)
  assert.deepEqual(calls, ['second released', 'newest']); newest.dispose()
})

test('closing and reopening the same unit rejects both old handles and pre-captured runtime targets', () => {
  const { studio, document, source } = fixture(), { handle, calls } = bind(document)
  const target = editors.captureDocumentHeaderFooterEditorTarget(studio, controller.resolveDocumentHeaderFooterRequest(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open))
  assert.equal(studio.disposeUnit(document.id), true); assert.deepEqual(calls, ['released']); assert.equal(handle.isCurrent(), false)
  studio.createDocument(copy(source), false); studio.activateDocument(document.id)
  const reopened = new FDocument(studio, document.id); reopened.setInputSelection(selection()); const current = bind(reopened)
  assert.equal(handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), false)
  assert.equal(editors.documentHeaderFooterEditorTargetCurrent(studio, target), false)
  assert.equal(editors.requestDocumentHeaderFooterEditor(studio, target), false)
  assert.equal(current.handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), true); current.handle.dispose()
})

test('history reset invalidates an owner even if its document ID and source model remain unchanged', () => {
  const { studio, document } = fixture(), { handle, calls } = bind(document), source = studio.getDocumentModel(document.id)
  studio.editor.history.clear(document.id)
  assert.equal(studio.getDocumentModel(document.id), source); assert.equal(handle.isCurrent(), false)
  assert.equal(handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), false); assert.deepEqual(calls, [])
  handle.dispose(); assert.deepEqual(calls, ['released'])
})

test.each(['activate', 'reopen', 'edit', 'selection', 'owner'])('a BeforeCommand %s change cannot redirect an implicit public open to a new target', change => {
  const { studio, document, source } = fixture(), first = bind(document), otherSource = { ...copy(source), id: 'other-header-footer' }
  studio.createDocument(otherSource, false); const other = new FDocument(studio, otherSource.id); other.setInputSelection(selection()); const second = bind(other)
  const stop = studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id !== controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open) return
    if (change === 'activate') studio.activateDocument(other.id)
    else if (change === 'edit') document.executeCommand({ id: 'doc.command.insert-text', text: 'Z' }, selection())
    else if (change === 'selection') document.setInputSelection(selection(1))
    else if (change === 'owner') document.bindHeaderFooterEditor({ execute: () => { second.calls.push('replacement opened'); return true }, release() {} })
    else { studio.disposeUnit(document.id); studio.createDocument(copy(source), false); studio.activateDocument(document.id)
      const reopened = new FDocument(studio, document.id); reopened.setInputSelection(selection()); bind(reopened, second.calls) }
  })
  assert.throws(() => studio.executeCommandSync(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open, undefined), /更改|变化/)
  assert.equal(first.calls.some(value => typeof value === 'object'), false); assert.equal(second.calls.length, 0)
  stop.dispose(); first.handle.dispose(); second.handle.dispose()
})

test('registering an initially absent owner in a BeforeCommand hook also invalidates the captured public request', () => {
  const { studio, document } = fixture(), calls = []
  const stop = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open) bind(document, calls) })
  assert.throws(() => studio.executeCommandSync(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open, undefined), /更改|变化/)
  assert.deepEqual(calls, []); stop.dispose()
})

test('separate studios with the same unit ID and separate units in one studio keep independent owners', () => {
  const first = fixture(), second = fixture(), one = bind(first.document), two = bind(second.document)
  assert.equal(one.handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.open), true)
  assert.equal(two.handle.execute(controller.DOCUMENT_HEADER_FOOTER_COMMANDS.close), true)
  assert.equal(one.calls[0].source, first.studio.getDocumentModel(first.document.id)); assert.equal(two.calls[0].source, second.studio.getDocumentModel(second.document.id))
  editors.disposeDocumentHeaderFooterEditors(first.studio, first.document.id)
  assert.equal(one.handle.isCurrent(), false); assert.equal(two.handle.isCurrent(), true); one.handle.dispose(); two.handle.dispose()
})

test('scoped story reads remain available when document Edit and unrelated body View are denied', () => {
  const { document } = fixture(), source = copy(document.getModel())
  source.permissionRules = [{ id: 'read-only', scope: 'document', action: 'edit', effect: 'deny', userIds: ['viewer'] },
    { id: 'body-private', scope: 'paragraph', targetId: 'p', action: 'view', effect: 'deny', userIds: ['viewer'] }]
  assert.deepEqual(controller.documentHeaderFooterStoryViewState(source, 'first', 'header', { documentId: source.id, userId: 'viewer' }), { available: true })
})

test.each([['document', undefined], ['section', 'first'], ['entity', 'first:header'], ['paragraph', 'header-p'], ['entity', 'header-image']])('%s View denial blocks the real header story rather than only disabling edit controls', (scope, targetId) => {
  const { document } = fixture(), source = copy(document.getModel())
  source.sectionHeaderFooters.first.storyContent = { header: { runs: [], paragraphs: [], blocks: [
    paragraph('header-p', 'Header'), { id: 'header-image', type: 'image', runs: [], metadata: { image: { src: 'data:image/png;base64,AAAA', width: 10, height: 10 } } },
  ] } }
  source.permissionRules = [{ id: 'denied', scope, ...(targetId ? { targetId } : {}), action: 'view', effect: 'deny', userIds: ['viewer'] }]
  assert.equal(controller.documentHeaderFooterStoryViewState(source, 'first', 'header', { documentId: source.id, userId: 'viewer' }).available, false)
})

test('linked headers check the previous story source as well as their displayed section scope', () => {
  const { document } = fixture(), source = copy(document.getModel())
  source.blocks.push({ id: 'section-b', type: 'sectionBreak', runs: [], metadata: { sectionBreak: { type: 'nextPage', sectionId: 'second' } } }, paragraph('after', 'After'))
  source.sectionHeaderFooters.second = { ...sections.EMPTY_SECTION_HEADER_FOOTER, linkToPrevious: true }
  source.permissionRules = [{ id: 'previous-header-denied', scope: 'entity', targetId: 'first:header', action: 'view', effect: 'deny', userIds: ['viewer'] }]
  assert.equal(controller.documentHeaderFooterStoryViewState(source, 'second', 'header', { documentId: source.id, userId: 'viewer' }).available, false)
})
