import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred real Studio/Facade/content-history/snapshot contracts; not run during implementation.
let server, OfficeStudio, FDocument, domain, model, fields, layoutEngine, pagination
const studios = new Set(), copy = value => structuredClone(value)
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-toc-insertion-runtime', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, model, fields] = await Promise.all(['domain/model', 'features/document/model/document-model',
    'features/document/model/document-text-fields'].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
  ;[layoutEngine, pagination] = await Promise.all(['features/document/layout/document-layout', 'api/document-toc-pagination']
    .map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const selection = (offset = 0) => ({ anchor: { blockId: 'body', offset }, focus: { blockId: 'body', offset } })
const insert = patch => ({ id: 'doc.command.insert-table-of-contents', pageNumbers: { heading: 7 }, ...patch })
function fixture(id = 'toc-runtime') {
  const source = domain.createBlankDocument()
  source.id = id
  source.blocks = [{ ...model.createParagraph('Body text'), id: 'body' },
    { ...model.createParagraph('Heading text'), id: 'heading', type: 'heading', level: 1 }]
  const studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(id)
  const document = new FDocument(studio, id)
  document.setInputSelection(selection(4))
  return { studio, document }
}
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0,
  redo: studio.editor.history.units[id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) })

// A renderer protocol fixture: geometry uses the real layout engine with a
// deterministic measurer; page labels may be varied to exercise nonconvergence.
function bindLayout(studio, document, options = {}) {
  let snapshot, generation = 0, fontVersion = 0
  const listeners = new Set()
  const publish = (patch = {}) => {
    const source = document.getModel()
    const layout = layoutEngine.layoutDocument(source, { measure: text => text.length * 7 })
    const label = options.pageNumber?.(generation + 1)
    if (label !== undefined) layout.pages = layout.pages.map(page => ({ ...page, pageNumber: label }))
    snapshot = { source, unitId: document.id, modelRevision: document.getModelRevision(), historyGeneration: document.getHistoryGeneration(),
      layoutGeneration: ++generation, fontVersion, complete: true, layout, engine: 'main-thread', ...patch }
    for (const listener of listeners) listener()
    return snapshot
  }
  const handle = document.bindTocLayout({ read: () => snapshot, subscribe: listener => { listeners.add(listener); return () => listeners.delete(listener) } })
  const subscription = studio.events.addEvent('CommandExecuted', event => {
    if (event.unitId === document.id && ['doc.command.insert-table-of-contents', 'doc.command.update-table-of-contents'].includes(event.id)
      && options.autoPublish !== false) queueMicrotask(() => { if (!studio.isDisposed && handle.isCurrent()) publish() })
  })
  publish({ complete: options.complete !== false })
  return { publish, invalidateFont: () => { fontVersion++; publish({ complete: false }) }, read: () => snapshot,
    listenerCount: () => listeners.size, dispose: () => { subscription.dispose(); handle.dispose() } }
}

test('registered insertion and Facade execution produce one isolated content transaction, undo, redo and public snapshot reopening', () => {
  const { studio, document } = fixture(), before = copy(document.getModel()), events = []
  const listener = studio.events.addEvent('CommandExecuted', event => { if (event.id === insert().id) events.push(event.id) })
  const result = document.executeCommand(insert({ title: '目录' }), selection(4))
  assert.equal(result.changed, true)
  const after = copy(document.getModel())
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  assert.deepEqual(events, ['doc.command.insert-table-of-contents'])
  assert.equal(fields.documentTextFieldRanges(after).filter(range => range.field.field === 'TOC').length, 1)
  assert.equal(fields.documentTextFieldRanges(after).find(range => range.field.field === 'PAGEREF').field.cachedResult, '7')
  document.undo()
  assert.deepEqual(document.getModel(), before)
  document.redo()
  assert.deepEqual(document.getModel(), after)
  const saved = document.save(), reopenedStudio = new OfficeStudio({ document: saved, currentUserId: 'editor' })
  studios.add(reopenedStudio)
  const reopened = new FDocument(reopenedStudio, saved.id).getModel()
  assert.ok(reopened.blocks.some(block => model.runText(block.runs) === 'Heading text\t7'))
  assert.equal(fields.documentTextFieldRanges(reopened).filter(range => range.field.field === 'TOC').length, 1)
  assert.equal(fields.documentTextFieldRanges(reopened).find(range => range.field.field === 'PAGEREF').field.cachedResult, '7')
  listener.dispose()
})

test('BeforeCommand cancellation and navigation before insertion preserve the source, input and history', () => {
  const { studio, document } = fixture(), before = copy(document.getModel()), input = copy(document.getInputState())
  const cancellation = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === insert().id) event.cancel = true })
  assert.throws(() => document.executeCommand(insert(), selection(4)), /canceled/)
  assert.deepEqual(document.getInputState(), input)
  assert.deepEqual(document.getModel(), before)
  cancellation.dispose()
  const navigation = studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === insert().id) document.setInputSelection(selection(1))
  })
  assert.throws(() => document.executeCommand(insert(), selection(4)))
  assert.equal(document.getInputState().location.selection.focus.offset, 1)
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  navigation.dispose()
})

test('a permission denial or no headings rejects the public command without adding an empty undo entry', () => {
  const { studio, document } = fixture()
  document.executeCommand({ id: 'doc.command.set-document-permission-rules', rules: [{ id: 'deny', scope: 'document', action: 'edit',
    effect: 'deny', userIds: [] }] }, selection())
  const before = copy(document.getModel()), previousHistory = history(studio, document.id)
  assert.throws(() => document.executeCommand(insert(), selection(4)))
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), previousHistory)
  const emptySource = domain.createBlankDocument(); emptySource.id = 'empty-toc'; emptySource.blocks[0].id = 'body'
  const emptyStudio = new OfficeStudio({ document: emptySource, currentUserId: 'editor' }); studios.add(emptyStudio)
  const empty = new FDocument(emptyStudio, emptySource.id), original = copy(empty.getModel())
  assert.throws(() => empty.executeCommand(insert(), selection()), /没有.*标题/)
  assert.deepEqual(empty.getModel(), original)
  assert.deepEqual(history(emptyStudio, empty.id), { undo: 0, redo: 0, dirty: false })
})

test('another document and its content history remain unchanged by TOC insertion', () => {
  const { studio, document } = fixture(), other = domain.createBlankDocument()
  other.id = 'toc-other'
  studio.createDocument(other)
  const before = copy(studio.getDocumentModel(other.id))
  document.executeCommand(insert(), selection(4))
  assert.deepEqual(studio.getDocumentModel(other.id), before)
  assert.deepEqual(history(studio, other.id), { undo: 0, redo: 0, dirty: false })
  assert.equal(history(studio, document.id).undo, 1)
})

test('the fixed public flat options await a renderer and resolve the live body caret through the canonical transaction', async () => {
  const { studio, document } = fixture(), before = copy(document.getModel())
  const renderer = bindLayout(studio, document, { pageNumber: () => 9 })
  const input = document.getInputState()
  const result = await studio.executeCommand('doc.command.insert-table-of-contents', { unitId: document.id, levels: 2,
    title: ' Contents ', pageNumbers: { heading: 9 }, expectedRevision: studio.getDocumentRevision(document.id),
    expectedHistoryGeneration: studio.editor.history.captureSavepoint(document.id).generation, inputGeneration: input.generation })
  assert.equal(result, true)
  assert.equal(model.runText(document.getModel().blocks[0].runs), 'Body')
  assert.ok(document.getModel().blocks.some(block => model.runText(block.runs) === 'Contents'))
  assert.equal(fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'PAGEREF').field.cachedResult, '9')
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  document.undo(); assert.deepEqual(document.getModel(), before)
  renderer.dispose()
})

test('public insertion requires real layout and cannot accept supplied or unavailable page labels as a renderer', async () => {
  const { studio, document } = fixture()
  const before = copy(document.getModel())
  assert.equal(await studio.executeCommand('doc.command.insert-table-of-contents'), false)
  assert.deepEqual(document.getModel(), before)
  assert.throws(() => studio.executeCommandSync('doc.command.insert-table-of-contents', { pageNumbers: { heading: 7 } }), /异步/)
  const renderer = bindLayout(studio, document, { pageNumber: () => 8 })
  assert.equal(await studio.executeCommand('doc.command.insert-table-of-contents'), true)
  assert.equal(fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'PAGEREF').field.cachedResult, '8')
  renderer.dispose()
})

test('flat insertion cannot manufacture an input caret on a newly activated document', async () => {
  const source = domain.createBlankDocument(); source.id = 'toc-no-input'
  source.blocks = [{ ...model.createParagraph('Body'), id: 'body' },
    { ...model.createParagraph('Heading'), id: 'heading', type: 'heading', level: 1 }]
  const studio = new OfficeStudio({ document: source }); studios.add(studio); studio.activateDocument(source.id)
  const before = copy(studio.getDocumentModel(source.id))
  assert.equal(await studio.executeCommand('doc.command.insert-table-of-contents', { unitId: source.id }), false)
  assert.deepEqual(studio.getDocumentModel(source.id), before)
  assert.equal(studio.getDocumentInputState(source.id), undefined)
  assert.deepEqual(history(studio, source.id), { undo: 0, redo: 0, dirty: false })
})

test.each(['selection', 'reopen', 'activate'])('a BeforeCommand %s change rejects a pre-captured public flat insertion', async change => {
  const { studio, document } = fixture(), before = copy(document.getModel()), other = domain.createBlankDocument()
  other.id = 'flat-toc-other'; studio.createDocument(other, false)
  const listener = studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id !== 'doc.command.insert-table-of-contents') return
    listener.dispose()
    if (change === 'selection') document.setInputSelection(selection(1))
    else if (change === 'activate') studio.activateDocument(other.id)
    else {
      studio.disposeUnit(document.id); studio.createDocument(copy(before), false); studio.activateDocument(document.id)
      new FDocument(studio, document.id).setInputSelection(selection(4))
    }
  })
  await assert.rejects(studio.executeCommand('doc.command.insert-table-of-contents', { pageNumbers: { heading: 7 } }))
  assert.deepEqual(studio.getDocumentModel(document.id), before)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  assert.equal(fields.documentTextFieldRanges(studio.getDocumentModel(document.id)).filter(range => range.field.field === 'TOC').length, 0)
  if (change === 'selection') assert.equal(document.getInputState().location.selection.focus.offset, 1)
  if (change === 'activate') assert.equal(studio.activeDocumentId, other.id)
})

test.each(['expectedRevision', 'expectedHistoryGeneration', 'inputGeneration'])('a stale flat %s guard preserves model and history', async key => {
  const { studio, document } = fixture(), before = copy(document.getModel())
  const renderer = bindLayout(studio, document)
  const current = key === 'expectedRevision' ? studio.getDocumentRevision(document.id)
    : key === 'expectedHistoryGeneration' ? studio.editor.history.captureSavepoint(document.id).generation : document.getInputState().generation
  await assert.rejects(studio.executeCommand('doc.command.insert-table-of-contents', { [key]: current + 1 }))
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  renderer.dispose()
})

test('post-insertion pagination uses real complete layout, preserves the caret and amends one undo/save/reopen transaction', async () => {
  const { studio, document } = fixture(), before = copy(document.getModel())
  const renderer = bindLayout(studio, document)
  const result = document.executeCommand(insert({ pageNumbers: { heading: 7 }, title: '目录' }), selection(4))
  const caret = copy(result.selection)
  const settled = await pagination.waitForDocumentTocPagination(studio, document.id)
  assert.equal(settled.status, 'settled')
  assert.ok(settled.iterations > 0 && settled.iterations <= 3)
  assert.deepEqual(document.getInputState().location.selection, caret)
  assert.equal(fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'PAGEREF').field.cachedResult, '1')
  assert.equal(renderer.listenerCount(), 0)
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  const after = copy(document.getModel())
  document.undo(); assert.deepEqual(document.getModel(), before)
  document.redo(); assert.deepEqual(document.getModel(), after)
  const snapshot = document.save(), reopened = new OfficeStudio({ document: snapshot, currentUserId: 'editor' }); studios.add(reopened)
  assert.equal(fields.documentTextFieldRanges(reopened.getDocumentModel(document.id)).find(range => range.field.field === 'PAGEREF').field.cachedResult, '1')
  renderer.dispose()
})

test('an incomplete layout or stale revision does not publish TOC content before initial public insertion', async () => {
  const { studio, document } = fixture(), before = copy(document.getModel())
  const renderer = bindLayout(studio, document, { complete: false })
  const insertion = studio.executeCommand('doc.command.insert-table-of-contents', { unitId: document.id })
  await Promise.resolve(); await Promise.resolve()
  assert.deepEqual(document.getModel(), before)
  renderer.publish({ modelRevision: document.getModelRevision() + 1 })
  await Promise.resolve()
  assert.deepEqual(document.getModel(), before)
  assert.equal(history(studio, document.id).undo, 0)
  renderer.publish()
  assert.equal(await insertion, true)
  assert.equal(history(studio, document.id).undo, 1)
  renderer.dispose()
})

test.each(['selection', 'save', 'undo', 'owner'])('a %s boundary cancels pending page refresh without changing a later transaction', async change => {
  const { studio, document } = fixture(), before = copy(document.getModel())
  const renderer = bindLayout(studio, document, { autoPublish: false })
  document.executeCommand(insert(), selection(4))
  const insertion = copy(document.getModel())
  const pending = pagination.waitForDocumentTocPagination(studio, document.id)
  await Promise.resolve()
  if (change === 'selection') document.setInputSelection({ anchor: { blockId: 'heading', offset: 0 }, focus: { blockId: 'heading', offset: 0 } })
  else if (change === 'save') document.save()
  else if (change === 'undo') document.undo()
  else renderer.dispose()
  if (change !== 'owner') renderer.publish()
  assert.equal((await pending).status, 'canceled')
  assert.deepEqual(document.getModel(), change === 'undo' ? before : change === 'save' ? { ...insertion, modified: false } : insertion)
  assert.equal(renderer.listenerCount(), 0)
  if (change !== 'owner') renderer.dispose()
})

test('layout/font invalidation inside a command hook refuses the prepared page amendment', async () => {
  const { studio, document } = fixture()
  const renderer = bindLayout(studio, document)
  const hook = studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === 'doc.command.update-table-of-contents') { renderer.invalidateFont(); hook.dispose() }
  })
  document.executeCommand(insert(), selection(4))
  const inserted = copy(document.getModel())
  const result = await pagination.waitForDocumentTocPagination(studio, document.id)
  assert.equal(result.status, 'rejected')
  assert.deepEqual(document.getModel(), inserted)
  assert.equal(history(studio, document.id).undo, 1)
  assert.equal(renderer.listenerCount(), 0)
  renderer.dispose()
})

test('nonconverging page labels stop after three amendments and retain one history entry', async () => {
  const { studio, document } = fixture(), renderer = bindLayout(studio, document, { pageNumber: generation => generation % 2 ? 9 : 10 })
  const before = copy(document.getModel())
  document.executeCommand(insert({ pageNumbers: { heading: 7 } }), selection(4))
  const result = await pagination.waitForDocumentTocPagination(studio, document.id)
  assert.equal(result.status, 'limit')
  assert.equal(result.iterations, 3)
  assert.equal(history(studio, document.id).undo, 1)
  const after = copy(document.getModel())
  document.undo(); assert.deepEqual(document.getModel(), before)
  document.redo(); assert.deepEqual(document.getModel(), after)
  renderer.dispose()
})

test('disposing and reopening the same ID releases the old layout wait and cannot amend the new history', async () => {
  const { studio, document } = fixture(), source = copy(document.getModel()), renderer = bindLayout(studio, document, { autoPublish: false })
  document.executeCommand(insert(), selection(4))
  const pending = pagination.waitForDocumentTocPagination(studio, document.id)
  await Promise.resolve()
  studio.disposeUnit(document.id); studio.createDocument(source); studio.activateDocument(document.id)
  assert.equal((await pending).status, 'canceled')
  assert.deepEqual(studio.getDocumentModel(document.id), source)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  assert.equal(renderer.listenerCount(), 0)
  renderer.dispose()
})

test.each(['entireTable', 'pageNumbersOnly'])('public %s update locates a field range outside the active caret, preserves selection and shares one history entry', async mode => {
  const { studio, document } = fixture()
  document.executeCommand(insert(), selection(4))
  const toc = fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'TOC')
  document.executeCommand({ id: 'doc.command.insert-text', text: 'New ' }, {
    anchor: { blockId: 'heading', offset: 0 }, focus: { blockId: 'heading', offset: 0 },
  })
  document.setInputSelection(selection(1))
  const before = copy(document.getModel()), caret = copy(document.getInputState().location), oldDepth = history(studio, document.id).undo
  const renderer = bindLayout(studio, document)
  assert.equal(await studio.executeCommand('doc.command.update-table-of-contents', { unitId: document.id, rangeId: toc.field.id, mode }), true)
  const settled = await pagination.waitForDocumentTocPagination(studio, document.id)
  assert.equal(settled.status, 'settled')
  assert.ok(settled.iterations >= 1 && settled.iterations <= 3)
  assert.deepEqual(document.getInputState().location, caret)
  assert.equal(history(studio, document.id).undo, oldDepth + 1)
  const page = fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'PAGEREF')
  assert.equal(page.field.cachedResult, '1')
  assert.equal(model.runText(page.block.runs), (mode === 'entireTable' ? 'New ' : '') + 'Heading text\t1')
  const after = copy(document.getModel())
  document.undo(); assert.deepEqual(document.getModel(), before)
  document.redo(); assert.deepEqual(document.getModel(), after)
  renderer.dispose()
})

test('an accurate page-only update succeeds without a new history entry or input generation', async () => {
  const { studio, document } = fixture()
  document.executeCommand(insert({ pageNumbers: { heading: 1 } }), selection(4))
  const toc = document.getModel().blocks.find(block => block.metadata.documentToc)
  document.setInputSelection({ anchor: { blockId: toc.id, offset: 0 }, focus: { blockId: toc.id, offset: 0 } })
  const renderer = bindLayout(studio, document), before = copy(document.getModel()), state = copy(document.getInputState()), oldDepth = history(studio, document.id)
  assert.equal(await studio.executeCommand('doc.command.update-table-of-contents', { mode: 'pageNumbersOnly' }), true)
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(document.getInputState(), state)
  assert.deepEqual(history(studio, document.id), oldDepth)
  renderer.dispose()
})

test('explicit-range deletion does not require a renderer and remains one isolated undo/redo transaction', async () => {
  const { studio, document } = fixture()
  document.executeCommand(insert(), selection(4))
  const rangeId = fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'TOC').field.id
  document.setInputSelection(selection(1))
  const before = copy(document.getModel()), caret = copy(document.getInputState().location), oldDepth = history(studio, document.id).undo
  assert.throws(() => studio.executeCommandSync('doc.command.delete-table-of-contents', { rangeId }), /异步/)
  assert.equal(await studio.executeCommand('doc.command.delete-table-of-contents', { rangeId }), true)
  assert.equal(fields.documentTextFieldRanges(document.getModel()).some(range => ['TOC', 'PAGEREF'].includes(range.field.field)), false)
  assert.deepEqual(document.getInputState().location, caret)
  assert.equal(history(studio, document.id).undo, oldDepth + 1)
  const after = copy(document.getModel())
  document.undo(); assert.deepEqual(document.getModel(), before)
  document.redo(); assert.deepEqual(document.getModel(), after)
  assert.equal(await studio.executeCommand('doc.command.delete-table-of-contents', { rangeId }), false)
})

test('manual typed updates keep the fixed three-pass total and amend only their own first update', async () => {
  const { studio, document } = fixture()
  document.executeCommand(insert(), selection(4))
  const toc = document.getModel().blocks.find(block => block.metadata.documentToc), tocId = toc.metadata.documentToc.id
  document.setInputSelection({ anchor: { blockId: toc.id, offset: 0 }, focus: { blockId: toc.id, offset: 0 } })
  const before = copy(document.getModel()), oldDepth = history(studio, document.id).undo
  const renderer = bindLayout(studio, document, { pageNumber: generation => generation % 2 ? 9 : 10 })
  assert.equal(document.executeCommand({ id: 'doc.command.update-table-of-contents', tocId, mode: 'all', pageNumbers: { heading: 8 } },
    document.getInputState().location.selection).changed, true)
  const result = await pagination.waitForDocumentTocPagination(studio, document.id)
  assert.equal(result.status, 'limit')
  assert.equal(result.iterations, 3)
  assert.equal(history(studio, document.id).undo, oldDepth + 1)
  document.undo(); assert.deepEqual(document.getModel(), before)
  assert.equal(fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'PAGEREF').field.cachedResult, '7')
  renderer.dispose()
})

test('an explicit range can update without inventing an input session', async () => {
  const first = fixture()
  first.document.executeCommand(insert(), selection(4))
  const source = copy(first.document.getModel())
  const studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(source.id)
  const document = new FDocument(studio, source.id), renderer = bindLayout(studio, document)
  const rangeId = fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'TOC').field.id
  assert.equal(document.getInputState(), undefined)
  assert.equal(await studio.executeCommand('doc.command.update-table-of-contents', { rangeId }), true)
  assert.equal(document.getInputState(), undefined)
  assert.equal(fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'PAGEREF').field.cachedResult, '1')
  assert.equal(history(studio, document.id).undo, 1)
  renderer.dispose()
})

test.each(['doc.command.update-table-of-contents', 'doc.command.delete-table-of-contents'])('a stale or canceled flat %s cannot mutate the TOC', async commandId => {
  const { studio, document } = fixture()
  document.executeCommand(insert(), selection(4))
  const rangeId = fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'TOC').field.id
  const renderer = bindLayout(studio, document), before = copy(document.getModel()), oldDepth = history(studio, document.id)
  await assert.rejects(studio.executeCommand(commandId, { rangeId, expectedRevision: document.getModelRevision() + 1 }))
  const cancellation = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === commandId) event.cancel = true })
  await assert.rejects(studio.executeCommand(commandId, { rangeId }), /canceled/)
  cancellation.dispose()
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), oldDepth)
  renderer.dispose()
})

test('manual update waits for a matching complete layout and cancels on later navigation', async () => {
  const { studio, document } = fixture()
  document.executeCommand(insert(), selection(4))
  const rangeId = fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'TOC').field.id
  const renderer = bindLayout(studio, document, { complete: false }), before = copy(document.getModel()), oldDepth = history(studio, document.id)
  const pending = studio.executeCommand('doc.command.update-table-of-contents', { rangeId, mode: 'pageNumbersOnly' })
  await Promise.resolve(); await Promise.resolve()
  document.setInputSelection(selection(1))
  assert.equal(await pending, false)
  renderer.publish()
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), oldDepth)
  assert.equal(renderer.listenerCount(), 0)
  renderer.dispose()
})

test.each(['doc.command.update-table-of-contents', 'doc.command.delete-table-of-contents'])('explicit-range %s preserves a separate story input and its typing preferences', async commandId => {
  const { studio, document } = fixture()
  document.executeCommand(insert(), selection(4))
  document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch: { header: 'Header' } }, selection(0))
  const rangeId = fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'TOC').field.id
  const story = { sectionId: 'first', story: 'header', start: 2, end: 2 }
  document.setInputSelection(selection(0), story)
  document.executeCommand({ id: 'doc.command.format-text', story, patch: { bold: true } }, selection(0))
  const before = copy(document.getInputState()), header = copy(document.getModel().sectionHeaderFooters.first)
  const renderer = bindLayout(studio, document)
  assert.equal(await studio.executeCommand(commandId, { rangeId }), true)
  assert.deepEqual(document.getInputState().location, before.location)
  assert.equal(document.getInputState().valid, true)
  assert.equal(document.getInputMarks().bold, true)
  assert.deepEqual(document.getModel().sectionHeaderFooters.first, header)
  renderer.dispose()
})

test.each(['doc.command.update-table-of-contents', 'doc.command.delete-table-of-contents'])('flat %s enforces the target field lock and document permissions', async commandId => {
  const { studio, document } = fixture()
  document.executeCommand(insert(), selection(4))
  const rangeId = fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'TOC').field.id
  studio.commitDocument(document.id, 'Lock TOC fixture', draft => {
    const owner = fields.documentTextFieldRanges(draft).find(range => range.field.id === rangeId)
    fields.setDocumentTextFields(owner.block, fields.documentTextFields(owner.block).map(field => field.id === rangeId ? { ...field, locked: true } : field))
  })
  const renderer = bindLayout(studio, document), before = copy(document.getModel()), oldDepth = history(studio, document.id)
  await assert.rejects(studio.executeCommand(commandId, { rangeId }), /锁定/)
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), oldDepth)
  studio.commitDocument(document.id, 'Unlock TOC fixture', draft => {
    const owner = fields.documentTextFieldRanges(draft).find(range => range.field.id === rangeId)
    fields.setDocumentTextFields(owner.block, fields.documentTextFields(owner.block).map(field => field.id === rangeId ? { ...field, locked: false } : field))
  })
  document.executeCommand({ id: 'doc.command.set-document-permission-rules', rules: [{ id: 'deny-toc', scope: 'document', action: 'edit',
    effect: 'deny', userIds: [] }] }, selection(0))
  renderer.publish()
  const denied = copy(document.getModel()), deniedHistory = history(studio, document.id)
  await assert.rejects(studio.executeCommand(commandId, { rangeId }))
  assert.deepEqual(document.getModel(), denied)
  assert.deepEqual(history(studio, document.id), deniedHistory)
  renderer.dispose()
})

test('flat actions reject invalid modes and missing targets without accepting a caller page map as layout', async () => {
  const { studio, document } = fixture()
  document.executeCommand(insert(), selection(4))
  const rangeId = fields.documentTextFieldRanges(document.getModel()).find(range => range.field.field === 'TOC').field.id
  const before = copy(document.getModel()), oldDepth = history(studio, document.id)
  assert.equal(await studio.executeCommand('doc.command.update-table-of-contents', { rangeId, pageNumbers: { heading: 1 } }), false)
  assert.equal(await studio.executeCommand('doc.command.update-table-of-contents', { rangeId: 'missing' }), false)
  assert.equal(await studio.executeCommand('doc.command.delete-table-of-contents', { rangeId: 'missing' }), false)
  await assert.rejects(studio.executeCommand('doc.command.update-table-of-contents', { rangeId, mode: 'all' }), /模式/)
  await assert.rejects(studio.executeCommand('doc.command.delete-table-of-contents', { rangeId, expectedRevision: null }), /版本/)
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), oldDepth)
})

test('the fixed public dialog operation checks the live TOC/layout/permission and follows replacement owner lifetime', () => {
  const { studio, document } = fixture(), calls = []
  document.executeCommand(insert(), selection(4))
  const block = document.getModel().blocks.find(candidate => candidate.metadata.documentToc)
  const caret = { anchor: { blockId: block.id, offset: 0 }, focus: { blockId: block.id, offset: 0 } }
  document.setInputSelection(caret)
  const renderer = bindLayout(studio, document)
  const first = document.bindTocDialogEditor({ openUpdate: () => { calls.push('first'); return true } })
  const before = copy(document.getModel()), oldDepth = history(studio, document.id)
  assert.equal(document.openTableOfContentsDialog(), true)
  const second = document.bindTocDialogEditor({ openUpdate: () => { calls.push('second'); return true } })
  assert.equal(first.isCurrent(), false)
  first.dispose()
  assert.equal(document.openTableOfContentsDialog(), true)
  assert.deepEqual(calls, ['first', 'second'])
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), oldDepth)
  document.setInputSelection(selection(0))
  assert.equal(document.openTableOfContentsDialog(), false)
  document.setInputSelection(caret)
  document.executeCommand({ id: 'doc.command.set-document-permission-rules', rules: [{ id: 'no-dialog-edit', scope: 'document', action: 'edit',
    effect: 'deny', userIds: [] }] }, caret)
  document.setInputSelection(caret)
  renderer.publish()
  assert.equal(document.openTableOfContentsDialog(), false)
  second.dispose(); renderer.dispose()
})

test('BeforeCommand navigation and unit disposal prevent an old public TOC dialog owner from running', () => {
  const { studio, document } = fixture(), calls = []
  document.executeCommand(insert(), selection(4))
  const block = document.getModel().blocks.find(candidate => candidate.metadata.documentToc)
  document.setInputSelection({ anchor: { blockId: block.id, offset: 0 }, focus: { blockId: block.id, offset: 0 } })
  const renderer = bindLayout(studio, document), owner = document.bindTocDialogEditor({ openUpdate: () => { calls.push('open'); return true } })
  const listener = studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === 'doc.operation.open-table-of-contents-dialog') document.setInputSelection(selection(0))
  })
  assert.throws(() => document.openTableOfContentsDialog(), /已更改/)
  listener.dispose()
  const snapshot = copy(document.getModel())
  studio.disposeUnit(document.id); studio.createDocument(snapshot); studio.activateDocument(document.id)
  assert.equal(owner.isCurrent(), false)
  assert.equal(new FDocument(studio, document.id).openTableOfContentsDialog(), false)
  assert.deepEqual(calls, [])
  owner.dispose(); renderer.dispose()
})
