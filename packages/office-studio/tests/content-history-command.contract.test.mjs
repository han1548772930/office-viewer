import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Drafts only, 2026-11-22 Asia/Shanghai. Validation is deferred by the user.
let server, api, createWorkspaceHistoryOperations
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  ;({ createWorkspaceHistoryOperations } = await server.ssrLoadModule('/examples/workbench/src/workspace-history-operations.ts'))
})
afterAll(async () => { await server?.close() })

function runtime(t) {
  const studio = new api.OfficeStudio(), office = api.FOffice.newAPI(studio)
  t.onTestFinished(() => studio.dispose())
  const book = office.getActiveWorkbook(), doc = office.getActiveDocument()
  book.getActiveSheet().getRange('A1').setInputText('sheet saved')
  doc.insertText(0, 'doc saved')
  studio.markSaved('workbook', book.id)
  studio.markSaved('document', doc.id)
  book.getActiveSheet().getRange('A1').setInputText('sheet changed')
  doc.insertText(3, ' changed')
  return { studio, office, book, doc }
}

function snapshot(studio, book, doc) {
  return {
    workbook: api.cloneModel(studio.getWorkbookModel(book.id)), document: api.cloneModel(doc.getModel()),
    workbookSavepoint: studio.editor.history.captureSavepoint(book.id),
    documentSavepoint: studio.editor.history.captureSavepoint(doc.id),
    stacks: [book.id, doc.id].map(id => {
      const unit = studio.editor.history.units[id]
      return { undo: unit.undo.length, redo: unit.redo.length, savedRevision: unit.savedRevision, dirty: studio.editor.history.isDirty(id) }
    }),
    selection: studio.getSelectionState(book.id, book.getActiveSheet().getSheetId()),
    input: studio.getDocumentInputState(doc.id),
  }
}

const commandPairs = [
  ['core.command.undo', 'core.command.redo'],
  ['univer.command.undo', 'univer.command.redo'],
  ['sheet.command.undo', 'sheet.command.redo'],
  ['doc.command.undo', 'doc.command.redo'],
]

test('history command aliases replay the same unit stack and publish only history events', t => {
  for (const [undo, redo] of commandPairs) {
    const { studio, office, book, doc } = runtime(t)
    const kind = undo.startsWith('doc.') ? 'document' : 'workbook', target = kind === 'workbook' ? book : doc
    // Product aliases resolve their product even when the other product is focused.
    studio.setActiveKind(undo.startsWith('sheet.') ? 'document' : undo.startsWith('doc.') ? 'workbook' : kind)
    const generic = [], observed = []
    office.addEvent(office.Event.BeforeCommandExecute, event => generic.push(event.id))
    office.addEvent(office.Event.CommandExecuted, event => generic.push(event.id))
    for (const name of ['BeforeUndo', 'Undo', 'BeforeRedo', 'Redo']) office.addEvent(office.Event[name], event => observed.push([name, event.unitId, event.kind]))
    const neighbor = api.cloneModel(kind === 'workbook' ? doc.getModel() : studio.getWorkbookModel(book.id))
    assert.equal(office.executeCommandSync(undo, undefined), true)
    assert.equal(studio.editor.state.activeKind, kind)
    assert.equal(studio.editor.history.isDirty(target.id), false)
    assert.equal(office.executeCommandSync(redo, undefined), true)
    assert.equal(studio.editor.history.isDirty(target.id), true)
    assert.deepEqual(kind === 'workbook' ? doc.getModel() : studio.getWorkbookModel(book.id), neighbor)
    assert.deepEqual(observed, ['BeforeUndo', 'Undo', 'BeforeRedo', 'Redo'].map(name => [name, target.id, kind]))
    assert.deepEqual(generic, [])
  }
})

test('global facade history returns actual async results while unit facade return types are retained', async t => {
  const { studio, office, book, doc } = runtime(t)
  studio.setActiveKind('workbook')
  const undo = office.undo()
  assert.ok(undo instanceof Promise)
  assert.equal(await undo, true)
  assert.equal(await office.redo(), true)
  assert.equal(book.undo(), book)
  assert.equal(book.redo(), book)
  studio.setActiveKind('document')
  assert.equal(await office.undo(), true)
  assert.equal(await office.redo(), true)
  assert.equal(doc.undo(), true)
  assert.equal(doc.redo(), true)
  const empty = office.createDocument({ name: 'empty history' })
  assert.equal(await office.undo(), false)
  assert.equal(await office.redo(), false)
  assert.equal(empty.getText(), '')
})

test('global facade resolves a deliberate cancellation as false and rejects an observer failure', async t => {
  for (const direction of ['undo', 'redo']) {
    const { studio, office, book, doc } = runtime(t)
    studio.setActiveKind('document')
    if (direction === 'redo') studio.undo('document', doc.id)
    const before = snapshot(studio, book, doc)
    const cancel = office.addEvent(office.Event[direction === 'undo' ? 'BeforeUndo' : 'BeforeRedo'], event => { event.cancel = true })
    assert.equal(await office[direction](), false)
    assert.deepEqual(snapshot(studio, book, doc), before)
    cancel.dispose()
    const failed = office.addEvent(office.Event[direction === 'undo' ? 'BeforeUndo' : 'BeforeRedo'], () => { throw new Error('observer failed') })
    await assert.rejects(office[direction](), /observer failed/)
    assert.deepEqual(snapshot(studio, book, doc), before)
    failed.dispose()
  }
})

test('invalid history requests and ambiguous ids do not activate a unit or modify history', t => {
  const { studio, office, book, doc } = runtime(t)
  studio.setActiveKind('workbook')
  const before = snapshot(studio, book, doc)
  for (const input of [null, false, 1, 'workbook', [], { unitId: '' }, { unitId: 42 }, { kind: 'sheet' }, { kind: 'document', unitId: book.id }]) {
    assert.throws(() => office.executeCommandSync('core.command.undo', input))
    assert.equal(studio.editor.state.activeKind, 'workbook')
    assert.deepEqual(snapshot(studio, book, doc), before)
  }
  assert.throws(() => office.executeCommandSync('sheet.command.undo', { kind: 'document', unitId: doc.id }), /类型/)
  assert.throws(() => office.executeCommandSync('doc.command.redo', { kind: 'workbook', unitId: book.id }), /类型/)
  assert.equal(api.contentHistoryCommandAction('toString'), undefined)
  const active = { kind: 'workbook', workbookId: book.id, documentId: doc.id }
  for (const input of [{ unitId: book.id }, { unitId: book.id, kind: 'workbook' }, undefined]) {
    assert.throws(() => api.resolveContentHistoryTarget('core.command.undo', input, active, () => true), /多个类型/)
  }
  assert.deepEqual(snapshot(studio, book, doc), before)
})

test('unit creation rejects cross-product id reuse before sharing populated history or savepoints', t => {
  const { studio, office, book, doc } = runtime(t), before = snapshot(studio, book, doc)
  for (const makeCurrent of [true, false]) {
    assert.throws(() => office.createDocument({ id: book.id, name: 'duplicate doc' }, { makeCurrent }), /不能共用内容历史/)
    assert.throws(() => office.createWorkbook({ id: doc.id, name: 'duplicate book' }, { makeCurrent }), /不能共用内容历史/)
    assert.deepEqual(snapshot(studio, book, doc), before)
  }
  assert.throws(() => new api.OfficeStudio({ workbook: { id: 'shared-id' }, document: { id: 'shared-id' } }), /不能共用内容历史/)
  book.undo()
  assert.deepEqual(doc.getModel(), before.document)
  assert.equal(studio.editor.history.isDirty(book.id), false)
  assert.equal(studio.editor.history.isDirty(doc.id), true)
})

test('cancellation through every registered history namespace preserves models, savepoints and input state', t => {
  for (const [undo, redo] of commandPairs) for (const direction of ['undo', 'redo']) {
    const { studio, office, book, doc } = runtime(t)
    const kind = undo.startsWith('doc.') ? 'document' : 'workbook', id = kind === 'workbook' ? book.id : doc.id
    studio.setActiveKind(kind)
    if (direction === 'redo') studio.undo(kind, id)
    const point = { anchor: { blockId: doc.getModel().blocks[0].id, offset: 1 }, focus: { blockId: doc.getModel().blocks[0].id, offset: 1 } }
    doc.setInputSelection(point)
    doc.executeCommand({ id: 'doc.command.format-text', patch: { fontSize: 27, color: undefined } }, point)
    studio.setActiveKind(kind)
    const before = snapshot(studio, book, doc), completed = []
    office.addEvent(office.Event[direction === 'undo' ? 'BeforeUndo' : 'BeforeRedo'], event => { event.cancel = true })
    office.addEvent(office.Event[direction === 'undo' ? 'Undo' : 'Redo'], event => completed.push(event))
    assert.throws(() => office.executeCommandSync(direction === 'undo' ? undo : redo, { kind, unitId: id }), /canceled/)
    assert.deepEqual(snapshot(studio, book, doc), before)
    assert.deepEqual(completed, [])
  }
})

test('a history namespace command stops if its Before listener switches the active product', t => {
  for (const kind of ['workbook', 'document']) for (const direction of ['undo', 'redo']) {
    const { studio, office, book, doc } = runtime(t)
    const id = kind === 'workbook' ? book.id : doc.id, neighbor = kind === 'workbook' ? 'document' : 'workbook'
    studio.setActiveKind(kind)
    if (direction === 'redo') studio.undo(kind, id)
    const before = snapshot(studio, book, doc), completed = []
    office.addEvent(office.Event[direction === 'undo' ? 'BeforeUndo' : 'BeforeRedo'], () => studio.setActiveKind(neighbor))
    office.addEvent(office.Event[direction === 'undo' ? 'Undo' : 'Redo'], event => completed.push(event))
    assert.equal(office.executeCommandSync(`univer.command.${direction}`, { kind, unitId: id }), false)
    assert.equal(studio.editor.state.activeKind, neighbor)
    const after = snapshot(studio, book, doc)
    // Product navigation may clear transient typing preferences; content and
    // both unit histories still belong to the unexecuted request.
    assert.deepEqual({ ...after, input: undefined }, { ...before, input: undefined })
    assert.deepEqual(completed, [])
  }
})

function workspaceHistory(studio) {
  const errors = []
  const history = createWorkspaceHistoryOperations({
    captureTarget: () => {
      const kind = studio.editor.state.activeKind, model = kind === 'workbook' ? studio.editor.state.workbook : studio.editor.state.document
      return { kind, unitId: model.id }
    },
    replay: (direction, { kind, unitId }) => studio[direction](kind, unitId),
    onError: (direction, error) => errors.push({ direction, message: error.message }),
  })
  return { ...history, errors }
}

test('production history adapter handles cancellation and rejects cached workbench events before replay', t => {
  const { studio, office, book, doc } = runtime(t), history = workspaceHistory(studio)
  studio.setActiveKind('workbook')
  const before = snapshot(studio, book, doc), observed = []
  const cancel = office.addEvent(office.Event.BeforeUndo, event => { observed.push(event.unitId); event.cancel = true })
  assert.equal(history.undo('document', doc.id), false)
  assert.equal(history.undo('workbook', 'replaced-workbook'), false)
  assert.deepEqual(observed, [])
  assert.equal(history.undo('workbook', book.id), false)
  assert.deepEqual(snapshot(studio, book, doc), before)
  assert.equal(history.errors.length, 1)
  assert.deepEqual(observed, [book.id])
  cancel.dispose()
  // Optional unit payloads retain compatibility with older workbench consumers.
  assert.equal(history.undo(), true)
  assert.equal(studio.editor.history.isDirty(book.id), false)
  assert.equal(history.redo('workbook', book.id), true)
  assert.equal(studio.editor.history.isDirty(book.id), true)
})

test('production adapter keeps a Before listener navigation authoritative for both products and directions', t => {
  for (const kind of ['workbook', 'document']) for (const direction of ['undo', 'redo']) {
    const { studio, office, book, doc } = runtime(t), history = workspaceHistory(studio)
    const id = kind === 'workbook' ? book.id : doc.id, neighbor = kind === 'workbook' ? 'document' : 'workbook'
    studio.setActiveKind(kind)
    if (direction === 'redo') studio.undo(kind, id)
    const before = snapshot(studio, book, doc)
    office.addEvent(office.Event[direction === 'undo' ? 'BeforeUndo' : 'BeforeRedo'], () => studio.setActiveKind(neighbor))
    assert.equal(history[direction](kind, id), false)
    assert.equal(studio.editor.state.activeKind, neighbor)
    const after = snapshot(studio, book, doc)
    assert.deepEqual({ ...after, input: undefined }, { ...before, input: undefined })
    assert.deepEqual(history.errors, [])
  }
})
