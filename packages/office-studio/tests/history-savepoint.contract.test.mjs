import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { createPinia } from 'pinia'
import { createServer } from 'vite'

let server
let useHistoryStore
let CommandBus

before(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ useHistoryStore } = await server.ssrLoadModule('/packages/office-studio/src/stores/historyStore.ts'))
  ;({ CommandBus } = await server.ssrLoadModule('/packages/office-studio/src/services/commandBus.ts'))
})

after(async () => { await server?.close() })

test('savepoints describe the exported revision, including an undo/redo branch', () => {
  const history = useHistoryStore(createPinia())
  history.activate('book')
  let value = 0
  history.execute('book', 'first', () => { value = 1 }, () => { value = 0 })
  const savepoint = history.captureSavepoint('book')
  history.execute('book', 'second', () => { value = 2 }, () => { value = 1 })
  assert.equal(history.markSaved('book', savepoint), true)
  assert.equal(history.isDirty('book'), true, 'finishing an older export must not save newer edits')
  assert.equal(history.undo(), 'second')
  assert.equal(value, 1)
  assert.equal(history.isDirty('book'), false)
  assert.equal(history.redo(), 'second')
  assert.equal(history.isDirty('book'), true)
  history.undo()
  history.execute('book', 'branch', () => { value = 3 }, () => { value = 1 })
  assert.equal(history.canRedo, false)
  history.undo()
  assert.equal(value, 1)
  assert.equal(history.isDirty('book'), false)
})

test('clearing a unit invalidates savepoints from its former history', () => {
  const history = useHistoryStore(createPinia())
  history.activate('reused-id')
  const token = history.captureSavepoint('reused-id')
  history.clear('reused-id')
  history.execute('reused-id', 'new document edit', () => {}, () => {})
  const before = history.units['reused-id'].savedRevision
  assert.equal(history.markSaved('reused-id', token), false)
  assert.equal(history.units['reused-id'].savedRevision, before)
  assert.equal(history.isDirty('reused-id'), true)
})

test('document IDs cannot collide with Object prototype properties', () => {
  const history = useHistoryStore(createPinia())
  for (const id of ['__proto__', 'constructor', 'toString']) {
    history.activate(id)
    let value = 0
    history.execute(id, `edit ${id}`, () => { value = 1 }, () => { value = 0 })
    assert.equal(history.canUndo, true)
    history.undo()
    assert.equal(value, 0)
    assert.equal(history.isDirty(id), false)
  }
})

test('failed undo and redo retain the command for retry', () => {
  const bus = new CommandBus()
  const context = { value: 0, blockUndo: false, blockRedo: false }
  bus.execute({
    id: 'test.command.edit', label: 'edit',
    execute(state) { if (state.blockRedo) throw new Error('redo rejected'); state.value = 1 },
    undo(state) { if (state.blockUndo) throw new Error('undo rejected'); state.value = 0 },
  }, context)
  context.blockUndo = true
  assert.throws(() => bus.undo(context), /undo rejected/)
  assert.equal(bus.canUndo(), true)
  assert.equal(bus.canRedo(), false)
  context.blockUndo = false
  bus.undo(context)
  context.blockRedo = true
  assert.throws(() => bus.redo(context), /redo rejected/)
  assert.equal(bus.canRedo(), true)
  assert.equal(bus.canUndo(), false)
  context.blockRedo = false
  bus.redo(context)
  assert.equal(context.value, 1)
})

test('a failed command leaves revision, savepoint and redo history unchanged', () => {
  const history = useHistoryStore(createPinia())
  history.activate('book')
  history.execute('book', 'accepted', () => {}, () => {})
  history.undo()
  const saved = history.captureSavepoint('book')
  assert.throws(() => history.execute('book', 'rejected', () => { throw new Error('validation failed') }, () => {}), /validation failed/)
  assert.deepEqual(history.captureSavepoint('book'), saved)
  assert.equal(history.canUndo, false)
  assert.equal(history.canRedo, true)
  assert.equal(history.isDirty('book'), false)
})
