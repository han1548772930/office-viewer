import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createPinia } from 'pinia'
import { watch } from 'vue'
import { useHistoryStore } from '../src/stores/historyStore.ts'
import { EventHub } from '../src/api/events.ts'

test('history publication resumes after stack and revision updates on execute, undo and redo', () => {
  const history = useHistoryStore(createPinia()), notifications = []
  history.activate('book')
  let content = 'before'
  const stop = watch(() => history.isApplying, applying => {
    if (!applying) notifications.push({ content, undo: history.units.book.undo.length, redo: history.units.book.redo.length, dirty: history.isDirty('book') })
  }, { flush: 'sync' })
  try {
    history.execute('book', 'change', () => { assert.equal(history.isApplying, true); content = 'after' },
      () => { assert.equal(history.isApplying, true); content = 'before' })
    history.undo(); history.redo()
    assert.deepEqual(notifications, [
      { content: 'after', undo: 1, redo: 0, dirty: true },
      { content: 'before', undo: 0, redo: 1, dirty: false },
      { content: 'after', undo: 1, redo: 0, dirty: true },
    ])
  } finally { stop(); history.$dispose() }
})

test('failed history callbacks release the publication scope and do not insert or remove an entry', () => {
  const history = useHistoryStore(createPinia())
  history.activate('book')
  try {
    assert.throws(() => history.execute('book', 'fail', () => { throw new Error('prepare refused') }, () => {}), /prepare refused/)
    assert.equal(history.isApplying, false); assert.equal(history.units.book.undo.length, 0)
    history.execute('book', 'valid', () => {}, () => { throw new Error('replay refused') })
    assert.throws(() => history.undo(), /replay refused/)
    assert.equal(history.isApplying, false); assert.equal(history.units.book.undo.length, 1); assert.equal(history.units.book.redo.length, 0)
  } finally { history.$dispose() }
})

test('same-document reentrant history changes reject before mutation while another document keeps an independent transaction', () => {
  const history = useHistoryStore(createPinia())
  history.activate('first')
  let first = 0, second = 0
  try {
    history.execute('first', 'first edit', () => {
      assert.throws(() => history.execute('first', 'nested first edit', () => { first = 99 }, () => {}), /历史事务正在提交/)
      assert.throws(() => history.clear('first'), /历史事务正在提交/)
      history.execute('second', 'second edit', () => { second = 2 }, () => { second = 0 })
      first = 1
    }, () => {
      assert.throws(() => history.undo(), /历史事务正在提交/)
      first = 0
    })
    assert.equal(first, 1); assert.equal(second, 2)
    assert.equal(history.units.first.undo.length, 1); assert.equal(history.units.second.undo.length, 1)
    history.undo(); assert.equal(first, 0); assert.equal(second, 2)
    history.activate('second'); history.undo(); assert.equal(second, 0)
    assert.equal(history.isApplying, false)
  } finally { history.$dispose() }
})

test('a settled history listener can start a second edit in the same unit with two independent replay steps', () => {
  const history = useHistoryStore(createPinia())
  history.activate('book')
  let content = '', nested = false
  const stop = watch(() => history.isApplying, applying => {
    if (!applying && !nested) {
      nested = true
      assert.equal(history.units.book.undo.length, 1)
      history.execute('book', 'second edit', () => { content = 'AB' }, () => { content = 'A' })
    }
  }, { flush: 'sync' })
  try {
    history.execute('book', 'first edit', () => { content = 'A' }, () => { content = '' })
    assert.equal(content, 'AB'); assert.equal(history.units.book.undo.length, 2); assert.equal(history.isApplying, false)
    history.undo(); assert.equal(content, 'A')
    history.undo(); assert.equal(content, '')
    history.redo(); assert.equal(content, 'A')
    history.redo(); assert.equal(content, 'AB')
  } finally { stop(); history.$dispose() }
})

test('settled observations copy payloads, continue after failures and leave Before dispatch exceptions observable', () => {
  const hub = new EventHub(), seen = [], failure = new Error('observer failed')
  hub.addEvent('changed', event => { event.value = 9; throw failure })
  hub.addEvent('changed', event => seen.push(event.value))
  assert.deepEqual(hub.emitSettled('changed', () => ({ value: 1 })), [{ eventName: 'changed', error: failure }])
  assert.deepEqual(seen, [1])
  hub.addEvent('before', () => { throw failure })
  assert.throws(() => hub.emit('before', { cancel: false }), error => error === failure)
  hub.clear()
})
