import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createPinia } from 'pinia'
import { watch } from 'vue'
import { useHistoryStore } from '../src/stores/historyStore.ts'

test('a paste option replaces the top redo while a single undo restores the original content and savepoint', () => {
  const history = useHistoryStore(createPinia())
  history.activate('book')
  const before = history.captureSavepoint('book')
  let cell = 'original'
  try {
    history.execute('book', 'paste all', () => { cell = 'all' }, () => { cell = 'original' })
    const paste = history.captureSavepoint('book'), originalUndo = history.units.book.undo[0].undo
    history.markSaved('book', paste)
    assert.equal(history.amend('book', paste, 'paste values', () => { cell = 'values' }), true)
    const amended = history.captureSavepoint('book')
    assert.equal(history.units.book.undo.length, 1)
    assert.equal(history.units.book.undo[0].undo, originalUndo)
    assert.equal(history.units.book.undo[0].before, before.revision)
    assert.ok(amended.revision > paste.revision)
    assert.equal(history.isDirty('book'), true)
    history.undo(); assert.equal(cell, 'original')
    assert.equal(history.captureSavepoint('book').revision, before.revision)
    history.redo(); assert.equal(cell, 'values')
    assert.deepEqual(history.captureSavepoint('book'), amended)
  } finally { history.$dispose() }
})

test('recent-paste amendment rejects a different unit, an undo branch, intervening edits and a reopened history', () => {
  const history = useHistoryStore(createPinia())
  history.activate('book')
  let calls = 0
  try {
    history.execute('book', 'paste', () => {}, () => {})
    const paste = history.captureSavepoint('book')
    assert.equal(history.amend('other', paste, 'wrong unit', () => { calls++ }), false)
    history.undo()
    assert.equal(history.amend('book', paste, 'undone paste', () => { calls++ }), false)
    history.redo()
    history.execute('book', 'later edit', () => {}, () => {})
    assert.equal(history.amend('book', paste, 'stale paste', () => { calls++ }), false)
    history.clear('book')
    history.execute('book', 'new paste', () => {}, () => {})
    assert.equal(history.amend('book', paste, 'old generation', () => { calls++ }), false)
    assert.equal(calls, 0)
  } finally { history.$dispose() }
})

test('amendment publication observes the final stack and rejects same-unit callback reentrancy', () => {
  const history = useHistoryStore(createPinia())
  history.activate('book')
  let content = 'before'
  history.execute('book', 'paste', () => { content = 'all' }, () => { content = 'before' })
  const paste = history.captureSavepoint('book'), seen = []
  const stop = watch(() => history.isApplying, applying => {
    if (!applying) seen.push({ content, count: history.units.book.undo.length, revision: history.units.book.revision })
  }, { flush: 'sync' })
  try {
    assert.equal(history.amend('book', paste, 'values', () => {
      assert.equal(history.amend('book', paste, 'nested', () => { content = 'bad' }), false)
      assert.throws(() => history.execute('book', 'nested edit', () => { content = 'bad' }, () => {}), /历史事务正在提交/)
      content = 'values'
    }), true)
    assert.deepEqual(seen, [{ content: 'values', count: 1, revision: history.units.book.revision }])
  } finally { stop(); history.$dispose() }
})

test('a refused amendment does not change history or saved revision and releases publication state', () => {
  const history = useHistoryStore(createPinia())
  history.activate('book')
  history.execute('book', 'paste', () => {}, () => {})
  const paste = history.captureSavepoint('book'), entry = history.units.book.undo[0]
  try {
    assert.throws(() => history.amend('book', paste, 'refused', () => { throw new Error('candidate refused') }), /candidate refused/)
    assert.equal(history.isApplying, false)
    assert.equal(history.units.book.undo[0], entry)
    assert.deepEqual(history.captureSavepoint('book'), paste)
    assert.equal(history.canAmend('book', paste), true)
  } finally { history.$dispose() }
})
