import assert from 'node:assert/strict'
import { test } from 'vitest'
import { equivalentModelData } from '../src/platform/commands/model-equivalence.ts'

test('model equivalence ignores key order and absent optional data, but only excludes root state', () => {
  assert.equal(equivalentModelData(
    { modified: true, name: 'book', optional: undefined, metadata: { modified: 'content', value: 1 } },
    { metadata: { value: 1, modified: 'content' }, name: 'book', modified: false },
    ['modified'],
  ), true)
  assert.equal(equivalentModelData(
    { modified: true, metadata: { modified: true } },
    { modified: false, metadata: { modified: false } },
    ['modified'],
  ), false)
})

test('model comparison preserves scalar types, array order and empty values', () => {
  for (const pair of [[0, false], [null, undefined], ['', null], [[1, 2], [2, 1]], [{ value: null }, {}]]) {
    assert.equal(equivalentModelData(...pair), false)
  }
  assert.equal(equivalentModelData({ nested: [{ value: 0 }] }, { nested: [{ value: 0 }] }), true)
})

test('comparing a frozen model does not rewrite its fields', () => {
  const left = Object.freeze({ modified: true, cells: Object.freeze([{ value: 'preserved' }]) })
  const right = Object.freeze({ cells: Object.freeze([{ value: 'preserved' }]), modified: false })
  assert.equal(equivalentModelData(left, right, ['modified']), true)
  assert.equal(left.modified, true)
  assert.equal(right.modified, false)
})
