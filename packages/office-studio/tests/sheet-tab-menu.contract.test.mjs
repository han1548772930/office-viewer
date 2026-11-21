import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetTabMenu } from '../src/components/spreadsheet/sheet-tab-menu.ts'

test('sheet-tab menu maps worksheet protection and existing tab commands to their target sheet', () => {
  const calls = []
  const menu = createSheetTabMenu('sheet-b', 2, false, (_key, fallback) => fallback, {
    rename: id => calls.push(['rename', id]),
    copy: id => calls.push(['copy', id]),
    protect: id => calls.push(['protect', id]),
    unprotect: id => calls.push(['unprotect', id]),
    setColor: (id, color) => calls.push(['color', id, color]),
    hide: id => calls.push(['hide', id]),
    delete: id => calls.push(['delete', id]),
  })
  assert.ok(menu.items.some(item => item.key === 'protect-sheet'))
  menu.onClick({ key: 'protect-sheet' })
  menu.onClick({ key: 'color:#4472c4' })
  menu.onClick({ key: 'delete' })
  assert.deepEqual(calls, [
    ['protect', 'sheet-b'],
    ['color', 'sheet-b', '#4472c4'],
    ['delete', 'sheet-b'],
  ])

  const protectedMenu = createSheetTabMenu('sheet-c', 2, true, (_key, fallback) => fallback, {
    rename() {}, copy() {}, protect() {}, unprotect: id => calls.push(['unprotect', id]), setColor() {}, hide() {}, delete() {},
  })
  assert.ok(protectedMenu.items.some(item => item.key === 'unprotect-sheet'))
  protectedMenu.onClick({ key: 'unprotect-sheet' })
  assert.deepEqual(calls.at(-1), ['unprotect', 'sheet-c'])

  const onlySheet = createSheetTabMenu('sheet-a', 1, false, (_key, fallback) => fallback, {
    rename() {}, copy() {}, protect() {}, unprotect() {}, setColor() {}, hide() { throw new Error('must be disabled') }, delete() { throw new Error('must be disabled') },
  })
  assert.equal(onlySheet.items.find(item => item.key === 'hide').disabled, true)
  assert.equal(onlySheet.items.find(item => item.key === 'delete').disabled, true)
  onlySheet.onClick({ key: 'hide' })
  onlySheet.onClick({ key: 'delete' })
})
