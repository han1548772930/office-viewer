import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createDocumentRibbon, documentCommandFromRibbon } from '../src/components/document/document-ribbon.ts'

test('restart numbering preserves the Office start boundary, including zero, through the Ribbon adapter', () => {
  const id = 'doc.command.restart-list-numbering'
  assert.deepEqual(documentCommandFromRibbon(id, undefined), { id })
  for (const start of [0, 1, 2147483647]) assert.deepEqual(documentCommandFromRibbon(id, { start }), { id, start })
  for (const start of [-1, 1.5, 2147483648, Infinity, NaN, '0', null]) assert.equal(documentCommandFromRibbon(id, { start }), undefined)
  assert.equal(documentCommandFromRibbon(id, []), undefined)
  assert.deepEqual(documentCommandFromRibbon('doc.command.continue-list-numbering', undefined), { id: 'doc.command.continue-list-numbering' })
})

test('numbering popup capabilities come from command context and are kept separate from bullet styles', () => {
  const context = {
    state: { marks: {}, mixedMarks: [], paragraph: {}, hasSelection: false, blockStyle: 'numbered', listPreset: 'ORDER_LIST' },
    canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false, paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: false,
  }
  const items = value => createDocumentRibbon(value).find(tab => tab.id === 'home').groups.flatMap(group => group.items)
  const ready = items({ ...context, numbering: { canRestart: true, canContinue: true } })
  assert.equal(ready.find(item => item.id === 'doc-numbered').popupProps.canRestart, true)
  assert.equal(ready.find(item => item.id === 'doc-numbered').popupProps.canContinue, true)
  assert.equal(ready.find(item => item.id === 'doc-bullet').popupProps.canRestart, false)
  assert.equal(ready.find(item => item.id === 'doc-bullet').popupProps.canContinue, false)
  const unavailable = items({ ...context, numbering: { canRestart: false, canContinue: false } })
  assert.equal(unavailable.find(item => item.id === 'doc-numbered').popupProps.canRestart, false)
  assert.equal(unavailable.find(item => item.id === 'doc-numbered').popupProps.canContinue, false)
})
