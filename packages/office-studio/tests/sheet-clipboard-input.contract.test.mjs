import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeSheetClipboardContentAliases } from '../src/features/sheet-workbench/sheet-clipboard-input.ts'

// Draft only: written for the later unified verification pass; not run in this batch.
test('canonical and fixed Univer content aliases normalize to the same explicit strings', () => {
  const html = '<table><tr><td>x</td></tr></table>'
  assert.deepEqual(normalizeSheetClipboardContentAliases({ textContent: 'x', htmlContent: html }), { text: 'x', html })
  assert.deepEqual(normalizeSheetClipboardContentAliases({ text: 'x', textContent: 'x', html, htmlContent: html }), { text: 'x', html })
})

test('omitted content stays omitted while empty content and meaningful whitespace stay explicit', () => {
  assert.deepEqual(normalizeSheetClipboardContentAliases({}), {})
  assert.deepEqual(normalizeSheetClipboardContentAliases({ textContent: '', htmlContent: '' }), { text: '', html: '' })
  assert.deepEqual(normalizeSheetClipboardContentAliases({ text: undefined, textContent: '\t \n', html: '' }), { text: '\t \n', html: '' })
})

test('conflicting content aliases reject rather than choose one or discard an explicit empty string', () => {
  for (const input of [{ text: '', textContent: 'old' }, { text: 'x', textContent: ' x' },
    { html: '<b>x</b>', htmlContent: '<i>x</i>' }, { html: '', htmlContent: '<br>' }]) {
    assert.throws(() => normalizeSheetClipboardContentAliases(input), TypeError)
  }
})

test('invalid content types and non-record inputs cannot become clipboard text', () => {
  for (const input of [null, undefined, [], 'x', { text: 1 }, { textContent: null }, { html: false }, { htmlContent: ['x'] }]) {
    assert.throws(() => normalizeSheetClipboardContentAliases(input), TypeError)
  }
})
