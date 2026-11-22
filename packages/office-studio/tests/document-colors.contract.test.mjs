import assert from 'node:assert/strict'
import { test } from 'vitest'
import { DOCUMENT_COLOR_PRESETS, DOCUMENT_STANDARD_COLORS, documentColorDraftIsValid, documentColorIsAutomatic, documentColorParams, documentColorPreview, rememberDocumentColors } from '../src/features/document/services/document-colors.ts'
import { documentCommandFromRibbon, createDocumentRibbon } from '../src/components/document/document-ribbon.ts'
import { createOfficeMessages } from '../src/i18n/office-i18n.ts'

test('color selections preserve explicit clear intent through JSON and reject invalid fields and hex values', () => {
  assert.deepEqual(documentColorParams('color', '#abc'), { patch: { color: '#aabbcc' } })
  assert.equal(documentColorParams('highlight', '#nothex'), undefined)
  const clear = JSON.parse(JSON.stringify(documentColorParams('highlight', null)))
  assert.deepEqual(documentCommandFromRibbon('doc.command.format-text', clear), { id: 'doc.command.format-text', patch: {}, clearFields: ['highlight'] })
  assert.equal(documentCommandFromRibbon('doc.command.format-text', { patch: {}, clearFields: ['fontSize'] }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.format-text', { patch: {}, clearFields: [] }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.format-text', { patch: {}, clearFields: 'highlight' }), undefined)
})

test('automatic color and absent highlight preview concrete document defaults without replacing the last applied color', () => {
  assert.equal(documentColorIsAutomatic('color', 'auto'), true)
  assert.equal(documentColorIsAutomatic('highlight', undefined), true)
  assert.equal(documentColorIsAutomatic('highlight', '#ffffff'), false)
  assert.equal(documentColorPreview('color', 'auto', '#1f2937'), '#1f2937')
  assert.equal(documentColorPreview('highlight', undefined, '#1677ff'), '#1677ff')
  assert.equal(documentColorPreview('color', 'rgb(10, 20, 30)', '#1f2937'), 'rgb(10, 20, 30)')
  const previous = { color: '#334455', highlight: '#ccddff' }
  assert.deepEqual(rememberDocumentColors(previous, { id: 'doc.command.format-text', patch: { color: 'auto', highlight: '' } }), previous)
  assert.deepEqual(rememberDocumentColors(previous, { id: 'doc.command.format-text', patch: { color: '#aaaaaa' }, clearFields: ['color'] }), previous)
})

test('a partial HEX or incomplete channel draft cannot silently apply the previously valid color', () => {
  assert.equal(documentColorDraftIsValid('hex', ['123456']), true)
  assert.equal(documentColorDraftIsValid('hex', ['#AbCdEf']), true)
  for (const input of ['', '12', 'abc', '12345g', '123456ff']) assert.equal(documentColorDraftIsValid('hex', [input]), false)
  assert.equal(documentColorDraftIsValid('rgb', ['0', '128', '255']), true)
  assert.equal(documentColorDraftIsValid('rgb', ['0', '', '255']), false)
  assert.equal(documentColorDraftIsValid('rgb', ['0', '12.5', '255']), false)
  assert.equal(documentColorDraftIsValid('rgb', ['256', '0', '0']), false)
  assert.equal(documentColorDraftIsValid('hsb', ['360', '12.5%', '100%']), true)
  assert.equal(documentColorDraftIsValid('hsb', ['0', '101%', '100%']), false)
  assert.equal(documentColorDraftIsValid('hsb', ['0', '100%']), false)
})

test('main color commands retain the last executed color while the panel derives the current and mixed selection', () => {
  const memory = rememberDocumentColors({}, { id: 'doc.command.format-text', patch: { color: '#334455', highlight: '#ccddff' } })
  const state = { marks: { color: '#aa0000' }, mixedMarks: ['highlight'], paragraph: {}, hasSelection: true }
  const context = { state, canUndo: true, canRedo: false, outline: false, focus: false, fullscreen: false, paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: false,
    lastTextColor: memory.color, lastHighlightColor: memory.highlight, defaultTextColor: '#1f2937', defaultHighlightColor: '#1677ff' }
  const items = createDocumentRibbon(context).find(tab => tab.id === 'home').groups.flatMap(group => group.items)
  const text = items.find(item => item.id === 'doc-text-color')
  const fill = items.find(item => item.id === 'doc-highlight')
  assert.equal(text.type, 'split')
  assert.equal(text.gridLayout.column, 8)
  assert.equal(fill.gridLayout.column, 7)
  assert.deepEqual(text.params, { patch: { color: '#334455' } })
  assert.equal(text.popupProps.selectedColor, '#aa0000')
  assert.equal(fill.popupProps.mixed, true)
  assert.deepEqual(DOCUMENT_COLOR_PRESETS.map(row => row.length), [9, 9, 9, 9, 9])
  assert.equal(new Set(DOCUMENT_STANDARD_COLORS).size, 45)
  assert.deepEqual(text.children.filter(item => item.id.includes('-preset-')).map(item => item.params.patch.color), DOCUMENT_STANDARD_COLORS)
  assert.deepEqual(fill.children.filter(item => item.id.includes('-preset-')).map(item => item.params.patch.highlight), DOCUMENT_STANDARD_COLORS)
  assert.deepEqual(fill.children.find(item => item.id === 'doc-highlight-reset').params, { patch: { highlight: '#1677ff' } })
  assert.deepEqual(fill.children.find(item => item.id === 'doc-highlight-clear').params, { patch: {}, clearFields: ['highlight'] })
  assert.deepEqual(rememberDocumentColors(memory, { id: 'doc.command.format-text', patch: {}, clearFields: ['highlight'] }), memory)
  assert.equal(createDocumentRibbon({ ...context, editingDisabled: true }).find(tab => tab.id === 'home').groups.flatMap(group => group.items).find(item => item.id === 'doc-text-color').disabled, true)
})

test('all locales distinguish background reset from clear and expose localized color panel actions', () => {
  const messages = createOfficeMessages()
  for (const locale of Object.keys(messages)) {
    const labels = messages[locale]
    assert.notEqual(labels['items.doc-highlight-reset'], labels['document.color.noBackground'], locale)
    assert.equal(labels['items.doc-highlight-clear'], labels['document.color.noBackground'], locale)
    for (const key of ['more', 'cancel', 'confirm', 'reset', 'automatic', 'standard', 'value', 'format', 'invalid']) {
      assert.equal(typeof labels[`document.color.${key}`], 'string', `${locale}: ${key}`)
      assert.ok(labels[`document.color.${key}`].length > 0, `${locale}: ${key}`)
    }
  }
  assert.equal(messages['zh-CN']['items.doc-highlight-reset'], '重置颜色')
  assert.equal(messages['en-US']['items.doc-highlight-reset'], 'Reset')
})
