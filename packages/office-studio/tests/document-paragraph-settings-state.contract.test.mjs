import assert from 'node:assert/strict'
import { test } from 'vitest'
import {
  changeDocumentParagraphSpacingRule, documentParagraphLineSpacingInput, documentParagraphSettingsStyle,
  isDocumentParagraphSpacingRule,
} from '../src/components/document/document-paragraph-settings-state.ts'

test('paragraph controls resolve missing defaults and legacy left indentation without changing the model', () => {
  const source = { align: 'right', indent: 2, lineHeight: 1.25, indentStart: undefined, spacingRule: undefined, wordWrap: undefined }
  const before = { ...source }
  assert.deepEqual(documentParagraphSettingsStyle(source), {
    align: 'right', indent: 2, lineHeight: 1.25, spaceBefore: 0, spaceAfter: 8,
    indentStart: 48, indentEnd: 0, indentFirstLine: 0, hanging: 0, spacingRule: 'auto', wordWrap: false, snapToGrid: false,
  })
  assert.deepEqual(source, before)
  assert.equal(documentParagraphSettingsStyle({ indent: 2, indentStart: 0 }).indentStart, 0)
})

test('switching line spacing units uses the displayed glyph and line heights rather than a guessed font size', () => {
  assert.equal(changeDocumentParagraphSpacingRule(1.5, 'auto', 'exact', { glyphLineHeight: 20, renderedLineHeight: 30 }, 11), 30)
  assert.equal(changeDocumentParagraphSpacingRule(30, 'exact', 'auto', { glyphLineHeight: 20, renderedLineHeight: 30 }, 100), 1.5)
  assert.equal(changeDocumentParagraphSpacingRule(8, 'atLeast', 'exact', { glyphLineHeight: 20, renderedLineHeight: 20 }), 20)
  assert.equal(changeDocumentParagraphSpacingRule(30, 'exact', 'atLeast', { glyphLineHeight: 20, renderedLineHeight: 30 }), 30)
  assert.equal(changeDocumentParagraphSpacingRule(1.15, 'auto', 'auto', { glyphLineHeight: 20, renderedLineHeight: 23 }), 1.15)
})

test('line spacing conversion has finite CSS-pixel fallbacks and obeys both unit boundaries', () => {
  assert.equal(changeDocumentParagraphSpacingRule(1.5, 'auto', 'exact', undefined, 11), 22)
  assert.equal(changeDocumentParagraphSpacingRule(22, 'exact', 'auto', undefined, 11), 1.5)
  assert.equal(changeDocumentParagraphSpacingRule(1.5, 'auto', 'exact', { glyphLineHeight: NaN, renderedLineHeight: Infinity }, NaN), 22)
  assert.equal(changeDocumentParagraphSpacingRule(100, 'exact', 'auto', { glyphLineHeight: 10, renderedLineHeight: 100 }), 5)
  assert.equal(changeDocumentParagraphSpacingRule(1, 'exact', 'auto', { glyphLineHeight: 20, renderedLineHeight: 1 }), 1)
  assert.equal(changeDocumentParagraphSpacingRule(5, 'auto', 'exact', { glyphLineHeight: 40, renderedLineHeight: 200 }), 100)
  assert.equal(documentParagraphLineSpacingInput('auto').max, 5)
  assert.equal(documentParagraphLineSpacingInput('exact').max, 100)
  for (const rule of ['auto', 'atLeast', 'exact']) assert.equal(isDocumentParagraphSpacingRule(rule), true)
  for (const rule of ['multiple', '', undefined, null, 1]) assert.equal(isDocumentParagraphSpacingRule(rule), false)
})
