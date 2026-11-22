import assert from 'node:assert/strict'
import { test } from 'vitest'
import { confirmRibbonDraft, ribbonFontOptions, fontSizeOptions } from '../src/platform/ribbon/ribbon-fonts.ts'
import { ribbonInvocation } from '../src/platform/ribbon/ribbon-schema.ts'
import { runFont } from '../src/features/document/layout/document-layout.ts'

test('font list override is instance-local, labels match in source order and an imported family stays visible', () => {
  const options = ribbonFontOptions('Imported', { override: true, families: [{ value: 'local-a', label: 'Office Sans' }, { value: 'local-b', label: 'Office Serif' }] })
  assert.deepEqual(options.map(x => x.value), ['local-a', 'local-b', 'Imported'])
  const item = { type: 'select', edit: { kind: 'font' }, commandId: 'font', options }
  assert.equal(confirmRibbonDraft(item, ' OFFICE '), 'local-a')
  assert.equal(confirmRibbonDraft(item, 'absent'), undefined)
  assert.equal(ribbonInvocation(item, 'absent'), undefined)
  assert.deepEqual(ribbonFontOptions().slice(0, 3).map(x => x.value), ['宋体', 'Arial', 'Baskerville'])
})

test('size confirmation clamps by editor, preserves decimals and keeps arbitrary input out of other selectors', () => {
  assert.deepEqual(fontSizeOptions.map(o => o.value), [9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72])
  assert.match(runFont({ fontFamily: 'Arial', fontSize: 3 }), /^4px Arial,/)
  for (const min of [1, 6]) {
    const item = { type: 'select', edit: { kind: 'number', min, max: 400 }, options: fontSizeOptions, commandId: 'size' }
    for (const invalid of ['', 'Infinity', '0x80', '12px']) assert.equal(confirmRibbonDraft(item, invalid), undefined)
    assert.equal(confirmRibbonDraft(item, '0'), min)
    assert.equal(confirmRibbonDraft(item, '401'), 400)
    assert.deepEqual(ribbonInvocation(item, confirmRibbonDraft(item, '13.5')), { commandId: 'size', params: 13.5 })
    for (const n of [NaN, 401, min - 1]) assert.equal(ribbonInvocation(item, n), undefined)
    assert.equal(ribbonInvocation({ ...item, disabled: true }, 13.5), undefined)
    assert.equal(ribbonInvocation({ ...item, edit: undefined }, 13.5), undefined)
  }
})
