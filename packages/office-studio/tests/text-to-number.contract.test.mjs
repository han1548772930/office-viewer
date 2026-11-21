import assert from 'node:assert/strict'
import { test } from 'vitest'
import { convertTextToNumber } from '../src/features/sheet-workbench/text-to-number.ts'

test('converts finite numeric text and removes an explicit text number format', () => {
  assert.deepEqual(convertTextToNumber({ value: '  -1.25e2  ', rawValue: '  -1.25e2  ', rawNumberFormat: '@', numberFormat: 'text' }), {
    value: '-125', rawValue: -125, numberFormat: 'general', rawNumberFormat: 'General', displayValue: undefined,
  })
})

test('does not convert formulas, non-numeric strings, empty values or non-finite numbers', () => {
  assert.equal(convertTextToNumber({ value: '=1+2', formula: '=1+2' }), undefined)
  assert.equal(convertTextToNumber({ value: '12x', rawValue: '12x' }), undefined)
  assert.equal(convertTextToNumber({ value: ' ', rawValue: ' ' }), undefined)
  assert.equal(convertTextToNumber({ value: '1e999', rawValue: '1e999' }), undefined)
  assert.equal(convertTextToNumber({ value: '42', rawValue: 42 }), undefined)
  assert.equal(convertTextToNumber({ value: '42', rawValue: 42, numberFormat: 'text' })?.rawValue, 42)
})
