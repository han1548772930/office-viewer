import assert from 'node:assert/strict'
import { test } from 'vitest'
import {
  canEditProtectedRanges,
  createSheetRangeProtectionRule,
  formatProtectionRange,
  formatProtectionUserIds,
  parseProtectionRange,
  removeSheetRangeProtectionRule,
  setSheetRangeProtectionRule,
} from '../src/features/sheet-protection/range-protection.ts'

test('protection panel range input normalizes A1 ranges and user lists', () => {
  assert.deepEqual(parseProtectionRange('$B$2:D4'), { startRow: 1, startCol: 1, endRow: 3, endCol: 3 })
  assert.equal(formatProtectionRange({ startRow: 1, startCol: 1, endRow: 3, endCol: 3 }), 'B2:D4')
  assert.deepEqual(formatProtectionUserIds('alex, pat;alex\nlee'), ['alex', 'pat', 'lee'])
})

test('range protection rule edits preserve ownership, authorized users and removal constraints', () => {
  const sheet = { rangeProtectionRules: [] }
  const rule = createSheetRangeProtectionRule({
    ranges: [], description: 'Budget', editState: 'designedUserCanEdit', viewState: 'othersCanView', allowedUserIds: ['editor'],
  }, [parseProtectionRange('A1:B4')], 'owner')
  setSheetRangeProtectionRule(sheet, rule)
  assert.equal(canEditProtectedRanges(sheet, [parseProtectionRange('B2')], 'editor'), true)
  assert.equal(canEditProtectedRanges(sheet, [parseProtectionRange('B2')], 'stranger'), false)
  assert.equal(canEditProtectedRanges(sheet, [parseProtectionRange('C1')], 'stranger'), true)
  assert.throws(() => removeSheetRangeProtectionRule(sheet, rule.id, 'stranger'), /创建者/)
  assert.equal(removeSheetRangeProtectionRule(sheet, rule.id, 'owner'), true)
  assert.deepEqual(sheet.rangeProtectionRules, [])
})
