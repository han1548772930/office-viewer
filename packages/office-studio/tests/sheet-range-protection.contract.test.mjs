import assert from 'node:assert/strict'
import { test } from 'vitest'
import {
  canEditProtectedRanges,
  canEditWorksheetProtection,
  canViewWorksheetProtection,
  createWorksheetProtectionRule,
  createSheetRangeProtectionRule,
  formatProtectionRange,
  formatProtectionUserIds,
  parseProtectionRange,
  removeSheetRangeProtectionRule,
  setSheetRangeProtectionRule,
  setWorksheetProtectionRule,
  validateWorksheetProtectionRule,
} from '../src/features/sheet-protection/range-protection.ts'

test('protection panel range input normalizes A1 ranges and user lists', () => {
  assert.deepEqual(parseProtectionRange('$B$2:D4'), { startRow: 1, startCol: 1, endRow: 3, endCol: 3 })
  assert.equal(formatProtectionRange({ startRow: 1, startCol: 1, endRow: 3, endCol: 3 }), 'B2:D4')
  assert.deepEqual(formatProtectionUserIds('alex, pat;alex\nlee'), ['alex', 'pat', 'lee'])
})

test('worksheet protection has a whole-sheet permission rule independent of range protection', () => {
  const sheet = { rangeProtectionRules: [] }
  const rule = createWorksheetProtectionRule({
    description: 'Forecast', editState: 'designedUserCanEdit', viewState: 'othersCanView', allowedUserIds: ['editor', 'editor'],
  }, 'owner')
  setWorksheetProtectionRule(sheet, rule)
  assert.equal(rule.unitType, 'worksheet')
  assert.deepEqual(rule.allowedUserIds, ['editor'])
  assert.equal(canEditWorksheetProtection(sheet, 'owner'), true)
  assert.equal(canEditWorksheetProtection(sheet, 'editor'), true)
  assert.equal(canEditWorksheetProtection(sheet, 'stranger'), false)
  assert.equal(canViewWorksheetProtection(sheet, 'stranger'), true)
  assert.equal(canEditProtectedRanges(sheet, [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }], 'stranger'), true)
  assert.throws(() => createWorksheetProtectionRule({
    editState: 'designedUserCanEdit', viewState: 'othersCanView', allowedUserIds: [],
  }, 'owner'), /至少填写一个/)
  setWorksheetProtectionRule(sheet, { ...rule, viewState: 'noOneElseCanView', editState: 'onlyMe', allowedUserIds: [] })
  assert.equal(canViewWorksheetProtection(sheet, 'stranger'), false)
  assert.throws(() => validateWorksheetProtectionRule({ ...rule, editState: 'unexpected' }), /编辑权限无效/)
  assert.throws(() => validateWorksheetProtectionRule({ ...rule, description: 'x'.repeat(256) }), /255 个字符/)
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
