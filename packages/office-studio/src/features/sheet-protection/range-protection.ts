import type { CellRange, SheetRangeProtectionDraft, SheetRangeProtectionRule, SheetWorksheetProtectionRule, SpreadsheetSheet } from '../../domain/model.ts'

const MAX_ROW_INDEX = 1_048_575
const MAX_COLUMN_INDEX = 16_383

export function parseProtectionRange(value: string): CellRange {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i.exec(value.trim())
  if (!match) throw new TypeError('请输入有效的单元格区域，例如 A1 或 A1:B4。')
  const firstColumn = columnIndex(match[1]!)
  const firstRow = Number(match[2]) - 1
  const lastColumn = match[3] ? columnIndex(match[3]) : firstColumn
  const lastRow = match[4] ? Number(match[4]) - 1 : firstRow
  if (firstColumn > lastColumn || firstRow > lastRow) throw new RangeError('保护区域必须从左上角延伸到右下角。')
  return validateProtectionRange({ startRow: firstRow, startCol: firstColumn, endRow: lastRow, endCol: lastColumn })
}

export function formatProtectionRange(range: CellRange): string {
  const first = `${columnLabel(range.startCol)}${range.startRow + 1}`
  const last = `${columnLabel(range.endCol)}${range.endRow + 1}`
  return first === last ? first : `${first}:${last}`
}

export function createSheetRangeProtectionRule(
  draft: SheetRangeProtectionDraft,
  ranges: readonly CellRange[],
  ownerUserId: string,
): SheetRangeProtectionRule {
  const normalizedRanges = ranges.map(validateProtectionRange)
  if (!normalizedRanges.length) throw new RangeError('请选择至少一个有效的保护区域。')
  const description = draft.description?.trim() ?? ''
  if (description.length > 255) throw new RangeError('保护说明不能超过 255 个字符。')
  const allowedUserIds = [...new Set((draft.allowedUserIds ?? []).map(userId => userId.trim()).filter(Boolean))]
  if (allowedUserIds.some(userId => userId.length > 255)) throw new RangeError('用户 ID 不能超过 255 个字符。')
  return {
    id: draft.id?.trim() || crypto.randomUUID(),
    permissionId: draft.permissionId?.trim() || crypto.randomUUID(),
    ranges: normalizedRanges,
    description,
    ownerUserId: ownerUserId.trim(),
    allowedUserIds,
    editState: draft.editState,
    viewState: draft.viewState ?? 'othersCanView',
    metadata: draft.metadata ? { ...draft.metadata } : undefined,
  }
}

export function canEditProtectedRanges(
  sheet: Pick<SpreadsheetSheet, 'rangeProtectionRules'>,
  ranges: readonly CellRange[],
  userId: string,
): boolean {
  const rules = sheet.rangeProtectionRules ?? []
  return rules.every(rule => !rule.ranges.some(protectedRange => ranges.some(range => rangesIntersect(protectedRange, range))) ||
    rule.ownerUserId === userId || (rule.editState === 'designedUserCanEdit' && rule.allowedUserIds.includes(userId)))
}

export function canEditWorksheetProtection(
  sheet: Pick<SpreadsheetSheet, 'worksheetProtection'>,
  userId: string,
): boolean {
  const rule = sheet.worksheetProtection
  if (!rule) return true
  return rule.ownerUserId === userId || (rule.editState === 'designedUserCanEdit' && rule.allowedUserIds.includes(userId))
}

export function canViewWorksheetProtection(
  sheet: Pick<SpreadsheetSheet, 'worksheetProtection'>,
  userId: string,
): boolean {
  const rule = sheet.worksheetProtection
  return !rule || rule.viewState !== 'noOneElseCanView' || rule.ownerUserId === userId ||
    (rule.editState === 'designedUserCanEdit' && rule.allowedUserIds.includes(userId))
}

export function createWorksheetProtectionRule(
  input: Omit<SheetWorksheetProtectionRule, 'id' | 'permissionId' | 'unitType' | 'ownerUserId'>,
  ownerUserId: string,
): SheetWorksheetProtectionRule {
  return validateWorksheetProtectionRule({
    ...input,
    id: crypto.randomUUID(),
    permissionId: crypto.randomUUID(),
    unitType: 'worksheet',
    ownerUserId,
  })
}

export function validateWorksheetProtectionRule(
  rule: SheetWorksheetProtectionRule,
): SheetWorksheetProtectionRule {
  if (!rule || typeof rule.id !== 'string' || !rule.id.trim() ||
      typeof rule.permissionId !== 'string' || !rule.permissionId.trim() || rule.unitType !== 'worksheet' ||
      typeof rule.ownerUserId !== 'string' || !rule.ownerUserId.trim()) {
    throw new TypeError('工作表保护规则不完整。')
  }
  if (rule.editState !== 'onlyMe' && rule.editState !== 'designedUserCanEdit') throw new TypeError('工作表编辑权限无效。')
  if (rule.viewState !== 'othersCanView' && rule.viewState !== 'noOneElseCanView') throw new TypeError('工作表查看权限无效。')
  if (rule.description !== undefined && typeof rule.description !== 'string') throw new TypeError('保护说明必须是文本。')
  const description = rule.description?.trim()
  if (description && description.length > 255) throw new RangeError('保护说明不能超过 255 个字符。')
  if (!Array.isArray(rule.allowedUserIds) || rule.allowedUserIds.some(userId => typeof userId !== 'string')) {
    throw new TypeError('允许编辑的用户必须是用户 ID 列表。')
  }
  const allowedUserIds = [...new Set(rule.allowedUserIds.map(userId => userId.trim()).filter(Boolean))]
  if (allowedUserIds.some(userId => userId.length > 255)) throw new RangeError('用户 ID 不能超过 255 个字符。')
  if (rule.editState === 'designedUserCanEdit' && !allowedUserIds.length) throw new RangeError('请至少填写一个允许编辑的用户 ID。')
  return { ...rule, id: rule.id.trim(), permissionId: rule.permissionId.trim(), ownerUserId: rule.ownerUserId.trim(), description, allowedUserIds }
}

export function setWorksheetProtectionRule(sheet: SpreadsheetSheet, rule: SheetWorksheetProtectionRule): void {
  sheet.worksheetProtection = validateWorksheetProtectionRule(rule)
}

export function assertWorksheetProtectionChanges(before: { sheets: SpreadsheetSheet[] }, after: { sheets: SpreadsheetSheet[] }, userId: string): void {
  for (const sheet of before.sheets) {
    const next = after.sheets.find(candidate => candidate.id === sheet.id)
    if (sheet.worksheetProtection && JSON.stringify(sheet.worksheetProtection) !== JSON.stringify(next?.worksheetProtection) &&
        sheet.worksheetProtection.ownerUserId !== userId) {
      throw new Error('只有工作表保护创建者可以修改或移除该保护。')
    }
    if (!canEditWorksheetProtection(sheet, userId)) {
      const { worksheetProtection: _beforeRule, ...previousContent } = sheet
      const { worksheetProtection: _afterRule, ...nextContent } = next ?? {}
      if (JSON.stringify(previousContent) !== JSON.stringify(nextContent)) throw new Error('该工作表受保护，当前用户没有编辑权限。')
    }
  }
  for (const sheet of after.sheets) {
    if (!sheet.worksheetProtection) continue
    const previous = before.sheets.find(candidate => candidate.id === sheet.id)?.worksheetProtection
    if (JSON.stringify(previous) === JSON.stringify(sheet.worksheetProtection)) continue
    const rule = validateWorksheetProtectionRule(sheet.worksheetProtection)
    if (rule.ownerUserId !== userId) throw new Error('工作表保护创建者必须是当前用户。')
    sheet.worksheetProtection = rule
  }
}

export function setSheetRangeProtectionRule(sheet: SpreadsheetSheet, rule: SheetRangeProtectionRule): void {
  const existing = sheet.rangeProtectionRules ?? []
  if (!rule.id || !rule.ownerUserId || !rule.ranges.length) throw new TypeError('保护规则不完整。')
  sheet.rangeProtectionRules = [...existing.filter(item => item.id !== rule.id), cloneRule(rule)]
}

export function removeSheetRangeProtectionRule(sheet: SpreadsheetSheet, ruleId: string, userId: string): boolean {
  const existing = sheet.rangeProtectionRules ?? []
  const rule = existing.find(item => item.id === ruleId)
  if (!rule) return false
  if (rule.ownerUserId !== userId) throw new Error('只有保护规则创建者可以移除该规则。')
  sheet.rangeProtectionRules = existing.filter(item => item.id !== ruleId)
  return true
}

export function formatProtectionUserIds(value: string): string[] {
  return [...new Set(value.split(/[\s,;]+/).map(userId => userId.trim()).filter(Boolean))]
}

function validateProtectionRange(range: CellRange): CellRange {
  if (![range.startRow, range.startCol, range.endRow, range.endCol].every(Number.isInteger) ||
      range.startRow < 0 || range.startCol < 0 || range.endRow < range.startRow || range.endCol < range.startCol ||
      range.endRow > MAX_ROW_INDEX || range.endCol > MAX_COLUMN_INDEX) {
    throw new RangeError('保护区域超出有效工作表范围。')
  }
  return { ...range }
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  return left.startRow <= right.endRow && left.endRow >= right.startRow &&
    left.startCol <= right.endCol && left.endCol >= right.startCol
}

function columnIndex(value: string): number {
  let index = 0
  for (const character of value.toUpperCase()) index = index * 26 + character.charCodeAt(0) - 64
  if (index < 1 || index > MAX_COLUMN_INDEX + 1) throw new RangeError('列地址超出工作表范围。')
  return index - 1
}

function columnLabel(index: number): string {
  let label = '', cursor = index
  while (cursor >= 0) {
    label = String.fromCharCode(cursor % 26 + 65) + label
    cursor = Math.floor(cursor / 26) - 1
  }
  return label
}

function cloneRule(rule: SheetRangeProtectionRule): SheetRangeProtectionRule {
  return { ...rule, ranges: rule.ranges.map(range => ({ ...range })), allowedUserIds: [...rule.allowedUserIds], metadata: rule.metadata ? { ...rule.metadata } : undefined }
}
