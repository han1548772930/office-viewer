import type { CellRange, SheetRangeThemeRule } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { assertSheetAxisEdit, sheetAxisEditSegments, sheetAxisMaximum, sheetRangeAxisInterval, sheetRangeWithAxisInterval,
  transformSheetAxisRange, type SheetAxisEdit } from '../../domain/sheet-axis-structure.ts'
import { assertSheetCommandRanges, sheetRangesEqual, sheetRangesIntersect, subtractSheetCommandRange } from './sheet-range-command.ts'

export interface SheetRangeThemeCellShift {
  kind: 'insert' | 'delete'
  axis: 'row' | 'column'
  start: number
  count: number
  bandStart: number
  bandEnd: number
}
const MAX_RULES = 100_000, MAX_RANGE_WORK = 5_000_000
const newId = () => `range-theme-${crypto.randomUUID()}`
function assertRuleBudget(length: number): void {
  if (length > MAX_RULES) throw new RangeError('范围主题片段数量超出当前同步计算预算。')
}
function assertRules(rules: readonly SheetRangeThemeRule[]): void {
  if (!Array.isArray(rules)) throw new TypeError('范围主题规则必须是数组。')
  assertRuleBudget(rules.length); assertSheetCommandRanges(rules.map(rule => rule.range))
  const ids = new Set<string>()
  for (const rule of rules) {
    if (typeof rule.id !== 'string' || !rule.id.trim() || ids.has(rule.id) || typeof rule.themeName !== 'string' || !rule.themeName.trim()) throw new TypeError('范围主题规则身份无效。')
    ids.add(rule.id)
    if (rule.styleRange !== undefined) assertSheetRangeThemeStyleRange(rule.styleRange, rule.range)
  }
}
/** A clipped fragment may retain a logical origin outside the sheet after moving to its edge. */
export function assertSheetRangeThemeStyleRange(range: unknown, physicalRange?: CellRange): asserts range is CellRange {
  if (!range || typeof range !== 'object' || Array.isArray(range)) throw new RangeError('范围主题逻辑样式范围无效。')
  const value = range as Record<string, unknown>
  if (![value.startRow, value.endRow, value.startCol, value.endCol].every(Number.isSafeInteger)
    || Number(value.startRow) > Number(value.endRow) || Number(value.startCol) > Number(value.endCol)
    || Number(value.startRow) < -1_048_576 || Number(value.endRow) > 2_097_151
    || Number(value.startCol) < -16_384 || Number(value.endCol) > 32_767) throw new RangeError('范围主题逻辑样式范围无效。')
  if (physicalRange && (Number(value.startRow) > physicalRange.startRow || Number(value.endRow) < physicalRange.endRow
    || Number(value.startCol) > physicalRange.startCol || Number(value.endCol) < physicalRange.endCol)) {
    throw new RangeError('范围主题逻辑样式范围必须包含实际应用区域。')
  }
}
function fragment(rule: SheetRangeThemeRule, range: CellRange, styleRange: CellRange, id = rule.id): SheetRangeThemeRule {
  assertSheetRangeThemeStyleRange(styleRange, range)
  const next = { ...cloneModel(rule), id, range: { ...range } }
  if (sheetRangesEqual(range, styleRange)) delete next.styleRange
  else next.styleRange = { ...styleRange }
  return next
}
function translate(range: CellRange, row: number, col: number): CellRange {
  return { startRow: range.startRow + row, endRow: range.endRow + row, startCol: range.startCol + col, endCol: range.endCol + col }
}
function shiftedStyleRange(range: CellRange, edit: SheetRangeThemeCellShift): CellRange {
  const interval = sheetRangeAxisInterval(range, edit.axis), end = edit.start + edit.count - 1
  if (interval.end < edit.start) return { ...range }
  if (edit.kind === 'insert') return sheetRangeWithAxisInterval(range, edit.axis, {
    start: interval.start >= edit.start ? interval.start + edit.count : interval.start, end: interval.end + edit.count,
  })
  if (interval.start > end) return sheetRangeWithAxisInterval(range, edit.axis, { start: interval.start - edit.count, end: interval.end - edit.count })
  return sheetRangeWithAxisInterval(range, edit.axis, {
    start: interval.start < edit.start ? interval.start : edit.start,
    end: interval.end > end ? interval.end - edit.count : edit.start - 1,
  })
}

/** Rectangle operations are bounded by rule/fragment count, never by the number of styled blank cells. */
export function planSheetRangeThemeCellShift(rules: readonly SheetRangeThemeRule[], edit: SheetRangeThemeCellShift): SheetRangeThemeRule[] {
  assertSheetAxisEdit(edit)
  if (edit.kind !== 'insert' && edit.kind !== 'delete') throw new TypeError('范围主题局部移位类型无效。')
  const bandAxis = edit.axis === 'row' ? 'column' : 'row', maximum = sheetAxisMaximum(bandAxis)
  if (!Number.isSafeInteger(edit.bandStart) || !Number.isSafeInteger(edit.bandEnd) || edit.bandStart < 0
    || edit.bandEnd < edit.bandStart || edit.bandEnd >= maximum) throw new RangeError('范围主题移位带无效。')
  assertRules(rules)
  const result: SheetRangeThemeRule[] = []
  for (const rule of rules) {
    const band = sheetRangeAxisInterval(rule.range, bandAxis), axis = sheetRangeAxisInterval(rule.range, edit.axis)
    if (band.end < edit.bandStart || band.start > edit.bandEnd || axis.end < edit.start) {
      result.push(cloneModel(rule)); continue
    }
    if (edit.kind === 'insert' && axis.end + edit.count >= sheetAxisMaximum(edit.axis)) {
      throw new RangeError('插入会把已有范围主题移出 Office 行列上限。')
    }
    const styleRange = rule.styleRange ?? rule.range, pieces: Array<{ range: CellRange; styleRange: CellRange }> = []
    if (band.start < edit.bandStart) pieces.push({ range: sheetRangeWithAxisInterval(rule.range, bandAxis, { start: band.start, end: edit.bandStart - 1 }), styleRange })
    if (band.end > edit.bandEnd) pieces.push({ range: sheetRangeWithAxisInterval(rule.range, bandAxis, { start: edit.bandEnd + 1, end: band.end }), styleRange })
    const selected = sheetRangeWithAxisInterval(rule.range, bandAxis, { start: Math.max(band.start, edit.bandStart), end: Math.min(band.end, edit.bandEnd) })
    for (const range of transformSheetAxisRange(selected, edit)) pieces.push({ range, styleRange: shiftedStyleRange(styleRange, edit) })
    for (const [index, piece] of pieces.entries()) result.push(fragment(rule, piece.range, piece.styleRange, index ? newId() : rule.id))
    assertRuleBudget(result.length)
  }
  assertRuleBudget(result.length)
  return result
}

/** Full axis changes reuse the same fragment geometry, including subsequent edits of a partial move. */
export function planSheetRangeThemeAxisEdit(rules: readonly SheetRangeThemeRule[], edit: SheetAxisEdit): SheetRangeThemeRule[] {
  assertSheetAxisEdit(edit); assertRules(rules)
  const bandEnd = sheetAxisMaximum(edit.axis === 'row' ? 'column' : 'row') - 1
  if (edit.kind === 'insert' || edit.kind === 'delete') return planSheetRangeThemeCellShift(rules, { ...edit, bandStart: 0, bandEnd })
  if (edit.kind === 'copy') {
    const result = planSheetRangeThemeCellShift(rules, { axis: edit.axis, kind: 'insert', start: edit.destination, count: edit.count, bandStart: 0, bandEnd })
    for (const rule of rules) {
      const interval = sheetRangeAxisInterval(rule.range, edit.axis), start = Math.max(interval.start, edit.start), end = Math.min(interval.end, edit.start + edit.count - 1)
      if (start > end) continue
      const offset = edit.destination - edit.start, selected = sheetRangeWithAxisInterval(rule.range, edit.axis, { start, end })
      const row = edit.axis === 'row' ? offset : 0, col = edit.axis === 'column' ? offset : 0
      const range = translate(selected, row, col)
      if (sheetRangeAxisInterval(range, edit.axis).end >= sheetAxisMaximum(edit.axis)) throw new RangeError('范围主题复制后越界。')
      result.push(fragment(rule, range, translate(rule.styleRange ?? rule.range, row, col), newId())); assertRuleBudget(result.length)
    }
    return result
  }
  if (edit.destination === edit.start || edit.destination === edit.start + edit.count) return cloneModel([...rules])
  const result: SheetRangeThemeRule[] = [], segments = sheetAxisEditSegments(edit)
  for (const rule of rules) {
    const interval = sheetRangeAxisInterval(rule.range, edit.axis)
    let index = 0
    for (const segment of segments) {
      const start = Math.max(interval.start, segment.start), end = Math.min(interval.end, segment.end)
      if (start > end) continue
      const row = edit.axis === 'row' ? segment.offset : 0, col = edit.axis === 'column' ? segment.offset : 0
      const selected = sheetRangeWithAxisInterval(rule.range, edit.axis, { start, end })
      result.push(fragment(rule, translate(selected, row, col), translate(rule.styleRange ?? rule.range, row, col), index++ ? newId() : rule.id))
      assertRuleBudget(result.length)
    }
  }
  return result
}

function intersection(first: CellRange, second: CellRange): CellRange | undefined {
  return sheetRangesIntersect(first, second) ? { startRow: Math.max(first.startRow, second.startRow), endRow: Math.min(first.endRow, second.endRow),
    startCol: Math.max(first.startCol, second.startCol), endCol: Math.min(first.endCol, second.endCol) } : undefined
}
function trim(rules: readonly SheetRangeThemeRule[], removed: readonly CellRange[]): SheetRangeThemeRule[] {
  const result: SheetRangeThemeRule[] = []
  for (const rule of rules) {
    let ranges = [{ ...rule.range }]
    for (const range of removed) ranges = ranges.flatMap(current => subtractSheetCommandRange(current, range))
    for (const [index, range] of ranges.entries()) result.push(fragment(rule, range, rule.styleRange ?? rule.range, index ? newId() : rule.id))
    assertRuleBudget(result.length)
  }
  return result
}
export function planSheetRangeThemeCut(sourceRules: readonly SheetRangeThemeRule[], targetRules: readonly SheetRangeThemeRule[],
  sourceRange: CellRange, targetRange: CellRange, sameSheet: boolean): { sourceRules: SheetRangeThemeRule[]; targetRules: SheetRangeThemeRule[] } {
  assertSheetCommandRanges([sourceRange, targetRange])
  assertRules(sourceRules); assertRules(targetRules)
  if (sourceRange.endRow - sourceRange.startRow !== targetRange.endRow - targetRange.startRow
    || sourceRange.endCol - sourceRange.startCol !== targetRange.endCol - targetRange.startCol) throw new RangeError('范围主题剪切目标尺寸不同。')
  const row = targetRange.startRow - sourceRange.startRow, col = targetRange.startCol - sourceRange.startCol
  if (sameSheet && !row && !col) { const rules = cloneModel([...sourceRules]); return { sourceRules: rules, targetRules: rules } }
  const retainedSource = trim(sourceRules, sameSheet ? [sourceRange, targetRange] : [sourceRange])
  const retainedTarget = sameSheet ? retainedSource : trim(targetRules, [targetRange])
  const ids = new Set([...retainedSource, ...retainedTarget].map(rule => rule.id))
  const additions: SheetRangeThemeRule[] = []
  for (const rule of sourceRules) {
    const selected = intersection(rule.range, sourceRange)
    if (!selected) continue
    const id = ids.has(rule.id) ? newId() : rule.id; ids.add(id)
    additions.push(fragment(rule, translate(selected, row, col), translate(rule.styleRange ?? rule.range, row, col), id))
    assertRuleBudget(retainedTarget.length + additions.length)
  }
  const nextTarget = [...retainedTarget, ...additions]; assertRuleBudget(nextTarget.length)
  return { sourceRules: sameSheet ? nextTarget : retainedSource, targetRules: nextTarget }
}

/** Group by observable source geometry so splitting IDs does not expand permissions into unchanged fragments. */
export function sheetRangeThemeChangedRanges(before: readonly SheetRangeThemeRule[], after: readonly SheetRangeThemeRule[]): CellRange[] {
  assertRules(before); assertRules(after)
  const groups = (rules: readonly SheetRangeThemeRule[]) => {
    const result = new Map<string, CellRange[]>()
    for (const rule of rules) {
      const key = JSON.stringify([rule.themeName, rule.styleRange ?? rule.range]), ranges = result.get(key) ?? []
      ranges.push(rule.range); result.set(key, ranges)
    }
    return result
  }
  const first = groups(before), next = groups(after), result: CellRange[] = []
  let work = 0
  function difference(original: readonly CellRange[], removed: readonly CellRange[]): void {
    let remaining = original.map(range => ({ ...range }))
    for (const cut of removed) {
      work += remaining.length
      if (work > MAX_RANGE_WORK) throw new RangeError('范围主题差异超出当前同步计算预算。')
      remaining = remaining.flatMap(range => subtractSheetCommandRange(range, cut)); assertRuleBudget(remaining.length)
    }
    result.push(...remaining); assertRuleBudget(result.length)
  }
  for (const key of new Set([...first.keys(), ...next.keys()])) {
    difference(first.get(key) ?? [], next.get(key) ?? []); difference(next.get(key) ?? [], first.get(key) ?? [])
  }
  return result
}
