import type { CellRange, SheetConditionalFormatRule, SheetConditionalFormatValueConfig } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { formulaRelativeReferenceSignature } from '../formula/formula-references.ts'
import { conditionalFormatFormulaAnchor, mapConditionalFormatFormulas, reanchorConditionalFormatRule as reanchorRule } from '../conditional-formatting/conditional-format-formulas.ts'
import { assertSheetCommandRanges, sheetRangesIntersect, subtractSheetCommandRange } from './sheet-range-command.ts'

export interface SheetConditionalFormatPainterSource {
  readonly sheetId: string
  /** Full original ranges retain the expression anchor even when only part of a rule is sampled. */
  readonly rules: readonly SheetConditionalFormatRule[]
}
export interface SheetConditionalFormatPainterPlan {
  changed: boolean
  rules: SheetConditionalFormatRule[] | undefined
}

const MAXIMUM_RULE_RANGES = 100_000
const MAXIMUM_RANGE_WORK = 5_000_000
interface RangeWork { operations: number }
function useRangeWork(work: RangeWork, count = 1): void {
  work.operations += count
  if (work.operations > MAXIMUM_RANGE_WORK) throw new RangeError('格式刷条件格式范围计算超出当前同步预算。')
}
function assertJsonRule(value: unknown, ancestors = new Set<object>()): void {
  if (value === undefined || value === null || typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number' && Number.isFinite(value)) return
  if (typeof value !== 'object' || ancestors.has(value)) throw new TypeError('格式刷条件格式规则必须是无循环的 JSON 数据。')
  const prototype = Object.getPrototypeOf(value)
  if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) throw new TypeError('格式刷条件格式规则必须是 JSON 数据。')
  ancestors.add(value)
  for (const nested of Object.values(value)) assertJsonRule(nested, ancestors)
  ancestors.delete(value)
}
function assertRule(rule: SheetConditionalFormatRule): void {
  if (!rule || typeof rule !== 'object' || typeof rule.id !== 'string' || !rule.id.trim() || !Array.isArray(rule.ranges)
    || typeof rule.type !== 'string' || typeof rule.operator !== 'string' || typeof rule.formula1 !== 'string'
    || !Number.isFinite(rule.priority) || typeof rule.stopIfTrue !== 'boolean' || !rule.style || typeof rule.style !== 'object' || Array.isArray(rule.style)) {
    throw new TypeError('格式刷条件格式规则无效。')
  }
  assertSheetCommandRanges(rule.ranges)
  if (rule.formulaAnchor !== undefined) {
    if (!rule.formulaAnchor || typeof rule.formulaAnchor !== 'object' || Array.isArray(rule.formulaAnchor)) throw new TypeError('格式刷条件格式公式基准无效。')
    assertSheetCommandRanges([{ startRow: rule.formulaAnchor.row, endRow: rule.formulaAnchor.row, startCol: rule.formulaAnchor.col, endCol: rule.formulaAnchor.col }])
  }
  if (rule.formula2 !== undefined && typeof rule.formula2 !== 'string') throw new TypeError('格式刷条件格式规则无效。')
  if (rule.config !== undefined && (!rule.config || typeof rule.config !== 'object' || Array.isArray(rule.config))) throw new TypeError('格式刷条件格式配置无效。')
  const threshold = (value: SheetConditionalFormatValueConfig) => {
    if (!value || typeof value !== 'object' || Array.isArray(value) || typeof value.type !== 'string'
      || value.value !== undefined && typeof value.value !== 'number' && typeof value.value !== 'string') throw new TypeError('格式刷条件格式阈值无效。')
  }
  if (rule.config?.min !== undefined) threshold(rule.config.min)
  if (rule.config?.max !== undefined) threshold(rule.config.max)
  for (const entries of [rule.config?.thresholds, rule.config?.icons]) if (entries !== undefined) {
    if (!Array.isArray(entries)) throw new TypeError('格式刷条件格式阈值无效。')
    for (const entry of entries) threshold(entry?.value)
  }
  assertJsonRule(rule)
}
function assertRangeBudget(ranges: readonly CellRange[]): void {
  if (ranges.length > MAXIMUM_RULE_RANGES) throw new RangeError(`格式刷当前最多同步处理 ${MAXIMUM_RULE_RANGES.toLocaleString()} 个条件格式范围。`)
}
function intersect(first: CellRange, second: CellRange): CellRange | undefined {
  return sheetRangesIntersect(first, second) ? { startRow: Math.max(first.startRow, second.startRow), endRow: Math.min(first.endRow, second.endRow),
    startCol: Math.max(first.startCol, second.startCol), endCol: Math.min(first.endCol, second.endCol) } : undefined
}
function subtract(ranges: readonly CellRange[], removed: readonly CellRange[], work: RangeWork): CellRange[] {
  let remaining = ranges.map(range => ({ ...range }))
  for (const cut of removed) {
    useRangeWork(work, remaining.length)
    remaining = remaining.flatMap(range => subtractSheetCommandRange(range, cut)); assertRangeBudget(remaining)
    if (!remaining.length) break
  }
  return remaining
}
/** Row events and compressed column coverage retain partial tiles and merge adjacent coverage deterministically. */
function normalizeRanges(ranges: readonly CellRange[], work: RangeWork): CellRange[] {
  assertRangeBudget(ranges)
  if (!ranges.length) return []
  const boundaries = [...new Set(ranges.flatMap(range => [range.startCol, range.endCol + 1]))].sort((left, right) => left - right)
  const indices = new Map(boundaries.map((column, index) => [column, index]))
  const events = ranges.flatMap(range => [
    { row: range.startRow, from: indices.get(range.startCol)!, to: indices.get(range.endCol + 1)! - 1, delta: 1 },
    { row: range.endRow + 1, from: indices.get(range.startCol)!, to: indices.get(range.endCol + 1)! - 1, delta: -1 },
  ]).sort((left, right) => left.row - right.row)
  const last = boundaries.length - 2, cover = new Int32Array(boundaries.length * 4), state = new Uint8Array(boundaries.length * 4)
  const update = (node: number, left: number, right: number, from: number, to: number, delta: number): void => {
    useRangeWork(work)
    if (from <= left && right <= to) cover[node] = cover[node]! + delta
    else {
      const middle = Math.floor((left + right) / 2)
      if (from <= middle) update(node * 2, left, middle, from, to, delta)
      if (to > middle) update(node * 2 + 1, middle + 1, right, from, to, delta)
    }
    state[node] = cover[node]! > 0 ? 2 : left === right ? 0
      : state[node * 2] === 2 && state[node * 2 + 1] === 2 ? 2 : state[node * 2] || state[node * 2 + 1] ? 1 : 0
  }
  const columns = (node: number, left: number, right: number, output: Array<{ startCol: number; endCol: number }>): void => {
    useRangeWork(work)
    if (!state[node]) return
    if (state[node] === 2) {
      const startCol = boundaries[left]!, endCol = boundaries[right + 1]! - 1, previous = output.at(-1)
      if (previous?.endCol === startCol - 1) previous.endCol = endCol
      else output.push({ startCol, endCol })
      return
    }
    const middle = Math.floor((left + right) / 2)
    columns(node * 2, left, middle, output); columns(node * 2 + 1, middle + 1, right, output)
  }
  const result: CellRange[] = []
  let previous = new Map<string, CellRange>(), index = 0
  while (index < events.length) {
    const row = events[index]!.row
    while (events[index]?.row === row) { const event = events[index++]!; update(1, 0, last, event.from, event.to, event.delta) }
    const endRow = events[index]?.row
    if (endRow === undefined) break
    const intervals: Array<{ startCol: number; endCol: number }> = [], current = new Map<string, CellRange>()
    columns(1, 0, last, intervals)
    for (const interval of intervals) {
      const key = `${interval.startCol}:${interval.endCol}`, existing = previous.get(key)
      const part = existing && existing.endRow === row - 1 ? existing : { startRow: row, endRow: endRow - 1, ...interval }
      if (part === existing) part.endRow = endRow - 1
      else result.push(part)
      current.set(key, part)
    }
    assertRangeBudget(result); previous = current
  }
  return result.sort((left, right) => left.startRow - right.startRow || left.startCol - right.startCol || left.endRow - right.endRow || left.endCol - right.endCol)
}
function stableJson(value: unknown): string {
  return JSON.stringify(value, (_key, item: unknown) => item && typeof item === 'object' && !Array.isArray(item)
    ? Object.fromEntries(Object.entries(item).sort(([left], [right]) => left.localeCompare(right))) : item) ?? 'undefined'
}
function ruleDefinition(rule: SheetConditionalFormatRule): string {
  const origin = conditionalFormatFormulaAnchor(rule)
  const normalized = origin ? mapConditionalFormatFormulas(rule, formula => formulaRelativeReferenceSignature(formula, origin.row, origin.col)) : rule
  const { id: _id, ranges: _ranges, priority: _priority, formulaAnchor: _origin, ...definition } = normalized
  return stableJson(definition)
}
function sameCoverage(first: readonly CellRange[], second: readonly CellRange[], work: RangeWork): boolean {
  return stableJson(normalizeRanges(first, work)) === stableJson(normalizeRanges(second, work))
}
function losslessReanchor(rule: SheetConditionalFormatRule, ranges: CellRange[]): boolean {
  return ruleDefinition(reanchorRule(reanchorRule(rule, ranges), rule.ranges)) === ruleDefinition(rule)
}
function preserveEquivalentRule(original: SheetConditionalFormatRule, next: SheetConditionalFormatRule, work: RangeWork): SheetConditionalFormatRule {
  return original.priority === next.priority && ruleDefinition(original) === ruleDefinition(next) && sameCoverage(original.ranges, next.ranges, work)
    ? cloneModel(original) : next
}
export function assertConditionalFormatPainterSource(source: SheetConditionalFormatPainterSource): void {
  if (!source || typeof source.sheetId !== 'string' || !source.sheetId.trim() || !Array.isArray(source.rules)) throw new TypeError('格式刷条件格式快照无效。')
  if (source.rules.length > MAXIMUM_RULE_RANGES) throw new RangeError('格式刷条件格式规则数量超出当前同步预算。')
  const ids = new Set<string>()
  for (const rule of source.rules) { assertRule(rule); if (ids.has(rule.id)) throw new TypeError('格式刷条件格式规则标识重复。'); ids.add(rule.id) }
  assertRangeBudget(source.rules.flatMap(rule => rule.ranges))
}
export function captureConditionalFormatPainterSource(sheetId: string, rules: readonly SheetConditionalFormatRule[] | undefined, sampled: CellRange): SheetConditionalFormatPainterSource {
  assertSheetCommandRanges([sampled])
  const captured = (rules ?? []).filter(rule => { assertRule(rule); return rule.ranges.some(range => sheetRangesIntersect(range, sampled)) }).map(cloneModel)
  const source = { sheetId, rules: captured }; assertConditionalFormatPainterSource(source)
  return source
}
function copiedRuleRanges(rule: SheetConditionalFormatRule, sampled: CellRange, targets: readonly CellRange[], work: RangeWork): CellRange[] {
  const sourceRanges = normalizeRanges(rule.ranges.flatMap(range => { const part = intersect(range, sampled); return part ? [part] : [] }), work)
  if (!sourceRanges.length) return []
  const height = sampled.endRow - sampled.startRow + 1, width = sampled.endCol - sampled.startCol + 1
  let copied: CellRange[] = []
  for (const target of targets) {
    // Later selection entries replace the earlier tiling phase in their overlap.
    copied = subtract(copied, [target], work)
    for (let row = target.startRow; row <= target.endRow; row += height) for (let col = target.startCol; col <= target.endCol; col += width) {
      for (const source of sourceRanges) {
        useRangeWork(work)
        const projected = { startRow: row + source.startRow - sampled.startRow, endRow: row + source.endRow - sampled.startRow,
          startCol: col + source.startCol - sampled.startCol, endCol: col + source.endCol - sampled.startCol }
        const part = intersect(projected, target)
        if (part) copied.push(part)
      }
      assertRangeBudget(copied)
    }
  }
  return normalizeRanges(copied, work)
}

/** Format painter replaces the target rule scope; source definitions and their relative formulas are snapshotted. */
export function createConditionalFormatPainterPlan(
  sheetId: string, currentRules: readonly SheetConditionalFormatRule[] | undefined, source: SheetConditionalFormatPainterSource | undefined,
  sampled: CellRange, targets: readonly CellRange[],
): SheetConditionalFormatPainterPlan {
  if (!source || !targets.length) return { changed: false, rules: currentRules ? cloneModel([...currentRules]) : undefined }
  assertConditionalFormatPainterSource(source); assertSheetCommandRanges([sampled, ...targets])
  for (const rule of currentRules ?? []) assertRule(rule)
  assertRangeBudget((currentRules ?? []).flatMap(rule => rule.ranges))
  const original = cloneModel([...(currentRules ?? [])]), work: RangeWork = { operations: 0 }
  if (new Set(original.map(rule => rule.id)).size !== original.length) throw new TypeError('格式刷目标条件格式规则标识重复。')
  const remaining = new Map<string, CellRange[]>(), next = new Map<string, SheetConditionalFormatRule>()
  for (const rule of original) {
    const overlaps = rule.ranges.some(range => targets.some(target => { useRangeWork(work); return sheetRangesIntersect(range, target) }))
    const ranges = overlaps ? normalizeRanges(subtract(rule.ranges, targets, work), work) : rule.ranges
    remaining.set(rule.id, ranges)
    if (ranges.length) next.set(rule.id, overlaps ? preserveEquivalentRule(rule, reanchorRule(rule, ranges), work) : rule)
  }
  const additions = [...source.rules].sort((left, right) => left.priority - right.priority).flatMap(rule => {
    const ranges = copiedRuleRanges(rule, sampled, targets, work)
    return ranges.length ? [{ rule, ranges, proposed: reanchorRule(rule, ranges) }] : []
  })
  // Sharing some live source rules and cloning others would change their relative target priority.
  // Reuse the full sampled stack only when every definition and its formula anchor remain equivalent.
  const shared = additions.map(({ rule, ranges }) => {
    useRangeWork(work, original.length)
    const candidate = source.sheetId === sheetId ? original.find(item => item.id === rule.id && item.priority === rule.priority
      && ruleDefinition(reanchorRule(item, rule.ranges)) === ruleDefinition(rule)) : undefined
    const combined = candidate ? normalizeRanges([...(remaining.get(candidate.id) ?? []), ...ranges], work) : []
    return candidate && losslessReanchor(candidate, combined) ? candidate : undefined
  })
  const reuseShared = shared.every(rule => rule !== undefined) && shared.every((rule, index) => index === 0
    || rule!.priority > shared[index - 1]!.priority
    || rule!.priority === shared[index - 1]!.priority && original.indexOf(rule!) > original.indexOf(shared[index - 1]!))
  const consumed = new Set<string>()
  const repeated = reuseShared ? [] : additions.map(({ ranges, proposed }) => {
    useRangeWork(work, original.length)
    const candidate = original.find(rule => !consumed.has(rule.id) && !(remaining.get(rule.id)?.length)
      && ruleDefinition(reanchorRule(rule, ranges)) === ruleDefinition(proposed))
    if (candidate) consumed.add(candidate.id)
    return candidate
  })
  const reuseRepeated = repeated.every(rule => rule !== undefined) && repeated.every((rule, index) => index === 0
    || rule!.priority > repeated[index - 1]!.priority
    || rule!.priority === repeated[index - 1]!.priority && original.indexOf(rule!) > original.indexOf(repeated[index - 1]!))
  const created: SheetConditionalFormatRule[] = []
  let priority = original.reduce((maximum, rule) => Math.max(maximum, rule.priority), 0)
  for (const [index, addition] of additions.entries()) {
    const reused = reuseShared ? shared[index] : reuseRepeated ? repeated[index] : undefined
    const ranges = reused ? normalizeRanges([...(remaining.get(reused.id) ?? []), ...addition.ranges], work) : addition.ranges
    const painted = reused ? preserveEquivalentRule(reused, reanchorRule(reused, ranges), work)
      : { ...addition.proposed, ranges, id: `painter-cf-${crypto.randomUUID()}`, priority: ++priority }
    if (reused) next.set(reused.id, painted)
    else created.push(painted)
  }
  // Fully replaced rules keep their original list position when their ID is reused.
  const resultRules = [...original.flatMap(rule => { const replacement = next.get(rule.id); return replacement ? [replacement] : [] }), ...created]
  assertRangeBudget(resultRules.flatMap(rule => rule.ranges))
  const result = resultRules.length || currentRules !== undefined ? resultRules : undefined
  return { changed: stableJson(result) !== stableJson(currentRules), rules: result }
}
