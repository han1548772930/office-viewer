import type { CellRange } from '../../domain/model.ts'
import { sheetRangeContainsCell } from './sheet-range-command.ts'
import { cloneSheetSelection, createSheetSelectionAxisIndexes, normalizeSheetSelection, resolveVisibleSheetSelection, type SheetSelectionCell, type SheetSelectionSheet, type SheetSelectionState } from './sheet-selection.ts'

export const SheetSelectionDirections = Object.freeze({ UP: 0, RIGHT: 1, DOWN: 2, LEFT: 3 } as const)
export type SheetSelectionDirection = typeof SheetSelectionDirections[keyof typeof SheetSelectionDirections]
export interface SheetSelectionNavigationResult { range: CellRange; primary: SheetSelectionCell; activeRangeIndex: number }
export type SheetSelectionNavigationSheet = SheetSelectionSheet & { mergedRanges?: readonly CellRange[] }
export type SheetSelectionShortcutKey = 'Enter' | 'Tab'
export interface SheetSelectionShortcutPlan { selection: SheetSelectionState; tabStart?: SheetSelectionState }

function assertDirection(direction: SheetSelectionDirection): void {
  if (direction !== 0 && direction !== 1 && direction !== 2 && direction !== 3) throw new TypeError('选区移动方向无效。')
}

function sameRange(first: CellRange, second: CellRange): boolean {
  return first.startRow === second.startRow && first.endRow === second.endRow && first.startCol === second.startCol && first.endCol === second.endCol
}

/** Tab visits rows; Enter visits columns. Hidden runs and merged interiors are skipped by interval bounds. */
export function nextSheetSelectionCell(state: SheetSelectionState, direction: SheetSelectionDirection,
  sheet: SheetSelectionNavigationSheet, filteredRows: readonly number[] = []): SheetSelectionNavigationResult | null {
  assertDirection(direction)
  const selection = resolveVisibleSheetSelection(state, sheet, filteredRows)
  if (!selection.primary) return null
  const { rows, columns } = createSheetSelectionAxisIndexes(sheet, filteredRows)
  const horizontal = direction === 1 || direction === 3, step = direction === 0 || direction === 3 ? -1 : 1
  const primaryAxis = horizontal ? columns : rows, secondaryAxis = horizontal ? rows : columns
  const mergedAt = (row: number, column: number) => sheet.mergedRanges?.find(range => sheetRangeContainsCell(range, row, column))
  const mergePrimary = (merge: CellRange): SheetSelectionCell | undefined => {
    const row = rows.nextVisible(merge.startRow, 1), column = columns.nextVisible(merge.startCol, 1)
    return row !== undefined && row <= merge.endRow && column !== undefined && column <= merge.endCol ? { row, column } : undefined
  }
  const result = (point: SheetSelectionCell, index: number): SheetSelectionNavigationResult => {
    const merge = mergedAt(point.row, point.column)
    return { range: merge ? { ...merge } : { startRow: point.row, endRow: point.row, startCol: point.column, endCol: point.column },
      primary: { ...point }, activeRangeIndex: index }
  }
  const within = (index: number, current?: SheetSelectionCell): SheetSelectionNavigationResult | undefined => {
    const range = selection.ranges[index]!
    const primaryStart = horizontal ? range.startCol : range.startRow, primaryEnd = horizontal ? range.endCol : range.endRow
    const secondaryStart = horizontal ? range.startRow : range.startCol, secondaryEnd = horizontal ? range.endRow : range.endCol
    const primaryEdge = step === 1 ? primaryStart : primaryEnd, secondaryEdge = step === 1 ? secondaryStart : secondaryEnd
    const insidePrimary = (value: number | undefined): value is number => value !== undefined && value >= primaryStart && value <= primaryEnd
    const insideSecondary = (value: number | undefined): value is number => value !== undefined && value >= secondaryStart && value <= secondaryEnd
    const edge = primaryAxis.nextVisible(primaryEdge, step)
    if (!insidePrimary(edge)) return undefined
    let secondary = current ? horizontal ? current.row : current.column : secondaryAxis.nextVisible(secondaryEdge, step)
    const currentMerge = current ? mergedAt(current.row, current.column) : undefined
    const currentPrimary = current ? horizontal ? current.column : current.row : undefined
    const currentBound = currentMerge
      ? horizontal ? step === 1 ? currentMerge.endCol : currentMerge.startCol : step === 1 ? currentMerge.endRow : currentMerge.startRow
      : currentPrimary
    let primary = currentBound === undefined ? edge : primaryAxis.nextVisible(currentBound + step, step)
    while (insideSecondary(secondary)) {
      let secondaryJump: number | undefined
      while (insidePrimary(primary)) {
        const point = horizontal ? { row: secondary, column: primary } : { row: primary, column: secondary }
        const merge = mergedAt(point.row, point.column)
        const representative = merge ? mergePrimary(merge) : undefined
        if (!merge || representative && point.row === representative.row && point.column === representative.column) return result(point, index)
        const mergePrimaryStart = horizontal ? merge.startCol : merge.startRow, mergePrimaryEnd = horizontal ? merge.endCol : merge.endRow
        const mergeSecondaryStart = horizontal ? merge.startRow : merge.startCol
        const representativePrimary = representative ? horizontal ? representative.column : representative.row : undefined
        const representativeSecondary = representative ? horizontal ? representative.row : representative.column : undefined
        if (step === -1 && representative && secondary === representativeSecondary && insidePrimary(representativePrimary)) return result(representative, index)
        if (mergePrimaryStart <= primaryStart && mergePrimaryEnd >= primaryEnd) {
          secondaryJump = step === 1 ? (horizontal ? merge.endRow : merge.endCol) + 1
            : representativeSecondary !== undefined && representativeSecondary < secondary ? representativeSecondary : mergeSecondaryStart - 1
          break
        }
        primary = primaryAxis.nextVisible(step === 1 ? mergePrimaryEnd + 1 : mergePrimaryStart - 1, step)
      }
      secondary = secondaryAxis.nextVisible(secondaryJump ?? secondary + step, step)
      primary = edge
    }
    return undefined
  }
  const next = within(selection.activeRangeIndex, selection.primary)
  if (next) return next
  for (let offset = 1; offset <= selection.ranges.length; offset++) {
    const index = (selection.activeRangeIndex + step * offset + selection.ranges.length) % selection.ranges.length
    const candidate = within(index)
    if (candidate) return candidate
  }
  return null
}

/** Shared keyboard plan. A range cycles within its selected cells; an isolated cell advances in the sheet. */
export function planSheetSelectionShortcut(state: SheetSelectionState, direction: SheetSelectionDirection, key: SheetSelectionShortcutKey,
  sheet: SheetSelectionNavigationSheet, tabStart?: SheetSelectionState, filteredRows: readonly number[] = []): SheetSelectionShortcutPlan | null {
  assertDirection(direction)
  if (key !== 'Enter' && key !== 'Tab' || key === 'Tab' && direction !== 1 && direction !== 3 || key === 'Enter' && direction !== 0 && direction !== 2) {
    throw new TypeError('选区移动按键与方向不一致。')
  }
  const selection = resolveVisibleSheetSelection(state, sheet, filteredRows)
  if (!selection.primary) return null
  const current = selection.primary, active = selection.ranges[selection.activeRangeIndex]!
  const merge = sheet.mergedRanges?.find(range => sheetRangeContainsCell(range, current.row, current.column))
  const single = { startRow: current.row, endRow: current.row, startCol: current.column, endCol: current.column }
  if (selection.ranges.length > 1 || !sameRange(active, single) && (!merge || !sameRange(active, merge))) {
    const next = nextSheetSelectionCell(selection, direction, sheet, filteredRows)
    return { selection: next ? { ranges: selection.ranges, activeRangeIndex: next.activeRangeIndex, primary: next.primary } : selection }
  }
  const origin = key === 'Enter' && tabStart ? resolveVisibleSheetSelection(tabStart, sheet, filteredRows) : selection
  const start = origin.primary ?? current
  const startMerge = sheet.mergedRanges?.find(range => sheetRangeContainsCell(range, start.row, start.column))
  const { rows, columns } = createSheetSelectionAxisIndexes(sheet, filteredRows)
  const horizontal = direction === 1 || direction === 3, step = direction === 0 || direction === 3 ? -1 : 1
  const bound = startMerge ? horizontal ? step === 1 ? startMerge.endCol : startMerge.startCol : step === 1 ? startMerge.endRow : startMerge.startRow
    : horizontal ? start.column : start.row
  const nextIndex = (horizontal ? columns : rows).nextVisible(bound + step, step)
  const remembered = key === 'Tab' ? tabStart ?? (direction === 1 ? cloneSheetSelection(selection) : undefined) : undefined
  if (nextIndex === undefined) return { selection, ...(remembered ? { tabStart: cloneSheetSelection(remembered) } : {}) }
  const point = horizontal ? { row: start.row, column: nextIndex } : { row: nextIndex, column: start.column }
  const targetMerge = sheet.mergedRanges?.find(range => sheetRangeContainsCell(range, point.row, point.column))
  if (targetMerge) {
    point.row = rows.nextVisible(targetMerge.startRow, 1) ?? point.row
    point.column = columns.nextVisible(targetMerge.startCol, 1) ?? point.column
  }
  const range = targetMerge ? { ...targetMerge, rangeType: 0 as const }
    : { startRow: point.row, endRow: point.row, startCol: point.column, endCol: point.column, rangeType: 0 as const }
  return { selection: normalizeSheetSelection([range], sheet, { activeRangeIndex: 0, primary: point }),
    ...(remembered ? { tabStart: cloneSheetSelection(remembered) } : {}) }
}
