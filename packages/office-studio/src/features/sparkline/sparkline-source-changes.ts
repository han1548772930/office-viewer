import type { CellRange, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { mergeSheetAxisIntervals, type SheetAxisInterval } from '../../domain/sheet-axis-structure.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { sheetAxisViewHiddenSpans } from '../sheet-workbench/sheet-row-visibility.ts'
import type { SparklineGroupInfo } from './sparkline-model.ts'
import { parseSparklineReference } from './sparkline-reference.ts'
import { resolveSparklineSource } from './sparkline-source.ts'

interface SourceChanges {
  rows: number[]
  columnsByRow: Map<number, number[]>
  beforeRows: SheetAxisInterval[]
  afterRows: SheetAxisInterval[]
  beforeColumns: SheetAxisInterval[]
  afterColumns: SheetAxisInterval[]
}

function sampleInput(cell: SpreadsheetCell | undefined): object {
  return { value: cell?.value ?? null, formula: cell?.formula ?? null, dataType: cell?.dataType,
    cachedDataType: cell?.formula ? cell.cachedDataType : undefined,
    cachedDisplay: cell?.formula && cell.value === null ? cell.display : undefined }
}
function visibility(sheet: SpreadsheetSheet, axis: 'row' | 'column'): SheetAxisInterval[] {
  const points = axis === 'row' ? [...(sheet.hiddenRows ?? []), ...(sheet.autoFilter?.filteredRows ?? [])] : sheet.hiddenColumns ?? []
  return mergeSheetAxisIntervals([
    ...sheetAxisViewHiddenSpans(sheet, axis).map(({ start, end }) => ({ start, end })),
    ...points.map(index => ({ start: index, end: index })),
  ])
}
function changes(before: SpreadsheetSheet, after: SpreadsheetSheet): SourceChanges {
  const columnsByRow = new Map<number, number[]>()
  const add = (cell: SpreadsheetCell): void => { const columns = columnsByRow.get(cell.row) ?? []; columns.push(cell.col); columnsByRow.set(cell.row, columns) }
  for (const [key, cell] of Object.entries(before.cells)) if (!equivalentModelData(sampleInput(cell), sampleInput(after.cells[key]))) add(cell)
  for (const [key, cell] of Object.entries(after.cells)) if (!Object.hasOwn(before.cells, key) && !equivalentModelData(sampleInput(undefined), sampleInput(cell))) add(cell)
  for (const columns of columnsByRow.values()) columns.sort((a, b) => a - b)
  return { rows: [...columnsByRow.keys()].sort((a, b) => a - b), columnsByRow,
    beforeRows: visibility(before, 'row'), afterRows: visibility(after, 'row'),
    beforeColumns: visibility(before, 'column'), afterColumns: visibility(after, 'column') }
}
function lowerBound(values: readonly number[], minimum: number): number {
  let start = 0, end = values.length
  while (start < end) { const middle = (start + end) >>> 1; if (values[middle]! < minimum) start = middle + 1; else end = middle }
  return start
}
function changedValues(delta: SourceChanges, range: CellRange): boolean {
  for (let index = lowerBound(delta.rows, range.startRow); index < delta.rows.length && delta.rows[index]! <= range.endRow; index++) {
    const columns = delta.columnsByRow.get(delta.rows[index]!)!
    const column = columns[lowerBound(columns, range.startCol)]
    if (column !== undefined && column <= range.endCol) return true
  }
  return false
}
function clip(intervals: readonly SheetAxisInterval[], start: number, end: number): SheetAxisInterval[] {
  const result: SheetAxisInterval[] = []
  for (const interval of intervals) {
    if (interval.end < start) continue
    if (interval.start > end) break
    result.push({ start: Math.max(start, interval.start), end: Math.min(end, interval.end) })
  }
  return result
}
/** One sparse delta per source sheet; never enumerate every cell in an A1 range. */
export function createSparklineSourceChangeDetector(before: WorkbookModel, after: WorkbookModel): (hostSheetId: string, group: SparklineGroupInfo) => boolean {
  const cache = new Map<string, SourceChanges>()
  return (hostSheetId, group) => group.sourceRanges.some((sourceRange, index) => {
    const oldHost = before.sheets.find(sheet => sheet.id === hostSheetId), host = after.sheets.find(sheet => sheet.id === hostSheetId)
    if (!oldHost || !host) return oldHost !== host
    const member = group.members[index], target = member && parseSparklineReference(member.address)
    const key = target ? `${target.range.startRow}:${target.range.startCol}` : ''
    const oldSource = resolveSparklineSource({ sourceRange, sourceReference: oldHost.cells[key]?.sparkline?.sourceReference }, oldHost, before)
    const source = resolveSparklineSource({ sourceRange, sourceReference: host.cells[key]?.sparkline?.sourceReference }, host, after)
    if (oldSource.status !== 'resolved' || source.status !== 'resolved') {
      return oldSource.status !== source.status || oldSource.status !== 'resolved' && source.status !== 'resolved' && oldSource.reason !== source.reason
    }
    const previous = oldSource.sheet, next = source.sheet
    if (previous.id !== next.id || !equivalentModelData(oldSource.range, source.range)) return true
    let delta = cache.get(next.id)
    if (!delta) { delta = changes(previous, next); cache.set(next.id, delta) }
    const { range } = source
    return changedValues(delta, range) || !group.settings.containHiddenCells && (
      !equivalentModelData(clip(delta.beforeRows, range.startRow, range.endRow), clip(delta.afterRows, range.startRow, range.endRow))
      || !equivalentModelData(clip(delta.beforeColumns, range.startCol, range.endCol), clip(delta.afterColumns, range.startCol, range.endCol)))
  })
}
