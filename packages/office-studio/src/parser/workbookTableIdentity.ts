import type { WorkbookModel } from '../domain/model.ts'
import type { FormulaCalendarTableDto } from './contracts.ts'

const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)

/** Prefer admitted identities, reserving them before assigning foreign native
 * objects. A new Office table/column cannot steal a restored domain identity. */
function uniqueIds(current: readonly string[], restored: readonly (string | undefined)[], fallback: (index: number) => string): string[] {
  const savedIds = restored.filter((id): id is string => id !== undefined)
  if (new Set(current).size !== current.length || new Set(savedIds).size !== savedIds.length) {
    throw new TypeError('TABLE_SOURCE_INVALID: duplicate table or column identity')
  }
  const occupied = new Set([...current, ...savedIds])
  const claimed = new Set(savedIds)
  return current.map((id, index) => {
    const saved = restored[index]
    if (saved !== undefined) return saved
    if (!claimed.has(id)) { claimed.add(id); return id }
    const base = fallback(index)
    let next = base, suffix = 1
    while (occupied.has(next)) next = `${base}:${suffix++}`
    occupied.add(next); claimed.add(next)
    return next
  })
}

export function restoreWorkbookTableIdentities(workbook: WorkbookModel, rawTables: unknown,
  sources: readonly FormulaCalendarTableDto[], sheetIds: ReadonlyMap<string, string>): void {
  const tables = workbook.tables ?? []
  const native: unknown[] = Array.isArray(rawTables) ? rawTables : []
  const byNativeId = new Map(sources.map(table => [table.officeTableId, table]))
  const nativeIds = tables.map((table, index) => {
    const raw = native[index]
    return table.officeTableId ?? (object(raw) && typeof raw.officeTableId === 'number' ? raw.officeTableId : undefined)
  })
  const saved = tables.map((table, index) => {
    const source = nativeIds[index] === undefined ? undefined : byNativeId.get(nativeIds[index]!)
    return source && source.name.toLowerCase() === table.name.toLowerCase() && sheetIds.get(source.sheetId) === table.sheetId ? source : undefined
  })
  const tableIds = uniqueIds(tables.map(table => table.id), saved.map(table => table?.tableId), index => `office-table:${nativeIds[index] ?? index + 1}`)
  const plans = tables.map((table, index) => {
    const restored = new Map(saved[index]?.columns?.map(column => [column.officeColumnId, column.columnId]) ?? [])
    const columnIds = uniqueIds(table.columns.map(column => column.id), table.columns.map(column =>
      column.officeColumnId === undefined ? undefined : restored.get(column.officeColumnId)),
    column => `${tableIds[index]}:office-column:${table.columns[column]!.officeColumnId ?? column + 1}`)
    return { table, previousId: table.id, id: tableIds[index]!, columns: new Map(table.columns.map((column, columnIndex) => [column.id, columnIds[columnIndex]!])) }
  })
  const byPreviousId = new Map(plans.map(plan => [plan.previousId, plan]))
  // All identities have been resolved before touching any table or binding.
  for (const plan of plans) {
    plan.table.id = plan.id
    for (const column of plan.table.columns) column.id = plan.columns.get(column.id)!
  }
  for (const slicer of workbook.slicers ?? []) if (slicer.tableId) {
    const plan = byPreviousId.get(slicer.tableId)
    if (plan) { slicer.tableId = plan.id; slicer.sourceFieldId = plan.columns.get(slicer.sourceFieldId) ?? slicer.sourceFieldId }
  }
  for (const cache of workbook.slicerCaches ?? []) if (cache.source.kind === 'table') {
    const plan = byPreviousId.get(cache.source.tableId)
    if (plan) { cache.source.tableId = plan.id; cache.source.columnId = plan.columns.get(cache.source.columnId) ?? cache.source.columnId }
  }
}
