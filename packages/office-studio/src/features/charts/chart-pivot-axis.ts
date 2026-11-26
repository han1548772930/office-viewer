import type { PivotAxisOutputEntry, PivotFieldModel } from '../../domain/model.ts'

/** Runtime category identity, aligned to normalized records rather than display-text matching. */
export interface PivotChartAxisMetadata {
  rowFields: Array<{ id: string; name: string; field: number }>
  columnFields: Array<{ id: string; name: string }>
  recordRows: PivotAxisOutputEntry[]
  valueColumns: PivotAxisOutputEntry[]
}

export function pivotChartAxisMetadata(rows: readonly PivotFieldModel[], columns: readonly PivotFieldModel[],
  records: readonly PivotAxisOutputEntry[], values: readonly PivotAxisOutputEntry[], rowFieldIndexes?: readonly number[]): PivotChartAxisMetadata {
  const field = (field: PivotFieldModel) => ({ id: field.id, name: field.displayName ?? field.name })
  const entry = (entry: PivotAxisOutputEntry): PivotAxisOutputEntry => ({ ...entry, labels: [...entry.labels],
    fieldFilters: { ...entry.fieldFilters }, fieldMembers: Object.fromEntries(Object.entries(entry.fieldMembers).map(([id, members]) => [id, [...members]])) })
  return { rowFields: rows.map((row, index) => ({ ...field(row), field: rowFieldIndexes?.[index] ?? index })),
    columnFields: columns.map(field), recordRows: records.map(entry), valueColumns: values.map(entry) }
}
