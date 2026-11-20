import type { CellRange, PivotAggregation, PivotFieldArea, PivotFieldModel, PivotTableModel, PivotTableOptions, SpreadsheetSheet, WorkbookModel } from '../../../domain/model'
import { parseSheetAddress } from '../../sheet-workbench/sheet-workbench.ts'
import { createPivotTable, normalizePivotSource, refreshPivotTable, type PivotTableAnchorInput, type PivotTableSourceInput } from './pivot-engine.ts'

export const CONFIGURE_PIVOT_COMMAND = 'sheet.command.configure-pivot-table'
export interface PivotPanelField {
  fieldId: string
  fieldName: string
  sourceColumn: number
  area: PivotFieldArea | 'unused'
  aggregation: PivotAggregation
}
export interface PivotPanelDraft {
  pivotId: string
  sourceSheetId: string
  sourceRange: string
  targetSheetId: string
  targetAddress: string
  fields: PivotPanelField[]
  options: PivotTableOptions
}
export interface ConfigurePivotFromPanelInput {
  unitId: string
  pivotTableId: string
  mode: 'create' | 'update'
  source: PivotTableSourceInput
  anchor: PivotTableAnchorInput
  fields: Array<{ fieldId: string; area: PivotFieldArea; aggregation?: PivotAggregation }>
  options: Partial<PivotTableOptions>
}
const maxRows = 1_048_576, maxColumns = 16_384
const areas: readonly PivotFieldArea[] = ['filter', 'row', 'column', 'value']
const aggregations: readonly PivotAggregation[] = ['sum', 'count', 'average', 'min', 'max']
const defaultOptions = (): PivotTableOptions => ({ repeatRowLabels: false, showRowGrandTotal: true, showColGrandTotal: true, showRowSubtotals: true, showColSubtotals: true, valuePosition: 'column' })

function columnLabel(column: number): string {
  let label = '', next = column + 1
  while (next > 0) { next -= 1; label = String.fromCharCode(next % 26 + 65) + label; next = Math.floor(next / 26) }
  return label
}
export function pivotRangeAddress(range: CellRange): string {
  const first = `${columnLabel(range.startCol)}${range.startRow + 1}`, last = `${columnLabel(range.endCol)}${range.endRow + 1}`
  return first === last ? first : `${first}:${last}`
}
function requireSheet(workbook: WorkbookModel, sheetId: string): SpreadsheetSheet {
  const sheet = workbook.sheets.find(candidate => candidate.id === sheetId)
  if (!sheet) throw new Error('找不到所选工作表，请重新选择数据源或放置位置。')
  return sheet
}
function sourceRange(workbook: WorkbookModel, sheetId: string, address: string): { sheet: SpreadsheetSheet; range: CellRange } {
  const sheet = requireSheet(workbook, sheetId)
  const parsed = parseSheetAddress(address, { rowCount: maxRows, columnCount: maxColumns })
  if (!parsed || !/^\$?[A-Z]+\$?[1-9]\d*(?::\$?[A-Z]+\$?[1-9]\d*)?$/i.test(address.trim())) throw new Error('请输入有效的数据区域，例如 A1:C20。')
  if (parsed.range.startRow === parsed.range.endRow) throw new Error('数据区域需要包含一行字段标题和至少一行数据。')
  return { sheet, range: { startRow: parsed.range.startRow, startCol: parsed.range.startColumn, endRow: parsed.range.endRow, endCol: parsed.range.endColumn } }
}
export function readPivotPanelSourceFields(workbook: WorkbookModel, sheetId: string, address: string): PivotPanelField[] {
  const { sheet, range } = sourceRange(workbook, sheetId, address)
  return Array.from({ length: range.endCol - range.startCol + 1 }, (_, index) => {
    const sourceColumn = range.startCol + index, header = sheet.cells[`${range.startRow}:${sourceColumn}`]?.value
    return { fieldId: String(index), fieldName: header == null || header === '' ? `Column ${index + 1}` : String(header), sourceColumn, area: 'unused', aggregation: 'sum' }
  })
}
export function createPivotPanelDraft(workbook: WorkbookModel, sheetId: string, address: string, pivotId = ''): PivotPanelDraft {
  const pivot = workbook.pivotTables?.find(candidate => candidate.id === pivotId)
  const draft: PivotPanelDraft = {
    pivotId: pivot?.id ?? '', sourceSheetId: pivot?.source.subUnitId ?? sheetId,
    sourceRange: pivot ? pivotRangeAddress(pivot.source.range) : address,
    targetSheetId: pivot?.anchor.subUnitId ?? sheetId,
    targetAddress: pivot ? `${columnLabel(pivot.anchor.col)}${pivot.anchor.row + 1}` : 'A1',
    fields: [], options: { ...(pivot?.options ?? defaultOptions()) },
  }
  try {
    draft.fields = readPivotPanelSourceFields(workbook, draft.sourceSheetId, draft.sourceRange)
    if (!pivot) {
      const range = sourceRange(workbook, sheetId, address).range
      draft.targetAddress = range.endCol + 2 < maxColumns ? `${columnLabel(range.endCol + 2)}${range.startRow + 1}` : `A${range.endRow + 3}`
    } else {
      draft.fields = draft.fields.map(field => {
        const existing = pivot.fields.find(candidate => candidate.id === field.fieldId)
        return existing ? { ...field, area: existing.area, aggregation: existing.aggregation ?? 'sum' } : field
      }).sort((a, b) => {
        const first = pivot.fields.find(candidate => candidate.id === a.fieldId), second = pivot.fields.find(candidate => candidate.id === b.fieldId)
        return first && second && first.area === second.area ? first.position - second.position : 0
      })
    }
  } catch { /* Keep an invalid current selection editable; validation is shown on field loading/submission. */ }
  return draft
}

export function preparePivotPanelInput(workbook: WorkbookModel, draft: PivotPanelDraft, pivotTableId: string): ConfigurePivotFromPanelInput {
  if (!workbook.id) throw new Error('工作簿没有稳定标识，无法创建数据透视表。')
  const { sheet, range } = sourceRange(workbook, draft.sourceSheetId, draft.sourceRange)
  requireSheet(workbook, draft.targetSheetId)
  const target = parseSheetAddress(draft.targetAddress, { rowCount: maxRows, columnCount: maxColumns })
  if (!target || target.range.startRow !== target.range.endRow || target.range.startColumn !== target.range.endColumn) throw new Error('放置位置必须是单个单元格，例如 E1。')
  const available = readPivotPanelSourceFields(workbook, draft.sourceSheetId, draft.sourceRange)
  const fields = draft.fields.filter(field => field.area !== 'unused').map(field => {
    if (!available.some(candidate => candidate.fieldId === field.fieldId)) throw new Error(`字段“${field.fieldName}”已不在当前数据区域中，请重新读取字段。`)
    if (!areas.some(area => area === field.area)) throw new Error('字段区域无效。')
    const area: PivotFieldArea = field.area === 'unused' ? 'row' : field.area
    return { fieldId: field.fieldId, area, ...(area === 'value' ? { aggregation: field.aggregation } : {}) }
  })
  const input: ConfigurePivotFromPanelInput = {
    unitId: workbook.id, pivotTableId, mode: draft.pivotId ? 'update' : 'create',
    source: { unitId: workbook.id, subUnitId: sheet.id, sheetName: sheet.name, range: { startRow: range.startRow, startColumn: range.startCol, endRow: range.endRow, endColumn: range.endCol } },
    anchor: { unitId: workbook.id, subUnitId: draft.targetSheetId, row: target.range.startRow, col: target.range.startColumn },
    fields, options: { ...draft.options },
  }
  validateConfiguration(workbook, input)
  return input
}

function validateConfiguration(workbook: WorkbookModel, input: ConfigurePivotFromPanelInput): void {
  if (!input.pivotTableId?.trim() || workbook.id !== input.unitId || input.anchor.unitId !== input.unitId) throw new Error('数据透视表所属工作簿不匹配。')
  const source = normalizePivotSource(input.source, workbook)
  requireSheet(workbook, input.anchor.subUnitId)
  if (![input.anchor.row, input.anchor.col].every(value => Number.isSafeInteger(value) && value >= 0) || input.anchor.row >= maxRows || input.anchor.col >= maxColumns) throw new Error('放置位置超出工作表有效范围。')
  if (input.anchor.subUnitId === source.subUnitId && input.anchor.row >= source.range.startRow && input.anchor.row <= source.range.endRow && input.anchor.col >= source.range.startCol && input.anchor.col <= source.range.endCol) throw new Error('数据透视表的放置位置不能与数据源重叠。')
  const existing = workbook.pivotTables?.find(candidate => candidate.id === input.pivotTableId)
  if (input.mode === 'update' && !existing) throw new Error('要修改的数据透视表已被删除，请重新选择。')
  if (input.mode === 'create' && existing) throw new Error('数据透视表标识已存在，请重新创建。')
  const seen = new Set<string>()
  for (const field of input.fields) {
    if (!/^\d+$/.test(field.fieldId) || Number(field.fieldId) > source.range.endCol - source.range.startCol || !areas.includes(field.area)) throw new Error('配置中存在无效字段。')
    if (seen.has(field.fieldId)) throw new Error('同一个字段不能重复放置。')
    seen.add(field.fieldId)
    if (field.area === 'value' && !aggregations.includes(field.aggregation ?? 'sum')) throw new Error('请选择有效的汇总方式。')
  }
  if (!input.fields.some(field => field.area === 'value')) throw new Error('请至少将一个字段放入“值”区域。')
  for (const key of ['repeatRowLabels', 'showRowGrandTotal', 'showColGrandTotal', 'showRowSubtotals', 'showColSubtotals'] as const) {
    if (input.options[key] !== undefined && typeof input.options[key] !== 'boolean') throw new Error('汇总选项必须是开关值。')
  }
  if (input.options.valuePosition !== undefined && !['row', 'column'].includes(input.options.valuePosition)) throw new Error('值字段排列方向无效。')
}

/** Mutates the command's detached workbook draft; OfficeStudio commits it as one history transaction. */
export function configurePivotFromPanel(workbook: WorkbookModel, input: ConfigurePivotFromPanelInput): PivotTableModel {
  validateConfiguration(workbook, input)
  let pivot = workbook.pivotTables?.find(candidate => candidate.id === input.pivotTableId)
  const previous = pivot
  const sourceUnchanged = previous && previous.source.subUnitId === input.source.subUnitId && previous.source.range.startRow === input.source.range.startRow && previous.source.range.endRow === input.source.range.endRow && previous.source.range.startCol === input.source.range.startColumn && previous.source.range.endCol === input.source.range.endColumn
  if (!pivot) pivot = createPivotTable(workbook, input.source, input.anchor, input.pivotTableId)
  else {
    const oldTarget = requireSheet(workbook, pivot.anchor.subUnitId)
    for (const key of pivot.generatedCellKeys ?? []) if (oldTarget.cells[key]?.metadata?.pivotTableId === pivot.id) delete oldTarget.cells[key]
    pivot.generatedCellKeys = []
    pivot.source = normalizePivotSource(input.source, workbook)
    pivot.anchor = { ...input.anchor }
  }
  const positions: Record<PivotFieldArea, number> = { row: 0, column: 0, filter: 0, value: 0 }
  const source = requireSheet(workbook, input.source.subUnitId)
  pivot.fields = input.fields.map(field => {
    const old = sourceUnchanged ? previous?.fields.find(candidate => candidate.id === field.fieldId) : undefined
    const sourceColumn = input.source.range.startColumn + Number(field.fieldId), header = source.cells[`${input.source.range.startRow}:${sourceColumn}`]?.value
    const next: PivotFieldModel = {
      ...old, id: field.fieldId, name: header == null || header === '' ? `Column ${Number(field.fieldId) + 1}` : String(header),
      sourceColumn, area: field.area, position: positions[field.area]++,
      ...(field.area === 'value' ? { aggregation: field.aggregation ?? 'sum' } : {}),
    }
    if (field.area !== 'value') delete next.aggregation
    return next
  })
  pivot.options = { ...pivot.options, ...input.options }
  refreshPivotTable(workbook, pivot)
  if (pivot.outputRange && (pivot.outputRange.endRow >= maxRows || pivot.outputRange.endCol >= maxColumns)) throw new Error('数据透视表结果超出工作表有效范围，请调整放置位置。')
  return pivot
}
