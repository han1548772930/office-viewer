import type { CellRange, PivotAggregation, PivotFieldArea, PivotFieldModel, PivotShowDataAsInfo, PivotSortDirection, PivotTableModel, PivotTableOptions, SpreadsheetSheet, WorkbookModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization.ts'
import { parseSheetAddress } from '../../sheet-workbench/sheet-workbench.ts'
import { createDetachedPivotReport, createPivotTable, getPivotSourceFields, normalizePivotSource, PIVOT_AGGREGATIONS, refreshPivotTable,
  setPivotFieldSetting, setPivotOptions, type PivotTableAnchorInput, type PivotTableSourceInput } from './pivot-engine.ts'
import type { PivotChartCommandContext } from './pivot-chart-api-types.ts'

export const CONFIGURE_PIVOT_COMMAND = 'sheet.command.configure-pivot-table'
export interface PivotPanelField {
  fieldId: string
  sourceFieldId?: string
  fieldName: string
  sourceColumn: number
  area: PivotFieldArea | 'unused'
  aggregation: PivotAggregation
  sort?: PivotSortDirection
  displayName?: string
  format?: string
  showDataAs?: PivotShowDataAsInfo | null
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
export interface ConfigurePivotFromPanelInput extends PivotChartCommandContext {
  unitId: string
  hostSheetId?: string
  pivotTableId: string
  mode: 'create' | 'update'
  source: PivotTableSourceInput
  anchor: PivotTableAnchorInput
  fields: Array<{ fieldId: string; sourceFieldId?: string; sourceColumn?: number; area: PivotFieldArea; aggregation?: PivotAggregation;
    sort?: PivotSortDirection; displayName?: string; format?: string; showDataAs?: PivotShowDataAsInfo | null }>
  options: Partial<PivotTableOptions>
}
const maxRows = 1_048_576, maxColumns = 16_384
const areas: readonly PivotFieldArea[] = ['filter', 'row', 'column', 'value']
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
export function readPivotPanelSourceFields(workbook: WorkbookModel, sheetId: string, address: string, pivot?: PivotTableModel): PivotPanelField[] {
  const { sheet, range } = sourceRange(workbook, sheetId, address)
  const report = createDetachedPivotReport(workbook, { unitId: workbook.id!, subUnitId: sheetId, sheetName: sheet.name,
    range: { startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol },
    ...(samePivotPanelSource(pivot, sheetId, range) && (pivot?.pivotCacheId ?? pivot?.source.pivotCacheId)
      ? { pivotCacheId: pivot?.pivotCacheId ?? pivot?.source.pivotCacheId } : {}) },
    { unitId: workbook.id!, subUnitId: sheetId, row: 0, col: 0 }, 'pivot-panel-source-fields')
  if (samePivotPanelSource(pivot, sheetId, range)) report.fields = cloneModel(pivot!.fields)
  return getPivotSourceFields(workbook, report).map(info => ({ fieldId: info.fieldId, sourceFieldId: info.fieldId,
    fieldName: info.fieldName, sourceColumn: info.sourceColumn, area: 'unused', aggregation: 'sum' }))
}

function samePivotPanelSource(pivot: PivotTableModel | undefined, sheetId: string, range: CellRange): boolean {
  return Boolean(pivot && pivot.source.subUnitId === sheetId && pivot.source.range.startRow === range.startRow
    && pivot.source.range.endRow === range.endRow && pivot.source.range.startCol === range.startCol && pivot.source.range.endCol === range.endCol)
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
    draft.fields = readPivotPanelSourceFields(workbook, draft.sourceSheetId, draft.sourceRange, pivot)
    if (!pivot) {
      const range = sourceRange(workbook, sheetId, address).range
      draft.targetAddress = range.endCol + 2 < maxColumns ? `${columnLabel(range.endCol + 2)}${range.startRow + 1}` : `A${range.endRow + 3}`
    } else {
      const sourceFields = draft.fields, usedIds = new Set(pivot.fields.map(field => field.id))
      // Preserve every configured instance; matching a local ID to a numeric column can otherwise drop grouped/repeated fields.
      draft.fields = sourceFields.flatMap(field => {
        const existing = pivot.fields.filter(candidate => candidate.sourceColumn === field.sourceColumn)
        if (existing.length) return existing.map(candidate => ({ ...field, fieldId: candidate.id, sourceFieldId: candidate.sourceFieldId ?? field.sourceFieldId,
          area: candidate.area, aggregation: candidate.aggregation ?? 'sum', sort: candidate.sort, displayName: candidate.displayName,
          format: candidate.format, ...(candidate.showDataAs ? { showDataAs: cloneModel(candidate.showDataAs) } : {}) }))
        let id = field.fieldId, suffix = 1
        while (usedIds.has(id)) id = `${field.fieldId}:source:${suffix++}`
        usedIds.add(id)
        return [{ ...field, fieldId: id }]
      })
      for (const area of areas) {
        const positions = draft.fields.map((field, index) => field.area === area ? index : -1).filter(index => index >= 0)
        const ordered = draft.fields.filter(field => field.area === area).sort((first, second) => {
          const firstPosition = pivot.fields.find(field => field.id === first.fieldId)?.position ?? 0
          const secondPosition = pivot.fields.find(field => field.id === second.fieldId)?.position ?? 0
          return firstPosition - secondPosition
        })
        positions.forEach((position, index) => { draft.fields[position] = ordered[index]! })
      }
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
  const previous = workbook.pivotTables?.find(pivot => pivot.id === draft.pivotId)
  const available = readPivotPanelSourceFields(workbook, draft.sourceSheetId, draft.sourceRange, previous)
  const fields = draft.fields.filter(field => field.area !== 'unused').map(field => {
    if (!available.some(candidate => candidate.sourceColumn === field.sourceColumn
      && (field.sourceFieldId === undefined || field.sourceFieldId === candidate.sourceFieldId))) throw new Error(`字段“${field.fieldName}”已不在当前数据区域中，请重新读取字段。`)
    if (!areas.some(area => area === field.area)) throw new Error('字段区域无效。')
    const area: PivotFieldArea = field.area === 'unused' ? 'row' : field.area
    return {
      fieldId: field.fieldId,
      sourceFieldId: field.sourceFieldId, sourceColumn: field.sourceColumn,
      area,
      ...(area === 'value' ? { aggregation: field.aggregation } : {}),
      ...((area === 'row' || area === 'column') && field.sort ? { sort: field.sort } : {}),
      ...(field.displayName !== undefined ? { displayName: field.displayName } : {}),
      ...(area === 'value' && field.format !== undefined ? { format: field.format } : {}),
      ...(area === 'value' && Object.hasOwn(field, 'showDataAs') ? { showDataAs: field.showDataAs ?? null } : {}),
    }
  })
  const input: ConfigurePivotFromPanelInput = {
    unitId: workbook.id, pivotTableId, mode: draft.pivotId ? 'update' : 'create',
    source: { unitId: workbook.id, subUnitId: sheet.id, sheetName: sheet.name, range: { startRow: range.startRow, startColumn: range.startCol, endRow: range.endRow, endColumn: range.endCol },
      ...(samePivotPanelSource(previous, sheet.id, range) && (previous?.pivotCacheId ?? previous?.source.pivotCacheId)
        ? { pivotCacheId: previous?.pivotCacheId ?? previous?.source.pivotCacheId } : {}) },
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
  const retained = samePivotPanelSource(existing, source.subUnitId, source.range) ? existing : undefined
  const identityReport = createDetachedPivotReport(workbook, input.source, input.anchor, input.pivotTableId)
  if (retained) identityReport.fields = cloneModel(retained.fields)
  const sourceFields = getPivotSourceFields(workbook, identityReport)
  for (const field of input.fields) {
    const column = pivotPanelSourceColumn(field, source.range, retained)
    if (typeof field.fieldId !== 'string' || !field.fieldId.trim() || !Number.isSafeInteger(column)
      || column < source.range.startCol || column > source.range.endCol || !areas.includes(field.area)) throw new Error('配置中存在无效字段。')
    if (field.sourceFieldId !== undefined && (typeof field.sourceFieldId !== 'string' || !field.sourceFieldId.trim())) throw new Error('字段的来源标识无效。')
    if (field.sourceFieldId !== undefined && sourceFields.find(candidate => candidate.sourceColumn === column)?.fieldId !== field.sourceFieldId) throw new Error('字段来源标识与来源列不匹配。')
    if (seen.has(field.fieldId)) throw new Error('同一个报告字段实例不能重复放置。')
    seen.add(field.fieldId)
    if (field.area === 'value' && !PIVOT_AGGREGATIONS.includes(field.aggregation ?? 'sum')) throw new Error('请选择有效的汇总方式。')
    if (field.sort !== undefined && (!['row', 'column'].includes(field.area) || !['ascending', 'descending'].includes(field.sort))) throw new Error('排序只能用于行字段或列字段。')
    if (field.displayName !== undefined && (typeof field.displayName !== 'string' || !field.displayName.trim())) throw new Error('字段名称不能为空。')
    if (field.format !== undefined && typeof field.format !== 'string') throw new Error('数字格式必须是文本。')
    if (Object.hasOwn(field, 'showDataAs') && field.area !== 'value') throw new Error('值显示方式只能用于值字段。')
  }
  if (!input.fields.some(field => field.area === 'value')) throw new Error('请至少将一个字段放入“值”区域。')
  for (const key of ['repeatRowLabels', 'showRowGrandTotal', 'showColGrandTotal', 'showRowSubtotals', 'showColSubtotals', 'allowMultipleFilters'] as const) {
    if (input.options[key] !== undefined && typeof input.options[key] !== 'boolean') throw new Error('汇总选项必须是开关值。')
  }
  if (input.options.valuePosition !== undefined && !['none', 'row', 'column'].includes(input.options.valuePosition)) throw new Error('值字段排列方向无效。')
  if (input.options.valueIndex !== undefined && (!Number.isSafeInteger(input.options.valueIndex) || input.options.valueIndex < -1)) throw new Error('值字段位置必须是 -1 或非负整数。')
}

function pivotPanelSourceColumn(field: ConfigurePivotFromPanelInput['fields'][number], range: CellRange, previous?: PivotTableModel): number {
  return field.sourceColumn ?? previous?.fields.find(candidate => candidate.id === field.fieldId)?.sourceColumn
    ?? (/^\d+$/.test(field.fieldId) ? range.startCol + Number(field.fieldId) : NaN)
}

function configuredPivotPanelFields(workbook: WorkbookModel, sourceInput: PivotTableSourceInput,
  fields: ConfigurePivotFromPanelInput['fields'], previous?: PivotTableModel): PivotFieldModel[] {
  const sourceInfo = normalizePivotSource(sourceInput, workbook), sheet = requireSheet(workbook, sourceInfo.subUnitId)
  const retained = samePivotPanelSource(previous, sourceInfo.subUnitId, sourceInfo.range) ? previous : undefined
  const identityReport = createDetachedPivotReport(workbook, sourceInput, { unitId: workbook.id!, subUnitId: sourceInfo.subUnitId, row: 0, col: 0 }, 'pivot-panel-field-identities')
  if (retained) identityReport.fields = cloneModel(retained.fields)
  const sourceFields = getPivotSourceFields(workbook, identityReport)
  const positions: Record<PivotFieldArea, number> = { row: 0, column: 0, filter: 0, value: 0 }
  return fields.map(field => {
    const old = retained?.fields.find(candidate => candidate.id === field.fieldId)
    const column = pivotPanelSourceColumn(field, sourceInfo.range, retained), header = sheet.cells[`${sourceInfo.range.startRow}:${column}`]?.value
    const next: PivotFieldModel = { ...(old ? cloneModel(old) : {}), id: field.fieldId,
      sourceFieldId: field.sourceFieldId ?? old?.sourceFieldId ?? sourceFields.find(candidate => candidate.sourceColumn === column)?.fieldId ?? String(column - sourceInfo.range.startCol),
      name: header == null || header === '' ? `Column ${column - sourceInfo.range.startCol + 1}` : String(header),
      sourceColumn: column, area: field.area, position: positions[field.area]++,
      ...(field.area === 'value' ? { aggregation: field.aggregation ?? 'sum' } : {}),
      ...((field.area === 'row' || field.area === 'column') && field.sort ? { sort: field.sort } : {}),
      ...(field.displayName !== undefined ? { displayName: field.displayName.trim() } : {}),
      ...(field.area === 'value' && field.format !== undefined ? { format: field.format } : {}) }
    if (field.area !== 'value') { delete next.aggregation; delete next.showDataAs }
    else if (Object.hasOwn(field, 'showDataAs')) {
      if (field.showDataAs == null) delete next.showDataAs
      else next.showDataAs = cloneModel(field.showDataAs)
    }
    if (field.area !== 'row' && field.area !== 'column' || !field.sort) delete next.sort
    return next
  })
}

/** A report-local, grouped field model for an unsaved picker; never reserves cells or modifies the workbook. */
export function pivotPanelDraftReport(workbook: WorkbookModel, draft: PivotPanelDraft): PivotTableModel {
  const { sheet, range } = sourceRange(workbook, draft.sourceSheetId, draft.sourceRange)
  const previous = workbook.pivotTables?.find(pivot => pivot.id === draft.pivotId)
  const source: PivotTableSourceInput = { unitId: workbook.id!, subUnitId: sheet.id, sheetName: sheet.name,
    range: { startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol },
    ...(samePivotPanelSource(previous, sheet.id, range) && (previous?.pivotCacheId ?? previous?.source.pivotCacheId)
      ? { pivotCacheId: previous?.pivotCacheId ?? previous?.source.pivotCacheId } : {}) }
  const target = parseSheetAddress(draft.targetAddress, { rowCount: maxRows, columnCount: maxColumns })
  if (!target || target.range.startRow !== target.range.endRow || target.range.startColumn !== target.range.endColumn) throw new Error('放置位置必须是单个单元格。')
  const report = createDetachedPivotReport(workbook, source, { unitId: workbook.id!, subUnitId: draft.targetSheetId,
    row: target.range.startRow, col: target.range.startColumn }, draft.pivotId || 'pivot-panel-draft')
  report.fields = configuredPivotPanelFields(workbook, source, draft.fields.filter(field => field.area !== 'unused').map(field => ({
    fieldId: field.fieldId, sourceFieldId: field.sourceFieldId, sourceColumn: field.sourceColumn, area: field.area === 'unused' ? 'row' : field.area,
    aggregation: field.aggregation, sort: field.sort, displayName: field.displayName, format: field.format,
    ...(Object.hasOwn(field, 'showDataAs') ? { showDataAs: field.showDataAs } : {}) })), previous)
  report.options = { ...(previous?.options ?? report.options) }
  setPivotOptions(workbook, report, draft.options, 'deferred')
  if (samePivotPanelSource(previous, sheet.id, range)) report.dateGroupingMode = previous?.dateGroupingMode
  return report
}

/** Mutates the command's detached workbook draft; OfficeStudio commits it as one history transaction. */
export function configurePivotFromPanel(workbook: WorkbookModel, input: ConfigurePivotFromPanelInput): PivotTableModel {
  validateConfiguration(workbook, input)
  let pivot = workbook.pivotTables?.find(candidate => candidate.id === input.pivotTableId)
  const previous = pivot
  const sourceUnchanged = previous && previous.source.subUnitId === input.source.subUnitId && previous.source.range.startRow === input.source.range.startRow && previous.source.range.endRow === input.source.range.endRow && previous.source.range.startCol === input.source.range.startColumn && previous.source.range.endCol === input.source.range.endColumn
  let fields = configuredPivotPanelFields(workbook, input.source, input.fields, sourceUnchanged ? previous : undefined)
  const validationReport = createDetachedPivotReport(workbook, input.source, input.anchor, input.pivotTableId)
  validationReport.fields = fields
  // The engine checks real grouped scalar members and collisions before any output/model is replaced.
  for (const field of fields.filter(field => field.area === 'value' && field.showDataAs)) {
    const rule = cloneModel(field.showDataAs!)
    delete field.showDataAs
    setPivotFieldSetting(workbook, validationReport, field.id, { showDataAs: rule }, 'deferred')
  }
  validationReport.options = { ...(pivot?.options ?? defaultOptions()) }
  setPivotOptions(workbook, validationReport, input.options, 'deferred')
  fields = validationReport.fields
  const options = validationReport.options
  const comparableFields = (items: PivotFieldModel[]) => items.map(field => ({ ...field,
    sourceFieldId: field.sourceFieldId ?? String(field.sourceColumn - validationReport.source.range.startCol),
    ...(field.area === 'value' ? { aggregation: field.aggregation ?? 'sum' } : {}) }))
  if (pivot && sourceUnchanged && sameConfigurationValue(pivot.source, validationReport.source)
    && sameConfigurationValue(pivot.anchor, validationReport.anchor) && sameConfigurationValue(comparableFields(pivot.fields), comparableFields(fields))
    && sameConfigurationValue({ ...defaultOptions(), ...pivot.options }, options)) return pivot
  if (!pivot) pivot = createPivotTable(workbook, input.source, input.anchor, input.pivotTableId)
  else {
    const oldTarget = requireSheet(workbook, pivot.anchor.subUnitId)
    for (const key of pivot.generatedCellKeys ?? []) if (oldTarget.cells[key]?.metadata?.pivotTableId === pivot.id) delete oldTarget.cells[key]
    pivot.generatedCellKeys = []
    pivot.source = normalizePivotSource(input.source, workbook)
    pivot.anchor = { ...input.anchor }
  }
  pivot.fields = fields
  pivot.options = options
  refreshPivotTable(workbook, pivot)
  if (pivot.outputRange && (pivot.outputRange.endRow >= maxRows || pivot.outputRange.endCol >= maxColumns)) throw new Error('数据透视表结果超出工作表有效范围，请调整放置位置。')
  return pivot
}

/** DTO property order and omitted undefined fields do not create content history. */
function sameConfigurationValue(first: unknown, second: unknown): boolean {
  if (Object.is(first, second)) return true
  if (Array.isArray(first) && Array.isArray(second)) return first.length === second.length && first.every((item, index) => sameConfigurationValue(item, second[index]))
  if (!first || !second || typeof first !== 'object' || typeof second !== 'object' || Array.isArray(first) || Array.isArray(second)) return false
  const left = first as Record<string, unknown>, right = second as Record<string, unknown>
  const keys = Object.keys(left).filter(key => left[key] !== undefined), otherKeys = Object.keys(right).filter(key => right[key] !== undefined)
  return keys.length === otherKeys.length && keys.every(key => Object.hasOwn(right, key) && sameConfigurationValue(left[key], right[key]))
}
