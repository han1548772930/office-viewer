import {
  createCell,
  defaultCellStyle,
  type CellRange,
  type CellScalar,
  type PivotAggregation,
  type PivotAxisOutputEntry,
  type PivotCollapseInfo,
  type PivotDateGroupType,
  type PivotDateGroupingMode,
  type PivotDateFilter,
  type PivotFieldArea,
  type PivotFieldFilter,
  type PivotFieldModel,
  type PivotLabelConditionFilter,
  type PivotManualGroup,
  type PivotNumberGroup,
  type PivotRecordSelection,
  type PivotShowDataAsInfo,
  type PivotTableModel,
  type PivotTableOptions,
  type PivotTableSource,
  type PivotValuePosition,
  type SpreadsheetSheet,
  type WorkbookModel,
} from '../../../domain/model'
import { randomId } from '../../../api/model-utils'
import { isRecordValue, pivotRecordScalar, pivotRecordValueKey, type PivotRecordValue, type RecordValue } from '../../../domain/record-value.ts'
import { sheetCellRecordValue } from '../../sheet-workbench/sheet-record-value.ts'
import { parseTimelineDate } from '../../slicer/timeline-date'
import type { ITimelineDateCondition, ITimelineRange } from '../../../domain/timeline-filters.ts'
import { compileTimelineSerialCondition, compileTimelineSerialRange, normalizeTimelineDateCondition, normalizeTimelineDateRange,
  timelineDateConditionToPivot } from '../../slicer/timeline-date-filter.ts'
import { classifySheetNumberFormat, formatSheetValue } from '../../sheet-workbench/sheet-number-format'
import { applyPivotShowDataAs, normalizePivotShowDataAs, pivotShowDataAsNumberFormat,
  resolvePivotShowDataAsBaseField, type PivotShowDataAsScope } from './pivot-show-data-as.ts'
import { assertPivotCollapseGroupingChange, equalPivotCollapseInfo, getPivotAxisState as pivotAxisState,
  normalizePivotCollapseInfo, pivotAxisField, pivotCategoryStepField, preparePivotCategoryLevel, preparePivotCollapse,
  projectPivotAxis, retainPivotCollapseInfo, type PivotAxisArea, type PivotAxisState } from './pivot-axis-state.ts'
export { getPivotCollapse } from './pivot-axis-state.ts'
import { clearPivotFieldFilterState, clearPivotReportFilterState, compilePivotDateFilter, compilePivotLabelConditionFilter,
  compilePivotRecordSelection, getPivotFieldFilterCategories, hasPivotMultipleFilters, normalizePivotDateFilter,
  normalizePivotLabelConditionFilter, normalizePivotRecordSelection, normalizePivotValueFilter, pivotFilterSourceKey, replacePivotFilterCategory } from './pivot-filter-predicate.ts'
export { getPivotFieldFilterCategories, hasPivotMultipleFilters } from './pivot-filter-predicate.ts'
import { adjustPivotValuesIndexForFieldEdit, expandPivotValuesAxis, getPivotValuesLayout, normalizePivotValuesPosition,
  retainPivotValuesIndexAfterRemoval } from './pivot-values-position.ts'
export { getPivotValuesPosition } from './pivot-values-position.ts'

export interface PivotSourceFieldInfo {
  fieldId: string
  fieldName: string
  sourceColumn: number
}

export interface PivotDrillDownEntry {
  fieldFilters: Record<string, CellScalar>
  fieldMembers?: Record<string, CellScalar[]>
  valueFieldId: string
}

export interface PivotTableModelWithDrillDown extends PivotTableModel {
  drillDownMap?: Record<string, PivotDrillDownEntry>
}

export interface PivotTableSourceInput {
  unitId: string
  subUnitId: string
  sheetName: string
  pivotCacheId?: string
  range: { startRow: number; startColumn: number; endRow: number; endColumn: number }
}

export interface PivotTableAnchorInput {
  unitId: string
  subUnitId: string
  row: number
  col: number
}

export interface PivotRenderResult {
  outputRange?: CellRange
  generatedCellKeys: string[]
  drillDownMap: Record<string, PivotDrillDownEntry>
}

export interface PivotCalculationResult {
  values: CellScalar[][]
  headerRows: number
  drillDownMap: Record<string, PivotDrillDownEntry>
  numberFormats: Record<string, string>
  rowAxis: PivotAxisOutputEntry[]
  columnAxis: PivotAxisOutputEntry[]
}

export type PivotRenderMode = 'worksheet' | 'data-only'
/** Deferred is for detached field batches only; their caller performs one final refresh. */
export type PivotRefreshMode = PivotRenderMode | 'deferred'

export class PivotOutputConflictError extends Error {
  constructor(message: string) { super(message); this.name = 'PivotOutputConflictError' }
}

export const PIVOT_AGGREGATIONS: readonly PivotAggregation[] = ['average', 'count', 'countNums', 'max', 'min', 'product', 'stdDev', 'stdDevp', 'sum', 'var', 'varp']

interface SourceRecord {
  row: number
  values: CellScalar[]
}

interface Group {
  values: CellScalar[]
  key: string
  memberValues: Record<string, CellScalar[]>
}

export type PivotDateGroup = PivotDateGroupType
export type PivotDateMode = PivotDateGroupingMode
export type PivotNumberGroupSettings = PivotNumberGroup
export type PivotManualGroupSettings = PivotManualGroup

interface OutputRow {
  group: Group | null
  valueFieldId?: string
  total: boolean
  subtotalLevel?: number
  collapsedLevel?: number
}

interface OutputColumn {
  group: Group | null
  valueFieldId?: string
  total: boolean
  subtotalLevel?: number
  collapsedLevel?: number
}

const defaultOptions = (): PivotTableOptions => ({
  repeatRowLabels: false,
  showRowGrandTotal: true,
  showColGrandTotal: true,
  showRowSubtotals: true,
  showColSubtotals: true,
  valuePosition: 'column',
})

export function normalizePivotSource(input: PivotTableSourceInput, workbook: WorkbookModel): PivotTableSource {
  const sheet = workbook.sheets.find((candidate) => candidate.id === input.subUnitId)
  if (!sheet || workbook.id !== input.unitId) throw new Error('Pivot source worksheet does not belong to the workbook')
  const range = {
    startRow: input.range.startRow,
    startCol: input.range.startColumn,
    endRow: input.range.endRow,
    endCol: input.range.endColumn,
  }
  validateRange(range)
  if (range.endRow <= range.startRow) throw new RangeError('Pivot source must contain a header and at least one data row')
  if (range.endCol < range.startCol) throw new RangeError('Pivot source must contain at least one column')
  if (range.endRow >= (sheet.rowCount ?? 1_048_576) || range.endCol >= (sheet.columnCount ?? 16_384)) throw new RangeError('Pivot source exceeds worksheet dimensions')
  if (input.pivotCacheId !== undefined && (typeof input.pivotCacheId !== 'string' || !input.pivotCacheId.trim())) throw new TypeError('Invalid pivot cache ID')
  return { unitId: input.unitId, subUnitId: input.subUnitId, sheetName: sheet.name, range,
    ...(input.pivotCacheId ? { pivotCacheId: input.pivotCacheId } : {}) }
}

/** A real report model without registration, output reservation, cells or derived output metadata. */
export function createDetachedPivotReport(workbook: WorkbookModel, sourceInput: PivotTableSourceInput,
  anchorInput: PivotTableAnchorInput, id = randomId('pivot')): PivotTableModel {
  const source = normalizePivotSource(sourceInput, workbook)
  const host = workbook.sheets.find(sheet => sheet.id === anchorInput.subUnitId)
  if (!host || anchorInput.unitId !== workbook.id) throw new Error('Pivot anchor worksheet does not belong to the workbook')
  validateIndex(anchorInput.row, 'anchor row'); validateIndex(anchorInput.col, 'anchor column')
  if (typeof id !== 'string' || !id.trim()) throw new TypeError('Invalid pivot report ID')
  return { id, source, ...(source.pivotCacheId ? { pivotCacheId: source.pivotCacheId } : {}),
    anchor: { ...anchorInput }, fields: [], options: defaultOptions(), revision: 0 }
}

export function createPivotTable(
  workbook: WorkbookModel,
  sourceInput: PivotTableSourceInput,
  anchorInput: PivotTableAnchorInput,
  id = randomId('pivot')
): PivotTableModelWithDrillDown {
  const detached = createDetachedPivotReport(workbook, sourceInput, anchorInput, id)
  const anchorSheet = workbook.sheets.find((sheet) => sheet.id === anchorInput.subUnitId)
  if (!anchorSheet) throw new Error('Pivot anchor worksheet does not belong to the workbook')
  const model: PivotTableModelWithDrillDown = {
    ...detached,
    generatedCellKeys: [],
    drillDownMap: {},
  }
  // Reserve the requested anchor as part of creation. Without this preflight,
  // a pivot with no value field has no output range yet and a later field
  // mutation could fail after leaving a partially configured pivot in history.
  ensureNoOutputConflict(
    workbook,
    model,
    { startRow: anchorInput.row, startCol: anchorInput.col, endRow: anchorInput.row, endCol: anchorInput.col },
    anchorSheet,
  )
  workbook.pivotTables ??= []
  workbook.pivotTables.push(model)
  renderPivotTable(workbook, model)
  return model
}

export function getPivotSourceFields(workbook: WorkbookModel, pivot: PivotTableModel): PivotSourceFieldInfo[] {
  const source = sourceSheet(workbook, pivot)
  const retained = workbook.pivotCaches?.find(cache => cache.id === (pivot.pivotCacheId ?? pivot.source.pivotCacheId))?.fields
  const result: PivotSourceFieldInfo[] = []
  for (let column = pivot.source.range.startCol; column <= pivot.source.range.endCol; column++) {
    const header = cellValue(source, pivot.source.range.startRow, column)
    const declared = pivot.fields.find(field => field.sourceColumn === column)?.sourceFieldId
    const fieldId = declared ?? retained?.find(field => field.sourceColumn === column)?.id ?? String(column - pivot.source.range.startCol)
    result.push({ fieldId, fieldName: header == null || header === '' ? `Column ${column - pivot.source.range.startCol + 1}` : String(header), sourceColumn: column })
  }
  return result
}

/** Source-order distinct raw values; grouping and report filters do not change this list. */
export function getPivotSourceFieldItems(workbook: WorkbookModel, pivot: PivotTableModel, sourceFieldId: string): CellScalar[] {
  const info = resolveField(workbook, pivot, sourceFieldId), sheet = sourceSheet(workbook, pivot)
  const items = new Set<CellScalar>()
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) items.add(cellValue(sheet, row, info.sourceColumn))
  return [...items]
}

export function setPivotField(
  workbook: WorkbookModel,
  pivot: PivotTableModel,
  fieldIdOrIndex: string | number,
  area: PivotFieldArea,
  position?: number,
  mode: PivotRefreshMode = 'worksheet',
  instanceId?: string,
): void {
  if (!['row', 'column', 'filter', 'value'].includes(area)) throw new TypeError('Invalid pivot field area')
  const info = resolveField(workbook, pivot, fieldIdOrIndex)
  const existing = instanceId === undefined
    ? pivot.fields.find((field) => field.sourceFieldId === info.fieldId || field.sourceColumn === info.sourceColumn)
    : requirePivotField(pivot, instanceId)
  if (existing && existing.sourceColumn !== info.sourceColumn) throw new Error('Pivot field instance belongs to another source column')
  if (existing?.area === 'value' && area !== 'value' && pivot.fields.some(field => field.valueFilter?.valueFieldId === existing.id)) {
    throw new Error('A value field used by a surviving value filter cannot leave the value area')
  }
  if (existing && area !== 'row' && area !== 'column' && pivot.fields.some(field => field !== existing && field.showDataAs?.baseFieldId === existing.id)) {
    throw new Error('A base field used by a surviving Show Values As rule must remain on a row or column axis')
  }
  const siblings = areaFields(pivot, area, existing?.id)
  const nextPosition = normalizePosition(position, siblings.length)
  const previousOrder = pivot.fields.map(field => ({ id: field.id, area: field.area, position: field.position }))
  let editedId: string
  if (existing) {
    if (existing.area === area && existing.position === nextPosition) return
    editedId = existing.id
    existing.area = area
    if (area === 'value') {
      existing.aggregation ??= 'sum'
      delete existing.selectedItems; delete existing.groupedItems; delete existing.partialItems; delete existing.recordSelection
      delete existing.labelFilter; delete existing.dateFilter; delete existing.timelineDateRange; delete existing.timelineDateCondition; delete existing.valueFilter; delete existing.sort
      delete existing.numberGroup; delete existing.dateGroupType; delete existing.manualGroups
    } else { delete existing.aggregation; delete existing.showDataAs }
    siblings.splice(nextPosition, 0, existing)
  } else {
    let localId = info.fieldId
    while (pivot.fields.some(field => field.id === localId)) localId = randomId('pivot-field')
    const field: PivotFieldModel = {
      id: localId,
      sourceFieldId: info.fieldId,
      sourceColumn: info.sourceColumn,
      name: info.fieldName,
      area,
      position: nextPosition,
      ...(area === 'value' ? { aggregation: 'sum' as const } : {}),
    }
    pivot.fields.push(field)
    editedId = field.id
    siblings.splice(nextPosition, 0, field)
  }
  siblings.forEach((field, index) => { field.position = index })
  reindexFields(pivot)
  adjustPivotValuesIndexForFieldEdit(pivot, previousOrder, editedId)
  publishPivotCollapseInfo(pivot, retainPivotCollapseInfo(pivot, pivot.collapseInfo))
  refreshPivotTable(workbook, pivot, mode)
}

export function removePivotFields(workbook: WorkbookModel, pivot: PivotTableModel, fieldIds: string | string[], mode: PivotRefreshMode = 'worksheet'): void {
  const ids = new Set(Array.isArray(fieldIds) ? fieldIds.map(String) : [String(fieldIds)])
  if (!pivot.fields.some(field => ids.has(field.id))) return
  if (pivot.fields.some(field => !ids.has(field.id) && field.valueFilter && ids.has(field.valueFilter.valueFieldId))) {
    throw new Error('A value field used by a surviving value filter cannot be removed')
  }
  if (pivot.fields.some(field => !ids.has(field.id) && field.showDataAs && ids.has(field.showDataAs.baseFieldId))) {
    throw new Error('A base field used by a surviving Show Values As rule cannot be removed')
  }
  const previousOrder = pivot.fields.map(field => ({ id: field.id, area: field.area, position: field.position }))
  pivot.fields = pivot.fields.filter((field) => !ids.has(field.id))
  reindexFields(pivot)
  retainPivotValuesIndexAfterRemoval(pivot, previousOrder)
  publishPivotCollapseInfo(pivot, retainPivotCollapseInfo(pivot, pivot.collapseInfo))
  refreshPivotTable(workbook, pivot, mode)
}

export function setPivotOptions(workbook: WorkbookModel, pivot: PivotTableModel, options: Partial<PivotTableOptions>, mode: PivotRefreshMode = 'worksheet'): void {
  if (options.valuePosition !== undefined || options.valueIndex !== undefined) {
    normalizePivotValuesPosition(options.valuePosition === undefined ? pivot.options.valuePosition : options.valuePosition,
      options.valueIndex === undefined ? pivot.options.valueIndex === undefined ? -1 : pivot.options.valueIndex : options.valueIndex)
  }
  for (const key of ['repeatRowLabels', 'showRowGrandTotal', 'showColGrandTotal', 'showRowSubtotals', 'showColSubtotals', 'allowMultipleFilters'] as const) {
    if (options[key] !== undefined && typeof options[key] !== 'boolean') throw new TypeError(`Invalid pivot option: ${key}`)
  }
  const next = { ...pivot.options }
  for (const key of ['repeatRowLabels', 'showRowGrandTotal', 'showColGrandTotal', 'showRowSubtotals', 'showColSubtotals', 'valuePosition', 'valueIndex', 'allowMultipleFilters'] as const) {
    if (key === 'allowMultipleFilters' && options[key] === false && next[key] === undefined) continue
    if (key === 'valueIndex' && options[key] === -1) { delete next.valueIndex; continue }
    if (options[key] !== undefined) Object.assign(next, { [key]: options[key] })
  }
  const cleared = options.allowMultipleFilters === false && hasPivotMultipleFilters(pivot) && clearPivotReportFilterState(pivot)
  if (JSON.stringify(next) === JSON.stringify(pivot.options) && !cleared) return
  pivot.options = next
  refreshPivotTable(workbook, pivot, mode)
}

export function setPivotValuesPosition(workbook: WorkbookModel, pivot: PivotTableModel, position: PivotValuePosition,
  index = -1, mode: PivotRefreshMode = 'worksheet'): void {
  const state = normalizePivotValuesPosition(position, index)
  setPivotOptions(workbook, pivot, { valuePosition: state.position, valueIndex: state.index }, mode)
}

export function setPivotFieldSetting(
  workbook: WorkbookModel,
  pivot: PivotTableModel,
  fieldId: string,
  setting: { displayName?: string; format?: string; subtotalType?: PivotAggregation; showDataAs?: PivotShowDataAsInfo | null },
  mode: PivotRefreshMode = 'worksheet',
): void {
  const field = requirePivotField(pivot, fieldId)
  if (setting.subtotalType !== undefined && !PIVOT_AGGREGATIONS.includes(setting.subtotalType)) throw new TypeError('Invalid pivot aggregation')
  if (setting.subtotalType !== undefined && field.area !== 'value') throw new Error('Aggregation settings require a value field')
  const hasShowDataAs = Object.hasOwn(setting, 'showDataAs')
  if (hasShowDataAs && field.area !== 'value') throw new Error('Show Values As requires a value field')
  const showDataAs = hasShowDataAs
    ? normalizePivotShowDataAs(pivot, setting.showDataAs ?? null, base => showDataAsBaseItems(workbook, pivot, base)) : undefined
  const before = JSON.stringify(field)
  if (setting.displayName !== undefined) field.displayName = setting.displayName
  if (setting.format !== undefined) field.format = setting.format
  if (setting.subtotalType !== undefined) field.aggregation = setting.subtotalType
  if (hasShowDataAs) {
    if (!showDataAs) delete field.showDataAs
    else if (!field.showDataAs || field.showDataAs.type !== showDataAs.type || field.showDataAs.baseFieldId !== showDataAs.baseFieldId
      || field.showDataAs.baseItem !== showDataAs.baseItem || field.showDataAs.baseItemType !== showDataAs.baseItemType) field.showDataAs = showDataAs
  }
  if (JSON.stringify(field) !== before) refreshPivotTable(workbook, pivot, mode)
}

export function setPivotShowDataAs(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string,
  showDataAs: PivotShowDataAsInfo | null, mode: PivotRefreshMode = 'worksheet'): void {
  if (showDataAs === undefined) throw new TypeError('Show Values As requires a rule or an explicit reset')
  setPivotFieldSetting(workbook, pivot, fieldId, { showDataAs }, mode)
}

export function getPivotShowDataAs(pivot: PivotTableModel, fieldId: string): PivotShowDataAsInfo | undefined {
  const field = requirePivotField(pivot, fieldId)
  return field.area === 'value' && field.showDataAs ? { ...field.showDataAs } : undefined
}

/** Report checklist keys share the exact grouping used by aggregation; filters do not hide picker members. */
export function getPivotGroupedFieldItems(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string): string[] {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Value fields cannot provide label filter items')
  return [...new Set(showDataAsBaseItems(workbook, pivot, field).map(value => String(value ?? '')))]
}

/** Grouped, real source members for the Show Values As base-item picker. */
export function getPivotShowDataAsBaseItems(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string): string[] {
  const field = resolvePivotShowDataAsBaseField(pivot, fieldId)
  return [...new Set(showDataAsBaseItems(workbook, pivot, field).map(value => String(value ?? '')))]
}

export function getPivotCollapseItems(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string): string[] {
  const field = pivotAxisField(pivot, fieldId, true)
  return [...new Set(showDataAsBaseItems(workbook, pivot, field).map(value => String(value ?? '')))]
}

export function getPivotAxisState(workbook: WorkbookModel, pivot: PivotTableModel, area: PivotAxisArea = 'row'): PivotAxisState {
  getPivotSourceFields(workbook, pivot)
  return pivotAxisState(pivot, area)
}

export function setPivotCollapse(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, collapse: boolean,
  item?: string, mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  const field = pivotAxisField(pivot, fieldId, true)
  const next = preparePivotCollapse(pivot, field.id, collapse, item, item === undefined ? [] : showDataAsBaseItems(workbook, pivot, field))
  if (publishPivotCollapseInfo(pivot, next, true)) refreshPivotTable(workbook, pivot, mode)
}

/** Full configuration can atomically reset or remap the hierarchy alongside new fields/grouping. */
export function setPivotCollapseInfo(workbook: WorkbookModel, pivot: PivotTableModel, info: PivotCollapseInfo | null,
  mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  const next = normalizePivotCollapseInfo(pivot, info === null ? {} : info, field => showDataAsBaseItems(workbook, pivot, field))
  if (publishPivotCollapseInfo(pivot, next, true)) refreshPivotTable(workbook, pivot, mode)
}

export function setPivotCategoryLevel(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string,
  mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  const next = preparePivotCategoryLevel(pivot, fieldId)
  if (publishPivotCollapseInfo(pivot, next, true)) refreshPivotTable(workbook, pivot, mode)
}

export function stepPivotCategoryLevel(workbook: WorkbookModel, pivot: PivotTableModel, direction: 'expand' | 'collapse',
  mode: PivotRefreshMode = 'worksheet'): boolean {
  assertPivotHierarchyRefreshMode(mode)
  normalizePivotCollapseInfo(pivot, pivot.collapseInfo ?? {})
  const fieldId = pivotCategoryStepField(pivot, direction)
  if (!fieldId) return false
  setPivotCollapse(workbook, pivot, fieldId, direction === 'collapse', undefined, mode)
  return true
}

function publishPivotCollapseInfo(pivot: PivotTableModel, next: PivotCollapseInfo, normalizeCurrent = false): boolean {
  const current = normalizeCurrent ? normalizePivotCollapseInfo(pivot, pivot.collapseInfo ?? {}) : pivot.collapseInfo
  if (equalPivotCollapseInfo(current, next)) return false
  if (Object.keys(next).length) pivot.collapseInfo = next
  else delete pivot.collapseInfo
  return true
}

function assertPivotHierarchyRefreshMode(mode: PivotRefreshMode): void {
  if (!['worksheet', 'data-only', 'deferred'].includes(mode)) throw new TypeError('Invalid pivot refresh mode')
}

export function setPivotSort(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, direction: 'ascending' | 'descending' | null, mode: PivotRefreshMode = 'worksheet'): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area !== 'row' && field.area !== 'column') throw new Error('Only row and column fields can be label-sorted')
  if (direction !== null && !['ascending', 'descending'].includes(direction)) throw new TypeError('Invalid pivot sort direction')
  const next = direction ?? undefined
  if (field.sort === next) return
  field.sort = next
  refreshPivotTable(workbook, pivot, mode)
}

export interface PivotManualFilterOptions {
  /** Existing callers use source strings; fixed checklist commands opt into grouped keys. */
  keyMode?: 'source' | 'grouped'
  partialItems?: string[]
}

function pivotMemberPredicate(field: PivotFieldModel, date1904: boolean): (value: PivotRecordValue, typedValue?: RecordValue) => boolean {
  const selected = field.selectedItems === undefined ? undefined : new Set(field.selectedItems)
  const records = field.recordSelection ? compilePivotRecordSelection(field.recordSelection) : undefined
  if (field.groupedItems === undefined) return records ?? (value => selected === undefined || selected.has(String(pivotRecordScalar(value) ?? '')))
  const grouped = new Set(field.groupedItems), partial = new Set(field.partialItems ?? [])
  return (value, typedValue) => {
    const key = String(groupedLabel(field, pivotRecordScalar(value), date1904) ?? '')
    return grouped.has(key) || partial.has(key) && (records?.(value, typedValue) ?? false)
  }
}

export function getPivotManualFilterItemState(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string): { checkedItems: string[]; partialItems: string[] } {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Value fields cannot provide manual filter items')
  const matches = pivotMemberPredicate(field, workbook.date1904), source = sourceSheet(workbook, pivot)
  const groups = new Map<string, { selected: boolean; unselected: boolean }>()
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    const value = cellValue(source, row, field.sourceColumn), key = String(groupedLabel(field, value, workbook.date1904) ?? '')
    const group = groups.get(key) ?? { selected: false, unselected: false }
    if (matches(value, sheetCellRecordValue(source.cells[`${row}:${field.sourceColumn}`], workbook.date1904))) group.selected = true; else group.unselected = true
    groups.set(key, group)
  }
  return { checkedItems: [...groups].filter(([, state]) => state.selected && !state.unselected).map(([key]) => key),
    partialItems: [...groups].filter(([, state]) => state.selected && state.unselected).map(([key]) => key) }
}

/** Raw scalar identity, independent of report-local field IDs and displayed grouping keys. */
export function getPivotSourceRecordSelection(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string): PivotRecordSelection | undefined {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Value fields cannot provide source member selections')
  const sourceId = pivotFilterSourceKey(pivot, field)
  const fields = pivot.fields.filter(candidate => candidate.area !== 'value' && pivotFilterSourceKey(pivot, candidate) === sourceId
    && getPivotFieldFilterCategories(candidate).has('member'))
  if (!fields.length) return undefined
  const direct = fields.length === 1 ? fields[0] : undefined
  if (direct && direct.groupedItems === undefined && direct.recordSelection) return normalizePivotRecordSelection(direct.recordSelection)
  const predicates = fields.map(candidate => pivotMemberPredicate(candidate, workbook.date1904))
  const matches = (value: PivotRecordValue, typedValue?: RecordValue) => predicates.every(predicate => predicate(value, typedValue))
  const source = sourceSheet(workbook, pivot), mode = direct?.recordSelection?.mode === 'exclude' ? 'exclude' : 'include'
  const typedAuthority = fields.some(candidate => candidate.recordSelection?.mode !== 'all' && candidate.recordSelection?.values.some(isRecordValue))
  const values: PivotRecordValue[] = []
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    const value = cellValue(source, row, field.sourceColumn)
    const typedValue = sheetCellRecordValue(source.cells[`${row}:${field.sourceColumn}`], workbook.date1904)
    if (matches(value, typedValue) === (mode === 'include')) values.push(typedAuthority ? typedValue : value)
  }
  // A source prune never discards explicitly retained raw include/exclude identities.
  if (direct?.recordSelection && direct.recordSelection.mode !== 'all') {
    for (const value of direct.recordSelection.values) if (matches(value) === (mode === 'include')) values.push(value)
  }
  return normalizePivotRecordSelection({ mode, values })
}

/** Effective source-string mirror; shared Slicer fan-out uses getPivotSourceRecordSelection. */
export function getPivotSourceSelectedItems(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string): string[] | undefined {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Value fields cannot provide source member selections')
  const selection = getPivotSourceRecordSelection(workbook, pivot, fieldId)
  if (!selection) return undefined
  const matches = compilePivotRecordSelection(selection), source = sourceSheet(workbook, pivot), values = new Set<string>()
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    const value = cellValue(source, row, field.sourceColumn)
    if (matches(value, sheetCellRecordValue(source.cells[`${row}:${field.sourceColumn}`], workbook.date1904))) values.add(String(value ?? ''))
  }
  return [...values]
}

export function setPivotLabelFilter(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, selectedItems: string[] | null,
  mode: PivotRefreshMode = 'worksheet', options: PivotManualFilterOptions = {}): void {
  assertPivotHierarchyRefreshMode(mode)
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Value fields cannot use label filters')
  if (!options || typeof options !== 'object' || Array.isArray(options)
    || Object.keys(options).some(key => !['keyMode', 'partialItems'].includes(key))
    || options.keyMode !== undefined && !['source', 'grouped'].includes(options.keyMode)
    || selectedItems !== null && (!Array.isArray(selectedItems) || selectedItems.some(value => typeof value !== 'string'))
    || options.partialItems !== undefined && (!Array.isArray(options.partialItems) || options.partialItems.some(value => typeof value !== 'string'))) {
    throw new TypeError('Invalid pivot manual filter items')
  }
  if (selectedItems === null) {
    if (options.partialItems?.length) throw new TypeError('All manual selection cannot contain partial items')
    if (replacePivotFilterCategory(pivot, field.id, 'member', {}, false)) refreshPivotTable(workbook, pivot, mode)
    return
  }
  const checked = new Set(selectedItems), partial = new Set(options.partialItems ?? [])
  if ([...partial].some(key => checked.has(key))) throw new TypeError('A Pivot item cannot be both checked and partial')
  if (options.keyMode !== 'grouped') {
    if (partial.size) throw new TypeError('Partial manual items require grouped keys')
    const items = [...checked].sort()
    if (replacePivotFilterCategory(pivot, field.id, 'member', { selectedItems: items }, true)) refreshPivotTable(workbook, pivot, mode)
    return
  }
  const keys = getPivotGroupedFieldItems(workbook, pivot, field.id)
  if ([...checked, ...partial].some(key => !keys.includes(key))) throw new TypeError('Invalid Pivot filter item')
  const source = sourceSheet(workbook, pivot), current = pivotMemberPredicate(field, workbook.date1904)
  const typedAuthority = field.recordSelection?.mode !== 'all' && field.recordSelection?.values.some(isRecordValue)
  const selected: PivotRecordValue[] = [], unselected: PivotRecordValue[] = [], seen = new Set<string>(), selectedLabels = new Set<string>()
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    const value = cellValue(source, row, field.sourceColumn), key = String(groupedLabel(field, value, workbook.date1904) ?? '')
    const typedValue = sheetCellRecordValue(source.cells[`${row}:${field.sourceColumn}`], workbook.date1904)
    const member = typedAuthority ? typedValue : value, identity = pivotRecordValueKey(member)
    const checkedValue = checked.has(key) || partial.has(key) && current(value, typedValue)
    if (checkedValue) selectedLabels.add(String(value ?? ''))
    if (!seen.has(identity)) {
      seen.add(identity)
      ;(checkedValue ? selected : unselected).push(member)
    }
  }
  const selectionMode = partial.size && field.recordSelection?.mode === 'exclude' ? 'exclude' : 'include'
  const values = selectionMode === 'exclude' ? unselected : selected
  if (partial.size && field.recordSelection && field.recordSelection.mode !== 'all' && field.recordSelection.mode === selectionMode) {
    for (const value of field.recordSelection.values) if (!seen.has(pivotRecordValueKey(value))) values.push(value)
  }
  const recordSelection = partial.size ? normalizePivotRecordSelection({ mode: selectionMode, values }) : undefined
  const state = { selectedItems: [...selectedLabels], groupedItems: keys.filter(key => checked.has(key)),
    ...(partial.size ? { partialItems: keys.filter(key => partial.has(key)), recordSelection } : {}) }
  if (replacePivotFilterCategory(pivot, field.id, 'member', state, true)) refreshPivotTable(workbook, pivot, mode)
}

export function setPivotSourceMemberFilter(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, items: string[] | null,
  mode: PivotRefreshMode = 'worksheet'): void {
  setPivotLabelFilter(workbook, pivot, fieldId, items, mode, { keyMode: 'source' })
}

export function setPivotRecordSelection(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, selection: PivotRecordSelection | null,
  mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  const field = requirePivotField(pivot, fieldId), next = selection === null ? undefined : normalizePivotRecordSelection(selection)
  if (field.area === 'value') throw new Error('Raw record selections require a label field')
  const predicate = next ? compilePivotRecordSelection(next) : undefined, source = sourceSheet(workbook, pivot), mirror = new Set<string>()
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    const value = cellValue(source, row, field.sourceColumn)
    if (predicate?.(value, sheetCellRecordValue(source.cells[`${row}:${field.sourceColumn}`], workbook.date1904))) mirror.add(String(value ?? ''))
  }
  const restrictive = !!next && (next.mode === 'include' || next.mode === 'exclude' && next.values.length > 0)
  if (replacePivotFilterCategory(pivot, field.id, 'member', restrictive ? { recordSelection: next, selectedItems: [...mirror] } : {}, restrictive)) {
    refreshPivotTable(workbook, pivot, mode)
  }
}

export function getPivotDateFilter(pivot: PivotTableModel, fieldId: string): PivotDateFilter | undefined {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Date filters require a label field')
  return field.dateFilter ? normalizePivotDateFilter(field.dateFilter) : undefined
}

export function getPivotTimelineDateRange(pivot: PivotTableModel, fieldId: string): ITimelineRange | undefined {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Timeline date ranges require a label field')
  return field.timelineDateRange === undefined ? undefined : normalizeTimelineDateRange(field.timelineDateRange)
}

export function getPivotLabelConditionFilter(pivot: PivotTableModel, fieldId: string): PivotLabelConditionFilter | undefined {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Caption conditions require a label field')
  return field.labelFilter ? normalizePivotLabelConditionFilter(field.labelFilter) : undefined
}

function sourceDateSerial(sheet: SpreadsheetSheet, row: number, field: PivotFieldModel, date1904: boolean): number | undefined {
  const typed = sourceTimelineDateSerial(sheet, row, field, date1904)
  if (typed !== undefined) return typed
  const value = cellValue(sheet, row, field.sourceColumn)
  const format = classifySheetNumberFormat(sheet.cells[`${row}:${field.sourceColumn}`]?.style?.numberFormat ?? 'General')
  const hinted = field.dateGroupType !== undefined || ['date', 'datetime', 'time'].includes(format)
  if (typeof value === 'number' ? !hinted : typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}(?:[Tt ]|$)/.test(value.trim())) return undefined
  return parseTimelineDate(value, date1904)?.serial
}

/** Canonical Timeline filtering uses stored typed dates, including cached formula dates. */
function sourceTimelineDateSerial(sheet: SpreadsheetSheet, row: number, field: PivotFieldModel, date1904: boolean): number | undefined {
  const value = sheetCellRecordValue(sheet.cells[`${row}:${field.sourceColumn}`], date1904)
  return value.type === 'date' ? value.value : undefined
}

export function getPivotFieldFilterCapabilities(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string): { supportsDateFilter: boolean; supportsTimelineRange: boolean; date1904: boolean } {
  const field = requirePivotField(pivot, fieldId), source = sourceSheet(workbook, pivot), date1904 = workbook.date1904 === true
  let supportsDateFilter = field.area !== 'value' && field.dateGroupType !== undefined
  let supportsTimelineRange = false
  if (field.area !== 'value') for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    if (sourceTimelineDateSerial(source, row, field, date1904) !== undefined) supportsTimelineRange = true
    if (!supportsDateFilter && sourceDateSerial(source, row, field, date1904) !== undefined) supportsDateFilter = true
    if (supportsDateFilter && supportsTimelineRange) break
  }
  return { supportsDateFilter, supportsTimelineRange, date1904 }
}

export function setPivotDateFilter(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, filter: PivotDateFilter | null,
  mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Date filters require a label field')
  if (filter === null) {
    if (clearPivotFieldFilterState(pivot, field.id, 'date')) refreshPivotTable(workbook, pivot, mode)
    return
  }
  const next = normalizePivotDateFilter(filter)
  if (!getPivotFieldFilterCapabilities(workbook, pivot, fieldId).supportsDateFilter) throw new Error('Date filters require a date source field')
  if (replacePivotFilterCategory(pivot, field.id, 'condition', { dateFilter: next }, true)) refreshPivotTable(workbook, pivot, mode)
}

/** Range and date conditions share one source category; the report's multiple-filter policy remains authoritative. */
export function setPivotTimelineDateRange(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, range: ITimelineRange | null,
  mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Timeline date ranges require a label field')
  if (range === null) {
    if (clearPivotFieldFilterState(pivot, field.id, 'date')) refreshPivotTable(workbook, pivot, mode)
    return
  }
  const next = normalizeTimelineDateRange(range)
  if (!getPivotFieldFilterCapabilities(workbook, pivot, fieldId).supportsTimelineRange) throw new Error('Timeline date ranges require a typed date source field')
  if (replacePivotFilterCategory(pivot, field.id, 'condition', { timelineDateRange: next }, true)) refreshPivotTable(workbook, pivot, mode)
}

/** Canonical Timeline conditions retain their typed-date authority beside the public Pivot condition. */
export function setPivotTimelineDateCondition(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, condition: ITimelineDateCondition | null,
  mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Timeline date conditions require a label field')
  if (condition === null) {
    if (clearPivotFieldFilterState(pivot, field.id, 'date')) refreshPivotTable(workbook, pivot, mode)
    return
  }
  const next = normalizeTimelineDateCondition(condition)
  if (!getPivotFieldFilterCapabilities(workbook, pivot, fieldId).supportsTimelineRange) throw new Error('Timeline date conditions require a typed date source field')
  if (replacePivotFilterCategory(pivot, field.id, 'condition', { dateFilter: timelineDateConditionToPivot(next), timelineDateCondition: next }, true)) {
    refreshPivotTable(workbook, pivot, mode)
  }
}

export function setPivotLabelConditionFilter(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, filter: PivotLabelConditionFilter | null,
  mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Caption conditions require a label field')
  const changed = filter === null ? clearPivotFieldFilterState(pivot, field.id, 'label')
    : replacePivotFilterCategory(pivot, field.id, 'condition', { labelFilter: normalizePivotLabelConditionFilter(filter) }, true)
  if (changed) refreshPivotTable(workbook, pivot, mode)
}

export function clearPivotFieldFilters(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  if (clearPivotFieldFilterState(pivot, fieldId, 'all')) refreshPivotTable(workbook, pivot, mode)
}

export function setPivotValueFilter(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, filter: PivotFieldFilter | null, mode: PivotRefreshMode = 'worksheet'): void {
  assertPivotHierarchyRefreshMode(mode)
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Value filters require a row, column, or filter field')
  if (filter && !pivot.fields.some((candidate) => candidate.id === filter.valueFieldId && candidate.area === 'value')) {
    throw new Error(`Value field not found: ${filter.valueFieldId}`)
  }
  const changed = filter ? replacePivotFilterCategory(pivot, field.id, 'value', { valueFilter: normalizePivotValueFilter(filter) }, true)
    : clearPivotFieldFilterState(pivot, field.id, 'value')
  if (changed) refreshPivotTable(workbook, pivot, mode)
}

/** Configure numeric buckets for a row/column/filter label field. Passing null removes grouping. */
export function setPivotNumberGroup(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, settings: PivotNumberGroup | null, mode: PivotRefreshMode = 'worksheet'): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Numeric grouping requires a label field')
  if (settings != null) validateNumberGroup(settings)
  const before = JSON.stringify(field)
  const next = { ...field, numberGroup: settings == null ? undefined : { ...settings } }
  if (settings != null) { next.manualGroups = undefined; next.dateGroupType = undefined }
  if (JSON.stringify(next) === before) return
  assertPivotFieldGroupingChange(workbook, pivot, field, next)
  Object.assign(field, next)
  if (JSON.stringify(field) !== before) refreshPivotTable(workbook, pivot, mode)
}

/** Select the calendar grouping mode used by date label fields. */
export function setPivotDateGroupingMode(workbook: WorkbookModel, pivot: PivotTableModel, groupingMode: PivotDateGroupingMode, mode: PivotRefreshMode = 'worksheet'): void {
  if (groupingMode !== 'derived' && groupingMode !== 'excelLike') throw new TypeError(`Invalid pivot date grouping mode: ${groupingMode}`)
  if ((pivot.dateGroupingMode ?? 'derived') === groupingMode) return
  pivot.dateGroupingMode = groupingMode
  refreshPivotTable(workbook, pivot, mode)
}

/** Configure a date bucket for a row/column/filter label field. */
export function setPivotDateGroupType(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, type: PivotDateGroupType | null, mode: PivotRefreshMode = 'worksheet'): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Date grouping requires a label field')
  if (type != null && !['year', 'quarter', 'month', 'day'].includes(type)) throw new TypeError(`Invalid pivot date group type: ${type}`)
  const before = JSON.stringify(field)
  const next = { ...field, dateGroupType: type == null ? undefined : type }
  if (type != null) { next.manualGroups = undefined; next.numberGroup = undefined }
  if (JSON.stringify(next) === before) return
  assertPivotFieldGroupingChange(workbook, pivot, field, next)
  Object.assign(field, next)
  if (JSON.stringify(field) !== before) refreshPivotTable(workbook, pivot, mode)
}

/** Excel-like calendar grouping entry point; requires the explicit mode to avoid accidental text parsing. */
export function setPivotExcelDateGroup(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, type: PivotDateGroupType, mode: PivotRefreshMode = 'worksheet'): void {
  if (pivot.dateGroupingMode !== 'excelLike') throw new Error('Excel-like date grouping requires date grouping mode excelLike')
  setPivotDateGroupType(workbook, pivot, fieldId, type, mode)
}

/** Set or clear manual member groups. Every non-empty group must contain at least two distinct members. */
export function setPivotManualGroups(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string, groups: PivotManualGroup[], mode: PivotRefreshMode = 'worksheet'): void {
  const field = requirePivotField(pivot, fieldId)
  if (field.area === 'value') throw new Error('Manual grouping requires a label field')
  const normalized = normalizeManualGroups(groups)
  const before = JSON.stringify(field)
  const next = { ...field, manualGroups: normalized.length ? normalized : undefined }
  if (normalized.length) { next.numberGroup = undefined; next.dateGroupType = undefined }
  if (JSON.stringify(next) === before) return
  assertPivotFieldGroupingChange(workbook, pivot, field, next)
  Object.assign(field, next)
  if (JSON.stringify(field) !== before) refreshPivotTable(workbook, pivot, mode)
}

export function refreshPivotTable(workbook: WorkbookModel, pivot: PivotTableModel, mode: PivotRefreshMode = 'worksheet'): void {
  if (!['worksheet', 'data-only', 'deferred'].includes(mode)) throw new TypeError('Invalid pivot refresh mode')
  if (mode !== 'deferred') renderPivotTable(workbook, pivot, mode)
  pivot.revision += 1
}

export function renderPivotTable(workbook: WorkbookModel, pivot: PivotTableModel, mode: PivotRenderMode = 'worksheet'): PivotRenderResult {
  if (mode !== 'worksheet' && mode !== 'data-only') throw new TypeError('Invalid pivot render mode')
  const result = buildPivotOutput(workbook, pivot)
  if (mode === 'data-only') {
    delete pivot.outputRange; delete pivot.generatedCellKeys; delete pivot.drillDownMap; delete pivot.axisOutput
    return { generatedCellKeys: [], drillDownMap: {} }
  }
  const target = targetSheet(workbook, pivot)
  if (!result.values.length) {
    removeGeneratedCells(target, pivot)
    pivot.outputRange = undefined
    pivot.generatedCellKeys = []
    pivot.drillDownMap = {}
    pivot.axisOutput = { rowAxis: [], columnAxis: [] }
    return { generatedCellKeys: [], drillDownMap: {} }
  }
  const output = {
    startRow: pivot.anchor.row,
    startCol: pivot.anchor.col,
    endRow: pivot.anchor.row + result.values.length - 1,
    endCol: pivot.anchor.col + result.values[0]!.length - 1,
  }
  ensureNoOutputConflict(workbook, pivot, output, target)
  removeGeneratedCells(target, pivot)
  const generatedCellKeys: string[] = []
  for (let row = 0; row < result.values.length; row++) {
    for (let col = 0; col < result.values[row]!.length; col++) {
      const targetRow = output.startRow + row
      const targetCol = output.startCol + col
      const key = `${targetRow}:${targetCol}`
      const cell = createCell(targetRow, targetCol, result.values[row]![col] ?? null)
      cell.style = { ...defaultCellStyle(), font: { ...defaultCellStyle().font, bold: row < result.headerRows } }
      const numberFormat = result.numberFormats[`${row}:${col}`]
      if (numberFormat) cell.style.numberFormat = numberFormat
      cell.display = formatSheetValue(cell.value, cell.style.numberFormat, workbook.date1904) ?? cell.display
      cell.metadata = { pivotTableId: pivot.id, pivotGenerated: true }
      target.cells[key] = cell
      generatedCellKeys.push(key)
    }
  }
  pivot.outputRange = output
  pivot.generatedCellKeys = generatedCellKeys
  pivot.drillDownMap = result.drillDownMap
  pivot.axisOutput = { rowAxis: result.rowAxis, columnAxis: result.columnAxis }
  return { outputRange: output, generatedCellKeys, drillDownMap: result.drillDownMap }
}

export function getPivotTableAtCell(workbook: WorkbookModel, sheetId: string, row: number, col: number): PivotTableModel | undefined {
  return workbook.pivotTables?.find((pivot) => pivot.anchor.subUnitId === sheetId && pivot.outputRange && contains(pivot.outputRange, row, col))
}

export function pivotOutputValues(workbook: WorkbookModel, pivot: PivotTableModel): CellScalar[][] {
  const range = pivot.outputRange
  if (!range) return []
  const sheet = targetSheet(workbook, pivot)
  const result: CellScalar[][] = []
  for (let row = range.startRow; row <= range.endRow; row++) {
    const values: CellScalar[] = []
    for (let col = range.startCol; col <= range.endCol; col++) values.push(cellValue(sheet, row, col))
    result.push(values)
  }
  return result
}

/** The same calculation engine serves worksheet reports and owned chart reports. */
export function calculatePivotReport(workbook: WorkbookModel, pivot: PivotTableModel): PivotCalculationResult {
  return buildPivotOutput(workbook, pivot)
}

export function pivotDrillDown(
  workbook: WorkbookModel,
  pivot: PivotTableModel,
  row: number,
  col: number
): { headers: string[]; rows: CellScalar[][] } {
  validateIndex(row, 'pivot row offset')
  validateIndex(col, 'pivot column offset')
  const map = (pivot as PivotTableModelWithDrillDown).drillDownMap ?? {}
  const entry = map[`${row}:${col}`] ?? map[`${pivot.anchor.row + row}:${pivot.anchor.col + col}`]
  if (!entry) throw new RangeError('The selected pivot cell does not contain drill-down detail')
  const source = sourceSheet(workbook, pivot)
  const fields = getPivotSourceFields(workbook, pivot)
  const records = filteredSourceRecords(readSourceRecords(source, pivot), pivot, workbook)
  const filtered = records.filter((record) => matchesFilters(record, pivot, entry.fieldFilters, entry.fieldMembers))
  return { headers: fields.map((field) => field.fieldName), rows: filtered.map((record) => record.values) }
}

export function pivotFieldIds(pivot: PivotTableModel, area: PivotFieldArea): string[] {
  return pivot.fields.filter((field) => field.area === area).sort((a, b) => a.position - b.position).map((field) => field.id)
}

function buildPivotOutput(workbook: WorkbookModel, pivot: PivotTableModel): PivotCalculationResult {
  getPivotSourceFields(workbook, pivot)
  const collapseInfo = normalizePivotCollapseInfo(pivot, pivot.collapseInfo ?? {})
  for (const field of pivot.fields) if (field.showDataAs !== undefined) {
    if (field.area !== 'value') throw new Error('Show Values As requires a value field')
    normalizePivotShowDataAs(pivot, field.showDataAs)
  }
  const byId = new Map(pivot.fields.map((field) => [field.id, field]))
  const rows = sortedFields(pivot, 'row')
  const columns = sortedFields(pivot, 'column')
  const values = sortedFields(pivot, 'value')
  if (!values.length) return { values: [], headerRows: 0, drillDownMap: {}, numberFormats: {}, rowAxis: [], columnAxis: [] }
  const allRecords = readSourceRecords(sourceSheet(workbook, pivot), pivot)
  const records = filteredSourceRecords(allRecords, pivot, workbook)

  const rowGroups = groupsFor(records, rows, pivot, workbook.date1904)
  const colGroups = groupsFor(records, columns, pivot, workbook.date1904)
  const aggregateCache = new Map<string, CellScalar>()
  const aggregateScope = (field: PivotFieldModel, filters: PivotShowDataAsScope): CellScalar => {
    const cacheKey = JSON.stringify([field.id, Object.entries(filters).sort(([a], [b]) => a.localeCompare(b))])
    if (aggregateCache.has(cacheKey)) return aggregateCache.get(cacheKey)!
    const matching = records.filter(record => Object.entries(filters).every(([id, label]) => {
      const dimension = byId.get(id)
      return Boolean(dimension && Object.is(groupedLabel(dimension, record.values[dimension.sourceColumn - pivot.source.range.startCol] ?? null, workbook.date1904), label))
    }))
    const aggregate = aggregateValue(matching, field, pivot)
    aggregateCache.set(cacheKey, aggregate)
    return aggregate
  }
  const baseItems = (fieldId: string, filters: PivotShowDataAsScope): CellScalar[] => {
    const rowIndex = rows.findIndex(field => field.id === fieldId), fields = rowIndex >= 0 ? rows : columns
    const groups = rowIndex >= 0 ? rowGroups : colGroups, index = fields.findIndex(field => field.id === fieldId)
    if (index < 0) throw new Error('Show Values As base field is not available on the axes')
    const peers = groups.filter(group => fields.slice(0, index).every((field, level) =>
      !Object.hasOwn(filters, field.id) || Object.is(group.values[level], filters[field.id])))
    return [...new Set(peers.map(group => group.values[index] ?? null))]
  }
  const valuesLayout = getPivotValuesLayout(pivot), valueRows = valuesLayout.area === 'row', valueIndex = valuesLayout.index
  const emptyGroup = (): Group => ({ values: [], key: '[]', memberValues: {} })
  const rowAxis = rows.length ? projectPivotAxis(rowGroups, rows, collapseInfo, pivot.options.showRowSubtotals) : [{ group: emptyGroup() }]
  const colAxis = columns.length ? projectPivotAxis(colGroups, columns, collapseInfo, pivot.options.showColSubtotals) : [{ group: emptyGroup() }]
  const rowScopes: OutputRow[] = rowAxis.map(entry => ({ ...entry, total: false }))
  if (pivot.options.showRowGrandTotal && rows.length) rowScopes.push({ group: null, total: true })
  const colScopes: OutputColumn[] = colAxis.map(entry => ({ ...entry, total: false }))
  if (pivot.options.showColGrandTotal && columns.length) colScopes.push({ group: null, total: true })
  const outputRows: OutputRow[] = valueRows ? expandPivotValuesAxis(rowScopes, values, valueIndex) : rowScopes
  const outputColumns: OutputColumn[] = valueRows ? colScopes : expandPivotValuesAxis(colScopes, values, valueIndex)

  const headerRows = Math.max(1, columns.length + (valueRows ? 0 : 1))
  const width = rows.length + (valueRows ? 1 : 0) + outputColumns.length
  // Allocate both the header and detail rows before populating the sparse output.
  // Pivot output is a rectangular range even when a field has no matching records.
  const matrix: CellScalar[][] = Array.from({ length: headerRows + outputRows.length }, () => Array.from({ length: width }, () => null))
  const drillDownMap: Record<string, PivotDrillDownEntry> = {}
  const numberFormats: Record<string, string> = {}
  const rowHeaderCount = rows.length + (valueRows ? 1 : 0)
  const rowLabelColumn = (index: number): number => index + (valueRows && index >= valueIndex ? 1 : 0)
  const columnLabelRow = (level: number): number => level + (!valueRows && level >= valueIndex ? 1 : 0)
  rows.forEach((field, index) => { matrix[headerRows - 1]![rowLabelColumn(index)] = field.displayName ?? field.name })
  if (valueRows) matrix[headerRows - 1]![valueIndex] = 'Values'
  for (let colIndex = 0; colIndex < outputColumns.length; colIndex++) {
    const outputColumn = outputColumns[colIndex]!
    const targetColumn = rowHeaderCount + colIndex
    columns.forEach((_field, level) => {
      const value = outputColumn.group?.values[level] ?? null
      const targetHeaderRow = columnLabelRow(level)
      if (outputColumn.total) matrix[targetHeaderRow]![targetColumn] = level === 0 ? 'Grand total' : null
      else if (outputColumn.subtotalLevel === level) matrix[targetHeaderRow]![targetColumn] = `${value ?? ''} Total`
      else matrix[targetHeaderRow]![targetColumn] = value
    })
    if (!valueRows) matrix[valueIndex]![targetColumn] = outputColumn.total
      ? (outputColumn.valueFieldId ? displayValueField(byId.get(outputColumn.valueFieldId)!, values.length) : 'Grand total')
      : outputColumn.valueFieldId ? displayValueField(byId.get(outputColumn.valueFieldId)!, values.length) : 'Value'
    else if (headerRows === 1) matrix[0]![targetColumn] = columns.length ? outputColumn.group?.values[0] ?? null : 'Value'
  }
  if (columns.length && !valueRows && values.length > 1) {
    for (let i = 0; i < outputColumns.length; i++) {
      const outputColumn = outputColumns[i]!
      if (outputColumn.total) matrix[valueIndex]![rowHeaderCount + i] = `Grand total ${displayValueField(byId.get(outputColumn.valueFieldId!)!, values.length)}`
    }
  }

  for (let rowIndex = 0; rowIndex < outputRows.length; rowIndex++) {
    const outputRow = outputRows[rowIndex]!
    const targetRow = headerRows + rowIndex
    const rowFieldFilters: Record<string, CellScalar> = Object.create(null)
    const rowFieldMembers: Record<string, CellScalar[]> = Object.create(null)
    if (outputRow.group) rows.forEach((field, index) => {
      if (index >= outputRow.group!.values.length) return
      rowFieldFilters[field.id] = outputRow.group!.values[index] ?? null
      rowFieldMembers[field.id] = outputRow.group!.memberValues[field.id] ?? [outputRow.group!.values[index] ?? null]
    })
    rows.forEach((_field, index) => {
      const value = outputRow.group?.values[index] ?? null
      const previous = outputRows[rowIndex - 1]
      const targetLabelColumn = rowLabelColumn(index)
      const repeats = previous?.group && outputRow.group && previous.subtotalLevel === undefined && outputRow.subtotalLevel === undefined
        && (!valueRows || index < valueIndex || previous.valueFieldId === outputRow.valueFieldId)
        && outputRow.group.values.slice(0, index + 1).every((label, level) => Object.is(label, previous.group!.values[level]))
      if (outputRow.total) matrix[targetRow]![targetLabelColumn] = index === 0 ? 'Grand total' : null
      else if (outputRow.subtotalLevel === index) matrix[targetRow]![targetLabelColumn] = `${value ?? ''} Total`
      else if (!pivot.options.repeatRowLabels && repeats) matrix[targetRow]![targetLabelColumn] = null
      else matrix[targetRow]![targetLabelColumn] = value
    })
    if (outputRow.valueFieldId && valueRows) {
      const valueField = byId.get(outputRow.valueFieldId)!
      matrix[targetRow]![valueIndex] = valueField.displayName ?? valueField.name
    }
    for (let colIndex = 0; colIndex < outputColumns.length; colIndex++) {
      const outputColumn = outputColumns[colIndex]!
      const valueFieldId = valueRows ? outputRow.valueFieldId! : outputColumn.valueFieldId!
      const colFieldFilters: Record<string, CellScalar> = Object.create(null)
      const colFieldMembers: Record<string, CellScalar[]> = Object.create(null)
      if (outputColumn.group) columns.forEach((field, index) => {
        if (index >= outputColumn.group!.values.length) return
        colFieldFilters[field.id] = outputColumn.group!.values[index] ?? null
        colFieldMembers[field.id] = outputColumn.group!.memberValues[field.id] ?? [outputColumn.group!.values[index] ?? null]
      })
      const groupFilter = { ...rowFieldFilters, ...colFieldFilters }
      const groupMembers = { ...rowFieldMembers, ...colFieldMembers }
      const valueField = byId.get(valueFieldId)!
      const aggregate = aggregateScope(valueField, groupFilter)
      const cellColumn = rowHeaderCount + colIndex
      matrix[targetRow]![cellColumn] = applyPivotShowDataAs(valueField.showDataAs, { value: aggregate, rowFields: rows, columnFields: columns,
        rowFilters: rowFieldFilters, columnFilters: colFieldFilters, aggregate: filters => aggregateScope(valueField, filters), baseItems })
      const numberFormat = valueField.format ?? pivotShowDataAsNumberFormat(valueField.showDataAs)
      if (numberFormat) numberFormats[`${targetRow}:${cellColumn}`] = numberFormat
      drillDownMap[`${targetRow}:${cellColumn}`] = { fieldFilters: groupFilter, fieldMembers: groupMembers, valueFieldId }
    }
  }
  return { values: matrix, headerRows, drillDownMap, numberFormats,
    rowAxis: outputRows.map((entry, index) => pivotAxisOutputEntry(entry, rows, headerRows + index)),
    columnAxis: outputColumns.map((entry, index) => pivotAxisOutputEntry(entry, columns, rowHeaderCount + index)) }
}

function pivotAxisOutputEntry(entry: OutputRow | OutputColumn, fields: readonly PivotFieldModel[], index: number): PivotAxisOutputEntry {
  const labels = entry.group ? [...entry.group.values] : [], included = fields.slice(0, labels.length)
  const fieldFilters = Object.fromEntries(included.map((field, level) => [field.id, labels[level] ?? null]))
  const fieldMembers = Object.fromEntries(included.map((field, level) => [field.id, [...(entry.group!.memberValues[field.id] ?? [labels[level] ?? null])]]))
  const fieldId = fields[labels.length - 1]?.id
  return { index, kind: entry.total ? 'grand-total' : entry.collapsedLevel !== undefined ? 'collapsed'
    : entry.subtotalLevel !== undefined ? 'subtotal' : 'detail', level: labels.length - 1, labels, fieldFilters, fieldMembers,
    ...(fieldId ? { fieldId } : {}), ...(entry.valueFieldId ? { valueFieldId: entry.valueFieldId } : {}) }
}

function filteredSourceRecords(records: SourceRecord[], pivot: PivotTableModel, workbook: WorkbookModel): SourceRecord[] {
  const dimensions = pivot.fields.filter(field => field.area !== 'value')
  const source = sourceSheet(workbook, pivot)
  const predicates = dimensions.map(field => ({ field, member: pivotMemberPredicate(field, workbook.date1904),
    label: field.labelFilter ? compilePivotLabelConditionFilter(field.labelFilter) : undefined,
    date: field.timelineDateRange ? compileTimelineSerialRange(field.timelineDateRange)
      : field.timelineDateCondition ? compileTimelineSerialCondition(field.timelineDateCondition)
        : field.dateFilter ? compilePivotDateFilter(field.dateFilter) : undefined }))
  const selected = records.filter(record => predicates.every(({ field, member, label, date }) => {
    const value = record.values[field.sourceColumn - pivot.source.range.startCol] ?? null
    if (!member(value, sheetCellRecordValue(source.cells[`${record.row}:${field.sourceColumn}`], workbook.date1904))
      || label && !label(String(groupedLabel(field, value, workbook.date1904) ?? ''))) return false
    const serial = !date ? undefined : field.timelineDateRange || field.timelineDateCondition
      ? sourceTimelineDateSerial(source, record.row, field, workbook.date1904)
      : sourceDateSerial(source, record.row, field, workbook.date1904)
    return !date || serial !== undefined && date(serial)
  }))
  return applyValueFilters(selected, dimensions, new Map(pivot.fields.map(field => [field.id, field])), pivot, workbook.date1904)
}

/** Cross-filter availability reuses the report's actual predicates, excluding every filter of this source field. */
export function getPivotSourceAvailableRows(workbook: WorkbookModel, pivot: PivotTableModel, fieldId: string): number[] {
  const report: PivotTableModel = { ...pivot, fields: pivot.fields.map(field => ({ ...field })) }
  clearPivotFieldFilterState(report, fieldId, 'all')
  return filteredSourceRecords(readSourceRecords(sourceSheet(workbook, report), report), report, workbook).map(record => record.row)
}

function applyValueFilters(records: SourceRecord[], dimensionFields: PivotFieldModel[], byId: Map<string, PivotFieldModel>, pivot: PivotTableModel, date1904: boolean): SourceRecord[] {
  const active = dimensionFields.filter((field) => field.valueFilter)
  if (!active.length) return records
  return records.filter((record) => active.every((field) => {
    const filter = field.valueFilter!
    const valueField = byId.get(filter.valueFieldId)
    if (!valueField) return false
    const label = groupedLabel(field, record.values[field.sourceColumn - pivot.source.range.startCol] ?? null, date1904)
    const matching = records.filter(candidate => Object.is(groupedLabel(field, candidate.values[field.sourceColumn - pivot.source.range.startCol] ?? null, date1904), label))
    const value = aggregateValue(matching, valueField, pivot)
    return typeof value === 'number' && testFilter(value, filter)
  }))
}

function testFilter(value: number, filter: PivotFieldFilter): boolean {
  switch (filter.operator) {
    case 'valueGreaterThan': return value > filter.expected
    case 'valueGreaterThanOrEqual': return value >= filter.expected
    case 'valueLessThan': return value < filter.expected
    case 'valueLessThanOrEqual': return value <= filter.expected
    case 'valueEqual': return value === filter.expected
    case 'valueNotEqual': return value !== filter.expected
    case 'valueBetween': return value >= Math.min(filter.expected, filter.expected2 ?? filter.expected) && value <= Math.max(filter.expected, filter.expected2 ?? filter.expected)
    case 'valueNotBetween': return value < Math.min(filter.expected, filter.expected2 ?? filter.expected) || value > Math.max(filter.expected, filter.expected2 ?? filter.expected)
  }
}

function groupsFor(records: SourceRecord[], fields: PivotFieldModel[], pivot: PivotTableModel, date1904: boolean): Group[] {
  if (!fields.length) return []
  const result = new Map<string, Group>()
  for (const record of records) {
    const values = fields.map((field) => groupedLabel(field, record.values[field.sourceColumn - pivot.source.range.startCol] ?? null, date1904))
    const key = JSON.stringify(values)
    if (!result.has(key)) {
      const memberValues: Record<string, CellScalar[]> = Object.create(null)
      fields.forEach((field) => { memberValues[field.id] = [record.values[field.sourceColumn - pivot.source.range.startCol] ?? null] })
      result.set(key, { values, key, memberValues })
    } else {
      const group = result.get(key)!
      fields.forEach((field) => {
        const sourceValue = record.values[field.sourceColumn - pivot.source.range.startCol] ?? null
        const members = group.memberValues[field.id] ?? (group.memberValues[field.id] = [])
        if (!members.some((member) => Object.is(member, sourceValue))) members.push(sourceValue)
      })
    }
  }
  return orderGroups([...result.values()], fields, 0)
}

function orderGroups(groups: Group[], fields: PivotFieldModel[], level: number): Group[] {
  if (level >= fields.length) return groups
  const siblings = new Map<string, Group[]>()
  for (const group of groups) {
    const key = JSON.stringify(group.values[level])
    const members = siblings.get(key) ?? []
    members.push(group)
    siblings.set(key, members)
  }
  const ordered = [...siblings.values()]
  const direction = fields[level]!.sort
  if (direction) ordered.sort((first, second) => {
    const order = compareLabels(first[0]!.values[level] ?? null, second[0]!.values[level] ?? null)
    return direction === 'ascending' ? order : -order
  })
  return ordered.flatMap(members => orderGroups(members, fields, level + 1))
}

function aggregateValue(records: SourceRecord[], field: PivotFieldModel, pivot: PivotTableModel): CellScalar {
  const index = field.sourceColumn - pivot.source.range.startCol
  const values = records.map((record) => record.values[index]).filter((value) => value !== null && value !== '')
  const aggregation = field.aggregation ?? 'sum'
  if (aggregation === 'count') return values.length
  const numericValues = values.filter((value): value is number => typeof value === 'number' && Number.isFinite(value))
  if (aggregation === 'countNums') return numericValues.length
  if (!numericValues.length) return null
  if (aggregation === 'product') return numericValues.reduce((product, value) => product * value, 1)
  if (aggregation === 'stdDev' || aggregation === 'stdDevp' || aggregation === 'var' || aggregation === 'varp') {
    if ((aggregation === 'stdDev' || aggregation === 'var') && numericValues.length < 2) return null
    const mean = numericValues.reduce((total, value) => total + value, 0) / numericValues.length
    const divisor = aggregation === 'stdDevp' || aggregation === 'varp' ? numericValues.length : numericValues.length - 1
    const variance = numericValues.reduce((total, value) => total + (value - mean) ** 2, 0) / divisor
    return aggregation === 'stdDev' || aggregation === 'stdDevp' ? Math.sqrt(variance) : variance
  }
  if (aggregation === 'min') return Math.min(...numericValues)
  if (aggregation === 'max') return Math.max(...numericValues)
  const sum = numericValues.reduce((total, value) => total + value, 0)
  if (aggregation === 'average') return sum / numericValues.length
  return sum
}

function readSourceRecords(sheet: SpreadsheetSheet, pivot: PivotTableModel): SourceRecord[] {
  const records: SourceRecord[] = []
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    records.push({ row, values: Array.from({ length: pivot.source.range.endCol - pivot.source.range.startCol + 1 }, (_, index) => cellValue(sheet, row, pivot.source.range.startCol + index)) })
  }
  return records
}

function resolveField(workbook: WorkbookModel, pivot: PivotTableModel, value: string | number): PivotSourceFieldInfo {
  const fields = getPivotSourceFields(workbook, pivot)
  const normalized = String(value)
  const field = fields.find((candidate) => candidate.fieldId === normalized)
  if (!field) throw new Error(`Pivot source field not found: ${normalized}`)
  return field
}

function sourceSheet(workbook: WorkbookModel, pivot: PivotTableModel): SpreadsheetSheet {
  if (workbook.id !== pivot.source.unitId) throw new Error('Pivot source workbook is not available')
  const sheet = workbook.sheets.find((candidate) => candidate.id === pivot.source.subUnitId)
  if (!sheet) throw new Error(`Pivot source worksheet not found: ${pivot.source.subUnitId}`)
  validateRange(pivot.source.range)
  if (pivot.source.range.endRow <= pivot.source.range.startRow) throw new RangeError('Pivot source must contain a header and at least one data row')
  if (pivot.source.range.endRow >= (sheet.rowCount ?? 1_048_576) || pivot.source.range.endCol >= (sheet.columnCount ?? 16_384)) throw new RangeError('Pivot source exceeds worksheet dimensions')
  return sheet
}

function targetSheet(workbook: WorkbookModel, pivot: PivotTableModel): SpreadsheetSheet {
  if (workbook.id !== pivot.anchor.unitId) throw new Error('Pivot target workbook is not available')
  const sheet = workbook.sheets.find((candidate) => candidate.id === pivot.anchor.subUnitId)
  if (!sheet) throw new Error(`Pivot target worksheet not found: ${pivot.anchor.subUnitId}`)
  return sheet
}

function cellValue(sheet: SpreadsheetSheet, row: number, col: number): CellScalar {
  const cell = sheet.cells[`${row}:${col}`]
  return cell?.value ?? null
}

function removeGeneratedCells(sheet: SpreadsheetSheet, pivot: PivotTableModel): void {
  for (const key of pivot.generatedCellKeys ?? []) {
    if (sheet.cells[key]?.metadata?.pivotTableId === pivot.id) delete sheet.cells[key]
  }
}

function ensureNoOutputConflict(workbook: WorkbookModel, pivot: PivotTableModel, range: CellRange, target: SpreadsheetSheet): void {
  if (pivot.source.subUnitId === pivot.anchor.subUnitId && overlaps(range, pivot.source.range)) throw new PivotOutputConflictError('Pivot output cannot overlap its source range')
  const allowed = new Set(pivot.generatedCellKeys ?? [])
  for (let row = range.startRow; row <= range.endRow; row++) for (let col = range.startCol; col <= range.endCol; col++) {
    const key = `${row}:${col}`
    const existing = target.cells[key]
    // Clearing a range leaves a sparse cell record with null content in the
    // public API. It must not keep an old pivot anchor blocked forever, while
    // formulas, values, and displayed text still count as real content.
    const hasContent = existing != null && (existing.value !== null || existing.formula !== null || existing.display !== '')
    if (hasContent && !allowed.has(key)) throw new PivotOutputConflictError(`Pivot output overlaps existing content at ${key}`)
  }
  for (const other of workbook.pivotTables ?? []) {
    if (other.id !== pivot.id && other.anchor.subUnitId === pivot.anchor.subUnitId && other.outputRange && overlaps(range, other.outputRange)) throw new PivotOutputConflictError(`Pivot output overlaps pivot table ${other.id}`)
  }
}

function matchesFilters(record: SourceRecord, pivot: PivotTableModel, fieldFilters: Record<string, CellScalar>, fieldMembers: Record<string, CellScalar[]> = {}): boolean {
  return Object.entries(fieldFilters).every(([fieldId, value]) => {
    const field = pivot.fields.find((candidate) => candidate.id === fieldId)
    if (!field) return false
    const sourceValue = record.values[field.sourceColumn - pivot.source.range.startCol] ?? null
    return (fieldMembers[fieldId] ?? [value]).some((member) => Object.is(member, sourceValue))
  })
}

function areaFields(pivot: PivotTableModel, area: PivotFieldArea, omitId?: string): PivotFieldModel[] {
  return pivot.fields.filter((field) => field.area === area && field.id !== omitId).sort((a, b) => a.position - b.position)
}

function sortedFields(pivot: PivotTableModel, area: PivotFieldArea): PivotFieldModel[] {
  return areaFields(pivot, area)
}

function requirePivotField(pivot: PivotTableModel, id: string): PivotFieldModel {
  const field = pivot.fields.find((candidate) => candidate.id === String(id))
  if (!field) throw new Error(`Pivot field not configured: ${id}`)
  return field
}

function reindexFields(pivot: PivotTableModel): void {
  for (const area of ['filter', 'row', 'column', 'value'] as const) areaFields(pivot, area).forEach((field, index) => { field.position = index })
}

function normalizePosition(position: number | undefined, length: number): number {
  if (position == null) return length
  if (!Number.isInteger(position) || position < 0 || position > length) throw new RangeError(`Invalid pivot field position: ${position}`)
  return position
}

function displayValueField(field: PivotFieldModel, count: number): string {
  return field.displayName ?? (count > 1 ? field.name : field.aggregation === 'count' ? `Count of ${field.name}` : field.name)
}

function compareLabels(a: CellScalar, b: CellScalar): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b
  return String(a ?? '').localeCompare(String(b ?? ''), undefined, { numeric: true, sensitivity: 'base' })
}

function groupedLabel(field: PivotFieldModel, value: CellScalar, date1904 = false): CellScalar {
  if (field.manualGroups?.length) {
    const group = field.manualGroups.find((candidate) => candidate.members.some((member) => Object.is(member, value)))
    return group?.name ?? value
  }
  if (field.numberGroup && typeof value === 'number' && Number.isFinite(value)) {
    const { start, end, interval } = field.numberGroup
    if (value < start) return `< ${start}`
    if (value > end) return `> ${end}`
    const bucket = Math.min(Math.floor((value - start) / interval), Math.max(0, Math.ceil((end - start) / interval) - 1))
    const lower = start + bucket * interval
    const upper = Math.min(end, lower + interval)
    return `${formatGroupNumber(lower)}-${formatGroupNumber(upper)}`
  }
  if (field.dateGroupType) return dateGroupLabel(value, field.dateGroupType, date1904)
  return value
}

function showDataAsBaseItems(workbook: WorkbookModel, pivot: PivotTableModel, field: PivotFieldModel): CellScalar[] {
  const source = sourceSheet(workbook, pivot), items = new Set<CellScalar>()
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row++) {
    items.add(groupedLabel(field, cellValue(source, row, field.sourceColumn), workbook.date1904))
  }
  return [...items]
}

function assertPivotFieldGroupingChange(workbook: WorkbookModel, pivot: PivotTableModel, field: PivotFieldModel, next: PivotFieldModel): void {
  const state = pivot.collapseInfo?.[field.id]
  if (state && typeof state === 'object') {
    assertPivotCollapseGroupingChange(state, showDataAsBaseItems(workbook, pivot, field), showDataAsBaseItems(workbook, pivot, next))
  }
}

function dateGroupLabel(value: CellScalar, type: PivotDateGroupType, date1904: boolean): CellScalar {
  const date = parseTimelineDate(value, date1904)
  if (!date) return value
  const { year, month, day } = date
  if (type === 'year') return String(year)
  if (type === 'quarter') return `${year} Q${Math.floor((month - 1) / 3) + 1}`
  if (type === 'month') return `${year}-${String(month).padStart(2, '0')}`
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
}

function formatGroupNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(12)))
}

function validateNumberGroup(settings: PivotNumberGroup): void {
  if (![settings.start, settings.end, settings.interval].every(Number.isFinite)) throw new TypeError('Pivot number group bounds must be finite')
  if (settings.end <= settings.start) throw new RangeError('Pivot number group end must be greater than start')
  if (settings.interval <= 0) throw new RangeError('Pivot number group interval must be positive')
}

function normalizeManualGroups(groups: PivotManualGroup[]): PivotManualGroup[] {
  if (!Array.isArray(groups)) throw new TypeError('Pivot manual groups must be an array')
  const names = new Set<string>()
  const usedMembers: CellScalar[] = []
  return groups.map((group, index) => {
    if (!group || typeof group.name !== 'string' || !group.name.trim()) throw new TypeError(`Pivot manual group ${index} requires a name`)
    if (!Array.isArray(group.members) || group.members.length < 2) throw new RangeError(`Pivot manual group ${index} must contain at least two members`)
    if (group.members.some(member => member !== null && typeof member !== 'string' && typeof member !== 'boolean' && (typeof member !== 'number' || !Number.isFinite(member)))) {
      throw new TypeError('Pivot manual group members must be scalar source values')
    }
    const members = group.members.filter((member, memberIndex, all) => all.findIndex((candidate) => Object.is(candidate, member)) === memberIndex)
    if (members.length < 2) throw new RangeError(`Pivot manual group ${index} must contain at least two distinct members`)
    if (names.has(group.name.trim())) throw new Error('Pivot manual group names must be unique')
    if (members.some(member => usedMembers.some(used => Object.is(used, member)))) throw new Error('Pivot manual group members cannot overlap')
    names.add(group.name.trim())
    usedMembers.push(...members)
    return { name: group.name.trim(), members: [...members] }
  })
}

function overlaps(a: CellRange, b: CellRange): boolean {
  return !(a.endRow < b.startRow || a.startRow > b.endRow || a.endCol < b.startCol || a.startCol > b.endCol)
}

function contains(range: CellRange, row: number, col: number): boolean {
  return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
}

function validateRange(range: CellRange): void {
  validateIndex(range.startRow, 'source start row')
  validateIndex(range.startCol, 'source start column')
  validateIndex(range.endRow, 'source end row')
  validateIndex(range.endCol, 'source end column')
  if (range.endRow < range.startRow || range.endCol < range.startCol) throw new RangeError('Invalid pivot source range')
}

function validateIndex(value: number, name: string): void {
  if (!Number.isInteger(value) || value < 0) throw new RangeError(`${name} must be a non-negative integer`)
}
