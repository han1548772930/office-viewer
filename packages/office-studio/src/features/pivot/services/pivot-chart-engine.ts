import type { ChartModel, PivotAggregation, PivotCollapseInfo, PivotDateFilter, PivotDateGroupType, PivotDateGroupingMode, PivotFieldArea,
  PivotFieldFilter, PivotFieldModel, PivotLabelConditionFilter, PivotManualGroup, PivotNumberGroup, PivotSortDirection,
  PivotShowDataAsInfo, PivotTableModel, PivotTableOptions, PivotValuePosition, WorkbookModel } from '../../../domain/model.ts'
import { cloneModel } from '../../../domain/serialization.ts'
import { equivalentModelData } from '../../../platform/commands/model-equivalence.ts'
import { ensurePivotCache, getSlicerConnectionReports, sourceFieldId, SlicerConnectionError } from '../../slicer/slicer-connections.ts'
import { clearPivotFieldFilters, createDetachedPivotReport, getPivotDateFilter, getPivotLabelConditionFilter, getPivotManualFilterItemState,
  getPivotSourceSelectedItems, getPivotFieldFilterCapabilities, getPivotSourceFields, getPivotSourceFieldItems, getPivotGroupedFieldItems, getPivotShowDataAsBaseItems, getPivotAxisState,
  getPivotCollapseItems, refreshPivotTable, removePivotFields, setPivotCategoryLevel, setPivotCollapse, setPivotCollapseInfo, setPivotDateGroupType,
  setPivotDateFilter, setPivotDateGroupingMode, setPivotExcelDateGroup, setPivotField, setPivotFieldSetting, setPivotLabelConditionFilter, setPivotLabelFilter,
  setPivotManualGroups, setPivotNumberGroup, setPivotOptions, setPivotShowDataAs, setPivotSort, setPivotValueFilter, setPivotValuesPosition, stepPivotCategoryLevel,
  type PivotSourceFieldInfo, type PivotTableSourceInput } from './pivot-engine.ts'
import { normalizePivotCollapseInfo, retainPivotCollapseInfo, type PivotAxisArea, type PivotAxisState } from './pivot-axis-state.ts'
import { getPivotFieldFilterCategories, normalizePivotDateFilter, normalizePivotLabelConditionFilter,
  normalizePivotRecordSelection, normalizePivotValueFilter } from './pivot-filter-predicate.ts'

export interface PivotChartFieldsConfig {
  fields: PivotFieldModel[]
  options?: Partial<PivotTableOptions>
  valuePosition?: PivotValuePosition
  valueIndex?: number
  dateGroupingMode?: PivotDateGroupingMode
  collapseInfo?: PivotCollapseInfo | null
}

export interface PivotChartFieldSetting {
  displayName?: string
  format?: string
  subtotalType?: PivotAggregation
  showDataAs?: PivotShowDataAsInfo | null
}

/** Source IDs are resolved to report-local field instances before calling the shared engine. */
export type PivotChartReportEdit =
  | { type: 'setFieldsConfig'; config: PivotChartFieldsConfig }
  | { type: 'resetFields'; area?: PivotFieldArea }
  | { type: 'addField'; dataFieldId: string | number; fieldArea: PivotFieldArea; index?: number }
  | { type: 'moveField'; fieldId: string; fieldArea: PivotFieldArea; index: number }
  | { type: 'removeFields'; fieldIds: string[] }
  | { type: 'setOptions'; options: Partial<PivotTableOptions> }
  | { type: 'updateValuePosition'; position: PivotValuePosition; index?: number }
  | { type: 'setFieldSetting'; fieldId: string; setting: PivotChartFieldSetting }
  | { type: 'setShowDataAs'; fieldId: string; showDataAs: PivotShowDataAsInfo | null }
  | { type: 'setCollapse'; fieldId: string; collapse: boolean; item?: string }
  | { type: 'setCategoryLevel'; fieldId: string }
  | { type: 'stepCategoryLevel'; direction: 'expand' | 'collapse' }
  | { type: 'setSort'; fieldId: string; direction: PivotSortDirection | null }
  | { type: 'setFilter'; fieldId: string; items: string[] | null; isAll?: boolean; partialItems?: string[]; clearAllFilters?: boolean }
  | { type: 'setValueFilter'; fieldId: string; filter: PivotFieldFilter | null }
  | { type: 'setDateFilter'; fieldId: string; filter: PivotDateFilter | null }
  | { type: 'setLabelConditionFilter'; fieldId: string; filter: PivotLabelConditionFilter | null }
  | { type: 'setNumberGroup'; fieldId: string; group: PivotNumberGroup | null }
  | { type: 'setDateGroupingMode'; mode: PivotDateGroupingMode }
  | { type: 'setDateGroup'; fieldId: string; dateType: PivotDateGroupType | null; excelLike?: boolean }
  | { type: 'setManualGroups'; fieldId: string; groups: PivotManualGroup[] }
  | { type: 'refresh' }

export class PivotChartEditError extends Error {
  constructor(readonly code: 'invalid-resource' | 'invalid-edit' | 'field-connected', message: string) {
    super(message); this.name = 'PivotChartEditError'
  }
}

const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
const identifier = (value: unknown): value is string => typeof value === 'string' && value.trim().length > 0
const own = (value: object, key: PropertyKey): boolean => Object.hasOwn(value, key)
const areas: readonly PivotFieldArea[] = ['filter', 'row', 'column', 'value']
const options = ['repeatRowLabels', 'showRowGrandTotal', 'showColGrandTotal', 'showRowSubtotals', 'showColSubtotals', 'valuePosition', 'valueIndex', 'allowMultipleFilters']
const editProperties: Record<PivotChartReportEdit['type'], readonly string[]> = {
  setFieldsConfig: ['config'], resetFields: ['area'], addField: ['dataFieldId', 'fieldArea', 'index'],
  moveField: ['fieldId', 'fieldArea', 'index'], removeFields: ['fieldIds'], setOptions: ['options'],
  updateValuePosition: ['position', 'index'], setFieldSetting: ['fieldId', 'setting'], setSort: ['fieldId', 'direction'],
  setShowDataAs: ['fieldId', 'showDataAs'],
  setCollapse: ['fieldId', 'collapse', 'item'], setCategoryLevel: ['fieldId'], stepCategoryLevel: ['direction'],
  setFilter: ['fieldId', 'items', 'isAll', 'partialItems', 'clearAllFilters'], setValueFilter: ['fieldId', 'filter'],
  setDateFilter: ['fieldId', 'filter'], setLabelConditionFilter: ['fieldId', 'filter'],
  setNumberGroup: ['fieldId', 'group'], setDateGroupingMode: ['mode'], setDateGroup: ['fieldId', 'dateType', 'excelLike'],
  setManualGroups: ['fieldId', 'groups'], refresh: [],
}
const fieldProperties = ['id', 'sourceFieldId', 'sourceColumn', 'name', 'area', 'position', 'displayName', 'aggregation', 'showDataAs',
  'format', 'sort', 'selectedItems', 'groupedItems', 'partialItems', 'recordSelection', 'labelFilter', 'dateFilter',
  'valueFilter', 'numberGroup', 'dateGroupType', 'manualGroups']

/** Own the report independently; its anchor records the real chart host, never an output reservation. */
export function createPivotChartReport(workbook: WorkbookModel, source: PivotTableSourceInput, hostSheetId: string,
  id = `owned-pivot-${crypto.randomUUID()}`): PivotTableModel {
  if (!identifier(workbook.id) || !identifier(hostSheetId) || !identifier(id) || (workbook.pivotTables ?? []).some(report => report.id === id)
    || (workbook.charts ?? []).some(chart => chart.ownedPivot?.id === id)) throw new PivotChartEditError('invalid-resource', '透视图报告 ID 必须唯一。')
  const report = createDetachedPivotReport(workbook, source, { unitId: workbook.id, subUnitId: hostSheetId, row: 0, col: 0 }, id)
  ensurePivotCache(workbook, report)
  return report
}

export function getPivotChartReport(workbook: WorkbookModel, chart: ChartModel): PivotTableModel {
  if (!identifier(workbook.id) || !workbook.sheets.some(sheet => sheet.id === chart.sheetId)
    || chart.pivotTableId !== undefined && !identifier(chart.pivotTableId)
    || Boolean(chart.pivotTableId) === Boolean(chart.ownedPivot)) {
    throw new PivotChartEditError('invalid-resource', '透视图必须持有一个范围报告或关联透视表。')
  }
  const report = chart.ownedPivot ?? workbook.pivotTables?.find(report => report.id === chart.pivotTableId)
  if (!report || !record(report.source) || !record(report.anchor) || !record(report.source.range)
    || !Array.isArray(report.fields) || !record(report.options)
    || report.source.unitId !== workbook.id || report.anchor.unitId !== workbook.id
    || !workbook.sheets.some(sheet => sheet.id === report.anchor.subUnitId)
    || chart.ownedPivot && report.anchor.subUnitId !== chart.sheetId) {
    throw new PivotChartEditError('invalid-resource', '透视图的报告或所属工作簿当前不可用。')
  }
  try { getPivotSourceFields(workbook, report) }
  catch (error) { if (error instanceof Error) throw new PivotChartEditError('invalid-resource', error.message); throw error }
  return report
}

export function getPivotChartSourceFields(workbook: WorkbookModel, chart: ChartModel): PivotSourceFieldInfo[] {
  return getPivotSourceFields(workbook, getPivotChartReport(workbook, chart))
}

export function getPivotChartConfiguredField(workbook: WorkbookModel, chart: ChartModel, fieldId: string): PivotFieldModel | undefined {
  return findConfiguredField(getPivotChartReport(workbook, chart), fieldId)
}

export function getPivotChartShowDataAsBaseItems(workbook: WorkbookModel, chart: ChartModel, fieldId: string): string[] {
  return getPivotShowDataAsBaseItems(workbook, getPivotChartReport(workbook, chart), fieldId)
}

export function getPivotChartAxisState(workbook: WorkbookModel, chart: ChartModel, area: PivotAxisArea = 'row'): PivotAxisState {
  return getPivotAxisState(workbook, getPivotChartReport(workbook, chart), area)
}

export function getPivotChartCollapseItems(workbook: WorkbookModel, chart: ChartModel, fieldId: string): string[] {
  const report = getPivotChartReport(workbook, chart), field = configuredField(report, fieldId)
  return getPivotCollapseItems(workbook, report, field.id)
}

export function getPivotChartDateFilter(workbook: WorkbookModel, chart: ChartModel, fieldId: string): PivotDateFilter | undefined {
  const report = getPivotChartReport(workbook, chart)
  return getPivotDateFilter(report, configuredField(report, fieldId).id)
}

export function getPivotChartLabelConditionFilter(workbook: WorkbookModel, chart: ChartModel, fieldId: string): PivotLabelConditionFilter | undefined {
  const report = getPivotChartReport(workbook, chart)
  return getPivotLabelConditionFilter(report, configuredField(report, fieldId).id)
}

export function getPivotChartManualFilterItemState(workbook: WorkbookModel, chart: ChartModel, fieldId: string): { checkedItems: string[]; partialItems: string[] } {
  const report = getPivotChartReport(workbook, chart)
  return getPivotManualFilterItemState(workbook, report, configuredField(report, fieldId).id)
}

export function getPivotChartSourceSelectedItems(workbook: WorkbookModel, chart: ChartModel, fieldId: string): string[] | undefined {
  const report = getPivotChartReport(workbook, chart)
  return getPivotSourceSelectedItems(workbook, report, configuredField(report, fieldId).id)
}

/** Configured label fields expose grouped keys; unconfigured source metadata keeps its raw member query. */
export function getPivotChartFieldItems(workbook: WorkbookModel, chart: ChartModel, fieldId: string): string[] {
  const report = getPivotChartReport(workbook, chart), fields = getPivotSourceFields(workbook, report)
  if (!identifier(fieldId)) throw new PivotChartEditError('invalid-edit', '透视图来源字段 ID 无效。')
  const configured = findConfiguredField(report, fieldId)
  if (configured) {
    if (configured.area === 'value') throw new PivotChartEditError('invalid-edit', '值字段没有标签筛选项目。')
    return getPivotGroupedFieldItems(workbook, report, configured.id)
  }
  const info = fields.find(field => field.fieldId === fieldId)
    ?? fields.find(field => field.sourceColumn === report.fields.find(field => field.id === fieldId)?.sourceColumn)
  if (!info) throw new PivotChartEditError('invalid-edit', `透视图来源字段不存在：${fieldId}`)
  return [...new Set(getPivotSourceFieldItems(workbook, report, info.fieldId).map(value => String(value ?? '')))]
}

/** Validate/stage fields first, then calculate and publish exactly once. Shared Slicer fan-out belongs to Studio prepare. */
export function applyPivotChartEdit(workbook: WorkbookModel, chart: ChartModel, edit: PivotChartReportEdit): boolean {
  if (!record(edit) || !identifier(edit.type)) throw new PivotChartEditError('invalid-edit', '透视图字段编辑参数无效。')
  const properties = own(editProperties, edit.type) ? editProperties[edit.type] : undefined
  if (!properties || Object.keys(edit).some(key => key !== 'type' && !properties.includes(key))) {
    throw new PivotChartEditError('invalid-edit', '透视图字段编辑类型或参数尚未支持。')
  }
  const current = getPivotChartReport(workbook, chart), draft = cloneModel(current)
  try { applyFieldEdit(workbook, draft, edit) }
  catch (error) {
    if (error instanceof PivotChartEditError) throw error
    if (error instanceof Error) throw new PivotChartEditError('invalid-edit', error.message)
    throw error
  }
  if (edit.type !== 'refresh' && equivalentModelData(editState(current), editState(draft))) return false
  assertConnectedFieldsSurvive(workbook, chart, current, draft)
  draft.revision = current.revision
  const outputId = chart.ownedPivot ? undefined : draft.anchor.subUnitId
  const staged: WorkbookModel = { ...workbook,
    sheets: workbook.sheets.map(sheet => sheet.id === outputId ? { ...sheet, cells: { ...sheet.cells } } : sheet),
    pivotTables: workbook.pivotTables?.map(report => report === current ? draft : report),
  }
  refreshPivotTable(staged, draft, chart.ownedPivot ? 'data-only' : 'worksheet')
  if (outputId) {
    const output = staged.sheets.find(sheet => sheet.id === outputId)!, target = workbook.sheets.find(sheet => sheet.id === outputId)!
    target.cells = output.cells
  }
  Object.assign(current, draft)
  if (chart.ownedPivot) {
    delete current.outputRange; delete current.generatedCellKeys; delete current.drillDownMap; delete current.axisOutput
    chart.revision += 1
  } else {
    for (const linked of workbook.charts ?? []) if (linked.pivotTableId === current.id && !linked.ownedPivot) {
      linked.source = { ...linked.source, sheetId: current.anchor.subUnitId, range: cloneModel(current.outputRange
        ?? { startRow: current.anchor.row, endRow: current.anchor.row, startCol: current.anchor.col, endCol: current.anchor.col }) }
      linked.revision += 1
    }
  }
  return true
}

function editState(report: PivotTableModel): unknown {
  const fields = report.fields.map(field => ({ ...field, selectedItems: field.selectedItems ? [...field.selectedItems].sort() : undefined }))
    .sort((a, b) => areas.indexOf(a.area) - areas.indexOf(b.area) || a.position - b.position || a.id.localeCompare(b.id))
  const collapseInfo = normalizePivotCollapseInfo(report, report.collapseInfo ?? {})
  return { fields, options: report.options, dateGroupingMode: report.dateGroupingMode ?? 'derived',
    collapseInfo: Object.keys(collapseInfo).length ? collapseInfo : undefined }
}

function configuredField(report: PivotTableModel, id: string): PivotFieldModel {
  const field = findConfiguredField(report, id)
  if (!field) throw new PivotChartEditError('invalid-edit', `透视图字段未配置：${id}`)
  return field
}

function findConfiguredField(report: PivotTableModel, id: string): PivotFieldModel | undefined {
  if (!identifier(id)) throw new PivotChartEditError('invalid-edit', '透视图来源字段 ID 无效。')
  const bySource = report.fields.filter(field => sourceFieldId(report, field) === id)
  if (bySource.length === 1) return bySource[0]!
  const exact = report.fields.find(field => field.id === id)
  if (exact && (bySource.length === 0 || bySource.includes(exact))) return exact
  if (bySource.length > 1) throw new PivotChartEditError('invalid-edit', '同一来源存在多个字段实例，请指定报告字段 ID。')
  return undefined
}

function assertOptions(value: Partial<PivotTableOptions>): void {
  if (!record(value) || Object.keys(value).some(key => !options.includes(key))) throw new PivotChartEditError('invalid-edit', '透视图报告选项无效或尚未支持。')
}

function applyFieldEdit(workbook: WorkbookModel, report: PivotTableModel, edit: PivotChartReportEdit): void {
  const mode = 'deferred'
  switch (edit.type) {
    case 'setFieldsConfig': return replaceFields(workbook, report, edit.config)
    case 'resetFields': {
      if (edit.area !== undefined && !areas.includes(edit.area)) throw new PivotChartEditError('invalid-edit', '透视图字段区域无效。')
      removePivotFields(workbook, report, report.fields.filter(field => edit.area === undefined || field.area === edit.area).map(field => field.id), mode)
      return
    }
    case 'addField': {
      const source = getPivotSourceFields(workbook, report)
      if (typeof edit.dataFieldId === 'number' && (!Number.isSafeInteger(edit.dataFieldId) || edit.dataFieldId < 0)) {
        throw new PivotChartEditError('invalid-edit', '透视图来源字段索引无效。')
      }
      const id = typeof edit.dataFieldId === 'number' ? source[edit.dataFieldId]?.fieldId : edit.dataFieldId
      if (!identifier(id)) throw new PivotChartEditError('invalid-edit', '透视图来源字段或索引无效。')
      setPivotField(workbook, report, id, edit.fieldArea, edit.index, mode); return
    }
    case 'moveField': {
      const field = configuredField(report, edit.fieldId)
      if (!Number.isSafeInteger(edit.index) || edit.index < 0) throw new PivotChartEditError('invalid-edit', '透视图字段顺序无效。')
      setPivotField(workbook, report, sourceFieldId(report, field), edit.fieldArea, edit.index, mode, field.id); return
    }
    case 'removeFields': {
      if (!Array.isArray(edit.fieldIds) || edit.fieldIds.some(id => !identifier(id))) throw new PivotChartEditError('invalid-edit', '透视图移除字段 ID 无效。')
      const local = [...new Set(edit.fieldIds.flatMap(id => {
        const source = report.fields.filter(field => sourceFieldId(report, field) === id)
        return source.length ? source.map(field => field.id) : report.fields.filter(field => field.id === id).map(field => field.id)
      }))]
      removePivotFields(workbook, report, local, mode); return
    }
    case 'setOptions': assertOptions(edit.options); setPivotOptions(workbook, report, edit.options, mode); return
    case 'updateValuePosition': {
      setPivotValuesPosition(workbook, report, edit.position, edit.index, mode); return
    }
    case 'setFieldSetting': {
      const field = configuredField(report, edit.fieldId), setting = edit.setting
      if (!record(setting) || Object.keys(setting).some(key => !['displayName', 'format', 'subtotalType', 'showDataAs'].includes(key))
        || setting.displayName !== undefined && typeof setting.displayName !== 'string' || setting.format !== undefined && typeof setting.format !== 'string') {
        throw new PivotChartEditError('invalid-edit', '透视图字段设置无效或尚未支持。')
      }
      setPivotFieldSetting(workbook, report, field.id, setting, mode)
      if (own(setting, 'displayName') && setting.displayName === undefined) delete field.displayName
      if (own(setting, 'format') && setting.format === undefined) delete field.format
      return
    }
    case 'setShowDataAs': setPivotShowDataAs(workbook, report, configuredField(report, edit.fieldId).id, edit.showDataAs, mode); return
    case 'setCollapse': {
      const field = configuredField(report, edit.fieldId)
      if (field.area !== 'row') throw new PivotChartEditError('invalid-edit', '透视图类别折叠需要行字段。')
      setPivotCollapse(workbook, report, field.id, edit.collapse, edit.item, mode); return
    }
    case 'setCategoryLevel': setPivotCategoryLevel(workbook, report, configuredField(report, edit.fieldId).id, mode); return
    case 'stepCategoryLevel': stepPivotCategoryLevel(workbook, report, edit.direction, mode); return
    case 'setSort': setPivotSort(workbook, report, configuredField(report, edit.fieldId).id, edit.direction, mode); return
    case 'setFilter': {
      const field = configuredField(report, edit.fieldId)
      if (edit.items !== null && (!Array.isArray(edit.items) || edit.items.some(item => typeof item !== 'string'))
        || edit.partialItems !== undefined && (!Array.isArray(edit.partialItems) || edit.partialItems.some(item => typeof item !== 'string'))
        || edit.isAll !== undefined && typeof edit.isAll !== 'boolean' || edit.clearAllFilters !== undefined && typeof edit.clearAllFilters !== 'boolean') {
        throw new PivotChartEditError('invalid-edit', '透视图筛选项目无效。')
      }
      if (edit.clearAllFilters) { clearPivotFieldFilters(workbook, report, field.id, mode); return }
      setPivotLabelFilter(workbook, report, field.id, edit.isAll ? null : edit.items, mode,
        { keyMode: 'grouped', partialItems: edit.isAll ? undefined : edit.partialItems }); return
    }
    case 'setValueFilter': {
      const field = configuredField(report, edit.fieldId), filter = edit.filter
      assertValueFilter(filter)
      const mapped = filter ? { ...filter, valueFieldId: configuredField(report, filter.valueFieldId).id } : null
      setPivotValueFilter(workbook, report, field.id, mapped, mode); return
    }
    case 'setDateFilter': setPivotDateFilter(workbook, report, configuredField(report, edit.fieldId).id, edit.filter, mode); return
    case 'setLabelConditionFilter': setPivotLabelConditionFilter(workbook, report, configuredField(report, edit.fieldId).id, edit.filter, mode); return
    case 'setNumberGroup': assertNumberGroup(edit.group); setPivotNumberGroup(workbook, report, configuredField(report, edit.fieldId).id, edit.group, mode); return
    case 'setDateGroupingMode': setPivotDateGroupingMode(workbook, report, edit.mode, mode); return
    case 'setDateGroup': {
      const field = configuredField(report, edit.fieldId)
      if (edit.excelLike !== undefined && typeof edit.excelLike !== 'boolean') throw new PivotChartEditError('invalid-edit', '透视图日期分组模式无效。')
      if (edit.excelLike && edit.dateType !== null) setPivotExcelDateGroup(workbook, report, field.id, edit.dateType, mode)
      else setPivotDateGroupType(workbook, report, field.id, edit.dateType, mode)
      return
    }
    case 'setManualGroups': assertManualGroups(edit.groups); setPivotManualGroups(workbook, report, configuredField(report, edit.fieldId).id, edit.groups, mode); return
    case 'refresh': return
    default: throw new PivotChartEditError('invalid-edit', '透视图字段编辑类型尚未支持。')
  }
}

function replaceFields(workbook: WorkbookModel, report: PivotTableModel, config: PivotChartFieldsConfig): void {
  if (!record(config) || !Array.isArray(config.fields) || Object.keys(config).some(key => !['fields', 'options', 'valuePosition', 'valueIndex', 'dateGroupingMode', 'collapseInfo'].includes(key))) {
    throw new PivotChartEditError('invalid-edit', '透视图完整字段配置无效。')
  }
  const source = getPivotSourceFields(workbook, report), ids = new Set<string>()
  const fields = config.fields.map(field => {
    if (!record(field) || Object.keys(field).some(key => !fieldProperties.includes(key))
      || !identifier(field.id) || ids.has(field.id) || !areas.includes(field.area)
      || field.sourceFieldId !== undefined && !identifier(field.sourceFieldId)
      || !Number.isSafeInteger(field.position) || field.position < 0 || typeof field.name !== 'string') {
      throw new PivotChartEditError('invalid-edit', '透视图字段实例、区域或顺序无效。')
    }
    ids.add(field.id)
    const info = source.find(info => info.fieldId === (field.sourceFieldId ?? source.find(info => info.sourceColumn === field.sourceColumn)?.fieldId))
    if (!info || field.sourceColumn !== info.sourceColumn) throw new PivotChartEditError('invalid-edit', '透视图字段来源与缓存列不一致。')
    return { ...cloneModel(field), sourceFieldId: info.fieldId }
  })
  for (const area of areas) {
    const siblings = fields.filter(field => field.area === area).sort((a, b) => a.position - b.position)
    if (siblings.some((field, index) => field.position !== index)) throw new PivotChartEditError('invalid-edit', '透视图字段区域顺序必须连续。')
  }
  const previousFields = report.fields, previousCollapseInfo = report.collapseInfo
  if (config.valuePosition !== undefined && config.options?.valuePosition !== undefined && config.valuePosition !== config.options.valuePosition
    || config.valueIndex !== undefined && config.options?.valueIndex !== undefined && config.valueIndex !== config.options.valueIndex) {
    throw new PivotChartEditError('invalid-edit', '透视图完整配置的值字段位置有冲突。')
  }
  if (config.options !== undefined) { assertOptions(config.options); setPivotOptions(workbook, report, config.options, 'deferred') }
  if (config.valuePosition !== undefined || config.valueIndex !== undefined) {
    setPivotOptions(workbook, report, {
      ...(config.valuePosition === undefined ? {} : { valuePosition: config.valuePosition }),
      ...(config.valueIndex === undefined ? {} : { valueIndex: config.valueIndex }),
    }, 'deferred')
  }
  report.fields = fields
  delete report.collapseInfo
  for (const field of fields) {
    if (field.displayName !== undefined && typeof field.displayName !== 'string' || field.format !== undefined && typeof field.format !== 'string') {
      throw new PivotChartEditError('invalid-edit', '透视图字段标题或数字格式无效。')
    }
    setPivotFieldSetting(workbook, report, field.id, { displayName: field.displayName, format: field.format,
      ...(field.showDataAs === undefined ? {} : { showDataAs: field.showDataAs }),
      ...(field.aggregation === undefined ? {} : { subtotalType: field.aggregation }) }, 'deferred')
    if (field.sort !== undefined) setPivotSort(workbook, report, field.id, field.sort, 'deferred')
    const groups = [field.numberGroup, field.dateGroupType, field.manualGroups].filter(group => group !== undefined)
    if (groups.length > 1) throw new PivotChartEditError('invalid-edit', '透视图字段不能同时持有多种分组。')
    if (field.numberGroup !== undefined) { assertNumberGroup(field.numberGroup); setPivotNumberGroup(workbook, report, field.id, field.numberGroup, 'deferred') }
    if (field.dateGroupType !== undefined) setPivotDateGroupType(workbook, report, field.id, field.dateGroupType, 'deferred')
    if (field.manualGroups !== undefined) { assertManualGroups(field.manualGroups); setPivotManualGroups(workbook, report, field.id, field.manualGroups, 'deferred') }
    for (const key of ['selectedItems', 'groupedItems', 'partialItems'] as const) if (field[key] !== undefined) {
      const items = field[key]!
      if (!Array.isArray(items) || items.some(item => typeof item !== 'string')) throw new PivotChartEditError('invalid-edit', '透视图完整配置的筛选项目无效。')
      field[key] = [...new Set(items)]
    }
    if (field.recordSelection !== undefined) field.recordSelection = normalizePivotRecordSelection(field.recordSelection)
    if (field.partialItems !== undefined && field.groupedItems === undefined
      || field.partialItems?.length && !field.recordSelection) throw new PivotChartEditError('invalid-edit', '部分透视项目需要分组键和原始成员选择。')
    if (field.groupedItems !== undefined) {
      const keys = getPivotGroupedFieldItems(workbook, report, field.id)
      if ([...field.groupedItems, ...(field.partialItems ?? [])].some(key => !keys.includes(key))
        || field.partialItems?.some(key => field.groupedItems!.includes(key))) throw new PivotChartEditError('invalid-edit', '透视图完整配置的分组筛选项目无效。')
      field.groupedItems = keys.filter(key => field.groupedItems!.includes(key))
      if (field.partialItems?.length) field.partialItems = keys.filter(key => field.partialItems!.includes(key))
      else delete field.partialItems
      field.selectedItems = getPivotSourceSelectedItems(workbook, report, field.id)
    }
    if (field.dateFilter !== undefined && field.labelFilter !== undefined) throw new PivotChartEditError('invalid-edit', '一个字段不能同时持有日期和标签条件。')
    if (field.dateFilter !== undefined) {
      field.dateFilter = normalizePivotDateFilter(field.dateFilter)
      if (!getPivotFieldFilterCapabilities(workbook, report, field.id).supportsDateFilter) throw new PivotChartEditError('invalid-edit', '日期筛选需要真实日期来源。')
    }
    if (field.labelFilter !== undefined) field.labelFilter = normalizePivotLabelConditionFilter(field.labelFilter)
    if (field.valueFilter !== undefined) {
      const normalized = normalizePivotValueFilter(field.valueFilter), valueField = configuredField(report, normalized.valueFieldId)
      if (valueField.area !== 'value') throw new PivotChartEditError('invalid-edit', '值筛选需要值字段实例。')
      field.valueFilter = { ...normalized, valueFieldId: valueField.id }
    }
    if (field.area === 'value' && (getPivotFieldFilterCategories(field).size || field.selectedItems !== undefined
      || field.recordSelection !== undefined)) throw new PivotChartEditError('invalid-edit', '值字段不能持有标签、日期或成员筛选。')
  }
  // Explicitly disabling multiple categories has the same whole-report clear behavior as setOptions.
  if (config.options?.allowMultipleFilters === false) setPivotOptions(workbook, report, { allowMultipleFilters: false }, 'deferred')
  if (own(config, 'dateGroupingMode')) setPivotDateGroupingMode(workbook, report, config.dateGroupingMode === undefined ? 'derived' : config.dateGroupingMode, 'deferred')
  setPivotCollapseInfo(workbook, report, own(config, 'collapseInfo') ? config.collapseInfo ?? null
    : retainPivotCollapseInfo(report, previousCollapseInfo, previousFields), 'deferred')
}

function assertValueFilter(value: PivotFieldFilter | null): void {
  if (value !== null && (!record(value) || Object.keys(value).some(key => !['operator', 'expected', 'expected2', 'valueFieldId'].includes(key)))) {
    throw new PivotChartEditError('invalid-edit', '透视图值筛选无效或尚未支持。')
  }
}

function assertNumberGroup(value: PivotNumberGroup | null): void {
  if (value !== null && (!record(value) || Object.keys(value).some(key => !['start', 'end', 'interval'].includes(key)))) {
    throw new PivotChartEditError('invalid-edit', '透视图数字分组无效或尚未支持。')
  }
}

function assertManualGroups(value: PivotManualGroup[]): void {
  if (!Array.isArray(value) || value.some(group => !record(group) || Object.keys(group).some(key => !['name', 'members'].includes(key)))) {
    throw new PivotChartEditError('invalid-edit', '透视图手工分组无效或尚未支持。')
  }
}

function assertConnectedFieldsSurvive(workbook: WorkbookModel, chart: ChartModel, current: PivotTableModel, next: PivotTableModel): void {
  const protectedIds = new Set<string>()
  const chartTargetsCurrent = (id: string): boolean => chart.ownedPivot ? id === chart.id
    : Boolean(workbook.charts?.some(candidate => candidate.id === id && candidate.pivotTableId === current.id && !candidate.ownedPivot))
  // Retained connections still protect fields when a sibling report is stale or a view is temporarily absent.
  for (const cache of workbook.slicerCaches ?? []) {
    if (cache.source.kind !== 'pivot') continue
    if (!chart.ownedPivot && cache.connections?.pivotTables.some(ref => ref.pivotTableId === current.id)
      || cache.connections?.ownedPivotCharts?.some(ref => chartTargetsCurrent(ref.pivotChartId))) protectedIds.add(cache.source.sourceFieldId)
  }
  for (const view of workbook.slicers ?? []) {
    let reports: ReturnType<typeof getSlicerConnectionReports>
    try { reports = getSlicerConnectionReports(workbook, view) }
    catch (error) {
      if (!(error instanceof SlicerConnectionError)) throw error
      if (!chart.ownedPivot && view.pivotTableId === current.id || view.ownedPivotChartId && chartTargetsCurrent(view.ownedPivotChartId)) {
        const field = current.fields.find(field => field.area !== 'value' && (sourceFieldId(current, field) === view.sourceFieldId || field.id === view.sourceFieldId))
        if (field) protectedIds.add(sourceFieldId(current, field))
      }
      continue
    }
    for (const report of reports) {
      if (chart.ownedPivot ? report.kind !== 'ownedPivotChart' || report.id !== chart.id : report.kind !== 'pivotTable' || report.id !== current.id) continue
      protectedIds.add(sourceFieldId(current, report.field))
    }
  }
  for (const id of protectedIds) if (current.fields.some(field => field.area !== 'value' && sourceFieldId(current, field) === id)
    && !next.fields.some(field => field.area !== 'value' && sourceFieldId(next, field) === id)) {
    throw new PivotChartEditError('field-connected', '该字段仍用于切片器或时间线连接，请先调整报告连接。')
  }
}
