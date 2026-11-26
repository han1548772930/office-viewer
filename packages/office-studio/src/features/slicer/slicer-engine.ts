import type {
  CellScalar,
  PivotRecordSelection,
  SlicerModel,
  SlicerCacheModel,
  SheetSlicerConnections,
  SheetFilterCriteria,
  SlicerTimelineRange,
  SlicerType,
  TimelineLevel,
  WorkbookModel,
} from '../../domain/model'
import { getPivotSourceFields, getPivotSourceRecordSelection } from '../pivot/services/pivot-engine'
import { compileSheetCellFilterCriteria, setAutoFilter, setFilterCriteria } from '../sheet-sort-filter/sort-filter-engine'
import { parseTimelineDate } from './timeline-date'
import { cloneModel } from '../../domain/serialization.ts'
import { normalizeDrawingAffineGeometry } from '../../domain/drawing-affine.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../rendering/spreadsheet/layout.ts'
import { sheetClipboardObjectAxes } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'
import { captureSlicerDrawing, resolveSlicerAnchorBounds, slicerBoundsInWorkbook, slicerGeometryWithBounds,
  slicerWithDrawingGeometry, validateSlicerDrawing, type SlicerAnchorType, type SlicerArrange } from './slicer-drawing'
import { arrangeSlicer } from './slicer-arrangement.ts'
import { ensureSlicerCache, findSlicerReportField, getSlicerCache, getSlicerConnectionReports, getSlicerSourceValues, getSlicerSourceRecordValues, getSlicerTableNativeFilter,
  planSlicerConnections, resolvePivotCacheId, sourceFieldId, SlicerConnectionError } from './slicer-connections.ts'
import type { SlicerConnectionReport } from './slicer-connections.ts'
import { applySlicerReportFilter, sameSlicerSelection } from './slicer-report-filter.ts'
import { normalizeSlicerRecordSelection, planSlicerRecordItemStates, sameSlicerRecordSelection, slicerRecordSelectionFromItems,
  slicerItemStateFromPredicate, slicerRecordSelectionFromNativeValues, slicerRecordSelectionItemState, slicerRecordSelectionItems } from './slicer-record-selection.ts'
import { assertSlicerSettingsAvailable, normalizeSlicerSettingsPatch, readSlicerSettings, resolveSlicerViewName, type SheetSlicerSettings } from './slicer-settings.ts'
import { getSlicerDisplayMembers } from './slicer-member-query.ts'
import { assertTimelineSettingsAvailable, normalizeTimelineSettingsPatch, readTimelineSettings } from './timeline-settings.ts'
import type { TimelineDateFilterState, TimelinePublicLevel } from '../../domain/timeline-filters.ts'
import { applyTimelineReportFilter, readTimelineFilterState, setTimelineDateCondition, setTimelineDateRange, updateTimelineCacheState } from './timeline-filter.ts'
import { readTimelineReportFilterState } from './timeline-state.ts'
import { timelineFilterStateFromLegacy } from './timeline-periods.ts'

export interface SlicerCreateInput {
  id: string
  unitId: string
  hostSheetId: string
  cacheId?: string
  pivotTableId?: string
  ownedPivotChartId?: string
  tableId?: string
  sourceFieldId: string
  type: SlicerType
  caption?: string
  name?: string
  style?: string
  columnCount?: number
  rowHeight?: number
  showCaption?: boolean
  lockedPosition?: boolean
  level?: TimelineLevel
  selectionLevel?: TimelineLevel
  showHeader?: boolean
  showSelectionLabel?: boolean
  showTimeLevel?: boolean
  showHorizontalScrollbar?: boolean
  scrollPosition?: number
  startItem?: number
  position?: { row: number; column: number }
  size?: { rows: number; columns: number }
  layout?: { left: number; top: number; width: number; height: number }
  drawing?: SlicerModel['drawing']
  drawingGeometry?: SlicerModel['drawingGeometry']
}

export function createSlicer(workbook: WorkbookModel, input: SlicerCreateInput): SlicerModel {
  if (input.unitId !== workbook.id) throw new Error('Slicer must belong to the target workbook')
  if (!workbook.sheets.some((sheet) => sheet.id === input.hostSheetId)) throw new Error(`Slicer host worksheet not found: ${input.hostSheetId}`)
  if (input.type !== 'slicer' && input.type !== 'timeline') throw new TypeError(`Invalid slicer type: ${input.type}`)
  const targetCount = [input.pivotTableId, input.ownedPivotChartId, input.tableId].filter(Boolean).length
  const retainedCache = input.cacheId ? workbook.slicerCaches?.find(cache => cache.id === input.cacheId && cache.unitId === workbook.id && cache.type === input.type) : undefined
  const retainedSource = retainedCache?.source
  if (targetCount !== 1 && !(targetCount === 0 && retainedCache)) throw new Error('Slicer requires one source report, Table, or retained source cache')
  const ownedChart = input.ownedPivotChartId ? workbook.charts?.find(chart => chart.id === input.ownedPivotChartId && chart.ownedPivot && !chart.pivotTableId) : undefined
  const pivot = input.pivotTableId ? workbook.pivotTables?.find((candidate) => candidate.id === input.pivotTableId) : ownedChart?.ownedPivot
  const table = input.tableId ? workbook.tables?.find((candidate) => candidate.id === input.tableId)
    : retainedSource?.kind === 'table' ? workbook.tables?.find(candidate => candidate.id === retainedSource.tableId) : undefined
  if (input.pivotTableId && !pivot) throw new Error(`Pivot table not found: ${input.pivotTableId}`)
  if (input.ownedPivotChartId && !ownedChart) throw new Error(`Owned Pivot chart not found: ${input.ownedPivotChartId}`)
  if (input.tableId && !table) throw new Error(`Table not found: ${input.tableId}`)
  if (pivot && pivot.source.unitId !== workbook.id) throw new Error('Pivot source must belong to the target workbook')
  if (table && input.type === 'timeline') throw new Error('Timeline requires a Pivot source')
  const reportField = pivot ? findSlicerReportField(pivot, String(input.sourceFieldId))
    ?? pivot.fields.find(field => field.id === String(input.sourceFieldId) && field.area !== 'value') : undefined
  const canonicalFieldId = pivot && reportField ? sourceFieldId(pivot, reportField) : String(input.sourceFieldId)
  const pivotField = pivot ? getPivotSourceFields(workbook, pivot).find(field => field.sourceColumn === reportField?.sourceColumn) : undefined
  const tableField = table?.columns.find((field) => field.id === String(input.sourceFieldId))
  const retainedField = retainedSource?.kind === 'pivot'
    ? workbook.pivotCaches?.find(cache => cache.id === retainedSource.pivotCacheId)?.fields?.find(field => field.id === canonicalFieldId) : undefined
  const sourceFieldName = pivotField?.fieldName ?? tableField?.name ?? retainedField?.name ?? (retainedField ? canonicalFieldId : undefined)
  if (!sourceFieldName) throw new Error(`Slicer source field not found: ${input.sourceFieldId}`)
  if (pivot && !reportField) {
    throw new Error(`Pivot source field must be placed in a label area: ${input.sourceFieldId}`)
  }
  const position = input.position ?? { row: 0, column: 0 }
  const size = input.size ?? { rows: 8, columns: 3 }
  validateNonNegative(position.row, 'slicer row')
  validateNonNegative(position.column, 'slicer column')
  validatePositive(size.rows, 'slicer rows')
  validatePositive(size.columns, 'slicer columns')
  const model: SlicerModel = {
    id: input.id,
    unitId: workbook.id!,
    hostSheetId: input.hostSheetId,
    ...(input.cacheId ? { cacheId: input.cacheId } : {}),
    ...(ownedChart ? { ownedPivotChartId: ownedChart.id } : pivot ? { pivotTableId: pivot.id } : table ? { tableId: table.id } : {}),
    sourceFieldId: canonicalFieldId,
    type: input.type,
    caption: input.caption === undefined ? sourceFieldName : input.caption,
    name: resolveSlicerViewName(workbook, input.name === undefined ? sourceFieldName : input.name, input.name !== undefined),
    position: { ...position },
    size: { ...size },
    layout: input.layout ? { ...input.layout } : { left: 200, top: 200, width: input.type === 'timeline' ? 420 : 220, height: input.type === 'timeline' ? 140 : 280 },
    drawing: input.drawing ? structuredClone(input.drawing) : undefined,
    ...(input.drawingGeometry !== undefined ? { drawingGeometry: normalizeDrawingAffineGeometry(input.drawingGeometry) } : {}),
    ...(input.style !== undefined ? { style: input.style } : {}),
    startItem: input.startItem ?? 0,
    columnCount: input.columnCount ?? 1,
    rowHeight: input.rowHeight ?? 24,
    showCaption: input.showCaption ?? true,
    lockedPosition: input.lockedPosition ?? false,
    ...(input.type === 'timeline' ? {
      level: input.level ?? 'month',
      selectionLevel: input.selectionLevel ?? 'month',
      showHeader: input.showHeader ?? true,
      showSelectionLabel: input.showSelectionLabel ?? true,
      showTimeLevel: input.showTimeLevel ?? true,
      showHorizontalScrollbar: input.showHorizontalScrollbar ?? true,
      scrollPosition: input.scrollPosition ?? 0,
    } : {}),
    revision: 0,
  }
  if (model.type === 'slicer') {
    Object.assign(model, normalizeSlicerSettingsPatch(readSlicerSettings(model)))
    assertSlicerSettingsAvailable(workbook, model, readSlicerSettings(model))
  } else {
    const settings = normalizeTimelineSettingsPatch(readTimelineSettings(model))
    Object.assign(model, settings); assertTimelineSettingsAvailable(workbook, model, settings)
  }
  validateSlicerPresentation(model)
  if (model.drawingGeometry) Object.assign(model, slicerWithDrawingGeometry(workbook, model, model.drawingGeometry))
  model.drawing ??= captureSlicerDrawing(workbook, model, model.layout!, 'position')
  if (input.cacheId) {
    const cache = getSlicerCache(workbook, model)
    if (pivot ? cache.source.kind !== 'pivot' || resolvePivotCacheId(workbook, pivot) !== cache.source.pivotCacheId
      : table ? cache.source.kind !== 'table' || cache.source.tableId !== table.id : targetCount !== 0) {
      throw new SlicerConnectionError('invalid-target', '新切片器与指定共享缓存必须使用同一来源和源字段。')
    }
  }
  if (model.type === 'timeline' && !getSlicerSourceRecordValues(workbook, model).some(value => value.type === 'date')) throw new SlicerConnectionError('invalid-target', '时间线需要真实日期源字段。')
  if (workbook.slicers?.some((candidate) => candidate.id === model.id)) throw new Error(`Slicer already exists: ${model.id}`)
  const cache = ensureSlicerCache(workbook, model)
  model.selectedItems = cache.selectedItems ? [...cache.selectedItems] : undefined
  model.timelineRange = cache.timelineRange ? cloneModel(cache.timelineRange) : undefined
  model.timelineDateRange = cache.timelineDateRange ? cloneModel(cache.timelineDateRange) : undefined
  model.timelineDateCondition = cache.timelineDateCondition ? cloneModel(cache.timelineDateCondition) : undefined
  workbook.slicers ??= []
  workbook.slicers.push(model)
  return model
}

export type SlicerPresentationPatch = Partial<Pick<SlicerModel,
  | 'name' | 'caption' | 'style' | 'startItem' | 'columnCount' | 'rowHeight' | 'showCaption' | 'lockedPosition'
  | 'level' | 'selectionLevel' | 'showHeader' | 'showSelectionLabel' | 'showTimeLevel' | 'showHorizontalScrollbar' | 'scrollPosition'
>> & Partial<Pick<SlicerModel, 'position' | 'size' | 'layout' | 'drawing' | 'drawingGeometry'>> & { anchorType?: SlicerAnchorType; keepRatio?: boolean; arrange?: SlicerArrange }
export type SlicerPresentationInputPatch = Omit<SlicerPresentationPatch, 'level' | 'selectionLevel'>
  & { level?: TimelineLevel | TimelinePublicLevel; selectionLevel?: TimelineLevel | TimelinePublicLevel }

export function patchSlicerSettings(workbook: WorkbookModel, slicer: SlicerModel, input: Partial<SheetSlicerSettings>): void {
  const patch = normalizeSlicerSettingsPatch(input)
  assertSlicerSettingsAvailable(workbook, slicer, patch)
  updateSlicerPresentation(slicer, patch, workbook)
}

export function updateSlicerPresentation(slicer: SlicerModel, input: SlicerPresentationInputPatch, workbook?: WorkbookModel): void {
  const { level: _level, selectionLevel: _selectionLevel, ...propertiesWithoutGrains } = input
  const grains = normalizeTimelineSettingsPatch(Object.fromEntries(['level', 'selectionLevel']
    .filter(key => Object.hasOwn(input, key)).map(key => [key, input[key as 'level' | 'selectionLevel']])))
  let patch: SlicerPresentationPatch = { ...propertiesWithoutGrains, ...grains }
  if (slicer.type === 'slicer') {
    const settings = normalizeSlicerSettingsPatch(Object.fromEntries(
      ['name', 'rowHeight', 'caption', 'startItem', 'columnCount', 'showCaption', 'lockedPosition', 'style']
        .filter(key => Object.hasOwn(patch, key)).map(key => [key, patch[key as keyof SlicerPresentationPatch]])))
    if (workbook) assertSlicerSettingsAvailable(workbook, slicer, settings)
    patch = { ...patch, ...settings }
  } else {
    const settings = normalizeTimelineSettingsPatch(Object.fromEntries(
      ['name', 'caption', 'style', 'level', 'selectionLevel', 'scrollPosition', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar']
        .filter(key => Object.hasOwn(patch, key)).map(key => [key, patch[key as keyof SlicerPresentationPatch]])))
    if (workbook) assertTimelineSettingsAvailable(workbook, slicer, settings)
    patch = { ...patch, ...settings }
  }
  validateSlicerPresentation(patch)
  const allowed = new Set([
    'name', 'caption', 'style', 'startItem', 'columnCount', 'rowHeight', 'showCaption', 'lockedPosition',
    'level', 'selectionLevel', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar', 'scrollPosition',
    'position', 'size', 'layout', 'drawing', 'drawingGeometry', 'anchorType', 'keepRatio', 'arrange',
  ])
  if (Object.keys(patch).some((key) => !allowed.has(key))) throw new TypeError('Unknown slicer presentation property')
  if (patch.drawingGeometry && (patch.layout || patch.position || patch.size || patch.drawing)) {
    throw new TypeError('切片器仿射几何和旧位置字段不能在同一更新中混用。')
  }
  if (slicer.lockedPosition && (patch.layout || patch.position || patch.size || patch.drawingGeometry
    || patch.drawing && !equivalentModelData(slicer.drawing, patch.drawing))) throw new Error('切片器位置已锁定。')
  const { anchorType, keepRatio, arrange, ...properties } = patch
  if ((anchorType !== undefined || arrange !== undefined || patch.drawingGeometry !== undefined
    || slicer.drawingGeometry && (patch.layout || patch.drawing || patch.position || patch.size)
    || (patch.layout !== undefined || patch.position !== undefined || patch.size !== undefined) && slicer.drawing) && !workbook) throw new Error('Slicer drawing changes require a workbook')
  let bounds = workbook ? patch.layout ?? slicerBoundsInWorkbook(workbook, slicer) : slicer.layout
  let next: SlicerModel = { ...cloneModel(slicer),
    ...properties,
    ...(patch.position ? { position: { ...patch.position } } : {}),
    ...(patch.size ? { size: { ...patch.size } } : {}),
    ...(patch.layout ? { layout: { ...patch.layout } } : {}),
    ...(patch.drawing ? { drawing: structuredClone(patch.drawing) } : {}),
  }
  if (workbook && !patch.layout && !patch.drawingGeometry && (patch.position || patch.size)) {
    const sheet = workbook.sheets.find(sheet => sheet.id === slicer.hostSheetId)
    if (!sheet) throw new Error('切片器所在工作表已移除。')
    const axes = sheetClipboardObjectAxes(workbook, sheet), prior = slicerBoundsInWorkbook(workbook, slicer)
    bounds = {
      left: patch.position ? GRID_HEADER_WIDTH + axes.columns.offset(next.position.column) : prior.left,
      top: patch.position ? GRID_HEADER_HEIGHT + axes.rows.offset(next.position.row) : prior.top,
      width: patch.size ? axes.columns.offset(next.position.column + next.size.columns) - axes.columns.offset(next.position.column) : prior.width,
      height: patch.size ? axes.rows.offset(next.position.row + next.size.rows) - axes.rows.offset(next.position.row) : prior.height,
    }
    next.layout = { ...bounds }
    if (slicer.drawingGeometry) next.drawingGeometry = slicerGeometryWithBounds(slicer.drawingGeometry, bounds)
    next.drawing = captureSlicerDrawing(workbook, next, bounds, anchorType ?? next.drawing?.anchorType ?? 'position')
  }
  if (workbook && patch.drawingGeometry) {
    next = slicerWithDrawingGeometry(workbook, next, patch.drawingGeometry)
    bounds = next.layout
  }
  else if (workbook && slicer.drawingGeometry && (patch.layout || patch.drawing)) {
    const sheet = workbook.sheets.find(sheet => sheet.id === slicer.hostSheetId)
    if (!sheet) throw new Error('切片器所在工作表已移除。')
    const target = patch.layout ?? resolveSlicerAnchorBounds(next, sheetClipboardObjectAxes(workbook, sheet))
    next.drawingGeometry = slicerGeometryWithBounds(slicer.drawingGeometry, target)
    next.layout = { ...target }
    bounds = target
  }
  if (workbook && bounds && (anchorType !== undefined || patch.layout !== undefined)) {
    next.layout = { ...bounds }
    next.drawing = captureSlicerDrawing(workbook, next, bounds, anchorType ?? next.drawing?.anchorType ?? 'position')
  }
  if (keepRatio !== undefined && next.drawing) next.drawing.keepRatio = keepRatio
  validateSlicerPresentation(next)
  const arrangedWorkbook = arrange && workbook ? { ...workbook, sheets: workbook.sheets.map(sheet => ({ ...sheet })),
    slicers: workbook.slicers?.map(control => control.id === slicer.id ? next : control) } : undefined
  const arranged = arrangedWorkbook && arrange ? arrangeSlicer(arrangedWorkbook, next, arrange) : false
  if (equivalentModelData(slicer, next) && !arranged) return
  Object.assign(slicer, next)
  if (arranged && workbook && arrangedWorkbook) {
    workbook.sheets.find(sheet => sheet.id === slicer.hostSheetId)!.drawingState = arrangedWorkbook.sheets.find(sheet => sheet.id === slicer.hostSheetId)!.drawingState
  }
  slicer.revision += 1
}

function validateSlicerPresentation(patch: SlicerPresentationPatch): void {
  if (patch.drawingGeometry !== undefined) normalizeDrawingAffineGeometry(patch.drawingGeometry)
  if (patch.drawing !== undefined) validateSlicerDrawing(patch.drawing)
  if (patch.anchorType !== undefined && !['both', 'position', 'none'].includes(patch.anchorType)) throw new TypeError('Invalid slicer anchor type')
  if (patch.keepRatio !== undefined && typeof patch.keepRatio !== 'boolean') throw new TypeError('Slicer aspect ratio lock must be a boolean')
  if (patch.arrange !== undefined && !['forward', 'backward', 'front', 'back'].includes(patch.arrange)) throw new TypeError('Invalid slicer arrangement')
  if (patch.name !== undefined && (typeof patch.name !== 'string' || !patch.name.trim())) throw new TypeError('Slicer name must not be empty')
  if (patch.caption !== undefined && (typeof patch.caption !== 'string' || !patch.caption.trim() || patch.caption.length > 32767)) throw new TypeError('Slicer caption must contain between 1 and 32767 characters')
  if (patch.style !== undefined && (typeof patch.style !== 'string' || !patch.style.trim() || patch.style.length > 255)) throw new TypeError('Slicer style must contain between 1 and 255 characters')
  if (patch.columnCount !== undefined && (!Number.isInteger(patch.columnCount) || patch.columnCount < 1 || patch.columnCount > 20000)) {
    throw new RangeError('Slicer column count must be an integer between 1 and 20000')
  }
  // This model stores CSS pixels. Upstream resources store button height as unsigned EMU.
  if (patch.rowHeight !== undefined && (!Number.isFinite(patch.rowHeight) || patch.rowHeight <= 0 || patch.rowHeight > 4294967295 / 9525)) throw new RangeError('Slicer row height is outside the supported range')
  if (patch.startItem !== undefined && (!Number.isInteger(patch.startItem) || patch.startItem < 0 || patch.startItem > 4294967295)) throw new RangeError('Slicer start item must be an unsigned 32-bit integer')
  if (patch.scrollPosition !== undefined && !Number.isFinite(patch.scrollPosition)) throw new RangeError('Timeline scroll position must be finite')
  if (patch.position !== undefined) {
    validateNonNegative(patch.position.row, 'slicer row')
    validateNonNegative(patch.position.column, 'slicer column')
  }
  if (patch.size !== undefined) {
    validatePositive(patch.size.rows, 'slicer rows')
    validatePositive(patch.size.columns, 'slicer columns')
  }
  if (patch.layout !== undefined) {
    for (const key of ['left', 'top'] as const) if (!Number.isFinite(patch.layout[key])) throw new RangeError(`Slicer ${key} must be finite`)
    for (const key of ['width', 'height'] as const) if (!Number.isFinite(patch.layout[key]) || patch.layout[key] <= 0) throw new RangeError(`Slicer ${key} must be positive and finite`)
  }
  if (patch.level !== undefined && !['day', 'month', 'quarter', 'year'].includes(patch.level)) throw new TypeError('Invalid timeline level')
  if (patch.selectionLevel !== undefined && !['day', 'month', 'quarter', 'year'].includes(patch.selectionLevel)) throw new TypeError('Invalid timeline selection level')
  for (const key of ['showCaption', 'lockedPosition', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar'] as const) {
    if (patch[key] !== undefined && typeof patch[key] !== 'boolean') throw new TypeError(`Slicer ${key} must be a boolean`)
  }
}

export function setSlicerItems(workbook: WorkbookModel, slicer: SlicerModel, selectedItems: string[] | null): void {
  if (slicer.type !== 'slicer') throw new Error('Timeline controls require a date range')
  const items = selectedItems == null ? null : normalizeItems(selectedItems)
  if (items && items.some((item) => !getSlicerItems(workbook, slicer).includes(item))) throw new Error('Slicer selection contains an unknown item')
  const selection = slicerRecordSelectionFromItems(getSlicerSourceValues(workbook, slicer), items)
  setSlicerRecordSelection(workbook, slicer, selection ?? null)
}

export function setSlicerRecordSelection(workbook: WorkbookModel, slicer: SlicerModel, recordSelection: PivotRecordSelection | null): void {
  if (slicer.type !== 'slicer') throw new Error('Timeline controls require a date range')
  const selection = normalizeSlicerRecordSelection(recordSelection), cache = getSlicerCache(workbook, slicer)
  const items = slicerRecordSelectionItems(selection, getSlicerSourceValues(workbook, slicer), getSlicerSourceRecordValues(workbook, slicer))
  if (cache.source.kind === 'table') setTableSlicerRecordFilter(workbook, slicer, selection ?? null)
  else applySlicerReportFilter(workbook, getSlicerConnectionReports(workbook, slicer), selection ?? null)
  updateSharedSelection(workbook, slicer, items ?? null, undefined, selection)
}

export function setSlicerItemStates(workbook: WorkbookModel, slicer: SlicerModel, checkedItems: string[], partialItems: string[]): void {
  if (slicer.type !== 'slicer') throw new Error('Timeline controls require a date range')
  const values = getSlicerSourceValues(workbook, slicer), typedValues = getSlicerSourceRecordValues(workbook, slicer),
    keys = new Set(values.map(value => String(value ?? '')))
  if (!Array.isArray(checkedItems) || !Array.isArray(partialItems) || [...checkedItems, ...partialItems].some(item => typeof item !== 'string')) {
    throw new TypeError('Slicer item states require string keys')
  }
  const checked = new Set(checkedItems)
  if ([...checked, ...partialItems].some(key => !keys.has(key))) throw new TypeError('Slicer item state contains an unknown item')
  if (partialItems.some(key => checked.has(key))) throw new TypeError('A slicer item cannot be both checked and partial')
  // Canvas checked-all is implicit all, including the empty list; direct raw include remains explicit.
  if (partialItems.length === 0 && checked.size === keys.size) { setSlicerItems(workbook, slicer, null); return }
  let current: PivotRecordSelection | undefined
  if (partialItems.length) {
    const cache = getSlicerCache(workbook, slicer), filter = cache.source.kind === 'table' ? getSlicerTableNativeFilter(workbook, slicer) : undefined
    if (filter !== undefined && filter.kind !== 'values') {
      const matches = tableNativeCellPredicate(workbook, slicer, filter)
      current = { mode: 'include', values: typedValues.filter((_value, index) => matches(values[index]!, index)) }
    } else current = readSlicerRecordSelection(workbook, slicer)
  }
  const selection = planSlicerRecordItemStates(current, values, checkedItems, partialItems, typedValues)
  setSlicerRecordSelection(workbook, slicer, selection)
}

export function setSlicerTimelineRange(workbook: WorkbookModel, slicer: SlicerModel, range: SlicerTimelineRange | null): void {
  if (slicer.type !== 'timeline') throw new Error('Slicer controls require selected items')
  ensureSlicerCache(workbook, slicer)
  const state: TimelineDateFilterState = range == null ? {} : timelineFilterStateFromLegacy(range, workbook.date1904)
  if (state.condition !== undefined) setTimelineDateCondition(workbook, slicer, state.condition)
  else setTimelineDateRange(workbook, slicer, state.range)
}

/** Replacement is shared by every view of this cache. Detached reports retain
 * their last field filter and can subsequently be edited independently. */
export function setSlicerConnections(workbook: WorkbookModel, slicer: SlicerModel, connections: SheetSlicerConnections): boolean {
  const plan = planSlicerConnections(workbook, slicer, connections), cache = ensureSlicerCache(workbook, slicer)
  const filter = cache.recordSelection !== undefined ? normalizeSlicerRecordSelection(cache.recordSelection) ?? null
    : cache.selectedItems ? [...cache.selectedItems] : null
  const legacy = cache.timelineLegacyMemberFilterPending === true || cache.recordSelection !== undefined || cache.selectedItems !== undefined
  const state = cache.type === 'timeline' ? readTimelineFilterState(cache) : undefined
  const canonicalTimeline = state !== undefined && (state.range !== undefined || state.condition !== undefined || !legacy || cache.timelineLegacyMemberFilterPending === true)
  const filtered = state !== undefined && canonicalTimeline ? applyTimelineReportFilter(workbook, plan.reports, state, legacy)
    : applySlicerReportFilter(workbook, plan.reports, filter)
  const retained = ensureSlicerCache(workbook, slicer)
  if (!plan.changed && !filtered && retained.timelineLegacyMemberFilterPending !== true) return true
  retained.connections = cloneModel(plan.connections)
  if (canonicalTimeline) delete retained.timelineLegacyMemberFilterPending
  retained.revision += 1
  for (const control of sharedControls(workbook, slicer)) {
    control.selectedItems = retained.selectedItems ? [...retained.selectedItems] : undefined
    control.timelineRange = retained.timelineRange ? cloneModel(retained.timelineRange) : undefined
    control.timelineDateRange = retained.timelineDateRange ? cloneModel(retained.timelineDateRange) : undefined
    control.timelineDateCondition = retained.timelineDateCondition ? cloneModel(retained.timelineDateCondition) : undefined
    control.revision += 1
  }
  return true
}

export function clearSlicerFilter(workbook: WorkbookModel, slicer: SlicerModel): void {
  if (slicer.type === 'timeline') { ensureSlicerCache(workbook, slicer); setTimelineDateRange(workbook, slicer, undefined) }
  else setSlicerItems(workbook, slicer, null)
}

export function removeSlicer(workbook: WorkbookModel, slicerId: string): void {
  const slicer = workbook.slicers?.find((candidate) => candidate.id === slicerId)
  if (!slicer) throw new Error(`Slicer not found: ${slicerId}`)
  workbook.slicers = (workbook.slicers ?? []).filter((candidate) => candidate.id !== slicerId)
  if (slicer.cacheId && !workbook.slicers.some(control => control.cacheId === slicer.cacheId)) {
    workbook.slicerCaches = workbook.slicerCaches?.filter(cache => cache.id !== slicer.cacheId)
  }
}

export function getSlicer(workbook: WorkbookModel, slicerId: string): SlicerModel | undefined {
  const model = workbook.slicers?.find((candidate) => candidate.id === slicerId)
  if (!model) return undefined
  const cache = getSlicerCache(workbook, model)
  const selectedItems = model.type === 'timeline' ? cache.selectedItems : readSlicerSelection(workbook, model)
  return { ...model, sourceFieldId: cache.source.kind === 'pivot' ? cache.source.sourceFieldId : cache.source.columnId,
    selectedItems, timelineRange: cache.timelineRange ? cloneModel(cache.timelineRange) : undefined,
    timelineDateRange: cache.timelineDateRange ? cloneModel(cache.timelineDateRange) : undefined,
    timelineDateCondition: cache.timelineDateCondition ? cloneModel(cache.timelineDateCondition) : undefined }
}

/** Source filters also change through Pivot and Table commands outside Slicer views. */
export function synchronizeSlicerSelections(workbook: WorkbookModel, before?: WorkbookModel): void {
  const visited = new Set<string>()
  for (const control of workbook.slicers ?? []) {
    let cache: SlicerCacheModel
    try { cache = getSlicerCache(workbook, control) } catch (error) { if (error instanceof SlicerConnectionError) continue; throw error }
    if (visited.has(cache.id)) continue
    visited.add(cache.id)
    if (cache.type === 'timeline') { synchronizeTimelineSelections(workbook, control, cache, before); continue }
    if (cache.source.kind === 'table') {
      const filter = getSlicerTableNativeFilter(workbook, control)
      // Condition/color/manual native restrictions retain their configuration and never become all.
      if (filter !== undefined && filter.kind !== 'values') continue
      const selected = filter ? slicerRecordSelectionFromNativeValues(filter) : undefined
      updateSharedSelection(workbook, control, slicerRecordSelectionItems(selected, getSlicerSourceValues(workbook, control), getSlicerSourceRecordValues(workbook, control)) ?? null, undefined, selected)
      continue
    }
    let reports: SlicerConnectionReport[]
    try { reports = getSlicerConnectionReports(workbook, control) } catch (error) { if (error instanceof SlicerConnectionError) continue; throw error }
    const changed = reports.filter(report => {
      const current = getPivotSourceRecordSelection(workbook, report.pivot, report.field.id)
      if (!before) {
        const cached = cache.recordSelection !== undefined ? cache.recordSelection
          : slicerRecordSelectionFromItems(getSlicerSourceValues(workbook, control), cache.selectedItems ?? null)
        return !sameSlicerRecordSelection(current, cached)
      }
      const previous = report.kind === 'pivotTable' ? before.pivotTables?.find(pivot => pivot.id === report.id)
        : before.charts?.find(chart => chart.id === report.id)?.ownedPivot
      const oldField = previous && findSlicerReportField(previous, cache.source.kind === 'pivot' ? cache.source.sourceFieldId : '')
      return oldField !== undefined && !sameSlicerRecordSelection(getPivotSourceRecordSelection(before, previous!, oldField.id), current)
    })
    if (!changed.length) {
      if (cache.recordSelection !== undefined) {
        const selected = normalizeSlicerRecordSelection(cache.recordSelection)
        applySlicerReportFilter(workbook, reports, selected ?? null)
        updateSharedSelection(workbook, control, slicerRecordSelectionItems(selected, getSlicerSourceValues(workbook, control), getSlicerSourceRecordValues(workbook, control)) ?? null,
          cache.timelineRange, selected)
      }
      continue
    }
    const selected = normalizeSlicerRecordSelection(getPivotSourceRecordSelection(workbook, changed[0]!.pivot, changed[0]!.field.id))
    if (changed.some(report => !sameSlicerRecordSelection(getPivotSourceRecordSelection(workbook, report.pivot, report.field.id), selected))) {
      throw new SlicerConnectionError('invalid-target', '一次编辑为同一共享切片器指定了不同的报告筛选。')
    }
    applySlicerReportFilter(workbook, reports, selected ?? null)
    updateSharedSelection(workbook, control, slicerRecordSelectionItems(selected, getSlicerSourceValues(workbook, control), getSlicerSourceRecordValues(workbook, control)) ?? null, undefined, selected)
  }
}

function synchronizeTimelineSelections(workbook: WorkbookModel, control: SlicerModel, cache: SlicerCacheModel, before?: WorkbookModel): void {
  let reports: SlicerConnectionReport[]
  try { reports = getSlicerConnectionReports(workbook, control) }
  catch (error) { if (error instanceof SlicerConnectionError) return; throw error }
  const state = readTimelineFilterState(cache)
  try {
    const changed = reports.filter(report => {
      const current = readTimelineReportFilterState(report.pivot, report.field.id)
      if (!before) return !equivalentModelData(current, state)
      const previous = report.kind === 'pivotTable' ? before.pivotTables?.find(pivot => pivot.id === report.id)
        : before.charts?.find(chart => chart.id === report.id)?.ownedPivot
      const oldField = previous && findSlicerReportField(previous, control.sourceFieldId)
      return oldField !== undefined && !equivalentModelData(readTimelineReportFilterState(previous!, oldField.id), current)
    })
    const next = changed.length ? readTimelineReportFilterState(changed[0]!.pivot, changed[0]!.field.id) : state
    if (changed.some(report => !equivalentModelData(readTimelineReportFilterState(report.pivot, report.field.id), next))) {
      throw new SlicerConnectionError('invalid-target', '一次编辑为同一共享时间线指定了不同的报告日期筛选。')
    }
    const legacy = cache.timelineLegacyMemberFilterPending === true || cache.recordSelection !== undefined || cache.selectedItems !== undefined || cache.timelineRange !== undefined
    // A member-only legacy snapshot is preserved until a real date edit supplies an unambiguous replacement.
    if (!changed.length && next.range === undefined && next.condition === undefined && legacy && cache.timelineLegacyMemberFilterPending !== true) return
    const filtered = applyTimelineReportFilter(workbook, reports, next, legacy)
    updateTimelineCacheState(workbook, control, next, filtered)
  } catch (error) {
    if (error instanceof TypeError || error instanceof RangeError) throw new SlicerConnectionError('invalid-target', error.message)
    throw error
  }
}

export function readSlicerSelection(workbook: WorkbookModel, slicer: SlicerModel): string[] | undefined {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind === 'table') {
    const filter = getSlicerTableNativeFilter(workbook, slicer)
    if (filter !== undefined && filter.kind !== 'values') {
      const state = readSlicerItemState(workbook, slicer)
      return [...state.checkedItems, ...state.partialItems]
    }
    return slicerRecordSelectionItems(readSlicerRecordSelection(workbook, slicer), getSlicerSourceValues(workbook, slicer), getSlicerSourceRecordValues(workbook, slicer))
  }
  if (cache.recordSelection !== undefined) return slicerRecordSelectionItems(cache.recordSelection, getSlicerSourceValues(workbook, slicer), getSlicerSourceRecordValues(workbook, slicer))
  return cache.selectedItems ? [...cache.selectedItems] : undefined
}

export function readSlicerRecordSelection(workbook: WorkbookModel, slicer: SlicerModel): PivotRecordSelection | undefined {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind === 'table') {
    const filter = getSlicerTableNativeFilter(workbook, slicer)
    if (filter !== undefined && filter.kind !== 'values') throw new SlicerConnectionError('invalid-target', '当前表格原生条件筛选不能表示为原始成员选择。')
    return filter ? slicerRecordSelectionFromNativeValues(filter) : undefined
  }
  if (cache.recordSelection !== undefined) return normalizeSlicerRecordSelection(cache.recordSelection)
  if (cache.selectedItems === undefined) return undefined
  const values = getSlicerSourceValues(workbook, slicer)
  if (cache.selectedItems.some(item => !values.some(value => String(value ?? '') === item))) {
    throw new SlicerConnectionError('invalid-resource', '旧切片器成员的原始类型尚无法从当前来源确定。')
  }
  return slicerRecordSelectionFromItems(values, cache.selectedItems)
}

/** Raw display items may be partial when scalar types have the same caption. */
export function readSlicerItemState(workbook: WorkbookModel, slicer: SlicerModel): { checkedItems: string[]; partialItems: string[] } {
  const cache = getSlicerCache(workbook, slicer), values = getSlicerSourceValues(workbook, slicer)
  if (cache.source.kind === 'table') {
    const filter = getSlicerTableNativeFilter(workbook, slicer)
    if (filter !== undefined && filter.kind !== 'values') return slicerItemStateFromPredicate(values, tableNativeCellPredicate(workbook, slicer, filter))
  }
  const selection = cache.source.kind === 'table' ? readSlicerRecordSelection(workbook, slicer)
    : cache.recordSelection !== undefined ? cache.recordSelection : slicerRecordSelectionFromItems(values, cache.selectedItems ?? null)
  return slicerRecordSelectionItemState(selection, values, getSlicerSourceRecordValues(workbook, slicer))
}

function tableNativeCellPredicate(workbook: WorkbookModel, slicer: SlicerModel, filter: SheetFilterCriteria): (value: CellScalar, index: number) => boolean {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind !== 'table') throw new SlicerConnectionError('invalid-target', '该切片器不是表格来源。')
  const source = cache.source, table = workbook.tables?.find(candidate => candidate.id === source.tableId)
  const sheet = table && workbook.sheets.find(candidate => candidate.id === table.sheetId)
  if (!table || !sheet) throw new SlicerConnectionError('source-missing', '切片器表格来源不存在。')
  const columnIndex = table.columns.findIndex(candidate => candidate.id === source.columnId)
  if (columnIndex < 0) throw new SlicerConnectionError('field-missing', '切片器表格源字段不存在。')
  const matches = compileSheetCellFilterCriteria(filter), startRow = table.range.startRow + (table.showHeaderRow ? 1 : 0), column = table.range.startCol + columnIndex
  return (_value, index) => matches(sheet.cells[`${startRow + index}:${column}`])
}

function updateSharedSelection(workbook: WorkbookModel, source: SlicerModel, items: string[] | null, range: SlicerTimelineRange | undefined,
  recordSelection?: PivotRecordSelection): void {
  const cache = ensureSlicerCache(workbook, source)
  const selection = normalizeSlicerRecordSelection(recordSelection)
  if (sameSlicerSelection(cache.selectedItems, items ?? undefined) && sameSlicerRecordSelection(cache.recordSelection, selection)
    && equivalentModelData(cache.timelineRange, range)) return
  cache.recordSelection = selection
  cache.selectedItems = items == null ? undefined : [...items]
  cache.timelineRange = range ? cloneModel(range) : undefined
  cache.revision += 1
  for (const control of sharedControls(workbook, source)) {
    control.selectedItems = items == null ? undefined : [...items]
    control.timelineRange = control.type === 'timeline' && range ? { ...range } : undefined
    control.revision += 1
  }
}

function sharedControls(workbook: WorkbookModel, source: SlicerModel): SlicerModel[] {
  const cache = getSlicerCache(workbook, source)
  return (workbook.slicers ?? []).filter(control => {
    try { return getSlicerCache(workbook, control).id === cache.id }
    catch (error) { if (error instanceof SlicerConnectionError) return false; throw error }
  }).map(control => { ensureSlicerCache(workbook, control); return control })
}

export function getSlicerItems(workbook: WorkbookModel, slicer: SlicerModel, locale = 'zh-CN'): string[] {
  return [...new Set(getSlicerDisplayMembers(workbook, slicer, locale).map(member => member.label))]
}

function setTableSlicerRecordFilter(workbook: WorkbookModel, slicer: SlicerModel, selection: PivotRecordSelection | null): void {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind !== 'table') throw new SlicerConnectionError('invalid-target', '该切片器不是表格来源。')
  const source = cache.source
  const table = workbook.tables?.find((candidate) => candidate.id === source.tableId)
  if (!table) throw new Error(`Table not found: ${slicer.tableId}`)
  const sheet = workbook.sheets.find((candidate) => candidate.id === table.sheetId)
  if (!sheet) throw new Error(`Table worksheet not found: ${table.sheetId}`)
  const columnIndex = table.columns.findIndex((candidate) => candidate.id === source.columnId)
  if (columnIndex < 0) throw new Error(`Table source field not found: ${slicer.sourceFieldId}`)
  const range = {
    startRow: table.range.startRow,
    endRow: table.range.endRow - (table.showTotalsRow ? 1 : 0),
    startCol: table.range.startCol,
    endCol: table.range.endCol,
  }
  if (range.endRow <= range.startRow) throw new Error('Table slicer source requires a header and data row')
  if (selection === null && table.autoFilter?.columns[table.range.startCol + columnIndex] === undefined) return
  const sourceView = { ...sheet, autoFilter: table.autoFilter ? structuredClone(table.autoFilter) : undefined }
  if (!sourceView.autoFilter) setAutoFilter(sourceView, range)
  const typed = selection && selection.mode !== 'all' && selection.values.some(value => value !== null && typeof value === 'object')
  const next = selection && selection.mode !== 'all' ? { kind: 'values' as const, selectedItems: selection.values.map(value =>
    value !== null && typeof value === 'object' ? value.type === 'blank' ? null : value.value : value),
    ...(selection.mode === 'exclude' ? { selectionMode: 'exclude' as const } : {}),
    ...(typed ? { recordSelection: cloneModel(selection), recordDate1904: workbook.date1904 === true } : {}) } : null
  const column = table.range.startCol + columnIndex, current = sourceView.autoFilter!.columns[column]
  const criteria = selection !== null && current?.kind === 'values' && sameSlicerRecordSelection(slicerRecordSelectionFromNativeValues(current), selection)
    ? current : next
  setFilterCriteria(sourceView, column, criteria)
  if (!equivalentModelData(table.autoFilter, sourceView.autoFilter)) table.autoFilter = sourceView.autoFilter
}

function normalizeItems(items: string[]): string[] {
  if (!Array.isArray(items)) throw new TypeError('Slicer items must be an array')
  return items.map((item) => String(item)).filter((item, index, all) => all.indexOf(item) === index)
}

function validateNonNegative(value: number, name: string): void {
  if (!Number.isInteger(value) || value < 0) throw new RangeError(`${name} must be a non-negative integer`)
}

function validatePositive(value: number, name: string): void {
  if (!Number.isInteger(value) || value <= 0) throw new RangeError(`${name} must be a positive integer`)
}
