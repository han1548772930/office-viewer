import type { SlicerCacheModel, SlicerCacheSource, SlicerModel, SheetSlicerCacheOptions, SheetSlicerConnections, WorkbookModel } from '../../domain/model.ts'
import { isRecordValue, normalizeRecordSelection, recordValueFromScalar, type RecordSelection } from '../../domain/record-value.ts'
import type { ITimelineDateCondition, ITimelineRange } from '../../domain/timeline-filters.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { drawingAffineBounds, drawingAffineFrame, drawingAffineWithFrame, type DrawingAffineGeometry } from '../../domain/drawing-affine.ts'
import { parseSheetA1Range } from '../sheet-workbench/sheet-a1-reference.ts'
import { sheetClipboardObjectAxes } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'
import { getSlicerCache, resolvePivotCacheId, sourceFieldId, SlicerConnectionError } from './slicer-connections.ts'
import { getPivotSourceFields } from '../pivot/services/pivot-engine.ts'
import { slicerDrawingGeometry } from './slicer-drawing.ts'
import { normalizeSlicerSettingsPatch, readSlicerSettings, type SheetSlicerSettings } from './slicer-settings.ts'
import { normalizeTimelineSettingsPatch, readTimelineSettings, type SheetTimelineSettings, type TimelineSettingsInputPatch } from './timeline-settings.ts'
import type { SlicerPresentationPatch } from './slicer-engine.ts'

export type SheetSlicerAnchorSpec = string | { row: number; column: number; rowOffset?: number; columnOffset?: number }
export type SheetSlicerSourceSpec = Extract<SlicerCacheSource, { kind: 'table' }>
  | (Extract<SlicerCacheSource, { kind: 'pivot' }> & SheetSlicerConnections)
export type SheetTimelineSourceSpec = Extract<SheetSlicerSourceSpec, { kind: 'pivot' }>
export interface SheetSlicerPlacement { anchor?: SheetSlicerAnchorSpec; position?: { x: number; y: number }; size?: { width: number; height: number } }
export interface ISheetSlicerInfo extends SheetSlicerPlacement { dataSource: SheetSlicerSourceSpec; settings: Omit<SheetSlicerSettings, 'name'> & { name?: string } }
export interface ISheetTimelineInfo extends SheetSlicerPlacement { dataSource: SheetTimelineSourceSpec; settings: Omit<SheetTimelineSettings, 'name'> & { name?: string } }
export type TimelinePresentationPatch = Omit<SlicerPresentationPatch, 'level' | 'selectionLevel'> & TimelineSettingsInputPatch
export type SlicerCacheDefinition = ({ id: string; kind: 'pivot'; name: string; sourceName: string; sourceFieldId: string;
  data: Omit<SheetSlicerCacheOptions, 'hideItemsWithNoData'> & { pivotCacheId: string }; selection: RecordSelection } & SheetSlicerConnections
  | { id: string; kind: 'table'; name: string; sourceName: string; tableSlicerCache: Omit<SheetSlicerCacheOptions, 'hideItemsWithNoData'>
    & { subUnitId: string; tableId: string; column: string } }) & { hideItemsWithNoData?: boolean }
export interface ITimelineCacheDefinition extends SheetSlicerConnections {
  id: string; name: string; sourceName: string; sourceFieldId: string; pivotCacheId: string
  state: { selection?: ITimelineRange; condition?: never } | { selection?: never; condition: ITimelineDateCondition }
}
type SheetSlicerDrawingMarker = { row: number; column: number; rowOffset: number; columnOffset: number }
export interface ISheetSlicerDrawing {
  unitId: string; subUnitId: string; drawingId: string; drawingType: 10 | 11; slicerViewId: string
  anchorType: '0' | '1' | '2'
  axisAlignSheetTransform: { from: SheetSlicerDrawingMarker; to: SheetSlicerDrawingMarker }
  transform: { left: number; top: number; width: number; height: number; angle: number; skewX: number; flipX: boolean; flipY: boolean }
  /** Exact authored affine basis, retained for shared Drawing APIs and lossless local save/reopen. */
  drawingGeometry: DrawingAffineGeometry
  zIndex?: number
}
export const SLICER_PUBLIC_COMMANDS = {
  insert: 'sheet.command.insert-slicer', insertTimeline: 'sheet.command.insert-timeline',
  update: 'sheet.command.update-slicer-info', updateTimeline: 'sheet.command.update-timeline-info',
  timelineSettings: 'sheet.command.update-timeline-settings', dateRange: 'sheet.command.set-timeline-date-range',
  dateCondition: 'sheet.command.set-timeline-date-condition', placement: 'sheet.command.set-slicer-placement',
} as const

const invalid = (message: string): never => { throw new SlicerConnectionError('invalid-target', message) }
const object = (value: unknown, label: string): Record<string, unknown> => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return invalid(`${label}无效。`)
  return value as Record<string, unknown>
}
const id = (value: unknown): string => typeof value === 'string' && value.trim() ? value : invalid('切片器来源标识无效。')
function only(value: Record<string, unknown>, keys: readonly string[], label: string): void {
  if (Object.keys(value).some(key => !keys.includes(key))) invalid(`${label}包含未知属性。`)
}
export function normalizeSheetSlicerAnchor(input: SheetSlicerAnchorSpec): Exclude<SheetSlicerAnchorSpec, string> {
  if (typeof input === 'string') {
    if (!/^\$?[A-Za-z]+\$?[1-9]\d*$/.test(input)) return invalid('切片器锚点必须是单个单元格地址。')
    const range = parseSheetA1Range(input)
    return { row: range.startRow, column: range.startColumn }
  }
  const value = object(input, '切片器锚点'); only(value, ['row', 'column', 'rowOffset', 'columnOffset'], '切片器锚点')
  if (!Number.isSafeInteger(value.row) || Number(value.row) < 0 || !Number.isSafeInteger(value.column) || Number(value.column) < 0) return invalid('切片器锚点行列无效。')
  const result = { row: Number(value.row), column: Number(value.column) } as Exclude<SheetSlicerAnchorSpec, string>
  for (const key of ['rowOffset', 'columnOffset'] as const) if (value[key] !== undefined) {
    if (typeof value[key] !== 'number' || !Number.isFinite(value[key])) return invalid('切片器锚点偏移必须是有限 CSS 像素。')
    result[key] = value[key]
  }
  return result
}
export function normalizeSheetSlicerPlacement(input: SheetSlicerPlacement): SheetSlicerPlacement {
  const result: SheetSlicerPlacement = {}
  if (input.anchor !== undefined) result.anchor = normalizeSheetSlicerAnchor(input.anchor)
  if (input.position !== undefined) {
    const position = object(input.position, '切片器位置'); only(position, ['x', 'y'], '切片器位置')
    if (typeof position.x !== 'number' || typeof position.y !== 'number' || !Number.isFinite(position.x) || !Number.isFinite(position.y)) return invalid('切片器位置必须是有限 CSS 像素。')
    result.position = { x: position.x, y: position.y }
  }
  if (input.size !== undefined) {
    const size = object(input.size, '切片器尺寸'); only(size, ['width', 'height'], '切片器尺寸')
    if (typeof size.width !== 'number' || typeof size.height !== 'number' || !Number.isFinite(size.width) || !Number.isFinite(size.height)
      || size.width <= 0 || size.height <= 0) return invalid('切片器尺寸必须是正的有限 CSS 像素。')
    result.size = { width: size.width, height: size.height }
  }
  if (result.anchor && result.position) return invalid('切片器不能同时使用单元格锚点与绝对位置。')
  return result
}
export function normalizeSheetSlicerSource(input: SheetSlicerSourceSpec): SheetSlicerSourceSpec {
  const value = object(input, '切片器来源')
  if (value.kind === 'table') {
    only(value, ['kind', 'subUnitId', 'tableId', 'columnId'], '表格切片器来源')
    return { kind: 'table', subUnitId: id(value.subUnitId), tableId: id(value.tableId), columnId: id(value.columnId) }
  }
  if (value.kind !== 'pivot') return invalid('切片器来源类型无效。')
  only(value, ['kind', 'pivotCacheId', 'sourceFieldId', 'pivotTables', 'ownedPivotCharts'], '透视切片器来源')
  if (!Array.isArray(value.pivotTables) || value.ownedPivotCharts !== undefined && !Array.isArray(value.ownedPivotCharts)) return invalid('透视来源连接列表无效。')
  const refs = <K extends 'pivotTableId' | 'pivotChartId'>(values: unknown[], key: K): Array<{ subUnitId: string } & Record<K, string>> => {
    const seen = new Set<string>()
    return values.flatMap(item => {
      const ref = object(item, '切片器报告引用'); only(ref, ['subUnitId', key], '切片器报告引用')
      const subUnitId = id(ref.subUnitId), reportId = id(ref[key]), identity = JSON.stringify([subUnitId, reportId])
      if (seen.has(identity)) return []; seen.add(identity)
      return [{ subUnitId, [key]: reportId } as { subUnitId: string } & Record<K, string>]
    })
  }
  return { kind: 'pivot', pivotCacheId: id(value.pivotCacheId), sourceFieldId: id(value.sourceFieldId), pivotTables: refs(value.pivotTables, 'pivotTableId'),
    ...(value.ownedPivotCharts === undefined ? {} : { ownedPivotCharts: refs(value.ownedPivotCharts as unknown[], 'pivotChartId') }) }
}

/** Read-only resolution never creates a cache or report field, and never treats a display name/index as identity. */
export function resolveSheetSlicerSource(workbook: WorkbookModel, fieldId: string,
  reference: { subUnitId: string; tableId: string } | { unitId: string; subUnitId: string; pivotTableId: string } | { unitId: string; subUnitId: string; pivotChartId: string }): SheetSlicerSourceSpec {
  id(fieldId)
  if ('tableId' in reference) {
    const table = workbook.tables?.find(table => table.id === reference.tableId && table.sheetId === reference.subUnitId)
    if (!table || !table.columns.some(column => column.id === fieldId)) return invalid('表格或稳定列标识不存在。')
    return { kind: 'table', subUnitId: table.sheetId, tableId: table.id, columnId: fieldId }
  }
  if (reference.unitId !== workbook.id) return invalid('切片器来源必须属于同一工作簿。')
  const chart = 'pivotChartId' in reference ? workbook.charts?.find(chart => chart.id === reference.pivotChartId && chart.sheetId === reference.subUnitId) : undefined
  const pivot = 'pivotTableId' in reference ? workbook.pivotTables?.find(pivot => pivot.id === reference.pivotTableId && pivot.anchor.subUnitId === reference.subUnitId)
    : chart?.pivotTableId ? workbook.pivotTables?.find(pivot => pivot.id === chart.pivotTableId) : chart?.ownedPivot
  if (!pivot || 'pivotChartId' in reference && !chart) return invalid('透视表或透视图来源不存在。')
  const pivotCacheId = resolvePivotCacheId(workbook, pivot), retained = workbook.pivotCaches?.find(cache => cache.id === pivotCacheId)
  const field = retained?.fields?.find(field => field.id === fieldId)
  const reportField = pivot.fields.find(field => sourceFieldId(pivot, field) === fieldId)
  if (!field && !reportField && !getPivotSourceFields(workbook, pivot).some(field => field.fieldId === fieldId)) return invalid('透视来源稳定字段标识不存在。')
  return { kind: 'pivot', pivotCacheId, sourceFieldId: fieldId,
    pivotTables: chart?.ownedPivot && !chart.pivotTableId ? [] : [{ subUnitId: pivot.anchor.subUnitId, pivotTableId: pivot.id }],
    ...(chart?.ownedPivot && !chart.pivotTableId ? { ownedPivotCharts: [{ subUnitId: chart.sheetId, pivotChartId: chart.id }] } : {}) }
}
export function readSheetSlicerSource(workbook: WorkbookModel, view: SlicerModel): SheetSlicerSourceSpec | undefined {
  let cache: SlicerCacheModel | undefined = view.cacheId ? workbook.slicerCaches?.find(cache => cache.id === view.cacheId && cache.type === view.type && cache.unitId === workbook.id) : undefined
  if (!cache && !view.cacheId) { try { cache = getSlicerCache(workbook, view) } catch (error) { if (!(error instanceof SlicerConnectionError)) throw error } }
  if (!cache || (cache.source.kind === 'table' ? cache.source.columnId : cache.source.sourceFieldId) !== view.sourceFieldId) return undefined
  return cache.source.kind === 'table' ? cloneModel(cache.source) : { ...cloneModel(cache.source), ...cloneModel(cache.connections ?? { pivotTables: [] }) }
}
function cacheSourceName(workbook: WorkbookModel, cache: SlicerCacheModel): string {
  if (cache.sourceName !== undefined) return cache.sourceName
  const source = cache.source
  return source.kind === 'table' ? workbook.tables?.find(table => table.id === source.tableId)?.columns.find(column => column.id === source.columnId)?.name ?? source.columnId
    : workbook.pivotCaches?.find(pivot => pivot.id === source.pivotCacheId)?.fields?.find(field => field.id === source.sourceFieldId)?.name ?? source.sourceFieldId
}
export function sheetSlicerCacheDefinition(workbook: WorkbookModel, cache: SlicerCacheModel): SlicerCacheDefinition {
  if (cache.type !== 'slicer') return invalid('时间线使用独立缓存定义。')
  const { hideItemsWithNoData, ...options } = cache.options ?? {}, base = { id: cache.id, name: cache.name ?? cache.id, sourceName: cacheSourceName(workbook, cache),
    ...(hideItemsWithNoData === undefined ? {} : { hideItemsWithNoData }) }
  if (cache.source.kind === 'table') return { ...base, kind: 'table', tableSlicerCache: { ...options,
    subUnitId: cache.source.subUnitId, tableId: cache.source.tableId, column: cache.source.columnId } }
  const raw = cache.recordSelection, selection: RecordSelection = raw
    ? raw.mode === 'all' ? { mode: 'all' } : normalizeRecordSelection({ mode: raw.mode, values: raw.values.map(value => isRecordValue(value) ? value : recordValueFromScalar(value)) })
    : cache.selectedItems === undefined ? { mode: 'all' } : normalizeRecordSelection({ mode: 'include', values: cache.selectedItems.map(value => ({ type: 'string', value })) })
  return { ...base, kind: 'pivot', sourceFieldId: cache.source.sourceFieldId, data: { ...options, pivotCacheId: cache.source.pivotCacheId },
    ...cloneModel(cache.connections ?? { pivotTables: [] }), selection }
}
export function sheetTimelineCacheDefinition(workbook: WorkbookModel, cache: SlicerCacheModel): ITimelineCacheDefinition {
  if (cache.type !== 'timeline' || cache.source.kind !== 'pivot') return invalid('时间线需要透视来源缓存。')
  if (cache.timelineDateRange !== undefined && cache.timelineDateCondition !== undefined) return invalid('时间线日期条件与范围不能同时存在。')
  return { id: cache.id, name: cache.name ?? cache.id, sourceName: cacheSourceName(workbook, cache), sourceFieldId: cache.source.sourceFieldId,
    pivotCacheId: cache.source.pivotCacheId, ...cloneModel(cache.connections ?? { pivotTables: [] }), state: cache.timelineDateCondition !== undefined
      ? { condition: cloneModel(cache.timelineDateCondition) } : cache.timelineDateRange !== undefined ? { selection: cloneModel(cache.timelineDateRange) } : {} }
}
export function sameSheetSlicerSource(first: SheetSlicerSourceSpec, second: SheetSlicerSourceSpec): boolean {
  const key = (input: SheetSlicerSourceSpec) => {
    const source = normalizeSheetSlicerSource(input)
    return source.kind === 'table' ? JSON.stringify(source) : JSON.stringify([source.kind, source.pivotCacheId, source.sourceFieldId,
      source.pivotTables.map(ref => JSON.stringify([ref.subUnitId, ref.pivotTableId])).sort(),
      (source.ownedPivotCharts ?? []).map(ref => JSON.stringify([ref.subUnitId, ref.pivotChartId])).sort()])
  }
  return key(first) === key(second)
}
export function readSheetSlicerPlacement(workbook: WorkbookModel, view: SlicerModel): SheetSlicerPlacement {
  const sheet = workbook.sheets.find(sheet => sheet.id === view.hostSheetId)
  if (!sheet) return invalid('切片器工作表不存在。')
  const geometry = slicerDrawingGeometry(view, sheetClipboardObjectAxes(workbook, sheet)), frame = drawingAffineFrame(geometry)
  return { ...(view.drawing?.anchorType === 'none' ? { position: { x: frame.left, y: frame.top } }
    : { anchor: view.drawing ? { row: view.drawing.from.row, column: view.drawing.from.column, rowOffset: view.drawing.from.offsetY, columnOffset: view.drawing.from.offsetX }
      : { row: view.position.row, column: view.position.column } }), size: { width: frame.width, height: frame.height } }
}
export function sheetSlicerPlacementPatch(workbook: WorkbookModel, view: SlicerModel, input: SheetSlicerPlacement): SlicerPresentationPatch {
  const placement = normalizeSheetSlicerPlacement(input), sheet = workbook.sheets.find(sheet => sheet.id === view.hostSheetId)
  if (!sheet) return invalid('切片器工作表不存在。')
  const axes = sheetClipboardObjectAxes(workbook, sheet), geometry = slicerDrawingGeometry(view, axes), prior = drawingAffineFrame(geometry)
  const frame = { ...prior, ...placement.size, ...(placement.position ? { left: placement.position.x, top: placement.position.y } : {}) }
  let next = drawingAffineWithFrame(geometry, frame)
  if (placement.anchor !== undefined) {
    const anchor = normalizeSheetSlicerAnchor(placement.anchor)
    if (anchor.row >= (sheet.rowCount ?? 1_048_576) || anchor.column >= (sheet.columnCount ?? 16_384)) return invalid('切片器锚点超出工作表。')
    const bounds = drawingAffineBounds(next), left = axes.columns.offset(anchor.column) + (anchor.columnOffset ?? 0), top = axes.rows.offset(anchor.row) + (anchor.rowOffset ?? 0)
    next = { ...next, matrix: { ...next.matrix, e: next.matrix.e + left - bounds.left, f: next.matrix.f + top - bounds.top } }
  }
  return { drawingGeometry: next, ...(placement.anchor !== undefined ? { anchorType: 'position' as const } : placement.position ? { anchorType: 'none' as const } : {}) }
}
export function readSheetSlicerDrawing(workbook: WorkbookModel, view: SlicerModel): ISheetSlicerDrawing {
  const sheet = workbook.sheets.find(sheet => sheet.id === view.hostSheetId)
  if (!sheet || !view.drawing) return invalid('切片器绘图宿主不可用。')
  const geometry = slicerDrawingGeometry(view, sheetClipboardObjectAxes(workbook, sheet)), frame = drawingAffineFrame(geometry), m = geometry.matrix
  const determinant = m.a * m.d - m.b * m.c, dot = m.a * m.c + m.b * m.d
  return { unitId: view.unitId, subUnitId: view.hostSheetId, drawingId: view.id, drawingType: view.type === 'timeline' ? 11 : 10, slicerViewId: view.id,
    anchorType: view.drawing.anchorType === 'none' ? '2' : view.drawing.anchorType === 'both' ? '0' : '1',
    axisAlignSheetTransform: { from: { row: view.drawing.from.row, column: view.drawing.from.column, rowOffset: view.drawing.from.offsetY, columnOffset: view.drawing.from.offsetX },
      to: { row: view.drawing.to.row, column: view.drawing.to.column, rowOffset: view.drawing.to.offsetY, columnOffset: view.drawing.to.offsetX } },
    transform: { ...frame, angle: Math.atan2(m.b, m.a) * 180 / Math.PI, skewX: Math.atan2(dot, Math.abs(determinant)) * 180 / Math.PI, flipX: false, flipY: determinant < 0 },
    drawingGeometry: cloneModel(geometry), ...(view.drawing.zIndex === undefined ? {} : { zIndex: view.drawing.zIndex }) }
}
export function readSheetSlicerInfo(workbook: WorkbookModel, view: SlicerModel): ISheetSlicerInfo {
  if (view.type !== 'slicer') return invalid('时间线使用独立信息接口。')
  const dataSource = readSheetSlicerSource(workbook, view)
  if (!dataSource) return invalid('切片器来源资源不可用。')
  return { dataSource, settings: readSlicerSettings(view), ...readSheetSlicerPlacement(workbook, view) }
}
export function readSheetTimelineInfo(workbook: WorkbookModel, view: SlicerModel): ISheetTimelineInfo {
  if (view.type !== 'timeline') return invalid('普通切片器使用独立信息接口。')
  const dataSource = readSheetSlicerSource(workbook, view)
  if (!dataSource || dataSource.kind !== 'pivot') return invalid('时间线来源资源不可用。')
  return { dataSource, settings: readTimelineSettings(view), ...readSheetSlicerPlacement(workbook, view) }
}
export function normalizeSheetSlicerInfo(input: ISheetSlicerInfo): ISheetSlicerInfo {
  const value = object(input, '切片器信息'); only(value, ['dataSource', 'settings', 'anchor', 'position', 'size'], '切片器信息')
  const raw = { ...object(value.settings, '切片器设置') }; if (raw.name === undefined) delete raw.name
  const settings = normalizeSlicerSettingsPatch({ ...raw, rowHeight: Object.hasOwn(raw, 'rowHeight') ? raw.rowHeight : 24 })
  return { dataSource: normalizeSheetSlicerSource(input.dataSource), settings: { ...settings, rowHeight: settings.rowHeight ?? 24 }, ...normalizeSheetSlicerPlacement(input) }
}
export function normalizeSheetTimelineInfo(input: ISheetTimelineInfo): ISheetTimelineInfo {
  const value = object(input, '时间线信息'); only(value, ['dataSource', 'settings', 'anchor', 'position', 'size'], '时间线信息')
  const raw = { ...object(value.settings, '时间线设置') }; if (raw.name === undefined) delete raw.name
  const settings = normalizeTimelineSettingsPatch({ level: 2, selectionLevel: 2, ...raw }), dataSource = normalizeSheetSlicerSource(input.dataSource)
  if (dataSource.kind !== 'pivot') return invalid('时间线必须使用透视日期来源。')
  const level = settings.level === 'year' ? 0 : settings.level === 'quarter' ? 1 : settings.level === 'day' ? 3 : 2
  const selectionLevel = settings.selectionLevel === 'year' ? 0 : settings.selectionLevel === 'quarter' ? 1 : settings.selectionLevel === 'day' ? 3 : 2
  return { dataSource, settings: { ...settings, level, selectionLevel }, ...normalizeSheetSlicerPlacement(input) }
}
