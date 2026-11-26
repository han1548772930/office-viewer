import type { OfficeStudio } from '../../api/office-studio'
import { FOffice, type FSheetSlicer } from '../../api/facade'
import type { CellRange, CellScalar, PivotTableModel, SlicerModel, SlicerTimelineRange, SlicerType, TimelineLevel, WorkbookModel } from '../../domain/model'
import { parseTimelineDate, timelineCalendarToSerial, timelineDateToInput } from './timeline-date'
import { sheetCellRecordValue } from '../sheet-workbench/sheet-record-value.ts'
import type { SlicerCreateInput, SlicerPresentationPatch } from './slicer-engine'
import { slicerBoundsInWorkbook } from './slicer-drawing'
import type { SlicerCommandContext } from '../../api/slicer-command-context.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { createSheetDrawingScene, sheetDrawingRefAffectedRanges } from '../sheet-drawings/sheet-drawing-host.ts'
import { planDrawingArrange } from '../drawing-groups/index.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { getPivotSourceFields } from '../pivot/services/pivot-engine'
import { sourceFieldId } from './slicer-connections'
import { normalizeSheetSlicerInfo, normalizeSheetTimelineInfo, resolveSheetSlicerSource,
  type ISheetSlicerInfo, type ISheetTimelineInfo, type SheetSlicerPlacement } from './slicer-public-info.ts'
import { toTimelinePublicLevel } from './timeline-date-filter.ts'

export interface SlicerPanelTarget { workbookId: string; sheetId: string }
export interface SlicerPanelField { id: string; name: string; available: boolean; controls: SlicerType[]; reason?: string; pivotTableId?: string; tableId?: string; ownedPivotChartId?: string }
export interface SlicerPanelPivot { id: string; label: string; fields: SlicerPanelField[]; kind: 'pivot' | 'table' | 'ownedPivotChart'; sheetId: string; range: CellRange }
export function slicerPanelSourceKey(source: Pick<SlicerPanelPivot, 'kind' | 'id'>): string { return JSON.stringify([source.kind, source.id]) }
export interface SlicerPanelControl { info: SlicerModel; items: string[]; selectedItems?: string[]; canConfigure: boolean; error?: string }
export interface SlicerPanelSnapshot {
  workbookId: string
  sheetId: string
  date1904: boolean
  revision: number
  pivots: SlicerPanelPivot[]
  controls: SlicerPanelControl[]
}
/** The insert panel supplies formal view configuration and CSS/anchor placement; raw drawing models use insertSlicerControl. */
export type SlicerPanelInsert = Pick<SlicerCreateInput, 'pivotTableId' | 'ownedPivotChartId' | 'tableId' | 'sourceFieldId' | 'type'
  | 'name' | 'caption' | 'style' | 'columnCount' | 'rowHeight' | 'showCaption' | 'lockedPosition' | 'startItem'
  | 'level' | 'selectionLevel' | 'scrollPosition' | 'showHeader' | 'showSelectionLabel' | 'showTimeLevel' | 'showHorizontalScrollbar'> & SheetSlicerPlacement
export interface SlicerPanelTimelineInput { start: string; end: string; level: TimelineLevel }

export class SlicerPanelError extends Error {
  readonly code: string
  constructor(code: string, message: string) {
    super(message)
    this.name = 'SlicerPanelError'
    this.code = code
  }
}

function requireTarget(studio: OfficeStudio, target: SlicerPanelTarget) {
  const workbook = FOffice.newAPI(studio).getWorkbook(target.workbookId)
  if (!workbook) throw new SlicerPanelError('WORKBOOK_NOT_FOUND', '当前工作簿已关闭，请重新打开工作簿。')
  const sheet = workbook.getSheetBySheetId(target.sheetId)
  if (!sheet) throw new SlicerPanelError('WORKSHEET_NOT_FOUND', '当前工作表已删除，请切换工作表后重试。')
  return { workbook, sheet }
}

function requireControl(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string): FSheetSlicer {
  const control = requireTarget(studio, target).sheet.getSlicerControl(slicerId)
  if (!control) throw new SlicerPanelError('SLICER_NOT_FOUND', '所选切片器已删除或属于其他工作表，请重新选择。')
  return control
}

function canReadReportSource(model: WorkbookModel, pivot: PivotTableModel, hostSheetId: string, userId: string): boolean {
  const source = model.sheets.find(sheet => sheet.id === pivot.source.subUnitId), host = model.sheets.find(sheet => sheet.id === hostSheetId)
  return Boolean(source && host && canViewWorksheetProtection(source, userId) && canViewWorksheetProtection(host, userId)
    && canViewProtectedRanges(source, [pivot.source.range], userId))
}

function reportPanelFields(model: WorkbookModel, pivot: PivotTableModel, reference: { pivotTableId: string } | { ownedPivotChartId: string }): SlicerPanelField[] {
  const source = model.sheets.find(candidate => candidate.id === pivot.source.subUnitId), sourceRange = pivot.source.range
  const labelFields = new Set(pivot.fields.filter(field => field.area !== 'value').map(field => sourceFieldId(pivot, field)))
  return getPivotSourceFields(model, pivot).map(field => {
    const reason = !labelFields.has(field.fieldId) ? '请先将字段放到透视表的行、列或筛选区域。' : undefined
    const controls: SlicerType[] = ['slicer']
    if (source) {
      for (let row = sourceRange.startRow + 1; row <= sourceRange.endRow; row += 1) {
        if (sheetCellRecordValue(source.cells[`${row}:${field.sourceColumn}`], model.date1904).type === 'date') { controls.push('timeline'); break }
      }
    }
    return { id: field.fieldId, name: field.fieldName, available: !reason, controls, reason, ...reference }
  })
}

/** Every list and selection is derived from the current workbook, including after undo. */
export function readSlicerPanelState(studio: OfficeStudio, target: SlicerPanelTarget): SlicerPanelSnapshot {
  const { workbook, sheet } = requireTarget(studio, target)
  const model = studio.getWorkbookModel(target.workbookId)!
  const pivots = (model.pivotTables ?? []).filter(pivot => canReadReportSource(model, pivot, pivot.anchor.subUnitId, studio.currentUserId)).map((pivotModel, index) => {
    const id = pivotModel.id, fields = reportPanelFields(model, pivotModel, { pivotTableId: id })
    const sheetName = workbook.getSheetBySheetId(pivotModel.anchor.subUnitId)?.getSheetName() ?? '工作表'
    return { id, label: `${sheetName} · 透视表 ${index + 1}`, fields, kind: 'pivot' as const, sheetId: pivotModel.anchor.subUnitId, range: pivotModel.outputRange ?? { startRow: pivotModel.anchor.row, endRow: pivotModel.anchor.row, startCol: pivotModel.anchor.col, endCol: pivotModel.anchor.col } }
  })
  const ownedPivotCharts = (model.charts ?? []).filter(chart => chart.ownedPivot && canReadReportSource(model, chart.ownedPivot, chart.sheetId, studio.currentUserId)).map((chart, index): SlicerPanelPivot => ({
    id: chart.id, label: `${workbook.getSheetBySheetId(chart.sheetId)?.getSheetName() ?? '工作表'} · ${chart.title || `透视图 ${index + 1}`}`,
    kind: 'ownedPivotChart', sheetId: chart.sheetId,
    fields: reportPanelFields(model, chart.ownedPivot!, { ownedPivotChartId: chart.id }),
    range: sheetDrawingRefAffectedRanges(model, chart.sheetId, [{ kind: 'chart', id: chart.id }])[0] ?? { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
  }))
  const tables = (model.tables ?? []).filter((table) => table.sheetId === target.sheetId).map((table) => ({
    id: table.id,
    label: `${workbook.getSheetBySheetId(table.sheetId)?.getSheetName() ?? '工作表'} · ${table.name}`,
    kind: 'table' as const,
    sheetId: table.sheetId,
    range: { ...table.range },
    fields: table.columns.map((column): SlicerPanelField => {
      return { id: column.id, name: column.name, available: true, controls: ['slicer'], tableId: table.id }
    }),
  }))
  const controls = sheet.getSlicerControls().map((control): SlicerPanelControl => {
    const current = control.getModel()
    const info = { ...current, layout: slicerBoundsInWorkbook(model, current) }
    const canConfigure = canConfigurePanelSlicer(studio, target, info.id)
    try {
      return { info, items: control.getItemLabels(), selectedItems: control.getSelectedItems(), canConfigure }
    } catch (error) {
      return { info, items: [], canConfigure, error: error instanceof Error ? error.message : String(error) }
    }
  })
  return { ...target, date1904: model.date1904, revision: studio.getWorkbookRevision(target.workbookId), pivots: [...pivots, ...ownedPivotCharts, ...tables], controls }
}

/** Same native footprint and protected ranges used by the content command. */
export function canConfigurePanelSlicer(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string): boolean {
  const model = studio.getWorkbookModel(target.workbookId), sheet = model?.sheets.find(candidate => candidate.id === target.sheetId),
    control = model?.slicers?.find(candidate => candidate.id === slicerId && candidate.hostSheetId === target.sheetId)
  if (!model || !sheet || !control) return false
  const userId = studio.currentUserId
  if (!canViewWorksheetProtection(sheet, userId) || !canEditWorksheetProtection(sheet, userId)) return false
  try {
    const ranges = sheetDrawingRefAffectedRanges(model, sheet.id, [{ kind: control.type, id: control.id }])
    return canViewProtectedRanges(sheet, ranges, userId) && canEditProtectedRanges(sheet, ranges, userId)
  } catch { return false }
}

/** Mixed pictures, charts, shapes, controls and groups share one sibling container order. */
export function readPanelSlicerArrangement(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string): Record<'forward' | 'backward' | 'front' | 'back', boolean> {
  const unavailable = { forward: false, backward: false, front: false, back: false }
  const model = studio.getWorkbookModel(target.workbookId), control = model?.slicers?.find(candidate => candidate.id === slicerId && candidate.hostSheetId === target.sheetId)
  if (!model || !control || !canConfigurePanelSlicer(studio, target, slicerId)) return unavailable
  try {
    const sheet = model.sheets.find(candidate => candidate.id === target.sheetId)!, userId = studio.currentUserId
    const scene = createSheetDrawingScene(model, target.sheetId, { canEdit: (ref) => {
      const ranges = sheetDrawingRefAffectedRanges(model, sheet.id, [ref])
      return canViewProtectedRanges(sheet, ranges, userId) && canEditProtectedRanges(sheet, ranges, userId)
    } })
    return {
      forward: planDrawingArrange(scene, { selection: [{ kind: control.type, id: control.id }], arrangeType: 'forward' }).status === 'ready',
      backward: planDrawingArrange(scene, { selection: [{ kind: control.type, id: control.id }], arrangeType: 'backward' }).status === 'ready',
      front: planDrawingArrange(scene, { selection: [{ kind: control.type, id: control.id }], arrangeType: 'front' }).status === 'ready',
      back: planDrawingArrange(scene, { selection: [{ kind: control.type, id: control.id }], arrangeType: 'back' }).status === 'ready',
    }
  } catch { return unavailable }
}

/** Unchanged projected bounds never overwrite the authored affine geometry. */
export function buildPanelSlicerPresentationPatch(current: SlicerModel, original: SlicerModel): SlicerPresentationPatch {
  const patch: SlicerPresentationPatch = {}
  for (const key of ['caption', 'style', 'columnCount', 'rowHeight', 'showCaption', 'lockedPosition', 'level', 'selectionLevel',
    'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar'] as const) {
    if (!equivalentModelData(current[key], original[key])) Object.assign(patch, { [key]: current[key] })
  }
  if (!equivalentModelData(current.layout, original.layout) && current.layout) patch.layout = { ...current.layout }
  if ((current.drawing?.anchorType ?? 'position') !== (original.drawing?.anchorType ?? 'position')) patch.anchorType = current.drawing?.anchorType ?? 'position'
  if ((current.drawing?.keepRatio ?? false) !== (original.drawing?.keepRatio ?? false)) patch.keepRatio = current.drawing?.keepRatio ?? false
  return patch
}

function panelCommandContext(studio: OfficeStudio, target: SlicerPanelTarget, context: SlicerCommandContext): SlicerCommandContext {
  const model = studio.getWorkbookModel(target.workbookId)
  if (studio.activeWorkbookId !== target.workbookId || studio.editor.state.activeKind !== 'workbook' || model?.activeSheetId !== target.sheetId) {
    throw new SlicerPanelError('STALE_PANEL_HOST', '切片器所在的工作台已更改，请重新打开面板。')
  }
  return { host: 'workbench', expectedRevision: context.expectedRevision ?? studio.getWorkbookRevision(target.workbookId) }
}

export async function insertPanelSlicer(studio: OfficeStudio, target: SlicerPanelTarget, input: SlicerPanelInsert, context: SlicerCommandContext = {}): Promise<string> {
  if (input.type !== 'slicer' && input.type !== 'timeline') throw new SlicerPanelError('INVALID_SLICER_TYPE', '请选择切片器或时间线。')
  const state = readSlicerPanelState(studio, target)
  const pivot = state.pivots.find((candidate) => candidate.kind === 'pivot' && candidate.id === input.pivotTableId
    || candidate.kind === 'table' && candidate.id === input.tableId || candidate.kind === 'ownedPivotChart' && candidate.id === input.ownedPivotChartId)
  if (!pivot) throw new SlicerPanelError('PIVOT_NOT_FOUND', '请选择现有的数据透视表或表格。')
  const field = pivot.fields.find((candidate) => candidate.id === input.sourceFieldId)
  if (!field) throw new SlicerPanelError('SOURCE_FIELD_NOT_FOUND', '请选择有效的源字段。')
  if (!field.available) throw new SlicerPanelError('SOURCE_FIELD_UNAVAILABLE', field.reason ?? '所选字段暂不可用。')
  if (!field.controls.includes(input.type)) throw new SlicerPanelError('UNSUPPORTED_SOURCE_FIELD', '所选字段不支持此控件，请选择包含有效日期的字段。')
  const sheet = requireTarget(studio, target).sheet, commandContext = panelCommandContext(studio, target, context)
  const model = studio.getWorkbookModel(target.workbookId)!
  const reference = pivot.kind === 'table' ? { subUnitId: pivot.sheetId, tableId: pivot.id }
    : pivot.kind === 'ownedPivotChart' ? { unitId: target.workbookId, subUnitId: pivot.sheetId, pivotChartId: pivot.id }
      : { unitId: target.workbookId, subUnitId: pivot.sheetId, pivotTableId: pivot.id }
  const dataSource = resolveSheetSlicerSource(model, field.id, reference)
  const placement: SheetSlicerPlacement = { ...(input.anchor === undefined ? {} : { anchor: input.anchor }),
    ...(input.position === undefined ? {} : { position: { ...input.position } }), ...(input.size === undefined ? {} : { size: { ...input.size } }) }
  const common = { ...(input.name === undefined ? {} : { name: input.name }), caption: input.caption?.trim() || field.name,
    ...(input.style === undefined ? {} : { style: input.style }) }
  if (input.type === 'timeline') {
    if (dataSource.kind !== 'pivot') throw new SlicerPanelError('INVALID_TIMELINE_SOURCE', '时间线需要原始透视日期来源。')
    const settings: ISheetTimelineInfo['settings'] = { ...common, level: toTimelinePublicLevel(input.level ?? 'month'), selectionLevel: toTimelinePublicLevel(input.selectionLevel ?? 'month') }
    for (const key of ['scrollPosition', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar'] as const) {
      if (Object.hasOwn(input, key)) Object.assign(settings, { [key]: input[key] })
    }
    const control = await sheet.insertTimeline(normalizeSheetTimelineInfo({ dataSource, settings, ...placement }), commandContext)
    return control.getId()
  }
  const settings: ISheetSlicerInfo['settings'] = { ...common, rowHeight: input.rowHeight ?? 24 }
  for (const key of ['columnCount', 'showCaption', 'lockedPosition', 'startItem'] as const) {
    if (Object.hasOwn(input, key)) Object.assign(settings, { [key]: input[key] })
  }
  const control = await sheet.insertSlicer(normalizeSheetSlicerInfo({ dataSource, settings, ...placement }), commandContext)
  return control.getId()
}

export async function applyPanelSlicerPresentation(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string, patch: SlicerPresentationPatch, context: SlicerCommandContext = {}): Promise<void> {
  await requireControl(studio, target, slicerId).updatePresentation(patch, panelCommandContext(studio, target, context))
}

export async function applyPanelSlicerItems(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string, selectedItems: readonly string[], context: SlicerCommandContext = {}): Promise<void> {
  if (!Array.isArray(selectedItems) || selectedItems.some((item) => typeof item !== 'string')) throw new SlicerPanelError('INVALID_SLICER_ITEMS', '筛选项必须是有效的字段值。')
  const control = requireControl(studio, target, slicerId)
  if (control.getModel().type !== 'slicer') throw new SlicerPanelError('INVALID_SLICER_TYPE', '时间线请使用日期范围筛选。')
  const known = new Set(control.getItemLabels())
  if (selectedItems.some((item) => !known.has(item))) throw new SlicerPanelError('UNKNOWN_SLICER_ITEM', '源数据已变化，请重新选择筛选项。')
  // [] means match nothing. Clearing is a separate command, never an empty-array shortcut.
  await control.setItems([...new Set(selectedItems)], panelCommandContext(studio, target, context))
}

/** Date controls use ISO dates; the command receives workbook-system Excel serials. */
export function timelineInputToSerial(value: string, date1904: boolean): number | undefined {
  if (!value.trim()) return undefined
  const serial = timelineCalendarToSerial(value, date1904)
  if (serial === undefined) {
    throw new SlicerPanelError('INVALID_TIMELINE_DATE', '请输入有效日期，格式为年-月-日。')
  }
  return serial
}

export function timelineValueToInput(value: CellScalar | undefined, date1904: boolean): string {
  return timelineDateToInput(value, date1904)
}

export function buildPanelTimelineRange(input: SlicerPanelTimelineInput, date1904: boolean): SlicerTimelineRange {
  if (!['day', 'month', 'quarter', 'year'].includes(input.level)) throw new SlicerPanelError('INVALID_TIMELINE_LEVEL', '请选择日、月、季度或年。')
  const start = timelineInputToSerial(input.start, date1904)
  const end = timelineInputToSerial(input.end, date1904)
  if (start != null && end != null && start > end) throw new SlicerPanelError('INVALID_TIMELINE_ORDER', '开始日期不能晚于结束日期。')
  return { level: input.level, ...(start != null ? { start } : {}), ...(end != null ? { end } : {}) }
}

export async function applyPanelTimelineRange(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string, input: SlicerPanelTimelineInput, context: SlicerCommandContext = {}): Promise<void> {
  const control = requireControl(studio, target, slicerId)
  if (control.getModel().type !== 'timeline') throw new SlicerPanelError('INVALID_SLICER_TYPE', '切片器请使用字段值筛选。')
  const date1904 = studio.getWorkbookModel(target.workbookId)!.date1904
  await control.setTimelineRange(buildPanelTimelineRange(input, date1904), panelCommandContext(studio, target, context))
}

export async function clearPanelSlicer(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string, context: SlicerCommandContext = {}): Promise<void> {
  if (!await requireControl(studio, target, slicerId).clearFilter(panelCommandContext(studio, target, context))) throw new SlicerPanelError('SLICER_COMMAND_REFUSED', '切片器筛选未能清除。')
}

export async function removePanelSlicer(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string, context: SlicerCommandContext = {}): Promise<void> {
  await requireControl(studio, target, slicerId).remove(panelCommandContext(studio, target, context))
}
