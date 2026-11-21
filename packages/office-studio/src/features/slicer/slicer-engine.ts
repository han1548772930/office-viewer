import type {
  SlicerModel,
  SlicerTimelineRange,
  SlicerType,
  TimelineLevel,
  WorkbookModel,
} from '../../domain/model'
import { getPivotSourceFields, setPivotLabelFilter } from '../pivot/services/pivot-engine'
import { parseTimelineDate, timelinePeriodKey, type TimelineDate } from './timeline-date'

export interface SlicerCreateInput {
  id: string
  unitId: string
  hostSheetId: string
  pivotTableId: string
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
}

export function createSlicer(workbook: WorkbookModel, input: SlicerCreateInput): SlicerModel {
  if (input.unitId !== workbook.id) throw new Error('Slicer must belong to the target workbook')
  if (!workbook.sheets.some((sheet) => sheet.id === input.hostSheetId)) throw new Error(`Slicer host worksheet not found: ${input.hostSheetId}`)
  const pivot = workbook.pivotTables?.find((candidate) => candidate.id === input.pivotTableId)
  if (!pivot) throw new Error(`Pivot table not found: ${input.pivotTableId}`)
  if (pivot.source.unitId !== workbook.id) throw new Error('Pivot source must belong to the target workbook')
  const sourceField = getPivotSourceFields(workbook, pivot).find((field) => field.fieldId === String(input.sourceFieldId))
  if (!sourceField) throw new Error(`Pivot source field not found: ${input.sourceFieldId}`)
  if (!pivot.fields.some((field) => field.id === sourceField.fieldId && field.area !== 'value')) {
    throw new Error(`Pivot source field must be placed in a label area: ${sourceField.fieldId}`)
  }
  if (workbook.slicers?.some((slicer) => slicer.pivotTableId === pivot.id && slicer.sourceFieldId === sourceField.fieldId)) {
    throw new Error(`A slicer is already bound to pivot field ${sourceField.fieldId}`)
  }
  if (input.type !== 'slicer' && input.type !== 'timeline') throw new TypeError(`Invalid slicer type: ${input.type}`)
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
    pivotTableId: input.pivotTableId,
    sourceFieldId: sourceField.fieldId,
    type: input.type,
    caption: input.caption?.trim() || sourceField.fieldName,
    name: input.name?.trim() || sourceField.fieldName,
    position: { ...position },
    size: { ...size },
    layout: input.layout ? { ...input.layout } : { left: 200, top: 200, width: input.type === 'timeline' ? 420 : 220, height: input.type === 'timeline' ? 140 : 280 },
    style: input.style || (input.type === 'timeline' ? 'TimelineStyleLight1' : 'SlicerStyleLight1'),
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
  validateSlicerPresentation(model)
  if (model.type === 'timeline' && getSlicerItems(workbook, model).length === 0) throw new Error('Timeline requires a date field with at least one valid date')
  workbook.slicers ??= []
  if (workbook.slicers.some((candidate) => candidate.id === model.id)) throw new Error(`Slicer already exists: ${model.id}`)
  workbook.slicers.push(model)
  return model
}

export type SlicerPresentationPatch = Partial<Pick<SlicerModel,
  | 'name' | 'caption' | 'style' | 'startItem' | 'columnCount' | 'rowHeight' | 'showCaption' | 'lockedPosition'
  | 'level' | 'selectionLevel' | 'showHeader' | 'showSelectionLabel' | 'showTimeLevel' | 'showHorizontalScrollbar' | 'scrollPosition'
>> & Partial<Pick<SlicerModel, 'position' | 'size' | 'layout'>>

export function updateSlicerPresentation(slicer: SlicerModel, patch: SlicerPresentationPatch): void {
  validateSlicerPresentation(patch)
  const allowed = new Set([
    'name', 'caption', 'style', 'startItem', 'columnCount', 'rowHeight', 'showCaption', 'lockedPosition',
    'level', 'selectionLevel', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar', 'scrollPosition',
    'position', 'size', 'layout',
  ])
  if (Object.keys(patch).some((key) => !allowed.has(key))) throw new TypeError('Unknown slicer presentation property')
  Object.assign(slicer, {
    ...patch,
    ...(patch.position ? { position: { ...patch.position } } : {}),
    ...(patch.size ? { size: { ...patch.size } } : {}),
    ...(patch.layout ? { layout: { ...patch.layout } } : {}),
  })
  slicer.revision += 1
}

function validateSlicerPresentation(patch: SlicerPresentationPatch): void {
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
  const pivot = requirePivot(workbook, slicer)
  const items = selectedItems == null ? null : normalizeItems(selectedItems)
  if (items && items.some((item) => !getSlicerItems(workbook, slicer).includes(item))) throw new Error('Slicer selection contains an unknown item')
  setPivotLabelFilter(workbook, pivot, slicer.sourceFieldId, items)
  slicer.selectedItems = items == null ? undefined : items
  slicer.timelineRange = undefined
  slicer.revision += 1
}

export function setSlicerTimelineRange(workbook: WorkbookModel, slicer: SlicerModel, range: SlicerTimelineRange | null): void {
  if (slicer.type !== 'timeline') throw new Error('Slicer controls require selected items')
  const pivot = requirePivot(workbook, slicer)
  if (range != null) validateTimelineRange(range, workbook.date1904)
  const items = range == null ? null : timelineItems(workbook, pivot, slicer.sourceFieldId, range)
  setPivotLabelFilter(workbook, pivot, slicer.sourceFieldId, items)
  slicer.timelineRange = range == null ? undefined : { ...range }
  slicer.selectedItems = items == null ? undefined : items
  slicer.revision += 1
}

export function clearSlicerFilter(workbook: WorkbookModel, slicer: SlicerModel): void {
  if (slicer.type === 'timeline') setSlicerTimelineRange(workbook, slicer, null)
  else setSlicerItems(workbook, slicer, null)
}

export function removeSlicer(workbook: WorkbookModel, slicerId: string): void {
  const slicer = workbook.slicers?.find((candidate) => candidate.id === slicerId)
  if (!slicer) throw new Error(`Slicer not found: ${slicerId}`)
  clearSlicerFilter(workbook, slicer)
  workbook.slicers = (workbook.slicers ?? []).filter((candidate) => candidate.id !== slicerId)
}

export function getSlicer(workbook: WorkbookModel, slicerId: string): SlicerModel | undefined {
  return workbook.slicers?.find((candidate) => candidate.id === slicerId)
}

export function getSlicerItems(workbook: WorkbookModel, slicer: SlicerModel): string[] {
  const pivot = requirePivot(workbook, slicer)
  const field = getPivotSourceFields(workbook, pivot).find((candidate) => candidate.fieldId === slicer.sourceFieldId)
  if (!field) throw new Error(`Pivot source field not found: ${slicer.sourceFieldId}`)
  const sheet = workbook.sheets.find((candidate) => candidate.id === pivot.source.subUnitId)
  if (!sheet) throw new Error(`Pivot source worksheet not found: ${pivot.source.subUnitId}`)
  const items = new Set<string>()
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row += 1) {
    const value = sheet.cells[`${row}:${field.sourceColumn}`]?.value ?? null
    if (slicer.type === 'timeline' ? parseTimelineDate(value, workbook.date1904) != null : value !== null && value !== '') items.add(String(value))
  }
  return [...items]
}

function requirePivot(workbook: WorkbookModel, slicer: SlicerModel) {
  if (slicer.unitId !== workbook.id) throw new Error('Slicer workbook does not match the target workbook')
  const pivot = workbook.pivotTables?.find((candidate) => candidate.id === slicer.pivotTableId)
  if (!pivot) throw new Error(`Pivot table not found: ${slicer.pivotTableId}`)
  if (!pivot.fields.some((field) => field.id === slicer.sourceFieldId && field.area !== 'value')) {
    throw new Error(`Pivot field is not configured as a label field: ${slicer.sourceFieldId}`)
  }
  return pivot
}

function normalizeItems(items: string[]): string[] {
  if (!Array.isArray(items)) throw new TypeError('Slicer items must be an array')
  return items.map((item) => String(item)).filter((item, index, all) => all.indexOf(item) === index)
}

function timelineItems(workbook: WorkbookModel, pivot: ReturnType<typeof requirePivot>, fieldId: string, range: SlicerTimelineRange): string[] {
  const field = getPivotSourceFields(workbook, pivot).find((candidate) => candidate.fieldId === fieldId)
  if (!field) throw new Error(`Pivot source field not found: ${fieldId}`)
  const sheet = workbook.sheets.find((candidate) => candidate.id === pivot.source.subUnitId)
  if (!sheet) throw new Error(`Pivot source worksheet not found: ${pivot.source.subUnitId}`)
  const values: string[] = []
  for (let row = pivot.source.range.startRow + 1; row <= pivot.source.range.endRow; row += 1) {
    const value = sheet.cells[`${row}:${field.sourceColumn}`]?.value ?? null
    const date = parseTimelineDate(value, workbook.date1904)
    if (!date) continue
    if (dateWithinRange(date, range, workbook.date1904)) {
      const key = String(value)
      if (!values.includes(key)) values.push(key)
    }
  }
  return values
}

function dateWithinRange(date: TimelineDate, range: SlicerTimelineRange, date1904: boolean): boolean {
  const start = range.start == null ? undefined : parseTimelineDate(range.start, date1904)
  const end = range.end == null ? undefined : parseTimelineDate(range.end, date1904)
  const key = timelinePeriodKey(date, range.level)
  return (!start || key >= timelinePeriodKey(start, range.level)) && (!end || key <= timelinePeriodKey(end, range.level))
}

function validateTimelineRange(range: SlicerTimelineRange, date1904: boolean): void {
  if (!['day', 'month', 'quarter', 'year'].includes(range.level)) throw new TypeError(`Invalid timeline level: ${range.level}`)
  const start = range.start == null ? undefined : parseTimelineDate(range.start, date1904)
  const end = range.end == null ? undefined : parseTimelineDate(range.end, date1904)
  if (range.start != null && !start || range.end != null && !end) throw new TypeError('Timeline range must contain valid dates')
  if (start && end && start.serial > end.serial) throw new RangeError('Timeline range start must not exceed end')
}

function validateNonNegative(value: number, name: string): void {
  if (!Number.isInteger(value) || value < 0) throw new RangeError(`${name} must be a non-negative integer`)
}

function validatePositive(value: number, name: string): void {
  if (!Number.isInteger(value) || value <= 0) throw new RangeError(`${name} must be a positive integer`)
}
