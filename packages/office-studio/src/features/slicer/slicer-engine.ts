import type {
  SlicerModel,
  SlicerTimelineRange,
  SlicerType,
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
  position?: { row: number; column: number }
  size?: { rows: number; columns: number }
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
    position: { ...position },
    size: { ...size },
    revision: 0,
  }
  if (model.type === 'timeline' && getSlicerItems(workbook, model).length === 0) throw new Error('Timeline requires a date field with at least one valid date')
  workbook.slicers ??= []
  if (workbook.slicers.some((candidate) => candidate.id === model.id)) throw new Error(`Slicer already exists: ${model.id}`)
  workbook.slicers.push(model)
  return model
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
