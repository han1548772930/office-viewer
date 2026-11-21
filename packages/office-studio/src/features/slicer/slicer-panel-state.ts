import type { OfficeStudio } from '../../api/office-studio'
import { FOffice, type FSheetSlicer } from '../../api/facade'
import type { CellScalar, SlicerModel, SlicerTimelineRange, SlicerType, TimelineLevel } from '../../domain/model'
import { parseTimelineDate, timelineCalendarToSerial, timelineDateToInput } from './timeline-date'
import { classifySheetNumberFormat } from '../sheet-workbench/sheet-number-format'
import type { SlicerCreateInput, SlicerPresentationPatch } from './slicer-engine'

export interface SlicerPanelTarget { workbookId: string; sheetId: string }
export interface SlicerPanelField { id: string; name: string; available: boolean; controls: SlicerType[]; reason?: string }
export interface SlicerPanelPivot { id: string; label: string; fields: SlicerPanelField[] }
export interface SlicerPanelControl { info: SlicerModel; items: string[]; selectedItems?: string[]; error?: string }
export interface SlicerPanelSnapshot {
  workbookId: string
  sheetId: string
  date1904: boolean
  pivots: SlicerPanelPivot[]
  controls: SlicerPanelControl[]
}
export type SlicerPanelInsert = Omit<SlicerCreateInput, 'id' | 'unitId' | 'hostSheetId'>
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
  const control = requireTarget(studio, target).sheet.getSlicer(slicerId)
  if (!control) throw new SlicerPanelError('SLICER_NOT_FOUND', '所选切片器已删除或属于其他工作表，请重新选择。')
  return control
}

/** Every list and selection is derived from the current workbook, including after undo. */
export function readSlicerPanelState(studio: OfficeStudio, target: SlicerPanelTarget): SlicerPanelSnapshot {
  const { workbook, sheet } = requireTarget(studio, target)
  const model = studio.getWorkbookModel(target.workbookId)!
  const allControls = workbook.getSlicers().map((control) => control.getInfo())
  const pivots = workbook.getPivotTables().map((pivot, index) => {
    const id = pivot.getPivotTableId()
    const pivotModel = model.pivotTables!.find((candidate) => candidate.id === id)!
    const source = model.sheets.find((candidate) => candidate.id === pivotModel.source.subUnitId)
    const sourceRange = pivotModel.source.range
    const labelFields = new Set(['row', 'column', 'filter'].flatMap((area) => pivot.getFieldIdsByArea(area as 'row' | 'column' | 'filter')))
    const fields = pivot.getSourceFieldsInfo().map((field) => {
      const occupied = allControls.some((control) => control.pivotTableId === id && control.sourceFieldId === field.fieldId)
      const reason = !labelFields.has(field.fieldId) ? '请先将字段放到透视表的行、列或筛选区域。'
        : occupied ? '此字段已绑定切片器或时间线。' : undefined
      const controls: SlicerType[] = ['slicer']
      if (source) {
        const heading = String(source.cells[`${sourceRange.startRow}:${field.sourceColumn}`]?.value ?? field.fieldName).trim()
        const dateHeading = /(?:date|time|日期|时间|年月|日付)/i.test(heading)
        for (let row = sourceRange.startRow + 1; row <= sourceRange.endRow; row += 1) {
          const cell = source.cells[`${row}:${field.sourceColumn}`]
          const numberFormat = classifySheetNumberFormat(cell?.style.numberFormat ?? 'General')
          const rawValue = cell?.value ?? null
          const numericSerial = typeof rawValue === 'number' && (numberFormat === 'date' || numberFormat === 'datetime' || dateHeading)
          const textualDate = typeof rawValue === 'string' && !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(rawValue.trim())
          if ((numericSerial || textualDate) && parseTimelineDate(rawValue, model.date1904)) {
            controls.push('timeline')
            break
          }
        }
      }
      return { id: field.fieldId, name: field.fieldName, available: !reason, controls, reason }
    })
    const sheetName = workbook.getSheetBySheetId(pivotModel.anchor.subUnitId)?.getSheetName() ?? '工作表'
    return { id, label: `${sheetName} · 透视表 ${index + 1}`, fields }
  })
  const controls = sheet.getSlicers().map((control): SlicerPanelControl => {
    const info = control.getInfo()
    // The pivot's filter is authoritative even if changed outside this panel.
    const field = model.pivotTables?.find((pivot) => pivot.id === info.pivotTableId)?.fields.find((candidate) => candidate.id === info.sourceFieldId && candidate.area !== 'value')
    try {
      return { info, items: control.getItems(), selectedItems: field?.selectedItems ? [...field.selectedItems] : undefined }
    } catch (error) {
      return { info, items: [], error: error instanceof Error ? error.message : String(error) }
    }
  })
  return { ...target, date1904: model.date1904, pivots, controls }
}

export async function insertPanelSlicer(studio: OfficeStudio, target: SlicerPanelTarget, input: SlicerPanelInsert): Promise<string> {
  if (input.type !== 'slicer' && input.type !== 'timeline') throw new SlicerPanelError('INVALID_SLICER_TYPE', '请选择切片器或时间线。')
  const state = readSlicerPanelState(studio, target)
  const pivot = state.pivots.find((candidate) => candidate.id === input.pivotTableId)
  if (!pivot) throw new SlicerPanelError('PIVOT_NOT_FOUND', '请选择现有的数据透视表。')
  const field = pivot.fields.find((candidate) => candidate.id === input.sourceFieldId)
  if (!field) throw new SlicerPanelError('SOURCE_FIELD_NOT_FOUND', '请选择有效的源字段。')
  if (!field.available) throw new SlicerPanelError('SOURCE_FIELD_UNAVAILABLE', field.reason ?? '所选字段暂不可用。')
  if (!field.controls.includes(input.type)) throw new SlicerPanelError('UNSUPPORTED_SOURCE_FIELD', '所选字段不支持此控件，请选择包含有效日期的字段。')
  const control = await requireTarget(studio, target).sheet.insertSlicer({ ...input, caption: input.caption?.trim() || field.name })
  return control.getId()
}

export async function applyPanelSlicerPresentation(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string, patch: SlicerPresentationPatch): Promise<void> {
  await requireControl(studio, target, slicerId).updatePresentation(patch)
}

export async function applyPanelSlicerItems(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string, selectedItems: readonly string[]): Promise<void> {
  if (!Array.isArray(selectedItems) || selectedItems.some((item) => typeof item !== 'string')) throw new SlicerPanelError('INVALID_SLICER_ITEMS', '筛选项必须是有效的字段值。')
  const control = requireControl(studio, target, slicerId)
  if (control.getInfo().type !== 'slicer') throw new SlicerPanelError('INVALID_SLICER_TYPE', '时间线请使用日期范围筛选。')
  const known = new Set(control.getItems())
  if (selectedItems.some((item) => !known.has(item))) throw new SlicerPanelError('UNKNOWN_SLICER_ITEM', '源数据已变化，请重新选择筛选项。')
  // [] means match nothing. Clearing is a separate command, never an empty-array shortcut.
  await control.setItems([...new Set(selectedItems)])
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

export async function applyPanelTimelineRange(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string, input: SlicerPanelTimelineInput): Promise<void> {
  const control = requireControl(studio, target, slicerId)
  if (control.getInfo().type !== 'timeline') throw new SlicerPanelError('INVALID_SLICER_TYPE', '切片器请使用字段值筛选。')
  const date1904 = studio.getWorkbookModel(target.workbookId)!.date1904
  await control.setTimelineRange(buildPanelTimelineRange(input, date1904))
}

export async function clearPanelSlicer(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string): Promise<void> {
  await requireControl(studio, target, slicerId).clearFilter()
}

export async function removePanelSlicer(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string): Promise<void> {
  await requireControl(studio, target, slicerId).remove()
}
