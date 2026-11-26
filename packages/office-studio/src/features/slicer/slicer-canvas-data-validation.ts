import { getRecordValueKey, normalizeRecordValue, type RecordValue } from '../../domain/record-value.ts'
import type { ISlicerDisplayMemberWindow, ISlicerSourceMetadata, SlicerDisplayMember } from './slicer-member-query.ts'

export type SlicerCanvasDataErrorCode = 'SLICER_CANVAS_INVALID_WINDOW' | 'SLICER_CANVAS_INVALID_RESULT'
  | 'SLICER_CANVAS_MEMORY_LIMIT' | 'SLICER_CANVAS_QUEUE_FULL' | 'SLICER_CANVAS_TIMEOUT' | 'SLICER_CANVAS_CANCELLED'

export class SlicerCanvasDataError extends Error {
  constructor(readonly code: SlicerCanvasDataErrorCode, message: string) { super(message); this.name = 'SlicerCanvasDataError' }
}

const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)
const invalid = () => new SlicerCanvasDataError('SLICER_CANVAS_INVALID_RESULT', '切片器画布查询返回了无效或不完整的数据。')

export function assertSlicerCanvasWindow(start: number, count: number, maximum: number): void {
  if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(count) || count <= 0
    || !Number.isSafeInteger(start + count) || count > maximum) {
    throw new SlicerCanvasDataError('SLICER_CANVAS_INVALID_WINDOW', '切片器画布成员窗口无效或超过缓存预算。')
  }
}

export function copySlicerCanvasMember(item: SlicerDisplayMember): SlicerDisplayMember {
  return { key: item.key, value: normalizeRecordValue(item.value), label: item.label, hasData: item.hasData,
    ...(item.selected !== undefined ? { selected: item.selected } : {}), ...(item.partial !== undefined ? { partial: item.partial } : {}) }
}

/** These bytes are a conservative admission estimate for retained DTOs, not a measured JS heap size. */
export function readSlicerCanvasWindow(input: unknown, start: number, count: number, budget: number): { window: ISlicerDisplayMemberWindow; bytes: number } {
  if (!object(input) || !Number.isSafeInteger(input.total) || typeof input.total !== 'number' || input.total < 0
    || !Array.isArray(input.items) || input.items.length !== Math.max(0, Math.min(count, input.total - start))) throw invalid()
  const items: SlicerDisplayMember[] = [], labels = new Set<string>()
  let bytes = 48
  for (const raw of input.items) {
    if (!object(raw) || typeof raw.key !== 'string' || typeof raw.label !== 'string' || typeof raw.hasData !== 'boolean'
      || raw.selected !== undefined && typeof raw.selected !== 'boolean' || raw.partial !== undefined && typeof raw.partial !== 'boolean'
      || raw.selected === true && raw.partial === true || labels.has(raw.label)) throw invalid()
    let value: RecordValue
    try { value = normalizeRecordValue(raw.value) }
    catch { throw invalid() }
    if (raw.key !== getRecordValueKey(value)) throw invalid()
    // Canvas is globally grouped by caption. Distinct captions can have the same typed raw key.
    labels.add(raw.label)
    bytes += 160 + raw.key.length * 2 + raw.label.length * 2
      + (value.type === 'string' || value.type === 'error' ? value.value.length * 2 : 8)
    if (!Number.isSafeInteger(bytes) || bytes > budget) {
      throw new SlicerCanvasDataError('SLICER_CANVAS_MEMORY_LIMIT', '切片器成员窗口超过画布缓存预算。')
    }
    items.push({ key: raw.key, value, label: raw.label, hasData: raw.hasData,
      ...(typeof raw.selected === 'boolean' ? { selected: raw.selected } : {}), ...(typeof raw.partial === 'boolean' ? { partial: raw.partial } : {}) })
  }
  if (bytes > budget) throw new SlicerCanvasDataError('SLICER_CANVAS_MEMORY_LIMIT', '切片器成员窗口超过画布缓存预算。')
  return { window: { total: input.total, items }, bytes }
}

export function copySlicerCanvasMetadata(value: ISlicerSourceMetadata): ISlicerSourceMetadata {
  return { fieldId: value.fieldId, name: value.name, ready: value.ready,
    ...(value.dateGroup ? { dateGroup: { ...value.dateGroup,
      ...(value.dateGroup.excelLike ? { excelLike: { ...value.dateGroup.excelLike } } : {}) } } : {}),
    ...(value.timeline ? { timeline: { ...value.timeline } } : {}) }
}

export function readSlicerCanvasMetadata(input: unknown, budget: number): { metadata: ISlicerSourceMetadata; bytes: number } {
  if (!object(input) || typeof input.fieldId !== 'string' || !input.fieldId.trim() || typeof input.name !== 'string'
    || input.ready !== true) throw invalid()
  let dateGroup: ISlicerSourceMetadata['dateGroup'], timeline: ISlicerSourceMetadata['timeline']
  if (input.dateGroup !== undefined) {
    const group = input.dateGroup
    if (!object(group) || group.type !== 3) throw invalid()
    const dateType = group.dateType
    if (dateType !== 'date' && dateType !== 'year' && dateType !== 'quarter' && dateType !== 'month') throw invalid()
    dateGroup = { type: 3, dateType }
    if (group.excelLike !== undefined) {
      const excel = group.excelLike
      if (!object(excel) || (excel.unit !== 'day' && excel.unit !== 'year' && excel.unit !== 'quarter' && excel.unit !== 'month')
        || typeof excel.autoStart !== 'boolean' || typeof excel.autoEnd !== 'boolean'
        || excel.unit !== (dateType === 'date' ? 'day' : dateType)) throw invalid()
      dateGroup.excelLike = { unit: excel.unit, autoStart: excel.autoStart, autoEnd: excel.autoEnd }
    }
  }
  if (input.timeline !== undefined) {
    const source = input.timeline
    if (!object(source) || typeof source.interactive !== 'boolean') throw invalid()
    const { minDate, maxDate } = source
    if (minDate !== undefined || maxDate !== undefined) {
      if (typeof minDate !== 'number' || !Number.isFinite(minDate) || typeof maxDate !== 'number'
        || !Number.isFinite(maxDate) || minDate > maxDate) throw invalid()
      timeline = { minDate, maxDate, interactive: source.interactive }
    } else {
      if (source.interactive) throw invalid()
      timeline = { interactive: false }
    }
  }
  const bytes = 320 + input.fieldId.length * 2 + input.name.length * 2 + (dateGroup ? 160 : 0) + (timeline ? 96 : 0)
  if (!Number.isSafeInteger(bytes) || bytes > budget) {
    throw new SlicerCanvasDataError('SLICER_CANVAS_MEMORY_LIMIT', '切片器时间线来源信息超过画布缓存预算。')
  }
  return { metadata: { fieldId: input.fieldId, name: input.name, ready: true, ...(dateGroup ? { dateGroup } : {}),
    ...(timeline ? { timeline } : {}) }, bytes }
}
