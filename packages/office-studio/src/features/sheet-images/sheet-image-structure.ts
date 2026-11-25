import type { SheetFloatingImage, SheetImageAnchor, SheetImageBounds } from '../../domain/sheet-images.ts'
import { assertSheetAxisEdit, sheetAxisInsertionEdit, sheetAxisMaximum, type SheetAxisEdit } from '../../domain/sheet-axis-structure.ts'
import { assertSheetCellShiftEdit, type SheetCellShiftEdit } from '../../domain/sheet-cell-shift.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { createSheetSelectionAxisIndexes, type SheetSelectionSheet } from '../sheet-workbench/sheet-selection.ts'
import { SHEET_IMAGE_MAX_COUNT } from './sheet-image-model.ts'
import { sheetImageFrame, sheetImageWithFrame, type SheetImageAxes, type SheetImageAxis } from './sheet-image-placement.ts'

type SheetImageDimensions = Partial<SheetSelectionSheet> & { autoFilter?: { filteredRows?: readonly number[] } }
export interface SheetImageStructureGeometry { before: SheetImageAxes; after: SheetImageAxes }
function imageAxes(dimensions: SheetImageDimensions): SheetImageAxes {
  return createSheetSelectionAxisIndexes({ ...dimensions, rowCount: dimensions.rowCount ?? 1_048_576,
    columnCount: dimensions.columnCount ?? 16_384 }, dimensions.autoFilter?.filteredRows ?? [])
}
function axisCount(axis: 'row' | 'column', dimensions: SheetImageDimensions): number {
  const count = (axis === 'row' ? dimensions.rowCount : dimensions.columnCount) ?? sheetAxisMaximum(axis)
  if (!Number.isSafeInteger(count) || count < 1 || count > sheetAxisMaximum(axis)) throw new RangeError('图片所在工作表容量无效。')
  return count
}
const axisOf = (axes: SheetImageAxes, axis: 'row' | 'column'): SheetImageAxis => axis === 'row' ? axes.rows : axes.columns
const markerIndex = (marker: SheetImageAnchor, axis: 'row' | 'column'): number => axis === 'row' ? marker.row : marker.column
function changedMarker(marker: SheetImageAnchor, axis: 'row' | 'column', index: number, offset?: number): SheetImageAnchor {
  return axis === 'row' ? { ...marker, row: index, ...(offset === undefined ? {} : { offsetY: offset }) }
    : { ...marker, column: index, ...(offset === undefined ? {} : { offsetX: offset }) }
}
/** Matches drawing-reference.ts: retain signed offsets, collapse removed markers, and clip to the surviving last cell. */
function referenceMarker(marker: SheetImageAnchor, edit: Pick<SheetAxisEdit, 'axis' | 'start' | 'count'> & { kind: 'insert' | 'delete' },
  dimensions: SheetImageDimensions, axes: SheetImageAxes): SheetImageAnchor {
  const index = markerIndex(marker, edit.axis), end = edit.start + edit.count - 1
  let mapped = index, offset: number | undefined
  if (edit.kind === 'insert' && index >= edit.start) mapped += edit.count
  else if (edit.kind === 'delete' && index > end) mapped -= edit.count
  else if (edit.kind === 'delete' && index >= edit.start) { mapped = edit.start; offset = 0 }
  const limit = axisCount(edit.axis, dimensions)
  if (mapped >= limit) { mapped = limit - 1; offset = axisOf(axes, edit.axis).size(mapped) }
  return changedMarker(marker, edit.axis, mapped, offset)
}
function shiftMarker(marker: SheetImageAnchor, axis: 'row' | 'column', delta: number): SheetImageAnchor {
  const index = markerIndex(marker, axis) + delta
  if (!Number.isSafeInteger(index) || index < 0 || index >= sheetAxisMaximum(axis)) throw new RangeError('移动或复制会把图片锚点移出 Office 行列上限。')
  return changedMarker(marker, axis, index)
}
function frameInterval(frame: SheetImageBounds, axis: 'row' | 'column'): { start: number; end: number } {
  return axis === 'row' ? { start: frame.top, end: frame.top + frame.height } : { start: frame.left, end: frame.left + frame.width }
}
function containedInAxis(frame: SheetImageBounds, axis: 'row' | 'column', index: SheetImageAxis, start: number, end: number): boolean {
  const interval = frameInterval(frame, axis)
  return interval.start >= index.offset(start) && interval.end <= index.offset(end + 1)
}
function transformReference(image: SheetFloatingImage, edit: Pick<SheetAxisEdit, 'axis' | 'start' | 'count'> & { kind: 'insert' | 'delete' },
  dimensions: SheetImageDimensions, after: SheetImageAxes): SheetFloatingImage | undefined {
  if (image.anchorType === '2') return cloneModel(image)
  if (image.anchorType === '1' && !image.endAnchor) throw new TypeError('随单元格缩放的图片缺少结束锚点。')
  if (edit.kind === 'delete' && image.anchorType === '1' && markerIndex(image.anchor, edit.axis) >= edit.start
    && markerIndex(image.endAnchor!, edit.axis) < edit.start + edit.count) return undefined
  const next = cloneModel(image)
  next.anchor = referenceMarker(image.anchor, edit, dimensions, after)
  if (image.anchorType === '1') next.endAnchor = referenceMarker(image.endAnchor!, edit, dimensions, after)
  // size/crop retain their authored basis. Both extents project from current axes, including hidden cells and rotation.
  return next
}
/** All modes share the before/after axes with Canvas and clipboard; content history owns the entire detached result. */
export function transformSheetImagesForAxisEdit(images: readonly SheetFloatingImage[], edit: SheetAxisEdit,
  dimensions: SheetImageDimensions = {}, geometry?: SheetImageStructureGeometry): SheetFloatingImage[] {
  assertSheetAxisEdit(edit)
  const after = geometry?.after ?? imageAxes(dimensions), before = geometry?.before ?? after, insertion = sheetAxisInsertionEdit(edit)
  const result: SheetFloatingImage[] = []
  for (const image of images) {
    if (edit.kind === 'move') {
      const frame = sheetImageFrame(image, before), next = cloneModel(image)
      // Univer carries only wholly contained Both drawings; Position/None stay on their host.
      if (image.anchorType === '1' && containedInAxis(frame, edit.axis, axisOf(before, edit.axis), edit.start, edit.start + edit.count - 1)) {
        if (!image.endAnchor) throw new TypeError('移动图片缺少结束锚点。')
        const target = edit.destination > edit.start ? edit.destination - edit.count : edit.destination, delta = target - edit.start
        next.anchor = shiftMarker(image.anchor, edit.axis, delta); next.endAnchor = shiftMarker(image.endAnchor, edit.axis, delta)
      }
      result.push(next); continue
    }
    const next = insertion ? transformReference(image, insertion, dimensions, after) : cloneModel(image)
    if (next) result.push(next)
    if (edit.kind !== 'copy') continue
    const frame = sheetImageFrame(image, before)
    if (!containedInAxis(frame, edit.axis, axisOf(before, edit.axis), edit.start, edit.start + edit.count - 1)) continue
    let copy = cloneModel(image); copy.id = crypto.randomUUID()
    const delta = edit.destination - edit.start
    if (copy.anchorType === '2') {
      const pixels = axisOf(after, edit.axis).offset(edit.destination) - axisOf(before, edit.axis).offset(edit.start)
      copy = sheetImageWithFrame(copy, { ...frame, ...(edit.axis === 'row' ? { top: frame.top + pixels } : { left: frame.left + pixels }) }, after)
    } else {
      copy.anchor = shiftMarker(image.anchor, edit.axis, delta)
      if (image.anchorType === '1' && image.endAnchor) copy.endAnchor = shiftMarker(image.endAnchor, edit.axis, delta)
    }
    result.push(copy)
  }
  if (result.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('复制后的图片数量超过 1000 张。')
  return result
}
export function transformSheetImagesForCellShift(images: readonly SheetFloatingImage[], edit: SheetCellShiftEdit,
  dimensions: SheetImageDimensions = {}, geometry?: SheetImageStructureGeometry): SheetFloatingImage[] {
  assertSheetCellShiftEdit(edit)
  const after = geometry?.after ?? imageAxes(dimensions), before = geometry?.before ?? after, perpendicular = edit.axis === 'row' ? 'column' : 'row'
  return images.flatMap(image => {
    if (image.anchorType === '2') return [cloneModel(image)]
    const frame = sheetImageFrame(image, before)
    // A drawing crossing the perpendicular band remains a single object and cannot be split by a cell edit.
    if (!containedInAxis(frame, perpendicular, axisOf(before, perpendicular), edit.bandStart, edit.bandEnd)) return [cloneModel(image)]
    if (edit.kind === 'delete' && containedInAxis(frame, edit.axis, axisOf(before, edit.axis), edit.start, edit.start + edit.count - 1)) return []
    const next = transformReference(image, edit, dimensions, after)
    return next ? [next] : []
  })
}
