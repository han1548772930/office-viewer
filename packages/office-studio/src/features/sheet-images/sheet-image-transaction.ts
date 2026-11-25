import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import type { SheetImagePlan } from './sheet-image-commands.ts'
import { sheetImageClipboardBounds } from './sheet-image-clipboard.ts'
import { sheetClipboardObjectRange } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'

/** Read and edit permissions cover the complete physical drawing footprint before and after a change. */
export function sheetImagePlanAffectedRanges(workbook: WorkbookModel, sheet: SpreadsheetSheet, plan: SheetImagePlan): CellRange[] {
  const ranges = plan.affected.map(range => ({ ...range }))
  if (plan.writesImages) {
    const before = new Map((sheet.images ?? []).map(image => [image.id, image])), after = new Map((plan.images ?? []).map(image => [image.id, image]))
    for (const id of new Set([...before.keys(), ...after.keys()])) {
      const previous = before.get(id), next = after.get(id)
      if (JSON.stringify(previous) === JSON.stringify(next)
        && sheet.images?.findIndex(image => image.id === id) === plan.images?.findIndex(image => image.id === id)) continue
      for (const image of [previous, next]) if (image) ranges.push(sheetClipboardObjectRange(workbook, sheet.id, sheetImageClipboardBounds(workbook, sheet, image)))
    }
  }
  const unique = new Map(ranges.map(range => [`${range.startRow}:${range.startCol}:${range.endRow}:${range.endCol}`, range]))
  return [...unique.values()]
}
