import type { SlicerModel, WorkbookModel } from '../../domain/model.ts'
import { planDrawingArrange, SheetDrawingError } from '../drawing-groups/index.ts'
import { applySheetDrawingTransaction, createSheetDrawingScene } from '../sheet-drawings/sheet-drawing-host.ts'
import type { SlicerArrange } from './slicer-drawing.ts'

/** Slicer settings and mixed-object menus use the same authoritative Sheet stacking order. */
export function arrangeSlicer(workbook: WorkbookModel, slicer: SlicerModel, action: SlicerArrange): boolean {
  const plan = planDrawingArrange(createSheetDrawingScene(workbook, slicer.hostSheetId), {
    selection: [{ kind: slicer.type, id: slicer.id }], arrangeType: action,
  })
  if (plan.status === 'rejected') throw new SheetDrawingError(plan.reason, plan.message, plan.ref)
  if (plan.status === 'noop') return false
  applySheetDrawingTransaction(workbook, slicer.hostSheetId, plan.transaction)
  return true
}
