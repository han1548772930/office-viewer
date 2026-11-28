import { computed, getCurrentScope, onScopeDispose, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { createBlankWorkbook, type WorkbookModel } from '../../domain/model.ts'
import { applyExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { sheetValidationCellFeedback, type SheetValidationFeedbackReader } from '../../features/sheet-validation/validation-cell-feedback.ts'
import { sheetValidationCellTagOptions, type SheetValidationTagOptionsReader } from '../../features/sheet-validation/validation-cell-options.ts'
import type { SheetValidationPosition } from '../../features/sheet-validation/validation-engine.ts'
import { SheetValidationListReadCache } from '../../features/sheet-validation/validation-list-cache.ts'

interface SheetValidationFeedbackState { studio?: OfficeStudio; model: ExcelWorkbookView; sheetId: string; userId: string; region: string; active: boolean; blocked: boolean }

export function useSheetValidationFeedback(state: () => SheetValidationFeedbackState) {
  const listCache = new SheetValidationListReadCache()
  if (getCurrentScope()) {
    watch(() => { const current = state(); return [current.studio, current.model.id, current.sheetId, current.userId, current.region, current.active] as const },
      () => listCache.clear(), { flush: 'sync' })
    onScopeDispose(() => listCache.clear())
  }
  const source = computed(() => {
    const current = state()
    return current.studio && current.model.id ? current.studio.getWorkbookModel(current.model.id)
      : applyExcelWorkbookView(createBlankWorkbook(), current.model, current.region)
  })
  function readCurrent<T>(positions: readonly SheetValidationPosition[], derive: (workbook: WorkbookModel, sheetId: string,
    positions: readonly SheetValidationPosition[], userId: string, region: string) => ReadonlyMap<string, T>, unblocked: boolean): ReadonlyMap<string, T> {
    const current = state(), workbook = source.value, studio = current.studio, unitId = current.model.id
    if (!current.active || unblocked && current.blocked || !workbook || studio && (studio.isDisposed || studio.activeWorkbookId !== unitId
      || studio.editor.state.activeKind !== 'workbook' || workbook.activeSheetId !== current.sheetId)) return new Map()
    const revision = studio && unitId ? studio.getWorkbookRevision(unitId) : undefined
    const generation = studio && unitId ? studio.editor.history.captureSavepoint(unitId).generation : undefined
    const result = derive(workbook, current.sheetId, positions, current.userId, current.region)
    const after = state()
    return after.active && (!unblocked || !after.blocked) && after.studio === studio && after.model === current.model && after.sheetId === current.sheetId
      && after.userId === current.userId && after.region === current.region && source.value === workbook && (!studio || unitId !== undefined
        && !studio.isDisposed && studio.activeWorkbookId === unitId && studio.editor.state.activeKind === 'workbook'
        && workbook.activeSheetId === current.sheetId && studio.getWorkbookRevision(unitId) === revision
        && studio.editor.history.captureSavepoint(unitId).generation === generation) ? result : new Map()
  }
  const read: SheetValidationFeedbackReader = positions => readCurrent(positions, sheetValidationCellFeedback, true)
  // Modal/popup interaction blocks feedback and writes, but must not reset the underlying cell's candidate colors.
  const readTagOptions: SheetValidationTagOptionsReader = positions => readCurrent(positions,
    (workbook, sheetId, positions, userId, region) => sheetValidationCellTagOptions(workbook, sheetId, positions, userId, region,
      rule => listCache.read(rule, workbook, sheetId, userId, region)), false)
  return { read, readTagOptions }
}
