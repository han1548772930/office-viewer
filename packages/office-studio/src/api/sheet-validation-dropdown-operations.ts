import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import { dataValidationsAt } from '../features/sheet-validation/validation-engine.ts'
import { SHEET_VALIDATION_DROPDOWN_OPERATIONS, sheetValidationDropdownHideInput, sheetValidationDropdownShowInput,
  sheetValidationDropdownUnit, type SheetValidationDropdownLocation, type SheetValidationDropdownRequest } from '../features/sheet-validation/validation-dropdown.ts'

export interface SheetValidationDropdownController {
  supports(request: SheetValidationDropdownRequest): boolean
  show(request: SheetValidationDropdownRequest): boolean
  hide(): boolean
  active(): SheetValidationDropdownLocation | undefined
}
interface DropdownCoordinator { controllers: Set<SheetValidationDropdownController>; active?: SheetValidationDropdownController; registered: boolean; dispatching: boolean }
const coordinators = new WeakMap<OfficeStudio, DropdownCoordinator>()
function coordinator(api: OfficeStudio): DropdownCoordinator {
  let state = coordinators.get(api)
  if (!state) { state = { controllers: new Set(), registered: false, dispatching: false }; coordinators.set(api, state) }
  return state
}

/** Workbench scopes attach controllers, while the Studio constructor owns one command registration. */
export function attachSheetValidationDropdownController(api: OfficeStudio, controller: SheetValidationDropdownController): Disposable {
  const state = coordinator(api)
  state.controllers.add(controller)
  return { dispose() { state.controllers.delete(controller); if (state.active === controller) { state.active = undefined; controller.hide() } } }
}
export function activeSheetValidationDropdown(api: OfficeStudio): SheetValidationDropdownLocation | undefined {
  const state = coordinators.get(api)
  let location = state?.active?.active()
  if (!location && state) {
    state.active = undefined
    for (const controller of [...state.controllers].reverse()) {
      location = controller.active()
      if (location) { state.active = controller; break }
    }
  }
  return location && { ...location }
}

export function registerSheetValidationDropdownOperations(api: OfficeStudio): Disposable {
  const state = coordinator(api)
  if (state.registered) throw new Error('数据验证下拉操作只能由 Studio 注册一次。')
  state.registered = true
  const registrations = [
    api.registerCommand<unknown, boolean>(SHEET_VALIDATION_DROPDOWN_OPERATIONS.show, raw => {
      if (raw === undefined || raw === null || state.dispatching) return false
      const target = sheetValidationDropdownUnit(raw), workbook = api.getWorkbookModel(target.unitId)
      const sheet = workbook?.sheets.find(candidate => candidate.id === target.sheetId)
      if (!workbook || !sheet || api.isDisposed || api.activeWorkbookId !== target.unitId
        || api.editor.state.activeKind !== 'workbook' || workbook.activeSheetId !== target.sheetId) return false
      const revision = api.getWorkbookRevision(target.unitId), generation = api.editor.history.captureSavepoint(target.unitId).generation
      const userId = api.currentUserId, signature = JSON.stringify(workbook)
      const input = sheetValidationDropdownShowInput(raw, target)
      const merge = sheet.mergedRanges.find(range => input.row >= range.startRow && input.row <= range.endRow
        && input.column >= range.startCol && input.column <= range.endCol)
      const row = merge?.startRow ?? input.row, column = merge?.startCol ?? input.column
      if (row >= (sheet.rowCount ?? 1_048_576) || column >= (sheet.columnCount ?? 16_384)) return false
      const range = merge ?? { startRow: row, endRow: row, startCol: column, endCol: column }
      const assertCurrent = (): void => {
        if (api.isDisposed || api.getWorkbookModel(target.unitId) !== workbook || api.getWorkbookRevision(target.unitId) !== revision
          || api.currentUserId !== userId || api.editor.history.captureSavepoint(target.unitId).generation !== generation
          || api.activeWorkbookId !== target.unitId || api.editor.state.activeKind !== 'workbook' || workbook.activeSheetId !== target.sheetId
          || JSON.stringify(workbook) !== signature || input.expectedRevision !== undefined && input.expectedRevision !== revision
          || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== generation
          || input.expectedUserId !== undefined && input.expectedUserId !== userId) throw new Error('下拉控件目标、内容、历史或用户已更改。')
        api.assertCanViewRange(target.unitId, target.sheetId, range)
        api.assertCanEditRange(target.unitId, target.sheetId, range)
      }
      assertCurrent()
      const rule = dataValidationsAt(sheet, row, column).find(rule => ['date', 'time', 'list', 'listMultiple'].includes(rule.type))
      if (!rule) return false
      const request: SheetValidationDropdownRequest = { ...target, row, column, ruleId: rule.id, expectedRuleSignature: JSON.stringify(rule),
        expectedRevision: revision, expectedHistoryGeneration: generation, expectedUserId: userId }
      activeSheetValidationDropdown(api)
      // An existing popup retains its scope and draft while that scope cannot accept a new target.
      const controllers = state.active ? [state.active] : [...state.controllers].reverse()
      state.dispatching = true
      try {
        for (const controller of controllers) {
          if (!state.controllers.has(controller) || !controller.supports(request)) continue
          assertCurrent()
          if (!controller.show(request)) continue
          assertCurrent()
          state.active = controller
          return true
        }
        return false
      } finally { state.dispatching = false }
    }),
    api.registerCommand<unknown, boolean>(SHEET_VALIDATION_DROPDOWN_OPERATIONS.hide, raw => {
      if (raw === undefined || raw === null || state.dispatching) return false
      const target = sheetValidationDropdownHideInput(raw)
      activeSheetValidationDropdown(api)
      const controllers = state.active ? [state.active] : [...state.controllers].reverse()
      state.dispatching = true
      try {
        for (const controller of controllers) {
          if (!state.controllers.has(controller)) continue
          const active = controller.active()
          if (!active || target.unitId !== undefined && active.unitId !== target.unitId || target.sheetId !== undefined && active.sheetId !== target.sheetId) continue
          if (!controller.hide()) return false
          if (state.active === controller) state.active = undefined
          return true
        }
        // Hiding an already closed dropdown is an idempotent operation and creates no document history.
        return !activeSheetValidationDropdown(api)
      } finally { state.dispatching = false }
    }),
  ]
  return { dispose() { registrations.forEach(registration => registration.dispose()); state.registered = false; state.active?.hide(); state.active = undefined } }
}
