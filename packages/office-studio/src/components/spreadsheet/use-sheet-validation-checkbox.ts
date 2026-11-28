import { computed, nextTick, onScopeDispose, ref, toRaw, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { createBlankWorkbook, type CellRange, type WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { applyExcelWorkbookView, cellKey, sheetModelCellViewAt, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { isFormulaError } from '../../features/formula/formula-types.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { createSheetCellEditSession, getSheetCellEditConflict, resolveSheetDate1904, type SheetCellEditSession } from '../../features/sheet-workbench/sheet-cell-edit-session.ts'
import type { SheetCellInputCommandResult, SheetCellInputDefaults, SheetCellInputPlan } from '../../features/sheet-workbench/sheet-cell-input-command.ts'
import { dataValidationsAt } from '../../features/sheet-validation/validation-engine.ts'
import { sheetCheckboxRuleValues, sheetCheckboxState } from '../../features/sheet-validation/validation-checkbox.ts'
import { COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, planSheetCheckboxInput, type SheetCheckboxInput } from '../../features/sheet-validation/validation-checkbox-input.ts'
import { sheetValidationCellValue, sheetValidationFormulaValue, sheetValidationSourceReader } from '../../features/sheet-validation/validation-value-read.ts'

export interface SheetCheckboxCellEntry {
  hostKey: string; key: string; row: number; column: number; ruleId: string; checked: boolean; disabled: boolean; inputOwner: string
  activate(checked?: boolean): SheetCellInputCommandResult
}
interface SheetValidationCheckboxState {
  studio?: OfficeStudio; model: ExcelWorkbookView; sheetId: string; userId: string; active: boolean; blocked: boolean
  inputSession?: SheetCellEditSession | null; region: string; currentYear: number; inputOwner: string
}
interface SheetValidationCheckboxOptions {
  state(): SheetValidationCheckboxState
  refresh(studio: OfficeStudio, unitId: string, current: () => boolean): void
  applyStandalone(plan: SheetCellInputPlan, workbook: WorkbookModel, current: () => boolean,
    published: (snapshot: ExcelWorkbookView) => void): SheetCellInputCommandResult
  focus(): void
  canRestoreFocus(): boolean
}

/** Visible controls capture their cell and rule. A stale pointer or keyboard callback cannot toggle a newer value. */
export function useSheetValidationCheckbox(options: SheetValidationCheckboxOptions) {
  const error = ref('')
  let stopped = false, submitting = false, sequence = 0
  const sourceModel = computed(() => {
    const state = options.state()
    return state.studio && state.model.id ? state.studio.getWorkbookModel(state.model.id)
      : applyExcelWorkbookView(createBlankWorkbook(), state.model, state.region)
  })
  function entryAt(row: number, column: number): SheetCheckboxCellEntry | undefined {
    const state = options.state(), workbook = sourceModel.value, sheet = workbook?.sheets.find(item => item.id === state.sheetId)
    if (!state.active || !workbook || !sheet || stopped) return undefined
    const merge = sheet.mergedRanges.find(range => row >= range.startRow && row <= range.endRow && column >= range.startCol && column <= range.endCol)
    row = merge?.startRow ?? row; column = merge?.startCol ?? column
    const range: CellRange = merge ?? { startRow: row, endRow: row, startCol: column, endCol: column }
    if (!canViewWorksheetProtection(sheet, state.userId) || !canViewProtectedRanges(sheet, [range], state.userId)) return undefined
    const rule = dataValidationsAt(sheet, row, column).find(item => item.type === 'checkbox')
    if (!rule) return undefined
    const canReadSource = sheetValidationSourceReader(state.userId)
    const values = sheetCheckboxRuleValues(rule, (formula, anchor) => sheetValidationFormulaValue(formula, workbook, sheet.id, anchor, canReadSource))
    const cell = sheet.cells[`${row}:${column}`], actual = sheetValidationCellValue(workbook, sheet, row, column, canReadSource)
    if (isFormulaError(actual)) return undefined
    const checkbox = sheetCheckboxState(values, actual, !!cell?.richText?.length, !!cell?.image)
    if (!checkbox.renderable) return undefined
    const studio = state.studio, unitId = state.model.id, key = cellKey(row, column), ruleSignature = JSON.stringify(rule)
    const revision = studio && unitId ? studio.getWorkbookRevision(unitId) : undefined
    const historyGeneration = studio && unitId ? studio.editor.history.captureSavepoint(unitId).generation : undefined
    const session = createSheetCellEditSession({ documentId: workbook.id, sheetId: sheet.id, key,
      date1904: resolveSheetDate1904(undefined, workbook), cell: sheetModelCellViewAt(workbook, sheet, row, column) },
    { region: state.region, currentYear: state.currentYear })
    const viewSource = toRaw(state.model), inputSessionId = state.inputSession?.sessionId
    function hostCurrent(): boolean {
      const current = options.state()
      return !stopped && current.active && !current.blocked && current.studio === studio && current.model.id === unitId
        && current.sheetId === sheet!.id && current.userId === state.userId && current.inputSession?.sessionId === inputSessionId
        && (!studio || !studio.isDisposed && unitId !== undefined && studio.activeWorkbookId === unitId
          && studio.editor.state.activeKind === 'workbook' && studio.getWorkbookModel(unitId)?.activeSheetId === sheet!.id)
    }
    function baselineCurrent(): boolean {
      const model = sourceModel.value, target = model?.sheets.find(item => item.id === sheet!.id)
      if (!hostCurrent() || !model || !target || !canViewWorksheetProtection(target, state.userId)
        || !canViewProtectedRanges(target, [range], state.userId) || !canEditWorksheetProtection(target, state.userId)
        || !canEditProtectedRanges(target, [range], state.userId)) return false
      if (studio && unitId ? model !== workbook || studio.getWorkbookRevision(unitId) !== revision
        || studio.editor.history.captureSavepoint(unitId).generation !== historyGeneration : toRaw(options.state().model) !== viewSource) return false
      const currentRule = dataValidationsAt(target, row, column).find(item => item.id === rule!.id)
      return JSON.stringify(currentRule) === ruleSignature && !getSheetCellEditConflict(session,
        { documentId: model.id, sheetId: target.id, key, date1904: resolveSheetDate1904(undefined, model),
          cell: sheetModelCellViewAt(model, target, row, column) })
    }
    const disabled = state.blocked || submitting || !canEditWorksheetProtection(sheet, state.userId)
      || !canEditProtectedRanges(sheet, [range], state.userId)
    return { hostKey: JSON.stringify([unitId, sheet.id, key, rule.id, state.userId]),
      key, row, column, ruleId: rule.id, checked: checkbox.checked, disabled, inputOwner: state.inputOwner,
      activate: checked => {
        const reject = (message: string): SheetCellInputCommandResult => ({ accepted: false, changed: false, issues: [], message })
        if (submitting || disabled || !baselineCurrent()) return reject('复选框目标、内容、权限或输入会话已更改，请重新操作。')
        const action = ++sequence
        const input: SheetCheckboxInput = { unitId: unitId ?? 'standalone-input', sheetId: sheet.id, row, column, ruleId: rule.id,
          checked, session, host: 'workbench', expectedRevision: revision, expectedHistoryGeneration: historyGeneration,
          expectedUserId: state.userId, expectedRuleSignature: ruleSignature }
        submitting = true
        try {
          let result: SheetCellInputCommandResult, afterModel: WorkbookModel | undefined, afterRevision = revision
          let publishedView: object | undefined
          if (studio && unitId) {
            result = studio.executeCommandSync(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, input)
            afterModel = studio.getWorkbookModel(unitId)
            afterRevision = (revision ?? 0) + (result.changed ? 1 : 0)
          } else {
            const defaults: SheetCellInputDefaults = { region: state.region, currentYear: state.currentYear, userId: state.userId }
            result = options.applyStandalone(planSheetCheckboxInput(workbook, input, defaults), workbook, baselineCurrent,
              snapshot => { publishedView = toRaw(snapshot) })
          }
          const afterCurrent = (): boolean => sequence === action && hostCurrent() && (studio && unitId
            ? afterModel === studio.getWorkbookModel(unitId) && studio.getWorkbookRevision(unitId) === afterRevision
              && studio.editor.history.captureSavepoint(unitId).generation === historyGeneration
            : toRaw(options.state().model) === (publishedView ?? viewSource))
          if (!result.accepted) { if (afterCurrent()) error.value = result.message ?? '复选框未能保存。'; return cloneModel(result) }
          if (!afterCurrent()) return cloneModel(result)
          error.value = ''
          if (studio && unitId && result.changed) options.refresh(studio, unitId, afterCurrent)
          void nextTick(() => { if (afterCurrent() && options.canRestoreFocus()) options.focus() })
          return cloneModel(result)
        } catch (cause) {
          const message = cause instanceof Error ? cause.message : '复选框未能保存。'
          if (hostCurrent() && sequence === action) error.value = message.startsWith('Command was canceled:') ? '复选框操作已取消，原值已保留。' : message
          return reject(error.value || message)
        } finally { submitting = false }
      } }
  }
  watch(() => { const state = options.state(); return [state.studio, state.model, state.model.id, state.sheetId, state.userId, state.active,
    state.studio && state.model.id ? state.studio.getWorkbookRevision(state.model.id) : undefined,
    state.studio && state.model.id ? state.studio.editor.history.captureSavepoint(state.model.id).generation : undefined] as const },
  () => { error.value = '' }, { flush: 'sync' })
  onScopeDispose(() => { stopped = true; sequence++; error.value = '' })
  return { entryAt, error, clearError: () => { error.value = '' } }
}
