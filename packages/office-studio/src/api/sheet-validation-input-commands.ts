import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import { cloneModel } from '../domain/serialization.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import type { SheetCellInputCommandResult, SheetCellInputDefaults, SheetCellInputPlan } from '../features/sheet-workbench/sheet-cell-input-command.ts'
import type { WorkbookModel } from '../domain/model.ts'
import { COMMIT_SHEET_DATE_INPUT_COMMAND, planSheetDateInput, sheetDateInputCommandInput } from '../features/sheet-validation/validation-date-input.ts'
import { COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, planSheetValidationListInput, sheetValidationListInput } from '../features/sheet-validation/validation-list-input.ts'
import { COMMIT_SHEET_CHECKBOX_INPUT_COMMAND, planSheetCheckboxInput, sheetCheckboxInput } from '../features/sheet-validation/validation-checkbox-input.ts'
import { resolveSheetValidationWidgetTarget, type SheetValidationWidgetTargetInput } from '../features/sheet-validation/validation-widget-input.ts'

function commitWidget<Input extends SheetValidationWidgetTargetInput>(api: OfficeStudio, input: Input, label: string,
  planInput: (workbook: WorkbookModel, input: Input, defaults: SheetCellInputDefaults) => SheetCellInputPlan): SheetCellInputCommandResult {
  const workbook = api.getWorkbookModel(input.unitId), userId = api.currentUserId
  if (!workbook) throw new Error('找不到选择控件的目标工作簿。')
  const sheet = workbook.sheets.find(candidate => candidate.id === input.sheetId), revision = api.getWorkbookRevision(input.unitId)
  const generation = api.editor.history.captureSavepoint(input.unitId).generation, signature = JSON.stringify(workbook)
  const defaults = { region: api.regionalSettings.region.value, currentYear: new Date().getFullYear(), userId }
  const assertCurrent = (): void => {
    if (api.isDisposed || api.getWorkbookModel(input.unitId) !== workbook || workbook.sheets.find(candidate => candidate.id === input.sheetId) !== sheet
      || api.currentUserId !== userId || api.getWorkbookRevision(input.unitId) !== revision
      || api.editor.history.captureSavepoint(input.unitId).generation !== generation || JSON.stringify(workbook) !== signature
      || input.expectedRevision !== undefined && input.expectedRevision !== revision
      || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== generation
      || input.expectedUserId !== undefined && input.expectedUserId !== userId
      || input.host === 'workbench' && (api.activeWorkbookId !== input.unitId || api.editor.state.activeKind !== 'workbook'
        || workbook.activeSheetId !== input.sheetId)) throw new Error('选择控件的目标、内容、历史或用户已更改，请重新打开。')
    const target = resolveSheetValidationWidgetTarget(workbook, input, defaults)
    if ('result' in target) return
    api.assertCanViewRange(input.unitId, input.sheetId, target.range)
    api.assertCanEditRange(input.unitId, input.sheetId, target.range)
  }
  assertCurrent()
  const plan = planInput(workbook, input, defaults)
  assertCurrent()
  if (!plan.result.accepted || !plan.result.changed || !plan.cell) return cloneModel(plan.result)
  const cell = plan.cell
  const changed = api.commitWorkbook(input.unitId, label, model => {
    assertCurrent()
    const candidateSheet = model.sheets.find(candidate => candidate.id === input.sheetId)
    if (!candidateSheet) throw new Error('选择控件的工作表已移除。')
    candidateSheet.cells[cell.id] = cloneModel(cell)
  }, { expectedRevision: revision, expectedHistoryGeneration: generation, preserveStyleInheritance: true,
    prepareView: model => { assertCurrent(); toExcelWorkbookView(model) } })
  if (changed) api.publishObserverEvent('SheetValueChanged', () => ({ workbookId: input.unitId, sheetId: input.sheetId,
    ranges: [{ unitId: input.unitId, sheetId: input.sheetId, startRow: plan.position.row, endRow: plan.position.row,
      startColumn: plan.position.column, endColumn: plan.position.column }] }))
  return { ...cloneModel(plan.result), changed }
}
export function registerSheetValidationInputCommands(api: OfficeStudio): Disposable {
  const subscriptions = [
    api.registerCommand<unknown, SheetCellInputCommandResult>(COMMIT_SHEET_DATE_INPUT_COMMAND,
      value => commitWidget(api, sheetDateInputCommandInput(value), '选择日期或时间', planSheetDateInput)),
    api.registerCommand<unknown, SheetCellInputCommandResult>(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND,
      value => commitWidget(api, sheetValidationListInput(value), '选择下拉列表项', planSheetValidationListInput)),
    api.registerCommand<unknown, SheetCellInputCommandResult>(COMMIT_SHEET_CHECKBOX_INPUT_COMMAND,
      value => commitWidget(api, sheetCheckboxInput(value), '切换复选框', planSheetCheckboxInput)),
  ]
  return { dispose: () => subscriptions.forEach(subscription => subscription.dispose()) }
}
