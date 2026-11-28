import type { Disposable } from './events.ts'
import type { OfficeStudio } from './office-studio.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import type { SheetRangeCommandInput } from '../features/sheet-workbench/sheet-range-command.ts'
import type { CellRange, SheetConditionalFormatRule } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { assertConditionalFormatCommandContext, planSheetConditionalFormatsCommand, SET_SHEET_CONDITIONAL_FORMATS_COMMAND,
  sheetConditionalFormatsCommandInput, type SetSheetConditionalFormatsInput } from '../features/conditional-formatting/conditional-format-command.ts'
import { QUICK_CONDITIONAL_FORMAT_COMMANDS, quickConditionalFormatRules, type QuickConditionalFormatInput } from '../features/conditional-formatting/quick-conditional-formatting.ts'

/** A rule-list edit is one document transaction and never changes cell values. */
export function registerSheetConditionalFormattingCommands(api: OfficeStudio): Disposable {
  const captureTarget = (input: SheetRangeCommandInput & Omit<SetSheetConditionalFormatsInput, 'rules'>) => {
    assertConditionalFormatCommandContext(input)
    const target = resolveSheetRangeCommandTarget(api, input)
    const revision = api.getWorkbookRevision(target.unitId), generation = api.editor.history.captureSavepoint(target.unitId).generation
    const userId = api.currentUserId, rulesSignature = JSON.stringify(target.sheet.conditionalFormats ?? [])
    const assertCurrent = () => {
      if (api.isDisposed || api.getWorkbookModel(target.unitId) !== target.workbook
        || target.workbook.sheets.find(sheet => sheet.id === target.sheetId) !== target.sheet
        || api.getWorkbookRevision(target.unitId) !== revision || api.editor.history.captureSavepoint(target.unitId).generation !== generation
        || JSON.stringify(target.sheet.conditionalFormats ?? []) !== rulesSignature
        || api.currentUserId !== userId || input.expectedRevision !== undefined && input.expectedRevision !== revision
        || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== generation
        || input.expectedUserId !== undefined && input.expectedUserId !== userId) throw new Error('工作簿内容、历史或用户已更改，请重新打开条件格式面板后编辑。')
      if (input.host === 'workbench' && (api.activeWorkbookId !== target.unitId || api.editor.state.activeKind !== 'workbook'
        || target.workbook.activeSheetId !== target.sheetId)) throw new Error('条件格式目标已切换，请重新打开面板后编辑。')
    }
    assertCurrent()
    return { ...target, revision, generation, assertCurrent }
  }
  const commitRules = (target: ReturnType<typeof captureTarget>, rules: readonly SheetConditionalFormatRule[]): boolean => {
    const assertPermissions = (ranges: readonly CellRange[]) => {
      target.assertCurrent()
      api.assertCanViewWorksheet(target.unitId, target.sheetId)
      api.assertCanEditWorksheet(target.unitId, target.sheetId)
      for (const range of ranges) {
        api.assertCanViewRange(target.unitId, target.sheetId, range)
        api.assertCanEditRange(target.unitId, target.sheetId, range)
      }
      target.assertCurrent()
    }
    assertPermissions([])
    const plan = planSheetConditionalFormatsCommand(target.sheet, rules)
    assertPermissions(plan.ranges)
    if (!plan.changed) return false
    return api.commitWorkbook(target.unitId, '编辑条件格式', model => {
      assertPermissions(plan.ranges)
      const sheet = model.sheets.find(candidate => candidate.id === target.sheetId)
      if (!sheet) throw new Error('找不到目标工作表。')
      sheet.conditionalFormats = cloneModel(plan.rules)
    }, { recalculate: false, preserveStyleInheritance: true, expectedRevision: target.revision, expectedHistoryGeneration: target.generation,
      prepareView: model => { target.assertCurrent(); toExcelWorkbookView(model); assertPermissions(plan.ranges) } })
  }
  const registrations = [api.registerCommand<SetSheetConditionalFormatsInput | undefined, boolean>(SET_SHEET_CONDITIONAL_FORMATS_COMMAND, params => {
    const input = sheetConditionalFormatsCommandInput(params), target = captureTarget(input)
    return commitRules(target, input.rules)
  })]
  for (const id of Object.values(QUICK_CONDITIONAL_FORMAT_COMMANDS)) {
    registrations.push(api.registerCommand<QuickConditionalFormatInput | undefined, boolean>(id, input => {
      if (input === undefined && id !== QUICK_CONDITIONAL_FORMAT_COMMANDS.clearWorksheet) return false
      if (input?.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('条件格式目标工作表不一致。')
      const target = captureTarget(input ?? {})
      if (!target.ranges.length && id !== QUICK_CONDITIONAL_FORMAT_COMMANDS.clearWorksheet) return false
      const rules = quickConditionalFormatRules(target.sheet.conditionalFormats ?? [], target.ranges, id, input ?? {})
      target.assertCurrent()
      return commitRules(target, rules)
    }))
  }
  return { dispose: () => registrations.forEach(registration => registration.dispose()) }
}
