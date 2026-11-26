import type { Disposable } from './events.ts'
import type { OfficeStudio } from './office-studio.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import type { SheetConditionalFormatRule } from '../domain/model.ts'
import { assertConditionalFormatExpectedRevision, planSheetConditionalFormatsCommand, SET_SHEET_CONDITIONAL_FORMATS_COMMAND,
  sheetConditionalFormatsCommandInput, type SetSheetConditionalFormatsInput } from '../features/conditional-formatting/conditional-format-command.ts'
import { QUICK_CONDITIONAL_FORMAT_COMMANDS, quickConditionalFormatRules, type QuickConditionalFormatInput } from '../features/conditional-formatting/quick-conditional-formatting.ts'

/** A rule-list edit is one document transaction and never changes cell values. */
export function registerSheetConditionalFormattingCommands(api: OfficeStudio): Disposable {
  const commitRules = (target: ReturnType<typeof resolveSheetRangeCommandTarget>, rules: readonly SheetConditionalFormatRule[]): boolean => {
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    api.assertCanEditWorksheet(target.unitId, target.sheetId)
    const plan = planSheetConditionalFormatsCommand(target.sheet, rules)
    for (const range of plan.ranges) api.assertCanEditRange(target.unitId, target.sheetId, range)
    if (!plan.changed) return false
    return api.commitWorkbook(target.unitId, '编辑条件格式', model => {
      model.sheets.find(candidate => candidate.id === target.sheetId)!.conditionalFormats = plan.rules
    }, { recalculate: false })
  }
  const assertRevision = (unitId: string, expectedRevision?: number) => {
    if (expectedRevision !== undefined && api.getWorkbookRevision(unitId) !== expectedRevision) throw new Error('工作簿已更改，请重新打开条件格式面板后编辑。')
  }
  const isCurrentHost = (target: ReturnType<typeof resolveSheetRangeCommandTarget>) => api.activeWorkbookId === target.unitId
    && api.editor.state.activeKind === 'workbook' && target.workbook.activeSheetId === target.sheetId
  const registrations = [api.registerCommand<SetSheetConditionalFormatsInput | undefined, boolean>(SET_SHEET_CONDITIONAL_FORMATS_COMMAND, params => {
    const input = sheetConditionalFormatsCommandInput(params), target = resolveSheetRangeCommandTarget(api, input)
    assertRevision(target.unitId, input.expectedRevision)
    if (input.host === 'workbench' && !isCurrentHost(target)) throw new Error('条件格式目标已切换，请重新打开面板后编辑。')
    return commitRules(target, input.rules)
  })]
  for (const id of Object.values(QUICK_CONDITIONAL_FORMAT_COMMANDS)) {
    registrations.push(api.registerCommand<QuickConditionalFormatInput | undefined, boolean>(id, input => {
      if (input === undefined) return false
      const target = resolveSheetRangeCommandTarget(api, input)
      if (input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('条件格式目标工作表不一致。')
      assertConditionalFormatExpectedRevision(input.expectedRevision)
      assertRevision(target.unitId, input.expectedRevision)
      if (!target.ranges.length || input.host === 'workbench' && !isCurrentHost(target)) return false
      const rules = quickConditionalFormatRules(target.sheet.conditionalFormats ?? [], target.ranges, id, input)
      return commitRules(target, rules)
    }))
  }
  return { dispose: () => registrations.forEach(registration => registration.dispose()) }
}
