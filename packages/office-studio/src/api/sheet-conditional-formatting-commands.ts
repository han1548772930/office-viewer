import type { Disposable } from './events.ts'
import type { OfficeStudio } from './office-studio.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { applySheetConditionalFormatsCommand, planSheetConditionalFormatsCommand, SET_SHEET_CONDITIONAL_FORMATS_COMMAND,
  sheetConditionalFormatsCommandInput, type SetSheetConditionalFormatsInput } from '../features/conditional-formatting/conditional-format-command.ts'

/** A rule-list edit is one document transaction and never changes cell values. */
export function registerSheetConditionalFormattingCommands(api: OfficeStudio): Disposable {
  return api.registerCommand<SetSheetConditionalFormatsInput | undefined, boolean>(SET_SHEET_CONDITIONAL_FORMATS_COMMAND, params => {
    const input = sheetConditionalFormatsCommandInput(params), target = resolveSheetRangeCommandTarget(api, input)
    if (input.expectedRevision !== undefined && api.getWorkbookRevision(target.unitId) !== input.expectedRevision) throw new Error('工作簿已更改，请重新打开条件格式面板后编辑。')
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    api.assertCanEditWorksheet(target.unitId, target.sheetId)
    const plan = planSheetConditionalFormatsCommand(target.sheet, input.rules)
    for (const range of plan.ranges) api.assertCanEditRange(target.unitId, target.sheetId, range)
    if (!plan.changed) return false
    return api.commitWorkbook(target.unitId, '编辑条件格式', model => {
      const sheet = model.sheets.find(candidate => candidate.id === target.sheetId)
      if (!sheet) throw new Error('找不到目标工作表。')
      applySheetConditionalFormatsCommand(sheet, plan.rules, api.currentUserId)
    }, { recalculate: false })
  })
}
