import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { applySheetRangeThemeWorkbookPlan, planWorkbookSheetRangeThemeCommand, sheetRangeThemeCommandInput, SHEET_RANGE_THEME_COMMANDS,
  type SheetRangeThemeCommandInput, type SheetRangeThemeCommandKind } from '../features/sheet-workbench/sheet-range-theme-command.ts'

export { SHEET_RANGE_THEME_COMMANDS, type SheetRangeThemeCommandInput } from '../features/sheet-workbench/sheet-range-theme-command.ts'
export function registerSheetRangeThemeCommands(api: OfficeStudio): Disposable {
  const registrations = Object.entries(SHEET_RANGE_THEME_COMMANDS).map(([kind, id]) => api.registerCommand<SheetRangeThemeCommandInput, boolean>(id, params => {
    const input = sheetRangeThemeCommandInput(params)
    const target = resolveSheetRangeCommandTarget(api, input)
    const baseRevision = api.getWorkbookRevision(target.unitId)
    const historyGeneration = api.editor.history.captureSavepoint(target.unitId).generation, userId = api.currentUserId
    const assertCurrent = () => {
      if (api.isDisposed || api.getWorkbookModel(target.unitId) !== target.workbook || api.getWorkbookRevision(target.unitId) !== baseRevision
        || api.editor.history.captureSavepoint(target.unitId).generation !== historyGeneration || api.currentUserId !== userId) throw new Error('工作簿内容或历史已更改，请重新选择范围主题。')
      if (input.expectedRevision !== undefined && input.expectedRevision !== baseRevision
        || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== historyGeneration) throw new Error('工作簿内容或历史已更改，请重新选择范围主题。')
      if (input.host && (input.host.unitId !== target.unitId || input.host.sheetId !== target.sheetId || api.activeWorkbookId !== input.host.unitId
        || api.editor.state.activeKind !== 'workbook' || target.workbook.activeSheetId !== input.host.sheetId)) throw new Error('范围主题宿主已切换，请重新选择范围。')
    }
    assertCurrent()
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    api.assertCanEditWorksheet(target.unitId, target.sheetId)
    const plan = planWorkbookSheetRangeThemeCommand(target.workbook, target.sheetId, target.ranges, kind as SheetRangeThemeCommandKind, input)
    for (const [sheetId, ranges] of Object.entries(plan.affected)) {
      api.assertCanViewWorksheet(target.unitId, sheetId); api.assertCanEditWorksheet(target.unitId, sheetId)
      for (const range of ranges) api.assertCanEditRange(target.unitId, sheetId, range)
    }
    assertCurrent()
    if (!plan.changed) return false
    const label = kind === 'register' ? '注册范围主题' : kind === 'unregister' ? '取消范围主题注册'
      : kind === 'remove' || input.themeName === null ? '移除范围主题' : plan.catalogChanged ? '保存并应用范围主题' : '设置范围主题'
    return api.commitWorkbook(target.unitId, label, workbook => {
      assertCurrent(); applySheetRangeThemeWorkbookPlan(workbook, plan)
    }, { preserveStyleInheritance: true, recalculate: false, expectedRevision: baseRevision, expectedHistoryGeneration: historyGeneration,
      prepareView: model => { toExcelWorkbookView(model) } })
  }))
  return { dispose: () => registrations.forEach(registration => registration.dispose()) }
}
