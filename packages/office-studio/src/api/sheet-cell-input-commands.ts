import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import { cloneModel } from '../domain/serialization.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { COMMIT_SHEET_CELL_INPUT_COMMAND, planSheetCellInput, sheetCellInputCommandInput,
  type SheetCellInputCommandResult } from '../features/sheet-workbench/sheet-cell-input-command.ts'

export function registerSheetCellInputCommands(api: OfficeStudio): Disposable {
  return api.registerCommand<unknown, SheetCellInputCommandResult>(COMMIT_SHEET_CELL_INPUT_COMMAND, params => {
    const input = sheetCellInputCommandInput(params), workbook = api.getWorkbookModel(input.unitId)
    if (!workbook) throw new Error('找不到输入的目标工作簿。')
    api.assertCanViewWorksheet(input.unitId, input.sheetId)
    const plan = planSheetCellInput(workbook, input, { region: api.regionalSettings.region.value,
      currentYear: new Date().getFullYear(), userId: api.currentUserId })
    if (!plan.result.accepted || !plan.result.changed || !plan.cell) return cloneModel(plan.result)
    const cell = plan.cell
    const changed = api.commitWorkbook(input.unitId, '编辑单元格', model => {
      const sheet = model.sheets.find(item => item.id === input.sheetId)
      if (!sheet) throw new Error('输入的目标工作表已移除。')
      sheet.cells[cell.id] = cloneModel(cell)
    }, { preserveStyleInheritance: true, prepareView: model => { toExcelWorkbookView(model) } })
    if (changed) api.publishObserverEvent('SheetValueChanged', () => ({ workbookId: input.unitId, sheetId: input.sheetId,
      ranges: [{ unitId: input.unitId, sheetId: input.sheetId, startRow: plan.position.row, endRow: plan.position.row,
        startColumn: plan.position.column, endColumn: plan.position.column }] }))
    return { ...cloneModel(plan.result), changed }
  })
}
