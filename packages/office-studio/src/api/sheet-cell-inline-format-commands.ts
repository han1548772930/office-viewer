import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import { cloneModel } from '../domain/serialization.ts'
import { sheetCellInputCommandInput } from '../features/sheet-workbench/sheet-cell-input-command.ts'
import { FORMAT_SHEET_CELL_INPUT_OPERATION, FORMAT_SHEET_CELL_RICH_TEXT_COMMAND,
  planSheetCellInlineFontFormat, planSheetCellInlineStyle } from '../features/sheet-workbench/sheet-cell-inline-format-command.ts'

export function registerSheetCellInlineFormatCommands(api: OfficeStudio): Disposable {
  const defaults = () => ({ region: api.regionalSettings.region.value, currentYear: new Date().getFullYear(), userId: api.currentUserId })
  const editing = api.registerCommand(FORMAT_SHEET_CELL_INPUT_OPERATION, params => {
    if (!params || typeof params !== 'object' || !('unitId' in params) || typeof params.unitId !== 'string') throw new TypeError('文字格式工作簿标识无效。')
    const workbook = api.getWorkbookModel(params.unitId)
    if (!workbook) throw new Error('找不到文字格式目标工作簿。')
    return cloneModel(planSheetCellInlineFontFormat(workbook, params, defaults()))
  })
  const persisted = api.registerCommand(FORMAT_SHEET_CELL_RICH_TEXT_COMMAND, params => {
    if (!params || typeof params !== 'object' || !('unitId' in params) || typeof params.unitId !== 'string') throw new TypeError('文字格式工作簿标识无效。')
    const unitId = params.unitId, workbook = api.getWorkbookModel(unitId)
    if (!workbook) throw new Error('找不到文字格式目标工作簿。')
    const plan = planSheetCellInlineStyle(workbook, params, defaults())
    if (!plan.result.accepted || !plan.result.changed || !plan.cell) return cloneModel(plan.result)
    const sheetId = sheetCellInputCommandInput({ ...params, input: '' }).sheetId
    const cell = plan.cell
    const changed = api.commitWorkbook(unitId, '设置局部文字格式', model => {
      const sheet = model.sheets.find(item => item.id === sheetId)
      if (!sheet) throw new Error('文字格式目标工作表已移除。')
      sheet.cells[cell.id] = cloneModel(cell)
    })
    return { ...cloneModel(plan.result), changed }
  })
  return { dispose: () => { editing.dispose(); persisted.dispose() } }
}
