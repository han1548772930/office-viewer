import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { applySheetAxisDimensionCommand, SHEET_AXIS_DIMENSION_COMMAND_IDS, sheetAxisDimensionActionForCommand,
  sheetAxisDimensionCommandInput, sheetAxisDimensionTargetRanges } from '../features/sheet-workbench/sheet-axis-dimensions.ts'
import { createSheetValidationTagOptionsResolver } from '../features/sheet-validation/validation-cell-options.ts'

export function registerSheetAxisDimensionCommands(api: OfficeStudio): Disposable {
  const registrations = SHEET_AXIS_DIMENSION_COMMAND_IDS.map(id => api.registerCommand<unknown, boolean>(id, params => {
    const input = sheetAxisDimensionCommandInput(params), action = sheetAxisDimensionActionForCommand(id)
    if (!action) return false
    const target = resolveSheetRangeCommandTarget(api, input)
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    if (input.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0)) throw new TypeError('工作簿版本无效。')
    if (input.expectedRevision !== undefined && api.getWorkbookRevision(target.unitId) !== input.expectedRevision) throw new Error('工作簿已更改，请重新打开尺寸设置。')
    const anchor = input.index ?? (action.axis === 'row' ? input.anchorRow : input.anchorCol)
    const ranges = sheetAxisDimensionTargetRanges(target.sheet, action.axis, target.ranges, anchor)
    if (!ranges.length) return false
    const changed = api.commitWorkbook(target.unitId, action.axis === 'row' ? action.kind === 'auto' ? '自动行高' : '设置行高'
      : action.kind === 'auto' ? '自动列宽' : '设置列宽', model => {
      const sheet = model.sheets.find(item => item.id === target.sheetId)!
      const region = api.regionalSettings.region.value
      applySheetAxisDimensionCommand(sheet, ranges, action, input,
        api.currentUserId, { measure: api.sheetTextMeasure, region, styleContext: model, userId: api.currentUserId,
          validationTagOptions: createSheetValidationTagOptionsResolver(model, sheet.id, api.currentUserId, region) }, model.date1904)
    }, { recalculate: false, prepareView: model => { toExcelWorkbookView(model) } })
    if (changed) api.publishObserverEvent('SheetDimensionsChanged', () => ({ workbookId: target.unitId, sheetId: target.sheetId,
      axis: action.axis, automatic: action.kind === 'auto', ranges: ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow,
        startColumn: range.startCol, endColumn: range.endCol, unitId: target.unitId, sheetId: target.sheetId })) }))
    return changed
  }))
  return { dispose: () => registrations.forEach(registration => registration.dispose()) }
}
