import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import { cloneModel } from '../domain/serialization.ts'
import { createOfficeTranslator } from '../i18n/office-i18n.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { assertSheetTableCommandContext, createSheetTableCommandCandidate, sheetTableCreationRange, sheetTableSelectorRange,
  SHEET_TABLE_COMMANDS, type AddSheetTableCommandInput, type OpenSheetTableSelectorInput } from '../features/sheet-tables/sheet-table-create.ts'

function tableTarget(api: OfficeStudio, input: OpenSheetTableSelectorInput) {
  assertSheetTableCommandContext(input)
  const target = resolveSheetRangeCommandTarget(api, input)
  const revision = api.getWorkbookRevision(target.unitId)
  const generation = api.editor.history.captureSavepoint(target.unitId).generation
  const userId = api.currentUserId
  const assertCurrent = () => {
    if (api.isDisposed || api.getWorkbookModel(target.unitId) !== target.workbook
      || api.getWorkbookRevision(target.unitId) !== revision || api.editor.history.captureSavepoint(target.unitId).generation !== generation
      || api.currentUserId !== userId || input.expectedRevision !== undefined && input.expectedRevision !== revision
      || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== generation
      || input.expectedUserId !== undefined && input.expectedUserId !== userId) throw new Error('表格目标内容或历史已更改，请重新打开范围选择器。')
    if (input.host === 'workbench' && (api.activeWorkbookId !== target.unitId || api.editor.state.activeKind !== 'workbook'
      || target.workbook.activeSheetId !== target.sheetId)) throw new Error('表格宿主已切换，请重新打开范围选择器。')
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    api.assertCanEditWorksheet(target.unitId, target.sheetId)
  }
  assertCurrent()
  return { ...target, revision, generation, userId, assertCurrent }
}

/** Selection only opens a draft. Creation, including direct API calls, uses one guarded content transaction. */
export function registerSheetTableCommands(api: OfficeStudio): Disposable {
  const registrations = [
    api.registerCommand<OpenSheetTableSelectorInput | undefined, boolean>(SHEET_TABLE_COMMANDS.openSelector, raw => {
      const input = raw === undefined ? {} : raw, target = tableTarget(api, input)
      if (!api.events.hasListeners('SheetTableSelectorRequested')) return false
      if (api.activeWorkbookId !== target.unitId || api.editor.state.activeKind !== 'workbook'
        || target.workbook.activeSheetId !== target.sheetId) return false
      const range = sheetTableSelectorRange(target.sheet, target.ranges.at(-1))
      api.assertCanViewRange(target.unitId, target.sheetId, range)
      api.assertCanEditRange(target.unitId, target.sheetId, range)
      target.assertCurrent()
      api.events.emit('SheetTableSelectorRequested', { requestId: crypto.randomUUID(), unitId: target.unitId, sheetId: target.sheetId,
        range: cloneModel(range), expectedRevision: target.revision, expectedHistoryGeneration: target.generation, expectedUserId: target.userId })
      return true
    }),
    api.registerCommand<AddSheetTableCommandInput, boolean>(SHEET_TABLE_COMMANDS.add, input => {
      assertSheetTableCommandContext(input)
      // The add command accepts an address string as well as the public startColumn/endColumn range DTO.
      const target = tableTarget(api, { unitId: input.unitId, sheetId: input.sheetId, subUnitId: input.subUnitId,
        expectedRevision: input.expectedRevision, expectedHistoryGeneration: input.expectedHistoryGeneration, expectedUserId: input.expectedUserId, host: input.host })
      const range = sheetTableCreationRange(target.sheet, input.range)
      const assertRangeCurrent = () => {
        target.assertCurrent()
        api.assertCanViewRange(target.unitId, target.sheetId, range)
        api.assertCanEditRange(target.unitId, target.sheetId, range)
      }
      assertRangeCurrent()
      const translate = createOfficeTranslator(api.localeSettings.locale.value)
      const candidate = createSheetTableCommandCandidate(target.workbook, target.sheetId, input,
        { tablePrefix: translate('sheet.table.name-prefix', 'Table'), columnPrefix: translate('sheet.table.column-prefix', 'Column') })
      assertRangeCurrent()
      return api.commitWorkbook(target.unitId, '创建表格', draft => {
        assertRangeCurrent()
        draft.tables = cloneModel(candidate.tables)
      }, { preserveStyleInheritance: true, expectedRevision: target.revision, expectedHistoryGeneration: target.generation,
        prepareView: model => { toExcelWorkbookView(model) } })
    }),
  ]
  return { dispose() { registrations.forEach(item => item.dispose()) } }
}
