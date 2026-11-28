import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import type { OfficeComputeClient } from '../compute/client.ts'
import { ComputeError } from '../compute/errors.ts'
import { cloneModel } from '../domain/serialization.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { planSheetCheckboxRuleCells, type SheetCheckboxRuleCellPlan } from '../features/sheet-validation/validation-checkbox-rule-cells.ts'
import { planSheetCheckboxRuleCellsAsync, type SheetDataValidationAsyncOptions } from '../features/sheet-validation/validation-checkbox-rule-cells-async.ts'
import { planSheetDataValidationCommand, SHEET_DATA_VALIDATION_COMMANDS, SHEET_DATA_VALIDATION_ASYNC_COMMANDS, sheetDataValidationCommandPayload,
  sheetDataValidationCommandRecord, sheetDataValidationCommandTarget, type SheetDataValidationAction } from '../features/sheet-validation/validation-rule-command.ts'

function prepareValidationCommand(api: OfficeStudio, action: SheetDataValidationAction, value: unknown) {
  const source = sheetDataValidationCommandRecord(value), target = sheetDataValidationCommandTarget(source)
  const workbook = api.getWorkbookModel(target.unitId), sheet = workbook?.sheets.find(candidate => candidate.id === target.sheetId)
  if (!workbook || !sheet) throw new Error('找不到数据验证的目标工作簿或工作表。')
  const revision = api.getWorkbookRevision(target.unitId), generation = api.editor.history.captureSavepoint(target.unitId).generation
  const userId = api.currentUserId, snapshot = JSON.stringify(workbook)
  function assertCurrent(fullSnapshot = true): void {
    if (api.isDisposed || api.getWorkbookModel(target.unitId) !== workbook || workbook!.sheets.find(candidate => candidate.id === target.sheetId) !== sheet
      || api.getWorkbookRevision(target.unitId) !== revision || api.currentUserId !== userId
      || api.editor.history.captureSavepoint(target.unitId).generation !== generation || fullSnapshot && JSON.stringify(workbook) !== snapshot
      || target.expectedRevision !== undefined && target.expectedRevision !== revision
      || target.expectedHistoryGeneration !== undefined && target.expectedHistoryGeneration !== generation
      || target.expectedUserId !== undefined && target.expectedUserId !== userId
      || target.expectedRulesSignature !== undefined && target.expectedRulesSignature !== JSON.stringify(sheet!.dataValidations ?? [])
      || target.host === 'workbench' && (api.activeWorkbookId !== target.unitId || api.editor.state.activeKind !== 'workbook'
        || workbook!.activeSheetId !== target.sheetId)) throw new Error('数据验证目标、内容、历史或用户已更改，请重新打开面板。')
  }
  assertCurrent()
  const input = sheetDataValidationCommandPayload(action, source, target)
  assertCurrent()
  const plan = planSheetDataValidationCommand(sheet, input)
  function assertAllowed(fullSnapshot = true): void {
    assertCurrent(fullSnapshot)
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    api.assertCanEditWorksheet(target.unitId, target.sheetId)
    for (const range of plan.ranges) { api.assertCanViewRange(target.unitId, target.sheetId, range); api.assertCanEditRange(target.unitId, target.sheetId, range) }
    assertCurrent(fullSnapshot)
  }
  assertAllowed()
  // apply/replace appends its canonical rule. Capability queries above never allocate cell initialization plans.
  const appliedRule = input.action === 'set' || input.action === 'replace' ? plan.rules.at(-1) : undefined
  const labels: Record<SheetDataValidationAction, string> = { set: '设置数据验证', replace: '更新数据验证规则', remove: '删除数据验证规则', clear: '清除数据验证', clearTypes: '清除快速数据验证' }
  return { source, workbook, sheet, appliedRule, userId, revision, changed: plan.changed, assertAllowed,
    previous: appliedRule ? sheet.dataValidations?.find(rule => rule.id === appliedRule.id) : undefined,
    commit(cellPlan: SheetCheckboxRuleCellPlan): boolean {
      assertAllowed()
      if (!plan.changed) return false
      const changed = api.commitWorkbook(target.unitId, labels[action], model => {
        assertAllowed()
        const targetSheet = model.sheets.find(candidate => candidate.id === target.sheetId)
        if (!targetSheet) throw new Error('数据验证目标工作表已移除。')
        targetSheet.dataValidations = cloneModel(plan.rules)
        for (const cell of cellPlan.cells) targetSheet.cells[cell.id] = cloneModel(cell)
      }, { expectedRevision: revision, expectedHistoryGeneration: generation, preserveStyleInheritance: true,
        prepareView: model => { assertAllowed(); toExcelWorkbookView(model) } })
      if (changed && cellPlan.cells.length) api.publishObserverEvent('SheetValueChanged', () => ({ workbookId: target.unitId, sheetId: target.sheetId,
        ranges: cellPlan.ranges.map(range => ({ unitId: target.unitId, sheetId: target.sheetId, startRow: range.startRow, endRow: range.endRow,
          startColumn: range.startCol, endColumn: range.endCol })) }))
      return changed
    },
  }
}
function executeValidationCommand(api: OfficeStudio, action: SheetDataValidationAction, value: unknown): boolean {
  const prepared = prepareValidationCommand(api, action, value)
  if (!prepared.changed) return false
  const cellPlan = prepared.appliedRule?.type === 'checkbox' ? planSheetCheckboxRuleCells(prepared.workbook, prepared.sheet,
    prepared.appliedRule, prepared.previous, prepared.userId, api.regionalSettings.region.value) : { cells: [], ranges: [] }
  return prepared.commit(cellPlan)
}
async function executeValidationCommandAsync(api: OfficeStudio, compute: Pick<OfficeComputeClient, 'submit'>, action: 'set' | 'replace', value: unknown): Promise<boolean> {
  const prepared = prepareValidationCommand(api, action, value)
  if (!prepared.changed) return false
  const source = prepared.source
  const signal = source.signal, onProgress = source.onProgress, maxMaterializationBytes = source.maxMaterializationBytes
  if (signal !== undefined && (!signal || typeof signal !== 'object' || typeof (signal as AbortSignal).aborted !== 'boolean'
    || typeof (signal as AbortSignal).addEventListener !== 'function' || typeof (signal as AbortSignal).removeEventListener !== 'function')) {
    throw new TypeError('数据验证取消参数必须是 AbortSignal。')
  }
  if (onProgress !== undefined && typeof onProgress !== 'function') throw new TypeError('数据验证进度回调必须是函数。')
  const options = { signal, onProgress, maxMaterializationBytes } as SheetDataValidationAsyncOptions
  const cellPlan = prepared.appliedRule?.type === 'checkbox' ? await planSheetCheckboxRuleCellsAsync(prepared.workbook, prepared.sheet,
    prepared.appliedRule, prepared.previous, prepared.userId, api.regionalSettings.region.value,
    { ...options, compute, modelRevision: prepared.revision, assertCurrent: () => prepared.assertAllowed(false) }) : { cells: [], ranges: [] }
  if (options.signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '数据验证操作已取消。')
  return prepared.commit(cellPlan)
}
export function registerSheetDataValidationCommands(api: OfficeStudio, compute: Pick<OfficeComputeClient, 'submit'>): Disposable {
  const subscriptions = Object.entries(SHEET_DATA_VALIDATION_COMMANDS).map(([action, id]) => api.registerCommand<unknown, boolean>(id,
    value => executeValidationCommand(api, action as SheetDataValidationAction, value)))
  for (const [action, id] of Object.entries(SHEET_DATA_VALIDATION_ASYNC_COMMANDS)) subscriptions.push(api.registerCommand<unknown, boolean>(id,
    value => executeValidationCommandAsync(api, compute, action as 'set' | 'replace', value)))
  return { dispose: () => subscriptions.forEach(subscription => subscription.dispose()) }
}
