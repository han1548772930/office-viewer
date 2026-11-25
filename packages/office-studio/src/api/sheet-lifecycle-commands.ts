import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { COPY_WORKSHEET_COMMAND, planWorksheetCopy, type CopyWorksheetCommandInput } from '../features/sheet-workbench/sheet-worksheet-copy.ts'
import { WORKSHEET_LIFECYCLE_COMMANDS, planWorksheetLifecycle, type WorksheetLifecycleCommandInput } from '../features/sheet-workbench/sheet-worksheet-lifecycle.ts'

export function copyWorksheetCommandInput(params: unknown): CopyWorksheetCommandInput {
  if (params === undefined) return {}
  if (!params || typeof params !== 'object' || Array.isArray(params)) throw new TypeError('复制工作表参数无效。')
  const input = params as Record<string, unknown>
  for (const key of ['unitId', 'sheetId', 'subUnitId', 'targetSheetId', 'targetSheetName']) {
    if (input[key] !== undefined && (typeof input[key] !== 'string' || !String(input[key]).trim())) throw new TypeError(`复制工作表参数 ${key} 无效。`)
  }
  if (input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('源工作表标识不一致。')
  for (const key of ['expectedRevision', 'expectedHistoryGeneration']) if (input[key] !== undefined
    && (!Number.isSafeInteger(input[key]) || Number(input[key]) < (key === 'expectedRevision' ? 0 : 1))) throw new TypeError('工作簿内容或历史版本无效。')
  if (input.host !== undefined) {
    const host = input.host
    if (!host || typeof host !== 'object' || Array.isArray(host) || !('unitId' in host) || !('sheetId' in host)
      || typeof host.unitId !== 'string' || !host.unitId.trim() || typeof host.sheetId !== 'string' || !host.sheetId.trim()) throw new TypeError('复制工作表宿主无效。')
  }
  return input as CopyWorksheetCommandInput
}

export function worksheetLifecycleCommandInput(params: unknown): WorksheetLifecycleCommandInput {
  const input = copyWorksheetCommandInput(params) as WorksheetLifecycleCommandInput
  for (const key of ['name', 'value'] as const) if (input[key] !== undefined && typeof input[key] !== 'string') throw new TypeError(`工作表参数 ${key} 无效。`)
  for (const key of ['index', 'order'] as const) if (input[key] !== undefined && (!Number.isSafeInteger(input[key]) || input[key]! < 0)) throw new TypeError(`工作表参数 ${key} 无效。`)
  if (input.hidden !== undefined && input.hidden !== 0 && input.hidden !== 1 && input.hidden !== 2) throw new TypeError('工作表隐藏状态无效。')
  if (input.sheet !== undefined && (!input.sheet || typeof input.sheet !== 'object' || Array.isArray(input.sheet))) throw new TypeError('新工作表参数无效。')
  return input
}

export function registerSheetLifecycleCommands(api: OfficeStudio): Disposable {
  const registrations: Disposable[] = [api.registerCommand<unknown, boolean>(COPY_WORKSHEET_COMMAND, params => {
    const input = copyWorksheetCommandInput(params), unitId = input.unitId ?? api.activeWorkbookId
    const workbook = api.getWorkbookModel(unitId)
    if (!workbook) throw new Error('工作簿不存在。')
    const sheetId = input.sheetId ?? input.subUnitId ?? workbook.activeSheetId
    if (input.host && (api.editor.state.activeKind !== 'workbook' || api.activeWorkbookId !== input.host.unitId
      || unitId !== input.host.unitId || workbook.activeSheetId !== input.host.sheetId)) return false
    api.assertCanViewWorksheet(unitId, sheetId)
    const revision = api.getWorkbookRevision(unitId), history = api.editor.history.captureSavepoint(unitId)
    if (input.expectedRevision !== undefined && input.expectedRevision !== revision
      || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== history.generation) throw new Error('工作簿已更改，请重新复制。')
    const plan = planWorksheetCopy(workbook, sheetId, { userId: api.currentUserId,
      targetSheetId: input.targetSheetId, targetSheetName: input.targetSheetName })
    return api.commitWorkbook(unitId, '复制工作表', model => Object.assign(model, plan.workbook), {
      expectedRevision: revision, expectedHistoryGeneration: history.generation,
      prepareView: model => { toExcelWorkbookView(model) },
    })
  })]
  const labels = { insert: '插入工作表', remove: '删除工作表', rename: '重命名工作表', hidden: '隐藏工作表', show: '显示工作表', order: '移动工作表', tabColor: '设置标签颜色' }
  for (const [kind, id] of Object.entries(WORKSHEET_LIFECYCLE_COMMANDS)) registrations.push(api.registerCommand<unknown, boolean>(id, params => {
    const input = worksheetLifecycleCommandInput(params), unitId = input.unitId ?? api.activeWorkbookId
    const workbook = api.getWorkbookModel(unitId)
    if (!workbook) throw new Error('工作簿不存在。')
    const sheetId = input.sheetId ?? input.subUnitId ?? workbook.activeSheetId
    const currentHost = () => !input.host || api.editor.state.activeKind === 'workbook' && api.activeWorkbookId === input.host.unitId
      && unitId === input.host.unitId && api.getWorkbookModel(unitId)?.activeSheetId === input.host.sheetId
    if (!currentHost()) return false
    const revision = api.getWorkbookRevision(unitId), history = api.editor.history.captureSavepoint(unitId), userId = api.currentUserId
    if (input.expectedRevision !== undefined && input.expectedRevision !== revision
      || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== history.generation) throw new Error('工作簿已更改，请重新执行。')
    const plan = planWorksheetLifecycle(workbook, id, input, userId)
    if (!plan.changed) return false
    const oldName = workbook.sheets.find(sheet => sheet.id === sheetId)?.name
    if (id === WORKSHEET_LIFECYCLE_COMMANDS.rename && oldName !== undefined && input.name !== undefined) {
      const event = { workbookId: unitId, sheetId, oldName, newName: input.name, cancel: false }
      api.events.emit('BeforeSheetNameChange', event)
      if (event.cancel) return false
      if (!currentHost() || api.currentUserId !== userId || api.getWorkbookRevision(unitId) !== revision
        || api.editor.history.captureSavepoint(unitId).generation !== history.generation) throw new Error('重命名前的工作簿上下文已更改，请重新执行。')
    }
    const changed = api.commitWorkbook(unitId, labels[kind as keyof typeof labels], model => Object.assign(model, plan.workbook), {
      expectedRevision: revision, expectedHistoryGeneration: history.generation,
      ...(id === WORKSHEET_LIFECYCLE_COMMANDS.hidden || id === WORKSHEET_LIFECYCLE_COMMANDS.show || id === WORKSHEET_LIFECYCLE_COMMANDS.tabColor ? { recalculate: false } : {}),
      prepareView: model => { toExcelWorkbookView(model) },
    })
    if (changed && id === WORKSHEET_LIFECYCLE_COMMANDS.rename && oldName !== undefined && input.name !== undefined) {
      api.events.emit('SheetNameChanged', { workbookId: unitId, sheetId, oldName, newName: input.name })
    }
    return changed
  }))
  registrations.push(api.registerCommand<unknown, boolean>('sheet.command.delete-sheet', params => api.executeCommandSync(WORKSHEET_LIFECYCLE_COMMANDS.remove, params)))
  return { dispose: () => { for (const registration of registrations.reverse()) registration.dispose() } }
}
