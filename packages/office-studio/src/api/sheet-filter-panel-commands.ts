import type { OfficeStudio } from './office-studio.ts'
import { bindWorkbookCommandGuard, workbookCommandGuard } from './workbook-command-guards.ts'
import { assertSheetRangesEditable } from '../features/sheet-workbench/sheet-range-command.ts'
import { sheetFilterPermissionRanges } from '../features/sheet-sort-filter/sheet-filter-column.ts'

export const SHEET_FILTER_PANEL_COMMANDS = { open: 'sheet.operation.open-filter-panel', close: 'sheet.operation.close-filter-panel' } as const
export interface SheetFilterPanelRequest {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  col?: number
  column?: number
  expectedRevision?: number
  expectedHistoryGeneration?: number
  expectedUserId?: string
}
export interface SheetFilterPanelTarget {
  unitId: string
  sheetId: string
  column: number
  revision: number
  historyGeneration: number
  userId: string
}
export interface SheetFilterPanelEditor { open(target: SheetFilterPanelTarget): boolean; close(): boolean; release(): void }
interface Owner { editor: SheetFilterPanelEditor; disposed: boolean }
const owners = new WeakMap<OfficeStudio, Map<string, Owner>>()

export function disposeSheetFilterPanelEditors(api: OfficeStudio, unitId?: string): void {
  const units = owners.get(api)
  if (!units) return
  const targets = unitId === undefined ? [...units.keys()] : [unitId]
  for (const id of targets) {
    const owner = units.get(id)
    if (!owner) continue
    units.delete(id); owner.disposed = true; owner.editor.release()
  }
  if (!units.size) owners.delete(api)
}

function request(raw: unknown): SheetFilterPanelRequest {
  if (raw === undefined) return {}
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new TypeError('筛选面板参数无效。')
  const input: SheetFilterPanelRequest = {}
  const fields = ['unitId', 'sheetId', 'subUnitId', 'col', 'column', 'expectedRevision', 'expectedHistoryGeneration', 'expectedUserId']
  if (Object.keys(raw).some(key => !fields.includes(key))) throw new TypeError('筛选面板参数包含未知字段。')
  for (const key of ['unitId', 'sheetId', 'subUnitId', 'expectedUserId'] as const) {
    const value: unknown = Reflect.get(raw, key)
    if (value !== undefined && (typeof value !== 'string' || !value.trim())) throw new TypeError('筛选面板目标 ID 无效。')
    if (typeof value === 'string') input[key] = value
  }
  for (const key of ['col', 'column', 'expectedRevision', 'expectedHistoryGeneration'] as const) {
    const value: unknown = Reflect.get(raw, key)
    if (value !== undefined && (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0)) throw new TypeError('筛选面板列或版本无效。')
    if (typeof value === 'number') input[key] = value
  }
  if (input.col !== undefined && input.column !== undefined && input.col !== input.column
    || input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('筛选面板目标参数不一致。')
  return input
}

export function registerSheetFilterPanelEditor(api: OfficeStudio, unitId: string, editor: SheetFilterPanelEditor) {
  if (!api.getWorkbookModel(unitId)) throw new RangeError('筛选面板目标工作簿不存在。')
  let units = owners.get(api)
  if (!units) { units = new Map(); owners.set(api, units) }
  const previous = units.get(unitId), owner: Owner = { editor, disposed: false }
  units.set(unitId, owner)
  if (previous) { previous.disposed = true; previous.editor.release() }
  const current = () => !owner.disposed && owners.get(api)?.get(unitId) === owner && !api.isDisposed
  return {
    execute(commandId: string, raw?: unknown): boolean {
      if (!current()) return false
      if (!Object.values(SHEET_FILTER_PANEL_COMMANDS).some(id => id === commandId)) throw new TypeError('筛选面板命令无效。')
      const input = request(raw), model = api.getWorkbookModel(unitId), revision = api.getWorkbookRevision(unitId),
        generation = api.editor.history.captureSavepoint(unitId).generation, userId = api.currentUserId
      if (input.unitId !== undefined && input.unitId !== unitId) throw new RangeError('筛选面板句柄的工作簿不匹配。')
      if (input.expectedRevision !== undefined && input.expectedRevision !== revision
        || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== generation
        || input.expectedUserId !== undefined && input.expectedUserId !== userId) throw new Error('筛选面板句柄的版本或用户已变化。')
      return api.executeCommandSync(commandId, bindWorkbookCommandGuard({ ...input, unitId, expectedRevision: revision,
        expectedHistoryGeneration: generation, expectedUserId: userId }, () => {
        if (!current() || api.getWorkbookModel(unitId) !== model || api.getWorkbookRevision(unitId) !== revision
          || api.editor.history.captureSavepoint(unitId).generation !== generation || api.currentUserId !== userId) throw new Error('筛选面板宿主已变化。')
      }))
    },
    dispose() {
      if (owner.disposed) return
      owner.disposed = true
      if (owners.get(api)?.get(unitId) === owner) owners.get(api)!.delete(unitId)
      editor.release()
    },
  }
}

export function registerSheetFilterPanelCommands(api: OfficeStudio): void {
  for (const id of Object.values(SHEET_FILTER_PANEL_COMMANDS)) api.registerCommand<unknown, boolean>(id, raw => {
    workbookCommandGuard(raw)?.()
    const input = request(raw), unitId = input.unitId ?? api.activeWorkbookId, model = api.getWorkbookModel(unitId)
    if (!model || api.activeWorkbookId !== unitId || api.editor.state.activeKind !== 'workbook') return false
    const sheetId = input.sheetId ?? input.subUnitId ?? model.activeSheetId, sheet = model.sheets.find(item => item.id === sheetId)
    const revision = api.getWorkbookRevision(unitId), historyGeneration = api.editor.history.captureSavepoint(unitId).generation
    if (input.expectedRevision !== undefined && input.expectedRevision !== revision
      || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== historyGeneration
      || input.expectedUserId !== undefined && input.expectedUserId !== api.currentUserId) throw new Error('筛选面板目标版本或用户已变化。')
    const owner = owners.get(api)?.get(unitId)
    if (!owner || owner.disposed || !sheet || sheet.id !== model.activeSheetId) return false
    if (id === SHEET_FILTER_PANEL_COMMANDS.close) return owner.editor.close()
    const column = input.column ?? input.col
    if (column === undefined || !sheet.autoFilter || column < sheet.autoFilter.range.startCol || column > sheet.autoFilter.range.endCol) return false
    assertSheetRangesEditable(sheet, sheetFilterPermissionRanges(sheet, sheet.autoFilter.range), api.currentUserId)
    return owner.editor.open({ unitId, sheetId: sheet.id, column, revision, historyGeneration, userId: api.currentUserId })
  })
}
