import type { OfficeStudio } from './office-studio.ts'
import type { SheetSelectionRange } from '../features/sheet-workbench/sheet-selection.ts'
import { cloneModel } from '../domain/serialization.ts'
import { CHANGE_SHEET_OUTLINE_COMMAND, assertOutlineAxis, getSheetOutlineUpdate, validateSheetOutlineState,
  type ChangeSheetOutlineInput, type SheetOutlineAction } from '../features/sheet-outline/sheet-outline.ts'
import { resolveSheetOutlineCommand, sheetOutlineCommandInput, SHEET_OUTLINE_COMMAND_IDS } from '../features/sheet-outline/sheet-outline-command.ts'

/** Direct callers and the fixed public aliases enter the same content transaction. */
export function sheetOutlineActionInput(value: unknown): SheetOutlineAction {
  if (!value || typeof value !== 'object' || Array.isArray(value) || !('kind' in value)) throw new TypeError('大纲动作必须是对象。')
  const action = value as Record<string, unknown>
  const bool = (key: string): boolean => { if (typeof action[key] !== 'boolean') throw new TypeError(`大纲参数 ${key} 必须是布尔值。`); return action[key] }
  const integer = (key: string): number => { const result = action[key]; if (typeof result !== 'number' || !Number.isSafeInteger(result) || result < 0) throw new TypeError(`大纲参数 ${key} 必须是非负安全整数。`); return result }
  const identity = (): string => { if (typeof action.groupId !== 'string' || !action.groupId.trim()) throw new TypeError('大纲动作需要有效 groupId。'); return action.groupId }
  if (action.kind === 'summary-settings') return { kind: action.kind, summaryAbove: bool('summaryAbove'), summaryLeft: bool('summaryLeft') }
  assertOutlineAxis(action.axis)
  const axis = action.axis
  switch (action.kind) {
    case 'summary-position': return { kind: action.kind, axis, atStart: bool('atStart') }
    case 'set-level': return { kind: action.kind, axis, level: integer('level') }
    case 'toggle-group': case 'remove-group': return { kind: action.kind, axis, groupId: identity() }
    case 'set-collapsed': return { kind: action.kind, axis, groupId: identity(), collapsed: bool('collapsed') }
    case 'group': case 'toggle': case 'remove': case 'collapse': case 'expand': case 'clear': {
      const range = { kind: action.kind, axis, start: integer('start'), end: integer('end') }
      return action.kind === 'group' && action.groupId !== undefined ? { ...range, kind: 'group', groupId: identity() } : range
    }
    default: throw new TypeError('未知的大纲动作。')
  }
}

export function registerSheetOutlineCommands(studio: OfficeStudio): void {
  studio.registerCommand<ChangeSheetOutlineInput, boolean>(CHANGE_SHEET_OUTLINE_COMMAND, value => {
    if (!value || typeof value !== 'object') throw new TypeError('大纲命令参数无效。')
    const owner = sheetOutlineCommandInput(value)
    if (!owner.unitId || !owner.sheetId) throw new TypeError('大纲命令需要工作簿和工作表。')
    const action = sheetOutlineActionInput(value.action)
    if (action.kind === 'group' && !action.groupId) action.groupId = crypto.randomUUID()
    const unitId = owner.unitId, sheetId = owner.sheetId
    const workbook = studio.getWorkbookModel(unitId), sheet = workbook?.sheets.find(sheet => sheet.id === sheetId)
    if (!workbook || !sheet) throw new Error('大纲所属工作簿或工作表已关闭。')
    const revision = studio.getWorkbookRevision(unitId), generation = studio.editor.history.captureSavepoint(unitId).generation, userId = studio.currentUserId
    const assertCurrent = (): void => {
      if (studio.isDisposed || studio.getWorkbookModel(unitId) !== workbook || studio.getWorkbookRevision(unitId) !== revision
        || studio.currentUserId !== userId || studio.editor.history.captureSavepoint(unitId).generation !== generation
        || owner.expectedRevision !== undefined && owner.expectedRevision !== revision
        || owner.expectedHistoryGeneration !== undefined && owner.expectedHistoryGeneration !== generation
        || owner.expectedUserId !== undefined && owner.expectedUserId !== userId) throw new Error('大纲目标内容、用户或历史已改变。')
      if (owner.host === 'workbench' && (studio.activeWorkbookId !== unitId || studio.editor.state.activeKind !== 'workbook'
        || workbook.activeSheetId !== sheetId)) throw new Error('大纲工作台宿主已改变。')
      studio.assertCanViewWorksheet(unitId, sheetId)
    }
    assertCurrent()
    const outline = getSheetOutlineUpdate(sheet, action, userId)
    if (!outline) return false
    return studio.commitWorkbook(unitId, '编辑工作表分组', candidate => {
      assertCurrent()
      const target = candidate.sheets.find(sheet => sheet.id === sheetId)
      if (!target) throw new Error('大纲目标工作表已不存在。')
      target.outline = cloneModel(outline)
    }, { expectedRevision: revision, expectedHistoryGeneration: generation, prepareView: candidate => {
      assertCurrent()
      // Recheck the original footprint after recalculation and all preparation callbacks.
      getSheetOutlineUpdate(sheet, action, userId)
      const target = candidate.sheets.find(sheet => sheet.id === sheetId)
      if (!target?.outline) throw new Error('大纲候选资源已不存在。')
      validateSheetOutlineState(target.outline, target)
    } })
  })
  for (const id of SHEET_OUTLINE_COMMAND_IDS) studio.registerCommand<unknown, boolean>(id, value => {
    const params = sheetOutlineCommandInput(value), unitId = params.unitId ?? studio.activeWorkbookId
    const workbook = studio.getWorkbookModel(unitId)
    if (!workbook) throw new Error('大纲所属工作簿已关闭。')
    const sheetId = params.subUnitId ?? params.sheetId ?? workbook.activeSheetId, sheet = workbook.sheets.find(sheet => sheet.id === sheetId)
    if (!sheet) throw new Error('大纲所属工作表已关闭。')
    studio.assertCanViewWorksheet(unitId, sheetId)
    const revision = studio.getWorkbookRevision(unitId), generation = studio.editor.history.captureSavepoint(unitId).generation, userId = studio.currentUserId
    if (params.expectedRevision !== undefined && params.expectedRevision !== revision
      || params.expectedHistoryGeneration !== undefined && params.expectedHistoryGeneration !== generation
      || params.expectedUserId !== undefined && params.expectedUserId !== userId) throw new Error('大纲目标内容、用户或历史已改变。')
    const selection = studio.getSelections(unitId, sheetId)
    const ranges: readonly SheetSelectionRange[] = params.ranges ?? selection.map(range => ({ startRow: range.startRow, endRow: range.endRow,
      startCol: range.startColumn, endCol: range.endColumn, ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) }))
    const result = resolveSheetOutlineCommand(id, params, { unitId, sheetId, sheet, userId: studio.currentUserId, selections: ranges })
    if (!result) return false
    if (result.kind === 'change') return studio.executeCommandSync(CHANGE_SHEET_OUTLINE_COMMAND, result.input)
    if (studio.activeWorkbookId !== unitId || studio.editor.state.activeKind !== 'workbook' || workbook.activeSheetId !== sheetId
      || !studio.events.hasListeners('SheetOutlineActionDialogRequested')) return false
    studio.publishObserverEvent('SheetOutlineActionDialogRequested', () => ({ unitId, sheetId, ranges: cloneModel([...ranges]), operation: result.operation,
      expectedRevision: revision, expectedHistoryGeneration: generation, expectedUserId: userId }), () => !studio.isDisposed
      && studio.getWorkbookModel(unitId) === workbook && studio.getWorkbookRevision(unitId) === revision
      && studio.editor.history.captureSavepoint(unitId).generation === generation && studio.currentUserId === userId
      && studio.activeWorkbookId === unitId && studio.editor.state.activeKind === 'workbook' && workbook.activeSheetId === sheetId)
    return true
  })
}
