import type { OfficeStudio } from '../../api/office-studio.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { CHANGE_SHEET_OUTLINE_COMMAND, createSheetOutlineState, getSheetOutlineUpdate,
  type ChangeSheetOutlineInput, type OutlineWorksheet, type SheetOutlineAction, type SheetOutlineState } from './sheet-outline.ts'

export interface SheetOutlinePanelTarget { workbookId: string; sheetId: string }
export interface SheetOutlinePanelOwner {
  studio: OfficeStudio
  target: SheetOutlinePanelTarget
  userId: string
  revision: number
  historyGeneration: number
}
export interface PreparedSheetOutlineChange {
  owner: SheetOutlinePanelOwner
  input: ChangeSheetOutlineInput
  after: SheetOutlineState
}

export function captureSheetOutlinePanelOwner(studio: OfficeStudio, target: SheetOutlinePanelTarget): SheetOutlinePanelOwner {
  const owner = { studio, target: { ...target }, userId: studio.currentUserId,
    revision: studio.getWorkbookRevision(target.workbookId),
    historyGeneration: studio.editor.history.captureSavepoint(target.workbookId).generation }
  if (!sheetOutlinePanelOwnerCurrent(owner)) throw new Error('大纲目标、用户或内容已更改，请重新打开对话框。')
  return owner
}

/** After a successful own edit, exactly one content revision is accepted. */
export function sheetOutlinePanelOwnerCurrent(owner: SheetOutlinePanelOwner, committed = false): boolean {
  const { studio, target } = owner, workbook = studio.getWorkbookModel(target.workbookId)
  return Boolean(workbook && studio.currentUserId === owner.userId && studio.activeWorkbookId === target.workbookId
    && studio.editor.state.activeKind === 'workbook' && workbook.activeSheetId === target.sheetId
    && workbook.sheets.some(sheet => sheet.id === target.sheetId)
    && studio.getWorkbookRevision(target.workbookId) === owner.revision + (committed ? 1 : 0)
    && studio.editor.history.captureSavepoint(target.workbookId).generation === owner.historyGeneration)
}

export function getSheetOutlinePanelSheet(owner: SheetOutlinePanelOwner): OutlineWorksheet {
  const sheet = owner.studio.getWorkbookModel(owner.target.workbookId)?.sheets.find(sheet => sheet.id === owner.target.sheetId)
  if (!sheet) throw new Error('大纲所在工作表已关闭。')
  return { outline: sheet.outline, worksheetProtection: sheet.worksheetProtection, rangeProtectionRules: sheet.rangeProtectionRules,
    rowCount: sheet.rowCount ?? 1_048_576, columnCount: sheet.columnCount ?? 16_384 }
}

/** Preview and publication share the same generated group identity. */
export function prepareSheetOutlinePanelChange(owner: SheetOutlinePanelOwner, action: SheetOutlineAction): PreparedSheetOutlineChange | undefined {
  if (!sheetOutlinePanelOwnerCurrent(owner)) throw new Error('大纲设置已过期，请重新读取后再保存。')
  const preparedAction = cloneModel(action)
  if (preparedAction.kind === 'group' && !preparedAction.groupId) preparedAction.groupId = crypto.randomUUID()
  const after = getSheetOutlineUpdate(getSheetOutlinePanelSheet(owner), preparedAction, owner.userId)
  return after ? { owner, after: cloneModel(after), input: { unitId: owner.target.workbookId, sheetId: owner.target.sheetId,
    action: preparedAction, host: 'workbench', expectedRevision: owner.revision,
    expectedHistoryGeneration: owner.historyGeneration, expectedUserId: owner.userId } } : undefined
}

export function applySheetOutlinePanelChange(prepared: PreparedSheetOutlineChange, hostCurrent: () => boolean = () => true): boolean {
  if (!hostCurrent() || !sheetOutlinePanelOwnerCurrent(prepared.owner)) return false
  // Permissions can change independently of a draft. Recheck the same command
  // candidate immediately before entering the history-backed command bus.
  const current = getSheetOutlineUpdate(getSheetOutlinePanelSheet(prepared.owner), prepared.input.action, prepared.owner.userId)
  if (!current || !equivalentModelData(current, prepared.after) || !hostCurrent() || !sheetOutlinePanelOwnerCurrent(prepared.owner)) return false
  return prepared.owner.studio.executeCommandSync<ChangeSheetOutlineInput, boolean>(CHANGE_SHEET_OUTLINE_COMMAND, cloneModel(prepared.input))
}

/** A reentrant edit must not close a replacement dialog or restore old focus. */
export function sheetOutlinePanelChangePublished(prepared: PreparedSheetOutlineChange): boolean {
  return sheetOutlinePanelOwnerCurrent(prepared.owner, true)
    && equivalentModelData(getSheetOutlinePanelSheet(prepared.owner).outline ?? createSheetOutlineState(), prepared.after)
}
