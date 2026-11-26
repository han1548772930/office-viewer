import type { OfficeStudio } from '../../api/office-studio'
import type { SlicerCommandContext } from '../../api/slicer-command-context.ts'
import type { SlicerPanelTarget } from './slicer-panel-state'

export interface SlicerPanelOwner {
  studio: OfficeStudio
  target: SlicerPanelTarget
  userId: string
  revision: number
  historyGeneration: number
}

/** A panel draft belongs to one workbench host and one content/history generation. */
export function captureSlicerPanelOwner(studio: OfficeStudio, target: SlicerPanelTarget, expectedRevision?: number): SlicerPanelOwner {
  const owner: SlicerPanelOwner = {
    studio, target: { ...target }, userId: studio.currentUserId,
    revision: expectedRevision ?? studio.getWorkbookRevision(target.workbookId),
    historyGeneration: studio.editor.history.captureSavepoint(target.workbookId).generation,
  }
  if (!slicerPanelOwnerCurrent(owner)) throw new Error('切片器目标、内容或历史已更改，请重新打开面板。')
  return owner
}

export function slicerPanelOwnerCurrent(owner: SlicerPanelOwner, published = false): boolean {
  const { studio, target, userId, revision, historyGeneration } = owner
  const workbook = studio.getWorkbookModel(target.workbookId)
  if (!workbook || studio.currentUserId !== userId || studio.activeWorkbookId !== target.workbookId
    || studio.editor.state.activeKind !== 'workbook' || workbook.activeSheetId !== target.sheetId
    || !workbook.sheets.some(sheet => sheet.id === target.sheetId)) return false
  const currentRevision = studio.getWorkbookRevision(target.workbookId)
  return (currentRevision === revision || published && currentRevision === revision + 1)
    && studio.editor.history.captureSavepoint(target.workbookId).generation === historyGeneration
}

export function slicerPanelOwnerContext(owner: SlicerPanelOwner): SlicerCommandContext {
  if (!slicerPanelOwnerCurrent(owner)) throw new Error('切片器面板内容已过期，请重新读取后再操作。')
  return { host: 'workbench', expectedRevision: owner.revision }
}
