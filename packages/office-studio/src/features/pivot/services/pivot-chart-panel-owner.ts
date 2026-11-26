import type { OfficeStudio } from '../../../api/office-studio.ts'
import type { PivotChartCommandContext } from './pivot-chart-api-types.ts'

export interface PivotChartPanelOwner {
  studio: OfficeStudio
  unitId: string
  sheetId: string
  userId: string
  revision: number
  historyGeneration: number
}

export function capturePivotChartPanelOwner(studio: OfficeStudio, unitId: string, sheetId: string): PivotChartPanelOwner {
  const owner = { studio, unitId, sheetId, userId: studio.currentUserId, revision: studio.getWorkbookRevision(unitId),
    historyGeneration: studio.editor.history.captureSavepoint(unitId).generation }
  if (!pivotChartPanelOwnerCurrent(owner)) throw new Error('透视图面板已离开原工作簿或工作表，请重新打开。')
  return owner
}

export function pivotChartPanelOwnerCurrent(owner: PivotChartPanelOwner, published = false): boolean {
  const { studio, unitId, sheetId, userId, revision, historyGeneration } = owner
  const workbook = studio.getWorkbookModel(unitId), currentRevision = studio.getWorkbookRevision(unitId)
  return Boolean(workbook && studio.activeWorkbookId === unitId && studio.editor.state.activeKind === 'workbook'
    && workbook.activeSheetId === sheetId && workbook.sheets.some(sheet => sheet.id === sheetId) && studio.currentUserId === userId
    && (currentRevision === revision || published && currentRevision === revision + 1)
    && studio.editor.history.captureSavepoint(unitId).generation === historyGeneration)
}

export function pivotChartPanelOwnerContext(owner: PivotChartPanelOwner): PivotChartCommandContext {
  if (!pivotChartPanelOwnerCurrent(owner)) throw new Error('透视图内容、用户或历史已更改，请先重新读取。')
  return { host: 'workbench', expectedRevision: owner.revision, expectedHistoryGeneration: owner.historyGeneration, expectedUserId: owner.userId }
}
