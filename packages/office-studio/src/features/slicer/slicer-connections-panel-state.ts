import { FOffice } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { SheetSlicerConnections, SlicerConnectionCandidate } from '../../domain/model'
import { canConfigurePanelSlicer, SlicerPanelError, type SlicerPanelTarget } from './slicer-panel-state'
import { captureSlicerPanelOwner, slicerPanelOwnerContext, slicerPanelOwnerCurrent, type SlicerPanelOwner } from './slicer-panel-owner'

export interface SlicerConnectionOption extends SlicerConnectionCandidate { connected: boolean }
export interface SlicerConnectionsPanelSnapshot {
  owner: SlicerPanelOwner
  slicerId: string
  cacheId: string
  options: SlicerConnectionOption[]
  connections: SheetSlicerConnections
}

/** Length/quoting are preserved; report IDs containing separators cannot collide. */
export function slicerConnectionKey(candidate: Pick<SlicerConnectionCandidate, 'kind' | 'subUnitId' | 'id'>): string {
  return JSON.stringify([candidate.kind, candidate.subUnitId, candidate.id])
}

export function slicerConnectionKeys(connections: SheetSlicerConnections): string[] {
  return [
    ...connections.pivotTables.map(({ subUnitId, pivotTableId }) => slicerConnectionKey({ kind: 'pivotTable', subUnitId, id: pivotTableId })),
    ...(connections.ownedPivotCharts ?? []).map(({ subUnitId, pivotChartId }) => slicerConnectionKey({ kind: 'ownedPivotChart', subUnitId, id: pivotChartId })),
  ]
}

export function selectedSlicerConnections(options: readonly SlicerConnectionOption[], selectedKeys: readonly string[]): SheetSlicerConnections {
  if (!Array.isArray(selectedKeys) || selectedKeys.some(key => typeof key !== 'string')) throw new SlicerPanelError('INVALID_CONNECTION_SELECTION', '请选择有效的报表连接。')
  if (options.some(option => (option.kind !== 'pivotTable' && option.kind !== 'ownedPivotChart') || typeof option.subUnitId !== 'string' || !option.subUnitId
    || typeof option.id !== 'string' || !option.id)) throw new SlicerPanelError('INVALID_CONNECTION_TARGET', '报表连接类型或标识无效，请重新读取列表。')
  const known = new Map(options.map(option => [slicerConnectionKey(option), option])), selected = [...new Set(selectedKeys)]
  if (selected.some(key => !known.has(key))) throw new SlicerPanelError('STALE_CONNECTION_SELECTION', '报表来源已更改，请重新读取连接列表。')
  const chosen = selected.map(key => known.get(key)!)
  return {
    pivotTables: chosen.filter(option => option.kind === 'pivotTable').map(({ subUnitId, id }) => ({ subUnitId, pivotTableId: id })),
    ownedPivotCharts: chosen.filter(option => option.kind === 'ownedPivotChart').map(({ subUnitId, id }) => ({ subUnitId, pivotChartId: id })),
  }
}

export async function readPanelSlicerConnections(studio: OfficeStudio, target: SlicerPanelTarget, slicerId: string): Promise<SlicerConnectionsPanelSnapshot> {
  const owner = captureSlicerPanelOwner(studio, target), control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getSlicerControl(slicerId)
  if (!control) throw new SlicerPanelError('SLICER_NOT_FOUND', '所选切片器已删除，请重新选择。')
  const connections = control.getConnections(), cacheId = control.getCacheId()
  if (!connections || !cacheId) throw new SlicerPanelError('UNSUPPORTED_CONNECTIONS', '表格切片器使用本表筛选；请选择透视表或透视图切片器。')
  const candidates = await control.getConnectionCandidates()
  if (!slicerPanelOwnerCurrent(owner) || control.getCacheId() !== cacheId) throw new SlicerPanelError('STALE_CONNECTION_SOURCE', '切片器来源或工作台已更改，请重新读取连接。')
  const connected = new Set(slicerConnectionKeys(connections))
  const available = new Set(candidates.map(slicerConnectionKey))
  if ([...connected].some(key => !available.has(key))) throw new SlicerPanelError('UNAVAILABLE_CONNECTED_REPORT', '已有连接的报表不可访问或来源已失效，请恢复报表权限或来源后重新读取。')
  return { owner, slicerId, cacheId,
    connections: { pivotTables: connections.pivotTables.map(ref => ({ ...ref })), ownedPivotCharts: connections.ownedPivotCharts?.map(ref => ({ ...ref })) },
    options: candidates.map(candidate => ({ ...candidate, connected: connected.has(slicerConnectionKey(candidate)) })),
  }
}

export function slicerConnectionsPanelCurrent(snapshot: SlicerConnectionsPanelSnapshot, published = false): boolean {
  if (!slicerPanelOwnerCurrent(snapshot.owner, published)) return false
  const { studio, target } = snapshot.owner
  try {
    return FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getSlicerControl(snapshot.slicerId)?.getCacheId() === snapshot.cacheId
  } catch { return false }
}

export async function applyPanelSlicerConnections(snapshot: SlicerConnectionsPanelSnapshot, selectedKeys: readonly string[]): Promise<boolean> {
  if (!slicerConnectionsPanelCurrent(snapshot)) throw new SlicerPanelError('STALE_CONNECTION_SOURCE', '连接列表已过期，请重新读取后再保存。')
  const { studio, target } = snapshot.owner
  if (!canConfigurePanelSlicer(studio, target, snapshot.slicerId)) return false
  const control = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)?.getSlicerControl(snapshot.slicerId)
  if (!control) throw new SlicerPanelError('SLICER_NOT_FOUND', '所选切片器已删除，请重新选择。')
  return control.setConnections(selectedSlicerConnections(snapshot.options, selectedKeys), slicerPanelOwnerContext(snapshot.owner))
}
