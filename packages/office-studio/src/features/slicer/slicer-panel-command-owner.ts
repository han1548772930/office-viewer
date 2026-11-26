import type { SlicerPanelOwner } from './slicer-panel-owner.ts'

/** Facade calls may wait before dispatch. Recheck the captured UI lease at the two real command boundaries. */
export async function withSlicerPanelCommandOwner<T>(owner: SlicerPanelOwner, slicerId: string, commandId: string,
  payload: object, current: () => boolean, action: () => Promise<T>, payloadKey: 'patch' | 'range' | 'condition' | 'params' = 'patch'): Promise<T> {
  let executing = false
  const before = owner.studio.events.addEvent('BeforeCommandExecute', event => {
    const params = event.params && typeof event.params === 'object' && !Array.isArray(event.params) ? event.params as Record<string, unknown> : undefined
    if (event.unitId !== owner.target.workbookId || event.id !== commandId || params?.slicerId !== slicerId
      || (payloadKey === 'params' ? params : params?.[payloadKey]) !== payload) return
    executing = true
    if (!current()) event.cancel = true
  })
  const prepare = owner.studio.events.addEvent('BeforeSlicerChange', event => {
    if (executing && event.unitId === owner.target.workbookId && event.slicerId === slicerId && event.commandId === commandId && !current()) event.cancel = true
  })
  try { return await action() }
  finally { before.dispose(); prepare.dispose() }
}
