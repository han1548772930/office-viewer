/** Map a visible-tab drop back to the persisted order, including hidden and inaccessible sheets. */
export function sheetTabDropOrder(sheetOrder: readonly string[], visibleSheetIds: readonly string[], sourceId: string,
  targetId: string, after: boolean): number | undefined {
  const source = visibleSheetIds.indexOf(sourceId), target = visibleSheetIds.indexOf(targetId)
  if (source < 0 || target < 0 || !sheetOrder.includes(sourceId) || !sheetOrder.includes(targetId)) return undefined
  if (source === target) return sheetOrder.indexOf(sourceId)
  const visibleIndex = target + (source < target ? after ? 0 : -1 : after ? 1 : 0)
  const id = visibleSheetIds[visibleIndex]
  return id === undefined ? undefined : sheetOrder.indexOf(id)
}
