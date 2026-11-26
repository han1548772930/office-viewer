import type { WorkbookModel } from '../../domain/model.ts'
import type { SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import { drawingDescendants, drawingGeometryEqual, drawingGraph, drawingGroupFrame, drawingRefKey } from '../drawing-groups/index.ts'
import { createSheetDrawingScene, reconcileSheetDrawingState } from './sheet-drawing-host.ts'
import { reconcileWorkbookImageGeometryChanges } from './sheet-drawing-image-geometry.ts'
import { reconcileWorkbookSlicerGeometryChanges } from './sheet-drawing-slicer-geometry.ts'

const hasDrawings = (workbook: WorkbookModel, sheetId: string) => {
  const sheet = workbook.sheets.find(sheet => sheet.id === sheetId)
  return !!(sheet?.drawingState || sheet?.images?.length || sheet?.shapes?.length || workbook.charts?.some(chart => chart.sheetId === sheetId)
    || workbook.slicers?.some(slicer => slicer.hostSheetId === sheetId))
}
/** Legacy snapshots retain the application's prior chart/slicer/image paint order until explicitly arranged. */
export function initializeWorkbookDrawingStates(workbook: WorkbookModel): void {
  for (const sheet of workbook.sheets) if (hasDrawings(workbook, sheet.id)) {
    const scene = createSheetDrawingScene(workbook, sheet.id)
    if (!sheet.drawingState) {
      const priority = { chart: 0, shape: 1, slicer: 2, timeline: 2, image: 3, group: 4 }
      scene.state.rootOrder.sort((left, right) => priority[left.kind] - priority[right.kind])
    }
    sheet.drawingState = scene.state
  }
}

/** A content candidate knows which native owners were actually removed; missing media alone is never a deletion. */
export function reconcileWorkbookDrawingStates(before: WorkbookModel, after: WorkbookModel): void {
  reconcileWorkbookImageGeometryChanges(before, after)
  reconcileWorkbookSlicerGeometryChanges(before, after)
  for (const sheet of after.sheets) {
    if (!hasDrawings(before, sheet.id) && !hasDrawings(after, sheet.id)) continue
    const previousSheet = before.sheets.find(candidate => candidate.id === sheet.id)
    if (!previousSheet) continue // A worksheet-copy planner must publish its complete identity mapping first.
    const oldScene = createSheetDrawingScene(before, sheet.id), oldObjects = new Map(oldScene.objects.map(object => [drawingRefKey(object.ref), object]))
    const probe = { ...after, sheets: after.sheets.map(candidate => candidate.id === sheet.id ? { ...candidate, drawingState: undefined } : candidate) },
      newObjects = createSheetDrawingScene(probe, sheet.id).objects,
      newKeys = new Set(newObjects.map(object => drawingRefKey(object.ref))), removedRefs = oldScene.objects.filter(object => !newKeys.has(drawingRefKey(object.ref))).map(object => object.ref)
    sheet.drawingState ??= oldScene.state
    // A feature may already have reconciled its deletion. Only declare refs still held by its candidate state.
    const stateRefs = new Set([...(sheet.drawingState.rootOrder ?? []), ...sheet.drawingState.groups.flatMap(group => group.members)].map(drawingRefKey))
    reconcileSheetDrawingState(after, sheet.id, { removedRefs: removedRefs.filter(ref => stateRefs.has(drawingRefKey(ref))) })
    const graph = drawingGraph(createSheetDrawingScene(after, sheet.id)), changedAncestors = new Set<string>()
    for (const object of newObjects) {
      const previous = oldObjects.get(drawingRefKey(object.ref))
      if (!previous || object.available === false || drawingGeometryEqual(previous.geometry, object.geometry)) continue
      let parent = graph.parents.get(drawingRefKey(object.ref)) ?? null
      while (parent !== null) { changedAncestors.add(parent); parent = graph.parents.get(drawingRefKey({ kind: 'group', id: parent })) ?? null }
    }
    const baseline = JSON.stringify(sheet.drawingState), groups = new Map(sheet.drawingState.groups.map(group => [group.id, group]))
    for (const ref of drawingDescendants(graph, graph.scene.state.rootOrder).reverse()) {
      if (ref.kind !== 'group' || !changedAncestors.has(ref.id)) continue
      const group = groups.get(ref.id)!
      const geometryFor = (member: SheetDrawingRef) => member.kind === 'group' ? groups.get(member.id)!.frame : graph.objects.get(drawingRefKey(member))!.geometry
      const fitted = drawingGroupFrame(group.members.map(geometryFor), group.frame)
      if (!drawingGeometryEqual(group.frame, fitted)) group.frame = fitted
    }
    if (JSON.stringify(sheet.drawingState) !== baseline) sheet.drawingState.revision++
  }
  initializeWorkbookDrawingStates(after)
}
