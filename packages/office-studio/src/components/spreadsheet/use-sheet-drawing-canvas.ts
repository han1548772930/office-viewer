import { computed, onActivated, onDeactivated, onUnmounted, ref, shallowRef, watch } from 'vue'
import type { SheetDrawingAffineGeometry, SheetDrawingLeafRef, SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import type { DrawCommand } from '../../rendering/contracts.ts'
import type { GridViewport, SpreadsheetGeometry } from '../../rendering/spreadsheet/layout.ts'
import { drawingDescendants, drawingGraph, drawingGroupFrame, drawingRefKey, planDrawingGroup, planDrawingSelectionTransform, planDrawingUngroup,
  SHEET_DRAWING_GROUP_COMMAND_IDS, type SheetDrawingScene, type SheetDrawingTransaction } from '../../features/drawing-groups/index.ts'
import { drawingGroupFrameDelta } from '../../features/drawing-groups/drawing-group-geometry.ts'
import { hitSheetDrawingFrame, pickSheetDrawing, sheetDrawingGestureFrame, sheetDrawingPaneLayouts, sheetDrawingWorldPoint,
  type SheetDrawingHandle, type SheetDrawingPaneLayout } from '../../features/sheet-drawings/sheet-drawing-viewport.ts'

interface DrawingCanvasContext {
  scene(): SheetDrawingScene | undefined
  geometry(): SpreadsheetGeometry
  viewport(): GridViewport
  selection(): readonly SheetDrawingRef[]
  revision(): number | undefined
  userId(): string
  blocked(): boolean
  color(): string
  background(): string
  visible(ref: SheetDrawingLeafRef, geometry: SheetDrawingAffineGeometry): boolean
  select(ref?: SheetDrawingRef, additive?: boolean, preserve?: boolean): void
  command(id: string, input: { sheetId: string; selection: SheetDrawingRef[]; expectedSelection?: SheetDrawingRef[]; expectedRevision?: number;
    frame?: SheetDrawingAffineGeometry; delta?: SheetDrawingAffineGeometry['matrix']; groupId?: string; arrangeType?: 'front' | 'forward' | 'back' | 'backward' }): void
  focus(): void
  capture(pointerId: number): void
  release(pointerId: number): void
  render(): void
}
interface DrawingGesture {
  scene: SheetDrawingScene
  signature: string
  selection: SheetDrawingRef[]
  frame: SheetDrawingAffineGeometry
  preview: SheetDrawingAffineGeometry
  start: { x: number; y: number }
  pane: Pick<SheetDrawingPaneLayout, 'frozenX' | 'frozenY'>
  viewport: GridViewport
  handle?: SheetDrawingHandle
  pointerId: number
  revision?: number
  userId: string
  transaction?: SheetDrawingTransaction
}
export function useSheetDrawingCanvas(context: DrawingCanvasContext) {
  const gesture = shallowRef<DrawingGesture>(), error = ref('')
  const editingGroupId = ref<string>()
  let active = true, disposed = false
  const selection = computed(() => context.selection().map(ref => ({ ...ref })))
  const graph = computed(() => { const scene = context.scene(); return scene ? drawingGraph(scene) : undefined })
  const sharedControls = computed(() => selection.value.length > 1 || selection.value.some(ref => ref.kind === 'group' || ref.kind === 'shape'))
  const frame = computed(() => {
    const index = graph.value
    if (!index || !selection.value.length) return undefined
    const geometries = selection.value.flatMap(ref => {
      const geometry = ref.kind === 'group' ? index.groups.get(ref.id)?.frame : index.objects.get(drawingRefKey(ref))?.geometry
      return geometry ? [geometry] : []
    })
    return geometries.length === selection.value.length ? geometries.length === 1 ? geometries[0]! : drawingGroupFrame(geometries) : undefined
  })
  const editable = computed(() => {
    const index = graph.value
    if (!index || !selection.value.length || context.blocked()) return false
    return drawingDescendants(index, selection.value).every(ref => {
      if (ref.kind === 'group') return true
      const object = index.objects.get(drawingRefKey(ref))
      return !!object && object.editable !== false && object.available !== false && context.visible(ref, object.geometry)
    })
  })
  const transformable = computed(() => editable.value && !!graph.value && drawingDescendants(graph.value, selection.value).every(ref =>
    ref.kind === 'group' || graph.value!.objects.get(drawingRefKey(ref))?.transformable !== false))
  const layouts = computed(() => sharedControls.value && frame.value ? sheetDrawingPaneLayouts(context.geometry(), context.viewport(), gesture.value?.preview ?? frame.value) : [])
  const canGroup = computed(() => {
    const scene = context.scene()
    return !!scene && editable.value && planDrawingGroup(scene, { selection: selection.value, groupId: '__group_capability__' }).status === 'ready'
  })
  const canUngroup = computed(() => {
    const scene = context.scene()
    return !!scene && editable.value && planDrawingUngroup(scene, { selection: selection.value }).status === 'ready'
  })
  const commands = computed<DrawCommand[]>(() => layouts.value.flatMap(layout => {
    const result: DrawCommand[] = [], corners = [layout.handles.nw, layout.handles.ne, layout.handles.se, layout.handles.sw]
    for (let index = 0; index < corners.length; index++) {
      const first = corners[index]!, last = corners[(index + 1) % corners.length]!
      result.push({ kind: 'line', x1: first.x, y1: first.y, x2: last.x, y2: last.y, color: context.color(), width: 1, clip: layout.clip })
    }
    if (!transformable.value) return result
    result.push({ kind: 'line', x1: layout.handles.n.x, y1: layout.handles.n.y, x2: layout.handles.rotate.x, y2: layout.handles.rotate.y,
      color: context.color(), width: 1, clip: layout.clip })
    for (const point of Object.values(layout.handles)) result.push({ kind: 'rect', x: point.x - 3, y: point.y - 3, width: 6, height: 6,
      fill: context.background(), stroke: context.color(), strokeWidth: 1, clip: layout.clip })
    return result
  }))
  const previewGeometry = computed(() => new Map((gesture.value?.transaction?.geometryChanges ?? []).map(change => [drawingRefKey(change.ref), change.after])))
  const menuPosition = computed(() => {
    if (context.blocked() || selection.value.length < 2 && !selection.value.some(ref => ref.kind === 'group' || ref.kind === 'shape')) return undefined
    const layout = layouts.value[0], view = context.viewport()
    return layout ? { left: Math.max(44, Math.min(view.width - 270, Math.max(layout.bounds.x, layout.clip.x))),
      top: Math.max(26, Math.min(view.height - 36, Math.max(layout.bounds.y, layout.clip.y) - 34)) } : undefined
  })
  const pick = (position: { x: number; y: number }) => {
    const scene = context.scene()
    return scene ? pickSheetDrawing(scene, context.geometry(), context.viewport(), position, { editingGroupId: editingGroupId.value, visible: context.visible }) : undefined
  }
  function current(state: DrawingGesture): boolean {
    return active && !disposed && !context.blocked() && context.revision() === state.revision && context.userId() === state.userId
      && JSON.stringify(context.scene()) === state.signature && JSON.stringify(selection.value) === JSON.stringify(state.selection)
      && (['zoom', 'scrollLeft', 'scrollTop', 'width', 'height'] as const).every(key => context.viewport()[key] === state.viewport[key])
      && transformable.value
  }
  function cancel(): void {
    const pointerId = gesture.value?.pointerId
    gesture.value = undefined
    if (pointerId !== undefined) context.release(pointerId)
    context.render()
  }
  function begin(position: { x: number; y: number }, event: PointerEvent): boolean {
    if (!active || disposed || context.blocked() || event.button !== 0 && event.button !== 2) return false
    let hit = pick(position)
    if (!hit && editingGroupId.value) { editingGroupId.value = undefined; hit = pick(position) }
    const additive = event.shiftKey || event.ctrlKey || event.metaKey,
      control = sharedControls.value ? hitSheetDrawingFrame(layouts.value, position, transformable.value) : undefined
    if (control?.handle && event.button === 0) return start(control.layout, control.handle)
    if (!hit) {
      if (control && selection.value.some(ref => ref.kind === 'group') && !additive && event.button === 0) return start(control.layout)
      return false
    }
    if (additive) { context.select(hit.target, true); event.preventDefault(); event.stopPropagation(); context.focus(); return true }
    const selected = selection.value.some(ref => drawingRefKey(ref) === drawingRefKey(hit!.target))
    if (hit.target.kind !== 'group' && hit.target.kind !== 'shape' && !(sharedControls.value && selected)) return false
    context.select(hit.target, false, selected)
    event.preventDefault(); event.stopPropagation(); context.focus()
    if (event.button !== 0 || !transformable.value || !frame.value) return true
    return start(hit.layout)
    function start(pane: SheetDrawingPaneLayout, handle?: SheetDrawingHandle): boolean {
      const scene = context.scene(), before = frame.value
      if (!scene || !before || !transformable.value) return true
      error.value = ''
      gesture.value = { scene, signature: JSON.stringify(scene), selection: selection.value.map(ref => ({ ...ref })), frame: before, preview: before,
        start: sheetDrawingWorldPoint(context.viewport(), position, pane), pane, viewport: { ...context.viewport() }, handle,
        pointerId: event.pointerId, revision: context.revision(), userId: context.userId() }
      event.preventDefault(); event.stopPropagation(); context.focus(); context.capture(event.pointerId); context.render(); return true
    }
  }
  function move(position: { x: number; y: number }, event: PointerEvent): boolean {
    const state = gesture.value
    if (!state || state.pointerId !== event.pointerId) return false
    if (!current(state)) { cancel(); return true }
    try {
      const preview = sheetDrawingGestureFrame(state.frame, state.start, sheetDrawingWorldPoint(state.viewport, position, state.pane), state.handle, event),
        plan = planDrawingSelectionTransform(state.scene, { selection: state.selection, delta: drawingGroupFrameDelta(state.frame, preview) })
      if (plan.status === 'rejected') throw new Error(plan.message)
      gesture.value = { ...state, preview, transaction: plan.status === 'ready' ? plan.transaction : undefined }
      error.value = ''; context.render()
    } catch (reason) { error.value = reason instanceof Error ? reason.message : '绘图变换失败。' }
    event.preventDefault(); return true
  }
  function end(position: { x: number; y: number }, event: PointerEvent): boolean {
    const state = gesture.value
    if (!state || state.pointerId !== event.pointerId) return false
    if (event.type !== 'pointerup' || !current(state)) { cancel(); return true }
    move(position, event)
    const final = gesture.value
    cancel()
    if (final?.transaction) context.command(SHEET_DRAWING_GROUP_COMMAND_IDS.transform, { sheetId: state.scene.sheetId,
      selection: state.selection, expectedSelection: state.selection, expectedRevision: state.revision,
      delta: drawingGroupFrameDelta(state.frame, final.preview) })
    return true
  }
  function selectAt(position: { x: number; y: number }): boolean {
    const hit = pick(position)
    if (!hit || hit.target.kind !== 'group' && hit.target.kind !== 'shape' && !sharedControls.value) return false
    context.select(hit.target, false, selection.value.some(ref => drawingRefKey(ref) === drawingRefKey(hit.target))); return true
  }
  function enterGroup(position: { x: number; y: number }): boolean {
    const hit = pick(position)
    if (hit?.target.kind !== 'group') return false
    editingGroupId.value = hit.target.id
    const child = pick(position)
    context.select(child?.target); context.render(); return true
  }
  function run(id: string, extra: { arrangeType?: 'front' | 'forward' | 'back' | 'backward' } = {}): void {
    const scene = context.scene()
    if (!scene || !selection.value.length || !editable.value) return
    cancel()
    context.command(id, { sheetId: scene.sheetId, selection: selection.value.map(ref => ({ ...ref })), expectedSelection: selection.value.map(ref => ({ ...ref })),
      expectedRevision: context.revision(), ...extra })
  }
  function keydown(event: KeyboardEvent): boolean {
    if (context.blocked() || event.isComposing || event.keyCode === 229 || !sharedControls.value) return false
    if (event.key === 'Escape') {
      if (gesture.value) cancel()
      else if (editingGroupId.value) { const id = editingGroupId.value; editingGroupId.value = undefined; context.select({ kind: 'group', id }) }
      else context.select()
    } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'g') {
      if (event.shiftKey ? canUngroup.value : canGroup.value) run(event.shiftKey ? SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup : SHEET_DRAWING_GROUP_COMMAND_IDS.group)
    } else if (event.key === 'Delete' || event.key === 'Backspace') run('sheet.command.remove-sheet-drawings')
    else if (event.key.startsWith('Arrow') && transformable.value) {
      const amount = event.shiftKey ? 10 : 1, dx = event.key === 'ArrowLeft' ? -amount : event.key === 'ArrowRight' ? amount : 0,
        dy = event.key === 'ArrowUp' ? -amount : event.key === 'ArrowDown' ? amount : 0, scene = context.scene()
      if (scene) context.command(SHEET_DRAWING_GROUP_COMMAND_IDS.transform, { sheetId: scene.sheetId, selection: selection.value,
        expectedSelection: selection.value, expectedRevision: context.revision(), delta: { a: 1, b: 0, c: 0, d: 1, e: dx, f: dy } })
    } else return false
    event.preventDefault(); event.stopPropagation(); return true
  }
  watch([() => context.scene(), () => context.selection(), () => context.revision(), () => context.userId(), () => context.blocked(), () => context.viewport()], () => {
    if (gesture.value && !current(gesture.value)) cancel()
    if (editingGroupId.value && !context.scene()?.state.groups.some(group => group.id === editingGroupId.value)) editingGroupId.value = undefined
  }, { deep: true, flush: 'sync' })
  onActivated(() => { active = true })
  onDeactivated(() => { active = false; cancel(); editingGroupId.value = undefined })
  onUnmounted(() => { disposed = true; active = false; cancel() })
  return { commands, selection, sharedControls, editable, transformable, canGroup, canUngroup, menuPosition, previewGeometry, error, editingGroupId,
    hasControlAt: (position: { x: number; y: number }) => Boolean(hitSheetDrawingFrame(layouts.value, position, transformable.value)?.handle),
    pick, begin, move, end, cancel, selectAt, enterGroup, keydown, run }
}
