import { computed, onScopeDispose, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { bindSheetDrawingSelection, getSheetDrawingSelection, hasSheetDrawingSelection, setSheetDrawingSelection } from '../../api/sheet-drawing-selection.ts'
import type { SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import { drawingGraph, drawingRefKey, normalizeDrawingRef, type SheetDrawingScene } from '../../features/drawing-groups/index.ts'

interface DrawingSelectionContext {
  studio(): OfficeStudio | undefined
  unitId(): string | undefined
  sheetId(): string
  active(): boolean
  scene(): SheetDrawingScene | undefined
  nativeSelection(): readonly SheetDrawingRef[]
  applyNative(refs: readonly SheetDrawingRef[]): void
}

/** A selected group and one of its descendants cannot both receive the same gesture. */
export function normalizeSheetDrawingUiSelection(scene: SheetDrawingScene | undefined, input: readonly SheetDrawingRef[]): SheetDrawingRef[] {
  if (!scene) return []
  const graph = drawingGraph(scene), selected = new Map<string, SheetDrawingRef>()
  for (const value of input) {
    const ref = normalizeDrawingRef(value), key = drawingRefKey(ref)
    if (ref.kind === 'group' ? graph.groups.has(ref.id) : graph.objects.has(key)) selected.set(key, ref)
  }
  return [...selected.values()].filter(ref => {
    let parent = graph.parents.get(drawingRefKey(ref)) ?? null
    while (parent !== null) {
      if (selected.has(drawingRefKey({ kind: 'group', id: parent }))) return false
      parent = graph.parents.get(drawingRefKey({ kind: 'group', id: parent })) ?? null
    }
    return true
  })
}

/** This broker owns only view selection. Content history remains in OfficeStudio. */
export function useSheetDrawingSelection(context: DrawingSelectionContext) {
  const refs = shallowRef<SheetDrawingRef[]>([]), scene = computed(() => context.scene())
  const sessions = new Map<string, SheetDrawingRef[]>()
  let synchronizing = false, initialized = false, binding: { dispose(): void } | undefined, previousStudio = context.studio()
  const key = () => `${context.unitId() ?? 'standalone'}\u0000${context.sheetId()}`
  function set(input: readonly SheetDrawingRef[], publish = true): void {
    const next = normalizeSheetDrawingUiSelection(scene.value, input)
    if (JSON.stringify(refs.value) !== JSON.stringify(next)) refs.value = next
    sessions.set(key(), next.map(ref => ({ ...ref })))
    synchronizing = true
    try { context.applyNative(next) } finally { synchronizing = false }
    const api = context.studio(), unitId = context.unitId()
    if (publish && context.active() && api && unitId && !api.isDisposed) setSheetDrawingSelection(api, unitId, context.sheetId(), next)
  }
  function select(ref?: SheetDrawingRef, additive = false, preserve = false): void {
    if (!context.active()) return
    if (!ref) { set([]); return }
    const target = normalizeDrawingRef(ref), targetKey = drawingRefKey(target), selected = refs.value.some(value => drawingRefKey(value) === targetKey)
    if (additive) {
      if (selected) { set(refs.value.filter(value => drawingRefKey(value) !== targetKey)); return }
      const graph = scene.value ? drawingGraph(scene.value) : undefined
      const related = (candidate: SheetDrawingRef): boolean => {
        if (!graph) return false
        for (const [child, ancestor] of [[candidate, target], [target, candidate]] as const) {
          if (ancestor.kind !== 'group') continue
          let parent = graph.parents.get(drawingRefKey(child)) ?? null
          while (parent !== null) {
            if (parent === ancestor.id) return true
            parent = graph.parents.get(drawingRefKey({ kind: 'group', id: parent })) ?? null
          }
        }
        return false
      }
      set([...refs.value.filter(value => !related(value)), target])
    } else if (preserve && selected) set([...refs.value.filter(value => drawingRefKey(value) !== targetKey), target])
    else set([target])
  }
  watch([() => context.studio(), () => context.unitId(), () => context.sheetId(), () => context.active()], ([api, unitId], _old, onCleanup) => {
    binding?.dispose(); binding = undefined
    if (api !== previousStudio) { sessions.clear(); previousStudio = api }
    const saved = api && unitId && hasSheetDrawingSelection(api, unitId, context.sheetId()) ? getSheetDrawingSelection(api, unitId, context.sheetId()) : undefined
    set(saved ?? sessions.get(key()) ?? (initialized ? [] : context.nativeSelection()), false)
    initialized = true
    if (context.active() && api && unitId && !api.isDisposed) {
      binding = bindSheetDrawingSelection(api, unitId, () => ({ sheetId: context.sheetId(), refs: refs.value }), selection => {
        if (selection.sheetId === context.sheetId() && context.active()) set(selection.refs, false)
      })
      setSheetDrawingSelection(api, unitId, context.sheetId(), refs.value)
    }
    const owner = binding
    onCleanup(() => { owner?.dispose(); if (binding === owner) binding = undefined })
  }, { immediate: true, flush: 'sync' })
  watch(() => context.nativeSelection(), native => {
    if (!synchronizing && context.active()) set(native)
  }, { deep: true, flush: 'sync' })
  watch(scene, () => set(refs.value), { flush: 'sync' })
  onScopeDispose(() => { binding?.dispose(); sessions.clear() })
  return { refs, scene, select, set }
}
