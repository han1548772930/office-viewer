import { computed, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { SheetSlicerCacheOptions } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { canConfigurePanelSlicer } from './slicer-panel-state.ts'
import type { SlicerPanelOwner } from './slicer-panel-owner.ts'
import { applyPanelSlicerCacheOptions, buildPanelSlicerCacheOptionsPatch, readPanelSlicerCacheOptions,
  setPanelSlicerCacheOption, slicerCacheOptionsPanelCurrent, type SlicerCacheOptionsPanelSnapshot } from './slicer-cache-options-panel-state.ts'

export interface SlicerCacheOptionsControllerOptions {
  studio: OfficeStudio; workbookId: string; sheetId: string; slicerId: string; open: boolean
  close: (owner: SlicerPanelOwner | undefined) => void
  changed: (slicerId: string) => void
}

/** One dialog owns its draft, source and in-flight generation; the content history stays in OfficeStudio. */
export function useSlicerCacheOptions(options: SlicerCacheOptionsControllerOptions) {
  const snapshot = shallowRef<SlicerCacheOptionsPanelSnapshot>(), draft = ref<SheetSlicerCacheOptions>({}), busy = ref(false), error = ref('')
  const version = ref(0)
  let generation = 0, busyGeneration = 0, stopped = false, publishing: SlicerCacheOptionsPanelSnapshot | undefined,
    executing: SlicerCacheOptionsPanelSnapshot | undefined, published: SlicerCacheOptionsPanelSnapshot | undefined
  let publishingPatch: Partial<SheetSlicerCacheOptions> | undefined
  let subscriptions: Array<{ dispose: () => void }> = []
  const sameHost = (captured: SlicerCacheOptionsPanelSnapshot) => !stopped && options.open && options.studio === captured.owner.studio
    && options.workbookId === captured.owner.target.workbookId && options.sheetId === captured.owner.target.sheetId
    && options.slicerId === captured.slicerId && options.studio.currentUserId === captured.owner.userId
  const stale = computed(() => {
    void version.value
    return Boolean(snapshot.value && (!sameHost(snapshot.value) || !slicerCacheOptionsPanelCurrent(snapshot.value,
      executing === snapshot.value || published === snapshot.value)))
  })
  const dirty = computed(() => Boolean(snapshot.value && Object.keys(buildPanelSlicerCacheOptionsPatch(draft.value, snapshot.value.options)).length))
  const canConfigure = computed(() => {
    void version.value
    return Boolean(snapshot.value && !stale.value && canConfigurePanelSlicer(options.studio,
      { workbookId: options.workbookId, sheetId: options.sheetId }, options.slicerId))
  })
  function read(): void {
    if (!options.open || stopped || busy.value) return
    error.value = ''
    try {
      const current = readPanelSlicerCacheOptions(options.studio, { workbookId: options.workbookId, sheetId: options.sheetId }, options.slicerId)
      snapshot.value = current; draft.value = cloneModel(current.options)
    } catch (failure) { snapshot.value = undefined; error.value = failure instanceof Error ? failure.message : String(failure) }
  }
  function isOwnCommand(command: { id: string; params?: unknown } | undefined, before = false): boolean {
    const params = command?.params && typeof command.params === 'object' && !Array.isArray(command.params) ? command.params as Record<string, unknown> : undefined
    return Boolean(publishing && command?.id === 'sheet.command.patch-slicer-cache-options' && params?.slicerId === publishing.slicerId
      && params.expectedRevision === publishing.owner.revision && params.host === 'workbench'
      && (before ? params.patch === publishingPatch : executing === publishing && equivalentModelData(params.patch, publishingPatch)))
  }
  function changedContent(command?: { id: string; params?: unknown }): void {
    if (isOwnCommand(command)) published = publishing
    version.value += 1
    if (options.open && !busy.value && !dirty.value) read()
  }
  watch(() => options.studio, studio => {
    subscriptions.forEach(subscription => subscription.dispose())
    subscriptions = [
      studio.events.addEvent('BeforeCommandExecute', event => { if (event.unitId === options.workbookId && isOwnCommand(event, true)) { executing = publishing; version.value += 1 } }),
      studio.events.addEvent('CommandExecuted', event => { if (event.unitId === options.workbookId) changedContent(event) }),
      studio.events.addEvent('Undo', event => { if (event.unitId === options.workbookId) changedContent() }),
      studio.events.addEvent('Redo', event => { if (event.unitId === options.workbookId) changedContent() }),
      studio.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === options.workbookId) changedContent() }),
    ]
  }, { immediate: true })
  watch([() => options.open, () => options.studio, () => options.workbookId, () => options.sheetId, () => options.slicerId], () => {
    generation += 1; busyGeneration += 1; busy.value = false; publishing = undefined; publishingPatch = undefined; executing = undefined; published = undefined; snapshot.value = undefined; draft.value = {}; error.value = ''
    read()
  }, { immediate: true, flush: 'sync' })
  watch(() => options.studio.editor.workbook.model, () => changedContent(), { flush: 'post' })
  watch(stale, value => {
    if (value && busy.value) { generation += 1; busyGeneration += 1; busy.value = false; publishing = undefined; publishingPatch = undefined; executing = undefined; published = undefined }
  }, { flush: 'sync' })
  function set(key: keyof SheetSlicerCacheOptions, value: unknown): void {
    if (busy.value || !canConfigure.value) return
    try { draft.value = setPanelSlicerCacheOption(draft.value, key, value); error.value = '' }
    catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
  }
  function reset(): void { if (!busy.value && canConfigure.value) { draft.value = {}; error.value = '' } }
  function close(): void {
    const owner = snapshot.value?.owner
    generation += 1; busyGeneration += 1; busy.value = false; publishing = undefined; publishingPatch = undefined; executing = undefined; published = undefined
    options.close(owner)
  }
  async function apply(): Promise<boolean> {
    const current = snapshot.value
    if (!current || busy.value || !canConfigure.value || !dirty.value) return false
    const request = generation, busyToken = ++busyGeneration, patch = buildPanelSlicerCacheOptionsPatch(draft.value, current.options)
    busy.value = true; error.value = ''; publishing = current; publishingPatch = undefined; executing = undefined; published = undefined
    try {
      const accepted = await applyPanelSlicerCacheOptions(current, patch, () => request === generation && sameHost(current),
        prepared => { if (publishing === current) publishingPatch = prepared })
      if (request !== generation || !sameHost(current) || !slicerCacheOptionsPanelCurrent(current, true)) return false
      if (!accepted) { error.value = '排序设置更新未接受，请检查编辑权限和来源。'; return false }
      options.changed(current.slicerId)
      if (request === generation && sameHost(current) && slicerCacheOptionsPanelCurrent(current, true)) close()
      return true
    } catch (failure) {
      if (request === generation && sameHost(current)) error.value = failure instanceof Error ? failure.message : String(failure)
      return false
    } finally {
      if (publishing === current) { publishing = undefined; publishingPatch = undefined }
      if (executing === current) { executing = undefined; version.value += 1 }
      if (published === current) { published = undefined; version.value += 1 }
      if (!stopped && busyToken === busyGeneration) busy.value = false
    }
  }
  onScopeDispose(() => { stopped = true; generation += 1; busyGeneration += 1; subscriptions.forEach(subscription => subscription.dispose()) })
  return { snapshot, draft, busy, error, stale, dirty, canConfigure, read, set, reset, apply, close }
}
