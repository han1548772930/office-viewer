import { computed, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { canConfigurePanelSlicer } from './slicer-panel-state.ts'
import type { SlicerPanelOwner } from './slicer-panel-owner.ts'
import { applyPanelTimelineDateFilter, createTimelineDateFilterDraft, fillTimelinePanelSourceDays, readPanelTimelineDateBounds,
  readPanelTimelineDateFilter, timelinePanelDateFilterDirty, timelinePanelDateFilterError, timelineRangePanelCurrent,
  type TimelineDateFilterDraft, type TimelinePanelDateBounds, type TimelinePanelPreparedCommand, type TimelineRangePanelSnapshot } from './timeline-range-panel-state.ts'

export interface TimelineDateFilterControllerOptions {
  studio: OfficeStudio; workbookId: string; sheetId: string; timelineId: string; open: boolean
  close: (owner: SlicerPanelOwner | undefined) => void; changed: (timelineId: string) => void
}
const emptyDraft = (): TimelineDateFilterDraft => ({ mode: 'none', range: {}, condition: { operator: 'equal' } })

/** Retained author state, inclusive source query and an in-flight edit have independent lifetimes. */
export function useTimelineDateFilter(options: TimelineDateFilterControllerOptions) {
  const snapshot = shallowRef<TimelineRangePanelSnapshot>(), draft = ref<TimelineDateFilterDraft>(emptyDraft())
  const bounds = shallowRef<TimelinePanelDateBounds>(), boundsLoading = ref(false), boundsError = ref(''), error = ref(''), busy = ref(false), composing = ref(false)
  const version = ref(0), readSession = ref(0)
  let generation = 0, boundsGeneration = 0, busyGeneration = 0, stopped = false,
    publishing: TimelineRangePanelSnapshot | undefined, executing: TimelineRangePanelSnapshot | undefined, published: TimelineRangePanelSnapshot | undefined,
    publishingCommand: TimelinePanelPreparedCommand | undefined
  let subscriptions: Array<{ dispose: () => void }> = []
  const sameHost = (captured: TimelineRangePanelSnapshot) => !stopped && options.open && options.studio === captured.owner.studio
    && options.workbookId === captured.owner.target.workbookId && options.sheetId === captured.owner.target.sheetId
    && options.timelineId === captured.timelineId && options.studio.currentUserId === captured.owner.userId
  const stale = computed(() => {
    void version.value
    return Boolean(snapshot.value && (!sameHost(snapshot.value) || !timelineRangePanelCurrent(snapshot.value,
      executing === snapshot.value || published === snapshot.value)))
  })
  const dirty = computed(() => Boolean(snapshot.value && timelinePanelDateFilterDirty(draft.value, snapshot.value.state)))
  const validation = computed(() => timelinePanelDateFilterError(draft.value))
  const canConfigure = computed(() => {
    void version.value
    return Boolean(snapshot.value && !stale.value && canConfigurePanelSlicer(options.studio,
      { workbookId: options.workbookId, sheetId: options.sheetId }, options.timelineId))
  })
  const hasDateBounds = computed(() => bounds.value?.minDate !== undefined && bounds.value.maxDate !== undefined)
  const canApply = computed(() => canConfigure.value && !busy.value && !composing.value && dirty.value && !validation.value
    && (draft.value.mode === 'none' || !boundsLoading.value && hasDateBounds.value && !boundsError.value))
  async function queryBounds(current: TimelineRangePanelSnapshot): Promise<void> {
    const request = ++boundsGeneration
    bounds.value = undefined; boundsError.value = ''; boundsLoading.value = true
    const active = () => request === boundsGeneration && snapshot.value === current && sameHost(current)
    try { const value = await readPanelTimelineDateBounds(current, active); if (active()) bounds.value = value }
    catch (failure) { if (active()) boundsError.value = failure instanceof Error ? failure.message : String(failure) }
    finally { if (active()) boundsLoading.value = false }
  }
  function read(): void {
    if (!options.open || stopped || busy.value || composing.value) return
    generation += 1; boundsGeneration += 1; publishing = undefined; executing = undefined; published = undefined; publishingCommand = undefined
    error.value = ''; boundsError.value = ''; bounds.value = undefined
    try {
      const current = readPanelTimelineDateFilter(options.studio, { workbookId: options.workbookId, sheetId: options.sheetId }, options.timelineId)
      snapshot.value = current; draft.value = createTimelineDateFilterDraft(current); readSession.value += 1
      void queryBounds(current)
    } catch (failure) { snapshot.value = undefined; boundsLoading.value = false; error.value = failure instanceof Error ? failure.message : String(failure) }
  }
  function isOwnCommand(command: { id: string; params?: unknown } | undefined, before = false): boolean {
    const params = command?.params && typeof command.params === 'object' && !Array.isArray(command.params) ? command.params as Record<string, unknown> : undefined
    if (!publishing || !publishingCommand || !params || command?.id !== publishingCommand.id || params.slicerId !== publishing.timelineId
      || params.expectedRevision !== publishing.owner.revision || params.host !== 'workbench') return false
    const payload = publishingCommand.payloadKey === 'params' ? params : params[publishingCommand.payloadKey]
    return before ? payload === publishingCommand.payload : executing === publishing && equivalentModelData(payload, publishingCommand.payload)
  }
  function changedContent(command?: { id: string; params?: unknown }): void {
    if (isOwnCommand(command)) published = publishing
    version.value += 1
    if (options.open && !busy.value && !dirty.value && !composing.value) read()
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
  watch([() => options.open, () => options.studio, () => options.workbookId, () => options.sheetId, () => options.timelineId], () => {
    generation += 1; boundsGeneration += 1; busyGeneration += 1; busy.value = false; boundsLoading.value = false; composing.value = false
    publishing = undefined; executing = undefined; published = undefined; publishingCommand = undefined; snapshot.value = undefined; draft.value = emptyDraft()
    read()
  }, { immediate: true, flush: 'sync' })
  watch(() => options.studio.editor.workbook.model, () => changedContent(), { flush: 'post' })
  watch(stale, value => {
    if (!value) return
    boundsGeneration += 1; boundsLoading.value = false
    if (busy.value) { generation += 1; busyGeneration += 1; busy.value = false; publishing = undefined; executing = undefined; published = undefined; publishingCommand = undefined }
  }, { flush: 'sync' })
  function edit(action: () => void): void { if (!busy.value && !composing.value && canConfigure.value) { error.value = ''; action() } }
  function editInput(action: () => void): void { if (!busy.value && (canConfigure.value || composing.value)) { error.value = ''; action() } }
  function setMode(value: unknown): void { if (value === 'none' || value === 'range' || value === 'condition') edit(() => { draft.value.mode = value }) }
  function setRange(patch: Partial<TimelineDateFilterDraft['range']>): void {
    const action = () => { Object.assign(draft.value.range, patch) }
    if (Object.hasOwn(patch, 'useWholeDay')) edit(action); else editInput(action)
  }
  function setCondition(patch: Partial<TimelineDateFilterDraft['condition']>): void {
    const action = () => { Object.assign(draft.value.condition, patch) }
    if (Object.hasOwn(patch, 'operator') || Object.hasOwn(patch, 'useWholeDay')) edit(action); else editInput(action)
  }
  function fillSourceDays(): void {
    edit(() => {
      try { if (bounds.value) draft.value = fillTimelinePanelSourceDays(draft.value, bounds.value) }
      catch (failure) { error.value = failure instanceof Error ? failure.message : String(failure) }
    })
  }
  function compositionStart(): void { if (!busy.value && canConfigure.value) composing.value = true }
  function compositionEnd(): void { composing.value = false }
  function close(): void {
    if (composing.value) return
    const owner = snapshot.value?.owner
    generation += 1; boundsGeneration += 1; busyGeneration += 1; busy.value = false; boundsLoading.value = false
    publishing = undefined; executing = undefined; published = undefined; publishingCommand = undefined
    options.close(owner)
  }
  async function apply(): Promise<boolean> {
    const current = snapshot.value
    if (!current || !canApply.value) return false
    const request = generation, busyToken = ++busyGeneration, value = cloneModel(draft.value)
    busy.value = true; error.value = ''; publishing = current; executing = undefined; published = undefined; publishingCommand = undefined
    try {
      const accepted = await applyPanelTimelineDateFilter(current, value, () => request === generation && sameHost(current),
        command => { if (publishing === current) publishingCommand = command })
      if (request !== generation || !sameHost(current) || !timelineRangePanelCurrent(current, true)) return false
      if (!accepted) { error.value = '日期筛选更新未接受，请检查修改权限和真实日期来源。'; return false }
      options.changed(current.timelineId)
      if (request === generation && sameHost(current) && timelineRangePanelCurrent(current, true)) close()
      return true
    } catch (failure) {
      if (request === generation && sameHost(current)) error.value = failure instanceof Error ? failure.message : String(failure)
      return false
    } finally {
      if (publishing === current) { publishing = undefined; publishingCommand = undefined }
      if (executing === current) { executing = undefined; version.value += 1 }
      if (published === current) { published = undefined; version.value += 1 }
      if (!stopped && busyToken === busyGeneration) busy.value = false
    }
  }
  onScopeDispose(() => { stopped = true; generation += 1; boundsGeneration += 1; busyGeneration += 1; subscriptions.forEach(subscription => subscription.dispose()) })
  return { snapshot, draft, bounds, boundsLoading, boundsError, busy, error, composing, stale, dirty, validation, canConfigure, hasDateBounds, canApply, readSession,
    read, setMode, setRange, setCondition, fillSourceDays, compositionStart, compositionEnd, apply, close }
}
