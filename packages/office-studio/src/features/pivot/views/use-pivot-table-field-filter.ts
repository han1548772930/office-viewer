import { computed, nextTick, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../../api/office-studio.ts'
import type { PivotPanelDraft } from '../services/pivot-panel-state.ts'
import { pivotChartPanelOwnerContext, pivotChartPanelOwnerCurrent,
  type PivotChartPanelOwner } from '../services/pivot-chart-panel-owner.ts'
import { applyPivotTableFieldFilterAction, pivotTableFieldFilterDraftCompatible, pivotTableFieldFilterFacade,
  pivotTableFieldFilterPanelSignature, readPivotTableFieldFilterSnapshot, type PivotTableFieldFilterAction,
  type PivotTableFieldFilterSnapshot, type PivotTableFieldFilterTarget } from '../services/pivot-table-field-filter-state.ts'
import { createPivotChartFieldFilterDraft, setPivotChartFilterAll, setPivotChartFilterSearch, setPivotChartFilterSelection,
  type PivotChartFieldFilterDraft } from '../services/pivot-chart-field-button-state.ts'
import { createPivotChartFieldConditionsDraft, setPivotChartDateConditionDraft, setPivotChartLabelConditionDraft,
  type PivotChartDateConditionDraft, type PivotChartFieldConditionsDraft, type PivotChartLabelConditionDraft } from '../services/pivot-chart-field-filter-condition.ts'

export interface PivotTableFieldFilterOptions {
  studio?: OfficeStudio
  unitId?: string
  sheetId: string
  pivotId: string
  panelDraft: PivotPanelDraft | null
  owner?: PivotChartPanelOwner
  active: boolean
  sourceReady: boolean
  busy?: boolean
  restoreFocus: (fieldId: string) => void
  applied: (fieldId: string, action: PivotTableFieldFilterAction, changed: boolean) => void
}
interface Session { owner: PivotChartPanelOwner; draft: PivotPanelDraft; signature: string; pivotId: string; fieldId: string; generation: number }

export function usePivotTableFieldFilter(options: PivotTableFieldFilterOptions) {
  const snapshot = shallowRef<PivotTableFieldFilterSnapshot>(), draft = ref<PivotChartFieldFilterDraft>(), conditions = ref<PivotChartFieldConditionsDraft>()
  const fieldId = ref(''), loading = ref(false), busy = ref(false), composing = ref(false), stale = ref(false), error = ref('')
  let session: Session | undefined, publishing: Session | undefined
  let disposed = false, generation = 0, busyGeneration = 0, compositionGeneration = 0
  const isOpen = computed(() => Boolean(fieldId.value) && options.active)
  const field = computed(() => snapshot.value?.field)
  const readable = computed(() => isOpen.value && !stale.value && Boolean(session && current(session, busy.value && publishing === session)))
  const navigationDisabled = computed(() => busy.value || Boolean(options.busy) || composing.value || stale.value || !readable.value)
  const editDisabled = computed(() => navigationDisabled.value || loading.value || !snapshot.value?.editable)
  function target(): PivotTableFieldFilterTarget {
    if (!options.studio || !options.unitId || !options.pivotId) throw new Error('请先选择已创建的数据透视表。')
    return { studio: options.studio, unitId: options.unitId, sheetId: options.sheetId, pivotId: options.pivotId }
  }
  function current(captured: Session, published = false): boolean {
    return !disposed && options.active && options.sourceReady && session === captured && captured.generation === generation
      && options.studio === captured.owner.studio && options.unitId === captured.owner.unitId && options.sheetId === captured.owner.sheetId
      && options.pivotId === captured.pivotId && options.panelDraft === captured.draft && options.owner === captured.owner
      && pivotTableFieldFilterPanelSignature(captured.draft) === captured.signature && pivotChartPanelOwnerCurrent(captured.owner, published)
  }
  function invalidate(): void {
    generation++; busyGeneration++; compositionGeneration++; session = undefined; publishing = undefined
    fieldId.value = ''; snapshot.value = undefined; draft.value = undefined; conditions.value = undefined
    loading.value = false; busy.value = false; composing.value = false; stale.value = false; error.value = ''
  }
  function synchronize(): void {
    const captured = session
    if (!captured || disposed) return
    if (!current(captured, publishing === captured)) {
      if (composing.value) { stale.value = true; loading.value = false }
      else invalidate()
      return
    }
    try { snapshot.value = readPivotTableFieldFilterSnapshot(target(), captured.fieldId) }
    catch (reason) { if (composing.value) { stale.value = true; error.value = message(reason) } else invalidate() }
  }
  async function open(field: string): Promise<void> {
    if (disposed || !options.active || options.busy || busy.value || composing.value) return
    let captured: Session | undefined
    try {
      const value = target(), panel = options.panelDraft, owner = options.owner, workbook = value.studio.getWorkbookModel(value.unitId)
      if (!panel || !owner || !options.sourceReady || !pivotChartPanelOwnerCurrent(owner)
        || !workbook || !pivotTableFieldFilterDraftCompatible(workbook, panel, value.pivotId, field)) throw new Error('请应用当前数据源配置或重新读取后再筛选字段。')
      const next = readPivotTableFieldFilterSnapshot(value, field)
      invalidate(); captured = { owner, draft: panel, signature: pivotTableFieldFilterPanelSignature(panel), pivotId: value.pivotId, fieldId: field, generation }
      session = captured; fieldId.value = field; snapshot.value = next; loading.value = true
      const pivot = pivotTableFieldFilterFacade(value), items = await pivot.getFieldItems(field)
      if (!current(captured)) return
      const fresh = readPivotTableFieldFilterSnapshot(value, field), itemState = pivot.getManualFilterItemState(field)
      const nextConditions = createPivotChartFieldConditionsDraft(pivot.getFilterCapabilities(field), pivot.getDateFilter(field), pivot.getLabelConditionFilter(field))
      const nextDraft = createPivotChartFieldFilterDraft(items, itemState)
      if (!current(captured)) return
      snapshot.value = fresh; draft.value = nextDraft; conditions.value = nextConditions
    } catch (reason) { if (!captured || current(captured)) error.value = message(reason) }
    finally {
      if (captured && session === captured && generation === captured.generation) {
        loading.value = false
        if (!current(captured)) synchronize()
      }
    }
  }
  function restore(captured: Session, published = false): void {
    const request = generation
    void nextTick(() => {
      if (!disposed && request === generation && options.active && !isOpen.value && !composing.value
        && options.studio === captured.owner.studio && options.unitId === captured.owner.unitId && options.sheetId === captured.owner.sheetId
        && options.pivotId === captured.pivotId && pivotChartPanelOwnerCurrent(captured.owner, published)) options.restoreFocus(captured.fieldId)
    })
  }
  function close(restoreFocus = true): void {
    if (busy.value || composing.value) return
    const captured = session; invalidate()
    if (captured && restoreFocus) restore(captured)
  }
  function updateMembers(update: (value: PivotChartFieldFilterDraft) => PivotChartFieldFilterDraft, text = false): void {
    if (!draft.value || busy.value || loading.value || options.busy || stale.value || composing.value && !text
      || !text && !snapshot.value?.editable || !composing.value && (!session || !current(session))) return
    try { draft.value = update(draft.value); error.value = '' } catch (reason) { error.value = message(reason) }
  }
  const search = (value: string): void => updateMembers(draft => setPivotChartFilterSearch(draft, value), true)
  const select = (key: string, checked: boolean): void => updateMembers(draft => setPivotChartFilterSelection(draft, key, checked))
  const selectAll = (checked: boolean): void => updateMembers(draft => setPivotChartFilterAll(draft, checked))
  function updateConditions(update: (value: PivotChartFieldConditionsDraft) => PivotChartFieldConditionsDraft): void {
    if (!conditions.value || busy.value || loading.value || options.busy || stale.value || !snapshot.value?.editable
      || !composing.value && (!session || !current(session))) return
    try { conditions.value = update(conditions.value); error.value = '' } catch (reason) { error.value = message(reason) }
  }
  const dateChange = (patch: Partial<PivotChartDateConditionDraft>): void => updateConditions(value => setPivotChartDateConditionDraft(value, patch))
  const labelChange = (patch: Partial<PivotChartLabelConditionDraft>): void => updateConditions(value => setPivotChartLabelConditionDraft(value, patch))
  async function mutate(action: PivotTableFieldFilterAction): Promise<boolean> {
    let captured = session, request = 0
    if (!captured || editDisabled.value) return false
    try {
      if (!current(captured)) throw new Error('字段筛选会话已过期。')
      const next = readPivotTableFieldFilterSnapshot(target(), captured.fieldId)
      if (!next.editable) throw new Error('当前用户不能修改透视报告。')
      const context = { ...pivotChartPanelOwnerContext(captured.owner), hostSheetId: captured.owner.sheetId }
      const savedDraft = draft.value ? { ...draft.value, items: [...draft.value.items], selectedKeys: [...draft.value.selectedKeys], partialKeys: [...draft.value.partialKeys] } : undefined
      const savedConditions = conditions.value ? { capabilities: { ...conditions.value.capabilities }, date: { ...conditions.value.date }, label: { ...conditions.value.label } } : undefined
      const pivot = pivotTableFieldFilterFacade(target()); request = ++busyGeneration; publishing = captured; busy.value = true; error.value = ''
      await applyPivotTableFieldFilterAction(pivot, captured.fieldId, action, context, savedDraft, savedConditions)
      if (!current(captured, true)) return false
      const changed = captured.owner.studio.getWorkbookRevision(captured.owner.unitId) !== captured.owner.revision
      invalidate(); options.applied(captured.fieldId, action, changed); restore(captured, true); return true
    } catch (reason) { if (current(captured, true)) error.value = message(reason); return false }
    finally {
      if (!disposed && request && request === busyGeneration) { busy.value = false; if (publishing === captured) publishing = undefined }
      if (session === captured && !current(captured)) synchronize()
    }
  }
  function compositionStart(): void { if (session) { compositionGeneration++; composing.value = true } }
  function compositionEnd(): void {
    const ended = compositionGeneration, captured = session
    void nextTick(() => { if (disposed || ended !== compositionGeneration || session !== captured) return; composing.value = false; synchronize() })
  }
  function keydown(event: KeyboardEvent): void {
    if (event.key !== 'Escape') return
    event.stopPropagation()
    if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || busy.value || composing.value) return
    event.preventDefault(); close()
  }
  function message(reason: unknown): string { return reason instanceof Error ? reason.message : String(reason) }
  watch([() => options.active, () => options.studio, () => options.unitId, () => options.sheetId, () => options.pivotId,
    () => options.panelDraft, () => options.owner, () => options.sourceReady,
    () => options.panelDraft ? pivotTableFieldFilterPanelSignature(options.panelDraft) : '',
    () => options.studio?.currentUserId, () => options.studio?.activeWorkbookId, () => options.studio?.editor.state.activeKind,
    () => options.studio?.getWorkbookModel(options.unitId)?.activeSheetId,
    () => options.studio?.getWorkbookModel(options.unitId) ? options.studio.editor.history.captureSavepoint(options.unitId!).generation : -1],
    synchronize, { flush: 'sync' })
  watch(() => options.studio, (studio, _previous, onCleanup) => {
    if (!studio) return
    const changed = (unitId?: string): void => { if (!unitId || unitId === options.unitId) synchronize() }
    const listeners = [
      studio.events.addEvent('CommandExecuted', event => changed(event.unitId)),
      studio.events.addEvent('Undo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
      studio.events.addEvent('Redo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
      studio.events.addEvent('PivotTableChanged', event => changed(event.unitId)),
      studio.events.addEvent('PivotTableRendered', event => changed(event.unitId)),
      studio.events.addEvent('SheetValueChanged', event => changed(event.workbookId)),
      studio.events.addEvent('SheetStructureChanged', event => changed(event.workbookId)),
      studio.events.addEvent('WorkbookDisposed', event => changed(event.workbookId)),
    ]
    onCleanup(() => listeners.forEach(listener => listener.dispose()))
  }, { immediate: true })
  onScopeDispose(() => { disposed = true; invalidate() })
  return { snapshot, field, fieldId, draft, conditions, isOpen, loading, busy, composing, stale, error, readable, navigationDisabled,
    editDisabled, open, close, mutate, search, select, selectAll, dateChange, labelChange, compositionStart, compositionEnd, keydown, synchronize }
}
