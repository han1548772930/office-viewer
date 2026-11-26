import { computed, nextTick, onScopeDispose, ref, shallowRef, watch } from 'vue'
import { applyPivotChartFieldButtonAction, createPivotChartFieldFilterDraft, pivotChartFieldButtonsFacade,
  readPivotChartFieldButtonsSnapshot, setPivotChartFilterAll, setPivotChartFilterSearch, setPivotChartFilterSelection,
  type PivotChartFieldButton, type PivotChartFieldButtonAction, type PivotChartFieldButtonsSnapshot,
  type PivotChartFieldButtonsTarget, type PivotChartFieldFilterDraft } from '../services/pivot-chart-field-button-state.ts'
import { capturePivotChartPanelOwner, pivotChartPanelOwnerContext, pivotChartPanelOwnerCurrent,
  type PivotChartPanelOwner } from '../services/pivot-chart-panel-owner.ts'
import { createPivotChartFieldConditionsDraft, setPivotChartDateConditionDraft, setPivotChartLabelConditionDraft,
  type PivotChartDateConditionDraft, type PivotChartFieldConditionsDraft, type PivotChartLabelConditionDraft } from '../services/pivot-chart-field-filter-condition.ts'
import type { PivotValuePosition } from '../../../domain/model.ts'
import { normalizePivotValuesPosition } from '../services/pivot-values-position.ts'

export interface PivotChartFieldButtonsOptions extends PivotChartFieldButtonsTarget {
  active?: boolean
  busy?: boolean
  openSettings: (fieldId: string) => boolean | Promise<boolean>
  restoreFocus: () => void
}
interface FieldPopupSession { owner: PivotChartPanelOwner; chartId: string; anchor: string; generation: number; fieldId?: string }

export function usePivotChartFieldButtons(options: PivotChartFieldButtonsOptions) {
  const snapshot = shallowRef<PivotChartFieldButtonsSnapshot>(), field = shallowRef<PivotChartFieldButton>()
  const draft = ref<PivotChartFieldFilterDraft>(), openAnchor = ref(''), loading = ref(false), busy = ref(false), composing = ref(false)
  const conditions = ref<PivotChartFieldConditionsDraft>()
  const error = ref(''), readError = ref(''), stale = ref(false)
  let disposed = false, generation = 0, busyGeneration = 0, compositionGeneration = 0
  let session: FieldPopupSession | undefined, publishing: FieldPopupSession | undefined
  const visible = computed(() => options.active !== false && Boolean(snapshot.value?.controls.showFieldButtons))
  const readable = computed(() => Boolean(openAnchor.value) && visible.value && !stale.value
    && Boolean(session && current(session, busy.value && publishing === session)))
  const navigationDisabled = computed(() => busy.value || Boolean(options.busy) || composing.value || stale.value || !snapshot.value)
  const editDisabled = computed(() => navigationDisabled.value || !snapshot.value?.editable || loading.value || !readable.value)
  const target = (): PivotChartFieldButtonsTarget => ({ studio: options.studio, unitId: options.unitId, sheetId: options.sheetId, chartId: options.chartId })

  function current(captured: FieldPopupSession, published = false): boolean {
    return !disposed && session === captured && generation === captured.generation && options.active !== false
      && options.studio === captured.owner.studio && options.unitId === captured.owner.unitId
      && options.sheetId === captured.owner.sheetId && options.chartId === captured.chartId
      && pivotChartPanelOwnerCurrent(captured.owner, published)
  }
  function invalidatePopup(): void {
    generation++; busyGeneration++; compositionGeneration++
    session = undefined; publishing = undefined; openAnchor.value = ''; field.value = undefined; draft.value = undefined; conditions.value = undefined
    loading.value = false; busy.value = false; composing.value = false; stale.value = false; error.value = ''
  }
  function synchronize(): void {
    if (disposed) return
    try { snapshot.value = options.active === false ? undefined : readPivotChartFieldButtonsSnapshot(target()); readError.value = '' }
    catch (reason) { snapshot.value = undefined; readError.value = message(reason) }
    if (!visible.value) { invalidatePopup(); return }
    if (session && !current(session, publishing === session)) {
      if (composing.value) { stale.value = true; loading.value = false }
      else invalidatePopup()
    }
  }
  function begin(anchor: string, fieldId?: string): FieldPopupSession {
    if (disposed || options.active === false || navigationDisabled.value) throw new Error('透视图字段操作当前不可用。')
    const next = readPivotChartFieldButtonsSnapshot(target())
    if (!next.controls.showFieldButtons) throw new Error('透视图字段按钮已隐藏。')
    invalidatePopup(); snapshot.value = next
    const captured = { owner: capturePivotChartPanelOwner(options.studio, options.unitId, options.sheetId),
      chartId: options.chartId, anchor, fieldId, generation }
    session = captured; openAnchor.value = anchor
    return captured
  }
  function openList(anchor: string): void {
    try { begin(anchor) }
    catch (reason) { readError.value = message(reason); synchronize() }
  }
  function openValues(anchor = 'values:position'): void {
    try {
      begin(anchor)
      if (!snapshot.value?.values.visible) throw new Error('当前报告没有可排列的多个值字段。')
    } catch (reason) { invalidatePopup(); readError.value = message(reason); synchronize() }
  }
  async function openField(fieldId: string, anchor = `field:${fieldId}`): Promise<void> {
    let captured: FieldPopupSession | undefined
    try {
      captured = begin(anchor, fieldId)
      const button = snapshot.value!.sections.flatMap(section => section.fields).find(button => button.fieldId === fieldId)
      if (!button) throw new Error('透视图字段已经被移除。')
      field.value = button
      if (button.area === 'value') return
      loading.value = true
      const pivot = pivotChartFieldButtonsFacade(target())
      const items = await pivot.getFieldItems(fieldId)
      if (!current(captured)) return
      const fresh = readPivotChartFieldButtonsSnapshot(target())
      if (!current(captured) || !fresh.controls.showFieldButtons) return
      const itemState = pivot.getManualFilterItemState(fieldId), capabilities = pivot.getFilterCapabilities(fieldId)
      const nextConditions = createPivotChartFieldConditionsDraft(capabilities, pivot.getDateFilter(fieldId), pivot.getLabelConditionFilter(fieldId))
      const nextDraft = createPivotChartFieldFilterDraft(items, itemState)
      if (!current(captured)) return
      snapshot.value = fresh
      field.value = fresh.sections.flatMap(section => section.fields).find(button => button.fieldId === fieldId)
      draft.value = nextDraft; conditions.value = nextConditions
    } catch (reason) {
      if (captured && current(captured)) error.value = message(reason)
      else if (!captured) { readError.value = message(reason); synchronize() }
    } finally {
      if (captured && session === captured && generation === captured.generation) {
        loading.value = false
        if (!current(captured)) synchronize()
      }
    }
  }
  function restore(owner: PivotChartPanelOwner, chartId: string, published = false): void {
    const request = generation
    void nextTick(() => {
      if (!disposed && request === generation && !openAnchor.value && !composing.value && options.active !== false
        && options.studio === owner.studio && options.unitId === owner.unitId && options.sheetId === owner.sheetId
        && options.chartId === chartId && pivotChartPanelOwnerCurrent(owner, published)) options.restoreFocus()
    })
  }
  function close(restoreFocus = true): void {
    if (busy.value || composing.value) return
    const captured = session
    invalidatePopup()
    if (restoreFocus && captured) restore(captured.owner, captured.chartId)
  }
  function requireSession(edit = false, values = false): FieldPopupSession {
    const captured = session
    if (!captured || !current(captured) || stale.value || navigationDisabled.value) throw new Error('透视图字段会话已过期，请重新打开。')
    const next = readPivotChartFieldButtonsSnapshot(target())
    if (!next.controls.showFieldButtons || edit && !next.editable) throw new Error('当前用户无法修改透视图字段或关联报告。')
    if (values ? captured.fieldId !== undefined || !captured.anchor.startsWith('values:') || !next.values.visible
      : !captured.fieldId || !next.sections.some(section => section.fields.some(field => field.fieldId === captured.fieldId))) throw new Error('透视图字段已经被移除。')
    snapshot.value = next
    return captured
  }
  function updateDraft(update: (current: PivotChartFieldFilterDraft) => PivotChartFieldFilterDraft, text = false): void {
    if (!draft.value || busy.value || loading.value || options.busy || !text && !snapshot.value?.editable || stale.value
      || composing.value && !text || !composing.value && (!session || !current(session))) return
    try { draft.value = update(draft.value); error.value = '' }
    catch (reason) { error.value = message(reason) }
  }
  const search = (value: string): void => updateDraft(draft => setPivotChartFilterSearch(draft, value), true)
  const select = (key: string, checked: boolean): void => updateDraft(draft => setPivotChartFilterSelection(draft, key, checked))
  const selectAll = (checked: boolean): void => updateDraft(draft => setPivotChartFilterAll(draft, checked))
  function updateConditions(update: (value: PivotChartFieldConditionsDraft) => PivotChartFieldConditionsDraft, text = false): void {
    if (!conditions.value || busy.value || loading.value || options.busy || !snapshot.value?.editable || stale.value
      || composing.value && !text || !composing.value && (!session || !current(session))) return
    try { conditions.value = update(conditions.value); error.value = '' }
    catch (reason) { error.value = message(reason) }
  }
  const dateChange = (patch: Partial<PivotChartDateConditionDraft>): void => updateConditions(draft => setPivotChartDateConditionDraft(draft, patch), true)
  const labelChange = (patch: Partial<PivotChartLabelConditionDraft>): void => updateConditions(draft => setPivotChartLabelConditionDraft(draft, patch), true)
  async function mutate(action: PivotChartFieldButtonAction): Promise<boolean> {
    if (editDisabled.value) return false
    let captured: FieldPopupSession | undefined, request = 0
    try {
      captured = requireSession(true)
      const context = pivotChartPanelOwnerContext(captured.owner), pivot = pivotChartFieldButtonsFacade(target())
      const savedDraft = draft.value ? { items: [...draft.value.items], selectedKeys: [...draft.value.selectedKeys], partialKeys: [...draft.value.partialKeys], search: draft.value.search } : undefined
      const savedConditions = conditions.value ? { capabilities: { ...conditions.value.capabilities }, date: { ...conditions.value.date }, label: { ...conditions.value.label } } : undefined
      request = ++busyGeneration; busy.value = true; publishing = captured; error.value = ''
      const accepted = await applyPivotChartFieldButtonAction(pivot, captured.fieldId!, action, context, savedDraft, savedConditions)
      if (!current(captured, true)) return false
      if (!accepted) throw new Error('透视图字段操作未能提交，请重新读取后重试。')
      const owner = captured.owner, chartId = captured.chartId
      invalidatePopup(); synchronize(); restore(owner, chartId, true)
      return true
    } catch (reason) {
      if (captured && current(captured, true)) error.value = message(reason)
      else if (!captured) { error.value = message(reason); synchronize() }
      return false
    } finally {
      if (!disposed && request && request === busyGeneration) { busy.value = false; if (publishing === captured) publishing = undefined }
      if (captured && session === captured && !current(captured)) synchronize()
    }
  }
  async function mutateValues(position: PivotValuePosition, index = -1): Promise<boolean> {
    if (editDisabled.value) return false
    let captured: FieldPopupSession | undefined, request = 0
    try {
      const state = normalizePivotValuesPosition(position, index)
      captured = requireSession(true, true)
      const context = pivotChartPanelOwnerContext(captured.owner), pivot = pivotChartFieldButtonsFacade(target())
      request = ++busyGeneration; busy.value = true; publishing = captured; error.value = ''
      const accepted = await pivot.updateValuePosition(state.position, state.index, context)
      if (!current(captured, true)) return false
      if (!accepted) throw new Error('ΣValues 位置未能提交，请重新读取后重试。')
      const owner = captured.owner, chartId = captured.chartId
      invalidatePopup(); synchronize(); restore(owner, chartId, true); return true
    } catch (reason) {
      if (captured && current(captured, true)) error.value = message(reason)
      else if (!captured) { error.value = message(reason); synchronize() }
      return false
    } finally {
      if (!disposed && request && request === busyGeneration) { busy.value = false; if (publishing === captured) publishing = undefined }
      if (captured && session === captured && !current(captured)) synchronize()
    }
  }
  async function settings(): Promise<void> {
    if (navigationDisabled.value) return
    let captured: FieldPopupSession | undefined, request = 0
    try {
      captured = requireSession()
      request = ++busyGeneration; busy.value = true; error.value = ''
      const accepted = await options.openSettings(captured.fieldId!)
      if (!current(captured)) return
      if (!accepted) throw new Error('所选字段设置面板未能打开。')
      invalidatePopup()
    } catch (reason) { if (captured && current(captured)) error.value = message(reason) }
    finally { if (!disposed && request && request === busyGeneration) busy.value = false }
  }
  function compositionStart(): void { if (session) { compositionGeneration++; composing.value = true } }
  function compositionEnd(): void {
    const ended = compositionGeneration, captured = session
    void nextTick(() => {
      if (disposed || ended !== compositionGeneration || session !== captured) return
      composing.value = false
      if (stale.value || captured && !current(captured)) synchronize()
    })
  }
  function keydown(event: KeyboardEvent): void {
    if (event.key !== 'Escape') return
    event.stopPropagation()
    if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || composing.value || busy.value) return
    event.preventDefault(); close()
  }
  function message(reason: unknown): string { return reason instanceof Error ? reason.message : String(reason) }

  watch([() => options.studio, () => options.unitId, () => options.sheetId, () => options.chartId, () => options.active,
    () => options.studio.currentUserId, () => options.studio.activeWorkbookId, () => options.studio.editor.state.activeKind,
    () => options.studio.getWorkbookModel(options.unitId)?.activeSheetId,
    () => options.studio.getWorkbookModel(options.unitId) ? options.studio.editor.history.captureSavepoint(options.unitId).generation : -1],
    () => { invalidatePopup(); synchronize() }, { immediate: true, flush: 'sync' })
  watch(() => options.studio, (studio, _previous, onCleanup) => {
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
  onScopeDispose(() => { disposed = true; invalidatePopup(); snapshot.value = undefined })
  return { snapshot, visible, field, draft, conditions, openAnchor, loading, busy, composing, error, readError, readable, stale,
    navigationDisabled, editDisabled, openList, openField, openValues, close, search, select, selectAll, mutate, mutateValues, settings,
    dateChange, labelChange, compositionStart, compositionEnd, keydown, synchronize }
}
