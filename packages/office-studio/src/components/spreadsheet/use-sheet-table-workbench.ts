import { computed, nextTick, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { Disposable } from '../../api/events.ts'
import type { CellRange } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { SHEET_TABLE_COMMANDS, type SheetTableSelectorRequest } from '../../features/sheet-tables/sheet-table-create.ts'

interface SheetTableWorkbenchState {
  studio?: OfficeStudio; unitId?: string; sheetId: string; active: boolean; blocked: boolean; userId: string
  ranges: readonly CellRange[]
}
interface SheetTableWorkbenchOptions {
  state(): SheetTableWorkbenchState
  refresh(studio: OfficeStudio, unitId: string): void
  focus(): void
  canRestoreFocus(): boolean
}
interface SheetTableDialogOwner { studio: OfficeStudio; request: SheetTableSelectorRequest; generation: number; model: object }

export function useSheetTableWorkbench(options: SheetTableWorkbenchOptions) {
  const open = ref(false), error = ref(''), pending = shallowRef<SheetTableDialogOwner>()
  const request = computed(() => pending.value?.request)
  let sequence = 0, stopped = false, subscription: Disposable | undefined, closeOwner: SheetTableDialogOwner | undefined
  function current(owner: SheetTableDialogOwner, revision = owner.request.expectedRevision): boolean {
    const state = options.state(), { studio, request } = owner
    return !stopped && state.active && state.studio === studio && state.unitId === request.unitId && state.sheetId === request.sheetId
      && owner.generation === sequence && !studio.isDisposed && studio.activeWorkbookId === request.unitId
      && studio.editor.state.activeKind === 'workbook' && studio.getWorkbookModel(request.unitId) === owner.model
      && studio.getWorkbookModel(request.unitId)?.activeSheetId === request.sheetId
      && studio.getWorkbookRevision(request.unitId) === revision && studio.currentUserId === request.expectedUserId
      && studio.editor.history.captureSavepoint(request.unitId).generation === request.expectedHistoryGeneration
  }
  function requested(value: SheetTableSelectorRequest): void {
    const state = options.state(), studio = state.studio
    if (!studio || state.blocked || open.value || !state.unitId || state.unitId !== value.unitId || state.sheetId !== value.sheetId) return
    const model = studio.getWorkbookModel(value.unitId)
    if (!model) return
    const owner = { studio, request: cloneModel(value), generation: sequence, model }
    if (!current(owner)) return
    error.value = ''; pending.value = owner; closeOwner = undefined; open.value = true
  }
  function cancel(): void {
    closeOwner = pending.value; open.value = false; pending.value = undefined; error.value = ''
  }
  function afterClose(): void {
    const owner = closeOwner; closeOwner = undefined
    if (!owner) return
    void nextTick(() => { if (!open.value && !pending.value && current(owner) && options.canRestoreFocus()) options.focus() })
  }
  function confirm(range: string): boolean {
    const owner = pending.value
    if (!owner || !open.value || options.state().blocked) return false
    try {
      if (!current(owner) || owner.studio.getWorkbookModel(owner.request.unitId) !== owner.model) throw new Error('表格目标已更改，请关闭后重新选择范围。')
      const changed = owner.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { ...owner.request, range, host: 'workbench' })
      if (!changed) return false
      const model = owner.studio.getWorkbookModel(owner.request.unitId)
      if (!model) { if (pending.value === owner) { cancel(); closeOwner = undefined }; return true }
      const completed = { ...owner, model,
        request: { ...owner.request, expectedRevision: owner.request.expectedRevision + 1 } }
      if (pending.value !== owner || !current(completed)) {
        if (pending.value === owner) { cancel(); closeOwner = undefined }
        return true
      }
      options.refresh(owner.studio, owner.request.unitId)
      if (pending.value !== owner || !current(completed)) {
        if (pending.value === owner) { cancel(); closeOwner = undefined }
        return true
      }
      closeOwner = completed
      open.value = false; pending.value = undefined; error.value = ''
      return true
    } catch (cause) { if (pending.value === owner) error.value = cause instanceof Error ? cause.message : '无法创建表格。'; return false }
  }
  function openSelector(): boolean {
    const state = options.state(), studio = state.studio
    if (!studio || !state.unitId || !state.active || state.blocked || open.value) return false
    return studio.executeCommandSync(SHEET_TABLE_COMMANDS.openSelector, { unitId: state.unitId, sheetId: state.sheetId,
      ranges: cloneModel([...state.ranges]), host: 'workbench' })
  }
  const canSubmit = computed(() => Boolean(pending.value && !options.state().blocked && current(pending.value)))
  watch(() => options.state().studio, studio => { subscription?.dispose(); subscription = studio?.events.addEvent('SheetTableSelectorRequested', requested) }, { immediate: true, flush: 'sync' })
  watch(() => { const state = options.state(); return [state.studio, state.unitId, state.sheetId, state.active, state.userId] as const }, (state, previous) => {
    if (previous && state.some((value, index) => value !== previous[index])) { sequence++; cancel(); closeOwner = undefined }
  }, { flush: 'sync' })
  onScopeDispose(() => { stopped = true; sequence++; subscription?.dispose(); cancel(); closeOwner = undefined })
  return { open, request, error, canSubmit, openSelector, confirm, cancel, afterClose }
}
