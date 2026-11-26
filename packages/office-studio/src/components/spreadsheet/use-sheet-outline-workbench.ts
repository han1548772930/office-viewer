import { computed, nextTick, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { ApiEventMap, Disposable } from '../../api/events.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { SheetSelectionRange } from '../../features/sheet-workbench/sheet-selection.ts'
import type { OutlineWorksheet, ChangeSheetOutlineInput } from '../../features/sheet-outline/sheet-outline.ts'
import { resolveSheetOutlineCommand, sheetOutlineCommandInput, type SheetOutlineCommandContext } from '../../features/sheet-outline/sheet-outline-command.ts'

export const OPEN_SHEET_OUTLINE_SETTINGS_OPERATION = 'sheet.operation.open-outline-settings'
export interface SheetOutlineWorkbenchState {
  studio?: OfficeStudio; unitId: string; sheetId: string; sheet?: OutlineWorksheet; ranges: readonly SheetSelectionRange[]; userId: string; active: boolean
}
export interface SheetOutlineWorkbenchOptions {
  state(): SheetOutlineWorkbenchState
  apply(input: ChangeSheetOutlineInput): void
  refresh(): void
  focus(): void
}
interface OutlineDialogScope {
  key: number; studio: OfficeStudio; workbookId: string; sheetId: string
  setOpen(open: boolean): void; changed(): void
}
export interface OutlineActionDialogRequest extends OutlineDialogScope { ranges: SheetSelectionRange[]; operation: 'group' | 'remove' }
export type OutlineSettingsDialogRequest = OutlineDialogScope

/** Windows and Mac bindings match the fixed outline UI shortcut declarations. */
export function sheetOutlineCommandForShortcut(event: Pick<KeyboardEvent, 'key' | 'altKey' | 'shiftKey' | 'ctrlKey' | 'metaKey' | 'isComposing' | 'keyCode'>, mac = false): string | undefined {
  if (event.isComposing || event.keyCode === 229) return undefined
  if (!event.ctrlKey && !event.metaKey && event.altKey && event.shiftKey) {
    if (event.key === 'ArrowRight') return 'sheet.command.outline-selected-dimension-shortcut'
    if (event.key === 'ArrowLeft') return 'sheet.command.clear-selected-dimension-outlines-shortcut'
  }
  if (mac && (event.metaKey || event.ctrlKey) && !event.altKey && !event.shiftKey) {
    if (event.key.toLowerCase() === 'k') return 'sheet.command.outline-selected-dimension-shortcut'
    if (event.key.toLowerCase() === 'j') return 'sheet.command.clear-selected-dimension-outlines-shortcut'
  }
  return undefined
}

export function useSheetOutlineWorkbench(options: SheetOutlineWorkbenchOptions) {
  const actionOpen = ref(false), settingsOpen = ref(false)
  const actionRequest = shallowRef<OutlineActionDialogRequest>(), settingsRequest = shallowRef<OutlineSettingsDialogRequest>()
  const uiOpen = computed(() => actionOpen.value || settingsOpen.value)
  let sequence = 0, stopped = false, subscription: Disposable | undefined
  const context = computed<SheetOutlineCommandContext | undefined>(() => {
    const state = options.state()
    const sheet = state.studio?.getWorkbookModel(state.unitId)?.sheets.find(sheet => sheet.id === state.sheetId) ?? state.sheet
    return !state.active || !sheet ? undefined : { unitId: state.unitId, sheetId: state.sheetId, sheet, userId: state.userId, selections: state.ranges }
  })
  function scopeCurrent(scope: OutlineDialogScope): boolean {
    const state = options.state()
    return !stopped && state.active && state.studio === scope.studio && state.unitId === scope.workbookId && state.sheetId === scope.sheetId
      && !scope.studio.isDisposed && scope.studio.activeWorkbookId === scope.workbookId && scope.studio.editor.state.activeKind === 'workbook'
      && scope.studio.getWorkbookModel(scope.workbookId)?.activeSheetId === scope.sheetId
  }
  function createScope(studio: OfficeStudio, unitId: string, sheetId: string, kind: 'action' | 'settings'): OutlineDialogScope {
    const key = ++sequence, userId = studio.currentUserId, generation = studio.editor.history.captureSavepoint(unitId).generation
    let focusRevision = studio.getWorkbookRevision(unitId)
    const current = () => sequence === key && (kind === 'action' ? actionRequest.value : settingsRequest.value)?.key === key
    const scope: OutlineDialogScope = { key, studio, workbookId: unitId, sheetId,
      setOpen(open) {
        if (!current()) return
        if (kind === 'action') actionOpen.value = open; else settingsOpen.value = open
        if (!open) void nextTick().then(() => nextTick(() => {
          if (current() && scopeCurrent(scope) && !uiOpen.value && studio.currentUserId === userId
            && studio.editor.history.captureSavepoint(unitId).generation === generation && studio.getWorkbookRevision(unitId) === focusRevision) options.focus()
        }))
      },
      changed() {
        if (!current() || !scopeCurrent(scope) || studio.currentUserId !== userId
          || studio.editor.history.captureSavepoint(unitId).generation !== generation) return
        focusRevision = studio.getWorkbookRevision(unitId); options.refresh()
      },
    }
    return scope
  }
  function openAction(operation: 'group' | 'remove', ranges: readonly SheetSelectionRange[] = options.state().ranges): boolean {
    const state = options.state(), studio = state.studio
    if (!studio || !state.active || !context.value || studio.isDisposed || studio.activeWorkbookId !== state.unitId
      || studio.editor.state.activeKind !== 'workbook' || studio.getWorkbookModel(state.unitId)?.activeSheetId !== state.sheetId) return false
    settingsOpen.value = false
    actionRequest.value = { ...createScope(studio, state.unitId, state.sheetId, 'action'), ranges: cloneModel([...ranges]), operation }
    actionOpen.value = true
    return true
  }
  function openSettings(): boolean {
    const state = options.state(), studio = state.studio
    if (!studio || !state.active || !context.value || studio.isDisposed || studio.activeWorkbookId !== state.unitId
      || studio.editor.state.activeKind !== 'workbook' || studio.getWorkbookModel(state.unitId)?.activeSheetId !== state.sheetId) return false
    studio.assertCanEditWorksheet(state.unitId, state.sheetId)
    actionOpen.value = false
    settingsRequest.value = createScope(studio, state.unitId, state.sheetId, 'settings'); settingsOpen.value = true
    return true
  }
  function requested(event: ApiEventMap['SheetOutlineActionDialogRequested']): void {
    const state = options.state(), studio = state.studio
    if (!studio || state.unitId !== event.unitId || state.sheetId !== event.sheetId || studio.isDisposed
      || studio.getWorkbookRevision(event.unitId) !== event.expectedRevision || studio.currentUserId !== event.expectedUserId
      || studio.editor.history.captureSavepoint(event.unitId).generation !== event.expectedHistoryGeneration) return
    openAction(event.operation, event.ranges)
  }
  watch(() => options.state().studio, studio => {
    subscription?.dispose(); subscription = studio?.events.addEvent('SheetOutlineActionDialogRequested', requested)
  }, { immediate: true, flush: 'sync' })
  watch(() => { const state = options.state(); return [state.studio, state.unitId, state.sheetId, state.active] as const }, (state, previous) => {
    if (previous && state.some((item, index) => item !== previous[index])) { actionOpen.value = false; settingsOpen.value = false; sequence += 1 }
  }, { flush: 'sync' })
  function run(id: string, value?: unknown): void {
    const current = context.value, state = options.state()
    if (!current) return
    const input = sheetOutlineCommandInput(value)
    if (input.unitId !== undefined && input.unitId !== state.unitId || (input.sheetId ?? input.subUnitId) !== undefined
      && (input.sheetId ?? input.subUnitId) !== state.sheetId) throw new Error('大纲菜单所属工作台已改变。')
    if (state.studio) {
      const changed = state.studio.executeCommandSync(id, { ...input, unitId: state.unitId, subUnitId: state.sheetId, host: 'workbench',
        expectedRevision: input.expectedRevision ?? state.studio.getWorkbookRevision(state.unitId),
        expectedHistoryGeneration: input.expectedHistoryGeneration ?? state.studio.editor.history.captureSavepoint(state.unitId).generation,
        expectedUserId: input.expectedUserId ?? state.userId })
      if (changed && !uiOpen.value) options.refresh()
    } else {
      const result = resolveSheetOutlineCommand(id, input, current)
      if (result?.kind === 'change') options.apply(result.input)
      // Standalone views retain the explicit row/column actions in the local details menu.
    }
    if (!uiOpen.value) void nextTick(() => { if (!stopped && options.state().active && !uiOpen.value) options.focus() })
  }
  onScopeDispose(() => { stopped = true; sequence += 1; subscription?.dispose(); subscription = undefined })
  return { context, actionRequest, settingsRequest, actionOpen, settingsOpen, uiOpen, openAction, openSettings, run }
}
