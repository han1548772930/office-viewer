import { computed, nextTick, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { SheetOutlineAction } from './sheet-outline.ts'
import { applySheetOutlinePanelChange, captureSheetOutlinePanelOwner, getSheetOutlinePanelSheet,
  prepareSheetOutlinePanelChange, sheetOutlinePanelChangePublished, sheetOutlinePanelOwnerCurrent,
  type SheetOutlinePanelOwner } from './sheet-outline-panel-owner.ts'
import { sheetOutlineErrorMessage } from './sheet-outline-types.ts'

export interface SheetOutlineDialogOptions {
  studio: OfficeStudio; workbookId: string; sheetId: string; open: boolean
  draftKey?: unknown
  close: (owner: SheetOutlinePanelOwner | undefined, committed: boolean) => void
  changed?: () => void
}

export function useSheetOutlineDialog(options: SheetOutlineDialogOptions) {
  const owner = shallowRef<SheetOutlinePanelOwner>(), busy = ref(false), error = ref(''), version = ref(0)
  let generation = 0, stopped = false, publishing: SheetOutlinePanelOwner | undefined
  let subscriptions: Array<{ dispose(): void }> = []
  const sameHost = (captured: SheetOutlinePanelOwner) => !stopped && options.open && options.studio === captured.studio
    && options.workbookId === captured.target.workbookId && options.sheetId === captured.target.sheetId
  const stale = computed(() => {
    void version.value
    return Boolean(owner.value && (!sameHost(owner.value) || !sheetOutlinePanelOwnerCurrent(owner.value)
      && !(publishing === owner.value && sheetOutlinePanelOwnerCurrent(owner.value, true))))
  })
  const sheet = computed(() => {
    void version.value
    if (!owner.value || stale.value) return undefined
    try { return getSheetOutlinePanelSheet(owner.value) } catch { return undefined }
  })
  function read(): void {
    if (stopped || !options.open || busy.value) return
    generation += 1; error.value = ''; publishing = undefined
    try { owner.value = captureSheetOutlinePanelOwner(options.studio, { workbookId: options.workbookId, sheetId: options.sheetId }) }
    catch (failure) { owner.value = undefined; error.value = sheetOutlineErrorMessage(failure) }
  }
  watch(() => options.studio, studio => {
    subscriptions.forEach(subscription => subscription.dispose())
    const changed = () => { version.value += 1 }
    subscriptions = [
      studio.events.addEvent('CommandExecuted', event => { if (event.unitId === options.workbookId) changed() }),
      studio.events.addEvent('Undo', event => { if (event.unitId === options.workbookId) changed() }),
      studio.events.addEvent('Redo', event => { if (event.unitId === options.workbookId) changed() }),
      studio.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === options.workbookId) changed() }),
    ]
  }, { immediate: true })
  watch([() => options.open, () => options.studio, () => options.workbookId, () => options.sheetId, () => options.draftKey], () => {
    generation += 1; publishing = undefined; busy.value = false; owner.value = undefined; error.value = ''; read()
  }, { immediate: true, flush: 'sync' })
  watch(() => options.studio.editor.workbook.model, () => { version.value += 1 }, { flush: 'sync' })
  function close(committed = false): void {
    const captured = owner.value
    generation += 1; busy.value = false; publishing = undefined
    options.close(captured, committed)
  }
  function apply(action: SheetOutlineAction): boolean {
    const captured = owner.value
    if (!captured || busy.value || stale.value) return false
    const request = generation, current = () => request === generation && sameHost(captured)
    busy.value = true; publishing = captured; error.value = ''
    try {
      const prepared = prepareSheetOutlinePanelChange(captured, action)
      if (!prepared) { error.value = '大纲设置没有发生变化。'; return false }
      const accepted = applySheetOutlinePanelChange(prepared, current)
      if (!current()) return false
      if (!accepted) { error.value = '大纲更新未接受，请检查当前权限和内容。'; return false }
      if (!sheetOutlinePanelChangePublished(prepared)) return false
      options.changed?.()
      if (current() && sheetOutlinePanelChangePublished(prepared)) close(true)
      return true
    } catch (failure) {
      if (current()) error.value = sheetOutlineErrorMessage(failure)
      return false
    } finally {
      if (publishing === captured) publishing = undefined
      if (!stopped && request === generation) busy.value = false
      version.value += 1
    }
  }
  onScopeDispose(() => { stopped = true; generation += 1; subscriptions.forEach(subscription => subscription.dispose()) })
  return { owner, sheet, busy, error, stale, read, apply, close }
}

export function useSheetOutlineDialogFocus(options: Pick<SheetOutlineDialogOptions, 'studio' | 'workbookId' | 'sheetId' | 'open' | 'draftKey'>
  & { owner: () => SheetOutlinePanelOwner | undefined; emitClose: () => void }) {
  let generation = 0, stopped = false, captured: { element: HTMLElement; owner: SheetOutlinePanelOwner } | undefined
  watch([() => options.open, () => options.studio, () => options.workbookId, () => options.sheetId, () => options.draftKey], (current, previous) => {
    if (current[0] || !previous || current.slice(1).some((value, index) => value !== previous[index + 1])) generation += 1
    captured = undefined
    const owner = options.owner()
    if (options.open && owner && typeof document !== 'undefined' && document.activeElement instanceof HTMLElement) {
      captured = { element: document.activeElement, owner }
    }
  }, { immediate: true, flush: 'sync' })
  function close(owner: SheetOutlinePanelOwner | undefined, committed: boolean): void {
    const saved = captured, request = ++generation
    captured = undefined; options.emitClose()
    if (!saved || !owner) return
    void nextTick(() => {
      if (stopped || generation !== request || options.open || options.studio !== owner.studio
        || options.workbookId !== owner.target.workbookId || options.sheetId !== owner.target.sheetId
        || saved.owner.studio !== owner.studio || saved.owner.userId !== owner.userId
        || saved.owner.target.workbookId !== owner.target.workbookId || saved.owner.target.sheetId !== owner.target.sheetId
        || !sheetOutlinePanelOwnerCurrent(owner, committed) || !saved.element.isConnected) return
      saved.element.focus({ preventScroll: true })
    })
  }
  onScopeDispose(() => { stopped = true; generation += 1; captured = undefined })
  return { close }
}
