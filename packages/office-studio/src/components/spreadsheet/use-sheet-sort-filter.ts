import { computed, nextTick, onBeforeUnmount, shallowRef, watch, type Ref } from 'vue'
import { FOffice } from '../../api/facade.ts'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { CellRange } from '../../domain/model.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { assertSheetRangesEditable, sheetRangeContainsCell, sheetRangesEqual } from '../../features/sheet-workbench/sheet-range-command.ts'
import { expandSheetDataRange, prepareSmartSheetFilterRange } from '../../features/sheet-sort-filter/sheet-data-range.ts'
import { dataPanelAddress, type PanelSortRangeSelection } from '../../features/sheet-sort-filter/sheet-data-panel-state.ts'
import { sheetFilterPermissionRanges } from '../../features/sheet-sort-filter/sheet-filter-column.ts'

export interface SheetSortFilterHostState {
  studio?: OfficeStudio
  unitId?: string
  sheetId: string
  userId: string
  active: boolean
  blocked: boolean
  ranges: readonly CellRange[]
  primary?: { row: number; column: number }
}
interface SheetSortFilterOwner {
  studio: OfficeStudio
  unitId: string
  sheetId: string
  userId: string
  lifetime: number
  revision: number
  historyGeneration: number
  signature: string
  sourceRange: CellRange
  primary: { row: number; column: number }
}
export interface SheetSortFilterWorkbenchOptions {
  state(): SheetSortFilterHostState
  open: Ref<boolean>
  mode: Ref<'sort' | 'filter'>
  sourceRange: Ref<string>
  sourcePrimary: Ref<{ row: number; column: number } | undefined>
  busy: Ref<boolean>
  progress: Ref<string>
  refresh(studio: OfficeStudio, unitId: string, current: () => boolean): void
  focus(): void
  notify(type: 'error' | 'info', message: string): void
}

/** Owns UI requests and cancellation only; all content uses the registered Studio commands. */
export function useSheetSortFilterWorkbench(options: SheetSortFilterWorkbenchOptions) {
  const pending = shallowRef<SheetSortFilterOwner>()
  let lifetime = 0, sortGeneration = 0, controller: AbortController | undefined
  const signature = (studio: OfficeStudio, unitId: string, sheetId: string) => {
    const selection = studio.getSelectionState(unitId, sheetId)
    return JSON.stringify(selection ? { ranges: selection.ranges, primary: selection.primary, activeRangeIndex: selection.activeRangeIndex } : null)
  }
  function current(owner: SheetSortFilterOwner, checkRevision = false, checkSelection = true): boolean {
    const state = options.state(), model = owner.studio.getWorkbookModel(owner.unitId)
    const sheet = model?.sheets.find(sheet => sheet.id === owner.sheetId)
    return !owner.studio.isDisposed && owner.lifetime === lifetime && state.studio === owner.studio && state.unitId === owner.unitId
      && state.sheetId === owner.sheetId && state.userId === owner.userId && state.active && !state.blocked
      && owner.studio.activeWorkbookId === owner.unitId && owner.studio.editor.state.activeKind === 'workbook'
      && model?.activeSheetId === owner.sheetId && !!sheet
      && canEditWorksheetProtection(sheet, state.userId) && canEditProtectedRanges(sheet, sheetFilterPermissionRanges(sheet, owner.sourceRange), state.userId)
      && owner.studio.editor.history.captureSavepoint(owner.unitId).generation === owner.historyGeneration
      && (!checkRevision || owner.studio.getWorkbookRevision(owner.unitId) === owner.revision)
      && (!checkSelection || signature(owner.studio, owner.unitId, owner.sheetId) === owner.signature)
  }
  function capture(target?: { range: CellRange; primary: { row: number; column: number } }): SheetSortFilterOwner {
    const state = options.state(), range = target?.range ?? (state.ranges.length === 1 ? state.ranges[0] : undefined), primary = target?.primary ?? state.primary
    if (!state.studio || !state.unitId || !range || !primary || !sheetRangeContainsCell(range, primary.row, primary.column)) {
      throw new Error('请选择一个数据区域及其活动单元格。')
    }
    const owner: SheetSortFilterOwner = { studio: state.studio, unitId: state.unitId, sheetId: state.sheetId,
      userId: state.userId, lifetime, revision: state.studio.getWorkbookRevision(state.unitId),
      historyGeneration: state.studio.editor.history.captureSavepoint(state.unitId).generation,
      signature: signature(state.studio, state.unitId, state.sheetId), sourceRange: { ...range }, primary: { ...primary } }
    if (!current(owner, true)) throw new Error('当前编辑器无法执行排序或筛选。')
    const sheet = owner.studio.getWorkbookModel(owner.unitId)!.sheets.find(sheet => sheet.id === owner.sheetId)!
    assertSheetRangesEditable(sheet, sheetFilterPermissionRanges(sheet, range), state.userId)
    return owner
  }
  const canUseDataRange = computed(() => {
    const state = options.state(), source = state.ranges.length === 1 ? state.ranges[0] : undefined
    const sheet = state.studio && state.unitId ? state.studio.getWorkbookModel(state.unitId)?.sheets.find(sheet => sheet.id === state.sheetId) : undefined
    return !!(state.studio && state.unitId && sheet && source && state.primary && state.active && !state.blocked
      && !options.open.value && !options.busy.value && sheetRangeContainsCell(source, state.primary.row, state.primary.column)
      && state.studio.activeWorkbookId === state.unitId && state.studio.editor.state.activeKind === 'workbook'
      && canEditWorksheetProtection(sheet, state.userId) && canEditProtectedRanges(sheet, sheetFilterPermissionRanges(sheet, source), state.userId))
  })
  const canSort = computed(() => canUseDataRange.value && options.state().ranges[0]!.endRow > options.state().ranges[0]!.startRow)
  const panelEditable = computed(() => !!pending.value && options.open.value && current(pending.value))

  function cancelSort(): void { sortGeneration++; controller?.abort(); controller = undefined; options.busy.value = false }
  function retire(): void { lifetime++; cancelSort(); pending.value = undefined; options.open.value = false }
  watch([() => options.state().studio, () => options.state().unitId, () => options.state().sheetId,
    () => options.state().userId, () => options.state().active], retire, { flush: 'sync' })
  watch(() => options.state().blocked, blocked => { if (blocked) cancelSort() }, { flush: 'sync' })
  watch(options.open, open => {
    if (open) return
    const owner = pending.value
    pending.value = undefined
    if (owner) void nextTick(() => { if (!options.open.value && current(owner)) options.focus() })
  }, { flush: 'sync' })
  watch(() => options.state().studio, (studio, _previous, onCleanup) => {
    if (!studio) return
    const disposed = studio.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === options.state().unitId) retire() })
    const lifecycle = studio.events.addEvent('LifeCycleChanged', event => { if (event.stage === 'disposed') retire() })
    onCleanup(() => { disposed.dispose(); lifecycle.dispose() })
  }, { immediate: true, flush: 'sync' })
  onBeforeUnmount(retire)

  function publishRange(owner: SheetSortFilterOwner, range: CellRange): void {
    if (!current(owner, true)) throw new Error('数据、选区或编辑状态已改变，请重新选择数据区域。')
    const sheet = owner.studio.getWorkbookModel(owner.unitId)!.sheets.find(sheet => sheet.id === owner.sheetId)!
    assertSheetRangesEditable(sheet, [owner.sourceRange, range], owner.userId)
    if (sheetRangesEqual(owner.sourceRange, range)) return
    owner.studio.setSelections(owner.unitId, owner.sheetId, [{ startRow: range.startRow, endRow: range.endRow,
      startColumn: range.startCol, endColumn: range.endCol }], { primary: owner.primary, activeRangeIndex: 0 })
    const selected = owner.studio.getSelectionState(owner.unitId, owner.sheetId)
    const actual = selected?.ranges[0]
    if (!current(owner, true, false) || selected?.ranges.length !== 1 || !actual
      || actual.startRow !== range.startRow || actual.endRow !== range.endRow || actual.startColumn !== range.startCol || actual.endColumn !== range.endCol
      || selected.primary?.row !== owner.primary.row || selected.primary?.column !== owner.primary.column) {
      throw new Error('排序选区已被新的编辑器操作替换。')
    }
    owner.signature = signature(owner.studio, owner.unitId, owner.sheetId)
  }
  function openDataPanel(mode: 'sort' | 'filter', column?: number): boolean {
    if (column === undefined && !canUseDataRange.value) return false
    try {
      const state = options.state(), filter = state.studio && state.unitId
        ? state.studio.getWorkbookModel(state.unitId)?.sheets.find(sheet => sheet.id === state.sheetId)?.autoFilter : undefined
      if (column !== undefined && (mode !== 'filter' || !filter || !Number.isSafeInteger(column)
        || column < filter.range.startCol || column > filter.range.endCol || options.open.value || options.busy.value)) return false
      const owner = capture(column === undefined ? undefined : { range: filter!.range, primary: { row: filter!.range.startRow, column } })
      options.mode.value = mode; options.sourceRange.value = dataPanelAddress(owner.sourceRange)
      options.sourcePrimary.value = { ...owner.primary }; pending.value = owner; options.open.value = true
      return true
    } catch (failure) { options.notify('error', failure instanceof Error ? failure.message : String(failure)); return false }
  }
  function selectDataPanelSortRange(input: PanelSortRangeSelection): boolean {
    const owner = pending.value
    if (!owner || !options.open.value || input.unitId !== owner.unitId || input.sheetId !== owner.sheetId
      || input.expectedRevision !== owner.revision || !sheetRangesEqual(input.sourceRange, owner.sourceRange)
      || input.primary.row !== owner.primary.row || input.primary.column !== owner.primary.column) return false
    publishRange(owner, input.range)
    return current(owner, true)
  }
  function assertDataPanelCurrent(expectedRevision?: number): void {
    const owner = pending.value
    if (!owner || !options.open.value || !current(owner)
      || expectedRevision !== undefined && owner.studio.getWorkbookRevision(owner.unitId) !== expectedRevision) {
      throw new Error('数据、选区或编辑状态已改变，请重新读取数据区域。')
    }
  }
  function receiveDataPanelChange(): void {
    const owner = pending.value
    if (owner && current(owner)) options.refresh(owner.studio, owner.unitId, () => current(owner))
  }
  async function sortSelection(ascending: boolean, extendSelection = false): Promise<void> {
    if (!canUseDataRange.value) return
    const owner = capture(), sheet = owner.studio.getWorkbookModel(owner.unitId)!.sheets.find(sheet => sheet.id === owner.sheetId)!
    const range = extendSelection ? expandSheetDataRange(sheet, owner.sourceRange) : owner.sourceRange
    if (range.endRow <= range.startRow) throw new RangeError('请选择至少两行的数据区域后再排序，或扩展到相邻连续区域。')
    publishRange(owner, range)
    const worksheet = FOffice.newAPI(owner.studio).getWorkbook(owner.unitId)?.getSheetBySheetId(owner.sheetId)
    if (!worksheet) return
    const version = ++sortGeneration, task = new AbortController()
    controller = task; options.busy.value = true; options.progress.value = '正在排序…'
    const assertCurrent = () => {
      task.signal.throwIfAborted()
      if (version !== sortGeneration || !current(owner, true)) throw new Error('排序目标已改变，旧结果没有应用。')
    }
    try {
      await worksheet.sortRange(dataPanelAddress(range), [{ column: owner.primary.column, direction: ascending ? 'ascending' : 'descending' }], {
        signal: task.signal, assertCurrent,
        onProgress: progress => { if (version === sortGeneration && current(owner)) options.progress.value = progress.phase === 'initializing'
          ? '正在准备排序…' : `排序：${progress.processed.toLocaleString()} / ${progress.total.toLocaleString()}` },
      })
      if (version === sortGeneration && current(owner)) options.refresh(owner.studio, owner.unitId, () => current(owner))
    } catch (failure) { if (!task.signal.aborted && version === sortGeneration) throw failure }
    finally { if (version === sortGeneration) { options.busy.value = false; controller = undefined } }
  }
  function runAutoFilterCommand(commandId: 'sheet.command.toggle-auto-filter' | 'sheet.command.clear-filter' | 'sheet.command.reapply-filter'): void {
    if (!canUseDataRange.value) return
    try {
      const owner = capture(), sheet = owner.studio.getWorkbookModel(owner.unitId)!.sheets.find(sheet => sheet.id === owner.sheetId)!
      const range = sheet.autoFilter?.range ?? (commandId === 'sheet.command.toggle-auto-filter' ? prepareSmartSheetFilterRange(sheet, owner.sourceRange) : undefined)
      if (!range) return
      assertSheetRangesEditable(sheet, [owner.sourceRange, ...sheetFilterPermissionRanges(sheet, range)], owner.userId)
      if (!current(owner, true)) return
      owner.studio.executeCommandSync(commandId, { unitId: owner.unitId, sheetId: owner.sheetId,
        expectedRevision: owner.revision, ...(commandId === 'sheet.command.toggle-auto-filter' ? { range } : {}) })
      if (current(owner)) options.refresh(owner.studio, owner.unitId, () => current(owner))
    } catch (failure) { options.notify('error', failure instanceof Error ? failure.message : String(failure)) }
  }
  return { canUseDataRange, canSort, panelEditable, openDataPanel, selectDataPanelSortRange,
    assertDataPanelCurrent, receiveDataPanelChange, sortSelection, runAutoFilterCommand, cancelSort }
}
