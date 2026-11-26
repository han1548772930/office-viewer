import { computed, nextTick, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { createBlankWorkbook, type CellRange, type WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { applyExcelWorkbookView, toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { sheetRangeCommandInputRanges } from '../../features/sheet-workbench/sheet-range-command.ts'
import { assertSheetCellShiftConfirmationAllowed, planWorkbookCellShiftCommand, SHEET_CELL_SHIFT_CONFIRM_COMMANDS,
  sheetCellShiftCommandInput, sheetCellShiftCommandRange, sheetCellShiftDirectionForCommand, sheetCellShiftEdit,
  type SheetCellShiftCommandInput, type SheetCellShiftConfirmationRequest } from '../../features/sheet-workbench/sheet-cell-shift-command.ts'
import { captureSheetAxisCommandHost, type SheetAxisHostChangeOptions } from './sheet-axis-command-host.ts'

interface SheetCellShiftUiContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  ranges(): CellRange[]
  activeRangeIndex(): number
  userId(): string
  region(): string
  active(): boolean
  finishEdit(): boolean
  prepareView(view: ExcelWorkbookView, previous: ExcelWorkbookView): ExcelWorkbookView
  applyView(view: ExcelWorkbookView): void
  selection(range: CellRange): void
  changed(options?: SheetAxisHostChangeOptions): void
  focus(): void
}
interface PendingCellShift {
  request: SheetCellShiftConfirmationRequest
  host: ReturnType<typeof captureSheetAxisCommandHost>
  signature?: string
  model?: WorkbookModel
  generation: number
}

/** Shared command dispatch and a captured confirmation; neither host owns a second content engine. */
export function useSheetCellShift(context: SheetCellShiftUiContext) {
  const pending = shallowRef<PendingCellShift | null>(null), cellShiftError = ref('')
  const request = computed(() => pending.value?.request ?? null)
  let generation = 0, closeCurrent: (() => boolean) | undefined
  function cancelCellShift(): void {
    if (pending.value) {
      const hostCurrent = pending.value.host.publishedCurrent(), closedGeneration = generation
      closeCurrent = () => closedGeneration === generation && hostCurrent()
    }
    pending.value = null; cellShiftError.value = ''
  }
  watch([context.studio, () => context.workbook().id ?? context.workbook().name, context.sheetId, context.active, context.userId], () => {
    generation++; cancelCellShift(); closeCurrent = undefined
  }, { flush: 'sync' })
  watch(context.studio, (studio, _previous, onCleanup) => {
    if (!studio) return
    const listener = studio.events.addEvent('SheetCellShiftConfirmationRequested', event => {
      if (!context.active() || event.workbookId !== context.workbook().id || event.sheetId !== context.sheetId()
        || studio.getWorkbookRevision(event.workbookId) !== event.baseRevision) return
      const host = captureSheetAxisCommandHost(context)
      if (!host.sameTarget()) return
      const request = cloneModel(event)
      request.input.host = { unitId: event.workbookId, sheetId: event.sheetId }
      pending.value = { request, host, model: studio.getWorkbookModel(event.workbookId), generation }; closeCurrent = undefined; cellShiftError.value = ''
    })
    onCleanup(() => listener.dispose())
  }, { immediate: true, flush: 'sync' })
  function afterCellShiftClose(): void {
    const current = closeCurrent; closeCurrent = undefined
    void nextTick(() => { if (!pending.value && current?.()) context.focus() })
  }
  function runCellShiftCommand(commandId: string, params?: unknown): boolean {
    const direction = sheetCellShiftDirectionForCommand(commandId)
    if (!direction || !context.active()) return false
    const input = sheetCellShiftCommandInput(params)
    if (pending.value && !input.dismantlePartialMerges || !context.finishEdit() || !context.active()) return false
    const host = captureSheetAxisCommandHost(context), { view, studio } = host, hostGeneration = generation
    const sameHost = () => hostGeneration === generation && host.sameTarget()
    if (!sameHost()) return false
    const sheetId = input.sheetId ?? input.subUnitId ?? context.sheetId(), explicit = sheetRangeCommandInputRanges(input)
    const ranges = explicit ?? context.ranges(), range = sheetCellShiftCommandRange(input, ranges, context.activeRangeIndex())
    if (!range) return false
    if (sheetId !== context.sheetId() || input.unitId !== undefined && input.unitId !== view.id) throw new Error('移位目标与当前工作台不一致。')
    const request: SheetCellShiftCommandInput = { ...input, unitId: view.id, sheetId, range: { ...range }, ranges: cloneModel(ranges) }
    if (studio && view.id) {
      const baseRevision = studio.getWorkbookRevision(view.id)
      request.expectedRevision ??= baseRevision
      request.expectedHistoryGeneration ??= studio.editor.history.captureSavepoint(view.id).generation
      request.host = { unitId: view.id, sheetId }
      const changed = studio.executeCommandSync<SheetCellShiftCommandInput, boolean>(commandId, request)
      if (!changed) return false
      if (!sameHost() || studio.getWorkbookRevision(view.id) !== baseRevision + 1) return true
      const model = studio.getWorkbookModel(view.id)
      if (!model) return true
      const next = toExcelWorkbookView(model)
      if (!sameHost() || studio.getWorkbookRevision(view.id) !== baseRevision + 1) return true
      const published = JSON.stringify(next)
      context.applyView(next)
      const publishedCurrent = host.publishedCurrent(baseRevision + 1)
      const current = () => sameHost() && publishedCurrent() && JSON.stringify(context.workbook()) === published
      if (current()) context.changed({ syncSelection: false, current })
      return true
    }
    const signature = JSON.stringify(view), model = applyExcelWorkbookView(createBlankWorkbook(), view, context.region())
    const sheet = model.sheets.find(item => item.id === sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, context.userId())) throw new Error('找不到可访问的移位工作表。')
    if (commandId === SHEET_CELL_SHIFT_CONFIRM_COMMANDS[direction]) assertSheetCellShiftConfirmationAllowed(model, sheetId, sheetCellShiftEdit(range, direction))
    const plan = planWorkbookCellShiftCommand(model, sheetId, range, direction, context.userId(), input.dismantlePartialMerges, ranges)
    if (!sameHost() || !host.sourceCurrent() || JSON.stringify(context.workbook()) !== signature) return false
    if (plan.requiresConfirmation) {
      pending.value = { request: { requestId: crypto.randomUUID(), workbookId: view.id ?? '', sheetId, commandId, baseRevision: 0,
        mergeRanges: cloneModel(plan.mergeRanges), input: { ...request, ranges: [cloneModel(range)] } }, host, signature, generation }
      closeCurrent = undefined; cellShiftError.value = ''; return false
    }
    if (JSON.stringify(model) === JSON.stringify(plan.workbook)) return false
    const next = context.prepareView(toExcelWorkbookView(plan.workbook), cloneModel(view))
    if (!sameHost() || !host.sourceCurrent() || JSON.stringify(context.workbook()) !== signature) return false
    if (view.id === undefined) delete next.id; else next.id = view.id
    const published = JSON.stringify(next)
    context.applyView(next)
    if (!sameHost() || JSON.stringify(context.workbook()) !== published) return true
    const publishedCurrent = host.publishedCurrent(), current = () => sameHost() && publishedCurrent()
    context.selection(plan.range)
    if (current()) context.changed({ current })
    return true
  }
  function confirmCellShift(): void {
    const state = pending.value
    if (!state || !context.active()) return
    try {
      if (state.generation !== generation || !state.host.sameTarget()) throw new Error('移位宿主已切换，请重新选择范围。')
      if (state.model && state.host.studio?.getWorkbookModel(state.request.workbookId) !== state.model) throw new Error('工作簿实例已更改，请重新选择移位范围。')
      if (state.signature !== undefined && !state.host.sourceCurrent()) throw new Error('移位工作簿宿主已更改，请重新选择范围。')
      if (state.signature !== undefined && state.signature !== JSON.stringify(context.workbook())) throw new Error('工作簿已更改，请重新选择移位范围。')
      const changed = runCellShiftCommand(state.request.commandId, { ...state.request.input, dismantlePartialMerges: true })
      if (changed && pending.value === state) cancelCellShift()
    } catch (error: unknown) { if (pending.value === state) cellShiftError.value = error instanceof Error ? error.message : '单元格移位未能提交。' }
  }
  return { pending, request, cellShiftError, runCellShiftCommand, confirmCellShift, cancelCellShift, afterCellShiftClose }
}
