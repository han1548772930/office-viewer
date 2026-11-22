import { createPinia } from 'pinia'
import { computed, nextTick, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { createBlankWorkbook, type CellRange } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { applyExcelWorkbookView, toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { applySheetViewFormatPainterCommand, captureSheetViewFormatPainterSource, createSheetFormatPainterPlan,
  type SheetFormatPainterCanvasState, type SheetFormatPainterSource } from '../../features/sheet-workbench/sheet-format-painter.ts'
import { formatPainterActionForCommand, SHEET_FORMAT_PAINTER_COMMANDS, sheetFormatPainterCommandInput,
  type SheetFormatPainterCommandInput, type SheetFormatPainterConfirmationRequest, type SheetFormatPainterState } from '../../features/sheet-workbench/sheet-format-painter-session.ts'
import { useSheetFormatPainterStore } from '../../features/sheet-workbench/sheet-format-painter-store.ts'
import { sheetRangeCommandInputRanges } from '../../features/sheet-workbench/sheet-range-command.ts'
import { assertSheetRangeChangedValuesValid } from '../../features/sheet-workbench/sheet-range-editor-validation.ts'
import { canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'

interface SheetFormatPainterContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  ranges(): CellRange[]
  userId(): string
  region(): string
  active(): boolean
  canStart(): boolean
  hasObjectSelection(): boolean
  applyView(view: ExcelWorkbookView): void
  selection(ranges: CellRange[]): void
  changed(recalculate: boolean): void
  focus(): void
  notify(message: string): void
}
type PendingPaint = { request: SheetFormatPainterConfirmationRequest; studio: OfficeStudio }
  | { request: SheetFormatPainterConfirmationRequest; signature: string }

/** Translates UI gestures to registered commands; it owns only confirmation and focus lifetimes. */
export function useSheetFormatPainter(context: SheetFormatPainterContext) {
  const standalone = useSheetFormatPainterStore(createPinia()), standaloneUnitId = crypto.randomUUID()
  const unitId = () => context.workbook().id ?? standaloneUnitId
  const painter = computed(() => context.workbook().id && context.studio() ? context.studio()!.sheetFormatPainter : standalone)
  const painterState = computed(() => painter.value.state)
  const painterActive = computed(() => painterState.value.mode !== 'off' && painterState.value.unitId === unitId())
  const painterUnavailable = computed(() => !context.active() || !context.canStart() || context.hasObjectSelection() || !context.ranges().length)
  const formatPainterCanvas = computed<SheetFormatPainterCanvasState | null>(() => {
    const state = painterState.value
    if (!painterActive.value || state.mode === 'off' || !state.sourceSheetId || !state.sourceRange) return null
    return { mode: state.mode, sourceSheetId: state.sourceSheetId, sourceRange: {
      startRow: state.sourceRange.startRow, endRow: state.sourceRange.endRow,
      startColumn: state.sourceRange.startCol, endColumn: state.sourceRange.endCol,
    } }
  })
  const pending = shallowRef<PendingPaint | null>(null), painterError = ref('')
  const painterRequest = computed(() => pending.value?.request ?? null)

  function clearPainterSession(): void { painter.value.cancel(); pending.value = null; painterError.value = '' }
  watch([context.studio, unitId, context.active] as const, ([studio, currentUnitId, active], [previousStudio]) => {
    if (previousStudio && previousStudio !== studio) previousStudio.sheetFormatPainter.cancel()
    standalone.cancel()
    if (!active || painterState.value.unitId !== currentUnitId) clearPainterSession()
    pending.value = null
    painterError.value = ''
  }, { flush: 'sync' })
  watch(context.sheetId, () => {
    if (pending.value) painter.value.cancelConfirmation(pending.value.request.requestId)
    pending.value = null
    painterError.value = ''
  }, { flush: 'sync' })
  watch(() => {
    const source = context.workbook().sheets.find(sheet => sheet.id === painterState.value.sourceSheetId)
    return source ? canViewWorksheetProtection(source, context.userId()) : false
  }, viewable => {
    if (painterActive.value && !viewable) clearPainterSession()
  }, { flush: 'sync' })
  watch(painterState, state => {
    if (pending.value && pending.value.request.requestId !== state.confirmationId) { pending.value = null; painterError.value = '' }
  }, { flush: 'sync' })
  watch(context.studio, (studio, _previous, onCleanup) => {
    if (!studio) return
    const confirmation = studio.events.addEvent('SheetFormatPainterConfirmationRequested', request => {
      if (!context.active() || request.workbookId !== context.workbook().id || request.sheetId !== context.sheetId()) return
      if (painterState.value.sessionId !== request.sessionId || painterState.value.confirmationId !== request.requestId) return
      pending.value = { studio, request: cloneModel(request) }
      painterError.value = ''
    })
    const disposed = studio.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === painterState.value.unitId) clearPainterSession() })
    const lifecycle = studio.events.addEvent('LifeCycleChanged', event => { if (event.stage === 'disposed') clearPainterSession() })
    onCleanup(() => { confirmation.dispose(); disposed.dispose(); lifecycle.dispose() })
  }, { immediate: true, flush: 'sync' })

  function standalonePaint(input: SheetFormatPainterCommandInput, source: SheetFormatPainterSource, state: SheetFormatPainterState, signature?: string): boolean {
    const current = context.workbook()
    if (signature !== undefined && JSON.stringify(current) !== signature) throw new Error('工作簿已更改，请取消后重新选择格式刷目标。')
    const targetId = input.sheetId ?? input.subUnitId ?? context.sheetId(), ranges = sheetRangeCommandInputRanges(input) ?? context.ranges()
    const candidate = cloneModel(current), sheet = candidate.sheets.find(item => item.id === targetId)
    if (!sheet) throw new Error('找不到目标工作表。')
    const result = applySheetViewFormatPainterCommand(sheet, source, ranges, { mergeConfirmed: input.mergeConfirmed,
      userId: context.userId(), date1904: candidate.date1904 ?? sheet.date1904, region: context.region() })
    if (result.requiresConfirmation) {
      const request: SheetFormatPainterConfirmationRequest = { requestId: crypto.randomUUID(), sessionId: state.sessionId!,
        workbookId: unitId(), sheetId: targetId, sourceSheetId: state.sourceSheetId!, source: cloneModel(source), baseRevision: 0,
        discardedContentCount: result.discardedContentCount,
        input: { ...input, unitId: unitId(), sheetId: targetId, ranges: ranges.map(range => ({ ...range })), sessionId: state.sessionId } }
      painter.value.requestConfirmation(request)
      if (painterState.value.sessionId !== request.sessionId || painterState.value.confirmationId !== request.requestId) return false
      pending.value = { request, signature: JSON.stringify(current) }
      painterError.value = ''
      return false
    }
    if (!result.applied) return false
    if (result.changed) {
      if (result.contentChanged) {
        const before = applyExcelWorkbookView(createBlankWorkbook(), current, context.region())
        assertSheetRangeChangedValuesValid(before.sheets.find(item => item.id === targetId)!, applyExcelWorkbookView(before, candidate, context.region()), targetId)
      }
      context.applyView(candidate)
      context.changed(result.contentChanged)
    }
    if (targetId === context.sheetId()) context.selection(result.ranges)
    return true
  }
  function runFormatPainterCommand(commandId: string, params?: unknown): boolean {
    const supplied = sheetFormatPainterCommandInput(params), action = formatPainterActionForCommand(commandId, supplied)
    if (!action || !context.active()) return false
    if ((action.kind === 'activate' || action.kind === 'set-mode' && action.mode !== 'off') && !painterActive.value && painterUnavailable.value) return false
    const input: SheetFormatPainterCommandInput = { ...supplied, unitId: supplied.unitId ?? unitId(),
      sheetId: supplied.sheetId ?? supplied.subUnitId ?? context.sheetId(), ranges: supplied.ranges ?? supplied.selections ?? (supplied.range ? [supplied.range] : context.ranges()) }
    const studio = context.studio()
    if (studio && context.workbook().id) {
      const isPaint = action.kind === 'apply' || action.kind === 'confirm', source = isPaint ? painter.value.getSource() : null
      const ranges = isPaint ? action.kind === 'confirm' ? sheetRangeCommandInputRanges(painter.value.getConfirmation()?.input) : sheetRangeCommandInputRanges(input) : undefined
      const actualRanges = source && ranges ? createSheetFormatPainterPlan([], source, ranges).ranges : undefined
      const revision = studio.getWorkbookRevision(input.unitId!)
      const accepted = studio.executeCommandSync<SheetFormatPainterCommandInput, boolean>(commandId, input)
      if (accepted && isPaint) {
        const model = studio.getWorkbookModel(input.unitId!)
        if (model) context.applyView(toExcelWorkbookView(model))
        if (studio.getWorkbookRevision(input.unitId!) !== revision) context.changed(false)
        if (actualRanges && input.sheetId === context.sheetId()) context.selection(actualRanges)
      }
      return accepted
    }
    if (action.kind === 'cancel') { clearPainterSession(); return true }
    if (action.kind === 'cancel-confirmation') return painter.value.cancelConfirmation(input.requestId)
    if (action.kind === 'confirm') {
      const request = painter.value.getConfirmation(), state = pending.value
      if (!request || !state || 'studio' in state || request.requestId !== input.requestId) throw new Error('格式刷确认请求已经失效。')
      return painter.value.apply(request.workbookId, (source, current) => standalonePaint({ ...request.input, mergeConfirmed: true }, source, current, state.signature), request.sessionId)
    }
    if (action.kind === 'apply') {
      if (painterState.value.confirmationId) return false
      return painter.value.apply(input.unitId!, (source, current) => standalonePaint(input, source, current), input.sessionId)
    }
    if (action.kind === 'set-mode' && action.mode === 'off') { clearPainterSession(); return true }
    const sheet = context.workbook().sheets.find(item => item.id === input.sheetId), range = sheetRangeCommandInputRanges(input)?.at(-1)
    if (!sheet || !range) return false
    const capture = () => captureSheetViewFormatPainterSource(sheet, range, context.userId())
    return action.kind === 'activate' ? painter.value.activate(input.unitId!, sheet.id, capture, action.continuous, action.legacyToggle)
      : painter.value.setMode(input.unitId!, sheet.id, capture, action.mode)
  }
  function armFormatPainter(continuous = false): void {
    try {
      if (runFormatPainterCommand(continuous ? SHEET_FORMAT_PAINTER_COMMANDS.continuous : SHEET_FORMAT_PAINTER_COMMANDS.activate)) context.focus()
    } catch (error: unknown) { context.notify(error instanceof Error ? error.message : '无法采样格式刷。') }
  }
  function cancelFormatPainter(): void {
    try { runFormatPainterCommand(SHEET_FORMAT_PAINTER_COMMANDS.cancel) }
    catch (error: unknown) { context.notify(error instanceof Error ? error.message : '无法取消格式刷。') }
  }
  function finishFormatPainterSelection(event: { canceled: boolean; ranges: GridRange[]; key: string }): void {
    if (event.canceled || !painterActive.value || !context.active() || context.hasObjectSelection()) return
    try {
      runFormatPainterCommand(SHEET_FORMAT_PAINTER_COMMANDS.apply, { sessionId: painterState.value.sessionId,
        ranges: event.ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn })) })
    } catch (error: unknown) { context.notify(error instanceof Error ? error.message : '格式刷应用失败，数据没有修改。') }
  }
  function cancelPainterConfirmation(): void {
    if (!pending.value) return
    try { runFormatPainterCommand(SHEET_FORMAT_PAINTER_COMMANDS.cancelConfirmation, { requestId: pending.value.request.requestId }) }
    catch (error: unknown) { painterError.value = error instanceof Error ? error.message : '无法取消格式刷确认。' }
  }
  function confirmPainter(): void {
    if (!pending.value) return
    try { runFormatPainterCommand(SHEET_FORMAT_PAINTER_COMMANDS.confirm, { requestId: pending.value.request.requestId, expectedRevision: pending.value.request.baseRevision }) }
    catch (error: unknown) { painterError.value = error instanceof Error ? error.message : '格式刷应用失败，数据没有修改。' }
  }
  function afterPainterClose(): void { void nextTick(() => { if (context.active() && !pending.value) context.focus() }) }
  return { painterActive, painterUnavailable, formatPainterCanvas, painterRequest, painterError,
    runFormatPainterCommand, armFormatPainter, cancelFormatPainter, clearPainterSession,
    finishFormatPainterSelection, confirmPainter, cancelPainterConfirmation, afterPainterClose }
}
