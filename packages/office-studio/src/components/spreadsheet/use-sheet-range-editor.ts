import { computed, nextTick, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import { createBlankWorkbook, type CellRange } from '../../domain/model.ts'
import { applyExcelWorkbookView, parseCellKey, toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { applySheetViewClearCommand, clearModeForCommand, type SheetClearCommandInput, type SheetClearOptions } from '../../features/sheet-workbench/sheet-clear-command.ts'
import { applySheetViewMergeCommand, mergeActionForCommand, type SheetMergeCommandInput, type SheetMergeConfirmationRequest } from '../../features/sheet-workbench/sheet-merge-command.ts'
import { assertSheetRangeChangedValuesValid } from '../../features/sheet-workbench/sheet-range-editor-validation.ts'
import { sheetRangeContainsCell, sheetRangesEqual, sheetRangesIntersect } from '../../features/sheet-workbench/sheet-range-command.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'

interface SheetRangeEditorContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  ranges(): CellRange[]
  primary(): { row: number; col: number }
  resetPrimary(): void
  userId(): string
  region(): string
  active(): boolean
  applyView(view: ExcelWorkbookView): void
  changed(recalculate: boolean): void
  focus(): void
}
type PendingMerge = { request: SheetMergeConfirmationRequest; studio: OfficeStudio }
  | { request: SheetMergeConfirmationRequest; signature: string }

function mergeUiInput(params: unknown): SheetMergeCommandInput {
  if (params === undefined) return {}
  if (params === 'all' || params === 'horizontal' || params === 'vertical') return { mode: params }
  if (!params || typeof params !== 'object' || Array.isArray(params)) throw new TypeError('合并参数无效。')
  const input: SheetMergeCommandInput = {}
  if ('mode' in params) {
    if (params.mode !== 'all' && params.mode !== 'horizontal' && params.mode !== 'vertical') throw new TypeError('合并方向无效。')
    input.mode = params.mode
  }
  if ('value' in params) {
    if (params.value !== 'all' && params.value !== 'horizontal' && params.value !== 'vertical' && params.value !== 0 && params.value !== 1) throw new TypeError('合并方向无效。')
    input.value = params.value
  }
  if ('defaultMerge' in params) {
    if (typeof params.defaultMerge !== 'boolean') throw new TypeError('合并选项必须是布尔值。')
    input.defaultMerge = params.defaultMerge
  }
  if ('isForceMerge' in params) {
    if (typeof params.isForceMerge !== 'boolean') throw new TypeError('合并选项必须是布尔值。')
    input.isForceMerge = params.isForceMerge
  }
  return input
}
function clearUiOptions(params: unknown): SheetClearOptions | undefined {
  if (params === undefined) return undefined
  if (!params || typeof params !== 'object' || Array.isArray(params)) throw new TypeError('清除参数无效。')
  if (!('options' in params)) return undefined
  const source = params.options
  if (!source || typeof source !== 'object' || Array.isArray(source)) throw new TypeError('清除选项无效。')
  const result: SheetClearOptions = {}
  for (const key of ['contentsOnly', 'formatOnly', 'contents', 'format'] as const) {
    if (!(key in source)) continue
    const value = source[key]
    if (typeof value !== 'boolean') throw new TypeError('清除选项必须是布尔值。')
    result[key] = value
  }
  return result
}

/** Owns only the UI request lifetime; range edits remain in shared commands and services. */
export function useSheetRangeEditor(context: SheetRangeEditorContext) {
  const pending = shallowRef<PendingMerge | null>(null)
  const mergeError = ref('')
  const mergeRequest = computed(() => pending.value?.request ?? null)
  const sheet = computed(() => context.workbook().sheets.find(candidate => candidate.id === context.sheetId()))
  const merges = computed(() => (sheet.value?.merges ?? []).flatMap(value => {
    const [start, end] = value.split(':'), first = parseCellKey(start ?? ''), last = parseCellKey(end ?? start ?? '')
    return first && last ? [{ startRow: first.row, startCol: first.col, endRow: last.row, endCol: last.col }] : []
  }))
  const affectedMerges = computed(() => merges.value.filter(merge => context.ranges().some(range => sheetRangesIntersect(range, merge))))
  const editable = computed(() => Boolean(sheet.value && canEditWorksheetProtection(sheet.value, context.userId())
    && canEditProtectedRanges(sheet.value, [...context.ranges(), ...affectedMerges.value], context.userId())))
  const canMerge = computed(() => {
    const ranges = context.ranges()
    return editable.value && ranges.some(range => range.startRow !== range.endRow || range.startCol !== range.endCol)
      && !ranges.some((range, index) => ranges.slice(0, index).some(previous => !sheetRangesEqual(previous, range) && sheetRangesIntersect(previous, range)))
  })
  const canUnmerge = computed(() => editable.value && affectedMerges.value.length > 0)
  const mergeActive = computed(() => {
    const primary = context.primary()
    return merges.value.some(merge => sheetRangeContainsCell(merge, primary.row, primary.col))
  })

  function cancelMerge(): void { pending.value = null; mergeError.value = '' }
  function afterMergeClose(): void {
    void nextTick(() => { if (context.active() && !pending.value) context.focus() })
  }
  watch([context.studio, () => context.workbook().id ?? context.workbook().name, context.sheetId, context.active], cancelMerge, { flush: 'sync' })
  watch(context.studio, (studio, _previous, onCleanup) => {
    if (!studio) return
    const confirmation = studio.events.addEvent('SheetMergeConfirmationRequested', request => {
      if (!context.active() || request.workbookId !== context.workbook().id || request.sheetId !== context.sheetId()) return
      pending.value = { request: structuredClone(request), studio }
      mergeError.value = ''
    })
    const disposed = studio.events.addEvent('WorkbookDisposed', event => {
      if (pending.value?.request.workbookId === event.workbookId) cancelMerge()
    })
    const lifecycle = studio.events.addEvent('LifeCycleChanged', event => { if (event.stage === 'disposed') cancelMerge() })
    onCleanup(() => { confirmation.dispose(); disposed.dispose(); lifecycle.dispose() })
  }, { immediate: true, flush: 'sync' })

  function refreshStudioView(studio: OfficeStudio, unitId: string): void {
    const model = studio.getWorkbookModel(unitId)
    if (model) context.applyView(toExcelWorkbookView(model))
  }
  function runClear(commandId: string, params?: unknown): boolean {
    if (!sheet.value || !context.ranges().length) return false
    const input: SheetClearCommandInput = { unitId: context.workbook().id, sheetId: sheet.value.id, ranges: context.ranges(), options: clearUiOptions(params) }
    const mode = clearModeForCommand(commandId, input)
    if (!mode || mode === 'none') return false
    const studio = context.studio(), unitId = context.workbook().id
    let changed: boolean
    if (studio && unitId) {
      changed = studio.executeCommandSync<SheetClearCommandInput, boolean>(commandId, input)
      if (changed) refreshStudioView(studio, unitId)
    } else {
      const before = applyExcelWorkbookView(createBlankWorkbook(), context.workbook(), context.region())
      const candidate: ExcelWorkbookView = JSON.parse(JSON.stringify(context.workbook()))
      const target = candidate.sheets.find(item => item.id === sheet.value!.id)!
      changed = applySheetViewClearCommand(target, context.ranges(), mode, context.userId())
      if (changed) {
        if (mode !== 'format') assertSheetRangeChangedValuesValid(before.sheets.find(item => item.id === target.id)!, applyExcelWorkbookView(before, candidate, context.region()), target.id)
        context.applyView(candidate)
      }
    }
    if (changed) context.changed(!studio && mode !== 'format')
    return changed
  }
  function runStandaloneMerge(commandId: string, input: SheetMergeCommandInput, signature?: string): boolean {
    const current = context.workbook()
    if (signature !== undefined && JSON.stringify(current) !== signature) throw new Error('工作簿已更改，请取消后重新选择区域。')
    const targetId = input.sheetId ?? context.sheetId(), ranges = input.ranges?.map(range => 'startCol' in range ? range
      : { startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }) ?? context.ranges()
    const candidate: ExcelWorkbookView = JSON.parse(JSON.stringify(current))
    const target = candidate.sheets.find(item => item.id === targetId)
    if (!target) throw new Error('找不到目标工作表。')
    const action = mergeActionForCommand(commandId, input)
    if (!action) return false
    const result = applySheetViewMergeCommand(target, ranges, action, context.userId())
    if (result.requiresConfirmation) {
      pending.value = { signature: JSON.stringify(current), request: {
        requestId: crypto.randomUUID(), workbookId: current.id ?? '', sheetId: targetId, commandId,
        baseRevision: 0, discardedContentCount: result.discardedContentCount,
        input: { ...input, sheetId: targetId, ranges: ranges.map(range => ({ ...range })) },
      } }
      mergeError.value = ''
      return false
    }
    if (!result.changed) return false
    const before = applyExcelWorkbookView(createBlankWorkbook(), current, context.region())
    if (action.kind === 'merge') assertSheetRangeChangedValuesValid(before.sheets.find(item => item.id === targetId)!, applyExcelWorkbookView(before, candidate, context.region()), targetId)
    context.applyView(candidate)
    if (action.kind === 'unmerge') context.resetPrimary()
    context.changed(action.kind === 'merge')
    return true
  }
  function runMerge(commandId: string, params?: unknown): boolean {
    if (!sheet.value || !context.ranges().length) return false
    const input: SheetMergeCommandInput = { ...mergeUiInput(params), unitId: context.workbook().id,
      sheetId: sheet.value.id, ranges: context.ranges(), isForceMerge: true }
    const studio = context.studio(), unitId = context.workbook().id
    if (!studio || !unitId) return runStandaloneMerge(commandId, input)
    const changed = studio.executeCommandSync<SheetMergeCommandInput, boolean>(commandId, input)
    if (changed) { refreshStudioView(studio, unitId); context.changed(false) }
    return changed
  }
  function confirmMerge(): void {
    const state = pending.value
    if (!state || !context.active()) return
    const request = state.request
    try {
      const input: SheetMergeCommandInput = { ...request.input, defaultMerge: true, expectedRevision: request.baseRevision }
      if ('studio' in state) {
        if (state.studio !== context.studio() || context.workbook().id !== request.workbookId || context.sheetId() !== request.sheetId) throw new Error('目标工作簿已切换，请重新选择区域。')
        const changed = state.studio.executeCommandSync<SheetMergeCommandInput, boolean>(request.commandId, input)
        if (changed) { refreshStudioView(state.studio, request.workbookId); context.changed(false) }
      } else runStandaloneMerge(request.commandId, input, state.signature)
      cancelMerge()
    } catch (error: unknown) { mergeError.value = error instanceof Error ? error.message : '合并失败，数据没有修改。' }
  }

  return { canMerge, canUnmerge, mergeActive, mergeRequest, mergeError, runClear, runMerge, confirmMerge, cancelMerge, afterMergeClose }
}
