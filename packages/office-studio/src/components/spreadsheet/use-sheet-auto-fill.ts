import { computed, nextTick } from 'vue'
import type { CellRange } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { createSheetAutoFillPlan, SHEET_AUTO_FILL_COMMAND, SHEET_REFILL_COMMAND, SHEET_COPY_DOWN_COMMAND,
  SHEET_COPY_RIGHT_COMMAND, type SheetAutoFillResult } from '../../features/sheet-workbench/sheet-auto-fill.ts'
import { parseSheetFillCommandInput, type SheetFillCommandInput } from '../../features/sheet-workbench/sheet-fill-command.ts'
import { sheetCopyFillRanges, sheetFillPointerTarget } from '../../features/sheet-workbench/sheet-fill-target.ts'
import { sheetRangesEqual } from '../../features/sheet-workbench/sheet-range-command.ts'
import { sheetFillSessionIsCurrent, type SheetFillOptionsView } from '../../features/sheet-workbench/sheet-fill-store.ts'
import type { SheetFillApplyType } from '../../features/sheet-workbench/sheet-fill.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { captureSheetClipboardHost, type SheetClipboardHostContext } from './sheet-clipboard-ui-host.ts'

export interface SheetAutoFillUiContext extends SheetClipboardHostContext {
  userId(): string
  finishEdit(): boolean
  prepareView(view: ExcelWorkbookView, previous: ExcelWorkbookView): ExcelWorkbookView
  applyView(view: ExcelWorkbookView): void
  selection(range: CellRange): void
  changed(options: { prepared: true; syncSelection?: boolean; current: () => boolean }): void
  focus(): void
  notify(value: { type: 'info' | 'warning' | 'error'; message: string }): void
}

let nextStandaloneFill = 0
const modelRange = (range: GridRange): CellRange => ({ startRow: range.startRow, endRow: range.endRow,
  startCol: range.startColumn, endCol: range.endColumn })

/** UI only captures its host and publishes a prepared view. The domain planner
 * and Studio commands own fill semantics; this controller has no undo stack. */
export function useSheetAutoFill(context: SheetAutoFillUiContext) {
  const standaloneUnitId = `sheet-fill-standalone-${++nextStandaloneFill}`
  let requestNumber = 0
  const fillOptions = computed<SheetFillOptionsView | undefined>(() => {
    const studio = context.studio(), unitId = context.workbook().id
    if (!studio || !unitId || !context.active() || context.editing()) return undefined
    const session = studio.sheetFill.getView(unitId)
    const selection = context.range()
    // Read the reactive workbook/history as well as the transient session.
    const model = studio.editor.workbook.model
    if (!session || session.sheetId !== context.sheetId() || session.userId !== context.userId()
      || studio.activeWorkbookId !== unitId || studio.editor.state.activeKind !== 'workbook' || model.id !== unitId
      || model.activeSheetId !== session.sheetId || studio.getWorkbookRevision(unitId) !== session.contentRevision
      || !selection || !sheetRangesEqual(selection, session.range)
      || !sheetFillSessionIsCurrent(session, studio.editor.history)) return undefined
    return session
  })

  function report(result: SheetAutoFillResult): void {
    const issue = result.issues.find(item => item.severity === 'stop') ?? result.issues[0]
    if (issue) context.notify({ type: issue.severity === 'stop' ? 'error' : issue.severity === 'warning' ? 'warning' : 'info',
      message: `${issue.title}：${issue.message}` })
    else if (!result.accepted && result.reason && result.reason !== 'same-range') {
      const messages = { 'empty-source': '填充源或目标没有可处理的可见单元格。', 'stale-target': '填充目标已更改，请重新操作。',
        'stale-fill': '上次填充之后内容或历史已更改，不能再切换其填充方式。', validation: '数据验证拒绝了本次填充。' }
      context.notify({ type: 'warning', message: messages[result.reason] })
    }
  }

  function runFillCommand(commandId: string, params?: unknown): boolean {
    const owner = context.studio(), unitId = context.workbook().id, sheetId = context.sheetId()
    if (!context.active() || !context.finishEdit() || context.studio() !== owner
      || context.workbook().id !== unitId || context.sheetId() !== sheetId) return false
    const host = captureSheetClipboardHost(context, standaloneUnitId)
    if (!host) return false
    const task = ++requestNumber
    try {
      const input = parseSheetFillCommandInput(params)
      if (input.unitId !== undefined && input.unitId !== host.unitId
        || (input.sheetId ?? input.subUnitId) !== undefined && (input.sheetId ?? input.subUnitId) !== host.sheetId) return false
      if (host.studio) {
        const studio = host.studio, request: SheetFillCommandInput = { ...input, unitId: host.unitId, sheetId: host.sheetId,
          range: input.range ?? host.range, expectedRevision: input.expectedRevision ?? host.revision,
          expectedHistoryGeneration: input.expectedHistoryGeneration ?? host.historyGeneration,
          host: { unitId: host.unitId, sheetId: host.sheetId } }
        const accepted = studio.executeCommandSync<SheetFillCommandInput, boolean>(commandId, request)
        const result = studio.sheetFill.getResult(host.unitId)
        if (result) report(result)
        if (!accepted || !result) return false
        if (!result.changed) {
          const current = () => task === requestNumber && host.current()
          if (current() && host.selectionUnchanged() && result.selection) context.selection(result.selection)
          void nextTick(() => { if (current() && host.focusUnchanged()) context.focus() })
          return true
        }
        const model = studio.getWorkbookModel(host.unitId), revision = host.revision! + 1
        const current = () => host.sameTarget() && task === requestNumber && studio.getWorkbookModel(host.unitId) === model
          && studio.getWorkbookRevision(host.unitId) === revision
          && studio.editor.history.captureSavepoint(host.unitId).generation === host.historyGeneration
        if (!model || !current()) return accepted
        const next = toExcelWorkbookView(model), signature = JSON.stringify(next)
        context.applyView(next)
        const published = context.workbook()
        const publishedCurrent = () => current() && context.workbook() === published && JSON.stringify(published) === signature
        if (publishedCurrent()) context.changed({ prepared: true, syncSelection: false, current: publishedCurrent })
        void nextTick(() => { if (publishedCurrent() && host.focusUnchanged()) context.focus() })
        return accepted
      }

      // Controlled standalone hosts keep their existing change/undo contract.
      // Refill requires a shared history amendment receipt, supplied by Studio.
      if (commandId === SHEET_REFILL_COMMAND) return false
      let source = input.sourceRange, target = input.targetRange
      if (commandId === SHEET_COPY_DOWN_COMMAND || commandId === SHEET_COPY_RIGHT_COMMAND) {
        const copy = sheetCopyFillRanges(input.range ?? host.range, commandId === SHEET_COPY_DOWN_COMMAND ? 'down' : 'right')
        if (!copy) return false
        source = copy.sourceRange; target = copy.targetRange; input.applyType = 'COPY'
      }
      if (!source || !target) throw new TypeError('自动填充需要源范围和目标范围。')
      const plan = createSheetAutoFillPlan(host.model, host.sheetId, source, target,
        { applyType: input.applyType, userId: context.userId(), region: context.region() })
      report(plan.result)
      if (!plan.result.accepted || !plan.sheet || !plan.result.selection) return false
      if (equivalentModelData(plan.sheet, host.sheet)) {
        const current = () => task === requestNumber && host.current()
        if (current() && host.selectionUnchanged()) context.selection(plan.result.selection!)
        void nextTick(() => { if (current() && host.focusUnchanged()) context.focus() })
        return true
      }
      const model = { ...host.model, sheets: host.model.sheets.map(sheet => sheet.id === host.sheetId ? plan.sheet! : sheet) }
      const next = context.prepareView(toExcelWorkbookView(model), cloneModel(host.view))
      if (!host.current() || task !== requestNumber) return false
      if (host.view.id === undefined) delete next.id
      const signature = JSON.stringify(next)
      context.applyView(next)
      const published = context.workbook()
      const current = () => host.sameTarget() && task === requestNumber && context.workbook() === published && JSON.stringify(published) === signature
      if (current() && host.selectionUnchanged()) context.selection(plan.result.selection)
      if (current()) context.changed({ prepared: true, current })
      void nextTick(() => { if (current() && host.focusUnchanged()) context.focus() })
      return true
    } catch (error) {
      context.notify({ type: 'error', message: error instanceof Error ? error.message : '填充操作未能提交。' })
      return false
    }
  }

  function fillSelection(source: GridRange, pointer: { row: number; column: number }, applyType?: SheetFillApplyType): boolean {
    const sourceRange = modelRange(source)
    return runFillCommand(SHEET_AUTO_FILL_COMMAND, { sourceRange,
      targetRange: sheetFillPointerTarget(sourceRange, pointer.row, pointer.column), applyType })
  }
  function refill(option: { sessionId: string; applyType: SheetFillApplyType }): boolean {
    return runFillCommand(SHEET_REFILL_COMMAND, option)
  }
  function handleFillKeydown(event: KeyboardEvent): boolean {
    if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || event.altKey || event.shiftKey
      || !(event.ctrlKey || event.metaKey) || context.editing()) return false
    const key = event.key.toLowerCase()
    if (key !== 'd' && key !== 'r') return false
    event.preventDefault()
    runFillCommand(key === 'd' ? SHEET_COPY_DOWN_COMMAND : SHEET_COPY_RIGHT_COMMAND)
    return true
  }
  return { fillOptions, runFillCommand, fillSelection, refill, handleFillKeydown }
}
