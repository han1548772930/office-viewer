import type { OfficeStudio } from './office-studio.ts'
import type { ApiSheetSelection, Disposable } from './events.ts'
import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { equivalentModelData } from '../platform/commands/model-equivalence.ts'
import type { HistorySavepoint } from '../platform/history/history-types.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { createSheetAutoFillPlan, SHEET_REFILL_COMMAND, SHEET_COPY_DOWN_COMMAND,
  SHEET_COPY_RIGHT_COMMAND, SHEET_AUTO_CLEAR_COMMAND, SHEET_FILL_COMMAND_IDS, type SheetAutoFillResult } from '../features/sheet-workbench/sheet-auto-fill.ts'
import { parseSheetFillCommandInput, type SheetFillCommandInput } from '../features/sheet-workbench/sheet-fill-command.ts'
import { resolveSheetFillTarget, sheetCopyFillRanges } from '../features/sheet-workbench/sheet-fill-target.ts'
import { sheetFillSessionIsCurrent, type SheetFillSession } from '../features/sheet-workbench/sheet-fill-store.ts'
import { assertSheetRangesEditable, normalizeSheetCommandRange, sheetRangesEqual } from '../features/sheet-workbench/sheet-range-command.ts'

function failed(input: SheetFillCommandInput, reason: SheetAutoFillResult['reason']): SheetAutoFillResult {
  return { accepted: false, changed: false, ranges: [], issues: [], supportsSeries: false, applyType: input.applyType ?? 'COPY', reason }
}
function targetFor(api: OfficeStudio, input: SheetFillCommandInput): { workbook: WorkbookModel; sheet: SpreadsheetSheet; unitId: string } | undefined {
  const unitId = input.unitId ?? api.activeWorkbookId, workbook = api.getWorkbookModel(unitId)
  const sheet = workbook?.sheets.find(item => item.id === (input.sheetId ?? input.subUnitId ?? workbook.activeSheetId))
  return workbook && sheet ? { workbook, sheet, unitId } : undefined
}
function currentTarget(api: OfficeStudio, input: SheetFillCommandInput, unitId: string, sheetId: string): boolean {
  const workbook = api.getWorkbookModel(unitId), savepoint = api.editor.history.captureSavepoint(unitId)
  return Boolean(workbook?.sheets.some(sheet => sheet.id === sheetId)
    && (input.expectedRevision === undefined || input.expectedRevision === api.getWorkbookRevision(unitId))
    && (input.expectedHistoryGeneration === undefined || input.expectedHistoryGeneration === savepoint.generation)
    && (!input.host || input.host.unitId === unitId && input.host.sheetId === sheetId && api.editor.state.activeKind === 'workbook'
      && api.activeWorkbookId === unitId && workbook?.activeSheetId === sheetId))
}
function selectionAfter(unitId: string, sheetId: string, range: CellRange): ApiSheetSelection {
  return { ranges: [{ unitId, sheetId, startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol }],
    activeRangeIndex: 0, primary: { row: range.startRow, column: range.startCol } }
}
function initialSelection(api: OfficeStudio, unitId: string, sheetId: string, selection: CellRange,
  before?: ApiSheetSelection): { sheetId: string; before: ApiSheetSelection; after: ApiSheetSelection } | undefined {
  const current = api.getSelectionState(unitId, sheetId)
  if (!current || api.editor.state.activeKind !== 'workbook' || api.activeWorkbookId !== unitId
    || api.getWorkbookModel(unitId)?.activeSheetId !== sheetId) return undefined
  return { sheetId, before: cloneModel(before ?? current), after: selectionAfter(unitId, sheetId, selection) }
}
function originalForRefill(workbook: WorkbookModel, sheet: SpreadsheetSheet, recent: SheetFillSession): WorkbookModel {
  const original = cloneModel(recent.originalSheet)
  original.worksheetProtection = sheet.worksheetProtection ? cloneModel(sheet.worksheetProtection) : undefined
  original.rangeProtectionRules = sheet.rangeProtectionRules ? cloneModel(sheet.rangeProtectionRules) : undefined
  return { ...workbook, sheets: workbook.sheets.map(item => item.id === sheet.id ? original : item) }
}

function runFill(api: OfficeStudio, commandId: string, input: SheetFillCommandInput): SheetAutoFillResult {
  const resolved = targetFor(api, input)
  if (!resolved) return failed(input, 'stale-target')
  const { workbook, sheet, unitId } = resolved
  if (!currentTarget(api, input, unitId, sheet.id)) return failed(input, 'stale-target')
  api.assertCanViewWorksheet(unitId, sheet.id)
  const history = api.editor.history, userId = api.currentUserId, baseRevision = api.getWorkbookRevision(unitId)
  const recent = commandId === SHEET_REFILL_COMMAND ? api.sheetFill.getSession(unitId) : undefined
  if (commandId === SHEET_REFILL_COMMAND && (!recent || recent.sheetId !== sheet.id || recent.userId !== userId
    || input.sessionId !== undefined && input.sessionId !== recent.id || recent.contentRevision !== api.getWorkbookRevision(unitId)
    || !sheetFillSessionIsCurrent(recent, history))) return failed(input, 'stale-fill')
  let source = input.sourceRange ?? recent?.sourceRange, selection = input.targetRange ?? recent?.targetRange
  if (commandId === SHEET_COPY_DOWN_COMMAND || commandId === SHEET_COPY_RIGHT_COMMAND) {
    const selected = api.getSelectionState(unitId, sheet.id), range = input.range ?? selected?.ranges[selected.activeRangeIndex]
    if (!range) return failed(input, 'empty-source')
    const copy = sheetCopyFillRanges(normalizeSheetCommandRange(range), commandId === SHEET_COPY_DOWN_COMMAND ? 'down' : 'right')
    if (!copy) return failed(input, 'same-range')
    source = copy.sourceRange; selection = copy.targetRange; input = { ...input, applyType: 'COPY' }
  }
  if (commandId === SHEET_AUTO_CLEAR_COMMAND) {
    if (!input.clearRange || !input.selectionRange) throw new TypeError('自动清除需要 clearRange 和 selectionRange。')
    selection = input.selectionRange
    source = { startRow: Math.min(selection.startRow, input.clearRange.startRow), endRow: Math.max(selection.endRow, input.clearRange.endRow),
      startCol: Math.min(selection.startCol, input.clearRange.startCol), endCol: Math.max(selection.endCol, input.clearRange.endCol) }
    const clearing = resolveSheetFillTarget(source, selection)
    if (!clearing?.clear || !sheetRangesEqual(clearing.destination, input.clearRange)) throw new RangeError('自动清除范围与保留选区不一致。')
  }
  if (!source || !selection) throw new TypeError('自动填充需要 sourceRange 和 targetRange。')
  if (recent && (!sheetRangesEqual(source, recent.sourceRange) || !sheetRangesEqual(selection, recent.targetRange))) return failed(input, 'stale-fill')
  const base = recent ? originalForRefill(workbook, sheet, recent) : workbook
  const plan = createSheetAutoFillPlan(base, sheet.id, source, selection, { applyType: input.applyType ?? (recent ? recent.applyType : undefined),
    userId, region: api.regionalSettings.region.value })
  if (!plan.result.accepted || !plan.sheet || !plan.result.selection) return plan.result
  assertSheetRangesEditable(sheet, plan.result.ranges, userId)
  const originalSheet = recent?.originalSheet ?? cloneModel(sheet)
  const beforeSelection = recent?.beforeSelection ?? api.getSelectionState(unitId, sheet.id) ?? undefined
  const finalSelection = plan.result.selection, nextSheet = plan.sheet
  const showOptions = !plan.target?.clear && commandId !== SHEET_COPY_DOWN_COMMAND && commandId !== SHEET_COPY_RIGHT_COMMAND
  const sessionId = recent?.id ?? crypto.randomUUID()
  const unchanged = (): SheetAutoFillResult => {
    const savepoint = recent?.savepoint ?? history.captureSavepoint(unitId)
    const hasHistoryEntry = recent?.hasHistoryEntry ?? false
    const selections = initialSelection(api, unitId, sheet.id, finalSelection)
    if (selections) api.setSelections(unitId, sheet.id, selections.after.ranges, selections.after)
    const current = { unitId, savepoint, hasHistoryEntry }
    const canShow = showOptions && api.currentUserId === userId && currentTarget(api, input, unitId, sheet.id)
      && api.getWorkbookRevision(unitId) === baseRevision && sheetFillSessionIsCurrent(current, history)
    if (canShow) api.sheetFill.setSession({ id: sessionId, unitId, sheetId: sheet.id, userId, sourceRange: source!, targetRange: selection!,
      originalSheet, beforeSelection, applyType: plan.result.applyType, supportsSeries: plan.result.supportsSeries,
      contentRevision: baseRevision, savepoint, hasHistoryEntry })
    else api.sheetFill.clearSession(unitId)
    return { ...plan.result, changed: false, ...(canShow ? { sessionId } : {}) }
  }
  if (equivalentModelData(sheet, nextSheet)) return unchanged()
  let receipt: HistorySavepoint | undefined
  const changed = api.commitWorkbook(unitId, plan.target?.clear ? '清除填充区域内容' : '自动填充', draft => {
    if (!currentTarget(api, input, unitId, sheet.id) || api.currentUserId !== userId) throw new Error('填充宿主或权限上下文已更改。')
    const index = draft.sheets.findIndex(item => item.id === sheet.id)
    if (index < 0) throw new Error('填充工作表已移除。')
    draft.sheets[index] = cloneModel(nextSheet)
  }, { expectedRevision: input.expectedRevision, expectedHistoryGeneration: input.expectedHistoryGeneration,
    amend: recent?.hasHistoryEntry ? recent.savepoint : undefined, recalculate: plan.result.applyType === 'ONLY_FORMAT' ? false : undefined,
    selections: initialSelection(api, unitId, sheet.id, finalSelection, beforeSelection),
    prepareView: model => { toExcelWorkbookView(model) }, onHistoryCommitted: value => { receipt = value } })
  if (!changed || !receipt) return unchanged()
  const canShow = showOptions && api.currentUserId === userId && api.getWorkbookRevision(unitId) === baseRevision + 1
    && history.canAmend(unitId, receipt)
  if (canShow) api.sheetFill.setSession({ id: sessionId, unitId, sheetId: sheet.id, userId, sourceRange: source, targetRange: selection,
    originalSheet, applyType: plan.result.applyType, supportsSeries: plan.result.supportsSeries, savepoint: receipt,
    contentRevision: api.getWorkbookRevision(unitId), beforeSelection, hasHistoryEntry: true })
  else api.sheetFill.clearSession(unitId)
  api.publishObserverEvent('SheetValueChanged', () => ({ workbookId: unitId, sheetId: sheet.id, ranges: plan.result.ranges.map(range => ({
    unitId, sheetId: sheet.id, startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol })) }))
  return { ...plan.result, changed: true, ...(showOptions ? { sessionId } : {}) }
}

/** Canonical commands keep Univer's boolean result. Detailed issues and menu state live in the application store. */
export function registerSheetFillCommands(api: OfficeStudio): Disposable {
  const commands = SHEET_FILL_COMMAND_IDS.map(commandId => api.registerCommand<unknown, boolean>(commandId, params => {
    const input = parseSheetFillCommandInput(params)
    if (commandId !== SHEET_REFILL_COMMAND) api.sheetFill.clear(input.unitId ?? api.activeWorkbookId)
    const result = runFill(api, commandId, input)
    api.sheetFill.setResult(input.unitId ?? api.activeWorkbookId, result)
    return result.accepted
  }))
  return { dispose: () => { for (const command of commands) command.dispose(); api.sheetFill.clear() } }
}
