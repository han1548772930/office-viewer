import { watch } from 'vue'
import type { CellRange } from '../domain/model.ts'
import type { Disposable } from './events.ts'
import type { OfficeStudio } from './office-studio.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { applySheetFormatPainterCommand, captureSheetFormatPainterSource, previewSheetFormatPainterCommand,
  type SheetFormatPainterSource } from '../features/sheet-workbench/sheet-format-painter.ts'
import { formatPainterActionForCommand, SHEET_FORMAT_PAINTER_COMMAND_IDS, sheetFormatPainterCommandInput,
  type SheetFormatPainterCommandInput, type SheetFormatPainterState } from '../features/sheet-workbench/sheet-format-painter-session.ts'
import { assertSheetRangeChangedValuesValid } from '../features/sheet-workbench/sheet-range-editor-validation.ts'
import { clearModeForCommand } from '../features/sheet-workbench/sheet-clear-command.ts'
import { sheetRangesEqual } from '../features/sheet-workbench/sheet-range-command.ts'
import { canViewWorksheetProtection } from '../features/sheet-protection/range-protection.ts'

function apiRange(unitId: string, sheetId: string, range: CellRange) {
  return { unitId, sheetId, startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol }
}
function applyCapturedFormat(
  api: OfficeStudio, input: SheetFormatPainterCommandInput, source: SheetFormatPainterSource, state: SheetFormatPainterState,
): boolean {
  const target = resolveSheetRangeCommandTarget(api, input)
  if (target.unitId !== state.unitId || !target.ranges.length) return false
  if (input.expectedRevision !== undefined && api.getWorkbookRevision(target.unitId) !== input.expectedRevision) throw new Error('工作簿已更改，请重新选择目标后应用格式刷。')
  const options = { mergeConfirmed: input.mergeConfirmed, userId: api.currentUserId, date1904: target.workbook.date1904,
    region: api.regionalSettings.region.value, styleContext: target.workbook }
  const preview = previewSheetFormatPainterCommand(target.sheet, source, target.ranges, options)
  if (preview.requiresConfirmation) {
    const request = { requestId: crypto.randomUUID(), sessionId: state.sessionId!, workbookId: target.unitId, sheetId: target.sheetId,
      sourceSheetId: state.sourceSheetId!, source: structuredClone(source), baseRevision: api.getWorkbookRevision(target.unitId),
      discardedContentCount: preview.discardedContentCount,
      input: { ...input, unitId: target.unitId, sheetId: target.sheetId, ranges: target.ranges.map(range => ({ ...range })), sessionId: state.sessionId } }
    api.sheetFormatPainter.requestConfirmation(request)
    if (api.sheetFormatPainter.state.sessionId !== state.sessionId || api.sheetFormatPainter.state.confirmationId !== request.requestId) return false
    api.events.emit('SheetFormatPainterConfirmationRequested', structuredClone(request))
    return false
  }
  let contentChanged = false
  const changed = api.commitWorkbook(target.unitId, '应用格式刷', model => {
    const sheet = model.sheets.find(candidate => candidate.id === target.sheetId)!
    const result = applySheetFormatPainterCommand(sheet, source, target.ranges, { ...options, styleContext: model })
    contentChanged = result.contentChanged
    if (result.contentChanged) assertSheetRangeChangedValuesValid(target.sheet, model, target.sheetId)
  }, { recalculate: preview.added.length > 0 || preview.removed.length > 0 })
  if (changed && contentChanged) api.events.emit('SheetValueChanged', { workbookId: target.unitId, sheetId: target.sheetId,
    ranges: [...preview.ranges, ...preview.added, ...preview.removed].map(range => apiRange(target.unitId, target.sheetId, range)) })
  const selection = api.getSelection(), original = target.ranges.at(-1), expanded = preview.ranges.at(-1)
  if (original && expanded && !sheetRangesEqual(original, expanded) && selection?.unitId === target.unitId && selection.sheetId === target.sheetId
    && sheetRangesEqual(original, { startRow: selection.startRow, endRow: selection.endRow, startCol: selection.startColumn, endCol: selection.endColumn })) {
    api.setSelection(target.unitId, target.sheetId, apiRange(target.unitId, target.sheetId, expanded))
  }
  // A successful redundant paint consumes ONCE, but never adds content history.
  return true
}

/** All brush entry points share the same immutable session and document transaction. */
export function registerSheetFormatPainterCommands(api: OfficeStudio): Disposable {
  const painter = api.sheetFormatPainter
  const registrations = SHEET_FORMAT_PAINTER_COMMAND_IDS.map(id => api.registerCommand<SheetFormatPainterCommandInput | undefined, boolean>(id, params => {
    const input = sheetFormatPainterCommandInput(params), action = formatPainterActionForCommand(id, input)
    if (!action) return false
    const state = painter.state
    if (action.kind === 'cancel') {
      if (input.unitId !== undefined && input.unitId !== state.unitId) return false
      painter.cancel()
      return true
    }
    if (action.kind === 'cancel-confirmation') return painter.cancelConfirmation(input.requestId)
    if (action.kind === 'confirm') {
      const request = painter.getConfirmation()
      if (!input.requestId || !request || request.requestId !== input.requestId || request.sessionId !== state.sessionId) throw new Error('格式刷确认请求已经失效。')
      if (input.unitId !== undefined && input.unitId !== request.workbookId) throw new Error('目标工作簿与格式刷请求不一致。')
      if (input.expectedRevision !== undefined && input.expectedRevision !== request.baseRevision) throw new Error('格式刷确认版本无效。')
      const confirmed = { ...request.input, expectedRevision: request.baseRevision, mergeConfirmed: true, sessionId: request.sessionId }
      return painter.apply(request.workbookId, (source, current) => applyCapturedFormat(api, confirmed, source, current), request.sessionId)
    }
    if (action.kind === 'apply') {
      if (state.mode === 'off') return false
      if (state.confirmationId) return false
      return painter.apply(input.unitId ?? api.activeWorkbookId, (source, current) => applyCapturedFormat(api, input, source, current), input.sessionId)
    }
    if (action.kind === 'set-mode' && action.mode === 'off') {
      if (state.mode !== 'off' && input.unitId !== undefined && input.unitId !== state.unitId) return false
      painter.cancel()
      return true
    }
    if (state.mode !== 'off') {
      if (input.unitId !== undefined && input.unitId !== state.unitId) return false
      return action.kind === 'activate' ? painter.activate(state.unitId!, state.sourceSheetId!, () => null, action.continuous, action.legacyToggle)
        : painter.setMode(state.unitId!, state.sourceSheetId!, () => null, action.mode)
    }
    // Shared UI operations never fall through from a Docs/object context to stale cell coordinates.
    if (api.editor.state.activeKind !== 'workbook') return false
    const target = resolveSheetRangeCommandTarget(api, input), range = target.ranges.at(-1)
    if (!range) return false
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    const capture = () => captureSheetFormatPainterSource(target.sheet, range, api.currentUserId, target.workbook)
    return action.kind === 'activate' ? painter.activate(target.unitId, target.sheetId, capture, action.continuous, action.legacyToggle)
      : painter.setMode(target.unitId, target.sheetId, capture, action.mode)
  }))
  const stopState = watch(() => painter.state, state => api.events.emit('SheetFormatPainterStateChanged', { state: { ...state, sourceRange: state.sourceRange ? { ...state.sourceRange } : undefined } }), { flush: 'sync' })
  const stopContext = watch([() => api.editor.workbook.model, () => api.editor.state.activeKind], () => {
    const state = painter.state
    if (state.mode === 'off') return
    const workbook = api.editor.workbook.model, source = workbook.sheets.find(sheet => sheet.id === state.sourceSheetId)
    if (api.editor.state.activeKind !== 'workbook' || workbook.id !== state.unitId || !source || !canViewWorksheetProtection(source, api.currentUserId)) painter.cancel()
  }, { flush: 'sync' })
  const disposed = api.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === painter.state.unitId) painter.cancel() })
  const clear = api.events.addEvent('CommandExecuted', event => {
    const mode = clearModeForCommand(event.id), owner = painter.state.unitId
    if (owner && (mode === 'format' || mode === 'all') && (event.unitId ?? api.activeWorkbookId) === owner) painter.cancel()
  })
  return { dispose: () => { stopState(); stopContext(); disposed.dispose(); clear.dispose(); registrations.forEach(command => command.dispose()); painter.cancel() } }
}
