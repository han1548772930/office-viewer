import type { OfficeStudio } from './office-studio.ts'
import type { ApiSheetSelection, Disposable } from './events.ts'
import type { CellRange, WorkbookModel } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { equivalentModelData } from '../platform/commands/model-equivalence.ts'
import type { HistorySavepoint } from '../platform/history/history-types.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { normalizeSheetCommandRange, sheetRangesEqual, type SheetCommandRange } from '../features/sheet-workbench/sheet-range-command.ts'
import { copySheetImageObjects, pasteSheetImageObjects } from './sheet-image-object-clipboard.ts'
import { SHEET_IMAGE_OBJECT_COMMANDS, sheetImageObjectClipboardFromHtml, sheetImageObjectClipboardIdFromHtml } from '../features/sheet-images/sheet-image-object-clipboard.ts'
import { applySheetClipboardPastePlan, assertSheetClipboardWorkbookChangesEditable, captureSheetClipboard, createSheetClipboardPastePlan,
  cloneSheetClipboardWorkbookResources, parseSheetClipboardHtml, parseSheetClipboardText,
  resolveSheetClipboardCommand, SHEET_CLIPBOARD_COMMAND_IDS, sheetClipboardContent, sheetClipboardIdFromHtml, sheetClipboardPasteMode,
  sheetClipboardChangedColumnWidthRanges, sheetClipboardChangedValueRanges, normalizeSheetClipboardContentAliases,
  type SheetClipboardCommandInput, type SheetClipboardCopyResult, type SheetClipboardLocation, type SheetClipboardPasteResult,
  type SheetClipboardRecentPaste, type SheetClipboardSnapshot, type SheetClipboardWorkbookResources } from '../features/sheet-workbench/sheet-clipboard.ts'

const failure = (reason: NonNullable<SheetClipboardPasteResult['reason']>): SheetClipboardPasteResult => ({ accepted: false, changed: false, ranges: [], issues: [], reason })
const copyFailure = (): SheetClipboardCopyResult => ({ accepted: false, plain: '', html: '' })
function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('剪贴板命令参数无效。')
  return value as Record<string, unknown>
}
function commandRange(value: unknown): CellRange {
  return normalizeSheetCommandRange(record(value) as unknown as SheetCommandRange)
}

export function parseSheetClipboardCommandInput(params: unknown): SheetClipboardCommandInput {
  if (params === undefined) return {}
  const input = record(params)
  for (const key of ['unitId', 'sheetId', 'subUnitId', 'clipboardId', 'value', 'mode', 'type']) {
    const value = input[key]
    if (value !== undefined && (typeof value !== 'string' || ['unitId', 'sheetId', 'subUnitId', 'clipboardId'].includes(key) && !value.trim())) {
      throw new TypeError(`剪贴板参数 ${key} 无效。`)
    }
  }
  for (const key of ['expectedRevision', 'expectedHistoryGeneration']) if (input[key] !== undefined
    && (!Number.isSafeInteger(input[key]) || Number(input[key]) < (key === 'expectedRevision' ? 0 : 1))) throw new TypeError('剪贴板内容或历史版本无效。')
  if (input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('工作表标识不一致。')
  const content = normalizeSheetClipboardContentAliases(input)
  if (input.imageIds !== undefined && (!Array.isArray(input.imageIds) || !input.imageIds.length
    || input.imageIds.some(id => typeof id !== 'string' || !id.trim()) || new Set(input.imageIds).size !== input.imageIds.length)) throw new TypeError('图片对象选区无效。')
  const result: SheetClipboardCommandInput = {
    unitId: input.unitId as string | undefined, sheetId: input.sheetId as string | undefined, subUnitId: input.subUnitId as string | undefined,
    clipboardId: input.clipboardId as string | undefined, ...content, value: input.value as string | undefined,
    ...(input.imageIds === undefined ? {} : { imageIds: [...input.imageIds as string[]] }),
    type: input.type as SheetClipboardCommandInput['type'],
    expectedRevision: input.expectedRevision as number | undefined, expectedHistoryGeneration: input.expectedHistoryGeneration as number | undefined,
    ...(input.mode === undefined && input.value === undefined && input.type === undefined ? {} : {
      mode: sheetClipboardPasteMode(input.mode as string | undefined, input.value as string | undefined, undefined, input.type as string | undefined),
    }),
    ...(input.range === undefined ? {} : { range: commandRange(input.range) }),
  }
  if (input.target !== undefined) {
    const target = record(input.target)
    if (typeof target.unitId !== 'string' || !target.unitId.trim() || typeof (target.sheetId ?? target.subUnitId) !== 'string'
      || !String(target.sheetId ?? target.subUnitId).trim() || target.sheetId !== undefined && target.subUnitId !== undefined && target.sheetId !== target.subUnitId) {
      throw new TypeError('剪贴板目标标识无效。')
    }
    const location = { unitId: target.unitId, sheetId: String(target.sheetId ?? target.subUnitId), range: commandRange(target.range) }
    if (result.unitId !== undefined && result.unitId !== location.unitId || (result.sheetId ?? result.subUnitId) !== undefined
      && (result.sheetId ?? result.subUnitId) !== location.sheetId || result.range !== undefined && !sheetRangesEqual(commandRange(result.range), location.range)) {
      throw new TypeError('剪贴板目标参数不一致。')
    }
    result.target = location
  }
  if (input.host !== undefined) {
    const host = record(input.host)
    if (typeof host.unitId !== 'string' || !host.unitId.trim() || typeof host.sheetId !== 'string' || !host.sheetId.trim()) throw new TypeError('剪贴板宿主标识无效。')
    result.host = { unitId: host.unitId, sheetId: host.sheetId }
  }
  return result
}

function locationFor(api: OfficeStudio, input: SheetClipboardCommandInput): SheetClipboardLocation | undefined {
  if (input.target) return { unitId: input.target.unitId, sheetId: input.target.sheetId ?? input.target.subUnitId!, range: commandRange(input.target.range) }
  const unitId = input.unitId ?? api.activeWorkbookId, workbook = api.getWorkbookModel(unitId)
  if (!workbook) return undefined
  const sheetId = input.sheetId ?? input.subUnitId ?? workbook.activeSheetId
  const selection = api.getSelectionState(unitId, sheetId), active = selection?.ranges[selection.activeRangeIndex]
  const range = input.range === undefined ? active && commandRange(active) : commandRange(input.range)
  return range ? { unitId, sheetId, range } : undefined
}

function currentTarget(api: OfficeStudio, input: SheetClipboardCommandInput, target: SheetClipboardLocation): boolean {
  const workbook = api.getWorkbookModel(target.unitId)
  if (!workbook?.sheets.some(sheet => sheet.id === target.sheetId)) return false
  const history = api.editor.history.captureSavepoint(target.unitId)
  if (input.expectedRevision !== undefined && input.expectedRevision !== api.getWorkbookRevision(target.unitId)
    || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== history.generation) return false
  if (input.host && (input.host.unitId !== target.unitId || input.host.sheetId !== target.sheetId
    || api.editor.state.activeKind !== 'workbook' || api.activeWorkbookId !== input.host.unitId || workbook.activeSheetId !== input.host.sheetId)) return false
  return true
}

function pasteSource(api: OfficeStudio, input: SheetClipboardCommandInput): SheetClipboardSnapshot | undefined {
  const supplied = input.text !== undefined || input.html !== undefined
  if (!supplied) return api.sheetClipboard.getClipboard(input.clipboardId)
  const token = input.html && sheetClipboardIdFromHtml(input.html)
  if (token && (input.clipboardId === undefined || input.clipboardId === token)) {
    const cached = api.sheetClipboard.getClipboard(token)
    if (cached) return cached
  }
  return input.html ? parseSheetClipboardHtml(input.html) ?? parseSheetClipboardText(input.text ?? '') : parseSheetClipboardText(input.text ?? '')
}

function workbookResources(workbook: WorkbookModel): SheetClipboardWorkbookResources {
  return cloneSheetClipboardWorkbookResources(workbook)
}
function restoreRecentOriginals(current: WorkbookModel, recent: SheetClipboardRecentPaste): WorkbookModel {
  const workbook = cloneModel(current), originals = new Map(recent.originalSheets.map(sheet => [sheet.id, sheet]))
  workbook.sheets = workbook.sheets.map(sheet => originals.has(sheet.id) ? cloneModel(originals.get(sheet.id)!) : sheet)
  if (recent.metadataChanged) {
    if (recent.originalMetadata === undefined) delete workbook.metadata
    else workbook.metadata = cloneModel(recent.originalMetadata)
  }
  if (recent.originalResources) for (const key of ['tables', 'pivotTables', 'pivotCaches', 'charts', 'slicers', 'slicerCaches'] as const) {
    const value = recent.originalResources[key]
    if (value === undefined) delete workbook[key]
    else Object.assign(workbook, { [key]: cloneModel(value) })
  }
  return workbook
}

function pasteSelection(api: OfficeStudio, input: SheetClipboardCommandInput, target: SheetClipboardLocation, ranges: readonly CellRange[]):
  { sheetId: string; before: ApiSheetSelection; after: ApiSheetSelection } | undefined {
  const current = api.getSelectionState(target.unitId, target.sheetId), active = current?.ranges[current.activeRangeIndex]
  if (!current || !active || api.editor.state.activeKind !== 'workbook' || api.activeWorkbookId !== target.unitId
    || api.getWorkbookModel(target.unitId)?.activeSheetId !== target.sheetId) return undefined
  // A user may navigate while the browser clipboard is being read. Preserve that newer full selection.
  const requested = input.range !== undefined ? commandRange(input.range) : input.target ? commandRange(input.target.range) : target.range
  if (input.host && !sheetRangesEqual(commandRange(active), requested)) return undefined
  const pasted = ranges.map(range => ({ unitId: target.unitId, sheetId: target.sheetId, startRow: range.startRow, endRow: range.endRow,
    startColumn: range.startCol, endColumn: range.endCol }))
  return pasted.length ? { sheetId: target.sheetId, before: current, after: { ranges: pasted, activeRangeIndex: pasted.length - 1,
    primary: { row: pasted.at(-1)!.startRow, column: pasted.at(-1)!.startColumn } } } : undefined
}

/** All clipboard callers share model planning, permission checks and one content-history transaction. */
export function registerSheetClipboardCommands(api: OfficeStudio): Disposable {
  const registrations = SHEET_CLIPBOARD_COMMAND_IDS.map(id => api.registerCommand<unknown, SheetClipboardCopyResult | SheetClipboardPasteResult | boolean>(id, params => {
    const action = resolveSheetClipboardCommand(id)!, input = parseSheetClipboardCommandInput(params)
    if (action.action === 'cancel-copy') { api.sheetClipboard.cancelCopyMark(); return true }
    const recent = action.action === 'optional-paste' ? api.sheetClipboard.getRecentPaste() : undefined
    const target = action.action === 'optional-paste' ? recent?.target : locationFor(api, input)
    const copying = action.action === 'copy' || action.action === 'cut'
    if (!target || !currentTarget(api, input, target)) return copying ? copyFailure() : failure(action.action === 'optional-paste' ? 'stale-paste' : 'stale-target')
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    const current = api.getWorkbookModel(target.unitId)!, baseRevision = api.getWorkbookRevision(target.unitId)
    const baseSavepoint = api.editor.history.captureSavepoint(target.unitId), clipboardGeneration = api.sheetClipboard.clipboardGeneration
    if (copying && input.imageIds) return copySheetImageObjects(api, input, target, action.action === 'cut')
    if (copying) {
      const clipboard = captureSheetClipboard(current, target, { copyType: action.action === 'cut' ? 'cut' : 'copy', formulaOnly: !!action.formulaOnly,
        revision: baseRevision, historyGeneration: baseSavepoint.generation, userId: api.currentUserId })
      const content = sheetClipboardContent(clipboard)
      if (!currentTarget(api, input, target) || api.getWorkbookRevision(target.unitId) !== baseRevision) return copyFailure()
      api.sheetClipboard.setClipboard(clipboard)
      return { accepted: true, clipboard, ...content }
    }
    if (recent && (recent.target.unitId !== target.unitId || recent.target.sheetId !== target.sheetId
      || recent.contentRevision !== baseRevision || !api.editor.history.canAmend(target.unitId, recent.savepoint)
      || input.clipboardId !== undefined && input.clipboardId !== recent.clipboard.id
      || input.unitId !== undefined && input.unitId !== target.unitId || (input.sheetId ?? input.subUnitId) !== undefined
      && (input.sheetId ?? input.subUnitId) !== target.sheetId)) return failure('stale-paste')
    const supplied = input.text !== undefined || input.html !== undefined, objectToken = input.html ? sheetImageObjectClipboardIdFromHtml(input.html) : undefined
    const imageClipboard = action.action === 'paste' && (!supplied || objectToken && (input.clipboardId === undefined || input.clipboardId === objectToken))
      ? api.sheetClipboard.getImageClipboard(supplied ? objectToken : input.clipboardId)
        ?? (input.html ? sheetImageObjectClipboardFromHtml(input.html) : undefined) : undefined
    if (imageClipboard) {
      const mode = sheetClipboardPasteMode(input.mode, input.value, action.mode, input.type)
      return mode === 'all' || mode === 'exceptBorders' ? pasteSheetImageObjects(api, input, target, imageClipboard) : failure('empty-clipboard')
    }
    const clipboard = recent?.clipboard ?? pasteSource(api, input)
    if (!clipboard) return failure('empty-clipboard')
    let mode = sheetClipboardPasteMode(input.mode, input.value, action.mode, input.type)
    if (clipboard.formulaOnly && mode === 'all') mode = 'formulas'
    const original = recent ? restoreRecentOriginals(current, recent) : current
    const plan = createSheetClipboardPastePlan(original, target, clipboard, mode, { userId: api.currentUserId,
      region: api.regionalSettings.region.value, sourceHistoryGeneration: baseSavepoint.generation })
    if (!plan.result.accepted) return cloneModel(plan.result)
    const candidate = cloneModel(original)
    applySheetClipboardPastePlan(candidate, plan)
    assertSheetClipboardWorkbookChangesEditable(current, candidate, api.currentUserId)
    const differs = !equivalentModelData(current, candidate, ['modified', 'activeSheetId'])
    if (!differs) {
      if (recent) api.sheetClipboard.setRecentPaste({ ...recent, mode, ranges: cloneModel(plan.result.ranges) })
      if (clipboard.copyType === 'cut') api.sheetClipboard.consumeCut(clipboard.id)
      return { ...cloneModel(plan.result), changed: false }
    }
    const originals = new Map((recent?.originalSheets ?? []).map(sheet => [sheet.id, cloneModel(sheet)]))
    for (const sheet of original.sheets) if (!originals.has(sheet.id) && plan.sheetChanges.some(changed => changed.id === sheet.id)) originals.set(sheet.id, cloneModel(sheet))
    const originalMetadata = recent ? recent.originalMetadata : original.metadata
    let committed: HistorySavepoint | undefined
    const selection = pasteSelection(api, input, target, plan.result.ranges)
    const changed = api.commitWorkbook(target.unitId, '粘贴', draft => {
      for (const key of ['metadata', 'tables', 'pivotTables', 'pivotCaches', 'charts', 'slicers', 'slicerCaches'] as const) if (!Object.hasOwn(candidate, key)) delete draft[key]
      Object.assign(draft, cloneModel(candidate))
    }, {
      expectedRevision: baseRevision, expectedHistoryGeneration: baseSavepoint.generation, amend: recent?.savepoint,
      selections: selection, prepareView: draft => { toExcelWorkbookView(draft) }, onHistoryCommitted: savepoint => { committed = savepoint },
    })
    if (changed && committed && api.getWorkbookRevision(target.unitId) === baseRevision + 1
      && api.editor.history.canAmend(target.unitId, committed) && api.sheetClipboard.clipboardGeneration === clipboardGeneration) {
      api.sheetClipboard.setRecentPaste({ clipboard, target: cloneModel(target), ranges: cloneModel(plan.result.ranges), mode, originalSheets: [...originals.values()],
        originalMetadata: originalMetadata === undefined ? undefined : cloneModel(originalMetadata), metadataChanged: !!recent?.metadataChanged || plan.metadataChanged,
        originalResources: recent?.originalResources ?? (plan.workbookResources ? workbookResources(original) : undefined),
        savepoint: committed, contentRevision: baseRevision + 1 })
    }
    if (changed && clipboard.copyType === 'cut') api.sheetClipboard.consumeCut(clipboard.id)
    if (changed) {
      for (const sheet of candidate.sheets) {
        const before = current.sheets.find(item => item.id === sheet.id)!
        const valueRanges = sheetClipboardChangedValueRanges(before, sheet)
        if (valueRanges.length) api.publishObserverEvent('SheetValueChanged', () => ({ workbookId: target.unitId, sheetId: sheet.id,
          ranges: valueRanges.map(range => ({ unitId: target.unitId, sheetId: sheet.id,
            startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol })) }))
      }
      for (const sheet of candidate.sheets) {
        const before = current.sheets.find(item => item.id === sheet.id)!
        const ranges = sheetClipboardChangedColumnWidthRanges(before, sheet)
        if (ranges.length) api.publishObserverEvent('SheetDimensionsChanged', () => ({ workbookId: target.unitId, sheetId: sheet.id,
          axis: 'column', automatic: false, ranges: ranges.map(range => ({ unitId: target.unitId, sheetId: sheet.id,
            startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol })) }))
      }
    }
    return { ...cloneModel(plan.result), changed }
  }))
  for (const [action, id] of Object.entries(SHEET_IMAGE_OBJECT_COMMANDS)) registrations.push(api.registerCommand<unknown, SheetClipboardCopyResult | SheetClipboardPasteResult>(id, params => {
    const input = parseSheetClipboardCommandInput(params)
    if (action !== 'paste' && !input.imageIds?.length) throw new TypeError('请提供要复制或剪切的图片身份。')
    return api.executeCommandSync(`sheet.command.${action}`, input)
  }))
  return { dispose: () => registrations.forEach(registration => registration.dispose()) }
}
