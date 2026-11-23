import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { assignSheetAxisSpan, sheetColumnWidth } from '../../domain/sheet-axis.ts'
import { validateSheetDataValidationChanges, type SheetValidationPosition } from '../sheet-validation/validation-engine.ts'
import { assertSheetClipboardArea, assertSheetClipboardSnapshot, captureSheetClipboard, sheetClipboardSourceFingerprint } from './sheet-clipboard-data.ts'
import { createSheetClipboardPastedCell } from './sheet-clipboard-cell.ts'
import { assertSheetClipboardWorkbookChangesEditable } from './sheet-clipboard-permissions.ts'
import { sheetClipboardPasteMode } from './sheet-clipboard-command.ts'
import { moveSheetClipboardCellReferences, moveSheetClipboardDefinedNames, type SheetClipboardMove } from './sheet-clipboard-references.ts'
import { moveSheetClipboardLinkedRuleReferences, pasteSheetClipboardRangeResources, trimSheetClipboardRangeResources } from './sheet-clipboard-resources.ts'
import { createSheetClipboardTargetGeometry } from './sheet-clipboard-target.ts'
import { pasteSheetClipboardWorkbookResources } from './sheet-clipboard-workbook-resources.ts'
import { assertSheetCommandRanges, assertSheetRangesEditable, sheetRangeContainsCell, sheetRangesEqual, sheetRangesIntersect } from './sheet-range-command.ts'
import type { SheetClipboardLocation, SheetClipboardPasteContext, SheetClipboardPastePlan, SheetClipboardSnapshot } from './sheet-clipboard-types.ts'
import type { SheetPasteMode } from './sheet-paste.ts'

export * from './sheet-clipboard-types.ts'
export * from './sheet-clipboard-data.ts'
export * from './sheet-clipboard-content.ts'
export * from './sheet-clipboard-command.ts'
export * from './sheet-clipboard-store.ts'
export * from './sheet-clipboard-changes.ts'
export * from './sheet-clipboard-permissions.ts'
export * from './sheet-clipboard-workbook-resources.ts'
export * from './sheet-clipboard-capabilities.ts'
export * from './sheet-clipboard-input.ts'

function failed(reason: NonNullable<SheetClipboardPastePlan['result']['reason']>): SheetClipboardPastePlan {
  return { sheetChanges: [], metadataChanged: false, result: { accepted: false, changed: false, ranges: [], issues: [], reason } }
}

function assertCompleteMerges(sheet: SpreadsheetSheet, ranges: readonly CellRange[]): void {
  for (const merge of sheet.mergedRanges) {
    if (ranges.some(range => sheetRangesIntersect(merge, range)) && !ranges.some(range => sheetRangeContainsCell(range, merge.startRow, merge.startCol)
      && sheetRangeContainsCell(range, merge.endRow, merge.endCol))) throw new RangeError('粘贴范围会覆盖部分合并单元格，请选择完整区域。')
  }
}

/** Planning never mutates the authoritative workbook; source clearing, references and resources form one candidate. */
export function createSheetClipboardPastePlan(
  workbook: WorkbookModel, target: SheetClipboardLocation, incoming: SheetClipboardSnapshot, requestedMode: SheetPasteMode,
  context: SheetClipboardPasteContext,
): SheetClipboardPastePlan {
  assertSheetClipboardSnapshot(incoming)
  const mode = sheetClipboardPasteMode(requestedMode)
  if (!target || target.unitId !== workbook.id || typeof target.sheetId !== 'string') throw new TypeError('粘贴目标不属于当前工作簿。')
  assertSheetCommandRanges([target.range])
  const originalTarget = workbook.sheets.find(sheet => sheet.id === target.sheetId)
  if (!originalTarget) return failed('stale-target')
  if (incoming.copyType === 'cut' && incoming.source?.unitId !== workbook.id) return failed('cross-workbook-cut')
  let snapshot = incoming
  if (incoming.copyType === 'cut') {
    const location = incoming.source!
    const source = workbook.sheets.find(sheet => sheet.id === location.sheetId)
    if (!source || location.historyGeneration !== undefined && context.sourceHistoryGeneration !== undefined
      && location.historyGeneration !== context.sourceHistoryGeneration
      || source && sheetClipboardSourceFingerprint(source, location.range, workbook) !== location.fingerprint) return failed('stale-source')
    // A caller cannot forge cells beside a valid fingerprint and then delete the real source.
    snapshot = captureSheetClipboard(workbook, location, { id: incoming.id, copyType: 'cut', userId: context.userId,
      revision: location.revision, historyGeneration: location.historyGeneration })
    if (JSON.stringify(snapshot.source!.range) !== JSON.stringify(location.range)) return failed('stale-source')
  }
  const actualMode = snapshot.formulaOnly && mode === 'all' ? 'formulas' : mode
  if (snapshot.externalFormat === 'text' && (actualMode === 'formats' || actualMode === 'columnWidths')) {
    return { sheetChanges: [], metadataChanged: false, result: { accepted: false, changed: false, ranges: [], issues: [] } }
  }
  const geometry = createSheetClipboardTargetGeometry(workbook, originalTarget, target.range, snapshot)
  assertSheetClipboardArea(geometry.rows.length, geometry.columns.length)
  assertCompleteMerges(originalTarget, geometry.ranges)
  assertSheetRangesEditable(originalTarget, actualMode === 'columnWidths'
    ? [{ startRow: 0, endRow: Math.max(1, originalTarget.rowCount ?? 1_048_576) - 1, startCol: geometry.columns[0]!, endCol: geometry.columns.at(-1)! }]
    : geometry.ranges, context.userId)
  const move: SheetClipboardMove | undefined = snapshot.copyType === 'cut' && actualMode === 'all' ? {
    sourceSheetId: snapshot.source!.sheetId, sourceSheetName: snapshot.source!.sheetName,
    targetSheetId: originalTarget.id, targetSheetName: originalTarget.name,
    sourceRange: { ...snapshot.source!.range }, targetRange: { ...geometry.ranges[0]! },
  } : undefined
  if (move && move.sourceSheetId === move.targetSheetId && sheetRangesEqual(move.sourceRange, move.targetRange)) {
    return { sheetChanges: [], metadataChanged: false, result: { accepted: true, changed: false, ranges: geometry.ranges.map(range => ({ ...range })), issues: [] } }
  }
  if (move) {
    const source = workbook.sheets.find(sheet => sheet.id === move.sourceSheetId)!
    assertCompleteMerges(source, [move.sourceRange]); assertSheetRangesEditable(source, [move.sourceRange], context.userId)
  }
  const drafts = new Map<string, SpreadsheetSheet>()
  const draftFor = (sheet: SpreadsheetSheet): SpreadsheetSheet => {
    let draft = drafts.get(sheet.id)
    if (!draft) { draft = cloneModel(sheet); drafts.set(sheet.id, draft) }
    return draft
  }
  if (move) for (const sheet of workbook.sheets) {
    const updatedCells = Object.values(sheet.cells).map(cell => moveSheetClipboardCellReferences(cell, sheet.name, sheet.name, workbook, move))
    const changed = updatedCells.filter(cell => JSON.stringify(cell) !== JSON.stringify(sheet.cells[cell.id]))
    const resources = { ...sheet }
    moveSheetClipboardLinkedRuleReferences(resources, move)
    if (changed.length || JSON.stringify(resources.conditionalFormats) !== JSON.stringify(sheet.conditionalFormats)
      || JSON.stringify(resources.dataValidations) !== JSON.stringify(sheet.dataValidations)) {
      const draft = draftFor(sheet)
      for (const cell of changed) draft.cells[cell.id] = cell
      draft.conditionalFormats = resources.conditionalFormats; draft.dataValidations = resources.dataValidations
    }
  }
  const targetDraft = draftFor(originalTarget)
  const validationPositions = new Map<string, Map<string, SheetValidationPosition>>()
  const includePosition = (sheetId: string, row: number, column: number): void => {
    const positions = validationPositions.get(sheetId) ?? new Map<string, SheetValidationPosition>()
    positions.set(`${row}:${column}`, { row, column }); validationPositions.set(sheetId, positions)
  }
  if (move) {
    const sourceDraft = draftFor(workbook.sheets.find(sheet => sheet.id === move.sourceSheetId)!)
    for (const [key, cell] of Object.entries(sourceDraft.cells)) if (sheetRangeContainsCell(move.sourceRange, cell.row, cell.col)) {
      delete sourceDraft.cells[key]; includePosition(sourceDraft.id, cell.row, cell.col)
    }
    sourceDraft.mergedRanges = sourceDraft.mergedRanges.filter(merge => !sheetRangesIntersect(merge, move.sourceRange))
    trimSheetClipboardRangeResources(sourceDraft, [move.sourceRange], true, true)
  }
  if (actualMode === 'columnWidths') {
    for (let column = 0; column < geometry.columns.length; column++) {
      const destination = geometry.columns[column]!, width = snapshot.columnWidths[column % snapshot.columns.length]!
      if (sheetColumnWidth(targetDraft, destination) !== width) {
        targetDraft.columnWidthSpans = assignSheetAxisSpan(targetDraft.columnWidthSpans, destination, destination, width)
        delete targetDraft.columnWidths[destination]
      }
    }
  } else {
    for (let row = 0; row < geometry.rows.length; row++) for (let column = 0; column < geometry.columns.length; column++) {
      const sourceRow = row % snapshot.rows.length, sourceCol = column % snapshot.columns.length
      const position = { row: geometry.rows[row]!, col: geometry.columns[column]! }, key = `${position.row}:${position.col}`
      let sourceCell = snapshot.cells[`${sourceRow}:${sourceCol}`]
      if (sourceCell && move) sourceCell = moveSheetClipboardCellReferences(sourceCell, move.sourceSheetName, move.targetSheetName, workbook, move)
      const pasted = createSheetClipboardPastedCell(sourceCell, targetDraft.cells[key], position, actualMode, {
        rowDelta: position.row - snapshot.rows[sourceRow]!, columnDelta: position.col - snapshot.columns[sourceCol]!,
        externalFormat: snapshot.externalFormat, date1904: workbook.date1904, region: context.region,
        preserveFormulaReferences: snapshot.copyType === 'cut', contentOnly: !!geometry.contentOnlyMerge,
      })
      if (pasted) targetDraft.cells[key] = pasted
      else delete targetDraft.cells[key]
      if (actualMode !== 'formats') includePosition(targetDraft.id, position.row, position.col)
    }
    if (!geometry.contentOnlyMerge && (actualMode === 'all' || actualMode === 'formats' || actualMode === 'exceptBorders')) {
      targetDraft.mergedRanges = [...targetDraft.mergedRanges.filter(merge => !geometry.ranges.some(range => sheetRangesIntersect(merge, range))),
        ...geometry.merges.map(merge => ({ ...merge }))]
    }
    pasteSheetClipboardRangeResources(targetDraft, snapshot, geometry, actualMode, move)
  }
  if (targetDraft.columnCount !== undefined) targetDraft.columnCount = Math.max(targetDraft.columnCount, geometry.columns.at(-1)! + 1)
  if (actualMode !== 'columnWidths' && targetDraft.rowCount !== undefined) targetDraft.rowCount = Math.max(targetDraft.rowCount, geometry.rows.at(-1)! + 1)
  const metadata = move ? moveSheetClipboardDefinedNames(workbook.metadata, workbook, move) : workbook.metadata
  const metadataChanged = JSON.stringify(metadata) !== JSON.stringify(workbook.metadata)
  const candidate: WorkbookModel = { ...workbook, sheets: workbook.sheets.map(sheet => drafts.get(sheet.id) ?? sheet), metadata }
  const workbookResources = pasteSheetClipboardWorkbookResources(workbook, candidate, target, snapshot, geometry, actualMode, move)
  const changes = candidate.sheets.filter(sheet => JSON.stringify(sheet) !== JSON.stringify(workbook.sheets.find(original => original.id === sheet.id)))
  for (const after of changes) {
    const before = workbook.sheets.find(sheet => sheet.id === after.id)!
    const positions = new Set([...Object.keys(before.cells), ...Object.keys(after.cells)])
    for (const key of positions) if (JSON.stringify(before.cells[key]) !== JSON.stringify(after.cells[key])) {
      const cell = after.cells[key] ?? before.cells[key]!
      if (after.cells[key]?.formula !== before.cells[key]?.formula) includePosition(after.id, cell.row, cell.col)
    }
  }
  const changedRanges = assertSheetClipboardWorkbookChangesEditable(workbook, candidate, context.userId)
  const validations = [...validationPositions].map(([sheetId, positions]) => validateSheetDataValidationChanges(candidate, sheetId, [...positions.values()]))
  const issues = validations.flatMap(result => result.issues)
  if (validations.some(result => result.blocked)) return { ...failed('validation'), result: { accepted: false, changed: false,
    ranges: geometry.ranges.map(range => ({ ...range })), issues, reason: 'validation' } }
  return { sheetChanges: changes.map(sheet => cloneModel(sheet)), changedRanges, workbookMetadata: metadata ? cloneModel(metadata) : undefined, metadataChanged,
    workbookResources,
    result: { accepted: true, changed: changes.length > 0 || metadataChanged || workbookResources !== undefined, ranges: geometry.ranges.map(range => ({ ...range })), issues } }
}

export function applySheetClipboardPastePlan(draft: WorkbookModel, plan: SheetClipboardPastePlan): void {
  if (!plan.result.accepted) throw new Error('不能发布已拒绝的粘贴候选。')
  if (plan.sheetChanges.some(sheet => !draft.sheets.some(original => original.id === sheet.id))) throw new Error('粘贴目标工作表已移除。')
  const changes = new Map(plan.sheetChanges.map(sheet => [sheet.id, sheet]))
  draft.sheets = draft.sheets.map(sheet => changes.has(sheet.id) ? cloneModel(changes.get(sheet.id)!) : sheet)
  if (plan.metadataChanged) {
    if (plan.workbookMetadata === undefined) delete draft.metadata
    else draft.metadata = cloneModel(plan.workbookMetadata)
  }
  if (plan.workbookResources) for (const key of ['tables', 'pivotTables', 'charts', 'slicers'] as const) {
    const resources = plan.workbookResources[key]
    if (resources === undefined) delete draft[key]
    else Object.assign(draft, { [key]: cloneModel(resources) })
  }
}
