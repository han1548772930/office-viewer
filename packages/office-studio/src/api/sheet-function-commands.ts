import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import type { CellRange, SpreadsheetSheet } from '../domain/model.ts'
import { DEFAULT_SHEET_COLUMN_COUNT, DEFAULT_SHEET_ROW_COUNT } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { applyCellValue } from './model-utils.ts'
import { bindWorkbookCommandGuard, workbookCommandGuard } from './workbook-command-guards.ts'
import { assertSheetRangesEditable, sheetRangeContainsCell, type SheetCommandRange } from '../features/sheet-workbench/sheet-range-command.ts'
import { assertSheetRangeChangedValuesValid } from '../features/sheet-workbench/sheet-range-editor-validation.ts'
import { expandSheetFunctionInsertions, planSheetFunctionInsertion, sheetInsertionFunctionName,
  type SheetFunctionInsertion, type SheetFunctionSelection } from '../features/formula/sheet-function-insertion.ts'

export const SHEET_INSERT_FUNCTION_OPERATION = 'formula-ui.operation.insert-function'
export const INSERT_SHEET_FUNCTION_COMMAND = 'formula.command.insert-function'
export interface SheetFunctionTargetInput {
  unitId?: string; sheetId?: string; subUnitId?: string
  expectedRevision?: number; expectedHistoryGeneration?: number; expectedUserId?: string
}
export interface SheetFunctionOperationInput extends SheetFunctionTargetInput { value: string }
export interface SheetFunctionInsertCommandInput extends SheetFunctionTargetInput {
  list: readonly SheetFunctionCommandInsertion[]
  listOfRangeHasNumber?: readonly SheetFunctionCommandInsertion[]
}
export interface SheetFunctionCommandInsertion { range: SheetCommandRange; primary: { row: number; column: number }; formula: string }
export interface SheetFunctionEditTarget extends SheetFunctionInsertion { unitId: string; sheetId: string }
export interface SheetFunctionEditor { canInsert(): boolean; open(target: SheetFunctionEditTarget): boolean; release(): void }
interface EditorOwner { editor: SheetFunctionEditor; disposed: boolean }
const editors = new WeakMap<OfficeStudio, Map<string, EditorOwner>>()

function record(raw: unknown): Record<string, unknown> {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new TypeError('函数插入参数必须为对象。')
  return raw as Record<string, unknown>
}
function targetInput(raw: Record<string, unknown>): SheetFunctionTargetInput {
  const input: SheetFunctionTargetInput = {}
  for (const key of ['unitId', 'sheetId', 'subUnitId', 'expectedUserId'] as const) {
    const value = raw[key]
    if (value !== undefined && (typeof value !== 'string' || !value.trim())) throw new TypeError('函数插入的目标 ID 无效。')
    if (typeof value === 'string') input[key] = value
  }
  if (input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('函数插入的工作表 ID 不一致。')
  for (const key of ['expectedRevision', 'expectedHistoryGeneration'] as const) {
    const value = raw[key]
    if (value !== undefined && (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0)) throw new TypeError('函数插入的版本无效。')
    if (typeof value === 'number') input[key] = value
  }
  return input
}
function coordinate(value: unknown, maximum: number): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0 || value >= maximum) throw new RangeError('函数插入坐标超出 Office 范围。')
  return value
}
function insertion(raw: unknown): SheetFunctionInsertion {
  const input = record(raw), source = record(input.range), primary = record(input.primary)
  if (source.startCol !== undefined && source.startColumn !== undefined && source.startCol !== source.startColumn
    || source.endCol !== undefined && source.endColumn !== undefined && source.endCol !== source.endColumn) throw new TypeError('函数插入范围的列坐标不一致。')
  const range: CellRange = { startRow: coordinate(source.startRow, 1_048_576), endRow: coordinate(source.endRow, 1_048_576),
    startCol: coordinate(source.startCol ?? source.startColumn, 16_384), endCol: coordinate(source.endCol ?? source.endColumn, 16_384) }
  if (range.startRow > range.endRow || range.startCol > range.endCol) throw new RangeError('函数插入范围必须按起止顺序排列。')
  if (typeof input.formula !== 'string' || !input.formula.startsWith('=') || input.formula.length > 32_768) throw new TypeError('插入的公式无效。')
  return { range, primary: { row: coordinate(primary.row, 1_048_576), column: coordinate(primary.column, 16_384) }, formula: input.formula }
}
function context(api: OfficeStudio, input: SheetFunctionTargetInput) {
  const unitId = input.unitId ?? api.activeWorkbookId, workbook = api.getWorkbookModel(unitId)
  if (!workbook) throw new Error('函数插入的工作簿不存在。')
  const sheetId = input.sheetId ?? input.subUnitId ?? workbook.activeSheetId, sheet = workbook.sheets.find(sheet => sheet.id === sheetId)
  if (!sheet) throw new Error('函数插入的工作表不存在。')
  api.assertCanViewWorksheet(unitId, sheetId)
  const historyGeneration = api.editor.history.captureSavepoint(unitId).generation, revision = api.getWorkbookRevision(unitId), userId = api.currentUserId
  if (input.expectedRevision !== undefined && input.expectedRevision !== revision
    || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== historyGeneration
    || input.expectedUserId !== undefined && input.expectedUserId !== userId) throw new Error('函数插入的内容版本或用户已变化。')
  const assertCurrent = (): void => {
    if (api.isDisposed || api.getWorkbookModel(unitId) !== workbook || api.getWorkbookRevision(unitId) !== revision
      || api.editor.history.captureSavepoint(unitId).generation !== historyGeneration || api.currentUserId !== userId) throw new Error('函数插入的目标内容或历史已变化。')
    api.assertCanViewWorksheet(unitId, sheetId)
  }
  return { unitId, sheetId, sheet, workbook, revision, historyGeneration, userId, assertCurrent }
}
function assertBounds(sheet: SpreadsheetSheet, range: CellRange): void {
  if (range.endRow >= (sheet.rowCount ?? DEFAULT_SHEET_ROW_COUNT) || range.endCol >= (sheet.columnCount ?? DEFAULT_SHEET_COLUMN_COUNT)) throw new RangeError('公式结果超出工作表边界。')
}
function selectionSignature(api: OfficeStudio, unitId: string, sheetId: string): string {
  const selected = api.getSelectionState(unitId, sheetId)
  return JSON.stringify(selected && { ranges: selected.ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow,
    startColumn: range.startColumn, endColumn: range.endColumn, rangeType: range.rangeType ?? 0 })),
    activeRangeIndex: selected.activeRangeIndex, primary: selected.primary })
}

export function registerSheetFunctionEditor(api: OfficeStudio, unitId: string, editor: SheetFunctionEditor) {
  if (!api.getWorkbookModel(unitId)) throw new Error('函数编辑器的目标工作簿不存在。')
  let units = editors.get(api)
  if (!units) { units = new Map(); editors.set(api, units) }
  const owner: EditorOwner = { editor, disposed: false }, previous = units.get(unitId)
  units.set(unitId, owner)
  if (previous) { previous.disposed = true; previous.editor.release() }
  const current = () => !api.isDisposed && !owner.disposed && editors.get(api)?.get(unitId) === owner
  return {
    execute(value: string): boolean {
      if (!current() || !editor.canInsert()) return false
      const target = context(api, { unitId }), selected = selectionSignature(api, unitId, target.sheetId)
      const assertCurrent = () => {
        target.assertCurrent()
        if (!current() || !editor.canInsert() || selected !== selectionSignature(api, unitId, target.sheetId)) throw new Error('函数插入的编辑器或选区已变化。')
      }
      return api.executeCommandSync(SHEET_INSERT_FUNCTION_OPERATION, bindWorkbookCommandGuard({ unitId, value,
        expectedRevision: target.revision, expectedHistoryGeneration: target.historyGeneration, expectedUserId: target.userId }, assertCurrent))
    },
    dispose() {
      if (owner.disposed) return
      owner.disposed = true
      if (editors.get(api)?.get(unitId) === owner) editors.get(api)!.delete(unitId)
      editor.release()
    },
  }
}
export function disposeSheetFunctionEditors(api: OfficeStudio, unitId?: string): void {
  const units = editors.get(api)
  if (!units) return
  for (const [id, owner] of [...units]) if (unitId === undefined || id === unitId) {
    units.delete(id); owner.disposed = true; owner.editor.release()
  }
  if (!units.size) editors.delete(api)
}

export function registerSheetFunctionCommands(api: OfficeStudio): Disposable {
  const command = api.registerCommand<unknown, boolean>(INSERT_SHEET_FUNCTION_COMMAND, raw => {
    const source = record(raw), input = targetInput(source), guard = workbookCommandGuard(raw)
    guard?.()
    if (!Array.isArray(source.list) || source.list.length > 4096 || source.listOfRangeHasNumber !== undefined
      && (!Array.isArray(source.listOfRangeHasNumber) || source.listOfRangeHasNumber.length > 4096)) throw new TypeError('函数插入列表无效或超过任务预算。')
    const list = source.list.map(insertion), numeric = source.listOfRangeHasNumber === undefined ? [] : source.listOfRangeHasNumber.map(insertion)
    const target = context(api, input), writes = expandSheetFunctionInsertions(list, numeric)
    if (!writes.length) return false
    const ranges = writes.map(write => ({ startRow: write.row, endRow: write.row, startCol: write.column, endCol: write.column }))
    for (const item of [...list, ...numeric]) assertBounds(target.sheet, item.range)
    for (const range of ranges) assertBounds(target.sheet, range)
    const assertCurrent = () => { guard?.(); target.assertCurrent(); assertSheetRangesEditable(target.sheet, ranges, api.currentUserId) }
    assertCurrent()
    const changed = api.commitWorkbook(target.unitId, '插入函数', draft => {
      const sheet = draft.sheets.find(sheet => sheet.id === target.sheetId)!
      for (const write of writes) {
        const key = `${write.row}:${write.column}`
        const cell = applyCellValue(sheet.cells[key], write.row, write.column, { f: write.formula }, draft, sheet)
        delete cell.richText
        sheet.cells[key] = cell
      }
      assertSheetRangeChangedValuesValid(target.sheet, draft, target.sheetId, { userId: api.currentUserId, region: api.regionalSettings.region.value })
    }, { expectedRevision: target.revision, expectedHistoryGeneration: target.historyGeneration, assertCurrent, preserveStyleInheritance: true })
    if (changed) api.publishObserverEvent('SheetValueChanged', () => ({ workbookId: target.unitId, sheetId: target.sheetId,
      ranges: ranges.map(range => ({ unitId: target.unitId, sheetId: target.sheetId, startRow: range.startRow, endRow: range.endRow,
        startColumn: range.startCol, endColumn: range.endCol })) }))
    return changed
  })
  const operation = api.registerCommand<unknown, boolean>(SHEET_INSERT_FUNCTION_OPERATION, raw => {
    const source = record(raw), input = targetInput(source), name = sheetInsertionFunctionName(source.value), guard = workbookCommandGuard(raw)
    guard?.()
    const target = context(api, input)
    if (api.activeWorkbookId !== target.unitId || api.editor.state.activeKind !== 'workbook' || target.workbook.activeSheetId !== target.sheetId) return false
    const selected = api.getSelectionState(target.unitId, target.sheetId)
    if (!selected?.ranges.length) return false
    const selections: SheetFunctionSelection[] = selected.ranges.map((range, index) => ({ range: {
      startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn,
    }, ...(index === selected.activeRangeIndex && selected.primary ? { primary: selected.primary } : {}) }))
    const plan = planSheetFunctionInsertion(target.sheet, selections, name), owner = editors.get(api)?.get(target.unitId)
    if (owner && (owner.disposed || !owner.editor.canInsert())) return false
    if (plan.edit && !owner) return false
    let selectedSignature = selectionSignature(api, target.unitId, target.sheetId), opened = false
    const assertCurrent = () => {
      target.assertCurrent()
      if (api.activeWorkbookId !== target.unitId || target.workbook.activeSheetId !== target.sheetId || api.editor.state.activeKind !== 'workbook'
        || selectedSignature !== selectionSignature(api, target.unitId, target.sheetId)
        || owner && (owner.disposed || editors.get(api)?.get(target.unitId) !== owner)) throw new Error('函数插入的工作台或选区已变化。')
    }
    if (plan.edit) {
      const primary = plan.edit.primary, merged = target.sheet.mergedRanges.find(range => sheetRangeContainsCell(range, primary.row, primary.column))
      assertSheetRangesEditable(target.sheet, [merged ?? { startRow: primary.row, endRow: primary.row, startCol: primary.column, endCol: primary.column }], api.currentUserId)
      assertCurrent()
      opened = owner!.editor.open({ ...cloneModel(plan.edit), unitId: target.unitId, sheetId: target.sheetId })
      if (!opened) return false
      selectedSignature = selectionSignature(api, target.unitId, target.sheetId)
      target.assertCurrent()
    }
    if (!plan.list.length && !plan.listOfRangeHasNumber.length) return opened
    assertCurrent()
    return api.executeCommandSync(INSERT_SHEET_FUNCTION_COMMAND, bindWorkbookCommandGuard({ unitId: target.unitId, sheetId: target.sheetId,
      list: plan.list, listOfRangeHasNumber: plan.listOfRangeHasNumber, expectedRevision: target.revision,
      expectedHistoryGeneration: target.historyGeneration, expectedUserId: target.userId }, assertCurrent)) || opened
  })
  return { dispose() { command.dispose(); operation.dispose(); disposeSheetFunctionEditors(api) } }
}
