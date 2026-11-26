import type { OfficeStudio } from '../../api/office-studio.ts'
import type { CellRange, SparklineSpec, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { createCell } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { assertSparklineGroupConsistency, normalizeSparklineSpec, planSparklineMutation, type SparklineMutationRequest } from './sparkline-model.ts'
import { parseSparklineReference, sparklineCellAddress } from './sparkline-reference.ts'
import { resolveSparklineSource } from './sparkline-source.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'

export const SET_SPARKLINES_COMMAND = 'sheet.command.set-sparklines'
export interface SparklineMutationContext {
  host?: 'workbench'
  expectedRevision?: number
  historyGeneration?: number
  expectedUserId?: string
}
export type SparklineCommandParams = { unitId: string; sheetId: string } & SparklineMutationContext & (
  | { request: SparklineMutationRequest; patches?: never }
  | { patches: Record<string, SparklineSpec | null>; request?: never }
)

function targetRange(sheet: SpreadsheetSheet, address: string): CellRange {
  const reference = parseSparklineReference(address)
  if (!reference || reference.range.startRow !== reference.range.endRow || reference.range.startCol !== reference.range.endCol
    || reference.sheetName !== undefined && reference.sheetName !== sheet.id && reference.sheetName.toLocaleLowerCase() !== sheet.name.toLocaleLowerCase()) {
    throw new TypeError('迷你图目标必须是当前工作表中的一个单元格。')
  }
  return reference.range
}

/** Old patch callers and the full settings planner share canonical, defensive cell patches. */
export function normalizeSparklinePatches(sheet: SpreadsheetSheet, patches: Readonly<Record<string, SparklineSpec | null>>): Record<string, SparklineSpec | null> {
  if (!patches || typeof patches !== 'object' || Array.isArray(patches)) throw new TypeError('迷你图修改数据无效。')
  const normalized: Record<string, SparklineSpec | null> = {}
  for (const [input, spec] of Object.entries(patches)) {
    const range = targetRange(sheet, input), address = sparklineCellAddress(range.startRow, range.startCol)
    if (Object.hasOwn(normalized, address)) throw new TypeError('迷你图修改包含重复目标单元格。')
    if (spec !== null && (!spec || typeof spec !== 'object' || Array.isArray(spec))) throw new TypeError('迷你图设置数据无效。')
    normalized[address] = spec === null ? null : normalizeSparklineSpec(spec, address)
  }
  assertSparklineGroupConsistency(sheet, normalized)
  return normalized
}

function assertViewRange(sheet: SpreadsheetSheet, ranges: readonly CellRange[], userId: string): void {
  if (!canViewWorksheetProtection(sheet, userId) || !canViewProtectedRanges(sheet, ranges, userId)) throw new Error('没有查看迷你图或其数据来源范围的权限。')
}

/** Configuration reads expose complete groups only when all member cells are readable. */
export function assertSparklineReadPermission(sheet: SpreadsheetSheet, addresses: readonly string[], userId: string): void {
  assertViewRange(sheet, addresses.map(address => targetRange(sheet, address)), userId)
}

function assertSourceRead(workbook: WorkbookModel, host: SpreadsheetSheet, spec: SparklineSpec, userId: string, retained: boolean): void {
  const source = resolveSparklineSource(spec, host, workbook)
  // A broken retained reference may be repaired or deleted. New authoring must name a real source.
  if (source.status !== 'resolved') {
    if (retained) return
    throw new Error('迷你图数据来源工作表或范围已不存在。')
  }
  assertViewRange(source.sheet, [source.range], userId)
}

export function assertSparklineMutationPermission(workbook: WorkbookModel, sheet: SpreadsheetSheet,
  patches: Readonly<Record<string, SparklineSpec | null>>, userId: string): void {
  const ranges = Object.keys(patches).map(address => targetRange(sheet, address))
  assertViewRange(sheet, ranges, userId)
  if (!canEditWorksheetProtection(sheet, userId) || !canEditProtectedRanges(sheet, ranges, userId)) throw new Error('没有编辑迷你图完整目标范围的权限。')
  for (const [address, spec] of Object.entries(patches)) {
    if (!spec) continue
    const range = targetRange(sheet, address), previous = sheet.cells[`${range.startRow}:${range.startCol}`]?.sparkline
    if (previous) assertSourceRead(workbook, sheet, previous, userId, true)
    const retained = spec.sourceReference !== undefined || spec.officeSource !== undefined
      || previous !== undefined && previous.sourceRange === spec.sourceRange && equivalentModelData(previous.sourceReference, spec.sourceReference)
    assertSourceRead(workbook, sheet, spec, userId, retained)
  }
}

export function applySparklinePatches(sheet: SpreadsheetSheet, patches: Readonly<Record<string, SparklineSpec | null>>): void {
  for (const [address, spec] of Object.entries(patches)) {
    const range = targetRange(sheet, address), key = `${range.startRow}:${range.startCol}`
    const cell = sheet.cells[key] ?? (spec ? createCell(range.startRow, range.startCol) : undefined)
    if (!cell) continue
    if (spec) {
      cell.sparkline = cloneModel(spec); sheet.cells[key] = cell
      sheet.rowCount = Math.max(sheet.rowCount, range.startRow + 1)
      sheet.columnCount = Math.max(sheet.columnCount, range.startCol + 1)
    } else delete cell.sparkline
  }
}

export function registerSparklineCommands(studio: OfficeStudio): void {
  studio.registerCommand<SparklineCommandParams, boolean>(SET_SPARKLINES_COMMAND, input => {
    if (!input || typeof input !== 'object' || typeof input.unitId !== 'string' || !input.unitId
      || typeof input.sheetId !== 'string' || !input.sheetId || (input.request === undefined) === (input.patches === undefined)) {
      throw new TypeError('迷你图命令须指定工作簿、工作表及一种修改数据。')
    }
    if (input.host !== undefined && input.host !== 'workbench'
      || input.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0)
      || input.historyGeneration !== undefined && (!Number.isSafeInteger(input.historyGeneration) || input.historyGeneration <= 0)
      || input.expectedUserId !== undefined && (typeof input.expectedUserId !== 'string' || !input.expectedUserId.trim())) throw new TypeError('迷你图交互上下文无效。')
    const workbook = studio.getWorkbookModel(input.unitId), sheet = workbook?.sheets.find(sheet => sheet.id === input.sheetId)
    if (!workbook || !sheet) throw new Error('迷你图所属工作簿或工作表已关闭。')
    const revision = studio.getWorkbookRevision(input.unitId), generation = studio.editor.history.captureSavepoint(input.unitId).generation, userId = studio.currentUserId
    const assertCurrent = (): void => {
      if (studio.isDisposed || studio.getWorkbookModel(input.unitId) !== workbook || studio.getWorkbookRevision(input.unitId) !== revision
        || studio.currentUserId !== userId || studio.editor.history.captureSavepoint(input.unitId).generation !== generation
        || input.expectedRevision !== undefined && input.expectedRevision !== revision
        || input.historyGeneration !== undefined && input.historyGeneration !== generation
        || input.expectedUserId !== undefined && input.expectedUserId !== userId) throw new Error('迷你图目标内容、用户或历史已改变。')
      if (input.host === 'workbench' && (studio.activeWorkbookId !== input.unitId || studio.editor.state.activeKind !== 'workbook'
        || workbook.activeSheetId !== input.sheetId)) throw new Error('迷你图工作台宿主已改变。')
    }
    assertCurrent()
    const patches = normalizeSparklinePatches(sheet, input.request === undefined ? input.patches!
      : planSparklineMutation(sheet, input.request, { sheets: workbook.sheets, workbook }).patches)
    assertSparklineMutationPermission(workbook, sheet, patches, userId)
    return studio.commitWorkbook(input.unitId, '编辑迷你图', candidate => {
      assertCurrent()
      const target = candidate.sheets.find(sheet => sheet.id === input.sheetId)!
      assertSparklineMutationPermission(candidate, target, patches, userId)
      applySparklinePatches(target, patches)
      assertSparklineGroupConsistency(target)
    }, { expectedRevision: revision, expectedHistoryGeneration: generation, prepareView: candidate => {
      assertCurrent()
      const target = candidate.sheets.find(sheet => sheet.id === input.sheetId)!
      assertSparklineMutationPermission(candidate, target, patches, userId)
      assertSparklineGroupConsistency(target)
    } })
  })
}
