import { createCell, type CellRange, type SpreadsheetCell, type SpreadsheetSheet, type WorkbookModel } from '../../domain/model.ts'
import { sheetColumnWidth } from '../../domain/sheet-axis.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { assertSheetCommandRanges, assertSheetRangesEditable, sheetRangeContainsCell, sheetRangesIntersect } from './sheet-range-command.ts'
import { assertSheetClipboardWorkbookResources, captureSheetClipboardWorkbookResources } from './sheet-clipboard-workbook-resources.ts'
import type { SheetClipboardLocation, SheetClipboardSnapshot } from './sheet-clipboard-types.ts'
import { resolveSheetCellStyle } from './sheet-style-inheritance.ts'
import { assertSheetStyleRegion, captureSheetStyleRegion, sheetStyleRegionCellStyle } from './sheet-style-region.ts'
import { assertSheetClipboardImages, captureSheetClipboardImages } from '../sheet-images/sheet-image-clipboard.ts'
import { normalizeSheetCellImage } from '../sheet-images/sheet-image-model.ts'

// These are synchronous safety budgets, not measured performance thresholds. Larger transfers need the Worker path.
export const SHEET_CLIPBOARD_SYNC_CELL_BUDGET = 1_000_000
export const SHEET_CLIPBOARD_SERIALIZED_BYTE_BUDGET = 64 * 1024 * 1024

export function assertSheetClipboardArea(rows: number, columns: number): void {
  if (!Number.isSafeInteger(rows) || rows < 1 || rows > 1_048_576 || !Number.isSafeInteger(columns) || columns < 1 || columns > 16_384
    || rows * columns > SHEET_CLIPBOARD_SYNC_CELL_BUDGET) {
    throw new RangeError('剪贴板范围超出当前同步处理预算。')
  }
}
export function assertSheetClipboardBytes(text: string): void {
  if (text.length > SHEET_CLIPBOARD_SERIALIZED_BYTE_BUDGET
    || new TextEncoder().encode(text).byteLength > SHEET_CLIPBOARD_SERIALIZED_BYTE_BUDGET) {
    throw new RangeError('剪贴板数据超出当前内存预算。')
  }
}

export function expandSheetClipboardRange(sheet: Pick<SpreadsheetSheet, 'mergedRanges'>, range: CellRange): CellRange {
  assertSheetCommandRanges([range])
  let expanded = { ...range }, changed = true
  while (changed) {
    changed = false
    for (const merge of sheet.mergedRanges) {
      if (!sheetRangesIntersect(merge, expanded)) continue
      const next = { startRow: Math.min(expanded.startRow, merge.startRow), endRow: Math.max(expanded.endRow, merge.endRow),
        startCol: Math.min(expanded.startCol, merge.startCol), endCol: Math.max(expanded.endCol, merge.endCol) }
      if (JSON.stringify(next) !== JSON.stringify(expanded)) { expanded = next; changed = true }
    }
  }
  assertSheetCommandRanges([expanded])
  return expanded
}

/** Formula cache refreshes do not make a pending cut stale; original content, runs, styles and anchored resources do. */
export function sheetClipboardSourceFingerprint(sheet: SpreadsheetSheet, range: CellRange, workbook?: WorkbookModel): string {
  const cells = Object.values(sheet.cells).filter(cell => sheetRangeContainsCell(range, cell.row, cell.col))
    .sort((first, second) => first.row - second.row || first.col - second.col)
    .map(cell => {
      const { display: _display, value, cachedDataType, ...source } = cell
      const metadata = source.metadata ? { ...source.metadata } : undefined
      if (metadata) delete metadata.formulaError
      return { ...source, ...(cell.formula ? {} : { value, cachedDataType }), ...(metadata ? { metadata } : {}) }
    })
  const resourceCapture = workbook ? captureSheetClipboardWorkbookResources(workbook, { unitId: workbook.id!, sheetId: sheet.id, range },
    Array.from({ length: range.endRow - range.startRow + 1 }, (_, offset) => range.startRow + offset)) : undefined
  const fingerprint = JSON.stringify({ cells, styleRegion: captureSheetStyleRegion(sheet, { start: range.startRow, end: range.endRow },
    { start: range.startCol, end: range.endCol }, workbook), merges: sheet.mergedRanges.filter(merge => sheetRangesIntersect(merge, range)),
    conditionalFormats: (sheet.conditionalFormats ?? []).filter(rule => rule.ranges.some(item => sheetRangesIntersect(item, range))),
    dataValidations: (sheet.dataValidations ?? []).filter(rule => rule.ranges.some(item => sheetRangesIntersect(item, range))),
    rangeThemeRules: (sheet.rangeThemeRules ?? []).filter(rule => sheetRangesIntersect(rule.range, range)),
    ...(resourceCapture ? { workbookResources: resourceCapture.resources, sourceOrigin: resourceCapture.origin,
      images: captureSheetClipboardImages(workbook!, sheet, range, Array.from({ length: range.endRow - range.startRow + 1 }, (_, offset) => range.startRow + offset)) } : {}) })
  assertSheetClipboardBytes(fingerprint)
  return fingerprint
}

export interface CaptureSheetClipboardOptions {
  id?: string
  copyType?: 'copy' | 'cut'
  formulaOnly?: boolean
  revision?: number
  historyGeneration?: number
  userId?: string
}
export function captureSheetClipboard(workbook: WorkbookModel, location: SheetClipboardLocation, options: CaptureSheetClipboardOptions = {}): SheetClipboardSnapshot {
  if (workbook.id !== location.unitId) throw new Error('复制目标不属于当前工作簿。')
  const sheet = workbook.sheets.find(item => item.id === location.sheetId)
  if (!sheet) throw new Error('复制的工作表已移除。')
  const range = expandSheetClipboardRange(sheet, location.range), userId = options.userId ?? 'local-user'
  if (!canViewWorksheetProtection(sheet, userId) || !canViewProtectedRanges(sheet, [range], userId)) throw new Error('当前用户没有复制此区域的权限。')
  if (options.copyType === 'cut') assertSheetRangesEditable(sheet, [range], userId)
  assertSheetClipboardArea(range.endRow - range.startRow + 1, range.endCol - range.startCol + 1)
  const filtered = new Set([...(sheet.autoFilter?.filteredRows ?? []), ...(workbook.tables ?? [])
    .filter(table => table.sheetId === sheet.id).flatMap(table => table.autoFilter?.filteredRows ?? [])])
  const copiedMerges = sheet.mergedRanges.filter(merge => sheetRangesIntersect(merge, range))
  const rows: number[] = []
  for (let row = range.startRow; row <= range.endRow; row++) {
    // A move preserves the physical rectangle. A merge crossing a filtered row must retain its complete geometry.
    if (options.copyType !== 'cut' && filtered.has(row) && !copiedMerges.some(merge => row >= merge.startRow && row <= merge.endRow)) continue
    rows.push(row)
  }
  if (!rows.length) throw new RangeError('所选区域没有可复制的可见行。')
  const columns = Array.from({ length: range.endCol - range.startCol + 1 }, (_, offset) => range.startCol + offset)
  const rowOffsets = new Map(rows.map((row, offset) => [row, offset])), cells: Record<string, SpreadsheetCell> = Object.create(null)
  for (const cell of Object.values(sheet.cells)) {
    const row = rowOffsets.get(cell.row), col = cell.col - range.startCol
    if (row === undefined || col < 0 || col >= columns.length || options.formulaOnly && !cell.formula) continue
    const id = `${row}:${col}`
    cells[id] = { ...cloneModel(cell), style: resolveSheetCellStyle(sheet, cell.row, cell.col, cell, workbook), id, row, col }
  }
  const resources = options.formulaOnly ? undefined : captureSheetClipboardWorkbookResources(workbook, { ...location, range }, rows)
  const snapshot: SheetClipboardSnapshot = {
    version: 1, id: options.id ?? crypto.randomUUID(), copyType: options.copyType ?? 'copy', formulaOnly: options.formulaOnly ?? false,
    rows, columns, cells,
    styleRegion: options.formulaOnly ? undefined : captureSheetStyleRegion(sheet, rows, columns, workbook),
    merges: options.formulaOnly ? [] : copiedMerges.map(merge => ({ startRow: rowOffsets.get(merge.startRow)!, endRow: rowOffsets.get(merge.endRow)!,
      startCol: merge.startCol - range.startCol, endCol: merge.endCol - range.startCol })),
    columnWidths: columns.map(column => sheetColumnWidth(sheet, column)),
    conditionalFormats: options.formulaOnly ? undefined : cloneModel((sheet.conditionalFormats ?? []).filter(rule => rule.ranges.some(item => sheetRangesIntersect(item, range)))),
    dataValidations: options.formulaOnly ? undefined : cloneModel((sheet.dataValidations ?? []).filter(rule => rule.ranges.some(item => sheetRangesIntersect(item, range)))),
    ...(resources ? { workbookResources: resources.resources, sourceOrigin: resources.origin, sourceBounds: resources.bounds } : {}),
    ...(resources ? { images: captureSheetClipboardImages(workbook, sheet, range, rows) } : {}),
    source: { ...location, range, sheetName: sheet.name, revision: options.revision, historyGeneration: options.historyGeneration,
      fingerprint: sheetClipboardSourceFingerprint(sheet, range, workbook) },
  }
  assertSheetClipboardSnapshot(snapshot)
  return snapshot
}

/** Formatting-only blank cells are projected without changing the sparse snapshot. */
export function sheetClipboardCellAt(snapshot: SheetClipboardSnapshot, row: number, col: number): SpreadsheetCell | undefined {
  const stored = snapshot.cells[`${row}:${col}`]
  if (stored || !snapshot.styleRegion) return stored
  const cell = createCell(row, col, null)
  cell.style = sheetStyleRegionCellStyle(snapshot.styleRegion, row, col)
  cell.authoredStyle = cloneModel(cell.style)
  return cell
}

/** API data is checked before its sparse coordinates, styles or formulas can reach a model mutation. */
export function assertSheetClipboardSnapshot(value: SheetClipboardSnapshot): void {
  if (!value || value.version !== 1 || typeof value.id !== 'string' || !value.id.trim()
    || !['copy', 'cut'].includes(value.copyType) || typeof value.formulaOnly !== 'boolean'
    || !Array.isArray(value.rows) || !Array.isArray(value.columns) || !Array.isArray(value.columnWidths)
    || !value.cells || typeof value.cells !== 'object' || Array.isArray(value.cells) || !Array.isArray(value.merges)) {
    throw new TypeError('剪贴板快照无效。')
  }
  assertSheetClipboardArea(value.rows.length, value.columns.length)
  if (value.styleRegion !== undefined) {
    if (value.formulaOnly || value.externalFormat) throw new TypeError('继承样式采样只适用于内部普通剪贴板。')
    assertSheetStyleRegion(value.styleRegion, value.rows.length, value.columns.length)
  }
  for (const [indices, limit] of [[value.rows, 1_048_576], [value.columns, 16_384]] as const) {
    if (indices.some((index, offset) => !Number.isSafeInteger(index) || index < 0 || index >= limit || offset > 0 && index <= indices[offset - 1]!)) {
      throw new RangeError('剪贴板源坐标无效。')
    }
  }
  if (value.columnWidths.length !== value.columns.length || value.columnWidths.some(width => !Number.isFinite(width) || width < 0 || width > 2000)) {
    throw new RangeError('剪贴板列宽无效。')
  }
  for (const [key, cell] of Object.entries(value.cells)) {
    if (!cell || !Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.col) || cell.row < 0 || cell.col < 0
      || cell.row >= value.rows.length || cell.col >= value.columns.length || key !== `${cell.row}:${cell.col}`
      || cell.id !== key || cell.formula !== null && typeof cell.formula !== 'string' || typeof cell.display !== 'string'
      || !cell.style || !cell.style.font || !Number.isFinite(cell.style.font.size) || typeof cell.style.font.family !== 'string'
      || typeof cell.style.numberFormat !== 'string' || ![null, 'string', 'number', 'boolean'].includes(cell.value === null ? null : typeof cell.value)
      || typeof cell.value === 'number' && !Number.isFinite(cell.value)) throw new TypeError('剪贴板单元格无效。')
    if (cell.image) normalizeSheetCellImage(cell.image)
  }
  assertSheetCommandRanges(value.merges)
  if (value.merges.some(merge => merge.endRow >= value.rows.length || merge.endCol >= value.columns.length)) throw new RangeError('剪贴板合并范围无效。')
  if (value.source) {
    if (![value.source.unitId, value.source.sheetId, value.source.sheetName, value.source.fingerprint].every(item => typeof item === 'string')
      || !value.source.unitId || !value.source.sheetId) throw new TypeError('剪贴板源标识无效。')
    assertSheetCommandRanges([value.source.range])
    if (value.rows.some(row => row < value.source!.range.startRow || row > value.source!.range.endRow)
      || value.columns.some(column => column < value.source!.range.startCol || column > value.source!.range.endCol)) throw new RangeError('剪贴板源范围无效。')
  }
  if (value.copyType === 'cut' && (!value.source || value.formulaOnly)) throw new TypeError('剪切快照缺少有效源区域。')
  for (const rules of [value.conditionalFormats, value.dataValidations]) if (rules !== undefined) {
    if (!Array.isArray(rules)) throw new TypeError('剪贴板范围资源无效。')
    for (const rule of rules) {
      if (!rule || typeof rule.id !== 'string' || !Array.isArray(rule.ranges)) throw new TypeError('剪贴板范围资源无效。')
      assertSheetCommandRanges(rule.ranges)
    }
  }
  if (value.workbookResources !== undefined) {
    if (!value.source || value.formulaOnly || value.externalFormat) throw new TypeError('工作簿对象需要内部剪贴板源。')
    assertSheetClipboardWorkbookResources(value.workbookResources, { source: value.source, rows: value.rows, columns: value.columns,
      sourceOrigin: value.sourceOrigin, sourceBounds: value.sourceBounds })
  }
  if (value.sourceOrigin !== undefined && (!value.sourceOrigin || ![value.sourceOrigin.left, value.sourceOrigin.top].every(Number.isFinite))) {
    throw new TypeError('剪贴板源位置无效。')
  }
  if (value.sourceBounds !== undefined && (!value.sourceBounds || ![value.sourceBounds.left, value.sourceBounds.top, value.sourceBounds.width, value.sourceBounds.height].every(Number.isFinite)
    || value.sourceBounds.left < 0 || value.sourceBounds.top < 0 || value.sourceBounds.width < 0 || value.sourceBounds.height < 0)) {
    throw new TypeError('剪贴板源范围几何无效。')
  }
  assertSheetClipboardBytes(JSON.stringify(value))
  assertSheetClipboardImages(value)
}
