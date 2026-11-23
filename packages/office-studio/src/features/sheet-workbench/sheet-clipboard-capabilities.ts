import type { CellRange, SheetAutoFilter, SheetTableModel, SpreadsheetSheet } from '../../domain/model.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { resolveSheetClipboardCommand } from './sheet-clipboard-command.ts'
import { expandSheetClipboardRange, SHEET_CLIPBOARD_SYNC_CELL_BUDGET } from './sheet-clipboard-data.ts'
import { sheetCommandRangesWithoutFilteredRows, sheetRangeContainsCell, sheetRangesIntersect } from './sheet-range-command.ts'
import type { SheetClipboardSnapshot } from './sheet-clipboard-types.ts'
import type { SheetPasteMode } from './sheet-paste.ts'

/** Cells are deliberately absent: entry prediction does not normalize or clone the workbook. */
export interface SheetClipboardCapabilitySheet extends Pick<SpreadsheetSheet, 'id' | 'worksheetProtection' | 'rangeProtectionRules' | 'rowCount' | 'columnCount'> {
  visible?: boolean
  mergedRanges?: CellRange[]
  filteredRows?: readonly number[]
  autoFilter?: Pick<SheetAutoFilter, 'filteredRows'>
}
export interface SheetClipboardCapabilityWorkbook {
  id?: string
  sheets: readonly SheetClipboardCapabilitySheet[]
  tables?: readonly Pick<SheetTableModel, 'sheetId' | 'autoFilter'>[]
}
export type SheetClipboardPermissionPoint = 'copy' | 'view' | 'edit' | 'setCellValue' | 'setCellStyle' | 'setColumnWidth'
export interface SheetClipboardPermissionRequest {
  point: SheetClipboardPermissionPoint
  phase: 'source' | 'target'
  workbook: SheetClipboardCapabilityWorkbook
  sheet: SheetClipboardCapabilitySheet
  unitId?: string
  sheetId: string
  ranges: readonly CellRange[]
  userId: string
  mode?: SheetPasteMode
}
/** Additional operation points refine, and cannot override, worksheet/range protection. */
export type SheetClipboardPermissionEvaluator = (request: SheetClipboardPermissionRequest) => boolean | { allowed: boolean; reason?: string }
export interface SheetClipboardCapabilitiesInput {
  workbook: SheetClipboardCapabilityWorkbook
  sheetId?: string
  selection: { ranges: readonly CellRange[]; activeRangeIndex?: number }
  userId: string
  clipboard?: SheetClipboardSnapshot
  /** True only when a reader is callable. Its current content remains unknown until the user pastes. */
  systemClipboardAvailable: boolean
  active?: boolean
  editing?: boolean
  /** Plugin interest in an exclusive range, not the number of selected rectangles. */
  exclusiveRange?: boolean
  shapeTextEditing?: boolean
  permission?: SheetClipboardPermissionEvaluator
  /** Recent-paste replacements must also remain allowed to undo the previous mode's writes. */
  previousPasteMode?: SheetPasteMode
}
export interface SheetClipboardEntryCapability {
  hidden: boolean
  disabled: boolean
  reason?: string
  effectiveMode?: SheetPasteMode
}
export interface SheetClipboardSourceState {
  kind: 'none' | 'internal' | 'text' | 'html'
  id?: string
  copyType?: 'copy' | 'cut'
  formulaOnly: boolean
  hasValue: boolean
  hasFormula: boolean
  hasFormat: boolean
  hasColumnWidths: boolean
  rows: number
  columns: number
  sourcePresent: boolean
  sameWorkbook: boolean
  systemClipboardAvailable: boolean
}
export interface SheetClipboardCapabilities {
  copy: SheetClipboardEntryCapability
  cut: SheetClipboardEntryCapability
  copyFormulas: SheetClipboardEntryCapability
  paste: Record<SheetPasteMode, SheetClipboardEntryCapability>
  copySpecial: SheetClipboardEntryCapability
  pasteSpecial: SheetClipboardEntryCapability
  ribbonPaste: SheetClipboardEntryCapability
  source: SheetClipboardSourceState
}

const pasteModes: readonly SheetPasteMode[] = ['all', 'values', 'formulas', 'formats', 'columnWidths', 'exceptBorders']
const permissionReasons: Readonly<Record<SheetClipboardPermissionPoint, string>> = {
  copy: '当前用户没有复制此区域的权限。', view: '当前用户没有查看此区域的权限。', edit: '此区域受保护，当前用户没有编辑权限。',
  setCellValue: '当前用户没有写入单元格内容的权限。', setCellStyle: '当前用户没有设置单元格格式的权限。', setColumnWidth: '当前用户没有设置列宽的权限。',
}
function validRange(range: CellRange | undefined): range is CellRange {
  return !!range && [range.startRow, range.endRow, range.startCol, range.endCol].every(Number.isSafeInteger)
    && range.startRow >= 0 && range.startCol >= 0 && range.startRow <= range.endRow && range.startCol <= range.endCol
    && range.endRow < 1_048_576 && range.endCol < 16_384
}
function expandedRange(sheet: SheetClipboardCapabilitySheet, range: CellRange): CellRange {
  return expandSheetClipboardRange({ mergedRanges: sheet.mergedRanges ?? [] }, range)
}
function filteredSheetRows(workbook: SheetClipboardCapabilityWorkbook, sheet: SheetClipboardCapabilitySheet): number[] {
  return [...new Set([...(sheet.filteredRows ?? []), ...(sheet.autoFilter?.filteredRows ?? []), ...(workbook.tables ?? [])
    .filter(table => table.sheetId === sheet.id).flatMap(table => table.autoFilter?.filteredRows ?? [])])]
    .filter(row => Number.isSafeInteger(row) && row >= 0 && row < 1_048_576).sort((first, second) => first - second)
}
function rowCount(ranges: readonly CellRange[]): number { return ranges.reduce((total, range) => total + range.endRow - range.startRow + 1, 0) }

/** First N visible rows in intervals; it never creates one array entry for every row of a selection. */
function firstVisibleRows(startRow: number, count: number, startCol: number, endCol: number, filtered: readonly number[]): CellRange[] | undefined {
  const available = sheetCommandRangesWithoutFilteredRows([{ startRow, endRow: 1_048_575, startCol, endCol }], filtered)
  const result: CellRange[] = []
  let remaining = count
  for (const range of available) {
    if (remaining <= 0) break
    const length = Math.min(remaining, range.endRow - range.startRow + 1)
    result.push({ ...range, endRow: range.startRow + length - 1 }); remaining -= length
  }
  return remaining === 0 ? result : undefined
}
interface TargetPreview { ranges: CellRange[]; reason?: string; contentOnlyMerge?: boolean }
function targetPreview(input: SheetClipboardCapabilitiesInput, sheet: SheetClipboardCapabilitySheet, selected: CellRange, snapshot?: SheetClipboardSnapshot): TargetPreview {
  const expanded = expandedRange(sheet, selected), filtered = filteredSheetRows(input.workbook, sheet)
  if (!snapshot) {
    const ranges = sheetCommandRangesWithoutFilteredRows([expanded], filtered)
    return { ranges, reason: ranges.length ? undefined : '粘贴目标没有可见行。' }
  }
  const rows = snapshot.rows.length, columns = snapshot.columns.length
  if (!rows || !columns) return { ranges: [], reason: '剪贴板中没有可粘贴的范围。' }
  let ranges: CellRange[]
  if (snapshot.copyType === 'cut') {
    const range = { startRow: selected.startRow, endRow: selected.startRow + rows - 1, startCol: selected.startCol, endCol: selected.startCol + columns - 1 }
    if (!validRange(range)) return { ranges: [], reason: '粘贴后会超出工作表范围。' }
    ranges = [range]
  } else {
    if ((expanded.endRow - expanded.startRow + 1) * (expanded.endCol - expanded.startCol + 1) > SHEET_CLIPBOARD_SYNC_CELL_BUDGET) {
      return { ranges: [], reason: '粘贴范围超出当前同步处理预算。' }
    }
    const initial = sheetCommandRangesWithoutFilteredRows([expanded], filtered), initialRows = rowCount(initial)
    if (!initialRows) return { ranges: [], reason: '粘贴目标没有可见行。' }
    const merges = sheet.mergedRanges ?? [], width = expanded.endCol - expanded.startCol + 1
    const singleMerge = merges.find(merge => merge.startRow === expanded.startRow && merge.endRow === expanded.endRow
      && merge.startCol === expanded.startCol && merge.endCol === expanded.endCol)
    if (singleMerge && rows === 1 && columns === 1 && !snapshot.merges.length) {
      return { ranges: [{ ...singleMerge }], contentOnlyMerge: true,
        reason: initialRows === singleMerge.endRow - singleMerge.startRow + 1 ? undefined : '合并单元格不能跨过目标筛选行粘贴。' }
    }
    const repeat = !merges.some(merge => initial.some(range => sheetRangesIntersect(merge, range))) && initialRows % rows === 0 && width % columns === 0
    const endCol = expanded.startCol + (repeat ? width : columns) - 1, count = repeat ? initialRows : rows
    if (endCol >= 16_384 || count * (endCol - expanded.startCol + 1) > SHEET_CLIPBOARD_SYNC_CELL_BUDGET) {
      return { ranges: [], reason: '粘贴范围超出当前同步处理预算或工作表边界。' }
    }
    const visible = firstVisibleRows(expanded.startRow, count, expanded.startCol, endCol, filtered)
    if (!visible) return { ranges: [], reason: '粘贴后会超出工作表的最大行数。' }
    ranges = visible
  }
  for (const merge of sheet.mergedRanges ?? []) {
    if (ranges.some(range => sheetRangesIntersect(merge, range)) && !ranges.some(range => sheetRangeContainsCell(range, merge.startRow, merge.startCol)
      && sheetRangeContainsCell(range, merge.endRow, merge.endCol))) return { ranges, reason: '粘贴范围会覆盖部分合并单元格，请选择完整区域。' }
  }
  return { ranges }
}
function permissionReason(input: SheetClipboardCapabilitiesInput, sheet: SheetClipboardCapabilitySheet, ranges: readonly CellRange[],
  point: SheetClipboardPermissionPoint, phase: 'source' | 'target', mode?: SheetPasteMode): string | undefined {
  if (point === 'view' && (!canViewWorksheetProtection(sheet, input.userId) || !canViewProtectedRanges(sheet, ranges, input.userId))) return permissionReasons.view
  if (point === 'edit' && (!canEditWorksheetProtection(sheet, input.userId) || !canEditProtectedRanges(sheet, ranges, input.userId))) return permissionReasons.edit
  const decision = input.permission?.({ point, phase, workbook: input.workbook, sheet, unitId: input.workbook.id, sheetId: sheet.id, ranges, userId: input.userId, mode })
  if (decision === false || typeof decision === 'object' && !decision.allowed) return typeof decision === 'object' ? decision.reason ?? permissionReasons[point] : permissionReasons[point]
  return undefined
}
function firstPermissionReason(input: SheetClipboardCapabilitiesInput, sheet: SheetClipboardCapabilitySheet, ranges: readonly CellRange[],
  points: readonly SheetClipboardPermissionPoint[], phase: 'source' | 'target', mode?: SheetPasteMode): string | undefined {
  for (const point of points) {
    const reason = permissionReason(input, sheet, ranges, point, phase, mode)
    if (reason) return reason
  }
  return undefined
}
function sourceState(input: SheetClipboardCapabilitiesInput): SheetClipboardSourceState {
  const snapshot = input.clipboard
  let hasValue = false, hasFormula = false
  if (snapshot) for (const key in snapshot.cells) {
    const cell = snapshot.cells[key]!
    hasValue ||= cell.value !== null && cell.value !== ''
    hasFormula ||= !!cell.formula
    if (hasValue && hasFormula) break
  }
  return { kind: snapshot ? snapshot.externalFormat ?? 'internal' : 'none', id: snapshot?.id, copyType: snapshot?.copyType,
    formulaOnly: snapshot?.formulaOnly ?? false, hasValue, hasFormula,
    hasFormat: !!snapshot && snapshot.externalFormat !== 'text' && !!snapshot.rows.length && !!snapshot.columns.length,
    hasColumnWidths: !!snapshot && snapshot.externalFormat !== 'text' && !!snapshot.columns.length
      && snapshot.columnWidths.length === snapshot.columns.length && snapshot.columnWidths.every(width => Number.isFinite(width) && width >= 0 && width <= 2000),
    rows: snapshot?.rows.length ?? 0, columns: snapshot?.columns.length ?? 0, sourcePresent: !!snapshot?.source,
    sameWorkbook: !!snapshot?.source && snapshot.source.unitId === input.workbook.id, systemClipboardAvailable: input.systemClipboardAvailable }
}
function capability(hidden: boolean, reason?: string, effectiveMode?: SheetPasteMode): SheetClipboardEntryCapability {
  return { hidden, disabled: !!reason, ...(reason ? { reason } : {}), ...(effectiveMode ? { effectiveMode } : {}) }
}
function groupCapability(hidden: boolean, entries: readonly SheetClipboardEntryCapability[]): SheetClipboardEntryCapability {
  const visible = entries.filter(entry => !entry.hidden)
  return capability(hidden, hidden ? '当前上下文不显示此入口。' : visible.some(entry => !entry.disabled) ? undefined
    : visible.find(entry => entry.reason)?.reason ?? '当前没有可执行的剪贴板操作。')
}

/** Conservative entry prediction only. The command planner remains authoritative for actual candidate changes. */
export function sheetClipboardCapabilities(input: SheetClipboardCapabilitiesInput): SheetClipboardCapabilities {
  const sheet = input.workbook.sheets.find(candidate => candidate.id === input.sheetId)
  const selected = input.selection.ranges[input.selection.activeRangeIndex ?? -1] ?? input.selection.ranges.at(-1)
  const hidden = input.active === false || !sheet || sheet.visible === false
  const baseReason = hidden ? '当前没有可用的工作表。' : !validRange(selected) ? '请选择有效的单元格范围。'
    : input.editing ? '请先完成单元格或公式编辑。' : input.shapeTextEditing ? '当前正在编辑绘图对象文本。' : undefined
  const source = sourceState(input), specialHidden = hidden || !!input.exclusiveRange
  let copyReason = baseReason, cutReason = baseReason
  if (!baseReason && sheet && selected) {
    const range = expandedRange(sheet, selected), ranges = [range]
    copyReason = firstPermissionReason(input, sheet, ranges, ['view', 'copy'], 'source')
    cutReason = copyReason ?? firstPermissionReason(input, sheet, ranges, ['edit'], 'source')
    if ((range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1) > SHEET_CLIPBOARD_SYNC_CELL_BUDGET) {
      copyReason ??= '复制范围超出当前同步处理预算。'; cutReason ??= '剪切范围超出当前同步处理预算。'
    }
    if (!sheetCommandRangesWithoutFilteredRows(ranges, filteredSheetRows(input.workbook, sheet)).length
      && !(sheet.mergedRanges ?? []).some(merge => sheetRangesIntersect(merge, range))) copyReason ??= '所选区域没有可复制的可见行。'
  }
  function pasteCapability(mode: SheetPasteMode, snapshot: SheetClipboardSnapshot | undefined, reader: boolean): SheetClipboardEntryCapability {
    const effective = snapshot?.formulaOnly && mode === 'all' ? 'formulas' : mode
    if (baseReason || !sheet || !selected) return capability(hidden, baseReason ?? '请选择有效的单元格范围。', effective)
    if (!snapshot && !reader) return capability(hidden, '系统剪贴板不可读，工作台内也没有可粘贴内容。', effective)
    if (snapshot?.externalFormat === 'text' && (effective === 'formats' || effective === 'columnWidths')) return capability(hidden, '文本剪贴板没有源格式或列宽。', effective)
    if (effective === 'columnWidths' && snapshot && (!snapshot.columns.length || snapshot.columnWidths.length !== snapshot.columns.length
      || !snapshot.columnWidths.every(width => Number.isFinite(width) && width >= 0 && width <= 2000))) {
      return capability(hidden, '剪贴板中没有可用的列宽。', effective)
    }
    if (snapshot?.copyType === 'cut') {
      if (!snapshot.source || snapshot.source.unitId !== input.workbook.id) return capability(hidden, '跨工作簿剪切移动尚未接入，源内容已保留。', effective)
      const origin = input.workbook.sheets.find(candidate => candidate.id === snapshot.source!.sheetId)
      if (!origin) return capability(hidden, '剪切源工作表已移除，请重新剪切。', effective)
      const points: SheetClipboardPermissionPoint[] = ['view', 'copy', 'edit']
      if (effective === 'all') points.push('setCellValue', 'setCellStyle')
      const reason = firstPermissionReason(input, origin, [snapshot.source.range], points, 'source', effective)
      if (reason) return capability(hidden, reason, effective)
    }
    const preview = targetPreview(input, sheet, selected, snapshot)
    if (preview.reason) return capability(hidden, preview.reason, effective)
    const ranges = effective === 'columnWidths' ? [{ startRow: 0, endRow: Math.max(1, sheet.rowCount ?? 1_048_576) - 1,
      startCol: preview.ranges[0]!.startCol, endCol: preview.contentOnlyMerge ? preview.ranges[0]!.startCol : preview.ranges[0]!.endCol }] : preview.ranges
    const points: SheetClipboardPermissionPoint[] = ['view', 'edit']
    if (effective === 'columnWidths') points.push('setColumnWidth')
    else if (effective === 'formats') points.push('setCellStyle')
    else {
      points.push('setCellValue')
      // An unknown reader may contain plain text, which preserves destination formatting.
      if (snapshot && snapshot.externalFormat !== 'text' && (effective === 'all' || effective === 'exceptBorders')) points.push('setCellStyle')
    }
    return capability(hidden, firstPermissionReason(input, sheet, ranges, points, 'target', effective), effective)
  }
  const previous = input.previousPasteMode ? pasteCapability(input.previousPasteMode, input.clipboard, false) : undefined
  const paste = Object.fromEntries(pasteModes.map(mode => {
    const cached = pasteCapability(mode, input.clipboard, false)
    const unknown = input.systemClipboardAvailable ? pasteCapability(mode, undefined, true) : undefined
    // A cached cut or text snapshot cannot rule out different, valid system content.
    let entry = unknown && !unknown.disabled ? unknown : cached.disabled && unknown ? unknown : cached
    if (previous?.disabled && mode !== input.previousPasteMode && !entry.disabled) {
      entry = { ...entry, disabled: true, reason: previous.reason ?? '当前没有撤回上一次粘贴内容的权限。' }
    }
    return [mode, { ...entry, hidden: mode === 'all' ? hidden : specialHidden }]
  })) as Record<SheetPasteMode, SheetClipboardEntryCapability>
  const copy = capability(hidden, copyReason), cut = capability(hidden, cutReason), copyFormulas = capability(specialHidden, copyReason)
  const copySpecial = groupCapability(specialHidden, [copyFormulas])
  const pasteSpecial = groupCapability(specialHidden, pasteModes.filter(mode => mode !== 'all').map(mode => paste[mode]))
  const ribbonPaste = groupCapability(specialHidden || !!input.shapeTextEditing, pasteModes.map(mode => paste[mode]))
  return { copy, cut, copyFormulas, paste, copySpecial, pasteSpecial, ribbonPaste, source }
}

/** Canonical IDs and the fixed Univer aliases resolve through the same command registry. */
export function sheetClipboardCommandCapability(capabilities: SheetClipboardCapabilities, commandId: string, mode?: SheetPasteMode): SheetClipboardEntryCapability | undefined {
  if (commandId === 'sheet.menu.copy-special') return capabilities.copySpecial
  if (commandId === 'sheet.menu.paste-special') return capabilities.pasteSpecial
  if (commandId === 'sheet.menu.paste') return capabilities.ribbonPaste
  const action = resolveSheetClipboardCommand(commandId)
  if (!action) return undefined
  if (action.action === 'copy') return action.formulaOnly ? capabilities.copyFormulas : capabilities.copy
  if (action.action === 'cut') return capabilities.cut
  if (action.action === 'cancel-copy') return { hidden: false, disabled: false }
  return capabilities.paste[mode ?? action.mode ?? 'all']
}
