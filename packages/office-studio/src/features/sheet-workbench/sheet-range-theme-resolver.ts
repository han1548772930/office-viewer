import { SHEET_RANGE_THEME_SECTIONS, type CellRange, type SheetCellStylePatch, type SheetRangeThemeRule,
  type SheetRangeThemeStyle, type SheetTableModel, type SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { canonicalSheetAxisSpans, effectiveSheetRowSizes, type SheetAxisDimensions, type SheetAxisSpan } from '../../domain/sheet-axis.ts'
import { SHEET_BUILTIN_RANGE_THEMES } from './sheet-range-theme-presets.ts'

export interface SheetRangeThemeContext { rangeThemes?: Readonly<Record<string, SheetRangeThemeStyle>>; tables?: readonly SheetTableModel[] }
export interface SheetRangeThemeSheet extends SheetAxisDimensions {
  id?: string
  /** Detached sheet views carry the workbook catalog; domain worksheets receive it via the context. */
  rangeThemes?: Readonly<Record<string, SheetRangeThemeStyle>>
  rangeThemeRules?: readonly SheetRangeThemeRule[]
  hiddenRows?: readonly number[]
  filteredRows?: readonly number[]
  outlineHiddenRows?: readonly number[]
  outline?: SpreadsheetSheet['outline']
  autoFilter?: SpreadsheetSheet['autoFilter']
}
export function sheetRangeThemeByName(name: string, context: SheetRangeThemeContext = {}): SheetRangeThemeStyle | undefined {
  return Object.hasOwn(SHEET_BUILTIN_RANGE_THEMES, name) ? SHEET_BUILTIN_RANGE_THEMES[name]
    : context.rangeThemes && Object.hasOwn(context.rangeThemes, name) ? context.rangeThemes[name] : undefined
}
export function sheetRangeThemeRangesIntersect(first: CellRange, second: CellRange): boolean {
  return first.startRow <= second.endRow && first.endRow >= second.startRow && first.startCol <= second.endCol && first.endCol >= second.startCol
}
/** Latest applied rule wins on overlap; exact range replacement is handled by the content command. */
export function sheetUsedRangeThemeRule(sheet: SheetRangeThemeSheet, range: CellRange): SheetRangeThemeRule | undefined {
  return sheet.rangeThemeRules?.findLast(rule => sheetRangeThemeRangesIntersect(rule.range, range))
}

/** Sparse size exceptions and compact manual/outline spans determine zebra visibility without scanning every row. */
export function sheetRangeThemeHiddenRowSpans(sheet: SheetRangeThemeSheet, context: SheetRangeThemeContext = {}): SheetAxisSpan<boolean>[] {
  const effective = effectiveSheetRowSizes(sheet)
  const zero = canonicalSheetAxisSpans([
    ...(effective.defaultSize <= 0 ? [{ start: 0, end: 1_048_575, value: true }] : []),
    ...effective.spans.filter(span => span.value <= 0).map(span => ({ ...span, value: true })),
    ...Object.entries(effective.sizes).map(([key, value]) => ({ start: Number(key), end: Number(key), value: value <= 0 })),
  ]).filter(span => span.value)
  const points = new Set([...(sheet.hiddenRows ?? []), ...(sheet.filteredRows ?? []), ...(sheet.outlineHiddenRows ?? []),
    ...(sheet.autoFilter?.filteredRows ?? []), ...(context.tables ?? []).filter(table => table.sheetId === sheet.id).flatMap(table => table.autoFilter?.filteredRows ?? [])])
  return canonicalSheetAxisSpans([
    ...zero, ...canonicalSheetAxisSpans(sheet.rowHiddenSpans ?? []).filter(span => span.value),
    ...(sheet.outline?.groups ?? []).filter(group => group.axis === 'row' && group.collapsed).map(group => ({ start: group.start, end: group.end, value: true })),
    ...[...points].filter(row => Number.isSafeInteger(row) && row >= 0 && row < 1_048_576).map(row => ({ start: row, end: row, value: true })),
  ])
}
export function sheetRangeThemeHiddenCount(spans: readonly SheetAxisSpan<boolean>[], start: number, end: number): number {
  let count = 0
  for (const span of spans) {
    if (span.start > end) break
    if (span.value) count += Math.max(0, Math.min(end, span.end) - Math.max(start, span.start) + 1)
  }
  return count
}
/** Theme sections compose individual properties; unlike cell direct formatting, borders merge per edge. */
export function composeSheetRangeThemeSections(sections: readonly (SheetCellStylePatch | null | undefined)[]): SheetCellStylePatch {
  const result: SheetCellStylePatch = {}
  for (const section of sections) {
    if (!section) continue
    const { font, border, metadata, ...rest } = cloneModel(section)
    Object.assign(result, rest)
    if (font) result.font = { ...result.font, ...font }
    if (border) result.border = { ...result.border, ...border }
    if (metadata) result.metadata = { ...result.metadata, ...metadata,
      ...((result.metadata?.sourceStyle || metadata.sourceStyle) ? { sourceStyle: {
        ...(result.metadata?.sourceStyle as Record<string, unknown> | undefined), ...(metadata.sourceStyle as Record<string, unknown> | undefined),
      } } : {}) }
  }
  return result
}
export function sheetRangeThemeStyleAt(theme: SheetRangeThemeStyle, offsetRow: number, offsetCol: number,
  isLastRow: boolean, isLastCol: boolean, isToggled = false): SheetCellStylePatch {
  const rowSection = offsetRow % 2 === 1 ? (isToggled ? 'secondRowStyle' : 'firstRowStyle')
    : offsetRow % 2 === 0 ? (isToggled ? 'firstRowStyle' : 'secondRowStyle') : undefined
  const columnSection = offsetCol % 2 === 1 ? 'firstColumnStyle' : offsetCol % 2 === 0 ? 'secondColumnStyle' : undefined
  const selected = new Set([...(offsetRow >= 0 && offsetCol >= 0 ? ['wholeStyle'] : []),
    ...(rowSection ? [rowSection] : []), ...(columnSection ? [columnSection] : []),
    ...(offsetCol === 0 ? ['headerColumnStyle'] : []), ...(isLastCol ? ['lastColumnStyle'] : []),
    ...(offsetRow === 0 ? ['headerRowStyle'] : []), ...(isLastRow ? ['lastRowStyle'] : [])])
  return composeSheetRangeThemeSections(SHEET_RANGE_THEME_SECTIONS.filter(section => selected.has(section)).map(section => theme[section]))
}
export function resolveSheetRangeThemeStyle(sheet: SheetRangeThemeSheet, row: number, col: number,
  context: SheetRangeThemeContext = {}): SheetCellStylePatch | undefined {
  const rule = sheetUsedRangeThemeRule(sheet, { startRow: row, endRow: row, startCol: col, endCol: col })
  if (!rule) return undefined
  const theme = sheetRangeThemeByName(rule.themeName, { ...context, rangeThemes: context.rangeThemes ?? sheet.rangeThemes })
  if (!theme) return undefined
  return rangeThemeRuleStyle(rule.styleRange ?? rule.range, theme, row, col, sheetRangeThemeHiddenRowSpans(sheet, context))
}
function rangeThemeRuleStyle(range: CellRange, theme: SheetRangeThemeStyle, row: number, col: number,
  hidden: readonly SheetAxisSpan<boolean>[]): SheetCellStylePatch {
  let lastRow = range.endRow
  for (let index = hidden.length - 1; index >= 0; index--) {
    const span = hidden[index]!
    if (span.value && lastRow >= span.start && lastRow <= span.end) lastRow = span.start - 1
  }
  const toggled = sheetRangeThemeHiddenCount(hidden, range.startRow, row - 1) % 2 === 1
  return sheetRangeThemeStyleAt(theme, row - range.startRow, col - range.startCol,
    row === lastRow && lastRow >= range.startRow, col === range.endCol, toggled)
}

export type SheetRangeThemeRegionAxis = readonly number[] | { start: number; end: number }
export interface SheetRangeThemeRegionSnapshot {
  rows: SheetRangeThemeRegionAxis
  columns: SheetRangeThemeRegionAxis
  /** Unresolved rules remain as blockers so a lower overlapping theme cannot leak into a copied region. */
  rules: Array<{ range: CellRange; styleRange?: CellRange; theme: SheetRangeThemeStyle | null }>
  hiddenRowSpans: SheetAxisSpan<boolean>[]
}
const regionPosition = (axis: SheetRangeThemeRegionAxis, index: number) => 'start' in axis ? axis.start + index : axis[index]!
const regionLength = (axis: SheetRangeThemeRegionAxis) => 'start' in axis ? axis.end - axis.start + 1 : axis.length
export function captureSheetRangeThemeRegion(sheet: SheetRangeThemeSheet, rows: SheetRangeThemeRegionAxis, columns: SheetRangeThemeRegionAxis,
  context: SheetRangeThemeContext = {}): SheetRangeThemeRegionSnapshot | undefined {
  const range = { startRow: regionPosition(rows, 0), endRow: regionPosition(rows, regionLength(rows) - 1),
    startCol: regionPosition(columns, 0), endCol: regionPosition(columns, regionLength(columns) - 1) }
  const rules = (sheet.rangeThemeRules ?? []).flatMap(rule => {
    const theme = sheetRangeThemeByName(rule.themeName, { ...context, rangeThemes: context.rangeThemes ?? sheet.rangeThemes })
    return sheetRangeThemeRangesIntersect(rule.range, range) ? [{ range: cloneModel(rule.range),
      ...(rule.styleRange ? { styleRange: cloneModel(rule.styleRange) } : {}), theme: theme ? cloneModel(theme) : null }] : []
  })
  return rules.length ? { rows: cloneModel(rows), columns: cloneModel(columns), rules, hiddenRowSpans: sheetRangeThemeHiddenRowSpans(sheet, context) } : undefined
}
export function sheetRangeThemeRegionStyle(snapshot: SheetRangeThemeRegionSnapshot, row: number, col: number): SheetCellStylePatch | undefined {
  const sourceRow = regionPosition(snapshot.rows, row), sourceCol = regionPosition(snapshot.columns, col)
  const range = { startRow: sourceRow, endRow: sourceRow, startCol: sourceCol, endCol: sourceCol }
  const rule = snapshot.rules.findLast(item => sheetRangeThemeRangesIntersect(item.range, range))
  return rule?.theme ? rangeThemeRuleStyle(rule.styleRange ?? rule.range, rule.theme, sourceRow, sourceCol, snapshot.hiddenRowSpans) : undefined
}
