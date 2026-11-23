import type {
  DocumentCell, DocumentTable, DocumentTableBorder, DocumentTableBorderSide, DocumentTableCellMargin,
  DocumentTableColumnTypeConfig,
} from './document-model'
import type { DocumentTableGrid, DocumentTableGridCell } from './document-table-grid'

export const DOCUMENT_TABLE_MIN_ROW_HEIGHT = 34
export const DOCUMENT_TABLE_DEFAULT_MARGIN: DocumentTableCellMargin = { top: 8, right: 8, bottom: 8, left: 8 }
export const DOCUMENT_TABLE_DEFAULT_BORDER: DocumentTableBorder = { width: 1, color: '#aab3bf', style: 'solid' }
export const DOCUMENT_TABLE_BORDER_SIDES: readonly DocumentTableBorderSide[] = ['top', 'right', 'bottom', 'left']
export type DocumentTableBorderPreset = 'all' | 'inner' | 'outer' | 'top' | 'bottom' | 'left' | 'right' | 'none' | 'innerHorizontal' | 'innerVertical'
export const DOCUMENT_TABLE_BORDER_PRESETS: readonly DocumentTableBorderPreset[] = ['all', 'inner', 'outer', 'top', 'bottom', 'left', 'right', 'none', 'innerHorizontal', 'innerVertical']

export function assertDocumentTableColor(color: string | null): void {
  if (color === null) return
  if (typeof color !== 'string' || !/^(?:#[\da-f]{3,4}|#[\da-f]{6}|#[\da-f]{8}|[a-z]+|(?:rgb|rgba|hsl|hsla)\([\d\s.,%+\-/]+\))$/i.test(color)) {
    throw new TypeError('表格颜色必须是有效的 CSS 颜色')
  }
}

export function assertDocumentTableMargin(margin: DocumentTableCellMargin): void {
  if (!margin || !DOCUMENT_TABLE_BORDER_SIDES.every(side => Number.isFinite(margin[side]) && margin[side] >= 0)) throw new RangeError('表格单元格边距必须是非负 CSS px')
}

export function assertDocumentTableBorder(border: DocumentTableBorder): void {
  if (!border || !Number.isFinite(border.width) || border.width < 0 || !['solid', 'dot', 'dash'].includes(border.style)) throw new RangeError('表格边框宽度或样式无效')
  assertDocumentTableColor(border.color)
}

export function assertDocumentTableBorderPatch(patch: Partial<DocumentTableBorder>): void {
  if (!patch || typeof patch !== 'object' || Array.isArray(patch) || Object.keys(patch).some(key => !['width', 'color', 'style'].includes(key))
    || !Object.values(patch).some(value => value !== undefined)) throw new TypeError('表格边框更改必须包含颜色、宽度或样式')
  assertDocumentTableBorder(documentTableBorderWithPatch(DOCUMENT_TABLE_DEFAULT_BORDER, patch))
}

/** Optional DTO fields have the same meaning whether omitted or explicitly undefined. */
export function documentTableBorderWithPatch(border: DocumentTableBorder, patch: Partial<DocumentTableBorder>): DocumentTableBorder {
  return {
    width: patch.width === undefined ? border.width : patch.width,
    color: patch.color === undefined ? border.color : patch.color,
    style: patch.style === undefined ? border.style : patch.style,
  }
}

export function assertDocumentTableMarginPatch(patch: Partial<DocumentTableCellMargin>): void {
  if (!patch || typeof patch !== 'object' || Array.isArray(patch) || Object.keys(patch).some(key => !(DOCUMENT_TABLE_BORDER_SIDES as readonly string[]).includes(key))
    || !Object.values(patch).some(value => value !== undefined)) throw new TypeError('表格边距更改必须包含至少一个方向')
  for (const side of DOCUMENT_TABLE_BORDER_SIDES) if (patch[side] !== undefined && (!Number.isFinite(patch[side]) || patch[side]! < 0)) throw new RangeError('表格单元格边距必须是非负 CSS px')
}

export interface DocumentTableCommonValue<T> { value: T | undefined; mixed: boolean }
export function documentTableCommonValue<T>(values: readonly T[]): DocumentTableCommonValue<T> {
  const value = values[0], mixed = values.length > 1 && values.some(candidate => !Object.is(candidate, value))
  return { value: mixed ? undefined : value, mixed }
}

export function assertDocumentTableColumnType(config: DocumentTableColumnTypeConfig): void {
  if (!config || !['text', 'number', 'date', 'checkbox', 'dropdown'].includes(config.type)) throw new TypeError('表格列类型无效')
  if (config.dropdownOptions !== undefined) {
    if (!Array.isArray(config.dropdownOptions)) throw new TypeError('表格下拉选项必须是数组')
    const ids = new Set<string>()
    for (const option of config.dropdownOptions) {
      if (!option || typeof option.id !== 'string' || !option.id || ids.has(option.id) || typeof option.label !== 'string') throw new TypeError('表格下拉选项标识或标签无效')
      ids.add(option.id)
      if (option.color !== undefined) assertDocumentTableColor(option.color)
    }
  }
}

export function documentTableCellMargin(table: DocumentTable, cell?: DocumentCell): DocumentTableCellMargin {
  const margin = cell?.style.margin ?? table.style?.cellMargin
  if (margin) assertDocumentTableMargin(margin)
  return margin ? { ...margin } : { ...DOCUMENT_TABLE_DEFAULT_MARGIN }
}

export function documentTableCellFill(table: DocumentTable, cell: DocumentCell): string {
  const fill = cell.style.fill ?? cell.style.background ?? table.style?.fill ?? table.style?.background
  return typeof fill === 'string' ? fill : '#ffffff'
}

export function documentTableCellBorder(grid: DocumentTableGrid, owner: DocumentTableGridCell, side: DocumentTableBorderSide): DocumentTableBorder {
  const outer = side === 'top' ? owner.row === 0 : side === 'bottom' ? owner.row + owner.rowSpan === grid.rowCount
    : side === 'left' ? owner.column === 0 : owner.column + owner.colSpan === grid.columnCount
  const tableSide = outer ? side : side === 'top' || side === 'bottom' ? 'insideHorizontal' : 'insideVertical'
  const border = owner.cell.style.borders?.[side] ?? grid.table.style?.borders?.[tableSide] ?? DOCUMENT_TABLE_DEFAULT_BORDER
  assertDocumentTableBorder(border)
  return { ...border }
}

export function documentTableCellBorders(grid: DocumentTableGrid, owner: DocumentTableGridCell): Record<DocumentTableBorderSide, DocumentTableBorder> {
  return { top: documentTableCellBorder(grid, owner, 'top'), right: documentTableCellBorder(grid, owner, 'right'),
    bottom: documentTableCellBorder(grid, owner, 'bottom'), left: documentTableCellBorder(grid, owner, 'left') }
}

/** Explicit column widths retain their actual size, including when the flow column changes. */
export function documentTableColumnWidths(table: DocumentTable, columnCount: number, availableWidth: number): number[] {
  if (!Number.isFinite(availableWidth) || availableWidth <= 0) throw new RangeError('表格可用宽度必须大于零')
  if (table.columnWidths?.length) {
    if (table.columnWidths.length !== columnCount || table.columnWidths.some(width => !Number.isFinite(width) || width <= 0)) throw new RangeError('表格列宽与网格不一致')
    return [...table.columnWidths]
  }
  const width = table.style?.width ?? availableWidth
  if (!Number.isFinite(width) || width <= 0) throw new RangeError('表格宽度必须大于零')
  return Array.from({ length: columnCount }, () => width / columnCount)
}

export function documentTableHorizontalOffset(table: DocumentTable, width: number, availableWidth: number): number {
  const indent = table.style?.indent ?? 0
  if (!Number.isFinite(indent)) throw new RangeError('表格缩进必须是有限 CSS px')
  return indent + (table.style?.align === 'center' ? (availableWidth - width) / 2 : table.style?.align === 'right' ? availableWidth - width : 0)
}

export function documentTableRowProperties(table: DocumentTable, row: number): {
  height?: number; heightRule: 'auto' | 'atLeast' | 'exact'; allowSplit: boolean
} {
  const source = table.rows[row], legacy = table.rowStyles?.[row]
  const rawHeight = source?.height ?? legacy?.height ?? legacy?.heightPx
  const height = typeof rawHeight === 'number' ? rawHeight : undefined
  const rawRule = source?.heightRule ?? legacy?.heightRule
  const heightRule = rawRule === 'auto' || rawRule === 'exact' || rawRule === 'atLeast' ? rawRule : height === undefined ? 'auto' : 'atLeast'
  if (height !== undefined && (!Number.isFinite(height) || height <= 0)) throw new RangeError('表格行高必须大于零')
  return { height, heightRule, allowSplit: source?.allowSplit ?? (legacy?.allowSplit !== false) }
}

export interface DocumentTableRowPropertiesPatch {
  height?: number
  heightRule?: 'auto' | 'atLeast' | 'exact'
  allowSplit?: boolean
  /** Height-only resize retains its existing atLeast default unless explicitly preserved. */
  preserveHeightRule?: boolean
}

export function assertDocumentTableRowPropertiesPatch(table: DocumentTable, rows: readonly number[], patch: DocumentTableRowPropertiesPatch): void {
  if (patch.height !== undefined && (!Number.isFinite(patch.height) || patch.height <= 0)) throw new RangeError('表格行高必须为正 CSS px')
  if (patch.heightRule !== undefined && !['auto', 'atLeast', 'exact'].includes(patch.heightRule)) throw new TypeError('表格行高规则无效')
  if (patch.allowSplit !== undefined && typeof patch.allowSplit !== 'boolean') throw new TypeError('表格跨页拆行参数无效')
  if (patch.preserveHeightRule !== undefined && typeof patch.preserveHeightRule !== 'boolean') throw new TypeError('表格行高保留参数无效')
  if (patch.height === undefined && patch.heightRule === undefined && patch.allowSplit === undefined) throw new RangeError('请更改至少一项行属性')
  for (const row of rows) {
    if (!Number.isSafeInteger(row) || row < 0 || row >= table.rows.length) throw new RangeError('表格尺寸行超出范围')
    if (patch.heightRule !== undefined && patch.heightRule !== 'auto' && patch.height === undefined && documentTableRowProperties(table, row).height === undefined) {
      throw new RangeError('所选行包含自动行高，请输入固定值或最小行高')
    }
  }
}

/** Validate the whole patch first so an invalid later row cannot leave a partial edit. */
export function applyDocumentTableRowPropertiesPatch(table: DocumentTable, rows: readonly number[], patch: DocumentTableRowPropertiesPatch): void {
  assertDocumentTableRowPropertiesPatch(table, rows, patch)
  for (const row of rows) {
    const source = table.rows[row]!, previousRule = documentTableRowProperties(table, row).heightRule
    if (patch.height !== undefined) source.height = patch.height
    if (patch.heightRule !== undefined) source.heightRule = patch.heightRule
    else if (patch.height !== undefined) source.heightRule = patch.preserveHeightRule ? previousRule : 'atLeast'
    if (patch.allowSplit !== undefined) source.allowSplit = patch.allowSplit
  }
}
