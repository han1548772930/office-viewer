import type { SheetDataValidationRule } from '../../domain/model.ts'
import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { sheetCellTextAlignment } from '../sheet-workbench/sheet-cell-alignment.ts'
import { sheetCellShrinkToFit, sheetCellTextRotation, sheetCellWrapStrategy, type SheetTextRotation } from '../sheet-workbench/sheet-cell-layout.ts'
import { layoutSheetCellText, layoutSheetStackedText, rotatedSheetTextSize, sheetCellTextFont, sheetCellTextIndent,
  type SheetCellTextFont, type SheetTextMeasure } from '../sheet-workbench/sheet-cell-text-layout.ts'
import { sheetColorHex } from '../sheet-workbench/sheet-colors.ts'
import { parseValidationListItems } from './validation-engine.ts'

/** Local Canvas padding is shared by painting and automatic dimensions; the Ant trigger occupies the trailing slot. */
export const SHEET_VALIDATION_TAG_PADDING = { cellX: 4, cellY: 4, tagX: 6, tagY: 3, gap: 4, icon: 22, radius: 4 } as const
export interface SheetValidationTagOption { label: string; color?: string }
export type SheetValidationTagOptionsAt = (row: number, column: number) => readonly SheetValidationTagOption[] | undefined
export interface SheetValidationTagLabel { text: string; color?: string }
export interface SheetValidationTagFragment { text: string; x: number; y: number; width: number; align: 'left' | 'center' }
export interface SheetValidationTagItem extends SheetValidationTagLabel {
  x: number; y: number; width: number; height: number; fragments: SheetValidationTagFragment[]
}
export interface SheetValidationTagLine { width: number; height: number; items: SheetValidationTagItem[] }
export interface SheetValidationTagCellLayout {
  font: SheetCellTextFont
  lines: SheetValidationTagLine[]
  width: number; height: number
  size: { width: number; height: number }
  rotation: number; direction: SheetTextRotation; scale: number
  leftInset: number; rightInset: number; verticalInset: number
  availableWidth: number; availableHeight: number
  paddingX: number; paddingY: number
}
export interface SheetValidationTagLayoutOptions {
  width: number
  height?: number
  zoom?: number
  measure: SheetTextMeasure
  /** Columns measure unwrapped label text; row sizing uses the current width and no height limit. */
  fit?: 'row' | 'column'
  /** The caller may supply the permission-aware, formatted candidate set without exposing workbook data to layout. */
  options?: readonly SheetValidationTagOption[]
}

export function sheetValidationListRuleAt(rules: readonly SheetDataValidationRule[] | undefined, row: number, column: number): SheetDataValidationRule | undefined {
  return rules?.find(rule => (rule.type === 'list' || rule.type === 'listMultiple')
    && rule.ranges.some(range => row >= range.startRow && row <= range.endRow && column >= range.startCol && column <= range.endCol))
}
export function sheetValidationCellArrowSpace(rule: SheetDataValidationRule | undefined, zoom = 1): number {
  return rule && (rule.type === 'list' || rule.type === 'listMultiple') && rule.renderMode !== 'text' && rule.showDropDown !== false
    ? SHEET_VALIDATION_TAG_PADDING.icon * zoom : 0
}
function canonicalTagOptions(rule: SheetDataValidationRule, options?: readonly SheetValidationTagOption[]): Map<string, SheetValidationTagOption> {
  const inline = rule.formula1?.trim().startsWith('=') ? [] : parseValidationListItems(rule.formula1 ?? '')
  const colors = rule.formula2?.split(',') ?? []
  const candidates = options ?? inline.map((label, index) => ({ label, color: sheetColorHex(colors[index]) }))
  const canonical = new Map<string, SheetValidationTagOption>()
  for (const option of candidates) {
    const key = option.label.trim().toLocaleLowerCase()
    if (!canonical.has(key)) canonical.set(key, option)
  }
  return canonical
}
/** The fixed single-list validator applies the selected color to arrow/plain-text cells without changing authored fill. */
export function sheetValidationSingleOptionColor(rule: SheetDataValidationRule | undefined, cell: ExcelCellView | undefined,
  options?: readonly SheetValidationTagOption[]): string | undefined {
  if (rule?.type !== 'list' || rule.renderMode !== 'arrow' && rule.renderMode !== 'text' || cell?.image) return undefined
  const value = String(cell?.rawValue ?? cell?.cachedValue ?? cell?.value ?? '').trim().toLocaleLowerCase()
  return value ? sheetColorHex(canonicalTagOptions(rule, options).get(value)?.color) : undefined
}
/** CSV storage and canonical labels are the same as typed commits. Layout never writes values or formulas. */
export function sheetValidationTagLabels(rule: SheetDataValidationRule | undefined, cell: ExcelCellView | undefined,
  options?: readonly SheetValidationTagOption[]): SheetValidationTagLabel[] | undefined {
  if (!rule || rule.type !== 'list' && rule.type !== 'listMultiple' || rule.renderMode === 'text'
    || rule.type === 'list' && rule.renderMode === 'arrow' || cell?.image) return undefined
  const raw = String(cell?.rawValue ?? cell?.cachedValue ?? cell?.value ?? '')
  const selected = rule.type === 'listMultiple' ? parseValidationListItems(`"${raw}"`) : raw ? [raw] : []
  const canonical = canonicalTagOptions(rule, options)
  return selected.map(text => {
    const match = canonical.get(text.trim().toLocaleLowerCase())
    return { text: match?.label ?? text, color: sheetColorHex(match?.color) }
  })
}

/** One pure layout supplies Canvas commands, automatic row height and automatic column width. */
export function layoutSheetValidationTagCell(rule: SheetDataValidationRule | undefined, cell: ExcelCellView | undefined,
  options: SheetValidationTagLayoutOptions): SheetValidationTagCellLayout | undefined {
  const labels = sheetValidationTagLabels(rule, cell, options.options)
  if (!labels) return undefined
  const zoom = options.zoom ?? 1, height = options.fit === 'row' ? Infinity : options.height ?? Infinity
  if (!Number.isFinite(zoom) || zoom <= 0 || !Number.isFinite(options.width) || options.width < 0 || Number.isNaN(height) || height < 0) return undefined
  const padding = SHEET_VALIDATION_TAG_PADDING, font = sheetCellTextFont(cell, zoom)
  const align = sheetCellTextAlignment(cell), direction = sheetCellTextRotation(cell), rotation = direction === 'v' ? 0 : direction
  const indent = align === 'center' ? 0 : sheetCellTextIndent(cell, font.fontSize)
  const leftInset = padding.cellX * zoom + (align === 'left' ? indent : 0)
  const rightInset = padding.cellX * zoom + sheetValidationCellArrowSpace(rule, zoom) + (align === 'right' ? indent : 0)
  const verticalInset = padding.cellY * zoom
  const availableWidth = Math.max(0, options.width - leftInset - rightInset), availableHeight = Math.max(0, height - verticalInset * 2)
  const paddingX = padding.tagX * zoom, paddingY = padding.tagY * zoom, gap = padding.gap * zoom
  const wrap = options.fit !== 'column' && sheetCellWrapStrategy(cell) === 'wrap'
  const measure = (text: string) => {
    const width = options.measure(text, font.font)
    return Number.isFinite(width) ? Math.max(0, width) : 0
  }
  const items = labels.map(label => {
    const fragments: SheetValidationTagFragment[] = []
    let textWidth: number, textHeight: number
    if (direction === 'v') {
      // Stacked labels keep complete graphemes; fixed/manual row heights clip instead of changing their authoring order.
      const layout = layoutSheetStackedText(label.text, { height: Infinity, lineHeight: font.fontSize, wrap: false, measure })
      textWidth = layout.width; textHeight = layout.height
      let x = paddingX
      for (const column of layout.columns) {
        column.glyphs.forEach((glyph, index) => fragments.push({ text: glyph.text, x: x + column.width / 2,
          y: paddingY + (index + .5) * font.fontSize, width: measure(glyph.text), align: 'center' }))
        x += column.width
      }
    } else {
      const layout = layoutSheetCellText(label.text, { width: Math.max(0, availableWidth - paddingX * 2), lineHeight: font.fontSize, wrap, measure })
      textWidth = layout.width; textHeight = layout.height
      layout.lines.forEach((line, index) => fragments.push({ text: line.text, x: paddingX,
        y: paddingY + (index + .5) * font.fontSize, width: line.width, align: 'left' }))
    }
    return { ...label, x: 0, y: 0, width: Math.max(12 * zoom, textWidth + paddingX * 2),
      height: Math.max(font.fontSize, textHeight) + paddingY * 2, fragments }
  })
  const lines: SheetValidationTagLine[] = []
  for (const item of items) {
    let line = lines.at(-1)
    // Multiple selections retain the fixed widget's item wrapping, including explicit non-wrapping cell styles.
    if (!line || rule?.type === 'listMultiple' && line.items.length && line.width + gap + item.width > availableWidth) {
      line = { width: 0, height: 0, items: [] }; lines.push(line)
    }
    item.x = line.items.length ? line.width + gap : 0
    line.items.push(item); line.width = item.x + item.width; line.height = Math.max(line.height, item.height)
  }
  const width = lines.reduce((maximum, line) => Math.max(maximum, line.width), 0)
  let blockHeight = 0
  for (const line of lines) {
    const x = align === 'center' ? (width - line.width) / 2 : align === 'right' ? width - line.width : 0
    for (const item of line.items) { item.x += x; item.y = blockHeight + (line.height - item.height) / 2 }
    blockHeight += line.height + gap
  }
  if (lines.length) blockHeight -= gap
  const rotated = rotatedSheetTextSize(width, blockHeight, rotation)
  const scale = options.fit !== 'column' && sheetCellShrinkToFit(cell)
    ? Math.max(0, Math.min(1, rotated.width > 0 ? availableWidth / rotated.width : 1, rotated.height > 0 ? availableHeight / rotated.height : 1)) : 1
  return { font, lines, width, height: blockHeight, size: { width: rotated.width * scale, height: rotated.height * scale }, rotation, direction, scale,
    leftInset, rightInset, verticalInset, availableWidth, availableHeight, paddingX, paddingY }
}
