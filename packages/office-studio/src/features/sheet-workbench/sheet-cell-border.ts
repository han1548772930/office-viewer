import type { CellRange } from '../../domain/model.ts'
import { sheetColorHex } from './sheet-colors.ts'

export const sheetBorderLabels = {
  top: '上框线', bottom: '下框线', left: '左框线', right: '右框线', none: '无框线',
  all: '所有边框', outside: '外侧框线', inside: '内部框线', horizontal: '内侧横线', vertical: '内侧竖线',
  'diagonal-down': '右下斜线', 'diagonal-down-pair': '右下双斜线', 'diagonal-down-triple': '右下三斜线',
  'diagonal-up': '右上斜线', 'diagonal-up-pair': '右上双斜线',
} as const
export type SheetBorderPreset = keyof typeof sheetBorderLabels
export type SheetBorderStyle = 'thin' | 'hair' | 'dotted' | 'dashed' | 'dash-dot' | 'dash-dot-dot' | 'medium' | 'medium-dashed' | 'medium-dash-dot' | 'medium-dash-dot-dot' | 'thick' | 'double'
export type SheetBorderLineStyle = SheetBorderStyle | 'none' | 'slant-dash-dot'
export interface SheetBorderOptions { style: SheetBorderLineStyle; color: string }
export const sheetBorderStyleLabels: Record<SheetBorderStyle, string> = {
  thin: '细实线', hair: '极细线', dotted: '点线', dashed: '虚线', 'dash-dot': '点划线', 'dash-dot-dot': '双点划线',
  medium: '中等实线', 'medium-dashed': '中等虚线', 'medium-dash-dot': '中等点划线', 'medium-dash-dot-dot': '中等双点划线', thick: '粗实线', double: '双线',
}

/** Fixed public enum values; NONE and SLANT_DASH_DOT are API pens, not extra menu entries. */
export const SheetBorderStyleTypes = Object.freeze({
  NONE: 0, THIN: 1, HAIR: 2, DOTTED: 3, DASHED: 4, DASH_DOT: 5, DASH_DOT_DOT: 6,
  DOUBLE: 7, MEDIUM: 8, MEDIUM_DASHED: 9, MEDIUM_DASH_DOT: 10, MEDIUM_DASH_DOT_DOT: 11,
  SLANT_DASH_DOT: 12, THICK: 13,
})
export const SheetBorderTypes = Object.freeze({
  TOP: 'top', BOTTOM: 'bottom', LEFT: 'left', RIGHT: 'right', NONE: 'none', ALL: 'all',
  OUTSIDE: 'outside', INSIDE: 'inside', HORIZONTAL: 'horizontal', VERTICAL: 'vertical',
  TLBR: 'tlbr', TLBC_TLMR: 'tlbc_tlmr', TLBR_TLBC_TLMR: 'tlbr_tlbc_tlmr', BLTR: 'bl_tr', MLTR_BCTR: 'mltr_bctr',
} as const)
export type SheetBorderType = typeof SheetBorderTypes[keyof typeof SheetBorderTypes]
export type SheetBorderStyleType = typeof SheetBorderStyleTypes[keyof typeof SheetBorderStyleTypes]

const numericStyles: readonly SheetBorderLineStyle[] = [
  'none', 'thin', 'hair', 'dotted', 'dashed', 'dash-dot', 'dash-dot-dot', 'double', 'medium',
  'medium-dashed', 'medium-dash-dot', 'medium-dash-dot-dot', 'slant-dash-dot', 'thick',
]
const presetAliases: Readonly<Record<string, SheetBorderPreset>> = {
  tlbr: 'diagonal-down', tlbc_tlmr: 'diagonal-down-pair', tlbr_tlbc_tlmr: 'diagonal-down-triple',
  bl_tr: 'diagonal-up', mltr_bctr: 'diagonal-up-pair',
}
export const SHEET_BORDER_SIDES = ['top', 'bottom', 'left', 'right'] as const
export type SheetBorderSide = typeof SHEET_BORDER_SIDES[number]
const diagonalKeys = ['diagonal', 'diagonalDown', 'diagonalUp', 'topLeftBottomCenter', 'topLeftMiddleRight', 'middleLeftTopRight', 'bottomCenterTopRight'] as const
const diagonalPresetKeys: Readonly<Partial<Record<SheetBorderPreset, readonly string[]>>> = {
  'diagonal-down': ['diagonalDown'], 'diagonal-down-pair': ['topLeftBottomCenter', 'topLeftMiddleRight'],
  'diagonal-down-triple': ['diagonalDown', 'topLeftBottomCenter', 'topLeftMiddleRight'],
  'diagonal-up': ['diagonalUp'], 'diagonal-up-pair': ['middleLeftTopRight', 'bottomCenterTopRight'],
}
const sideOffsets: Readonly<Record<SheetBorderSide, { row: number; col: number; opposite: SheetBorderSide }>> = {
  top: { row: -1, col: 0, opposite: 'bottom' }, bottom: { row: 1, col: 0, opposite: 'top' },
  left: { row: 0, col: -1, opposite: 'right' }, right: { row: 0, col: 1, opposite: 'left' },
}
const maxRow = 1_048_575, maxCol = 16_383

export function isSheetBorderPreset(value: unknown): value is SheetBorderPreset {
  return typeof value === 'string' && Object.hasOwn(sheetBorderLabels, value)
}
export function normalizeSheetBorderPreset(value: unknown): SheetBorderPreset {
  if (isSheetBorderPreset(value)) return value
  if (typeof value === 'string' && Object.hasOwn(presetAliases, value)) return presetAliases[value]!
  throw new TypeError('边框位置无效。')
}
export function normalizeSheetBorderStyle(value: unknown): SheetBorderLineStyle {
  const style = typeof value === 'number' && Number.isInteger(value) ? numericStyles[value] : value
  if (typeof style === 'string' && numericStyles.includes(style as SheetBorderLineStyle)) return style as SheetBorderLineStyle
  throw new TypeError('边框线型无效。')
}
export function normalizeSheetBorderColor(value: unknown): string {
  if (typeof value !== 'string' || !value.trim()) throw new TypeError('边框颜色必须是非空字符串。')
  return sheetColorHex(value) ?? value
}
export function assertSheetBorderRanges(ranges: readonly CellRange[]): void {
  for (const range of ranges) if (!range || ![range.startRow, range.startCol, range.endRow, range.endCol].every(Number.isSafeInteger)
    || range.startRow < 0 || range.startCol < 0 || range.endRow < range.startRow || range.endCol < range.startCol
    || range.endRow > maxRow || range.endCol > maxCol) throw new RangeError('边框选区无效。')
}

export interface SheetBorderPosition { row: number; col: number }
export interface SheetBorderSource {
  readBorder: (row: number, col: number) => Record<string, unknown> | undefined
  existingCells: () => Iterable<SheetBorderPosition>
  mergedRanges?: readonly CellRange[]
}
export interface SheetBorderPatch extends SheetBorderPosition { border: Record<string, unknown> }
export interface SheetBorderPlan { ranges: CellRange[]; patches: SheetBorderPatch[] }
export interface SheetMergedBorderSegment extends SheetBorderPosition { side: SheetBorderSide; edge: unknown }

/**
 * An explicit null on a physical segment suppresses the imported main-cell fallback.
 * A viewport clips only the physical perimeter; fallback still reads the original main cell.
 */
export function sheetMergedBorderSegments(
  range: CellRange, readBorder: SheetBorderSource['readBorder'], visibleRange?: CellRange,
): SheetMergedBorderSegment[] {
  const primary = readBorder(range.startRow, range.startCol)
  const segments: SheetMergedBorderSegment[] = []
  for (const side of SHEET_BORDER_SIDES) for (const { row, col } of sidePositions(range, side, visibleRange)) {
    const actual = readBorder(row, col)
    segments.push({ side, row, col, edge: actual && Object.hasOwn(actual, side) ? actual[side] : primary?.[side] })
  }
  return segments
}

function contains(range: CellRange, row: number, col: number): boolean {
  return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
}
function intersects(left: CellRange, right: CellRange): boolean {
  return left.startRow <= right.endRow && left.endRow >= right.startRow && left.startCol <= right.endCol && left.endCol >= right.startCol
}
function* sidePositions(range: CellRange, side: SheetBorderSide, visibleRange?: CellRange): Generator<SheetBorderPosition> {
  if (side === 'top' || side === 'bottom') {
    const row = side === 'top' ? range.startRow : range.endRow
    if (visibleRange && (row < visibleRange.startRow || row > visibleRange.endRow)) return
    const startCol = Math.max(range.startCol, visibleRange?.startCol ?? range.startCol)
    const endCol = Math.min(range.endCol, visibleRange?.endCol ?? range.endCol)
    for (let col = startCol; col <= endCol; col += 1) yield { row, col }
  } else {
    const col = side === 'left' ? range.startCol : range.endCol
    if (visibleRange && (col < visibleRange.startCol || col > visibleRange.endCol)) return
    const startRow = Math.max(range.startRow, visibleRange?.startRow ?? range.startRow)
    const endRow = Math.min(range.endRow, visibleRange?.endRow ?? range.endRow)
    for (let row = startRow; row <= endRow; row += 1) yield { row, col }
  }
}
function same(left: unknown, right: unknown): boolean {
  if (left === right) return true
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object' || Array.isArray(left) !== Array.isArray(right)) return false
  const leftRecord = left as Record<string, unknown>, rightRecord = right as Record<string, unknown>, keys = Object.keys(leftRecord)
  return keys.length === Object.keys(rightRecord).length && keys.every(key => Object.hasOwn(rightRecord, key) && same(leftRecord[key], rightRecord[key]))
}
function visibleEdge(edge: unknown): boolean {
  if (!edge) return false
  if (typeof edge === 'string') return !!edge && edge.split(':')[0] !== 'none'
  return typeof edge === 'object' && (!('style' in edge) || edge.style !== 'none')
}

/**
 * Paint the union of all ranges. Merges are logical cells, cardinal sides are physical segments,
 * and diagonals belong to their main cell. Border commands include filtered/manual-hidden rows,
 * matching the fixed border command's direct range iterator rather than the font iterator.
 * The source is never mutated; permissions can check the complete plan before applying any patch.
 */
export function createSheetBorderPlan(
  source: SheetBorderSource, inputRanges: readonly CellRange[], preset: SheetBorderPreset,
  inputOptions: SheetBorderOptions = { style: 'thin', color: '#000000' },
): SheetBorderPlan {
  assertSheetBorderRanges(inputRanges)
  const type = normalizeSheetBorderPreset(preset)
  const options: SheetBorderOptions = { style: normalizeSheetBorderStyle(inputOptions.style), color: normalizeSheetBorderColor(inputOptions.color) }
  const merges = source.mergedRanges ?? []
  assertSheetBorderRanges(merges)
  const ranges = inputRanges.map(range => ({ ...range }))
  // A partial API range selects the complete logical merged cell without adding its bounding gaps.
  for (let added = true; added;) {
    added = false
    for (const merge of merges) if (ranges.some(range => intersects(range, merge))
      && !ranges.some(range => contains(range, merge.startRow, merge.startCol) && contains(range, merge.endRow, merge.endCol))) {
      ranges.push({ ...merge }); added = true
    }
  }
  const selected = (row: number, col: number) => ranges.some(range => contains(range, row, col))
  const mergedAt = (row: number, col: number) => merges.find(range => contains(range, row, col))
  const pending = new Map<string, SheetBorderPatch>()
  const read: SheetBorderSource['readBorder'] = (row, col) => pending.get(`${row}:${col}`)?.border ?? source.readBorder(row, col)
  const draft = (row: number, col: number): Record<string, unknown> => {
    const key = `${row}:${col}`, existing = pending.get(key)
    if (existing) return existing.border
    const border = { ...source.readBorder(row, col) }
    pending.set(key, { row, col, border })
    return border
  }
  const set = (row: number, col: number, key: string, value: unknown) => {
    if (row < 0 || col < 0 || row > maxRow || col > maxCol) return
    const current = read(row, col)
    if (value === undefined) { if (current && Object.hasOwn(current, key)) delete draft(row, col)[key] }
    else if (!current || !Object.hasOwn(current, key) || !same(current[key], value)) draft(row, col)[key] = value
  }
  const clearSegment = (row: number, col: number, side: SheetBorderSide) => {
    if (row < 0 || col < 0 || row > maxRow || col > maxCol) return
    const merge = mergedAt(row, col), actual = read(row, col)
    const edge = actual && Object.hasOwn(actual, side) ? actual[side] : merge ? read(merge.startRow, merge.startCol)?.[side] : undefined
    if (!visibleEdge(edge)) return
    if (!merge) { set(row, col, side, undefined); return }
    if (row === merge.startRow && col === merge.startCol) {
      // Replacing the primary segment must not change untouched segments that inherited its edge.
      for (const position of sidePositions(merge, side)) if (position.row !== row || position.col !== col) {
        const border = read(position.row, position.col)
        if (!border || !Object.hasOwn(border, side)) set(position.row, position.col, side, edge)
      }
    }
    set(row, col, side, null)
  }
  const clearOutside = (row: number, col: number, side: SheetBorderSide) => {
    const offset = sideOffsets[side], nextRow = row + offset.row, nextCol = col + offset.col
    if (!selected(nextRow, nextCol)) clearSegment(nextRow, nextCol, offset.opposite)
  }
  if (type === 'none') {
    // Clear only existing cells, even for a full row/column/sheet selection.
    for (const { row, col } of source.existingCells()) if (selected(row, col)) {
      for (const key of [...SHEET_BORDER_SIDES, ...diagonalKeys]) set(row, col, key, undefined)
    }
    const clearNeighbor = (row: number, col: number, side: SheetBorderSide) => {
      if (selected(row, col)) return
      const offset = sideOffsets[side]
      if (selected(row + offset.row, col + offset.col)) clearSegment(row, col, side)
    }
    for (const { row, col } of source.existingCells()) for (const side of SHEET_BORDER_SIDES) clearNeighbor(row, col, side)
    // An imported merge may store only a main-cell edge, so its other physical segments are virtual.
    for (const merge of merges) if (!selected(merge.startRow, merge.startCol)) for (const side of SHEET_BORDER_SIDES) {
      const offset = sideOffsets[side]
      for (const range of ranges) for (const { row, col } of sidePositions(merge, side, {
        startRow: range.startRow - offset.row, endRow: range.endRow - offset.row,
        startCol: range.startCol - offset.col, endCol: range.endCol - offset.col,
      })) clearNeighbor(row, col, side)
    }
  } else {
    const visited = new Set<number>(), visitedDiagonals = new Set<number>()
    for (const range of ranges) for (let row = range.startRow; row <= range.endRow; row += 1) for (let col = range.startCol; col <= range.endCol; col += 1) {
      const index = row * (maxCol + 1) + col
      if (visited.has(index)) continue
      visited.add(index)
      const merge = mergedAt(row, col)
      const keys = diagonalPresetKeys[type]
      if (keys) {
        const primaryRow = merge?.startRow ?? row, primaryCol = merge?.startCol ?? col, primary = primaryRow * (maxCol + 1) + primaryCol
        if (!visitedDiagonals.has(primary)) {
          visitedDiagonals.add(primary)
          for (const key of keys) set(primaryRow, primaryCol, key, options.style === 'none' ? undefined : { ...options })
        }
        continue
      }
      for (const side of SHEET_BORDER_SIDES) {
        const offset = sideOffsets[side], nextRow = row + offset.row, nextCol = col + offset.col
        if (merge && contains(merge, nextRow, nextCol)) continue
        const inside = selected(nextRow, nextCol)
        const apply = type === 'all' || type === 'outside' && !inside || type === 'inside' && inside
          || type === side && !inside || type === 'horizontal' && inside && (side === 'top' || side === 'bottom')
          || type === 'vertical' && inside && (side === 'left' || side === 'right')
        if (!apply) continue
        if (options.style === 'none') clearSegment(row, col, side)
        else set(row, col, side, { ...options })
        if (!inside) clearOutside(row, col, side)
      }
    }
    // Keep uniform main-cell edges available to existing snapshots/Office readers as well as segments.
    for (const merge of merges) if (selected(merge.startRow, merge.startCol)) for (const side of SHEET_BORDER_SIDES) {
      const positions = [...sidePositions(merge, side)]
      if (!positions.some(({ row, col }) => pending.get(`${row}:${col}`)?.border && Object.hasOwn(pending.get(`${row}:${col}`)!.border, side))) continue
      const values = positions.map(({ row, col }) => { const border = read(row, col); return border && Object.hasOwn(border, side) ? border[side] : read(merge.startRow, merge.startCol)?.[side] })
      if (values.every(value => same(value, values[0]))) set(merge.startRow, merge.startCol, side, visibleEdge(values[0]) ? values[0] : undefined)
    }
  }
  const patches = [...pending.values()].filter(patch => !same(source.readBorder(patch.row, patch.col) ?? {}, patch.border))
  return { ranges, patches }
}
