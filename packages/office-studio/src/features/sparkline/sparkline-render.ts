import type { SparklineSpec } from '../../domain/model.ts'
import type { ExcelCellView, ExcelSheetView } from '../../services/excelViewAdapter.ts'
import type { DrawCommand, RenderRect } from '../../rendering/contracts.ts'

interface SparklineRange { startRow: number; endRow: number; startColumn: number; endColumn: number }
interface SparklinePoint { x: number; y: number; value: number; sourceIndex: number }

export function parseSparklineRange(sourceRange: string): SparklineRange | undefined {
  const parts = sourceRange.replaceAll('$', '').trim().split(':')
  if (parts.length < 1 || parts.length > 2) return undefined
  const first = parseAddress(parts[0] ?? '')
  const last = parseAddress(parts[1] ?? parts[0] ?? '')
  if (!first || !last) return undefined
  return {
    startRow: Math.min(first.row, last.row), endRow: Math.max(first.row, last.row),
    startColumn: Math.min(first.column, last.column), endColumn: Math.max(first.column, last.column),
  }
}

export function readSparklineValues(spec: SparklineSpec, sheet: ExcelSheetView): Array<number | null> {
  const range = parseSparklineRange(spec.sourceRange)
  if (!range) return []
  const values: Array<number | null> = []
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const cell = sheet.cells[cellAddress(row, column)]
      values.push(readNumber(cell))
    }
  }
  return values
}

export function createSparklineCommands(spec: SparklineSpec, input: readonly (number | null)[], bounds: RenderRect): DrawCommand[] {
  const inset = Math.min(4, bounds.width / 8, bounds.height / 8)
  const rect = { x: bounds.x + inset, y: bounds.y + inset, width: Math.max(0, bounds.width - inset * 2), height: Math.max(0, bounds.height - inset * 2) }
  if (rect.width < 2 || rect.height < 2 || !input.length) return []
  const values = spec.emptyCells === 'zero' ? input.map(value => value ?? 0) : [...input]
  const numeric = values.filter((value): value is number => value !== null && Number.isFinite(value))
  if (!numeric.length) return []
  const transformed = spec.type === 'winloss' ? values.map(value => value === null ? null : Math.sign(value)) : values
  const transformedNumeric = transformed.filter((value): value is number => value !== null)
  const min = spec.axis === 'custom' && Number.isFinite(spec.min) ? spec.min! : Math.min(...transformedNumeric)
  const max = spec.axis === 'custom' && Number.isFinite(spec.max) ? spec.max! : Math.max(...transformedNumeric)
  if (min > max) return []
  const span = max - min || 1
  const yFor = (value: number) => rect.y + rect.height - ((value - min) / span) * rect.height

  if (spec.type === 'line') return lineCommands(spec, transformed, rect, yFor)
  return columnCommands(spec, transformed, rect, min, max, yFor)
}

export function sparklineCommandsForCell(sheet: ExcelSheetView, cell: ExcelCellView | undefined, bounds: RenderRect): DrawCommand[] {
  return cell?.sparkline ? createSparklineCommands(cell.sparkline, readSparklineValues(cell.sparkline, sheet), bounds) : []
}

function lineCommands(spec: SparklineSpec, values: readonly (number | null)[], rect: RenderRect, yFor: (value: number) => number): DrawCommand[] {
  const points: SparklinePoint[] = []
  values.forEach((value, sourceIndex) => {
    if (value === null || !Number.isFinite(value)) return
    const x = rect.x + (values.length === 1 ? rect.width / 2 : sourceIndex / (values.length - 1) * rect.width)
    points.push({ x, y: yFor(value), value, sourceIndex })
  })
  const commands: DrawCommand[] = []
  for (let index = 1; index < points.length; index += 1) {
    const previous = points[index - 1]!
    const current = points[index]!
    if (spec.emptyCells === 'gap' && current.sourceIndex !== previous.sourceIndex + 1) continue
    commands.push({ kind: 'line', x1: previous.x, y1: previous.y, x2: current.x, y2: current.y, color: spec.color, width: 1.5, clip: rect })
  }
  const extrema = points.reduce<{ high: SparklinePoint; low: SparklinePoint }>((result, point) => ({
    high: point.value > result.high.value ? point : result.high,
    low: point.value < result.low.value ? point : result.low,
  }), { high: points[0]!, low: points[0]! })
  if (spec.highPointColor) commands.push(marker(extrema.high, spec.highPointColor, rect))
  if (spec.lowPointColor) commands.push(marker(extrema.low, spec.lowPointColor, rect))
  return commands
}

function columnCommands(spec: SparklineSpec, values: readonly (number | null)[], rect: RenderRect, min: number, max: number, yFor: (value: number) => number): DrawCommand[] {
  const zero = min <= 0 && max >= 0 ? yFor(0) : min >= 0 ? rect.y + rect.height : rect.y
  const step = rect.width / values.length
  const width = Math.max(1, step * 0.66)
  const finite = values.filter((value): value is number => value !== null && Number.isFinite(value))
  const high = Math.max(...finite)
  const low = Math.min(...finite)
  const commands: DrawCommand[] = []
  values.forEach((value, index) => {
    if (value === null || !Number.isFinite(value) || value === 0) return
    const y = yFor(value)
    const top = Math.min(zero, y)
    const height = Math.max(1, Math.abs(zero - y))
    const color = value === high && spec.highPointColor ? spec.highPointColor
      : value === low && spec.lowPointColor ? spec.lowPointColor
        : value < 0 ? spec.negativeColor ?? spec.color : spec.color
    commands.push({ kind: 'rect', x: rect.x + index * step + (step - width) / 2, y: top, width, height, fill: color, clip: rect })
  })
  return commands
}

function marker(point: SparklinePoint, color: string, clip: RenderRect): DrawCommand {
  return { kind: 'rect', x: point.x - 2, y: point.y - 2, width: 4, height: 4, fill: color, clip }
}

function readNumber(cell: ExcelCellView | undefined): number | null {
  const value = cell?.rawValue ?? cell?.cachedValue ?? cell?.value
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value !== 'string' || !value.trim()) return null
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : null
}

function parseAddress(address: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(address.trim())
  if (!match) return undefined
  let column = 0
  for (const character of match[1]!.toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64
  const row = Number(match[2]) - 1
  column -= 1
  if (row > 1_048_575 || column > 16_383) return undefined
  return { row, column }
}

function cellAddress(row: number, column: number): string {
  let label = ''
  for (let index = column; index >= 0; index = Math.floor(index / 26) - 1) label = String.fromCharCode(index % 26 + 65) + label
  return `${label}${row + 1}`
}
