import type { CellRange, SparklineSpec } from '../../domain/model.ts'
import { normalizeSheetCellValueType } from '../../domain/sheet-cell-value.ts'
import { sheetAxisSpanAt } from '../../domain/sheet-axis.ts'
import type { ExcelCellView, ExcelSheetView } from '../../services/excelViewAdapter.ts'
import type { DrawCommand, RenderRect } from '../../rendering/contracts.ts'
import { resolveSparklineSettings, type ResolvedSparklineSettings, type SparklinePointName } from './sparkline-model.ts'
import { parseSparklineReference, sparklineCellAddress } from './sparkline-reference.ts'

interface SparklineRange { startRow: number; endRow: number; startColumn: number; endColumn: number }
interface SparklinePoint { x: number; y: number; value: number; sourceIndex: number }
export type SparklineSample = number | 'gap' | 'connect'
export interface SparklineExtrema { min: number; max: number }
export interface SparklineRenderContext {
  sheets?: readonly ExcelSheetView[]
  /** Frame-scoped caches: rebuilding the scene or changing source permissions requires a fresh context. */
  groupCache?: Map<string, SparklineExtrema>
  sampleCache?: Map<SparklineSpec, SparklineSample[]>
  canReadSource?: (sheet: ExcelSheetView, range: CellRange) => boolean
}

export function parseSparklineRange(sourceRange: string): SparklineRange | undefined {
  const reference = parseSparklineReference(sourceRange)
  return reference ? { startRow: reference.range.startRow, endRow: reference.range.endRow,
    startColumn: reference.range.startCol, endColumn: reference.range.endCol } : undefined
}

function missing(mode: 'gap' | 'zero' | 'connect'): SparklineSample { return mode === 'zero' ? 0 : mode }
function readSample(cell: ExcelCellView | undefined, settings: ResolvedSparklineSettings): SparklineSample {
  if (!cell) return missing(settings.emptyCells)
  const type = normalizeSheetCellValueType(cell.formula ? cell.cachedDataType : cell.dataType)
  const cached = cell.formula && (cell.rawValue === undefined || cell.rawValue === null) ? cell.cachedValue : undefined
  const value = cached !== undefined ? cached : cell.rawValue !== undefined ? cell.rawValue : cell.value === '' ? null : cell.value
  if (value === null || value === undefined || type === 'blank') return missing(settings.emptyCells)
  if (type === 'text' || type === 'error') return missing(settings.nonNumericCells)
  if (typeof value === 'boolean') return value ? 1 : 0
  if (type === 'boolean') return /^(true|1)$/i.test(String(value)) ? 1 : /^(false|0)$/i.test(String(value)) ? 0 : missing(settings.nonNumericCells)
  if (typeof value === 'number') return Number.isFinite(value) ? value : missing(settings.nonNumericCells)
  if (type === 'number' && typeof value === 'string' && value.trim()) {
    const number = Number(value)
    return Number.isFinite(number) ? number : missing(settings.nonNumericCells)
  }
  return missing(settings.nonNumericCells)
}
function hiddenSource(sheet: ExcelSheetView, row: number, column: number): boolean {
  return sheet.hiddenRows?.includes(row) === true || sheet.outlineHiddenRows?.includes(row) === true || sheet.filteredRows?.includes(row) === true
    || sheet.hiddenColumns?.includes(column) === true || sheet.outlineHiddenColumns?.includes(column) === true
    || sheetAxisSpanAt(sheet.rowHiddenSpans, row)?.value === true || sheetAxisSpanAt(sheet.columnHiddenSpans, column)?.value === true
    || sheetAxisSpanAt(sheet.outlineHiddenRowSpans, row)?.value === true || sheetAxisSpanAt(sheet.outlineHiddenColumnSpans, column)?.value === true
}

export function readSparklineSamples(spec: SparklineSpec, sheet: ExcelSheetView, context: SparklineRenderContext = {}): SparklineSample[] {
  const reference = parseSparklineReference(spec.sourceRange)
  if (!reference) return []
  const source = reference.sheetName === undefined || reference.sheetName === sheet.id || reference.sheetName.toLocaleLowerCase() === sheet.name.toLocaleLowerCase()
    ? sheet : context.sheets?.find(candidate => candidate.id === reference.sheetName || candidate.name.toLocaleLowerCase() === reference.sheetName!.toLocaleLowerCase())
  if (!source || context.canReadSource?.(source, reference.range) === false) return []
  const cached = context.sampleCache?.get(spec)
  if (cached) return cached
  const settings = resolveSparklineSettings(spec), values: SparklineSample[] = []
  for (let row = reference.range.startRow; row <= reference.range.endRow; row += 1) {
    for (let column = reference.range.startCol; column <= reference.range.endCol; column += 1) {
      if (!settings.containHiddenCells && hiddenSource(source, row, column)) continue
      values.push(readSample(source.cells[sparklineCellAddress(row, column)], settings))
    }
  }
  context.sampleCache?.set(spec, values)
  return values
}

/** Compatibility read exposes missing points as null; the renderer keeps gap and connect distinct. */
export function readSparklineValues(spec: SparklineSpec, sheet: ExcelSheetView, context: SparklineRenderContext = {}): Array<number | null> {
  return readSparklineSamples(spec, sheet, context).map(value => typeof value === 'number' ? value : null)
}
function extrema(values: readonly SparklineSample[], winloss: boolean): SparklineExtrema | undefined {
  let min = Infinity, max = -Infinity
  for (const value of values) if (typeof value === 'number' && Number.isFinite(value)) {
    const point = winloss ? Math.sign(value) : value
    min = Math.min(min, point); max = Math.max(max, point)
  }
  return min === Infinity ? undefined : { min, max }
}
function groupExtrema(spec: SparklineSpec, sheet: ExcelSheetView, context: SparklineRenderContext): SparklineExtrema | undefined {
  if (!spec.groupId) return undefined
  const members = Object.values(sheet.cells).flatMap(cell => cell.sparkline?.groupId === spec.groupId ? [cell.sparkline] : [])
  const permissionMask = members.map(member => {
    const reference = parseSparklineReference(member.sourceRange)
    if (!reference) return '0'
    const source = reference.sheetName === undefined || reference.sheetName === sheet.id || reference.sheetName.toLocaleLowerCase() === sheet.name.toLocaleLowerCase()
      ? sheet : context.sheets?.find(candidate => candidate.id === reference.sheetName || candidate.name.toLocaleLowerCase() === reference.sheetName!.toLocaleLowerCase())
    return source && context.canReadSource?.(source, reference.range) !== false ? '1' : '0'
  }).join('')
  const key = `${sheet.id}:${spec.groupId}:${permissionMask}`, cached = context.groupCache?.get(key)
  if (cached) return cached
  let min = Infinity, max = -Infinity
  for (const member of members) {
    const range = extrema(readSparklineSamples(member, sheet, context), spec.type === 'winloss')
    if (range) { min = Math.min(min, range.min); max = Math.max(max, range.max) }
  }
  if (min === Infinity) return undefined
  const result = { min, max }
  context.groupCache?.set(key, result)
  return result
}

export function createSparklineCommands(spec: SparklineSpec, input: readonly (SparklineSample | null)[], bounds: RenderRect,
  options: { groupExtrema?: SparklineExtrema } = {}): DrawCommand[] {
  if (spec.axis === 'custom' && spec.extremumMin === undefined && spec.extremumMax === undefined && Number.isFinite(spec.min) && Number.isFinite(spec.max) && spec.min! >= spec.max!) return []
  const settings = resolveSparklineSettings(spec), inset = Math.min(4, bounds.width / 8, bounds.height / 8)
  const rect = { x: bounds.x + inset, y: bounds.y + inset, width: Math.max(0, bounds.width - inset * 2), height: Math.max(0, bounds.height - inset * 2) }
  if (rect.width < 2 || rect.height < 2 || !input.length) return []
  const values: SparklineSample[] = input.map(value => value === null ? missing(settings.emptyCells) : typeof value === 'number' && !Number.isFinite(value) ? missing(settings.nonNumericCells) : value)
  if (settings.type === 'pie') return pieCommands(values, rect)
  const transformed = settings.type === 'winloss' ? values.map(value => typeof value === 'number' ? Math.sign(value) : value) : values
  const own = extrema(transformed, false)
  if (!own) return []
  const extent = (end: 'Min' | 'Max') => {
    const axis = settings[`extremum${end}`], property = end === 'Min' ? 'min' : 'max'
    return axis.type === 'custom' ? axis.value! : axis.type === 'group' ? options.groupExtrema?.[property] ?? own[property] : own[property]
  }
  let min = extent('Min'), max = extent('Max')
  if (settings.type !== 'line') {
    if (settings.extremumMin.type !== 'custom') min = Math.min(0, min)
    if (settings.extremumMax.type !== 'custom') max = Math.max(0, max)
  }
  if (settings.type === 'winloss' && settings.extremumMin.type !== 'custom' && settings.extremumMax.type !== 'custom') { min = -1; max = 1 }
  if (!Number.isFinite(min) || !Number.isFinite(max) || min > max) return []
  const span = max - min, yFor = (value: number) => {
    if (span === 0) return rect.y + rect.height / 2
    const ratio = Number.isFinite(span) ? (value - min) / span : (value / 2 - min / 2) / (max / 2 - min / 2)
    const finiteRatio = Number.isFinite(ratio) ? Math.max(-1_000_000, Math.min(1_000_000, ratio)) : value > max ? 1_000_000 : -1_000_000
    return rect.y + rect.height - finiteRatio * rect.height
  }
  const commands = settings.type === 'line' ? lineCommands(settings, transformed, rect, yFor) : columnCommands(settings, transformed, rect, min, max, yFor)
  if (settings.axisOptions.visible) {
    const zero = Math.max(rect.y, Math.min(rect.y + rect.height, yFor(0)))
    commands.push({ kind: 'line', x1: rect.x, y1: zero, x2: rect.x + rect.width, y2: zero, color: settings.axisOptions.color, width: 1, clip: rect })
  }
  return commands
}

export function sparklineCommandsForCell(sheet: ExcelSheetView, cell: ExcelCellView | undefined, bounds: RenderRect, context: SparklineRenderContext = {}): DrawCommand[] {
  if (!cell?.sparkline) return []
  const spec = cell.sparkline, samples = readSparklineSamples(spec, sheet, context), settings = resolveSparklineSettings(spec)
  const group = settings.extremumMin.type === 'group' || settings.extremumMax.type === 'group' ? groupExtrema(spec, sheet, context) : undefined
  return createSparklineCommands(spec, samples, bounds, { groupExtrema: group })
}

function lineCommands(settings: ResolvedSparklineSettings, values: readonly SparklineSample[], rect: RenderRect, yFor: (value: number) => number): DrawCommand[] {
  const points: SparklinePoint[] = []
  values.forEach((value, sourceIndex) => {
    if (typeof value !== 'number') return
    const ratio = values.length === 1 ? .5 : sourceIndex / (values.length - 1)
    points.push({ x: rect.x + (settings.axisOptions.reverse ? 1 - ratio : ratio) * rect.width, y: yFor(value), value, sourceIndex })
  })
  if (!points.length) return []
  const strokes: DrawCommand[] = [], fills: DrawCommand[] = []
  for (let index = 1; index < points.length; index += 1) {
    const previous = points[index - 1]!, current = points[index]!
    let disconnected = false
    for (let offset = previous.sourceIndex + 1; offset < current.sourceIndex; offset += 1) if (values[offset] === 'gap') { disconnected = true; break }
    if (disconnected) continue
    if (settings.showGradient) fills.push(...gradientSegment(previous, current, rect, settings.color))
    strokes.push({ kind: 'line', x1: previous.x, y1: previous.y, x2: current.x, y2: current.y, color: settings.color, width: settings.lineWidth, clip: rect })
  }
  const commands = [...fills, ...strokes], range = extrema(values, false)!
  const markerNames: SparklinePointName[] = ['markersPoint', 'firstPoint', 'lastPoint', 'highPoint', 'lowPoint', 'negativePoint']
  for (const name of markerNames) {
    const pointSetting = settings.points[name]
    if (!pointSetting.visible) continue
    for (const point of points) {
      if (name === 'firstPoint' && point !== points[0] || name === 'lastPoint' && point !== points[points.length - 1]
        || name === 'highPoint' && point.value !== range.max || name === 'lowPoint' && point.value !== range.min || name === 'negativePoint' && point.value >= 0) continue
      commands.push({ kind: 'polygon', points: circlePoints(point.x, point.y, 2, 12), fill: pointSetting.color, clip: rect })
    }
  }
  return commands
}

function columnCommands(settings: ResolvedSparklineSettings, values: readonly SparklineSample[], rect: RenderRect, min: number, max: number,
  yFor: (value: number) => number): DrawCommand[] {
  const zero = min <= 0 && max >= 0 ? yFor(0) : min >= 0 ? rect.y + rect.height : rect.y
  const step = rect.width / values.length, width = Math.min(step, Math.max(.5, step * .66)), range = extrema(values, false)!
  const first = values.findIndex(value => typeof value === 'number'), commands: DrawCommand[] = []
  let last = -1
  for (let index = values.length - 1; index >= 0; index -= 1) if (typeof values[index] === 'number') { last = index; break }
  values.forEach((value, index) => {
    if (typeof value !== 'number' || value === 0) return
    const y = yFor(value), top = Math.min(zero, y), height = Math.max(.5, Math.abs(zero - y)), position = settings.axisOptions.reverse ? values.length - index - 1 : index
    let color = settings.color
    for (const name of ['firstPoint', 'lastPoint', 'highPoint', 'lowPoint', 'negativePoint'] as const) {
      if (!settings.points[name].visible) continue
      if (name === 'firstPoint' && index === first || name === 'lastPoint' && index === last || name === 'highPoint' && value === range.max
        || name === 'lowPoint' && value === range.min || name === 'negativePoint' && value < 0) color = settings.points[name].color
    }
    const bar = { x: rect.x + position * step + (step - width) / 2, y: top, width, height }
    commands.push(settings.showRadius ? { kind: 'polygon', points: roundedRectPoints(bar, Math.min(2, width / 2, height / 2)), fill: color, clip: rect }
      : { kind: 'rect', ...bar, fill: color, clip: rect })
  })
  return commands
}
function circlePoints(x: number, y: number, radius: number, count: number): Array<{ x: number; y: number }> {
  return Array.from({ length: count }, (_, index) => { const angle = index / count * Math.PI * 2; return { x: x + Math.cos(angle) * radius, y: y + Math.sin(angle) * radius } })
}
function roundedRectPoints(rect: RenderRect, radius: number): Array<{ x: number; y: number }> {
  const corners = [
    { x: rect.x + rect.width - radius, y: rect.y + radius, angle: -Math.PI / 2 },
    { x: rect.x + rect.width - radius, y: rect.y + rect.height - radius, angle: 0 },
    { x: rect.x + radius, y: rect.y + rect.height - radius, angle: Math.PI / 2 },
    { x: rect.x + radius, y: rect.y + radius, angle: Math.PI },
  ]
  return corners.flatMap(corner => Array.from({ length: 4 }, (_, index) => { const angle = corner.angle + index / 3 * Math.PI / 2; return { x: corner.x + Math.cos(angle) * radius, y: corner.y + Math.sin(angle) * radius } }))
}
function clippedY(points: readonly { x: number; y: number }[], edge: number, below: boolean): Array<{ x: number; y: number }> {
  const output: Array<{ x: number; y: number }> = []
  if (!points.length) return output
  for (let index = 0; index < points.length; index += 1) {
    const a = points[index]!, b = points[(index + 1) % points.length]!, insideA = below ? a.y >= edge : a.y <= edge, insideB = below ? b.y >= edge : b.y <= edge
    if (insideA) output.push(a)
    if (insideA !== insideB) { const ratio = (edge - a.y) / (b.y - a.y); output.push({ x: a.x + (b.x - a.x) * ratio, y: edge }) }
  }
  return output
}
function alphaColor(color: string, alpha: number): string {
  if (color === 'transparent') return color
  const hex = /^#([0-9a-f]{3,8})$/i.exec(color)
  if (hex) {
    const expanded = hex[1]!.length === 3 || hex[1]!.length === 4 ? [...hex[1]!].map(character => character + character).join('') : hex[1]!
    const ownAlpha = expanded.length === 8 ? Number.parseInt(expanded.slice(6), 16) / 255 : 1
    return `rgba(${Number.parseInt(expanded.slice(0, 2), 16)},${Number.parseInt(expanded.slice(2, 4), 16)},${Number.parseInt(expanded.slice(4, 6), 16)},${alpha * ownAlpha})`
  }
  const rgb = /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)(?:\s*,\s*([\d.]+))?\s*\)$/i.exec(color)
  return rgb ? `rgba(${rgb[1]},${rgb[2]},${rgb[3]},${alpha * Number(rgb[4] ?? 1)})` : color
}
function gradientSegment(a: SparklinePoint, b: SparklinePoint, rect: RenderRect, color: string): DrawCommand[] {
  const polygon = [{ x: a.x, y: a.y }, { x: b.x, y: b.y }, { x: b.x, y: rect.y + rect.height }, { x: a.x, y: rect.y + rect.height }], output: DrawCommand[] = []
  // Shared polygons give both backends the same area; 12 bands approximate the public .20 -> .05 alpha gradient.
  for (let band = 0; band < 12; band += 1) {
    const start = rect.y + band / 12 * rect.height, end = rect.y + (band + 1) / 12 * rect.height
    const points = clippedY(clippedY(polygon, start, true), end, false)
    if (points.length >= 3) output.push({ kind: 'polygon', points, fill: alphaColor(color, .2 - (band + .5) / 12 * .15), clip: rect })
  }
  return output
}
function pieCommands(values: readonly SparklineSample[], rect: RenderRect): DrawCommand[] {
  const slices = values.filter((value): value is number => typeof value === 'number' && value !== 0).map(Math.abs)
  const total = slices.reduce((sum, value) => sum + value, 0)
  if (!Number.isFinite(total) || total === 0) return []
  const palette = ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF', '#FF9F40', '#E7E9ED', '#B4A0E5'], radius = Math.min(rect.width, rect.height) / 2
  const center = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }, output: DrawCommand[] = []
  let angle = -Math.PI / 2
  slices.forEach((value, index) => {
    const sweep = value / total * Math.PI * 2, segments = Math.max(2, Math.ceil(sweep / (Math.PI * 2) * 64)), points = [center]
    for (let step = 0; step <= segments; step += 1) { const next = angle + sweep * step / segments; points.push({ x: center.x + Math.cos(next) * radius, y: center.y + Math.sin(next) * radius }) }
    output.push({ kind: 'polygon', points, fill: palette[index % palette.length]!, clip: rect })
    if (slices.length > 1) output.push({ kind: 'line', x1: center.x, y1: center.y, x2: points[1]!.x, y2: points[1]!.y, color: '#ffffff', width: 1, clip: rect })
    angle += sweep
  })
  return output
}
