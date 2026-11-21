import type { SparklineSpec, SparklineType, SparklineEmptyCells } from '../../domain/model.ts'
import { parseSparklineRange } from './sparkline-render.ts'

export interface SparklineGroupRequest {
  targetRange: string
  sourceRange: string
  type: SparklineType
  color: string
  negativeColor?: string
  highPointColor?: string
  lowPointColor?: string
  emptyCells?: SparklineEmptyCells
  axis?: 'individual' | 'custom'
  min?: number
  max?: number
}

export function createSparklineSpec(request: SparklineGroupRequest): SparklineSpec {
  const target = parseSparklineRange(request.targetRange)
  const source = parseSparklineRange(request.sourceRange)
  if (!target || !source) throw new Error('迷你图源区域和目标区域必须是有效单元格范围。')
  if (request.type !== 'line' && request.type !== 'column' && request.type !== 'winloss') throw new Error('不支持的迷你图类型。')
  if (!validColor(request.color) || !optionalColor(request.negativeColor) || !optionalColor(request.highPointColor) || !optionalColor(request.lowPointColor)) {
    throw new Error('迷你图颜色格式无效。')
  }
  const axis = request.axis ?? 'individual'
  if (axis === 'custom' && (!Number.isFinite(request.min) || !Number.isFinite(request.max) || request.min! >= request.max!)) {
    throw new Error('自定义迷你图坐标轴必须提供递增的最小值和最大值。')
  }
  const emptyCells = request.emptyCells ?? 'gap'
  if (emptyCells !== 'gap' && emptyCells !== 'zero' && emptyCells !== 'connect') throw new Error('不支持的空单元格处理方式。')
  return {
    type: request.type,
    sourceRange: normalizeRange(request.sourceRange),
    color: request.color,
    ...(request.negativeColor ? { negativeColor: request.negativeColor } : {}),
    ...(request.highPointColor ? { highPointColor: request.highPointColor } : {}),
    ...(request.lowPointColor ? { lowPointColor: request.lowPointColor } : {}),
    emptyCells,
    axis,
    ...(axis === 'custom' ? { min: request.min, max: request.max } : {}),
  }
}

export function createSparklineGroupPatches(request: SparklineGroupRequest): Record<string, SparklineSpec> {
  const target = parseSparklineRange(request.targetRange)
  if (!target) throw new Error('迷你图目标区域必须是有效单元格范围。')
  const source = parseSparklineRange(request.sourceRange)
  if (!source) throw new Error('迷你图源区域必须是有效单元格范围。')
  const rowCount = target.endRow - target.startRow + 1
  const columnCount = target.endColumn - target.startColumn + 1
  const sourceRows = source.endRow - source.startRow + 1
  const sourceColumns = source.endColumn - source.startColumn + 1
  if (rowCount !== 1 && columnCount !== 1) throw new Error('迷你图目标区域必须为单行或单列。')
  if (rowCount === 1 && columnCount !== sourceColumns || columnCount === 1 && rowCount !== sourceRows) {
    throw new Error('迷你图目标数必须与源数据系列数一致。')
  }
  const spec = createSparklineSpec(request)
  const patches: Record<string, SparklineSpec> = {}
  for (let offset = 0; offset < Math.max(rowCount, columnCount); offset += 1) {
    const row = target.startRow + (rowCount === 1 ? 0 : offset)
    const column = target.startColumn + (columnCount === 1 ? 0 : offset)
    const seriesRange = rowCount === 1
      ? `${cellAddress(source.startRow, source.startColumn + offset)}:${cellAddress(source.endRow, source.startColumn + offset)}`
      : `${cellAddress(source.startRow + offset, source.startColumn)}:${cellAddress(source.startRow + offset, source.endColumn)}`
    patches[cellAddress(row, column)] = { ...spec, sourceRange: seriesRange }
  }
  return patches
}

export function translateSparklineForMove(spec: SparklineSpec, rowDelta: number, columnDelta: number): SparklineSpec {
  const source = spec.sourceRange.replaceAll('$', '').split(':')
  const shifted = source.map(address => {
    const range = parseSparklineRange(address)
    if (!range || range.startRow !== range.endRow || range.startColumn !== range.endColumn) throw new Error('迷你图源区域格式无效。')
    const row = range.startRow + rowDelta
    const column = range.startColumn + columnDelta
    if (row < 0 || column < 0) throw new Error('移动迷你图会使源区域超出工作表边界。')
    return `${columnLabel(column)}${row + 1}`
  })
  return { ...spec, sourceRange: shifted.join(':') }
}

function normalizeRange(value: string): string {
  const parts = value.replaceAll('$', '').trim().toUpperCase().split(':')
  return parts.length === 1 ? parts[0]! : `${parts[0]}:${parts[1]}`
}
function validColor(value: string): boolean { return /^#[0-9a-f]{6}$/i.test(value) }
function optionalColor(value: string | undefined): boolean { return value === undefined || validColor(value) }
function cellAddress(row: number, column: number): string { return `${columnLabel(column)}${row + 1}` }
function columnLabel(index: number): string {
  let label = ''
  for (let current = index; current >= 0; current = Math.floor(current / 26) - 1) label = String.fromCharCode(current % 26 + 65) + label
  return label
}
