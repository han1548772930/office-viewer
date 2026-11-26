import type { SparklineSpec } from '../domain/model.ts'
import { createSparklineSpec, SPARKLINE_POINT_ORDER, type SparklineSettingsInput } from '../features/sparkline/sparkline-model.ts'
import { sparklineCellAddress } from '../features/sparkline/sparkline-reference.ts'
import { normalizeOfficeSparklineSource } from '../features/sparkline/sparkline-office-source.ts'

function object(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`迷你图${label}数据无效。`)
  return value as Record<string, unknown>
}
function boolean(value: unknown): boolean | undefined {
  if (value === undefined) return undefined
  if (typeof value !== 'boolean') throw new TypeError('迷你图开关设置必须是布尔值。')
  return value
}
function number(value: unknown): number | undefined {
  if (value === undefined) return undefined
  if (typeof value !== 'number' || !Number.isFinite(value)) throw new TypeError('迷你图数值设置必须是有限数值。')
  return value
}
function string(value: unknown): string | undefined {
  if (value === undefined) return undefined
  if (typeof value !== 'string') throw new TypeError('迷你图文本设置无效。')
  return value
}
function emptyMode(value: unknown): SparklineSpec['emptyCells'] | undefined {
  if (value === undefined || value === 'gap' || value === 'zero' || value === 'connect') return value
  throw new TypeError('迷你图空白或非数字处理方式无效。')
}
function extremum(value: unknown): SparklineSpec['extremumMin'] {
  if (value === undefined) return undefined
  const raw = object(value, '极值'), type = raw.type
  if (type !== 'individual' && type !== 'group' && type !== 'custom') throw new TypeError('迷你图极值类型无效。')
  const result = number(raw.value)
  return { type, ...(result === undefined ? {} : { value: result }) }
}

/** Parser DTO validation uses the same settings/reference authority as authoring. */
export function normalizeWorkbookSparkline(value: unknown, row: number, column: number): SparklineSpec | undefined {
  if (value === undefined || value === null) return undefined
  const raw = object(value, '单元格'), type = raw.type
  if (type !== 'line' && type !== 'column' && type !== 'winloss' && type !== 'pie') throw new TypeError('迷你图类型无效。')
  const sourceRange = string(raw.sourceRange)
  if (!sourceRange) throw new TypeError('迷你图缺少源范围。')
  const axis = raw.axis === undefined ? 'individual' : raw.axis
  if (axis !== 'individual' && axis !== 'custom') throw new TypeError('迷你图坐标轴设置无效。')
  const settings: SparklineSettingsInput = { type, axis, emptyCells: emptyMode(raw.emptyCells) ?? 'gap' }
  for (const key of ['color', 'negativeColor', 'highPointColor', 'lowPointColor'] as const) if (raw[key] !== undefined) settings[key] = string(raw[key])
  for (const key of ['min', 'max', 'lineWidth'] as const) if (raw[key] !== undefined) settings[key] = number(raw[key])
  for (const key of ['containHiddenCells', 'showGradient', 'showRadius'] as const) if (raw[key] !== undefined) settings[key] = boolean(raw[key])
  if (raw.themeType !== undefined) {
    const theme = raw.themeType
    if (theme !== 'standard' && theme !== 'specialty' && theme !== 'distinctive' && theme !== 'soft' && theme !== 'fresh' && theme !== 'custom') throw new TypeError('迷你图主题无效。')
    settings.themeType = theme
  }
  if (raw.nonNumericCells !== undefined) settings.nonNumericCells = emptyMode(raw.nonNumericCells)
  for (const key of ['extremumMin', 'extremumMax'] as const) if (raw[key] !== undefined) settings[key] = extremum(raw[key])
  if (raw.axisOptions !== undefined) {
    const axis = object(raw.axisOptions, '轴配置')
    settings.axisOptions = { visible: boolean(axis.visible), reverse: boolean(axis.reverse), color: string(axis.color) }
  }
  if (raw.points !== undefined) {
    const points = object(raw.points, '标记'), result: NonNullable<SparklineSpec['points']> = {}
    for (const name of SPARKLINE_POINT_ORDER) {
      if (points[name] === undefined) continue
      const point = object(points[name], '标记')
      result[name] = { visible: boolean(point.visible), color: string(point.color) }
    }
    settings.points = result
  }
  return createSparklineSpec({ ...settings, type, sourceRange, color: settings.color ?? '#364ef4', axis, emptyCells: settings.emptyCells ?? 'gap',
    targetRange: sparklineCellAddress(row, column), ...(raw.groupId === undefined ? {} : { groupId: string(raw.groupId) }),
    ...(raw.officeSource === undefined ? {} : { officeSource: normalizeOfficeSparklineSource(raw.officeSource) }) })
}
