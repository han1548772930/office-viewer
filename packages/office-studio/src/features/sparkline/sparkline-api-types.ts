import type { ApiRange } from '../../api/events.ts'
import type { SparklineEmptyCells, SparklineType } from '../../domain/model.ts'
import type { ResolvedSparklineSettings, SparklineSettingsInput } from './sparkline-model.ts'
import { SPARKLINE_POINT_ORDER } from './sparkline-model.ts'
import type { SparklineRangeMatrix } from './sparkline-range-matrix.ts'

export const SparklineTypeEnum = Object.freeze({ LINE_CHART: 1, BAR_CHART: 2, PROFIT_AND_LOSS_CHART: 3, PIE_CHART: 4 } as const)
export type SparklineTypeEnum = typeof SparklineTypeEnum[keyof typeof SparklineTypeEnum]
export const SparklineThemeTypeEnum = Object.freeze({ STANDARD: 'standard', SPECIALTY: 'specialty', DISTINCTIVE: 'distinctive', SOFT: 'soft', FRESH: 'fresh', CUSTOM: 'custom' } as const)
export type SparklineThemeTypeEnum = typeof SparklineThemeTypeEnum[keyof typeof SparklineThemeTypeEnum]
export const SpacialShowAsEnum = Object.freeze({ SPACING: 'spacing', ZERO_VALUE: 'zero_value', CONNECT_WITH_LINE: 'connect_with_line' } as const)
export type SpacialShowAsEnum = typeof SpacialShowAsEnum[keyof typeof SpacialShowAsEnum]
export const SparklineExtremumTypeEnum = Object.freeze({ SELF_EXTREMUM: 1, GROUP_EXTREMUM: 2, CUSTOM_EXTREMUM: 3 } as const)
export type SparklineExtremumTypeEnum = typeof SparklineExtremumTypeEnum[keyof typeof SparklineExtremumTypeEnum]
export interface ISparklinePointConfig { visible?: boolean; color?: string }
export interface ISparklineAxisConfig { color?: string; reverse?: boolean; visible?: boolean }
export interface ISparklineExtremumConfig { type?: SparklineExtremumTypeEnum; value?: number }
export type ISparklinePoints = Partial<Record<typeof SPARKLINE_POINT_ORDER[number], ISparklinePointConfig>>
export interface ISparklineGroupConfig {
  type?: SparklineTypeEnum
  axis?: ISparklineAxisConfig
  extremumMax?: ISparklineExtremumConfig
  extremumMin?: ISparklineExtremumConfig
  themeType?: SparklineThemeTypeEnum
  nonNumShowAs?: SpacialShowAsEnum
  emptyShowAs?: SpacialShowAsEnum
  containHiddenCells?: boolean
  seriesColor?: string
  points?: ISparklinePoints
  lineWidth?: number
  showGradient?: boolean
  showRadius?: boolean
}
export interface ISparklineGroup { config: ISparklineGroupConfig; sparklines: SparklineRangeMatrix<ApiRange> }
export interface ISparklineSerializedGroup { config: ISparklineGroupConfig; sparklines: Record<number, Record<number, ApiRange>> }

const publicTypes: Readonly<Record<SparklineTypeEnum, SparklineType>> = { 1: 'line', 2: 'column', 3: 'winloss', 4: 'pie' }
const publicEmpty: Readonly<Record<SpacialShowAsEnum, SparklineEmptyCells>> = { spacing: 'gap', zero_value: 'zero', connect_with_line: 'connect' }
const publicExtremum = { 1: 'individual', 2: 'group', 3: 'custom' } as const

export function sparklineTypeFromPublic(type: SparklineTypeEnum): SparklineType {
  const value = publicTypes[type]
  if (!value) throw new TypeError('迷你图类型枚举无效。')
  return value
}
/** Public partial configurations stay partial so unspecified group settings survive. */
export function sparklineConfigFromPublic(config: ISparklineGroupConfig): SparklineSettingsInput {
  if (!config || typeof config !== 'object' || Array.isArray(config)) throw new TypeError('迷你图组配置无效。')
  const settings: SparklineSettingsInput = {}
  if (config.type !== undefined) settings.type = sparklineTypeFromPublic(config.type)
  if (config.axis !== undefined) {
    if (!config.axis || typeof config.axis !== 'object' || Array.isArray(config.axis)) throw new TypeError('迷你图轴配置无效。')
    settings.axisOptions = { ...config.axis }
  }
  for (const key of ['extremumMin', 'extremumMax'] as const) {
    const extremum = config[key]
    if (extremum === undefined) continue
    if (!extremum || typeof extremum !== 'object' || Array.isArray(extremum)) throw new TypeError('迷你图极值配置无效。')
    const type = extremum.type === undefined ? undefined : publicExtremum[extremum.type]
    if (extremum.type !== undefined && !type) throw new TypeError('迷你图极值类型枚举无效。')
    // A value-only patch retains the current extremum mode in mergeSparklineSettings.
    settings[key] = { ...(type === undefined ? {} : { type }), ...(extremum.value === undefined ? {} : { value: extremum.value }) }
  }
  if (config.themeType !== undefined) settings.themeType = config.themeType
  if (config.seriesColor !== undefined) settings.color = config.seriesColor
  for (const [publicKey, key] of [['emptyShowAs', 'emptyCells'], ['nonNumShowAs', 'nonNumericCells']] as const) {
    const value = config[publicKey]
    if (value === undefined) continue
    const mode = publicEmpty[value]
    if (!mode) throw new TypeError('迷你图空白或非数字枚举无效。')
    settings[key] = mode
  }
  for (const key of ['containHiddenCells', 'lineWidth', 'showGradient', 'showRadius'] as const) {
    if (config[key] !== undefined) Object.assign(settings, { [key]: config[key] })
  }
  if (config.points !== undefined) {
    if (!config.points || typeof config.points !== 'object' || Array.isArray(config.points)) throw new TypeError('迷你图标记配置无效。')
    settings.points = {}
    for (const name of SPARKLINE_POINT_ORDER) {
      const point = config.points[name]
      if (point === undefined) continue
      if (!point || typeof point !== 'object' || Array.isArray(point)) throw new TypeError('迷你图标记配置无效。')
      settings.points[name] = { ...point }
    }
  }
  return settings
}
export function sparklineConfigToPublic(settings: ResolvedSparklineSettings): ISparklineGroupConfig {
  const types: Record<SparklineType, SparklineTypeEnum> = { line: 1, column: 2, winloss: 3, pie: 4 }
  const empty: Record<SparklineEmptyCells, SpacialShowAsEnum> = { gap: 'spacing', zero: 'zero_value', connect: 'connect_with_line' }
  const extrema = { individual: 1, group: 2, custom: 3 } as const
  return {
    type: types[settings.type], axis: { ...settings.axisOptions },
    extremumMin: { ...settings.extremumMin, type: extrema[settings.extremumMin.type] },
    extremumMax: { ...settings.extremumMax, type: extrema[settings.extremumMax.type] },
    themeType: settings.themeType, seriesColor: settings.color, emptyShowAs: empty[settings.emptyCells], nonNumShowAs: empty[settings.nonNumericCells],
    points: Object.fromEntries(SPARKLINE_POINT_ORDER.map(name => [name, { ...settings.points[name] }])),
    containHiddenCells: settings.containHiddenCells, lineWidth: settings.lineWidth, showGradient: settings.showGradient, showRadius: settings.showRadius,
  }
}
