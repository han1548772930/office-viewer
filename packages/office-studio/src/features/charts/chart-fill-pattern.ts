import type { ChartFillPatternPreset, ChartSeriesFillPattern } from '../../domain/model'
import { chartColorChannels, normalizeChartColor } from './chart-color'

/** DrawingML ST_PresetPatternVal; the selector is a local extension of the published style contract. */
const PATTERN_CATALOG: ReadonlyArray<readonly [ChartFillPatternPreset, string]> = [
  ['pct5', '5% 点纹'], ['pct10', '10% 点纹'], ['pct20', '20% 点纹'], ['pct25', '25% 点纹'], ['pct30', '30% 点纹'], ['pct40', '40% 点纹'],
  ['pct50', '50% 点纹'], ['pct60', '60% 点纹'], ['pct70', '70% 点纹'], ['pct75', '75% 点纹'], ['pct80', '80% 点纹'], ['pct90', '90% 点纹'],
  ['horz', '水平线'], ['vert', '垂直线'], ['ltHorz', '浅色水平线'], ['ltVert', '浅色垂直线'], ['dkHorz', '深色水平线'], ['dkVert', '深色垂直线'],
  ['narHorz', '窄水平线'], ['narVert', '窄垂直线'], ['dashHorz', '水平虚线'], ['dashVert', '垂直虚线'], ['cross', '十字线'],
  ['dnDiag', '向下对角线'], ['upDiag', '向上对角线'], ['ltDnDiag', '浅色向下对角线'], ['ltUpDiag', '浅色向上对角线'],
  ['dkDnDiag', '深色向下对角线'], ['dkUpDiag', '深色向上对角线'], ['wdDnDiag', '宽向下对角线'], ['wdUpDiag', '宽向上对角线'],
  ['dashDnDiag', '向下对角虚线'], ['dashUpDiag', '向上对角虚线'], ['diagCross', '对角十字线'],
  ['smCheck', '小棋盘'], ['lgCheck', '大棋盘'], ['smGrid', '小网格'], ['lgGrid', '大网格'], ['dotGrid', '点网格'],
  ['smConfetti', '小纸屑'], ['lgConfetti', '大纸屑'], ['horzBrick', '水平砖块'], ['diagBrick', '对角砖块'],
  ['solidDmnd', '实心菱形'], ['openDmnd', '空心菱形'], ['dotDmnd', '点菱形'], ['plaid', '格子布'], ['sphere', '球形'],
  ['weave', '编织'], ['divot', '草屑'], ['shingle', '瓦片'], ['wave', '波纹'], ['trellis', '棚架'], ['zigZag', '之字纹'],
]
export const CHART_FILL_PATTERN_PRESETS: ReadonlyArray<readonly [ChartFillPatternPreset, string]> = Object.freeze(PATTERN_CATALOG.map(entry => Object.freeze(entry)))

const PRESETS = new Set<string>(CHART_FILL_PATTERN_PRESETS.map(([preset]) => preset))
export function isChartFillPatternPreset(value: unknown): value is ChartFillPatternPreset { return typeof value === 'string' && PRESETS.has(value) }

/** Paint colors are detached from the author model; fill opacity never attenuates the border. */
export function chartFillPatternPaint(pattern: ChartSeriesFillPattern | null | undefined, opacity = 1, shade = 1): ChartSeriesFillPattern | undefined {
  if (!pattern) return undefined
  const paint = (color: string): string => {
    if (opacity === 1 && shade === 1) return normalizeChartColor(color)
    const [r, g, b, a] = chartColorChannels(color)
    return `rgba(${Math.round(Math.min(1, r * shade) * 255)}, ${Math.round(Math.min(1, g * shade) * 255)}, ${Math.round(Math.min(1, b * shade) * 255)}, ${a * opacity})`
  }
  return { preset: pattern.preset, foregroundColor: paint(pattern.foregroundColor), backgroundColor: paint(pattern.backgroundColor) }
}
