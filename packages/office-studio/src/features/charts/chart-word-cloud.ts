import type { ChartType, ChartWordCloudOptions, ChartWordCloudShape } from '../../domain/model'
import type { ChartData } from './chart-engine'

export const ChartWordCloudShapes = Object.freeze({ Cloud: 'cloud', Rect: 'rect', Circle: 'circle', Cardioid: 'cardioid',
  Heart: 'heart', Diamond: 'diamond', TriangleForward: 'triangle-forward', Triangle: 'triangle', Pentagon: 'pentagon', Star: 'star' } as const)
export const ChartWordCloudShapeSources = Object.freeze({ Shape: 'shape', MaskImage: 'maskImage' } as const)
export const CHART_WORD_CLOUD_DEFAULTS = Object.freeze({ shape: 'cloud' as const, repeat: false, minFontSize: 16, maxFontSize: 56,
  rotations: Object.freeze([0, -45, 45, 90, -90]), randomSeed: 1, gridSize: 4, padding: 1 })

export interface ChartWordCloudDatum { id: string; text: string; categoryIndex: number; value: number; realValue: number; repeated: boolean }

export function assertChartWordCloudType(type: ChartType): void {
  if (type !== 'wordCloud') throw new TypeError('此操作仅适用于词云图。')
}
export function chartWordCloudShape(value: unknown): ChartWordCloudShape {
  const shape = Object.values(ChartWordCloudShapes).find(shape => shape === value)
  if (shape === undefined) throw new TypeError('词云形状无效。')
  return shape
}
export function chartWordCloudMaskImageId(value: string): string {
  if (typeof value !== 'string' || !value.trim()) throw new TypeError('词云遮罩需要非空的图片资源标识。')
  return value
}
export function chartWordCloudRepeat(value: boolean): boolean {
  if (typeof value !== 'boolean') throw new TypeError('词云重复填充必须为布尔值。')
  return value
}
export function chartWordCloudEffectiveMask(options?: ChartWordCloudOptions): { type: 'shape'; value: ChartWordCloudShape } | { type: 'image'; value: string } {
  const image = options?.maskImageId ?? options?.maskImage
  return image ? { type: 'image', value: image } : { type: 'shape', value: options?.shape ?? CHART_WORD_CLOUD_DEFAULTS.shape }
}

/** Fixed v1.6.0 converter keeps source identities and uses absolute weights.
 * Repeats carry their smaller layout weight and the original tooltip value. */
export function chartWordCloudData(data: ChartData, options?: ChartWordCloudOptions): ChartWordCloudDatum[] {
  const source = data.series[0]
  if (!source) return []
  let minimum = Number.MAX_VALUE, zeroCount = 0
  const words = source.values.map((raw, categoryIndex): ChartWordCloudDatum => {
    const value = raw === null || raw === undefined || Number.isNaN(raw) ? 0 : Math.abs(Number(raw))
    if (!Number.isFinite(value)) throw new RangeError('词云权重必须能够转换为有限数值。')
    minimum = Math.min(minimum, value)
    if (value === 0) zeroCount++
    const label = data.categories[categoryIndex] ?? '', text = label.length > 100 ? `${label.slice(0, 100)}...` : label
    return { id: String(categoryIndex), text, categoryIndex, value, realValue: value, repeated: false }
  })
  if (zeroCount === words.length) {
    words.length = Math.min(5, words.length)
    for (const word of words) word.value = 1
    minimum = 1
  }
  if ((options?.repeat ?? false) && words.length > 0 && words.length < 200) {
    const sourceCount = words.length
    for (let index = sourceCount; index < 200; index++) {
      const word = words[index % sourceCount]!
      words.push({ ...word, id: `${word.id}-repeat-${index}`, value: minimum * ((index % sourceCount + 1) / (sourceCount + 1)) * 0.3,
        realValue: word.realValue, repeated: true })
    }
  }
  return words
}

/** Equal multi-item domains use the smaller endpoint; a single observation uses
 * the larger endpoint, matching the published signed-sqrt extent contract. */
export function chartWordCloudFontSize(value: number, extent: readonly [number, number], range: readonly [number, number], count: number): number {
  if (count === 1) return range[1]
  if (extent[0] === extent[1]) return range[0]
  const minimum = Math.sqrt(extent[0]), maximum = Math.sqrt(extent[1])
  return range[0] + (Math.sqrt(value) - minimum) / (maximum - minimum) * (range[1] - range[0])
}
