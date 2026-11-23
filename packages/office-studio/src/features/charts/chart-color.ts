import { color as echartsColor } from 'echarts/core'

/** ECharts' pure CSS parser supports named colors and HSL without consulting DOM or the current theme. */
export function chartColorChannels(value: string): [number, number, number, number] {
  const parsed = echartsColor.parse(value)
  if (!parsed || parsed.length < 4 || parsed.some(channel => !Number.isFinite(channel))) throw new TypeError(`无效或尚不支持的图表颜色：${value}`)
  return [parsed[0]! / 255, parsed[1]! / 255, parsed[2]! / 255, parsed[3]!]
}

/** Canvas2D and the shared WebGPU geometry parser receive the same canonical color. */
export function normalizeChartColor(value: string): string {
  const [r, g, b, a] = chartColorChannels(value)
  return `rgba(${Math.round(r * 255)}, ${Math.round(g * 255)}, ${Math.round(b * 255)}, ${a})`
}
