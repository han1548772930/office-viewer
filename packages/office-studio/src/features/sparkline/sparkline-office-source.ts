import type { OfficeSparklineSource } from '../../domain/sparkline-office-source.ts'

const COLOR_NAMES = new Set(['colorSeries', 'colorNegative', 'colorAxis', 'colorMarkers', 'colorFirst', 'colorLast', 'colorHigh', 'colorLow'])
const SOURCE_FIELDS = new Set(['version', 'groupXml', 'resolvedColors', 'themeXml', 'indexedColors'])
const encoder = new TextEncoder()

function object(value: unknown, label: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) throw new Error(`迷你图 Office ${label}格式无效。`)
  return value as Record<string, unknown>
}
function xml(value: unknown, limit: number, label: string): string {
  if (typeof value !== 'string' || !value.trim().startsWith('<') || value.length > limit || encoder.encode(value).byteLength > limit) {
    throw new Error(`迷你图 Office ${label}缺失或超出保留预算。`)
  }
  return value
}

/** XML is inert here. The shared Rust adapter validates names, namespaces and resources before writing. */
export function normalizeOfficeSparklineSource(value: unknown): OfficeSparklineSource | undefined {
  if (value === undefined) return undefined
  const raw = object(value, '源配置')
  if (raw.version !== 1 || Object.keys(raw).some(key => !SOURCE_FIELDS.has(key))) throw new Error('迷你图 Office 源配置版本或字段无效。')
  const colors = object(raw.resolvedColors, '颜色基线'), resolvedColors: Record<string, string> = {}
  for (const [name, color] of Object.entries(colors)) {
    if (!COLOR_NAMES.has(name) || typeof color !== 'string' || !color.length || color.length > 256) throw new Error('迷你图 Office 颜色基线无效。')
    resolvedColors[name] = color
  }
  const result: OfficeSparklineSource = { version: 1, groupXml: xml(raw.groupXml, 128 * 1024, '分组源'), resolvedColors }
  if (raw.themeXml !== undefined) result.themeXml = xml(raw.themeXml, 16 * 1024 * 1024, '主题源')
  if (raw.indexedColors !== undefined) {
    if (!Array.isArray(raw.indexedColors) || raw.indexedColors.length > 256 || raw.indexedColors.some(color => typeof color !== 'string' || !/^[\da-f]{8}$/i.test(color))) {
      throw new Error('迷你图 Office 索引调色板无效。')
    }
    result.indexedColors = raw.indexedColors.map(color => String(color).toUpperCase())
  }
  return result
}
