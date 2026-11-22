export type SheetColorField = 'color' | 'background'
export type SheetColorInputFormat = 'hex' | 'rgb' | 'hsb'
export type SheetColorMemory = Partial<Record<SheetColorField, string>>

/** Copyright DreamNum Co., Ltd. (Apache-2.0): five rows of nine preset colors. */
export const SHEET_COLOR_PRESETS = [
  ['#ffffff', '#e1effe', '#fde8e8', '#feecdc', '#fff4b9', '#def7ec', '#d5f5f6', '#edebfe', '#fce8f3'],
  ['#cdd0d8', '#a4cafe', '#f8b4b4', '#fdba8c', '#fac815', '#84e1bc', '#7edce2', '#cabffd', '#f8b4d9'],
  ['#979dac', '#3f83f8', '#f05252', '#ff5a1f', '#d49d0f', '#0da471', '#0694a2', '#9061f9', '#e74694'],
  ['#414657', '#1a56db', '#c81e1e', '#b43403', '#9a6d15', '#046c4e', '#036672', '#6c2bd9', '#bf125d'],
  ['#000000', '#233876', '#771d1d', '#8a2c0d', '#634312', '#014737', '#014451', '#4a1d96', '#751a3d'],
] as const
export const SHEET_STANDARD_COLORS = SHEET_COLOR_PRESETS.flat()

/** Cell color strings remain model data; this helper only canonicalizes picker hex values. */
export function sheetColorHex(value: unknown): string | undefined {
  if (typeof value !== 'string' || !/^#(?:[\da-f]{3}|[\da-f]{6})$/i.test(value.trim())) return undefined
  const hex = value.trim().slice(1).toLowerCase()
  return `#${hex.length === 3 ? [...hex].map(character => character + character).join('') : hex}`
}

const channelsToHex = (channels: readonly number[]) => `#${channels.map(channel => Math.round(channel).toString(16).padStart(2, '0')).join('')}`

/** The opaque Sheets picker uses the RGB portion of a ConfigProvider text token. */
export function sheetOpaqueColorHex(value: string): string | undefined {
  const hex = sheetColorHex(value)
  if (hex) return hex
  const rgb = /^rgba?\(\s*(\d+(?:\.\d+)?)\s*[, ]\s*(\d+(?:\.\d+)?)\s*[, ]\s*(\d+(?:\.\d+)?)(?:\s*[,/]\s*(?:0(?:\.\d+)?|1(?:\.0+)?))?\s*\)$/i.exec(value.trim())
  if (!rgb) return undefined
  const channels = rgb.slice(1, 4).map(Number)
  return channels.every(channel => channel >= 0 && channel <= 255) ? channelsToHex(channels) : undefined
}

/** Partial or invalid input must never apply the picker's previous valid scalar. */
export function sheetColorDraftValue(format: SheetColorInputFormat, fields: readonly string[]): string | undefined {
  if (format === 'hex') {
    const hex = fields[0]?.trim() ?? ''
    return fields.length === 1 && /^#?[\da-f]{6}$/i.test(hex) ? sheetColorHex(hex.startsWith('#') ? hex : `#${hex}`) : undefined
  }
  if (fields.length !== 3) return undefined
  const channels: number[] = []
  for (const [index, field] of fields.entries()) {
    const text = field.trim()
    if (!/^\d+(?:\.\d+)?%?$/.test(text) || text.endsWith('%') && (format !== 'hsb' || index === 0)) return undefined
    const value = Number(text.replace(/%$/, ''))
    const maximum = format === 'rgb' ? 255 : index === 0 ? 360 : 100
    if (!Number.isFinite(value) || value < 0 || value > maximum || format === 'rgb' && !Number.isInteger(value)) return undefined
    channels.push(value)
  }
  if (format === 'rgb') return channelsToHex(channels)
  const hue = (channels[0]! % 360) / 60, saturation = channels[1]! / 100, brightness = channels[2]! / 100
  const chroma = brightness * saturation, x = chroma * (1 - Math.abs(hue % 2 - 1)), m = brightness - chroma
  const rgb = hue < 1 ? [chroma, x, 0] : hue < 2 ? [x, chroma, 0] : hue < 3 ? [0, chroma, x]
    : hue < 4 ? [0, x, chroma] : hue < 5 ? [x, 0, chroma] : [chroma, 0, x]
  return channelsToHex(rgb.map(channel => (channel + m) * 255))
}

export function rememberSheetColor(previous: SheetColorMemory, field: SheetColorField, value: unknown): SheetColorMemory {
  const color = sheetColorHex(value)
  return color && previous[field] !== color ? { ...previous, [field]: color } : previous
}

/** An unset border color and a border whose edges have different colors are distinct. */
export function sheetBorderColor(border: Record<string, unknown> | undefined): string | null | undefined {
  const colors: string[] = []
  for (const edge of Object.values(border ?? {})) {
    if (!edge || typeof edge !== 'object' || !('color' in edge) || typeof edge.color !== 'string') continue
    colors.push(sheetColorHex(edge.color) ?? edge.color)
  }
  return !colors.length ? null : colors.every(color => color === colors[0]) ? colors[0] : undefined
}
