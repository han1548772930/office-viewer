import type { DocumentBlock } from '../../../domain/model'
import definitions from '../../../domain/document-list-presets.json'

export type ListPreset = keyof typeof definitions
export type ListKind = 'bullet' | 'numbered'
type Definition = { symbols?: string[]; formats?: string[]; suffix?: string; compound?: boolean; levels?: number }
export const listPresets = Object.keys(definitions) as ListPreset[]
export const isListPreset = (value: unknown): value is ListPreset => typeof value === 'string' && Object.hasOwn(definitions, value)
export const listKind = (preset: ListPreset): ListKind => preset.startsWith('BULLET_') ? 'bullet' : 'numbered'
export function blockListPreset(block: DocumentBlock): ListPreset | undefined {
  if (block.type !== 'bullet' && block.type !== 'numbered') return
  const preset = block.list?.preset
  return isListPreset(preset) && listKind(preset) === block.type ? preset : block.type === 'bullet' ? 'BULLET_LIST' : 'ORDER_LIST'
}
export function listLevelSpec(preset: ListPreset, level: number) {
  const def: Definition = definitions[preset]
  const index = Math.max(0, Math.min((def.levels ?? 9) - 1, Math.trunc(level) || 0))
  const symbol = def.symbols?.[index % def.symbols.length]
  const format = def.formats?.[index % def.formats.length] ?? 'bullet'
  const pattern = symbol ?? (def.compound ? Array.from({ length: index + 1 }, (_, i) => `%${i + 1}`).join('.') : `%${index + 1}`) + def.suffix
  return { index, format, pattern }
}
function formattedOrdinal(value: number, format: string): string {
  if (format === 'lowerLetter' || format === 'upperLetter') {
    let result = ''
    for (let n = value; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result
    return format === 'lowerLetter' ? result.toLowerCase() : result
  }
  if (format === 'lowerRoman' && value <= 3999) {
    let result = '', n = value
    for (const [v, glyph] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']] as const) {
      while (n >= v) { result += glyph; n -= v }
    }
    return result
  }
  return format === 'decimalZero' ? String(value).padStart(2, '0') : String(value)
}
export function listMarker(preset: ListPreset, level: number, counters: readonly number[] = []): string {
  return listLevelSpec(preset, level).pattern.replace(/%([1-9])/g, (_, slot: string) => {
    const index = Number(slot) - 1
    return formattedOrdinal(counters[index] || 1, listLevelSpec(preset, index).format)
  })
}

/** Named lists resume across intervening paragraphs; legacy blocks without ids form contiguous lists. */
export function paragraphListMarkers(blocks: readonly DocumentBlock[]) {
  const markers = new Map<string, { marker: string; inset: number }>()
  const sequences = new Map<string, number[]>()
  let legacySequence = 0, previous: ListPreset | undefined
  for (const block of blocks) {
    const preset = blockListPreset(block)
    if (!preset) { previous = undefined; continue }
    const { index } = listLevelSpec(preset, block.level)
    if (previous !== preset || block.list) legacySequence++
    previous = block.list ? undefined : preset
    const key = JSON.stringify([preset, block.list?.id ?? legacySequence])
    const counters = sequences.get(key) ?? []
    counters.length = index + 1
    counters[index] = (counters[index] ?? 0) + 1
    sequences.set(key, counters)
    markers.set(block.id, { marker: listMarker(preset, index, counters), inset: 28 * (index + 1) })
  }
  return markers
}
