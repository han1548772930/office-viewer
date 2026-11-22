import type { DocumentBlock } from '../../../domain/model'
import definitions from '../../../domain/document-list-presets.json'
import { tableForBlock } from './document-model'
import { blockHasNumbering, blockNumberingDefinition, blockNumberingId, blockNumberingLevelIndex, normalizeNumberingLevel, numberingRecord, type DocumentNumberingDefinition, type DocumentNumberingLevel, type DocumentNumberingPicture } from './document-numbering'

export type ListPreset = keyof typeof definitions
export type ListKind = 'bullet' | 'numbered'
type Definition = { symbols?: string[]; formats?: string[]; suffix?: string; compound?: boolean; levels?: number }
export const listPresets = Object.keys(definitions) as ListPreset[]
export const isListPreset = (value: unknown): value is ListPreset => typeof value === 'string' && Object.hasOwn(definitions, value)
export const listKind = (preset: ListPreset): ListKind => preset.startsWith('BULLET_') ? 'bullet' : 'numbered'
export function blockListPreset(block: DocumentBlock): ListPreset | undefined {
  if (!blockHasNumbering(block) || block.type === 'checklist') return
  const preset = block.list?.preset
  if (isListPreset(preset) && (!['bullet', 'numbered'].includes(block.type) || listKind(preset) === block.type)) return preset
  if (blockNumberingDefinition(block)) return
  if (block.type !== 'bullet' && block.type !== 'numbered') return
  return block.type === 'bullet' ? 'BULLET_LIST' : 'ORDER_LIST'
}
export function listLevelSpec(preset: ListPreset, level: number) {
  const def: Definition = definitions[preset]
  const index = Math.max(0, Math.min((def.levels ?? 9) - 1, Math.trunc(level) || 0))
  const symbol = def.symbols?.[index % def.symbols.length]
  const format = def.formats?.[index % def.formats.length] ?? 'bullet'
  const pattern = symbol ?? (def.compound ? Array.from({ length: index + 1 }, (_, i) => `%${i + 1}`).join('.') : `%${index + 1}`) + def.suffix
  return { index, format, pattern }
}
export function formatListOrdinal(value: number, format: string, minimumDigits = 0): string {
  if (format === 'lowerLetter' || format === 'upperLetter') {
    if (value < 1 || Math.ceil(value / 26) > 4096) return String(value)
    const result = String.fromCharCode(65 + (value - 1) % 26).repeat(Math.ceil(value / 26))
    return format === 'lowerLetter' ? result.toLowerCase() : result
  }
  if ((format === 'lowerRoman' || format === 'upperRoman') && value > 0 && value <= 3999999) {
    let result = '', n = value
    for (const [v, glyph] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']] as const) {
      while (n >= v) { result += glyph; n -= v }
    }
    return format === 'upperRoman' ? result.toUpperCase() : result
  }
  if (format === 'ordinal') return `${value}${value % 100 >= 11 && value % 100 <= 13 ? 'th' : value % 10 === 1 ? 'st' : value % 10 === 2 ? 'nd' : value % 10 === 3 ? 'rd' : 'th'}`
  if (format === 'hex') return value.toString(16).toUpperCase()
  const width = Math.max(format === 'decimalZero' ? 2 : 0, Math.min(255, Math.trunc(minimumDigits)))
  return value === 0 ? '0' : String(value).padStart(width, '0')
}
export function listMarker(preset: ListPreset, level: number, counters: readonly number[] = []): string {
  return listLevelSpec(preset, level).pattern.replace(/%([1-9])/g, (_, slot: string) => {
    const index = Number(slot) - 1
    return formatListOrdinal(counters[index] ?? 1, listLevelSpec(preset, index).format)
  })
}

export function presetNumberingDefinition(preset: ListPreset): DocumentNumberingDefinition {
  const levels: DocumentNumberingLevel[] = Array.from({ length: listLevelSpec(preset, 8).index + 1 }, (_, ilvl) => {
    const spec = listLevelSpec(preset, ilvl)
    return { ilvl, start: 1, format: spec.format, text: spec.pattern, suffix: 'tab', indent: { leftTwips: 420 * (ilvl + 1), hangingTwips: 420 }, bulletFontFamily: 'Arial' }
  })
  return { abstract: { abstractNumId: `preset-${preset}`, levels }, instance: {}, levels }
}

export function blockListLevel(block: DocumentBlock): { index: number; maximum: number; level?: DocumentNumberingLevel } {
  const definition = blockNumberingDefinition(block), preset = blockListPreset(block)
  const maximum = definition ? Math.max(0, ...definition.levels.map(level => level.ilvl)) : preset ? listLevelSpec(preset, 8).index : 8
  const index = Math.min(maximum, blockNumberingLevelIndex(block))
  return { index, maximum, level: definition?.levels.find(level => level.ilvl === index) ?? (preset ? presetNumberingDefinition(preset).levels[index] : undefined) }
}

export interface ParagraphListMarker {
  marker: string
  ordinal?: number
  inset: number
  markerIndent: number
  rightIndent: number
  gap: number
  fontFamily: string
  fontSize: number
  color: string
  bold: boolean
  italic: boolean
  picture?: DocumentNumberingPicture
  diagnostics: string[]
}

const finite = (value: unknown, fallback = 0): number => typeof value === 'number' && Number.isFinite(value) ? value : fallback
const text = (value: unknown, fallback: string): string => typeof value === 'string' && value ? value : fallback

function levelStart(definition: DocumentNumberingDefinition, index: number, firstUse: boolean): number {
  const base = (Array.isArray(definition.abstract.levels) ? definition.abstract.levels : []).map(normalizeNumberingLevel).find(level => level?.ilvl === index)
  const override = (Array.isArray(definition.instance.levelOverrides) ? definition.instance.levelOverrides : []).map(normalizeNumberingLevel).find(level => level?.ilvl === index)
  const value = firstUse ? numberingRecord(definition.instance.levelStartOverrides)[String(index)] : undefined
  return Math.max(0, Math.trunc(finite(value, override?.start ?? base?.start ?? definition.levels.find(level => level.ilvl === index)?.start ?? 1)))
}

function markerAppearance(block: DocumentBlock, level: DocumentNumberingLevel | undefined, index: number): Omit<ParagraphListMarker, 'marker' | 'ordinal' | 'diagnostics'> {
  const style = level?.runStyle ?? {}, indent = level?.indent ?? {}, run = block.runs[0]
  const left = finite(block.style.indentStart, finite(indent.leftTwips, 420 * (index + 1)) / 15)
  const hanging = finite(block.style.hanging, finite(indent.hangingTwips, 420) / 15)
  const firstLine = finite(block.style.indentFirstLine, finite(indent.firstLineTwips) / 15)
  return {
    inset: left, markerIndent: left - hanging + firstLine, rightIndent: finite(block.style.indentEnd, finite(indent.rightTwips) / 15),
    gap: level?.suffix === 'nothing' ? 0 : level?.suffix === 'space' ? 4 : 6,
    fontFamily: text(level?.bulletFontFamily, text(style.fontFamilyAscii, text(style.resolvedFontFamilyAscii, text(style.fontFamily, run?.fontFamily || 'Arial')))),
    fontSize: Math.max(1, finite(style.fontSizePt, run?.fontSize ?? 11)), color: text(level?.bulletColor, text(style.color, run?.color ?? '#1f2937')),
    bold: typeof style.bold === 'boolean' ? style.bold : run?.bold === true, italic: typeof style.italic === 'boolean' ? style.italic : run?.italic === true,
    picture: level?.pictureBullet,
  }
}

/** Numbering instances share one sequence across the body and table cells, in document order. */
export function paragraphListMarkers(blocks: readonly DocumentBlock[]): Map<string, ParagraphListMarker> {
  const markers = new Map<string, ParagraphListMarker>()
  type Sequence = { counters: Array<number | undefined>; total?: number; initialized: Map<string, Set<number>> }
  const sequences = new Map<string, Sequence>()
  let legacySequence = 0
  const visit = (container: readonly DocumentBlock[]) => {
    let previous: string | undefined, legacyId = 0
    for (const block of container) {
      const table = tableForBlock(block)
      if (table) {
        previous = undefined
        for (const row of table.rows) for (const cell of row.cells) visit(cell.blocks)
        continue
      }
      if (!blockHasNumbering(block) || block.type === 'checklist') { previous = undefined; continue }
      const imported = blockNumberingDefinition(block), preset = blockListPreset(block)
      const definition = imported ?? (preset ? presetNumberingDefinition(preset) : undefined)
      const { index, level } = blockListLevel(block)
      const styleKey = String(definition?.abstract.abstractNumId ?? preset ?? block.type)
      const id = blockNumberingId(block)
      if (!id && previous !== styleKey) legacyId = ++legacySequence
      previous = id ? undefined : styleKey
      const consecutive = definition?.abstract.consecutiveNumbering === true
      const instanceKey = id ?? `legacy-${legacyId}`
      const counterId = consecutive ? String(definition?.instance.consecutiveListId ?? instanceKey) : instanceKey
      const key = JSON.stringify([styleKey, counterId])
      const state: Sequence = sequences.get(key) ?? { counters: [], initialized: new Map<string, Set<number>>() }
      const initialized = state.initialized.get(instanceKey) ?? new Set<number>()
      state.initialized.set(instanceKey, initialized)
      const starts = (slot: number, firstUse: boolean) => definition ? levelStart(definition, slot, firstUse) : 1
      for (let slot = 0; slot < index; slot++) state.counters[slot] ??= starts(slot, !initialized.has(slot))
      if (consecutive) {
        const overrides = numberingRecord(definition?.instance.levelStartOverrides)
        const restart = !initialized.size && instanceKey !== counterId ? Object.values(overrides).find(value => typeof value === 'number') : undefined
        state.total = typeof restart === 'number' ? restart : state.total === undefined ? starts(0, true) : state.total + 1
        state.counters[index] = state.total
      } else state.counters[index] = state.counters[index] === undefined ? starts(index, !initialized.has(index)) : state.counters[index]! + 1
      initialized.add(index)
      for (let slot = index + 1; slot < 9; slot++) {
        const restart = definition?.levels.find(candidate => candidate.ilvl === slot)?.lvlRestart
        const trigger = restart !== undefined && restart >= 0 && restart <= slot ? restart : slot
        if (consecutive || index < trigger) state.counters[slot] = undefined
      }
      sequences.set(key, state)
      const diagnostics: string[] = []
      let marker = level?.text
      if (marker === undefined) {
        marker = text(block.metadata.listLabel, '')
        diagnostics.push('The imported list level has no usable numbering pattern.')
      }
      marker = marker.replace(/%([1-9])/g, (_whole, slot: string) => {
        const referenced = definition?.levels.find(candidate => candidate.ilvl === Number(slot) - 1)
        const format = level?.isLgl ? 'decimal' : referenced?.format ?? 'decimal'
        if (!['decimal', 'decimalZero', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'ordinal', 'hex'].includes(format)) diagnostics.push(`Number format ${format} uses a decimal display fallback.`)
        return formatListOrdinal(state.counters[Number(slot) - 1] ?? starts(Number(slot) - 1, true), format, level?.isLgl ? 0 : referenced?.minimumDigits)
      })
      if (level?.format === 'bullet') {
        const symbols: Record<string, string> = { '\uF0B7': '•', '\uF0A7': '□', '\uF0A8': '▪', '\uF0D8': '➢' }
        marker = marker.replace(/[\uF0B7\uF0A7\uF0A8\uF0D8]/g, glyph => symbols[glyph] ?? glyph)
      }
      if (level?.format === 'none' || level?.pictureBullet) marker = ''
      markers.set(block.id, { ...markerAppearance(block, level, index), marker, ordinal: state.counters[index], diagnostics })
    }
  }
  visit(blocks)
  return markers
}
