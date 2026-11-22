import { listLevelSpec, listPresets, type ListPreset } from '../features/document/model/document-lists.ts'
import { numberingDefinitionFromParts, type DocumentNumberingDefinition, type DocumentNumberingLevel } from '../features/document/model/document-numbering.ts'
import { array, record } from './validation.ts'

function matchPreset(levels: Map<number, DocumentNumberingLevel>): ListPreset | undefined {
  return listPresets.find(preset => {
    if (levels.size !== listLevelSpec(preset, 8).index + 1) return false
    for (let i = 0; i < levels.size; i++) {
      const level = levels.get(i), spec = listLevelSpec(preset, i)
      if (!level || spec.format !== level.format || spec.pattern !== level.text ||
          (level.start ?? 1) !== 1 || level.lvlRestart != null || level.isLgl === true ||
          level.pictureBulletId != null || level.pictureBullet != null || level.minimumDigits != null || level.customFormat != null) return false
    }
    return true
  })
}

export function documentNumberingIndex(value: unknown) {
  const definitions = record(value)
  const abstracts = new Map<number, Record<string, unknown>>()
  for (const a of array(definitions.abstracts).map(record)) {
    if (typeof a.abstractNumId !== 'number') continue
    abstracts.set(a.abstractNumId, a)
  }
  const index = new Map<number, { preset?: ListPreset; levels: Map<number, DocumentNumberingLevel>; definition: DocumentNumberingDefinition }>()
  for (const n of array(definitions.instances).map(record)) {
    if (typeof n.numId !== 'number' || typeof n.abstractNumId !== 'number') continue
    const base = abstracts.get(n.abstractNumId)
    if (!base) continue
    const definition = numberingDefinitionFromParts(base, n)
    const levels = new Map(definition.levels.map(level => [level.ilvl, level]))
    const custom = array(n.levelOverrides).length || Object.keys(record(n.levelStartOverrides)).length || n.consecutiveListId != null || base.consecutiveNumbering === true
    index.set(n.numId, { levels, definition, preset: custom ? undefined : matchPreset(levels) })
  }
  return index
}
