import type { DocumentBlock } from '../../../domain/model'

export interface DocumentNumberingPicture extends Record<string, unknown> {
  src?: string
  widthPx?: number
  heightPx?: number
}

/** Effective level fields use the duke-pages DTO names; the raw definitions remain alongside them. */
export interface DocumentNumberingLevel extends Record<string, unknown> {
  ilvl: number
  start?: number
  format?: string
  customFormat?: string
  text?: string
  minimumDigits?: number
  suffix?: 'tab' | 'space' | 'nothing'
  lvlRestart?: number
  isLgl?: boolean
  indent?: Record<string, unknown>
  runStyle?: Record<string, unknown>
  bulletFontFamily?: string
  bulletColor?: string
  pictureBulletId?: number
  pictureBullet?: DocumentNumberingPicture
}

export interface DocumentNumberingDefinition {
  abstract: Record<string, unknown>
  instance: Record<string, unknown>
  levels: DocumentNumberingLevel[]
}

function isNumberingRecord(value: unknown): value is Record<string, unknown> { return typeof value === 'object' && value !== null && !Array.isArray(value) }
export function numberingRecord(value: unknown): Record<string, unknown> { return isNumberingRecord(value) ? value : {} }

function integer(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isSafeInteger(value) ? value : undefined
}

const string = (value: unknown): string | undefined => typeof value === 'string' ? value : undefined
const positive = (value: unknown): number | undefined => typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : undefined

export function normalizeNumberingLevel(value: unknown): DocumentNumberingLevel | undefined {
  const raw = numberingRecord(value), ilvl = integer(raw.ilvl)
  if (ilvl === undefined || ilvl < 0 || ilvl > 8) return
  const picture = numberingRecord(raw.pictureBullet)
  return {
    ...raw, ilvl,
    start: integer(raw.start), format: string(raw.format), customFormat: string(raw.customFormat), text: string(raw.text),
    minimumDigits: positive(raw.minimumDigits), lvlRestart: integer(raw.lvlRestart),
    isLgl: typeof raw.isLgl === 'boolean' ? raw.isLgl : undefined,
    suffix: raw.suffix === 'tab' || raw.suffix === 'space' || raw.suffix === 'nothing' ? raw.suffix : undefined,
    indent: Object.keys(numberingRecord(raw.indent)).length ? numberingRecord(raw.indent) : undefined,
    runStyle: Object.keys(numberingRecord(raw.runStyle)).length ? numberingRecord(raw.runStyle) : undefined,
    bulletFontFamily: string(raw.bulletFontFamily), bulletColor: string(raw.bulletColor),
    pictureBulletId: integer(raw.pictureBulletId),
    pictureBullet: Object.keys(picture).length ? { ...picture, src: string(picture.src), widthPx: positive(picture.widthPx), heightPx: positive(picture.heightPx) } : undefined,
  }
}

export function numberingDefinitionFromParts(abstractValue: unknown, instanceValue: unknown): DocumentNumberingDefinition {
  const abstract = numberingRecord(abstractValue), instance = numberingRecord(instanceValue)
  const levels = new Map<number, Record<string, unknown>>()
  for (const raw of Array.isArray(abstract.levels) ? abstract.levels : []) {
    const level = numberingRecord(raw), index = integer(level.ilvl)
    if (index !== undefined && index >= 0 && index <= 8) levels.set(index, { ...level })
  }
  for (const raw of Array.isArray(instance.levelOverrides) ? instance.levelOverrides : []) {
    const level = numberingRecord(raw), index = integer(level.ilvl)
    if (index !== undefined && index >= 0 && index <= 8) levels.set(index, { ...levels.get(index), ...level })
  }
  for (const [key, raw] of Object.entries(numberingRecord(instance.levelStartOverrides))) {
    const index = Number(key), start = integer(raw), level = levels.get(index)
    if (level && start !== undefined) levels.set(index, { ...level, start })
  }
  return { abstract: { ...abstract }, instance: { ...instance }, levels: [...levels.values()].flatMap(raw => {
    const normalized = normalizeNumberingLevel(raw)
    return normalized ? [normalized] : []
  }).sort((a, b) => a.ilvl - b.ilvl) }
}

export function normalizeNumberingDefinition(value: unknown): DocumentNumberingDefinition | undefined {
  const raw = numberingRecord(value)
  if (!Array.isArray(raw.levels)) return
  const levels = raw.levels.flatMap(level => {
    const normalized = normalizeNumberingLevel(level)
    return normalized ? [normalized] : []
  })
  return { abstract: { ...numberingRecord(raw.abstract) }, instance: { ...numberingRecord(raw.instance) }, levels }
}

export function cloneNumberingDefinition(value: DocumentNumberingDefinition): DocumentNumberingDefinition {
  const result = normalizeNumberingDefinition(JSON.parse(JSON.stringify(value)))
  if (!result) throw new RangeError('The numbering definition could not be serialized')
  return result
}

export function blockNumberingDefinition(block: DocumentBlock): DocumentNumberingDefinition | undefined {
  if (!['paragraph', 'heading', 'quote', 'numbered', 'bullet', 'checklist'].includes(block.type)) return
  return normalizeNumberingDefinition(block.metadata.numberingDefinition)
}

export function blockHasNumbering(block: DocumentBlock): boolean {
  if (['numbered', 'bullet', 'checklist'].includes(block.type)) return true
  if (!['paragraph', 'heading', 'quote'].includes(block.type)) return false
  const numId = integer(numberingRecord(block.metadata.numbering).numId)
  return !!block.list || !!blockNumberingDefinition(block) || numId !== undefined && numId > 0
}

/** Heading depth and list depth are independent paragraph properties. */
export function blockNumberingLevelIndex(block: DocumentBlock): number {
  const value = ['numbered', 'bullet', 'checklist'].includes(block.type) ? block.level
    : block.metadata.numberingLevel ?? numberingRecord(block.metadata.numbering).ilvl
  return Math.max(0, Math.min(8, integer(value) ?? 0))
}

export function blockNumberingId(block: DocumentBlock): string | undefined {
  if (block.list?.id) return block.list.id
  if (typeof block.metadata.numberingListId === 'string' && block.metadata.numberingListId) return block.metadata.numberingListId
  const definition = blockNumberingDefinition(block)
  return definition && typeof definition.instance.numId === 'number' ? `num-${definition.instance.numId}` : block.checklist?.listId
}

export function consecutiveNumberingGroup(definition: DocumentNumberingDefinition, instanceId: string): { id: string; primary: boolean } {
  const group = definition.instance.consecutiveListId
  if (typeof group === 'number' && Number.isSafeInteger(group)) return { id: `num-${group}`, primary: definition.instance.numId === group || instanceId === `num-${group}` }
  if (typeof group === 'string' && group) return { id: group, primary: group === instanceId }
  return { id: instanceId, primary: true }
}

export function clearBlockNumbering(block: DocumentBlock): void {
  delete block.list
  delete block.metadata.numberingDefinition
  delete block.metadata.numberingListId
  delete block.metadata.numbering
  delete block.metadata.numberingLevel
  delete block.metadata.listLabel
}

export function copyBlockNumbering(source: DocumentBlock, target: DocumentBlock): void {
  clearBlockNumbering(target)
  if (source.list) target.list = { ...source.list }
  const definition = blockNumberingDefinition(source)
  if (definition) target.metadata.numberingDefinition = cloneNumberingDefinition(definition)
  const id = blockNumberingId(source)
  if (id) target.metadata.numberingListId = id
  if (!['numbered', 'bullet', 'checklist'].includes(target.type) && blockHasNumbering(source)) target.metadata.numberingLevel = blockNumberingLevelIndex(source)
}
