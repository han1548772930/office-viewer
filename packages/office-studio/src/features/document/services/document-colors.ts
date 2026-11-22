import type { DocumentCommand } from '../commands/document-commands'
import type { DocumentRunClearField, RunPatch } from '../model/document-model'

export type DocumentColorField = DocumentRunClearField
export type DocumentColorMemory = Partial<Record<DocumentColorField, string>>
export type DocumentColorInputFormat = 'hex' | 'rgb' | 'hsb'
export interface DocumentColorParams { patch: RunPatch; clearFields?: DocumentRunClearField[] }

export const DOCUMENT_DEFAULT_TEXT_COLOR = '#1f2937'
/** DreamNum (Apache-2.0), Univer eaf4cfa design/color-picker/presets.ts: five rows of nine. */
export const DOCUMENT_COLOR_PRESETS = [
  ['#ffffff', '#e1effe', '#fde8e8', '#feecdc', '#fff4b9', '#def7ec', '#d5f5f6', '#edebfe', '#fce8f3'],
  ['#cdd0d8', '#a4cafe', '#f8b4b4', '#fdba8c', '#fac815', '#84e1bc', '#7edce2', '#cabffd', '#f8b4d9'],
  ['#979dac', '#3f83f8', '#f05252', '#ff5a1f', '#d49d0f', '#0da471', '#0694a2', '#9061f9', '#e74694'],
  ['#414657', '#1a56db', '#c81e1e', '#b43403', '#9a6d15', '#046c4e', '#036672', '#6c2bd9', '#bf125d'],
  ['#000000', '#233876', '#771d1d', '#8a2c0d', '#634312', '#014737', '#014451', '#4a1d96', '#751a3d'],
] as const
export const DOCUMENT_STANDARD_COLORS = DOCUMENT_COLOR_PRESETS.flat()

export function isDocumentColorField(value: unknown): value is DocumentColorField {
  return value === 'color' || value === 'highlight'
}

/** The model keeps automatic text color as `auto`; a picker needs a concrete preview. */
export function documentColorIsAutomatic(field: DocumentColorField, value?: string): boolean {
  return value === undefined || value === '' || (field === 'color' && value === 'auto')
}

export function documentColorPreview(field: DocumentColorField, value: string | undefined, defaultColor: string): string {
  return documentColorIsAutomatic(field, value) ? defaultColor : value ?? defaultColor
}

/** ColorPicker supplies opaque hex values; clear intent remains explicit in the command. */
export function documentColorParams(field: DocumentColorField, value: string | null): DocumentColorParams | undefined {
  if (value === null) return { patch: {}, clearFields: [field] }
  if (!/^#(?:[\da-f]{3}|[\da-f]{6})$/i.test(value)) return undefined
  const hex = value.slice(1).toLowerCase()
  return { patch: { [field]: `#${hex.length === 3 ? [...hex].map(character => character + character).join('') : hex}` } }
}

/** A partially typed picker value must not commit the previous, still valid color. */
export function documentColorDraftIsValid(format: DocumentColorInputFormat, fields: readonly string[]): boolean {
  if (format === 'hex') return fields.length === 1 && /^#?[\da-f]{6}$/i.test(fields[0]?.trim() ?? '')
  if (fields.length !== 3) return false
  return fields.every((field, index) => {
    const text = field.trim().replace(/%$/, '')
    if (!/^\d+(?:\.\d+)?$/.test(text)) return false
    const value = Number(text)
    const max = format === 'rgb' ? 255 : index === 0 ? 360 : 100
    return Number.isFinite(value) && value >= 0 && value <= max && (format !== 'rgb' || Number.isInteger(value))
  })
}

/** The main button remembers executed colors independently of selection-derived panel values. */
export function rememberDocumentColors(previous: DocumentColorMemory, command: DocumentCommand): DocumentColorMemory {
  if (command.id !== 'doc.command.format-text') return previous
  const next = { ...previous }
  for (const field of ['color', 'highlight'] as const) {
    const value = command.patch[field]
    if (typeof value === 'string' && !documentColorIsAutomatic(field, value) && !command.clearFields?.includes(field)) next[field] = value
  }
  return next
}
