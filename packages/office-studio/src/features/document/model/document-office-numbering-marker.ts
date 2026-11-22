import type { DocumentBlock, InlineRun } from '../../../domain/model'
import { blockHasNumbering, blockNumberingLevelIndex, numberingRecord } from './document-numbering'

export interface OfficeNumberingMarkerBaseline {
  version: 1
  numbering: Record<string, unknown>
  prefix: Array<Record<string, unknown>>
}

/** This contract is also checked by the DOCX writer before replaying an imported field. */
export function officeNumberingModelContract(block: DocumentBlock): Record<string, unknown> {
  const kind = ['numbered', 'bullet', 'checklist'].includes(block.type) ? block.type : blockHasNumbering(block) ? 'numbered' : null
  return numberingRecord(JSON.parse(JSON.stringify({
    kind, level: kind ? blockNumberingLevelIndex(block) : null,
    list: block.list ?? null, checklist: block.checklist ?? null,
    definition: block.metadata.numberingDefinition ?? null,
    listId: block.metadata.numberingListId ?? null,
    numbering: block.metadata.numbering ?? null,
  })))
}

function runAppearance(run: InlineRun): Record<string, unknown> {
  const metadata = numberingRecord(run.metadata)
  const properties: Record<string, unknown> = {}
  for (const key of ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'eastAsiaFontFamily', 'fontSize', 'color', 'verticalAlign', 'href', 'footnoteId', 'endnoteId'] as const) {
    properties[key] = Reflect.get(run, key) ?? null
  }
  properties.metadata = Object.fromEntries(['highlight', 'verticalAlign', 'href', 'link', 'note', 'footnoteId', 'endnoteId', 'latexFormula'].map(key => [key, metadata[key] ?? null]))
  return properties
}

/** Ignore run IDs and segmentation, while retaining every editable property used by the writer. */
export function officeNumberingPrefixContract(block: DocumentBlock, text: string): Array<Record<string, unknown>> | undefined {
  const current = block.runs.map(run => run.text).join('')
  if (!current.startsWith(text)) return
  const previous = current.charCodeAt(text.length - 1), next = current.charCodeAt(text.length)
  if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) return
  let remaining = text.length
  const prefix: Array<Record<string, unknown>> = []
  for (const run of block.runs) {
    if (!remaining) break
    const metadata = numberingRecord(run.metadata)
    if (Reflect.get(run, 'footnoteId') != null || Reflect.get(run, 'endnoteId') != null || metadata.footnoteId != null || metadata.endnoteId != null || metadata.note != null || metadata.latexFormula != null) return
    const length = Math.min(run.text.length, remaining)
    if (!length) continue
    const properties = runAppearance(run)
    const previous = prefix.at(-1)
    if (previous && JSON.stringify(previous.properties) === JSON.stringify(properties)) previous.text = String(previous.text) + run.text.slice(0, length)
    else prefix.push({ text: run.text.slice(0, length), properties })
    remaining -= length
  }
  return prefix.map(part => numberingRecord(JSON.parse(JSON.stringify(part))))
}

export function officeNumberingMarkerBaseline(block: DocumentBlock, text: string): OfficeNumberingMarkerBaseline | undefined {
  const prefix = officeNumberingPrefixContract(block, text)
  return prefix && { version: 1, numbering: officeNumberingModelContract(block), prefix }
}

/** Only the unchanged imported prefix can be consumed when the user replaces its numbering. */
export function officeNumberingPrefixLength(block: DocumentBlock): number | undefined {
  const marker = numberingRecord(block.metadata.officeNumberingMarker)
  const baseline = numberingRecord(marker.baseline)
  if (marker.version !== 1 || baseline.version !== 1 || typeof marker.text !== 'string' || !Array.isArray(baseline.prefix)) return
  if (baseline.prefix.map(part => numberingRecord(part).text).join('') !== marker.text) return
  return officeNumberingPrefixContract(block, marker.text) ? marker.text.length : undefined
}
