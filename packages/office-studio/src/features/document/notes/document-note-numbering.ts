import type { DocumentModel } from '../../../domain/model'
import { numberingOrdinal } from '../model/document-numbering-format'
import { tableForBlock } from '../model/document-model'
import { documentTableCellContentContainers } from '../model/document-table-continuations'
import { documentNoteReferenceHosts, documentNoteReferenceRecords, type DocumentNoteReferenceRecord } from './document-note-references'
import { effectiveDocumentNoteSettings, type EffectiveDocumentNoteSettings } from './document-note-settings'

export const DOCUMENT_NOTE_NUMBER_FORMATS = [
  ['decimal', '1, 2, 3'], ['upperLetter', 'A, B, C'], ['lowerLetter', 'a, b, c'], ['upperRoman', 'I, II, III'],
  ['lowerRoman', 'i, ii, iii'], ['chicago', '*, †, ‡'], ['decimalEnclosedCircle', '①, ②, ③'],
  ['decimalFullWidth', '１, ２, ３'], ['chineseCounting', '一, 二, 三'], ['chineseLegalSimplified', '壹, 贰, 叁'],
  ['ideographEnclosedCircle', '㈠, ㈡, ㈢'], ['ideographTraditional', '甲, 乙, 丙'], ['ideographZodiac', '子, 丑, 寅'],
] as const

function chineseCounting(value: number, legal = false): string {
  const digits = legal ? '零壹贰叁肆伍陆柒捌玖' : '零一二三四五六七八九'
  const units = legal ? ['', '拾', '佰', '仟'] : ['', '十', '百', '千']
  const bigUnits = ['', '万', '亿', '兆']
  if (value === 0) return digits[0]!
  const groups: number[] = []
  for (let remaining = value; remaining; remaining = Math.floor(remaining / 10_000)) groups.push(remaining % 10_000)
  let result = '', zero = false
  for (let group = groups.length - 1; group >= 0; group--) {
    const current = groups[group]!
    if (!current) { zero = !!result; continue }
    if (result && (zero || current < 1000)) result += digits[0]
    zero = false
    let part = '', gap = false
    for (let slot = 3; slot >= 0; slot--) {
      const digit = Math.floor(current / 10 ** slot) % 10
      if (!digit) { if (part) gap = true; continue }
      if (gap) part += digits[0]
      part += digits[digit]! + units[slot]!
      gap = false
    }
    result += part + bigUnits[group]!
  }
  return !legal && result.startsWith('一十') ? result.slice(1) : result
}

/** Public format names stay persisted; unsupported imported formats use an explicit decimal display fallback. */
export function formatDocumentNoteNumber(value: number, format = 'decimal'): { text: string; recognized: boolean } {
  if (!Number.isSafeInteger(value) || value < 1) throw new RangeError('注释编号超出正的安全整数范围。')
  if (format === 'none') return { text: '', recognized: true }
  if (format === 'chicago') {
    const symbols = ['*', '†', '‡', '§', '‖', '¶'], index = value - 1
    const count = Math.floor(index / symbols.length) + 1
    if (count > 4096) throw new RangeError('符号注释编号超过显示预算。')
    return { text: symbols[index % symbols.length]!.repeat(count), recognized: true }
  }
  if (format === 'decimalEnclosedCircle' && value <= 50) return { text: String.fromCodePoint(value <= 20 ? 0x2460 + value - 1 : value <= 35 ? 0x3251 + value - 21 : 0x32b1 + value - 36), recognized: true }
  if (format === 'ideographEnclosedCircle' && value <= 10) return { text: String.fromCodePoint(0x3220 + value - 1), recognized: true }
  if (format === 'ideographTraditional' || format === 'ideographZodiac') {
    const alphabet = format === 'ideographTraditional' ? '甲乙丙丁戊己庚辛壬癸' : '子丑寅卯辰巳午未申酉戌亥'
    if (value <= alphabet.length) return { text: alphabet[value - 1]!, recognized: true }
  }
  if (format === 'chineseCounting' || format === 'chineseLegalSimplified') return { text: chineseCounting(value, format === 'chineseLegalSimplified'), recognized: true }
  if (format === 'decimalFullWidth' || format === 'decimalFullWidth2') return { text: String(value).replace(/\d/g, digit => String.fromCodePoint(0xff10 + Number(digit))), recognized: true }
  if (format === 'ideographDigital' || format === 'taiwaneseDigital') return { text: String(value).replace(/\d/g, digit => '〇一二三四五六七八九'[Number(digit)]!), recognized: true }
  return numberingOrdinal(value, format)
}

export interface DocumentNoteLabel {
  reference: DocumentNoteReferenceRecord
  label: string
  number?: number
  settings: EffectiveDocumentNoteSettings
  recognized: boolean
  page?: number
}

/** Physical page indices drive eachPage; printed page numbers and stable note IDs remain independent. */
export function resolveDocumentNoteLabels(model: DocumentModel, referencePages: ReadonlyMap<string, number> = new Map()): Map<string, DocumentNoteLabel> {
  const references = documentNoteReferenceRecords(model).filter(reference => !reference.linked)
  const hosts = new Map<string, number>(), blocks = new Map<string, number>()
  let position = 0
  // Collection insertion order never defines the displayed sequence.
  const ordered = (items: import('../../../domain/model').DocumentBlock[], host = 'body') => {
    for (const block of items) {
      blocks.set(`${host}:${block.id}`, position++)
      const table = tableForBlock(block)
      for (const row of table?.rows ?? []) for (const cell of row.cells) {
        for (const content of documentTableCellContentContainers(cell)) ordered(content, host)
      }
    }
  }
  for (const owner of documentNoteReferenceHosts(model)) {
    const host = owner.host ? JSON.stringify(owner.host) : 'body'
    hosts.set(host, hosts.size)
    ordered(owner.blocks, host)
  }
  references.sort((left, right) => {
    const a = left.host ? JSON.stringify(left.host) : 'body', b = right.host ? JSON.stringify(right.host) : 'body'
    return hosts.get(a)! - hosts.get(b)! || blocks.get(`${a}:${left.block.id}`)! - blocks.get(`${b}:${right.block.id}`)! || left.offset - right.offset
  })
  const result = new Map<string, DocumentNoteLabel>(), counters = new Map<string, { next: number; sectionId: string; page?: number }>()
  for (const reference of references) {
    const settings = effectiveDocumentNoteSettings(model, reference.kind, reference.sectionId)
    const key = `${reference.kind}:${reference.note.id}`, page = referencePages.get(key) ?? referencePages.get(reference.note.id)
    const previous = counters.get(reference.kind)
    const restart = previous === undefined || settings.restart === 'eachSect' && previous.sectionId !== reference.sectionId
      || settings.restart === 'eachPage' && page !== undefined && previous.page !== page
    const next = restart ? settings.startNumber : previous!.next
    const display = reference.note.customMark === undefined ? formatDocumentNoteNumber(next, settings.numberFormat) : { text: reference.note.customMark, recognized: true }
    const number = reference.note.customMark === undefined ? next : undefined
    result.set(key, { reference, settings, label: display.text, recognized: display.recognized, ...(number === undefined ? {} : { number }), ...(page === undefined ? {} : { page }) })
    counters.set(reference.kind, { next: number === undefined ? next : next + 1, sectionId: reference.sectionId, page })
  }
  return result
}
