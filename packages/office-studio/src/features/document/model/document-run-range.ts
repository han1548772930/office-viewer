import type { DocumentDefaultStyle } from '../../../domain/model'
import { normalizeDocumentHighlight } from './document-highlight'
import { applyRunPatch, createRun, documentId, runText, type DocumentRun, type RunPatch } from './document-model'

/** The same font and text patch contract is used by commands and typing marks. */
export function validateDocumentRunPatch(patch: RunPatch): void {
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) throw new TypeError('字符格式补丁必须是对象。')
  const keys = ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'eastAsiaFontFamily', 'fontSize', 'verticalAlign', 'color', 'highlight', 'href', 'metadata']
  if (Object.keys(patch).some(key => !keys.includes(key))) throw new TypeError('字符格式补丁包含未知属性。')
  for (const key of ['bold', 'italic', 'underline', 'strike'] as const) {
    if (Object.hasOwn(patch, key) && typeof patch[key] !== 'boolean') throw new TypeError('字符强调格式必须为布尔值。')
  }
  for (const key of ['fontFamily', 'eastAsiaFontFamily'] as const) {
    if (Object.hasOwn(patch, key) && (typeof patch[key] !== 'string' || !patch[key]!.trim() || patch[key]!.includes('\0'))) throw new TypeError('字体名称无效。')
  }
  if (Object.hasOwn(patch, 'fontSize') && (typeof patch.fontSize !== 'number' || !Number.isFinite(patch.fontSize) || patch.fontSize <= 0 || patch.fontSize > 400)) throw new RangeError('字号必须大于 0 且不超过 400 pt。')
  if (Object.hasOwn(patch, 'verticalAlign') && !['baseline', 'subscript', 'superscript'].includes(patch.verticalAlign ?? '')) throw new TypeError('字符上下标格式无效。')
  if (patch.color !== undefined && (typeof patch.color !== 'string' || !/^(?:auto|#?[\da-f]{6})$/i.test(patch.color))) throw new TypeError('字符颜色无效。')
  if (patch.highlight !== undefined && (typeof patch.highlight !== 'string' || patch.highlight !== '' && patch.highlight.toLowerCase() !== 'none' && !normalizeDocumentHighlight(patch.highlight))) throw new TypeError('字符高亮颜色无效。')
  if (patch.href !== undefined && typeof patch.href !== 'string') throw new TypeError('链接格式无效。')
  if (patch.metadata !== undefined && (!patch.metadata || typeof patch.metadata !== 'object' || Array.isArray(patch.metadata))) throw new TypeError('字符来源数据无效。')
}

function validBoundary(text: string, offset: number): boolean {
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > text.length) return false
  const before = text.charCodeAt(offset - 1), after = text.charCodeAt(offset)
  return !(before >= 0xD800 && before <= 0xDBFF && after >= 0xDC00 && after <= 0xDFFF)
}

/** Merge only supplied properties, preserving empty entities and unchanged run IDs. */
export function formatDocumentRunRange(runs: readonly DocumentRun[], start: number, end: number, patch: RunPatch): DocumentRun[] {
  validateDocumentRunPatch(patch)
  const text = runText(runs)
  if (!validBoundary(text, start) || !validBoundary(text, end) || end < start) throw new RangeError('字符格式选区的 UTF-16 位置无效。')
  const result: DocumentRun[] = []
  let offset = 0
  for (const run of runs) {
    const from = Math.max(0, start - offset), to = Math.min(run.text.length, end - offset)
    offset += run.text.length
    if (to <= from) { result.push(run); continue }
    const updated = applyRunPatch(run, patch)
    if (JSON.stringify(updated) === JSON.stringify(run)) { result.push(run); continue }
    if (from > 0) result.push({ ...run, text: run.text.slice(0, from) })
    result.push({ ...updated, id: from > 0 ? documentId('run') : run.id, text: run.text.slice(from, to) })
    if (to < run.text.length) result.push({ ...run, id: documentId('run'), text: run.text.slice(to) })
  }
  return result
}

/** Replacement inherits the first selected character; insertion inherits its left neighbor. */
export function documentRunStyleAt(runs: readonly DocumentRun[], start: number, end = start): RunPatch | undefined {
  const index = end > start ? start : Math.max(0, start - 1)
  let offset = 0
  const run = runs.find(candidate => {
    const from = offset
    offset += candidate.text.length
    return candidate.text.length > 0 && from <= index && offset > index
  }) ?? runs.at(-1)
  if (!run) return undefined
  const { id: _id, text: _text, ...style } = run
  if (style.metadata) {
    const metadata = { ...style.metadata }
    // A newly typed character is not a duplicate note or embedded formula.
    delete metadata.note
    delete metadata.noteReference
    delete metadata.latexFormula
    style.metadata = metadata
  }
  return style
}

/** Keep empty entities outside a text replacement and assign a boundary entity only once. */
export function splitDocumentRunRange(runs: readonly DocumentRun[], start: number, end: number): { before: DocumentRun[]; after: DocumentRun[] } {
  const value = runText(runs)
  if (!validBoundary(value, start) || !validBoundary(value, end) || end < start) throw new RangeError('字符替换选区的 UTF-16 位置无效。')
  const before: DocumentRun[] = [], after: DocumentRun[] = []
  let offset = 0
  for (const run of runs) {
    const from = offset, to = from + run.text.length
    offset = to
    if (!run.text.length) {
      // A boundary entity belongs to the half-open suffix once, matching
      // sliceRuns and the before-reference default used by public body ranges.
      if (from < start) before.push(run)
      else if (from >= end) after.push(run)
      continue
    }
    const prefix = Math.max(0, Math.min(run.text.length, start - from))
    const suffix = Math.max(0, Math.min(run.text.length, end - from))
    if (prefix) before.push(prefix === run.text.length ? run : { ...run, text: run.text.slice(0, prefix) })
    if (suffix < run.text.length) after.push(suffix === 0 ? run : { ...run, id: prefix ? documentId('run') : run.id, text: run.text.slice(suffix) })
  }
  return { before, after }
}

/** Preserve run formatting and entity identities outside the edited half-open range. */
export function replaceDocumentRunRange(runs: readonly DocumentRun[], start: number, end: number, text: string, defaults: DocumentDefaultStyle['text'], marks: RunPatch = {}, sourceStyle?: RunPatch): DocumentRun[] {
  validateDocumentRunPatch(marks)
  const { before, after } = splitDocumentRunRange(runs, start, end)
  // Field caches can provide an inherited model style independently of the
  // literal range. Imported values are not newly requested toolbar overrides.
  const inherited = sourceStyle ?? documentRunStyleAt(runs, start, end) ?? defaults
  const result = [...before, ...(text ? [createRun(text, { ...inherited, ...marks })] : []), ...after]
  return result.length ? result : [createRun('', { ...inherited, ...marks })]
}
