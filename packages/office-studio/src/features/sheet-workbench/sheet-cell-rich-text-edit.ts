import { cloneModel } from '../../domain/serialization.ts'
import { applySheetCellTextEdit, inferSheetCellTextEdit, type SheetCellTextEditOperation } from './sheet-cell-input-trace.ts'

export function sheetCellRichTextMatches(runs: ReadonlyArray<Record<string, unknown>>, text: string): boolean {
  return runs.every(run => run && typeof run === 'object' && typeof run.text === 'string')
    && runs.map(run => run.text).join('') === text
}

/** Replace one known interval. Empty source markers keep their identity instead of becoming the typing donor. */
export function spliceSheetCellRichText(runs: ReadonlyArray<Record<string, unknown>>, before: string,
  operation: SheetCellTextEditOperation, emptyDonor?: Record<string, unknown>, emptyDonorPosition?: number): Record<string, unknown>[] | undefined {
  if (!sheetCellRichTextMatches(runs, before) || applySheetCellTextEdit(before, operation) === undefined) return undefined
  const { start, end, text: inserted } = operation
  let offset = 0
  const spans = runs.map((run, index) => {
    const first = offset
    offset += (run.text as string).length
    return { run, index, first, last: offset }
  })
  // Replacement inherits the first replaced character; insertion at a boundary inherits the preceding run.
  const donor = spans.find(span => end > start || start === 0
    ? span.first <= start && span.last > start
    : span.first < start && span.last >= start) ?? spans.find(span => span.last > span.first)
  const result: Record<string, unknown>[] = []
  for (const span of spans) {
    const text = span.run.text as string
    const prefixText = text.slice(0, Math.max(0, Math.min(text.length, start - span.first)))
    const suffixText = text.slice(Math.max(0, Math.min(text.length, end - span.first)))
    const nextText = prefixText + (inserted && span.index === donor?.index ? inserted : '') + suffixText
    // Each source entity is emitted at most once: empty markers stay in order and IDs never get duplicated by splitting.
    if (nextText || text === '') result.push({ ...cloneModel(span.run), text: nextText })
  }
  if (!donor && inserted) {
    // A full deletion can retain the former caret style without consuming an original empty marker.
    const inherited = emptyDonor ? cloneModel(emptyDonor) : {}
    if (inherited.id !== undefined && result.some(run => run.id === inherited.id)) delete inherited.id
    result.splice(Math.min(result.length, emptyDonorPosition ?? result.length), 0, { ...inherited, text: inserted })
  }
  return result
}

/** Explicit operations preserve repeated-text positions. The no-trace API keeps its older inference contract. */
export function editSheetCellRichText(runs: ReadonlyArray<Record<string, unknown>>, before: string,
  after: string, operations?: ReadonlyArray<SheetCellTextEditOperation>): Record<string, unknown>[] | undefined {
  if (!sheetCellRichTextMatches(runs, before)) return undefined
  if (operations === undefined && before === after) return runs.map(run => cloneModel(run))
  const edits = operations ?? [inferSheetCellTextEdit(before, after)]
  let current = before, result = runs.map(run => cloneModel(run)), emptyDonor: Record<string, unknown> | undefined
  let emptyDonorPosition: number | undefined
  for (const operation of edits) {
    const next = applySheetCellTextEdit(current, operation)
    if (next === undefined) return undefined
    if (operation.end > operation.start) {
      let offset = 0
      for (let index = 0; index < result.length; index++) {
        const run = result[index]!
        const end = offset + (run.text as string).length
        if (offset <= operation.start && end > operation.start) { emptyDonor = cloneModel(run); emptyDonorPosition = index; break }
        offset = end
      }
    }
    const edited = spliceSheetCellRichText(result, current, operation, emptyDonor, emptyDonorPosition)
    if (!edited) return undefined
    current = next; result = edited
  }
  return current === after ? result : undefined
}
