import type { DocumentDefaultStyle } from '../../../domain/model'
import { createRun, documentId, type DocumentRun, type RunPatch } from './document-model'

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
    delete metadata.noteReference
    delete metadata.latexFormula
    style.metadata = metadata
  }
  return style
}

/** Preserve run formatting and entity identities outside the edited half-open range. */
export function replaceDocumentRunRange(runs: readonly DocumentRun[], start: number, end: number, text: string, defaults: DocumentDefaultStyle['text'], marks: RunPatch = {}): DocumentRun[] {
  const before: DocumentRun[] = [], after: DocumentRun[] = []
  let offset = 0
  for (const run of runs) {
    const from = offset, to = from + run.text.length
    offset = to
    if (!run.text.length) {
      if (from <= start) before.push(run)
      else if (from >= end) after.push(run)
      continue
    }
    const prefix = Math.max(0, Math.min(run.text.length, start - from))
    const suffix = Math.max(0, Math.min(run.text.length, end - from))
    if (prefix) before.push(prefix === run.text.length ? run : { ...run, text: run.text.slice(0, prefix) })
    if (suffix < run.text.length) after.push(suffix === 0 ? run : { ...run, id: prefix ? documentId('run') : run.id, text: run.text.slice(suffix) })
  }
  const inherited = documentRunStyleAt(runs, start, end) ?? defaults
  const result = [...before, ...(text ? [createRun(text, { ...inherited, ...marks })] : []), ...after]
  return result.length ? result : [createRun('', { ...inherited, ...marks })]
}
