export interface SheetValidationListHighlightRange { start: number; end: number }

let segmenter: Intl.Segmenter | undefined
function* segments(label: string): Generator<{ text: string; start: number }> {
  if (typeof Intl.Segmenter === 'function') {
    segmenter ??= new Intl.Segmenter(undefined, { granularity: 'grapheme' })
    for (const segment of segmenter.segment(label)) yield { text: segment.segment, start: segment.index }
  } else {
    // Matches the existing text-layout fallback: code points preserve surrogate
    // pairs. Full grapheme boundaries require the browser Segmenter capability.
    let start = 0
    for (const text of label) { yield { text, start }; start += text.length }
  }
}

/** First folded match, mapped back to whole visible graphemes. Filtering and
 * Rust queries keep the same ambient locale boundary; highlighting never
 * changes candidate spelling, order or membership. Iterate only through the
 * matched prefix rather than building a full code-unit/grapheme mapping array.
 */
export function sheetValidationListHighlightRange(label: string, query: string): SheetValidationListHighlightRange | undefined {
  if (!query) return undefined
  const needle = query.toLocaleLowerCase(), folded = label.toLocaleLowerCase(), first = folded.indexOf(needle)
  if (!needle || first < 0) return undefined
  const last = first + needle.length
  if (/^[\x00-\x7f]*$/.test(label)) return { start: first, end: last }
  let offset = 0, start: number | undefined, end: number | undefined
  for (const segment of segments(label)) {
    const next = offset + segment.text.toLocaleLowerCase().length
    if (next > first && offset < last) {
      start ??= segment.start
      end = segment.start + segment.text.length
    }
    offset = next
    if (offset >= last) break
  }
  return start === undefined || end === undefined ? undefined : { start, end }
}
