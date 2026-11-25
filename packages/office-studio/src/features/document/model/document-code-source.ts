import { isDocumentTextBoundary } from './document-story'

export interface DocumentCodeSourceEdit { startOffset: number; endOffset: number; text: string }

export function normalizeDocumentCodeSource(source: string): string {
  if (typeof source !== 'string') throw new TypeError('代码源文本必须为字符串。')
  return source.replace(/\r\n?/g, '\n')
}

/** The unchanged prefix and suffix retain their runs, fields and stable paragraph IDs. */
export function documentCodeSourceEdit(before: string, after: string): DocumentCodeSourceEdit {
  let startOffset = 0
  while (startOffset < before.length && startOffset < after.length && before[startOffset] === after[startOffset]) startOffset++
  while (startOffset > 0 && (!isDocumentTextBoundary(before, startOffset) || !isDocumentTextBoundary(after, startOffset))) startOffset--
  let suffix = 0
  while (suffix < before.length - startOffset && suffix < after.length - startOffset
    && before[before.length - suffix - 1] === after[after.length - suffix - 1]) suffix++
  while (suffix > 0 && (!isDocumentTextBoundary(before, before.length - suffix) || !isDocumentTextBoundary(after, after.length - suffix))) suffix--
  return { startOffset, endOffset: before.length - suffix, text: after.slice(startOffset, after.length - suffix) }
}

/** Descending, disjoint edits keep each untouched line and its rich-text anchors intact. */
export function documentCodeIndentEdits(source: string, startOffset: number, endOffset: number, direction: -1 | 1, tabSize: number): DocumentCodeSourceEdit[] {
  if (!Number.isSafeInteger(startOffset) || !Number.isSafeInteger(endOffset) || startOffset < 0 || endOffset < startOffset || endOffset > source.length
    || !isDocumentTextBoundary(source, startOffset) || !isDocumentTextBoundary(source, endOffset)
    || direction !== -1 && direction !== 1 || !Number.isInteger(tabSize) || tabSize < 1 || tabSize > 16) throw new RangeError('代码缩进范围无效。')
  if (startOffset === endOffset && direction === 1) return [{ startOffset, endOffset, text: '\t' }]
  const starts = [0]
  for (let offset = 0; offset < source.length; offset++) if (source[offset] === '\n' || source[offset] === '\u2028') starts.push(offset + 1)
  const first = starts.findLastIndex(offset => offset <= startOffset), last = starts.findLastIndex(offset => offset <= (endOffset > startOffset ? endOffset - 1 : endOffset))
  const edits: DocumentCodeSourceEdit[] = []
  for (let index = last; index >= first; index--) {
    const start = starts[index]!
    if (direction === 1) edits.push({ startOffset: start, endOffset: start, text: '\t' })
    else {
      let count = source[start] === '\t' ? 1 : 0
      if (!count) while (count < tabSize && source[start + count] === ' ') count++
      if (count) edits.push({ startOffset: start, endOffset: start + count, text: '' })
    }
  }
  return edits
}
