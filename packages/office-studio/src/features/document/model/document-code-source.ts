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
