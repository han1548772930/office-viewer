import type { DocumentSectionHeaderFooter } from '../../../domain/model'
import { createRun, runText, sliceRuns, type DocumentRun } from '../model/document-model'
import { resolveDocumentPageFieldResult } from '../model/document-page-field-result'
import { documentStoryContent, documentStoryDefaultStyle, documentStoryFieldDisplayRuns, type DocumentStoryKey, type StoryDefaultsSource } from '../model/document-story'
import { storyBuffer } from './story-editor'

interface StoryDisplaySegment { start: number; end: number; bufferStart: number; bufferEnd: number; field: boolean }
export interface StoryDisplayParagraph { start: number; end: number; bufferStart: number; bufferEnd: number; emptyRun: DocumentRun; style: ReturnType<typeof documentStoryContent>['paragraphs'][number]['style'] }

/** Field results are rendered with their real width but retain a single input-buffer object range. */
export function documentStoryDisplay(section: DocumentSectionHeaderFooter, story: DocumentStoryKey, page: number, total: number, model?: StoryDefaultsSource) {
  const content = documentStoryContent(section, story, model), defaults = documentStoryDefaultStyle(model)
  const buffer = storyBuffer(section, story), runs: DocumentRun[] = [], segments: StoryDisplaySegment[] = []
  let literalOffset = 0
  const sourceRuns = content.runs.map(run => { const start = literalOffset; literalOffset += run.text.length; return { start, end: literalOffset, run } })
  let offset = 0
  const append = (run: DocumentRun, bufferStart: number, bufferEnd: number, field = false) => {
    runs.push(run)
    segments.push({ start: offset, end: offset + run.text.length, bufferStart, bufferEnd, field })
    offset += run.text.length
  }
  if (!buffer.value.length) for (const run of content.runs) append(run, 0, 0)
  for (let position = 0; position < buffer.value.length;) {
    const field = buffer.objects.get(position)
    if (field) {
      const text = resolveDocumentPageFieldResult(field, page, total)
      for (const run of documentStoryFieldDisplayRuns(field, text, defaults.text)) append(run, position, position + 1, true)
      position++
    } else if (buffer.literalOffsets[position] === buffer.literalOffsets[position + 1]) {
      append(createRun(buffer.value[position]!, defaults.text), position, position + 1)
      position++
    } else {
      let end = position + 1
      while (end < buffer.value.length && !buffer.objects.has(end) && buffer.literalOffsets[end] !== buffer.literalOffsets[end + 1]) end++
      let cursor = position
      for (const run of sliceRuns(content.runs, buffer.literalOffsets[position]!, buffer.literalOffsets[end]!)) {
        append(run, cursor, cursor + run.text.length)
        cursor += run.text.length
      }
      position = end
    }
  }
  const displayOffset = (at: number) => {
    const segment = segments.find(segment => at >= segment.bufferStart && at < segment.bufferEnd)
    return segment ? segment.start + (segment.field ? 0 : at - segment.bufferStart) : offset
  }
  const paragraphs: StoryDisplayParagraph[] = content.paragraphs.map((paragraph, index) => {
    const bufferStart = buffer.literalOffsets.findIndex(offset => offset === paragraph.startOffset)
    const bufferEnd = index === content.paragraphs.length - 1 ? buffer.value.length
      : buffer.literalOffsets.findIndex((offset, position) => offset === paragraph.endOffset && buffer.value[position] === '\n')
    const source = sourceRuns.filter(span => span.start === paragraph.startOffset && span.start === span.end).at(-1)?.run
      ?? sourceRuns.find(span => span.start <= paragraph.startOffset && span.end > paragraph.startOffset)?.run
      ?? sourceRuns.at(-1)?.run ?? createRun('', defaults.text)
    return { start: displayOffset(bufferStart), end: displayOffset(bufferEnd), bufferStart, bufferEnd, emptyRun: { ...source, text: '' }, style: paragraph.style }
  })
  const inputRange = (from: number, to: number) => {
    const start = Math.min(from, to), end = Math.max(from, to)
    if (start === end) {
      const next = segments.find(segment => start >= segment.start && start < segment.end)
      const prior = [...segments].reverse().find(segment => segment.end <= start)
      const point = next ? next.field ? next.bufferStart : next.bufferStart + start - next.start
        : prior?.bufferEnd ?? 0
      return { start: point, end: point }
    }
    const first = segments.find(segment => start >= segment.start && start < segment.end)
    const last = segments.find(segment => end > segment.start && end <= segment.end)
    return { start: first ? first.field ? first.bufferStart : first.bufferStart + start - first.start : buffer.value.length,
      end: last ? last.field ? last.bufferEnd : last.bufferStart + end - last.start : buffer.value.length }
  }
  return { runs, paragraphs, text: runText(runs), inputRange }
}
