type BodyRecord = Record<string, unknown>
const record = (value: unknown): BodyRecord | undefined => value && typeof value === 'object' && !Array.isArray(value) ? value as BodyRecord : undefined
const indexedLists = ['paragraphs', 'sectionBreaks', 'customBlocks', 'docxRawCustomBlocks', 'docxRawBlocks', 'textRuns', 'customRanges',
  'customDecorations', 'tables', 'columnGroups', 'blockRanges', 'docxExportExcludedRanges'] as const

/** All offsets are UTF-16 offsets. Only the drawing sentinel changes; text and opaque extensions are retained. */
function shiftBodyOffsets(body: BodyRecord, at: number, delta: 1 | -1): void {
  const boundary = (index: number): number => delta > 0 ? index >= at ? index + 1 : index : index > at ? index - 1 : index
  const inclusiveEnd = (index: number): number => delta < 0 && index >= at ? index - 1 : boundary(index)
  const mapEntry = (entry: unknown): unknown[] => {
    const item = record(entry)
    if (!item) throw new TypeError('单元格正文索引项无效。')
    const next = { ...item }
    const start = typeof item.st === 'number' ? item.st : typeof item.startIndex === 'number' ? item.startIndex : item.start
    const end = typeof item.ed === 'number' ? item.ed : typeof item.endIndex === 'number' ? item.endIndex : item.end
    if (delta < 0 && typeof start === 'number' && typeof end === 'number') {
      const exclusive = typeof item.ed === 'number' || typeof item.start === 'number'
      if (start === at && end === at + (exclusive ? 1 : 0)) return []
    }
    for (const key of ['st', 'ed', 'startIndex', 'start', 'end'] as const) if (typeof item[key] === 'number') next[key] = boundary(item[key])
    if (typeof item.endIndex === 'number') next.endIndex = inclusiveEnd(item.endIndex)
    // Columns and table entries contain their own body offsets, while properties/metadata can contain unrelated indices.
    for (const key of ['columns', 'rows', 'cells'] as const) if (Array.isArray(item[key])) next[key] = item[key].flatMap(mapEntry)
    return [next]
  }
  for (const key of indexedLists) if (Array.isArray(body[key])) body[key] = body[key].flatMap(mapEntry)
  if (Array.isArray(body.renderedPageBreaks)) body.renderedPageBreaks = body.renderedPageBreaks.flatMap(index => {
    if (typeof index !== 'number' || !Number.isSafeInteger(index)) throw new TypeError('单元格正文分页索引无效。')
    return delta < 0 && index === at ? [] : [boundary(index)]
  })
}

export function removeSheetImageBodyDrawing(body: BodyRecord, drawingId: string): void {
  if (typeof body.dataStream !== 'string') throw new TypeError('单元格正文缺少 dataStream。')
  let stream = body.dataStream
  const blocks = Array.isArray(body.customBlocks) ? body.customBlocks : []
  const removed = blocks.filter(block => record(block)?.blockId === drawingId).map(block => record(block)!.startIndex)
  if (removed.some(index => typeof index !== 'number' || !Number.isSafeInteger(index) || index < 0 || stream[index] !== '\u0008')) {
    throw new TypeError('单元格图片引用未指向有效的绘图字符。')
  }
  body.customBlocks = blocks.filter(block => record(block)?.blockId !== drawingId)
  for (const index of [...new Set(removed as number[])].sort((a, b) => b - a)) {
    stream = stream.slice(0, index) + stream.slice(index + 1)
    body.dataStream = stream
    shiftBodyOffsets(body, index, -1)
  }
}

export function ensureSheetImageBodyDrawing(body: BodyRecord, drawingId: string, priorId?: string): void {
  if (typeof body.dataStream !== 'string') throw new TypeError('单元格正文缺少 dataStream。')
  const blocks = Array.isArray(body.customBlocks) ? body.customBlocks : []
  if (blocks.some(block => record(block)?.blockId === drawingId)) return
  const prior = priorId && blocks.find(block => record(block)?.blockId === priorId)
  if (prior) { record(prior)!.blockId = drawingId; return }
  // An unbound sentinel can be reused; otherwise prepend a real sentinel and rebase every body index.
  const hasInitialDrawing = blocks.some(block => record(block)?.startIndex === 0)
  if (!body.dataStream.startsWith('\u0008') || hasInitialDrawing) {
    shiftBodyOffsets(body, 0, 1)
    body.dataStream = '\u0008' + body.dataStream
  }
  const current = Array.isArray(body.customBlocks) ? body.customBlocks : []
  body.customBlocks = [{ startIndex: 0, blockId: drawingId, blockType: 0 }, ...current]
}
