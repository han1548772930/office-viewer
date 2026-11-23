import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { runText } from '../model/document-model'
import { documentFieldLocked, documentFieldRetainsNestedFields, documentTextFieldParagraphs, documentTextFieldRanges, type DocumentTextFieldRange } from '../model/document-text-fields'
import { DEFAULT_DOCUMENT_TOC_OPTIONS, documentTocIdAtBlock, type DocumentNativeTocMetadata, type DocumentTocMetadata, type DocumentTocOptions, type DocumentTocTabLeader } from './document-toc'

export interface DocumentNativeTocHydrationEntry {
  fieldId: string
  tocId: string
  blockIds: string[]
  fullUpdateUnavailableReason?: string
  pageUpdateUnavailableReason?: string
}

export interface DocumentNativeTocHydrationReport {
  recognized: DocumentNativeTocHydrationEntry[]
  retained: DocumentNativeTocHydrationEntry[]
  unmapped: Array<{ fieldId: string; reason: string }>
}

function record(value: unknown): Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
}

/** Tokenization is intentionally separate from result text; field codes remain untouched. */
function instructionTokens(instruction: string): string[] | undefined {
  const tokens: string[] = []
  let offset = 0
  while (offset < instruction.length) {
    if (/\s/.test(instruction[offset]!)) { offset++; continue }
    if (instruction[offset] !== '"') {
      const start = offset
      while (offset < instruction.length && !/\s/.test(instruction[offset]!)) offset++
      tokens.push(instruction.slice(start, offset))
      continue
    }
    offset++
    let value = '', closed = false
    while (offset < instruction.length) {
      const character = instruction[offset++]!
      if (character === '"') { closed = true; break }
      if (character === '\\' && (instruction[offset] === '"' || instruction[offset] === '\\')) value += instruction[offset++]!
      else value += character
    }
    if (!closed || offset < instruction.length && !/\s/.test(instruction[offset]!)) return undefined
    tokens.push(value)
  }
  return tokens
}

function nativeOptions(instruction: string): { options: DocumentTocOptions; hyperlinks: boolean; reason?: string } {
  const options = { ...DEFAULT_DOCUMENT_TOC_OPTIONS }, tokens = instructionTokens(instruction)
  if (!tokens || tokens[0]?.toUpperCase() !== 'TOC') return { options, hyperlinks: false, reason: '原生目录指令尚未完整解析，不能重建其结果。' }
  let hyperlinks = false, reason: string | undefined
  const parameterSwitches = new Set<string>()
  for (let index = 1; index < tokens.length; index++) {
    const token = tokens[index]!.toLowerCase()
    if (token === '\\h') { hyperlinks = true; continue }
    if (token === '\\u' || token === '\\z') continue
    if (token === '\\n') {
      if (tokens[index + 1] && !tokens[index + 1]!.startsWith('\\')) {
        index++
        reason ??= '原生目录按标题级别隐藏页码的指令尚未接入完整重建。'
      } else options.showPageNumbers = false
      continue
    }
    if (token === '\\o' || token === '\\p') {
      if (parameterSwitches.has(token)) reason ??= '原生目录包含重复的范围或分隔参数，不能推断完整重建结果。'
      parameterSwitches.add(token)
      const value = tokens[index + 1]
      if (value === undefined || value.startsWith('\\')) { reason ??= '原生目录指令缺少参数。'; continue }
      index++
      if (token === '\\p') {
        options.rightAlignPageNumbers = false
        if (value !== ' ') reason ??= '原生目录的自定义页码分隔符尚未接入完整重建。'
      } else {
        const range = /^(\d+)\s*-\s*(\d+)$/.exec(value)
        const minimum = Number(range?.[1]), maximum = Number(range?.[2])
        if (!range || minimum < 1 || maximum > 9 || maximum < minimum) reason ??= '原生目录的标题级别范围无效。'
        else { options.levels = maximum; if (minimum > 1) options.minimumLevel = minimum }
      }
      continue
    }
    reason ??= '原生目录含有尚未接入完整重建的指令开关或条目来源。'
  }
  if (!options.showPageNumbers) options.rightAlignPageNumbers = false
  return { options, hyperlinks, ...(reason ? { reason } : {}) }
}

function headingTargets(model: DocumentModel): (target: string) => DocumentBlock | undefined {
  const heading = (block: DocumentBlock) => block.type === 'heading' && Number.isInteger(block.level) && block.level >= 1 && block.level <= 9
  const direct = new Map(model.blocks.filter(heading).map(block => [block.id, block]))
  const bookmarks = new Map<string, Array<{ block: DocumentBlock; childIndex: unknown }>>()
  for (const block of model.blocks) {
    const source = record(block.metadata.source), details = record(source.details)
    for (const value of Array.isArray(details.bookmarks) ? details.bookmarks : []) {
      const bookmark = record(value)
      if (typeof bookmark.name !== 'string' || !bookmark.name) continue
      const values = bookmarks.get(bookmark.name) ?? []
      values.push({ block, childIndex: bookmark.childIndex })
      bookmarks.set(bookmark.name, values)
    }
  }
  return target => {
    const candidates = bookmarks.get(target)
    // A bookmark takes priority over a same-named paragraph ID, as upstream does.
    if (candidates) return candidates.length === 1 && candidates[0]!.childIndex === 0 && heading(candidates[0]!.block)
      ? candidates[0]!.block : undefined
    return direct.get(target)
  }
}

function pageReference(range: DocumentTextFieldRange): { target?: string; reason?: string } {
  const tokens = instructionTokens(range.field.instruction ?? '')
  if (!tokens || tokens[0]?.toUpperCase() !== 'PAGEREF' || !tokens[1] || tokens[1].startsWith('\\')) return { reason: '原生目录的 PAGEREF 目标尚未解析。' }
  for (let index = 2; index < tokens.length; index++) {
    if (tokens[index]!.toLowerCase() === '\\h') continue
    if (tokens[index] === '\\*' && ['ARABIC', 'MERGEFORMAT'].includes(tokens[index + 1]?.toUpperCase() ?? '')) { index++; continue }
    return { target: tokens[1], reason: '原生 PAGEREF 的位置或数字格式指令尚未接入目录更新。' }
  }
  return { target: tokens[1] }
}

function linkedHeading(block: DocumentBlock, resolve: (target: string) => DocumentBlock | undefined): DocumentBlock | undefined {
  const links = [...new Set(block.runs.map(run => run.metadata?.link).filter((link): link is string => typeof link === 'string' && link.startsWith('#')))]
  return links.length === 1 ? resolve(links[0]!.slice(1)) : undefined
}

function nativeLeader(blocks: readonly DocumentBlock[]): { leader?: DocumentTocTabLeader; reason?: string } {
  const leaders = new Set<DocumentTocTabLeader>()
  for (const block of blocks) {
    const source = record(block.metadata.source), details = record(source.details), style = record(details.style)
    const tabs = Array.isArray(style.tabStops) ? style.tabStops : []
    for (const value of tabs) {
      const tab = record(value)
      if (tab.alignment !== 'right' || tab.leader === undefined) continue
      const leader = tab.leader === 'none' ? 'none' : tab.leader === 'dot' ? 'dots'
        : tab.leader === 'hyphen' ? 'dashes' : tab.leader === 'underscore' ? 'underline' : undefined
      if (!leader) return { reason: '原生目录的制表位前导符尚未接入完整重建。' }
      leaders.add(leader)
    }
  }
  return leaders.size > 1 ? { reason: '原生目录的各行前导符不同，尚未接入完整重建。' }
    : { ...(leaders.size ? { leader: [...leaders][0] } : {}) }
}

/**
 * Called after native body field ranges and heading/bookmark provenance exist.
 * This derives context metadata only: text, fields, source XML, dirty state and
 * history are unchanged. Partial/native instruction-only ranges stay preserved
 * in the report; no headings are inferred by comparing displayed strings.
 */
export function hydrateDocumentTocsFromNativeFields(model: DocumentModel): DocumentNativeTocHydrationReport {
  const report: DocumentNativeTocHydrationReport = { recognized: [], retained: [], unmapped: [] }
  const paragraphs = documentTextFieldParagraphs(model), ranges = documentTextFieldRanges(model)
  const outers = ranges.filter(range => range.field.field.toUpperCase() === 'TOC'), resolve = headingTargets(model)
  for (const outer of outers) {
    const reject = (reason: string) => { report.unmapped.push({ fieldId: outer.field.id, reason }) }
    const start = outer.startParagraph, tail = outer.endParagraph
    if (record(outer.field.metadata?.source).nestedInInstruction === true) { reject('TOC 位于其他字段指令中，尚未映射为正文目录。'); continue }
    if (start.container !== model.blocks || tail.container !== model.blocks || outer.field.startOffset !== 0 || outer.field.endOffset !== tail.text.length) {
      reject('原生 TOC 结果不是完整的顶层段落范围；保留原字段，部分范围的目录交互待接入。'); continue
    }
    const blocks = model.blocks.slice(start.index, tail.index + 1)
    const span = paragraphs.filter(paragraph => paragraph.absoluteStart >= start.absoluteStart && paragraph.absoluteStart <= tail.absoluteStart)
    if (!blocks.length || span.length !== blocks.length || span.some((paragraph, index) => paragraph.container !== model.blocks || paragraph.block.id !== blocks[index]?.id)) {
      reject('原生 TOC 结果跨过表格、对象或非连续文本段落，结构范围尚未接入。'); continue
    }
    if (outers.some(other => other.field.id !== outer.field.id && (other.absoluteStart < outer.absoluteEnd && outer.absoluteStart < other.absoluteEnd
      || other.absoluteStart === other.absoluteEnd && other.absoluteStart === outer.absoluteStart && outer.absoluteStart === outer.absoluteEnd))) {
      reject('原生 TOC 与另一个 TOC 结果重叠，不能分配唯一目录上下文。'); continue
    }
    const existingIds = new Set(blocks.map(documentTocIdAtBlock)), existingId = documentTocIdAtBlock(blocks[0])
    const native = blocks.every(block => record(record(block.metadata.documentToc).native).fieldId === outer.field.id)
    if ([...existingIds].some(id => id !== undefined)) {
      if (existingIds.size !== 1 || !existingId) { reject('原生 TOC 与已有目录元数据只部分重叠，保留原状态等待协调。'); continue }
      if (!native) { report.retained.push({ fieldId: outer.field.id, tocId: existingId, blockIds: blocks.map(block => block.id) }); continue }
    }
    const parsed = nativeOptions(outer.field.instruction ?? ''), leader = nativeLeader(blocks)
    const options = { ...parsed.options, ...(leader.leader ? { tabLeader: leader.leader } : {}) }, tocId = existingId ?? outer.field.id
    const pages = ranges.filter(range => range.field.field.toUpperCase() === 'PAGEREF' && record(range.field.metadata?.source).nestedInInstruction !== true
      && range.absoluteStart >= outer.absoluteStart && range.absoluteEnd <= outer.absoluteEnd)
    let fullReason = parsed.reason ?? leader.reason, pageReason: string | undefined
    if (documentFieldRetainsNestedFields(outer.field)) fullReason ??= '原生目录的相对嵌套字段范围尚未转换。'
    const metadata: DocumentTocMetadata[] = []
    for (const block of blocks) {
      const localPages = pages.filter(range => range.block.id === block.id)
      const page = localPages.length === 1 ? localPages[0] : undefined, reference = page && pageReference(page)
      const linked = linkedHeading(block, resolve), target = reference?.target && resolve(reference.target)
      const heading = target && !blocks.includes(target) ? target : !localPages.length && linked && !blocks.includes(linked) ? linked : undefined
      let rowPageReason: string | undefined
      if (localPages.some(range => !documentFieldLocked(range.field))) {
        if (localPages.length !== 1 || page?.endParagraph.block.id !== block.id) rowPageReason = '原生目录行包含多个或跨段 PAGEREF，目标范围尚未接入。'
        else rowPageReason = reference?.reason ?? (!heading ? '原生目录的 PAGEREF 尚未关联到唯一的段首标题书签。' : undefined)
      }
      pageReason ??= rowPageReason
      fullReason ??= reference?.reason ?? (!heading ? '原生目录结果含有尚未关联到真实标题的段落，不能安全重建。' : undefined)
      const text = page ? runText(block.runs).slice(page.field.startOffset, page.field.endOffset) : ''
      const number = /^\d+$/.test(text) && Number.isSafeInteger(Number(text)) && Number(text) > 0 ? Number(text) : undefined
      metadata.push({ id: tocId, role: 'entry', options: { ...options },
        ...(heading ? { headingId: heading.id, headingText: runText(heading.runs).replace(/[\r\n\t]/g, ' ').trim(), level: heading.level } : {}),
        ...(number !== undefined ? { pageNumber: number } : {}),
        native: { fieldId: outer.field.id, instruction: outer.field.instruction ?? '', hyperlinks: parsed.hyperlinks || linked !== undefined } })
    }
    const provenance: Omit<DocumentNativeTocMetadata, 'hyperlinks'> = { fieldId: outer.field.id,
      ...(fullReason ? { fullUpdateUnavailableReason: fullReason } : {}), ...(pageReason ? { pageUpdateUnavailableReason: pageReason } : {}) }
    for (const [index, block] of blocks.entries()) {
      const next = metadata[index]!
      next.native = { ...next.native!, ...provenance }
      if (JSON.stringify(block.metadata.documentToc) !== JSON.stringify(next)) block.metadata = { ...block.metadata, documentToc: next }
    }
    report.recognized.push({ fieldId: outer.field.id, tocId, blockIds: blocks.map(block => block.id),
      ...(fullReason ? { fullUpdateUnavailableReason: fullReason } : {}), ...(pageReason ? { pageUpdateUnavailableReason: pageReason } : {}) })
  }
  return report
}
