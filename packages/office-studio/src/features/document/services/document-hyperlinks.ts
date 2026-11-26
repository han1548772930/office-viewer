import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { blockAtPoint, documentBlockAncestry, documentId, orderedSelection, runsForTarget, safeUrl, sameTarget,
  type DocumentSelection } from '../model/document-model'
import { formatDocumentRunRange } from '../model/document-run-range'
import { documentTextFieldRanges } from '../model/document-text-fields'
import { assertDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../permissions/document-permissions'
import { flattenDocumentBody } from './document-facade'
import { assertDocumentTextSelectionEditable } from './document-text-formatting'

export type DocumentHyperlinkCommand =
  | { id: 'docs.command.add-hyper-link'; payload: string }
  | { id: 'docs.command.update-hyper-link'; linkId: string; payload: string; label: string }
  | { id: 'docs.command.delete-hyper-link'; linkId: string }

export interface DocumentHyperlink {
  id: string
  target: string
  label: string
  selection: DocumentSelection
}

export interface DocumentHyperlinkSelectionState {
  available: boolean
  reason?: string
  link?: DocumentHyperlink
}

interface HyperlinkSource extends Record<string, unknown> {
  rangeId?: string
  properties?: Record<string, unknown>
}

interface HyperlinkBodySource {
  body?: {
    customRanges?: Array<{ rangeType: number; rangeId: string; properties?: { bookmarkId?: string } }>
    paragraphs?: Array<{ paragraphId?: string; paragraphStyle?: { headingId?: string } }>
  }
}

export function isDocumentHyperlinkCommand(command: { id: string }): command is DocumentHyperlinkCommand {
  return ['docs.command.add-hyper-link', 'docs.command.update-hyper-link', 'docs.command.delete-hyper-link'].includes(command.id)
}

/** UI and API submit raw payloads; only the editing command normalizes them. */
export function normalizeDocumentHyperlinkTarget(value: string): string {
  if (typeof value !== 'string' || !value.trim() || /[\u0000-\u001f]/.test(value)) throw new TypeError('请输入有效的链接地址。')
  const input = value.trim()
  if (input.startsWith('#')) {
    if (input.length === 1) throw new TypeError('文档内链接需要指定书签或标题。')
    return input
  }
  const target = /^[a-z][a-z\d+.-]*:/i.test(input) ? input
    : /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(input) ? `mailto:${input}` : `https://${input}`
  if (!safeUrl(target)) throw new TypeError('链接地址的协议不受支持。')
  const url = new URL(target)
  if (url.protocol === 'mailto:' && !url.pathname && !url.hostname) throw new TypeError('请输入有效的邮件地址。')
  return target
}

/** Range identity lives in existing run provenance, including imported mixed runs. */
export function documentHyperlinks(model: DocumentModel): DocumentHyperlink[] {
  const body = flattenDocumentBody(model)
  const links: Array<DocumentHyperlink & { sourceId?: string; startOffset: number; endOffset: number }> = []
  let previous: (typeof links)[number] | undefined
  for (const segment of body.segments) {
    let offset = 0
    for (const run of runsForTarget(model, segment)) {
      const from = offset
      offset += run.text.length
      const properties = run.metadata?.link as Record<string, unknown> | undefined
      const target = run.href ?? properties?.target
      if (!run.text.length) continue
      if (typeof target !== 'string' || !target) { previous = undefined; continue }
      const source = run.metadata?.univerBodyHyperlinkSource as HyperlinkSource | undefined
      const sourceId = source?.rangeId
      const startOffset = segment.startOffset + from, endOffset = segment.startOffset + offset
      if (previous && previous.target === target && previous.sourceId === sourceId
        && /^\n*$/.test(body.text.slice(previous.endOffset, startOffset))) {
        previous.endOffset = endOffset
        previous.selection.focus = { blockId: segment.blockId, row: segment.row, col: segment.col, cellBlockId: segment.cellBlockId, offset }
        continue
      }
      const baseId = sourceId ?? `hyperlink:${run.id}`
      const link = { id: links.some(link => link.id === baseId) ? `${baseId}:${run.id}` : baseId, target, label: '', sourceId, startOffset, endOffset,
        selection: { anchor: { blockId: segment.blockId, row: segment.row, col: segment.col, cellBlockId: segment.cellBlockId, offset: from },
          focus: { blockId: segment.blockId, row: segment.row, col: segment.col, cellBlockId: segment.cellBlockId, offset } } }
      links.push(link)
      previous = link
    }
  }
  return links.map(link => ({ id: link.id, target: link.target, label: body.text.slice(link.startOffset, link.endOffset), selection: link.selection }))
}

export function documentHyperlinkForSelection(model: DocumentModel, selection: DocumentSelection): DocumentHyperlink | undefined {
  if (selection.tableSelection) return undefined
  const body = flattenDocumentBody(model), range = orderedSelection(model, selection)
  const first = body.segments.find(segment => sameTarget(segment, range.start)), last = body.segments.find(segment => sameTarget(segment, range.end))
  if (!first || !last) return undefined
  const start = first.startOffset + range.start.offset, end = last.startOffset + range.end.offset
  return documentHyperlinks(model).find(candidate => {
    const from = body.segments.find(segment => sameTarget(segment, candidate.selection.anchor))!
    const to = body.segments.find(segment => sameTarget(segment, candidate.selection.focus))!
    return from.startOffset + candidate.selection.anchor.offset <= start && to.startOffset + candidate.selection.focus.offset >= end
  })
}

function assertDocumentHyperlinkEditingAllowed(model: DocumentModel, selection: DocumentSelection): void {
  const body = flattenDocumentBody(model), range = orderedSelection(model, selection)
  const first = body.segments.find(segment => sameTarget(segment, range.start))!, last = body.segments.find(segment => sameTarget(segment, range.end))!
  const start = first.startOffset + range.start.offset, end = last.startOffset + range.end.offset, fields = documentTextFieldRanges(model)
  const generated = documentHyperlinks(model).some(link => {
    const from = body.segments.find(segment => sameTarget(segment, link.selection.anchor))!, to = body.segments.find(segment => sameTarget(segment, link.selection.focus))!
    const linkStart = from.startOffset + link.selection.anchor.offset, linkEnd = to.startOffset + link.selection.focus.offset
    if (start >= linkEnd || end <= linkStart) return false
    if (fields.some(field => field.absoluteStart <= linkStart && field.absoluteEnd >= linkEnd
      && ['TOC', 'INDEX', 'TOA', 'REF', 'PAGEREF', 'NOTEREF', 'CITATION', 'BIBLIOGRAPHY'].includes(field.field.field.toUpperCase()))) return true
    return body.segments.filter(segment => segment.startOffset < linkEnd && segment.endOffset > linkStart).some(segment => {
      let offset = segment.startOffset
      return runsForTarget(model, segment).some(run => {
        const from = offset
        offset += run.text.length
        if (from >= linkEnd || offset <= linkStart) return false
        const properties = run.metadata?.link as Record<string, unknown> | undefined
        const source = run.metadata?.univerBodyHyperlinkSource as HyperlinkSource | undefined
        return (properties?.referenceFieldLink ?? source?.properties?.referenceFieldLink) === true
      })
    })
  })
  if (generated) throw new RangeError('引用字段生成的链接需通过所属字段更新。')
}

export function documentHyperlinkSelectionState(model: DocumentModel, selection: DocumentSelection,
  permissionContext: DocumentPermissionContext = { documentId: model.id ?? '', userId: 'local-user' }): DocumentHyperlinkSelectionState {
  if (selection.tableSelection) return { available: false, reason: '请选择文字以添加或编辑链接。' }
  const body = flattenDocumentBody(model), range = orderedSelection(model, selection)
  const first = body.segments.find(segment => sameTarget(segment, range.start)), last = body.segments.find(segment => sameTarget(segment, range.end))
  if (!first || !last) return { available: false, reason: '当前内容没有可编辑的文字。' }
  const start = first.startOffset + range.start.offset, end = last.startOffset + range.end.offset
  const link = documentHyperlinkForSelection(model, selection)
  if (!link && !body.segments.some(segment => segment.startOffset < end && segment.endOffset > start)) {
    return { available: false, reason: '请选择需要添加链接的文字。' }
  }
  try {
    const target = link?.selection ?? selection
    assertDocumentPermission(model.permissionRules ?? [], permissionContext, 'edit', permissionTargetsForSelection(model, target))
    assertDocumentTextSelectionEditable(model, target)
    assertDocumentHyperlinkEditingAllowed(model, target)
    return { available: true, ...(link ? { link } : {}) }
  } catch (error) {
    return { available: false, reason: error instanceof Error ? error.message : String(error), ...(link ? { link } : {}) }
  }
}

function hyperlinkTargetProperties(target: string, model: DocumentModel, previous: Record<string, unknown>): Record<string, unknown> {
  const properties = { ...previous }
  for (const key of ['url', 'bookmarkId', 'headingId', 'anchor', 'target']) delete properties[key]
  if (!target.startsWith('#')) return { ...properties, url: target }
  const id = target.slice(1), body = (model.metadata?.univerBodySource as HyperlinkBodySource | undefined)?.body
  const bookmark = body?.customRanges?.some(range => range.rangeType === 3 && (range.rangeId === id || range.properties?.bookmarkId === id))
  const heading = !bookmark && (body?.paragraphs?.some(paragraph => paragraph.paragraphId === id || paragraph.paragraphStyle?.headingId === id)
    || flattenDocumentBody(model).segments.some(segment => (segment.cellBlockId ?? segment.blockId) === id))
  return { ...properties, ...(previous.headingId === id || heading ? { headingId: id } : { bookmarkId: id }) }
}

export function executeDocumentHyperlinkCommand(model: DocumentModel, selection: DocumentSelection, command: DocumentHyperlinkCommand,
  permissionContext: DocumentPermissionContext, replaceText: (model: DocumentModel, selection: DocumentSelection, text: string) => DocumentSelection,
  targetModel: DocumentModel = model): DocumentSelection {
  if (selection.tableSelection) throw new RangeError('链接命令需要文字选区。')
  const target = command.id === 'docs.command.delete-hyper-link' ? undefined : normalizeDocumentHyperlinkTarget(command.payload)
  if (command.id !== 'docs.command.add-hyper-link' && (typeof command.linkId !== 'string' || !command.linkId)) throw new TypeError('链接标识无效。')
  if (command.id === 'docs.command.update-hyper-link' && (typeof command.label !== 'string' || !command.label)) throw new TypeError('链接显示文字不能为空。')
  const link = command.id === 'docs.command.add-hyper-link' ? undefined : documentHyperlinks(model).find(candidate => candidate.id === command.linkId)
  if (command.id !== 'docs.command.add-hyper-link' && !link) throw new RangeError('需要编辑的链接已不存在。')
  const range = link?.selection ?? selection
  const replaceLabel = command.id === 'docs.command.update-hyper-link' && command.label !== link!.label
  assertDocumentPermission(model.permissionRules ?? [], permissionContext, 'edit', permissionTargetsForSelection(model, range))
  if (!replaceLabel) assertDocumentTextSelectionEditable(model, range)
  assertDocumentHyperlinkEditingAllowed(model, range)
  const candidate = replaceLabel ? cloneModel(model) : model
  const result = candidate === model ? selection : replaceText(candidate, range, command.id === 'docs.command.update-hyper-link' ? command.label : '')
  const editedRange = candidate === model ? range : { anchor: range.anchor, focus: result.focus }
  const body = flattenDocumentBody(candidate), ordered = orderedSelection(candidate, editedRange)
  const first = body.segments.find(segment => sameTarget(segment, ordered.start))!, last = body.segments.find(segment => sameTarget(segment, ordered.end))!
  const start = first.startOffset + ordered.start.offset, end = last.startOffset + ordered.end.offset
  const affected = body.segments.filter(segment => segment.startOffset < end && segment.endOffset > start)
  if (!affected.length) throw new RangeError('请选择需要添加链接的文字。')
  const id = link?.id ?? documentId('hyperlink')
  const updates = affected.map(segment => ({ segment, runs: formatDocumentRunRange(runsForTarget(candidate, segment),
    Math.max(0, start - segment.startOffset), Math.min(segment.text.length, end - segment.startOffset), run => {
      const metadata = { ...run.metadata }
      if (command.id === 'docs.command.delete-hyper-link') {
        delete metadata.link
        delete metadata.href
        delete metadata.univerBodyHyperlinkSource
        return { href: undefined, metadata }
      }
      const source = run.metadata?.univerBodyHyperlinkSource as HyperlinkSource | undefined
      const previous = command.id === 'docs.command.update-hyper-link'
        ? { ...source?.properties, ...(run.metadata?.link as Record<string, unknown> | undefined) } : {}
      const properties = hyperlinkTargetProperties(target!, targetModel, previous)
      delete metadata.href
      metadata.link = { ...properties, target }
      metadata.univerBodyHyperlinkSource = { ...(command.id === 'docs.command.update-hyper-link' ? source : {}), rangeId: id, rangeType: 0, properties }
      return { href: target, metadata }
    }) }))
  for (const update of updates) {
    const block = blockAtPoint(candidate, update.segment)!
    block.runs = update.runs
    delete block.metadata.html
    for (const ancestor of documentBlockAncestry(candidate, block.id) ?? []) delete ancestor.metadata.html
  }
  if (candidate !== model) Object.assign(model, candidate)
  return result
}
