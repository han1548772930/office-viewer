import type { DocumentContentControl } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { applyRunPatch, runText, sliceRuns, type DocumentRun } from '../features/document/model/document-model'
import { documentTextFields, setDocumentTextFields } from '../features/document/model/document-text-fields'
import { isDocumentTextBoundary } from '../features/document/model/document-story'
import { readDocumentContentControlSourceOrder } from '../features/document/model/content-control-source-order'
import { bodyBoolean, bodyRecord } from './document-body-style'
import { documentBodyOfficeSource, documentBodySourceMetadata } from './document-body-source'
import { applyDocumentBodyControlOwnership } from './document-body-control-ranges'
import { documentBodyNoteReferenceCount } from './document-body-note-boundaries'
import { isDocumentBodyLatexRange } from './document-body-latex'
import { getLatexFormula, latexFormulaRangeId } from '../features/document/latex/latex-formula'
import { DOCUMENT_BODY_TOKENS as T, documentBodyArray, documentBodyIdentity, documentBodyLocalOffset, documentBodyOffset,
  type DocumentBodyDecodeOptions, type DocumentBodyParagraphPosition, type DocumentBodyStructureIndex } from './document-body-types'

const controlTypes = new Set<string>(['richText', 'text', 'checkbox', 'date', 'dropDownList', 'comboBox', 'picture', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj'])
const placements = new Set<string>(['inline', 'block', 'table', 'row', 'cell', 'header', 'footer'])
const locks = new Set<string>(['unlocked', 'controlLocked', 'contentLocked', 'sdtLocked', 'sdtContentLocked'])

/** Range endpoints include sentinels in the public stream, but never in model text. */
export function applyDocumentBodyRanges(value: unknown, positions: readonly DocumentBodyParagraphPosition[], diagnostics: string[],
  options: Pick<DocumentBodyDecodeOptions, 'note'> = {}, structures: DocumentBodyStructureIndex = { blocks: [], rows: [], cells: [] },
  modeledRangeIds: Set<string> = new Set()): DocumentContentControl[] {
  const body = bodyRecord(value), stream = typeof body.dataStream === 'string' ? body.dataStream : '', controls: DocumentContentControl[] = [], ids = new Set<string>()
  for (const [sourceOrder, raw] of documentBodyArray(body.customRanges, 'customRanges').entries()) {
    const start = documentBodyOffset(raw.startIndex, stream.length, 'customRanges.startIndex'), inclusiveEnd = documentBodyOffset(raw.endIndex, stream.length - 1, 'customRanges.endIndex')
    if (!isDocumentTextBoundary(stream, start) || !isDocumentTextBoundary(stream, inclusiveEnd + 1)) throw new RangeError('文档范围不能截断 UTF-16 字符。')
    if (inclusiveEnd < start) throw new RangeError('文档自定义范围结束位置早于起点。')
    const id = documentBodyIdentity(raw.rangeId, '文档范围')
    if (ids.has(id)) throw new TypeError(`文档范围标识重复：${id}`)
    ids.add(id)
    const from = stream[start] === T.rangeStart ? start + 1 : start, to = stream[inclusiveEnd] === T.rangeEnd ? inclusiveEnd : inclusiveEnd + 1
    const affected = positions.filter(position => from === to ? position.start <= from && position.end >= from
      : position.start < to && position.end > from || position.start === position.end && position.start >= from && position.start < to)
    const first = affected[0], last = affected.at(-1)
    if (!first || !last) { diagnostics.push(`自定义范围 ${id} 仅覆盖结构标记，其原始定义已保留。`); continue }
    const startOffset = documentBodyLocalOffset(first, Math.max(first.start, from)), endOffset = documentBodyLocalOffset(last!, Math.min(last!.end, to), 'end')
    const startNoteReferenceCount = documentBodyNoteReferenceCount(first, Math.max(first.start, from))
    const endNoteReferenceCount = documentBodyNoteReferenceCount(last!, Math.min(last!.end, to))
    const properties = bodyRecord(raw.properties), rangeType = raw.rangeType
    if (isDocumentBodyLatexRange(raw)) {
      const runs = first.block.runs as DocumentRun[]
      if (first !== last || endOffset - startOffset !== 1 || !runs.some(run => getLatexFormula(run) && latexFormulaRangeId(run) === id)) throw new RangeError('公式公开范围未映射到一个完整的模型实体。')
      modeledRangeIds.add(id)
    } else if (rangeType === 0) {
      const url = typeof properties.url === 'string' ? properties.url : typeof properties.anchor === 'string' ? `#${properties.anchor}` : undefined
      if (!url) { diagnostics.push(`超链接范围 ${id} 缺少可导航目标，定义已保留。`); continue }
      for (const position of affected) {
        const begin = documentBodyLocalOffset(position, Math.max(position.start, from)), end = documentBodyLocalOffset(position, Math.min(position.end, to), 'end')
        if (end <= begin) continue
        const runs = position.block.runs as DocumentRun[], selected = sliceRuns(runs, begin, end).map(run => applyRunPatch(run, {
          href: url, metadata: { ...run.metadata, link: { ...properties, target: url }, univerBodyHyperlinkSource: cloneModel(raw) },
        }))
        position.block.runs = [...sliceRuns(runs, 0, begin), ...selected, ...sliceRuns(runs, end)]
        modeledRangeIds.add(id)
      }
    } else if (rangeType === 1) {
      if (properties.fieldType === 'FOOTNOTE_SEPARATOR' || properties.fieldType === 'FOOTNOTE_CONTINUATION_SEPARATOR') {
        first.block.metadata.noteSeparator = { version: 1, widthFraction: properties.fieldType === 'FOOTNOTE_SEPARATOR' ? 1 / 3 : 1 }
        first.block.metadata.univerNoteSeparatorRangeSource = cloneModel(raw)
        modeledRangeIds.add(id)
        continue
      }
      if (typeof properties.instruction !== 'string' || !properties.instruction.trim()) throw new TypeError('文档字段必须包含指令。')
      const field = typeof properties.fieldType === 'string' ? properties.fieldType.toUpperCase() : properties.instruction.trim().split(/\s+/)[0]!.toUpperCase()
      setDocumentTextFields(first.block, [...documentTextFields(first.block), { id, field, instruction: properties.instruction, startOffset, endOffset,
        ...(startNoteReferenceCount === undefined ? {} : { startNoteReferenceCount }), ...(endNoteReferenceCount === undefined ? {} : { endNoteReferenceCount }),
        ...(first !== last ? { endBlockId: last!.block.id } : {}),
        ...(typeof properties.cachedResult === 'string' ? { cachedResult: properties.cachedResult } : {}), locked: bodyBoolean(properties.locked),
        containsNestedFields: documentBodyArray(body.customRanges, 'customRanges').some(other => other.rangeType === 1 && other.rangeId !== id
          && typeof other.startIndex === 'number' && typeof other.endIndex === 'number' && other.startIndex > start && other.endIndex < inclusiveEnd),
        metadata: { ...documentBodySourceMetadata(documentBodyOfficeSource(raw, '字段').metadata, '字段', ['univerBodyFieldSource']),
          univerBodyFieldSource: cloneModel(raw), source: { properties: cloneModel(properties),
          ...(properties.instructionParts === undefined ? {} : { instructionParts: cloneModel(properties.instructionParts) }),
          ...(properties.instructionTextRuns === undefined ? {} : { instructionTextRuns: cloneModel(properties.instructionTextRuns) }) } } }])
      modeledRangeIds.add(id)
    } else if (rangeType === 2) {
      if (typeof properties.kind !== 'string' || !controlTypes.has(properties.kind) || typeof properties.placement !== 'string' || !placements.has(properties.placement)) {
        throw new TypeError('文档内容控件类型或位置无效。')
      }
      if (properties.lock !== undefined && (typeof properties.lock !== 'string' || !locks.has(properties.lock))) throw new TypeError('文档内容控件锁无效。')
      const checkbox = bodyRecord(properties.checkbox), date = bodyRecord(properties.date), text = bodyRecord(properties.text)
      const retainedSourceOrder = documentBodySourceMetadata(documentBodyOfficeSource(raw, '内容控件').metadata, '内容控件').univerBodySdtOrder
      if (retainedSourceOrder !== undefined && readDocumentContentControlSourceOrder(retainedSourceOrder) === undefined) throw new TypeError('内容控件来源顺序无效。')
      const control: DocumentContentControl = { id, type: properties.kind as DocumentContentControl['type'], placement: properties.placement as DocumentContentControl['placement'],
        blockId: first.block.id, startOffset, endOffset: first === last ? endOffset : runText(first.block.runs).length,
        ...(startNoteReferenceCount === undefined ? {} : { startNoteReferenceCount }), ...(first !== last || endNoteReferenceCount === undefined ? {} : { endNoteReferenceCount }),
        ...(properties.lock === undefined ? {} : { lock: properties.lock as DocumentContentControl['lock'] }),
        // Keep the imported ordering when a public save sorts its customRanges.
        metadata: { univerBodySdtSource: cloneModel(raw), univerBodySdtOrder: retainedSourceOrder ?? { version: 1, order: sourceOrder } } }
      for (const key of ['alias', 'title', 'tag', 'appearance', 'color', 'placeholder', 'value', 'rawSdtPr', 'rawSdtEndPr'] as const) {
        if (properties[key] !== undefined) {
          const current = properties[key]; if (typeof current !== 'string') throw new TypeError(`内容控件 ${key} 必须为字符串。`)
          control[key] = current
        }
      }
      for (const key of ['showingPlaceholder', 'temporary'] as const) { const current = bodyBoolean(properties[key]); if (current !== undefined) control[key] = current }
      if (properties.checkbox !== undefined) { const checked = bodyBoolean(checkbox.checked) ?? false; control.checked = checked; control.checkbox = { ...checkbox, checked } }
      if (text.multiLine !== undefined) control.multiLine = bodyBoolean(text.multiLine)
      for (const [source, target] of [['fullDate', 'dateValue'], ['format', 'dateFormat'], ['locale', 'dateLocale'], ['calendar', 'dateCalendar'], ['storage', 'dateStorage']] as const) {
        const current = date[source]; if (current !== undefined) { if (typeof current !== 'string') throw new TypeError('内容控件日期属性必须为字符串。'); control[target] = current }
      }
      if (properties.listItems !== undefined) control.items = documentBodyArray(properties.listItems, 'listItems').map(item => {
        if (typeof item.value !== 'string' || item.displayText !== undefined && typeof item.displayText !== 'string') throw new TypeError('内容控件列表项无效。')
        return { value: item.value, ...(typeof item.displayText === 'string' ? { displayText: item.displayText } : {}) }
      })
      for (const key of ['placeholderDocPart', 'repeatingSection', 'docPart', 'dataBinding'] as const) {
        if (properties[key] !== undefined) Object.assign(control, { [key]: cloneModel(bodyRecord(properties[key])) })
      }
      applyDocumentBodyControlOwnership(control, raw, from, to, positions, structures, stream)
      if (first !== last && !control.blockRange && !control.tableRowRange) throw new RangeError('跨段内容控件必须拥有真实的完整连续块。')
      controls.push(control)
      modeledRangeIds.add(id)
    } else if (rangeType === 8 || rangeType === 9) {
      const noteId = documentBodyIdentity(properties.noteId, '注释范围'), kind = rangeType === 8 ? 'footnote' : 'endnote'
      if (options.note?.kind === kind && (options.note.id === undefined || options.note.id === noteId)) continue
      if (first !== last || endOffset <= startOffset) throw new RangeError('注释引用必须位于一个文本原子中。')
      const runs = first.block.runs as DocumentRun[]
      const selected = sliceRuns(runs, startOffset, endOffset)
      if (selected.length !== 1) throw new RangeError('注释引用不能跨越多个文本运行。')
      selected[0]!.metadata = { ...selected[0]!.metadata, noteReference: { kind, noteId }, univerBodyNoteRangeSource: cloneModel(raw) }
      first.block.runs = [...sliceRuns(runs, 0, startOffset), ...selected, ...sliceRuns(runs, endOffset)]
      modeledRangeIds.add(id)
    } else {
      diagnostics.push(`自定义范围 ${id} 的 ${String(rangeType)} 类型已保留，尚未接入全部插件行为。`)
    }
  }
  // Parent identity is derived from the actual containing range, never array order.
  for (const control of controls) {
    if (control.parentId) continue
    const own = bodyRecord(control.metadata?.univerBodySdtSource), containing = controls.filter(parent => {
      const raw = bodyRecord(parent.metadata?.univerBodySdtSource)
      return parent !== control && ['group', 'repeatingSection', 'repeatingSectionItem'].includes(parent.type)
        && typeof raw.startIndex === 'number' && typeof raw.endIndex === 'number'
        && typeof own.startIndex === 'number' && typeof own.endIndex === 'number' && raw.startIndex <= own.startIndex && raw.endIndex >= own.endIndex
        && (raw.startIndex < own.startIndex || raw.endIndex > own.endIndex)
    }).sort((a, b) => Number(bodyRecord(a.metadata?.univerBodySdtSource).endIndex) - Number(bodyRecord(a.metadata?.univerBodySdtSource).startIndex)
      - Number(bodyRecord(b.metadata?.univerBodySdtSource).endIndex) + Number(bodyRecord(b.metadata?.univerBodySdtSource).startIndex))
    if (containing[0]) control.parentId = containing[0].id
  }
  for (const control of controls) if (control.parentId) {
    const parent = controls.find(parent => parent.id === control.parentId), own = bodyRecord(control.metadata?.univerBodySdtSource)
    const ancestor = bodyRecord(parent?.metadata?.univerBodySdtSource), seen = new Set([control.id])
    if (!parent || !['group', 'repeatingSection', 'repeatingSectionItem'].includes(parent.type)
      || Number(ancestor.startIndex) > Number(own.startIndex) || Number(ancestor.endIndex) < Number(own.endIndex)) throw new RangeError('内容控件父身份不包含其真实公开范围。')
    let current: DocumentContentControl | undefined = control
    while (current?.parentId) {
      if (seen.has(current.parentId)) throw new RangeError('内容控件父身份不能形成循环。')
      seen.add(current.parentId); current = controls.find(parent => parent.id === current!.parentId)
    }
  }
  return controls
}
