import type { DocumentDefaultStyle, DocumentModel, DocumentNote } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/thread-comments'
import { documentEditingBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentDefaultStyle } from '../model/document-default-style'
import { blockAtPoint, collapsedSelection, createRun, documentBlockAncestry, documentId, marksAtPoint, orderedSelection,
  runText, runsForTarget, sameTarget, sliceRuns, textTargets, type DocumentRun, type DocumentSelection } from '../model/document-model'
import { documentRunHasNoteReference } from '../model/document-note-reference'
import { documentNoteReferenceOffsets } from '../model/document-note-boundary'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentTextFieldParagraphs, transformDocumentTextFieldsForBlockEdit } from '../model/document-text-fields'
import { applyDocumentBlockHostModel, documentModelForBlockHost, documentPointAtBlockHostOffset } from '../services/document-block-host'
import { replaceDocumentRange } from '../services/document-facade'
import { assertTextRangeEditable, transformContentControlsForBlockEdit, validateDocumentContentControlRanges } from '../services/content-controls'
import { assertDocumentRepeatingNoteContentRemovable, removeDeletedDocumentRepeatingNoteScopes,
  visitDocumentRepeatingContentControlBlocks } from '../services/repeating-content-control-copy'
import { appendDocumentNoteContentControls, assertDocumentNoteSerializable, documentNoteContentText, prepareDocumentNoteContentWithControls } from './document-note-content'
import { documentNoteReferenceRecords } from './document-note-references'
import { renumberDocumentNotes } from './document-note-reference-sync'
import { setDocumentNoteSettings } from './document-note-settings'
import type { DocumentNoteInsertOptions, DocumentNoteKind, DocumentNotePatch } from './document-note-types'

export function documentSupportsNotes(model: DocumentModel): boolean {
  const record = (value: unknown): Record<string, unknown> => value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
  const flavor = model.metadata?.documentFlavor ?? record(model.metadata?.documentStyle).documentFlavor
    ?? record(record(model.metadata?.source).documentStyle).documentFlavor
  return flavor !== 2 && !(typeof flavor === 'string' && flavor.toLowerCase() === 'modern')
}

function assertNoteKind(kind: DocumentNoteKind): void {
  if (kind !== 'footnote' && kind !== 'endnote') throw new TypeError('注释类型无效。')
}

function assertText(text: string): void {
  if (typeof text !== 'string') throw new TypeError('注释文字无效。')
  if (text.length * 2 > 32 * 1024 * 1024) throw new RangeError('注释内容超过文本预算。')
}

function assertReferenceStyle(style: DocumentDefaultStyle['text']): void {
  assertDocumentNoteSerializable(style)
  if (!style || typeof style !== 'object' || Array.isArray(style) || typeof style.fontFamily !== 'string' || !style.fontFamily.trim()
    || !Number.isFinite(style.fontSize) || style.fontSize <= 0 || style.fontSize > 10_000 || typeof style.color !== 'string'
    || ['bold', 'italic', 'underline', 'strike'].some(key => typeof Reflect.get(style, key) !== 'boolean')) throw new TypeError('注释引用的字符样式无效。')
  const allowed = ['bold', 'italic', 'underline', 'strike', 'color', 'fontFamily', 'eastAsiaFontFamily', 'fontSize', 'verticalAlign', 'highlight']
  if (Object.keys(style).some(key => !allowed.includes(key))) throw new TypeError('注释引用样式不能包含字符身份或来源属性。')
  for (const key of ['fontFamily', 'eastAsiaFontFamily'] as const) if (style[key] !== undefined && (!style[key]!.trim() || style[key]!.includes('\0'))) throw new TypeError('注释引用字体名称无效。')
  if (style.verticalAlign !== undefined && !['superscript', 'subscript', 'baseline'].includes(style.verticalAlign)) throw new TypeError('注释引用的垂直对齐无效。')
  documentDefaultStyle({ defaultStyle: { version: 1, text: style, paragraph: documentDefaultStyle().paragraph } })
}

export function validateDocumentNotePatch(patch: DocumentNotePatch): void {
  assertDocumentNoteSerializable(patch)
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) throw new TypeError('注释修改需要对象参数。')
  for (const key of Object.keys(patch)) if (!['text', 'content', 'contentControls', 'customMark', 'referenceTextStyle', 'showBodyReference', 'metadata'].includes(key)) throw new TypeError('未知的注释属性：' + key)
  if (patch.contentControls !== undefined && (!Array.isArray(patch.contentControls) || patch.content === undefined)) throw new TypeError('注释内容控件需要与其来源 content 一起提供。')
  if (patch.text !== undefined) assertText(patch.text)
  if (patch.customMark !== undefined && patch.customMark !== null && (typeof patch.customMark !== 'string' || patch.customMark.length > 4096 || /[\r\n\u0000]/.test(patch.customMark))) throw new TypeError('自定义注释标记需要单行文字。')
  if (patch.referenceTextStyle !== undefined && patch.referenceTextStyle !== null) assertReferenceStyle(patch.referenceTextStyle)
  if (patch.showBodyReference !== undefined && typeof patch.showBodyReference !== 'boolean') throw new TypeError('注释正文引用显隐无效。')
  if (patch.metadata !== undefined && patch.metadata !== null && (typeof patch.metadata !== 'object' || Array.isArray(patch.metadata))) throw new TypeError('注释元数据无效。')
}

function atomicNoteEdit<T>(model: DocumentModel, edit: (candidate: DocumentModel) => T): T {
  const candidate = cloneModel(model), result = edit(candidate)
  for (const key of Object.keys(model)) if (!Object.hasOwn(candidate, key)) Reflect.deleteProperty(model, key)
  Object.assign(model, candidate)
  return result
}

function requireNote(model: DocumentModel, kind: DocumentNoteKind, id: string): DocumentNote {
  assertNoteKind(kind)
  if (typeof id !== 'string' || !id.trim()) throw new TypeError('注释标识无效。')
  const note = (kind === 'footnote' ? model.footnotes : model.endnotes)?.[id]
  if (!note || note.id !== id) throw new RangeError('注释已不存在：' + id)
  return note
}

function blockTextLock(model: DocumentModel, blockId: string, start: number, end: number): void {
  const paragraph = documentTextFieldParagraphs(model).find(paragraph => paragraph.block.id === blockId)
  if (paragraph) assertTextRangeEditable(model, paragraph.absoluteStart + start, paragraph.absoluteStart + end)
}

/** Read-only capability check shared by the Ribbon and insertion planner. */
export function assertDocumentNoteInsertion(model: DocumentModel, selection: DocumentSelection): void {
  if (!documentSupportsNotes(model) || documentEditingBlockHost(model)) throw new RangeError('脚注和尾注引用需要在传统文档正文中插入。')
  if (!selection || selection.tableSelection) throw new RangeError('脚注或尾注需要正文或单元格内的文字插入位置。')
  const targets = textTargets(model)
  for (const point of [selection.anchor, selection.focus]) {
    if (!point || !targets.some(target => sameTarget(target, point)) || !Number.isSafeInteger(point.offset) || point.offset < 0) throw new RangeError('注释引用的文字位置无效。')
    const text = runText(runsForTarget(model, point))
    if (point.offset > text.length || !isDocumentTextBoundary(text, point.offset)) throw new RangeError('注释引用需要合法的 UTF-16 文字边界。')
  }
  const point = orderedSelection(model, selection).end, block = blockAtPoint(model, point)!
  blockTextLock(model, block.id, point.offset, point.offset)
  let offset = 0
  for (const run of block.runs) {
    if (documentRunHasNoteReference(run) && point.offset > offset && point.offset < offset + run.text.length) throw new RangeError('不能在注释引用字符内部插入新引用。')
    offset += run.text.length
  }
}

export function insertDocumentNote(model: DocumentModel, selection: DocumentSelection, kind: DocumentNoteKind,
  options: DocumentNoteInsertOptions = {}): { selection: DocumentSelection; note: DocumentNote; host: DocumentAuxiliaryBlockHost } {
  assertNoteKind(kind); assertDocumentNoteSerializable(options)
  const { noteId, settings, sectionId, suppressEndnotes, decorationContentControls, ...properties } = options
  if (properties.customMark === null || properties.referenceTextStyle === null || properties.metadata === null) throw new TypeError('插入注释不能使用属性删除参数。')
  validateDocumentNotePatch(properties)
  assertDocumentNoteInsertion(model, selection)
  return atomicNoteEdit(model, candidate => {
    // Upstream inserts at the logical end of a nonempty selection without replacing it.
    const point = orderedSelection(candidate, selection).end, block = blockAtPoint(candidate, point)!
    const id = noteId ?? documentId(kind)
    if (typeof id !== 'string' || !id.trim() || id.length > 4096 || Object.hasOwn(kind === 'footnote' ? candidate.footnotes ?? {} : candidate.endnotes ?? {}, id)) throw new RangeError('注释标识为空、重复或超过预算。')
    const prepared = prepareDocumentNoteContentWithControls(candidate, kind, id, properties.content, properties.text ?? '', true, properties.contentControls)
    const content = prepared.content
    if (properties.content !== undefined && properties.text !== undefined && properties.text !== documentNoteContentText(content)) throw new TypeError('注释文字与块内容投影不一致。')
    const marks = marksAtPoint(candidate, point), reference = createRun(properties.customMark ?? '1', {
      ...marks, verticalAlign: 'superscript', ...properties.referenceTextStyle,
      metadata: { ...marks.metadata, note: { type: kind, noteId: id } },
    })
    const note: DocumentNote = { id, text: documentNoteContentText(content), content, referenceBlockId: block.id, referenceRunId: reference.id,
      ...(properties.customMark !== undefined ? { customMark: properties.customMark } : {}),
      ...(properties.referenceTextStyle ? { referenceTextStyle: cloneModel(properties.referenceTextStyle) } : {}),
      ...(properties.showBodyReference !== undefined ? { showBodyReference: properties.showBodyReference } : {}),
      ...(properties.metadata ? { metadata: cloneModel(properties.metadata) } : {}) }
    const sourceNoteReferences = documentNoteReferenceOffsets(block.runs)
    // References are independent entities, including a legitimate empty custom mark.
    const futureRuns = [...sliceRuns(block.runs as DocumentRun[], 0, point.offset), reference, ...sliceRuns(block.runs as DocumentRun[], point.offset)]
    transformContentControlsForBlockEdit(candidate, block.id, point.offset, point.offset, reference.text.length, undefined, false,
      { projectedNoteReferences: documentNoteReferenceOffsets(futureRuns) })
    transformDocumentTextFieldsForBlockEdit(block, point.offset, point.offset, reference.text, undefined, candidate, documentNoteReferenceOffsets(futureRuns))
    block.runs = futureRuns
    transformDocumentCommentAnchorsForTextEdit(candidate, block.id, point.offset, point.offset, reference.text, Date.now(), sourceNoteReferences)
    for (const ancestor of documentBlockAncestry(candidate, block.id) ?? []) delete ancestor.metadata.html
    if (kind === 'footnote') candidate.footnotes = { ...candidate.footnotes, [id]: note }
    else candidate.endnotes = { ...candidate.endnotes, [id]: note }
    appendDocumentNoteContentControls(candidate, prepared.contentControls)
    if (settings !== undefined) setDocumentNoteSettings(candidate, kind, settings, sectionId, suppressEndnotes, decorationContentControls)
    else if (suppressEndnotes !== undefined || sectionId !== undefined || decorationContentControls !== undefined) throw new TypeError('注释设置作用域和分隔控件需要与 settings 一起提供。')
    renumberDocumentNotes(candidate)
    validateDocumentContentControlRanges(candidate)
    const host: DocumentAuxiliaryBlockHost = { kind: 'note', noteKind: kind, noteId: id }
    return { note: requireNote(candidate, kind, id), host, selection: collapsedSelection(documentPointAtBlockHostOffset(candidate, host, 0)) }
  })
}

function removeNoteAnchors(candidate: DocumentModel, kind: DocumentNoteKind, note: DocumentNote): void {
  const ids = new Set<string>()
  const scoped = documentModelForBlockHost(candidate, { kind: 'note', noteKind: kind, noteId: note.id })
  visitDocumentRepeatingContentControlBlocks(scoped.blocks, block => ids.add(block.id))
  candidate.contentControls = candidate.contentControls?.filter(control => !ids.has(control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId))
  for (const [id, comment] of Object.entries(candidate.comments ?? {})) {
    const anchor = comment.anchor.filter(segment => !ids.has(segment.blockId))
    if (!anchor.length) delete candidate.comments![id]
    else if (anchor.length !== comment.anchor.length) candidate.comments![id] = { ...comment, anchor, referenceText: anchor.map(segment => segment.text).join('\n') }
  }
}

function updateNoteText(candidate: DocumentModel, kind: DocumentNoteKind, id: string, text: string): void {
  const note = requireNote(candidate, kind, id)
  if (text === note.text) return
  const host: DocumentAuxiliaryBlockHost = { kind: 'note', noteKind: kind, noteId: id }, scoped = documentModelForBlockHost(candidate, host, true)
  const paragraphs = documentTextFieldParagraphs(scoped)
  if (scoped.blocks.some(block => block.type !== 'paragraph') || paragraphs.some(paragraph => paragraph.container !== scoped.blocks)) throw new RangeError('包含表格或对象的注释需要在画布中编辑对应内容。')
  assertTextRangeEditable(scoped, 0, paragraphs.at(-1)?.absoluteEnd ?? 0)
  const updated = replaceDocumentRange(scoped, 0, paragraphs.at(-1)?.absoluteEnd ?? 0, text)
  applyDocumentBlockHostModel(candidate, host, updated)
}

export function patchDocumentNote(model: DocumentModel, kind: DocumentNoteKind, id: string, patch: DocumentNotePatch): boolean {
  validateDocumentNotePatch(patch); requireNote(model, kind, id)
  return atomicNoteEdit(model, candidate => {
    const before = cloneModel(candidate), original = requireNote(candidate, kind, id)
    if (patch.content !== undefined) {
      const prepared = prepareDocumentNoteContentWithControls(candidate, kind, id, patch.content, '', true, patch.contentControls)
      const content = prepared.content
      if (patch.text !== undefined && patch.text !== documentNoteContentText(content)) throw new TypeError('注释文字与块内容投影不一致。')
      assertDocumentRepeatingNoteContentRemovable(candidate, original)
      removeNoteAnchors(candidate, kind, original)
      original.content = content; original.text = documentNoteContentText(content)
      appendDocumentNoteContentControls(candidate, prepared.contentControls)
    } else if (patch.text !== undefined) updateNoteText(candidate, kind, id, patch.text)
    const note = requireNote(candidate, kind, id)
    for (const key of ['customMark', 'referenceTextStyle', 'metadata'] as const) {
      const value = patch[key]
      if (value === null) delete note[key]
      else if (value !== undefined) Object.assign(note, { [key]: cloneModel(value) })
    }
    if (patch.showBodyReference !== undefined) note.showBodyReference = patch.showBodyReference
    const reference = documentNoteReferenceRecords(candidate).find(reference => reference.kind === kind && reference.note.id === id)
    if (reference && patch.referenceTextStyle !== undefined) {
      const scoped = reference.host ? documentModelForBlockHost(candidate, reference.host, true) : candidate
      const block = documentTextFieldParagraphs(scoped).find(paragraph => paragraph.block.id === reference.block.id)?.block
      const run = block?.runs.find(run => run.id === reference.run.id)
      if (run) {
        const metadata = run.metadata
        Object.assign(run, patch.referenceTextStyle ?? { ...documentDefaultStyle(candidate).text, verticalAlign: 'superscript' })
        run.metadata = metadata
      }
      if (reference.host) applyDocumentBlockHostModel(candidate, reference.host, scoped)
    }
    // Imported nested note references disappear with the replaced scope in the same candidate.
    removeDeletedDocumentRepeatingNoteScopes(before, candidate)
    renumberDocumentNotes(candidate)
    validateDocumentContentControlRanges(candidate)
    return JSON.stringify(before) !== JSON.stringify(candidate)
  })
}

/** Legacy text setters delegate to the same guarded story edit path. */
export function updateDocumentNote(model: DocumentModel, kind: DocumentNoteKind, id: string, text: string): void {
  patchDocumentNote(model, kind, id, { text })
}

export function deleteDocumentNote(model: DocumentModel, kind: DocumentNoteKind, id: string): boolean {
  assertNoteKind(kind)
  if (!(kind === 'footnote' ? model.footnotes : model.endnotes)?.[id]) return false
  return atomicNoteEdit(model, candidate => {
    const before = cloneModel(candidate), note = requireNote(candidate, kind, id)
    assertDocumentRepeatingNoteContentRemovable(candidate, note)
    const reference = documentNoteReferenceRecords(candidate).find(reference => reference.kind === kind && reference.note.id === id)
    if (reference?.linked) throw new RangeError('请先取消“链接到前一节”再删除该内容区中的注释引用。')
    if (reference) {
      const scoped = reference.host ? documentModelForBlockHost(candidate, reference.host, true) : candidate
      const block = documentTextFieldParagraphs(scoped).find(paragraph => paragraph.block.id === reference.block.id)?.block
      if (!block) throw new RangeError('注释引用没有合法的文字宿主。')
      const index = block.runs.findIndex(run => run.id === reference.run.id)
      if (index < 0) throw new RangeError('注释引用字符已失效。')
      const length = block.runs[index]!.text.length
      blockTextLock(scoped, block.id, reference.offset, reference.offset + length)
      const sourceNoteReferences = documentNoteReferenceOffsets(block.runs)
      const futureRuns = block.runs.filter((_run, position) => position !== index)
      if (!futureRuns.length) futureRuns.push(createRun())
      transformContentControlsForBlockEdit(scoped, block.id, reference.offset, reference.offset + length, 0, undefined, false,
        { projectedNoteReferences: documentNoteReferenceOffsets(futureRuns) })
      transformDocumentTextFieldsForBlockEdit(block, reference.offset, reference.offset + length, '', undefined, scoped, documentNoteReferenceOffsets(futureRuns))
      block.runs = futureRuns
      transformDocumentCommentAnchorsForTextEdit(scoped, block.id, reference.offset, reference.offset + length, '', Date.now(), sourceNoteReferences)
      for (const ancestor of documentBlockAncestry(scoped, block.id) ?? []) delete ancestor.metadata.html
      if (reference.host) applyDocumentBlockHostModel(candidate, reference.host, scoped)
    }
    removeNoteAnchors(candidate, kind, note)
    delete (kind === 'footnote' ? candidate.footnotes : candidate.endnotes)![id]
    removeDeletedDocumentRepeatingNoteScopes(before, candidate)
    renumberDocumentNotes(candidate)
    return true
  })
}
