import type { DocumentCommentMessage, DocumentModel, DocumentPermissionRule, DocumentThreadComment } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { documentStoredBlockHosts } from '../features/document/model/document-block-host'
import { findDocumentBlock, runText } from '../features/document/model/document-model'
import { isDocumentTextBoundary } from '../features/document/model/document-story'
import { assertDocumentNoteReferenceBoundaryRank, documentNoteReferenceOffsets } from '../features/document/model/document-note-boundary'
import { validateDocumentPermissionRules } from '../features/document/permissions/document-permissions'
import { publicDocumentRecord } from './document-input-common'
import { documentBodyIdentity, documentBodyOffset } from './document-body-types'
import type { OfficeDocumentSnapshotSource } from './document-snapshot-types'
import type { UniverBodyRecord } from './document-body-style'

const SCOPES = ['document', 'section', 'paragraph', 'entity'] as const
const ACTIONS = ['view', 'edit', 'copy', 'comment', 'print', 'export'] as const
function list(value: unknown, label: string): unknown[] {
  if (!Array.isArray(value)) throw new TypeError(`${label} 必须为数组。`)
  if (value.length > 100_000) throw new RangeError(`${label} 超过快照数量预算。`)
  return value
}
function string(value: unknown, label: string): string {
  if (typeof value !== 'string') throw new TypeError(`${label} 必须为字符串。`)
  return value
}
function permissionRules(value: unknown): DocumentPermissionRule[] {
  const rules = list(value, 'officeStudio.permissionRules').map((entry, index): DocumentPermissionRule => {
    const raw = publicDocumentRecord(entry, `权限规则 ${index}`)
    if (!SCOPES.includes(raw.scope as DocumentPermissionRule['scope']) || !ACTIONS.includes(raw.action as DocumentPermissionRule['action'])
      || raw.effect !== 'allow' && raw.effect !== 'deny') throw new TypeError('快照权限范围、动作或效果无效。')
    return { id: documentBodyIdentity(raw.id, '权限规则'), scope: raw.scope as DocumentPermissionRule['scope'],
      action: raw.action as DocumentPermissionRule['action'], effect: raw.effect,
      ...(raw.targetId === undefined ? {} : { targetId: documentBodyIdentity(raw.targetId, '权限目标') }),
      userIds: list(raw.userIds, '权限用户').map(value => documentBodyIdentity(value, '权限用户')) }
  })
  validateDocumentPermissionRules(rules)
  return rules
}
function message(value: unknown): DocumentCommentMessage {
  const raw = publicDocumentRecord(value, '评论消息')
  return { id: documentBodyIdentity(raw.id, '评论消息'), authorId: documentBodyIdentity(raw.authorId, '评论作者'),
    authorName: string(raw.authorName, '评论作者名称'), text: string(raw.text, '评论正文'), createdAt: string(raw.createdAt, '评论创建时间'),
    ...(raw.updatedAt === undefined ? {} : { updatedAt: string(raw.updatedAt, '评论更新时间') }),
    ...(raw.mentions === undefined ? {} : { mentions: list(raw.mentions, '评论提及').map(value => {
      const mention = publicDocumentRecord(value, '评论提及')
      return { userId: documentBodyIdentity(mention.userId, '提及用户'), label: string(mention.label, '提及标签') }
    }) }) }
}
function comments(value: unknown): Record<string, DocumentThreadComment> {
  const result: Record<string, DocumentThreadComment> = {}
  for (const [id, value] of Object.entries(publicDocumentRecord(value, 'officeStudio.comments'))) {
    const raw = publicDocumentRecord(value, `评论 ${id}`)
    if (documentBodyIdentity(raw.id, '评论') !== id || typeof raw.resolved !== 'boolean') throw new TypeError('快照评论身份或解决状态无效。')
    const messages = list(raw.messages, '评论消息').map(message), ids = new Set(messages.map(message => message.id))
    if (ids.size !== messages.length) throw new TypeError('同一评论中的消息身份重复。')
    result[id] = { id, anchor: list(raw.anchor, '评论锚点').map(value => {
      const anchor = publicDocumentRecord(value, '评论锚点'), text = string(anchor.text, '评论锚点文字')
      const startOffset = documentBodyOffset(anchor.startOffset, Number.MAX_SAFE_INTEGER, '评论锚点起点')
      const endOffset = documentBodyOffset(anchor.endOffset, Number.MAX_SAFE_INTEGER, '评论锚点终点')
      if (endOffset < startOffset || endOffset - startOffset !== text.length) throw new RangeError('快照评论锚点长度无效。')
      return { blockId: documentBodyIdentity(anchor.blockId, '评论段落'), startOffset, endOffset, text,
        ...(anchor.startNoteReferenceCount === undefined ? {} : { startNoteReferenceCount: documentBodyOffset(anchor.startNoteReferenceCount, Number.MAX_SAFE_INTEGER, '评论起点累计注释序号') }),
        ...(anchor.endNoteReferenceCount === undefined ? {} : { endNoteReferenceCount: documentBodyOffset(anchor.endNoteReferenceCount, Number.MAX_SAFE_INTEGER, '评论终点累计注释序号') }) }
    }), referenceText: string(raw.referenceText, '评论引用文字'), messages, resolved: raw.resolved,
    createdAt: string(raw.createdAt, '评论创建时间'), updatedAt: string(raw.updatedAt, '评论更新时间'),
    ...(raw.metadata === undefined ? {} : { metadata: cloneModel(publicDocumentRecord(raw.metadata, '评论来源')) }) }
  }
  return result
}

/** Versioned service/provenance data only; body text is decoded separately. */
export function readOfficeDocumentSnapshotSource(value: unknown): OfficeDocumentSnapshotSource | undefined {
  if (value === undefined) return undefined
  const raw = publicDocumentRecord(value, 'officeStudio')
  if (raw.version !== 1) throw new RangeError('Office 文档快照扩展版本不匹配。')
  const metadata: UniverBodyRecord | undefined = raw.metadata === undefined ? undefined : cloneModel(publicDocumentRecord(raw.metadata, 'officeStudio.metadata'))
  for (const key of ['univerDocumentSource', 'univerBodySource', 'univerBodyFirstSectionSource', 'documentEditingBlockHost',
    'documentEditingBlockSourceId', 'documentEditingSectionId', 'documentEditingSectionIds']) if (metadata) delete metadata[key]
  return { version: 1, ...(metadata === undefined ? {} : { metadata }),
    ...(raw.permissionRules === undefined ? {} : { permissionRules: permissionRules(raw.permissionRules) }),
    ...(raw.comments === undefined ? {} : { comments: comments(raw.comments) }) }
}

/** Service anchors must address current UTF-16 text, including stored stories,
 * note bodies and table continuations. Empty orphan threads remain valid. */
export function assertOfficeDocumentSnapshotCommentAnchors(model: DocumentModel): void {
  const contents = [model.blocks, ...documentStoredBlockHosts(model).map(entry => entry.blocks)]
  for (const comment of Object.values(model.comments ?? {})) for (const anchor of comment.anchor) {
    const block = contents.map(blocks => findDocumentBlock({ blocks }, anchor.blockId)).find(block => block !== undefined)
    if (!block || ['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)) throw new RangeError('评论锚点未指向真实文字段落：' + anchor.blockId)
    const text = runText(block.runs)
    if (!Number.isSafeInteger(anchor.startOffset) || !Number.isSafeInteger(anchor.endOffset) || anchor.startOffset < 0
      || anchor.endOffset < anchor.startOffset || anchor.endOffset > text.length
      || !isDocumentTextBoundary(text, anchor.startOffset) || !isDocumentTextBoundary(text, anchor.endOffset)
      || text.slice(anchor.startOffset, anchor.endOffset) !== anchor.text) throw new RangeError('评论锚点与当前段落文字或 UTF-16 边界不匹配：' + comment.id)
    const references = documentNoteReferenceOffsets(block.runs)
    const startRank = assertDocumentNoteReferenceBoundaryRank(references, anchor.startOffset, anchor.startNoteReferenceCount)
    const endRank = assertDocumentNoteReferenceBoundaryRank(references, anchor.endOffset, anchor.endNoteReferenceCount)
    if (anchor.startOffset === anchor.endOffset && endRank < startRank) throw new RangeError('评论锚点的同位注释引用范围倒置：' + comment.id)
  }
}
