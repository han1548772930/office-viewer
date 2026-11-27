import type { DocumentDefaultStyle, DocumentModel, DocumentSectionSetup } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { normalizeDocumentCustomXmlParts } from '../domain/document-custom-xml'
import { documentDefaultStyle } from '../features/document/model/document-default-style'
import { normalizeDocument } from '../features/document/model/document-model'
import { validateDocumentContentControlRanges } from '../features/document/services/content-controls'
import { applyDocumentSectionSetupPatch } from '../features/document/services/document-settings'
import { DEFAULT_DOCUMENT_SECTION_ID } from '../features/document/services/section-header-footer'
import { assertDocumentBodySerializable, decodeDocumentBody } from './document-body-codec'
import { bodyRecord, documentParagraphStyleFromBody, documentRunStyleFromBody, type UniverBodyRecord } from './document-body-style'
import { documentSectionSetupFromBody } from './document-body-section'
import { DOCUMENT_BODY_TOKENS as T, documentBodyIdentity, type DocumentBodyDecodeOptions } from './document-body-types'
import { publicDocumentRecord, publicDocumentResources, reservePublicDocumentContent } from './document-input-common'
import { applyPublicDocumentStories } from './document-input-stories'
import { applyPublicDocumentNotes } from './document-input-notes'
import { assertOfficeDocumentSnapshotCommentAnchors, readOfficeDocumentSnapshotSource } from './document-snapshot-source'

function plainTextBody(text: string): { dataStream: string } {
  const stream = text.replace(/\r\n?/g, '\n').replace(/\n/g, T.paragraph)
  return { dataStream: `${stream}${stream.endsWith(T.paragraph) ? '' : T.paragraph}${T.section}` }
}

function publicDefaults(style: UniverBodyRecord): DocumentDefaultStyle {
  const defaults = documentDefaultStyle(), text = documentRunStyleFromBody(style.textStyle)
  delete text.metadata
  const resolved = { ...defaults, text: { ...defaults.text, ...text } }
  resolved.paragraph = documentParagraphStyleFromBody(style.defaultParagraphStyle, resolved)
  return documentDefaultStyle({ defaultStyle: resolved })
}

/** The API shares the real model used by Canvas, commands, history and native
 * IO. Public stream tokens are decoded before any model enters the workbench. */
export function documentFromPublicInput(input: UniverBodyRecord, candidate: DocumentModel): DocumentModel {
  assertDocumentBodySerializable(input)
  if (input.rev !== undefined && (typeof input.rev !== 'number' || !Number.isSafeInteger(input.rev) || input.rev < 0)) throw new RangeError('文档公开版本必须为非负安全整数。')
  const extension = readOfficeDocumentSnapshotSource(input.officeStudio)
  const style = input.documentStyle === undefined ? {} : publicDocumentRecord(input.documentStyle, 'documentStyle')
  const styles: Record<string, UniverBodyRecord> = {}
  if (input.styles !== undefined) for (const [id, value] of Object.entries(publicDocumentRecord(input.styles, 'styles'))) {
    styles[documentBodyIdentity(id, '文档样式')] = publicDocumentRecord(value, `styles.${id}`)
  }
  const defaults = publicDefaults(style), resources = publicDocumentResources(input)
  const options: DocumentBodyDecodeOptions = { defaults, allowSections: true, styles,
    ...(style.defaultParagraphStyleId === undefined ? {} : { defaultParagraphStyleId: documentBodyIdentity(style.defaultParagraphStyleId, '默认段落样式') }) }
  if (input.text !== undefined && typeof input.text !== 'string') throw new TypeError('文档 text 必须为字符串。')
  if (input.body !== undefined && input.text !== undefined) throw new TypeError('文档 body 与 text 不能同时提供。')
  const decoded = decodeDocumentBody(input.body ?? plainTextBody(typeof input.text === 'string' ? input.text : ''), resources, options)
  candidate.blocks = decoded.content.blocks ?? []
  candidate.defaultStyle = defaults
  candidate.contentControls = decoded.contentControls
  const source = cloneModel(input)
  delete source.body
  delete source.officeStudio
  candidate.metadata = { ...candidate.metadata, ...extension?.metadata, ...decoded.content.metadata, documentStyle: cloneModel(style), documentStyles: cloneModel(styles),
    source: { ...bodyRecord(extension?.metadata?.source), documentStyle: cloneModel(style) }, univerDocumentSource: { version: 1, snapshot: source } }
  if (extension?.permissionRules !== undefined) candidate.permissionRules = extension.permissionRules
  const flavor = style.documentFlavor ?? 1
  if (![0, 1, 2, 3].includes(Number(flavor)) || typeof flavor !== 'number') throw new RangeError('文档模式无效。')
  candidate.metadata.documentFlavor = flavor
  candidate.metadata.documentFlavorContract = 'univer'
  if (input.customXmlParts !== undefined) candidate.customXmlParts = normalizeDocumentCustomXmlParts(input.customXmlParts)
  const documentSetup: DocumentSectionSetup = { ...candidate.page, ...documentSectionSetupFromBody(style) }
  // Later sections inherit document defaults, rather than explicit overrides
  // belonging only to the first outgoing public section.
  for (const block of candidate.blocks) if (block.type === 'sectionBreak') block.sectionBreak = {
    ...block.sectionBreak!, setup: { ...documentSetup, ...block.sectionBreak?.setup },
  }
  const firstSetup = documentSectionSetupFromBody(candidate.metadata.univerBodyFirstSectionSource)
  if (flavor === 2) {
    applyDocumentSectionSetupPatch(candidate, DEFAULT_DOCUMENT_SECTION_ID, documentSetup)
    candidate.sectionSetup = { ...candidate.sectionSetup, ...firstSetup }
    candidate.metadata.documentPageOrientation = documentSetup.orientation ?? (candidate.page.width > candidate.page.height ? 'landscape' : 'portrait')
  }
  if (flavor !== 2) applyDocumentSectionSetupPatch(candidate, DEFAULT_DOCUMENT_SECTION_ID, { ...documentSetup, ...firstSetup })
  const claimed = new Set<string>()
  reservePublicDocumentContent(decoded.content, decoded.contentControls, claimed)
  applyPublicDocumentStories(candidate, input, resources, options, claimed)
  applyPublicDocumentNotes(candidate, input, resources, options, claimed)
  // Snapshot service anchors already use the saved model's display offsets.
  // Install them after public U+FFFC references regain their current labels.
  if (extension?.comments !== undefined) candidate.comments = extension.comments
  validateDocumentContentControlRanges(candidate)
  candidate.modified = false
  const model = normalizeDocument(candidate)
  assertOfficeDocumentSnapshotCommentAnchors(model)
  return model
}
