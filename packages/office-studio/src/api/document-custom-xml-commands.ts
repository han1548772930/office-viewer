import type { DocumentCustomXmlPart } from '../domain/model'
import type { OfficeStudio } from './office-studio'
import { DocumentCustomXmlError, normalizeDocumentCustomXmlParts } from '../domain/document-custom-xml'
import { assertDocumentPermission } from '../features/document/permissions/document-permissions'
import { prepareDocumentCustomXmlCandidate } from '../features/document/services/document-custom-xml-candidate'

export const SET_DOCUMENT_CUSTOM_XML_PARTS_COMMAND = 'doc.command.set-custom-xml-parts'

export interface DocumentCustomXmlPartsWriteOptions {
  expectedRevision?: number
  expectedHistoryGeneration?: number
}

export interface DocumentCustomXmlPartsCommandInput extends DocumentCustomXmlPartsWriteOptions {
  unitId: string
  parts: readonly DocumentCustomXmlPart[]
}

/** A local source-edit extension; the upstream SDT entry count remains unchanged. */
export function registerDocumentCustomXmlCommands(api: OfficeStudio): void {
  api.registerCommand<DocumentCustomXmlPartsCommandInput, boolean>(SET_DOCUMENT_CUSTOM_XML_PARTS_COMMAND, input => {
    if (!input || typeof input.unitId !== 'string' || !input.unitId.trim() || !Array.isArray(input.parts)) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_INVALID_SOURCE', '请提供文档 ID 和完整的 XML 数据源数组；清空使用空数组。')
    const current = api.getDocumentModel(input.unitId)
    if (!current) throw new Error('Document not found: ' + input.unitId)
    const revision = api.getDocumentRevision(input.unitId), historyGeneration = api.editor.history.captureSavepoint(input.unitId).generation
    if (input.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision !== revision)
      || input.expectedHistoryGeneration !== undefined && (!Number.isSafeInteger(input.expectedHistoryGeneration) || input.expectedHistoryGeneration !== historyGeneration)) {
      throw new DocumentCustomXmlError('DOC_CUSTOM_XML_STALE_DOCUMENT', '文档内容或历史已经更改，请重新修改 XML 数据源。')
    }
    const permissionContext = { documentId: input.unitId, userId: api.currentUserId }
    assertDocumentPermission(current.permissionRules ?? [], permissionContext, 'edit')
    const parts = normalizeDocumentCustomXmlParts(input.parts)
    if (api.getDocumentModel(input.unitId) !== current || api.getDocumentRevision(input.unitId) !== revision
      || api.editor.history.captureSavepoint(input.unitId).generation !== historyGeneration) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_STALE_DOCUMENT', '准备数据源期间文档已经更改，请重新执行。')
    return api.commitDocument(input.unitId, '修改自定义 XML 数据源', model => {
      model.customXmlParts = parts
      prepareDocumentCustomXmlCandidate(current, model, permissionContext, true)
    })
  })
}
