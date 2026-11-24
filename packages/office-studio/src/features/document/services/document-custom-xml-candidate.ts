import type { DocumentModel } from '../../../domain/model'
import { DocumentCustomXmlError } from '../../../domain/document-custom-xml'
import type { DocumentPermissionContext } from '../permissions/document-permissions'
import { runText, runsForTarget, textTargets } from '../model/document-model'
import { synchronizeDocumentContentControlBindings } from './content-control-bindings'
import { assertDocumentCustomXmlBindings, validateDocumentCustomXmlParts } from './document-custom-xml-validation'

function mappings(model: DocumentModel): string {
  return JSON.stringify((model.contentControls ?? []).filter(control => control.dataBinding).map(control => ({
    id: control.id, type: control.type, xpath: control.dataBinding!.xpath,
    storeItemID: control.dataBinding!.storeItemID, prefixMappings: control.dataBinding!.prefixMappings,
  })))
}

/** The same candidate boundary protects source setters and arbitrary Facade mutations. */
export function prepareDocumentCustomXmlCandidate(before: DocumentModel, candidate: DocumentModel,
  permissionContext: DocumentPermissionContext, forceRefresh = false): void {
  const sourcesChanged = JSON.stringify(before.customXmlParts) !== JSON.stringify(candidate.customXmlParts)
  const mappingsChanged = mappings(before) !== mappings(candidate)
  if (!sourcesChanged && !mappingsChanged && !forceRefresh) { assertDocumentCustomXmlBindings(candidate); return }
  if (before.customXmlParts !== undefined && candidate.customXmlParts === undefined) {
    throw new DocumentCustomXmlError('DOC_CUSTOM_XML_INVALID_SOURCE', '移除 XML 数据源请使用空数组，并先解除引用它的内容控件绑定。')
  }
  if (candidate.customXmlParts !== undefined) candidate.customXmlParts = validateDocumentCustomXmlParts(candidate.customXmlParts)
  assertDocumentCustomXmlBindings(candidate)
  const changes = synchronizeDocumentContentControlBindings(before, candidate, { refreshCustomXml: candidate.customXmlParts !== undefined, permissionContext })
  if (changes.textEdits.length) candidate.wordCount = textTargets(candidate).reduce((count, target) => count + runText(runsForTarget(candidate, target)).replace(/\s/g, '').length, 0)
}
