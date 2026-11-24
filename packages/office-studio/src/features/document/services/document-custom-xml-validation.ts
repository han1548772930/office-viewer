import type { DocumentCustomXmlPart, DocumentCustomXmlResource, DocumentModel } from '../../../domain/model'
import { DocumentCustomXmlError, documentCustomXmlPartKey, documentCustomXmlRelationshipOwner,
  DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH, normalizeDocumentCustomXmlParts, resolveDocumentCustomXmlTarget } from '../../../domain/document-custom-xml'
import { browserDocumentCustomXmlPackageRuntime, type DocumentCustomXmlPackageRelationship,
  type DocumentCustomXmlPackageRuntime } from '../../../platform/dom/custom-xml-runtime'

const PROPERTIES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'
const RELATIONSHIP_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml'
const RELATIONSHIP_PREFIXES = ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/', 'http://purl.oclc.org/ooxml/officeDocument/relationships/']

interface Part { contentType: string; xml?: string; resource?: DocumentCustomXmlResource }
interface Owner { name: string; xml: string; relationships: readonly DocumentCustomXmlPackageRelationship[] }

function fail(code: string, message: string): never { throw new DocumentCustomXmlError(code, message) }
function propertiesRelationship(type: string): boolean { return RELATIONSHIP_PREFIXES.some(prefix => type === prefix + 'customXmlProps') }

/** Incomplete upstream binding definitions stay dormant; active bindings need an explicit store. */
export function assertDocumentCustomXmlBindings(model: DocumentModel): void {
  const stores = model.customXmlParts === undefined ? undefined
    : new Set(model.customXmlParts.flatMap(part => part.storeItemID ? [part.storeItemID.toLowerCase()] : []))
  for (const control of model.contentControls ?? []) {
    const binding = control.dataBinding
    if (binding === undefined) continue
    if (typeof binding !== 'object' || binding === null || Array.isArray(binding)
      || binding.storeItemID !== undefined && (typeof binding.storeItemID !== 'string' || binding.storeItemID.length > 128 || /[\x00-\x1f\x7f-\x9f]/.test(binding.storeItemID))
      || binding.xpath !== undefined && (typeof binding.xpath !== 'string' || binding.xpath.length > DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH)
      || binding.prefixMappings !== undefined && (typeof binding.prefixMappings !== 'string' || binding.prefixMappings.length > DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH)
      || binding.cachedResult !== undefined && typeof binding.cachedResult !== 'string' || binding.dirty !== undefined && typeof binding.dirty !== 'boolean') {
      fail('DOC_CUSTOM_XML_INVALID_BINDING', '内容控件的 XML 绑定参数无效：' + control.id)
    }
    if (stores && binding.storeItemID && binding.xpath && !stores.has(binding.storeItemID.toLowerCase())) {
      fail('DOC_CUSTOM_XML_STORE_NOT_FOUND', '内容控件仍引用不存在的数据源，请先解除绑定再移除：' + control.id)
    }
  }
}

/** Validate every source and its whole retained OPC graph before an editing candidate is published. */
export function validateDocumentCustomXmlParts(value: unknown,
  runtime: DocumentCustomXmlPackageRuntime = browserDocumentCustomXmlPackageRuntime): DocumentCustomXmlPart[] {
  const normalized = normalizeDocumentCustomXmlParts(value), parts = new Map<string, Part>(), owners = new Map<string, Owner>()
  const addOwner = (name: string, xml: string) => {
    const key = documentCustomXmlPartKey(name), previous = owners.get(key)
    if (previous) {
      if (previous.xml !== xml) fail('DOC_CUSTOM_XML_RELATIONSHIP_CONFLICT', '同一个自定义 XML 部件具有不同的关系内容：' + name)
      return
    }
    owners.set(key, { name, xml, relationships: runtime.relationships(xml) })
  }
  for (const part of normalized) {
    runtime.validate(part.xml)
    parts.set(documentCustomXmlPartKey(part.partName), { contentType: part.contentType ?? 'application/xml', xml: part.xml })
    if (part.propertiesPartName && part.propertiesXml !== undefined) {
      runtime.validate(part.propertiesXml)
      parts.set(documentCustomXmlPartKey(part.propertiesPartName), { contentType: PROPERTIES_CONTENT_TYPE, xml: part.propertiesXml })
    }
    if (part.relationshipsXml !== undefined) addOwner(part.partName, part.relationshipsXml)
    for (const resource of part.relatedParts ?? []) {
      const key = documentCustomXmlPartKey(resource.partName)
      if (key.endsWith('.rels')) {
        const owner = documentCustomXmlRelationshipOwner(resource.partName)
        if (!owner || resource.contentType.toLowerCase() !== RELATIONSHIP_CONTENT_TYPE) fail('DOC_CUSTOM_XML_INVALID_RELATIONSHIPS', '关联关系部件的路径、宿主或内容类型无效。')
        addOwner(owner, runtime.decodeXmlResource(resource.dataBase64))
      } else parts.set(key, { contentType: resource.contentType, resource })
    }
  }
  const stores = new Set<string>()
  for (const part of normalized) {
    const propertyEdges = (owners.get(documentCustomXmlPartKey(part.partName))?.relationships ?? []).filter(edge => propertiesRelationship(edge.type))
    if (propertyEdges.length > 1 || propertyEdges.some(edge => edge.external)) fail('DOC_CUSTOM_XML_INVALID_PROPERTIES', '数据源不能引用多个或外部 properties 部件。')
    const related = propertyEdges[0] && resolveDocumentCustomXmlTarget(part.partName, propertyEdges[0].target)
    if (part.propertiesPartName && related && documentCustomXmlPartKey(part.propertiesPartName) !== documentCustomXmlPartKey(related)) fail('DOC_CUSTOM_XML_INVALID_PROPERTIES', 'properties 关系与声明的 properties 部件不一致。')
    const name = part.propertiesPartName ?? related
    if (name) {
      const property = parts.get(documentCustomXmlPartKey(name))
      if (!property || property.contentType.toLowerCase() !== PROPERTIES_CONTENT_TYPE) fail('DOC_CUSTOM_XML_INVALID_PROPERTIES', '自定义 XML properties 部件缺失或内容类型无效。')
      const xml = property.xml ?? runtime.decodeXmlResource(property.resource!.dataBase64)
      const id = runtime.propertiesStoreItemID(xml)
      if (part.storeItemID && part.storeItemID.toLowerCase() !== id.toLowerCase()) fail('DOC_CUSTOM_XML_INVALID_PROPERTIES', 'properties 的实际数据源 ID 与 storeItemID 不一致。')
      part.storeItemID ??= id
    }
    if (part.storeItemID) {
      const id = part.storeItemID.toLowerCase()
      if (stores.has(id)) fail('DOC_CUSTOM_XML_INVALID_PROPERTIES', '自定义 XML properties 的实际数据源 ID 重复。')
      stores.add(id)
    }
  }
  for (const [key, owner] of owners) {
    if (!parts.has(key)) fail('DOC_CUSTOM_XML_MISSING_PART', '自定义 XML 关系的宿主部件不存在：' + owner.name)
    for (const edge of owner.relationships) {
      if (edge.external) continue
      const target = resolveDocumentCustomXmlTarget(owner.name, edge.target)
      if (!parts.has(documentCustomXmlPartKey(target))) fail('DOC_CUSTOM_XML_MISSING_PART', '自定义 XML 关系引用的部件不存在：' + target)
    }
  }
  return normalized
}
