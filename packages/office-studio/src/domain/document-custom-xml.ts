import type { DocumentCustomXmlPart } from './model'

export const DOCUMENT_CUSTOM_XML_MAX_PART_BYTES = 8 * 1024 * 1024
export const DOCUMENT_CUSTOM_XML_MAX_TOTAL_BYTES = 64 * 1024 * 1024
export const DOCUMENT_CUSTOM_XML_MAX_PARTS = 256
export const DOCUMENT_CUSTOM_XML_MAX_RELATED_PARTS = 2048
export const DOCUMENT_CUSTOM_XML_MAX_ELEMENTS = 100_000
export const DOCUMENT_CUSTOM_XML_MAX_DEPTH = 128
export const DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH = 8192

export class DocumentCustomXmlError extends Error {
  constructor(readonly code: string, message: string, readonly details?: string) {
    super(message)
    this.name = 'DocumentCustomXmlError'
  }
}

/** Count Unicode bytes without allocating a duplicate TextEncoder buffer. */
export function documentCustomXmlUtf8Bytes(value: string): number {
  let bytes = 0
  for (const character of value) {
    const point = character.codePointAt(0)!
    if (point >= 0xd800 && point <= 0xdfff) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_INVALID_UNICODE', 'XML 数据包含不完整的 Unicode 字符。')
    bytes += point < 0x80 ? 1 : point < 0x800 ? 2 : point < 0x10000 ? 3 : 4
  }
  return bytes
}

/** Estimate the actual encoded output, including a corrected XML declaration and BOM. */
export function documentCustomXmlEncodedBytes(xml: string, encoding: DocumentCustomXmlPart['encoding'] = 'utf8', byteOrderMark = false): number {
  const utf8Bytes = documentCustomXmlUtf8Bytes(xml), end = xml.indexOf('?>')
  const declaration = /^<\?xml\s/.test(xml) && end >= 0 ? xml.slice(0, end) : ''
  const match = /(?:^|\s)encoding\s*=\s*(['"])([^'"]*)\1/.exec(declaration)
  const expected = encoding === 'utf16le' ? 'UTF-16LE' : encoding === 'utf16be' ? 'UTF-16BE' : 'UTF-8'
  const previous = match?.[2], matching = previous?.toUpperCase() === expected || encoding !== 'utf8' && previous?.toUpperCase() === 'UTF-16'
  const utf8Delta = previous === undefined || matching ? 0 : expected.length - documentCustomXmlUtf8Bytes(previous)
  const unitDelta = previous === undefined || matching ? 0 : expected.length - previous.length
  const serializedBytes = utf8Bytes + utf8Delta
  const encodedBytes = encoding === 'utf16le' || encoding === 'utf16be' ? (xml.length + unitDelta) * 2 + (byteOrderMark ? 2 : 0) : serializedBytes + (byteOrderMark ? 3 : 0)
  return Math.max(serializedBytes, encodedBytes)
}

export function isDocumentCustomXmlPartName(value: unknown): value is string {
  if (typeof value !== 'string') return false
  try { canonicalDocumentCustomXmlPartName(value); return true } catch { return false }
}

export function isDocumentCustomXmlContentType(value: unknown): value is string {
  // A negative any-character lookahead requires the actual end, including final line breaks.
  return typeof value === 'string' && /^[A-Za-z0-9!#$%&'*+.^_`|~-]+\/[A-Za-z0-9!#$%&'*+.^_`|~-]+(?![\s\S])/.test(value)
}

/** Canonical relative OPC URI; ZIP entry spelling is a separate representation. */
export function canonicalDocumentCustomXmlPartName(value: string): string {
  const invalid = () => { throw new DocumentCustomXmlError('DOC_CUSTOM_XML_INVALID_PART_NAME', '自定义 XML 部件路径无效。') }
  if (!value || value.length > 1024 || value.startsWith('/') || /[\\:#?\x00-\x1f\x7f-\x9f]/.test(value)) invalid()
  const segments = value.split('/').map(segment => {
    let decoded: string
    try { decoded = decodeURIComponent(segment) } catch { return invalid() }
    if (!decoded || decoded === '.' || decoded === '..' || decoded.endsWith('.') || /[\/\\\x00-\x1f\x7f-\x9f]/.test(decoded)) invalid()
    try { return encodeURIComponent(decoded) } catch { return invalid() }
  })
  const name = segments.join('/')
  if (name.length > 1024) invalid()
  return name
}

/** OPC part identity ignores ASCII case, including percent-escape spelling. */
export function documentCustomXmlPartKey(value: string): string { return canonicalDocumentCustomXmlPartName(value).toLowerCase() }

export function documentCustomXmlRelationshipPart(owner: string): string {
  const name = canonicalDocumentCustomXmlPartName(owner), separator = name.lastIndexOf('/')
  return separator < 0 ? '_rels/' + name + '.rels' : name.slice(0, separator) + '/_rels/' + name.slice(separator + 1) + '.rels'
}

export function documentCustomXmlRelationshipOwner(part: string): string | undefined {
  const name = canonicalDocumentCustomXmlPartName(part)
  if (!name.toLowerCase().endsWith('.rels')) return undefined
  const separator = name.toLowerCase().lastIndexOf('/_rels/')
  const owner = separator < 0 ? name.toLowerCase().startsWith('_rels/') ? name.slice(6, -5) : undefined
    : name.slice(0, separator) + '/' + name.slice(separator + 7, -5)
  return owner && isDocumentCustomXmlPartName(owner) ? canonicalDocumentCustomXmlPartName(owner) : undefined
}

export function resolveDocumentCustomXmlTarget(owner: string, target: string): string {
  if (!target || target.length > 2048 || target.startsWith('//') || /[\\:#?\x00-\x1f\x7f-\x9f]/.test(target)) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_INVALID_TARGET', '自定义 XML 关系目标无效。')
  const segments = target.startsWith('/') || !owner ? [] : canonicalDocumentCustomXmlPartName(owner).split('/').slice(0, -1)
  for (const segment of (target.startsWith('/') ? target.slice(1) : target).split('/')) {
    if (segment === '.') continue
    if (segment === '..') {
      if (!segments.pop()) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_INVALID_TARGET', '自定义 XML 关系目标超出文件包。')
    } else segments.push(canonicalDocumentCustomXmlPartName(segment))
  }
  return canonicalDocumentCustomXmlPartName(segments.join('/'))
}

function plainObject(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function denseArray(value: unknown, maximum: number): value is unknown[] {
  if (!Array.isArray(value) || value.length > maximum) return false
  for (let index = 0; index < value.length; index++) if (!Object.hasOwn(value, index)) return false
  return true
}

function fail(message: string): never { throw new DocumentCustomXmlError('DOC_CUSTOM_XML_INVALID_SOURCE', message) }

/** Validate canonical base64 in one pass without a decoded resource copy. */
export function documentCustomXmlResourceByteLength(value: string): number {
  if (value.length > Math.ceil(DOCUMENT_CUSTOM_XML_MAX_PART_BYTES / 3) * 4) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_RESOURCE_LIMIT', '自定义 XML 关联资源超过 8 MiB 限额。')
  if (value.length % 4 !== 0) fail('自定义 XML 关联资源不是完整的 base64。')
  const padding = value.endsWith('==') ? 2 : value.endsWith('=') ? 1 : 0, limit = value.length - padding
  const digit = (code: number): number => code >= 65 && code <= 90 ? code - 65 : code >= 97 && code <= 122 ? code - 71
    : code >= 48 && code <= 57 ? code + 4 : code === 43 ? 62 : code === 47 ? 63 : -1
  for (let index = 0; index < limit; index++) if (digit(value.charCodeAt(index)) < 0) fail('自定义 XML 关联资源包含非法 base64 字符。')
  if (padding && (limit < 2 || digit(value.charCodeAt(limit - 1)) % (padding === 2 ? 16 : 4) !== 0)) fail('自定义 XML 关联资源包含非标准 base64 尾位。')
  const bytes = value.length / 4 * 3 - padding
  if (bytes > DOCUMENT_CUSTOM_XML_MAX_PART_BYTES) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_RESOURCE_LIMIT', '自定义 XML 关联资源超过 8 MiB 限额。')
  return bytes
}

/** Structural boundary shared by parser normalization, snapshot APIs and binding sessions. */
export function normalizeDocumentCustomXmlParts(value: unknown): DocumentCustomXmlPart[] {
  if (value === undefined) return []
  if (!denseArray(value, DOCUMENT_CUSTOM_XML_MAX_PARTS)) fail('自定义 XML 数据源列表不是有效的有界密集数组。')
  let totalBytes = 0, resourceCount = 0
  const partNames = new Set<string>(), stores = new Set<string>()
  const resources = new Map<string, { contentType: string; dataBase64: string }>()
  const count = (text: string, maximum = DOCUMENT_CUSTOM_XML_MAX_PART_BYTES, encodedBytes = 0) => {
    const bytes = documentCustomXmlUtf8Bytes(text)
    if (bytes > maximum || encodedBytes > maximum) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_RESOURCE_LIMIT', '自定义 XML 部件超过 8 MiB 限额。')
    totalBytes += Math.max(bytes, encodedBytes)
    if (totalBytes > DOCUMENT_CUSTOM_XML_MAX_TOTAL_BYTES) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_RESOURCE_LIMIT', '自定义 XML 数据源超过整份文档的 64 MiB 限额。')
  }
  const normalized = value.map((part): DocumentCustomXmlPart => {
    if (!plainObject(part) || !isDocumentCustomXmlPartName(part.partName) || typeof part.xml !== 'string' || !part.xml.trim()) fail('自定义 XML 部件的名称或内容无效。')
    const allowed = new Set(['partName', 'xml', 'contentType', 'storeItemID', 'propertiesPartName', 'propertiesXml', 'relationshipsXml', 'relatedParts', 'encoding', 'byteOrderMark'])
    if (Object.keys(part).some(key => !allowed.has(key))) fail('自定义 XML 部件包含未知字段。')
    const partName = canonicalDocumentCustomXmlPartName(part.partName), partKey = documentCustomXmlPartKey(partName)
    if (partKey.endsWith('.rels') || partKey === '%5bcontent_types%5d.xml') fail('数据源不能占用 OPC 系统部件。')
    if (partNames.has(partKey)) fail('自定义 XML 部件名称重复：' + part.partName)
    partNames.add(partKey)
    const result: DocumentCustomXmlPart = { partName, xml: part.xml }
    if (part.contentType !== undefined) {
      if (!isDocumentCustomXmlContentType(part.contentType)
        || part.contentType.toLowerCase() !== 'application/xml' && part.contentType.toLowerCase() !== 'text/xml' && !part.contentType.toLowerCase().endsWith('+xml')) fail('自定义 XML 数据源的 contentType 必须是 XML 类型。')
      result.contentType = part.contentType
    }
    if (part.storeItemID !== undefined) {
      if (typeof part.storeItemID !== 'string' || !part.storeItemID || part.storeItemID.length > 128 || /[\x00-\x1f\x7f-\x9f]/.test(part.storeItemID)) fail('自定义 XML storeItemID 无效。')
      const key = part.storeItemID.toLowerCase()
      if (stores.has(key)) fail('自定义 XML storeItemID 重复：' + part.storeItemID)
      stores.add(key); result.storeItemID = part.storeItemID
    }
    if (part.propertiesPartName !== undefined || part.propertiesXml !== undefined) {
      if (!isDocumentCustomXmlPartName(part.propertiesPartName) || typeof part.propertiesXml !== 'string' || !part.propertiesXml.trim()) fail('自定义 XML properties 部件的名称和 XML 必须一起提供。')
      const propertyName = canonicalDocumentCustomXmlPartName(part.propertiesPartName), propertyKey = documentCustomXmlPartKey(propertyName)
      if (propertyKey.endsWith('.rels') || propertyKey === '%5bcontent_types%5d.xml') fail('properties 不能占用 OPC 系统部件。')
      if (!result.storeItemID) fail('自定义 XML properties 必须提供 storeItemID。')
      if (partNames.has(propertyKey)) fail('自定义 XML properties 部件名称重复。')
      partNames.add(propertyKey)
      result.propertiesPartName = propertyName; result.propertiesXml = part.propertiesXml; count(part.propertiesXml, DOCUMENT_CUSTOM_XML_MAX_PART_BYTES, documentCustomXmlEncodedBytes(part.propertiesXml))
    }
    if (part.relationshipsXml !== undefined) {
      if (typeof part.relationshipsXml !== 'string' || !part.relationshipsXml.trim()) fail('自定义 XML relationships 内容无效。')
      result.relationshipsXml = part.relationshipsXml; count(part.relationshipsXml, DOCUMENT_CUSTOM_XML_MAX_PART_BYTES, documentCustomXmlEncodedBytes(part.relationshipsXml))
    }
    if (part.encoding !== undefined) {
      if (part.encoding !== 'utf8' && part.encoding !== 'utf16le' && part.encoding !== 'utf16be') fail('自定义 XML 字节编码无效。')
      result.encoding = part.encoding
    }
    if (part.byteOrderMark !== undefined) {
      if (typeof part.byteOrderMark !== 'boolean') fail('自定义 XML BOM 标志无效。')
      result.byteOrderMark = part.byteOrderMark
    }
    const encodedBytes = documentCustomXmlEncodedBytes(part.xml, result.encoding, result.byteOrderMark)
    count(part.xml, DOCUMENT_CUSTOM_XML_MAX_PART_BYTES, encodedBytes)
    if (part.relatedParts !== undefined) {
      if (!denseArray(part.relatedParts, DOCUMENT_CUSTOM_XML_MAX_RELATED_PARTS)) fail('自定义 XML 关联资源列表无效。')
      resourceCount += part.relatedParts.length
      if (resourceCount > DOCUMENT_CUSTOM_XML_MAX_RELATED_PARTS) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_RESOURCE_LIMIT', '自定义 XML 关联资源数量超过限额。')
      result.relatedParts = part.relatedParts.map(resource => {
        if (!plainObject(resource) || !isDocumentCustomXmlPartName(resource.partName) || !isDocumentCustomXmlContentType(resource.contentType) || typeof resource.dataBase64 !== 'string'
          || Object.keys(resource).some(key => !['partName', 'contentType', 'dataBase64'].includes(key))) fail('自定义 XML 关联资源格式无效。')
        documentCustomXmlResourceByteLength(resource.dataBase64)
        const resourceName = canonicalDocumentCustomXmlPartName(resource.partName), resourceKey = documentCustomXmlPartKey(resourceName)
        if (resourceKey === '%5bcontent_types%5d.xml') fail('关联资源不能覆盖 OPC content types 部件。')
        const previous = resources.get(resourceKey)
        if (previous && (previous.contentType.toLowerCase() !== resource.contentType.toLowerCase() || previous.dataBase64 !== resource.dataBase64)) fail('自定义 XML 关联资源同名内容冲突：' + resource.partName)
        if (!previous) { resources.set(resourceKey, { contentType: resource.contentType, dataBase64: resource.dataBase64 }); count(resource.dataBase64, Math.ceil(DOCUMENT_CUSTOM_XML_MAX_PART_BYTES / 3) * 4) }
        return { partName: resourceName, contentType: resource.contentType, dataBase64: resource.dataBase64 }
      })
    }
    return result
  })
  if ([...resources.keys()].some(name => partNames.has(name))) fail('自定义 XML 数据部件与关联资源名称冲突。')
  return normalized
}
