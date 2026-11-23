import type { DocumentOfficeSourceDto, DocumentOfficeSourceOriginDto } from './contracts.ts'

// Keep these limits aligned with office-io-common/document_legacy_source.rs.
const XML_BYTE_LIMIT = 16 * 1024 * 1024
const BINARY_BYTE_LIMIT = 16 * 1024 * 1024
const TOTAL_RESOURCE_BYTE_LIMIT = 32 * 1024 * 1024
const JSON_BYTE_LIMIT = 64 * 1024 * 1024
const ORIGIN_COUNT_LIMIT = 32
const RESOURCE_COUNT_LIMIT = 1024
const NAME_BYTE_LIMIT = 1024
const ERROR_COUNT_LIMIT = 1024
const ERROR_BYTE_LIMIT = 4096
const TOTAL_ERROR_BYTE_LIMIT = 1024 * 1024
const U32_MAX = 0xFFFFFFFF

function object(value: unknown, label: string): asserts value is Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(label + '必须是对象。')
}

function fields(value: Record<string, unknown>, allowed: string[]): void {
  if (Object.keys(value).some(key => !allowed.includes(key))) throw new TypeError('文档来源包含尚未支持的字段，不能静默丢弃。')
}

// Count bytes without making another complete UTF-8/JSON buffer. Reject lone
// surrogates because the Rust JSON boundary cannot represent them losslessly.
function stringBytes(value: string, limit: number, json = false): number {
  if (value.length > limit) throw new RangeError('文档来源字符串超过字节预算。')
  let size = json ? 2 : 0
  for (const character of value) {
    const code = character.codePointAt(0)!
    if (code >= 0xD800 && code <= 0xDFFF) throw new TypeError('文档来源含有不完整的 Unicode 字符。')
    if (json && (code === 0x22 || code === 0x5C || [8, 9, 10, 12, 13].includes(code))) size += 2
    else if (json && code < 0x20) size += 6
    else size += code < 0x80 ? 1 : code < 0x800 ? 2 : code < 0x10000 ? 3 : 4
    if (size > limit) throw new RangeError('文档来源字符串超过字节预算。')
  }
  return size
}

function name(value: unknown): asserts value is string {
  if (typeof value !== 'string' || !value || value.includes('\0')) throw new TypeError('文档来源资源名称无效。')
  stringBytes(value, NAME_BYTE_LIMIT)
}

function xmlBytes(value: unknown): number {
  if (typeof value !== 'string') throw new TypeError('文档来源 XML 必须是字符串。')
  return stringBytes(value, XML_BYTE_LIMIT)
}

function xmlPart(value: unknown): void {
  if (value === undefined || value === null) return
  object(value, '文档来源 XML part')
  fields(value, ['partName', 'xml'])
  name(value.partName)
  xmlBytes(value.xml)
}

function packagePath(value: unknown): string {
  name(value)
  if (/[\\\0?#:]/.test(value)) throw new TypeError('编号来源资源路径无效。')
  const path = value.replace(/^\/+/, '')
  if (!path || path.split('/').some(part => !part || part === '.' || part === '..')) throw new TypeError('编号来源资源路径必须为规范包路径。')
  return path
}

function mediaBytes(value: unknown): number {
  const limit = Math.ceil(BINARY_BYTE_LIMIT / 3) * 4
  if (typeof value !== 'string' || !value || value.length % 4 || value.length > limit
    || !/^[A-Za-z0-9+/]+={0,2}$/.test(value)) throw new TypeError('编号图片必须为预算内的标准 Base64。')
  const padding = value.endsWith('==') ? 2 : value.endsWith('=') ? 1 : 0
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
  const digit = alphabet.indexOf(value[value.length - padding - 1]!)
  if (padding === 2 && (digit & 15) !== 0 || padding === 1 && (digit & 3) !== 0) throw new TypeError('编号图片 Base64 填充位无效。')
  const bytes = value.length / 4 * 3 - padding
  if (bytes > BINARY_BYTE_LIMIT) throw new RangeError('单个编号图片超过 16 MiB。')
  return bytes
}

function numberingPart(value: unknown): void {
  if (value === undefined || value === null) return
  object(value, '文档编号来源')
  fields(value, ['partName', 'xml', 'relationshipsXml', 'resources'])
  packagePath(value.partName)
  let bytes = xmlBytes(value.xml)
  if (value.relationshipsXml !== undefined && value.relationshipsXml !== null) bytes += xmlBytes(value.relationshipsXml)
  const resources = value.resources === undefined ? [] : value.resources
  if (!Array.isArray(resources) || resources.length > RESOURCE_COUNT_LIMIT) throw new TypeError('编号媒体资源列表无效。')
  const paths = new Set<string>()
  for (const resource of resources) {
    object(resource, '编号媒体资源')
    fields(resource, ['partName', 'contentType', 'dataBase64'])
    const path = packagePath(resource.partName)
    if (paths.has(path)) throw new TypeError('编号媒体资源路径重复。')
    paths.add(path)
    name(resource.contentType)
    if (/[\r\n]/.test(resource.contentType)) throw new TypeError('编号媒体内容类型无效。')
    bytes += mediaBytes(resource.dataBase64)
    if (bytes > TOTAL_RESOURCE_BYTE_LIMIT) throw new RangeError('编号 XML 和媒体来源合计超过 32 MiB。')
  }
  if (bytes > TOTAL_RESOURCE_BYTE_LIMIT) throw new RangeError('编号 XML 和媒体来源合计超过 32 MiB。')
}

function resourceErrors(value: unknown): string[] {
  if (value === undefined) return []
  if (!Array.isArray(value) || value.length > ERROR_COUNT_LIMIT) throw new TypeError('文档来源资源诊断列表无效。')
  let bytes = 0
  for (const error of value) {
    if (typeof error !== 'string') throw new TypeError('文档来源资源诊断必须是字符串。')
    bytes += stringBytes(error, ERROR_BYTE_LIMIT)
    if (bytes > TOTAL_ERROR_BYTE_LIMIT) throw new RangeError('文档来源资源诊断超过 1 MiB。')
  }
  return value
}

interface OriginBudget { origin: DocumentOfficeSourceOriginDto; rawBytes: number; binaryBytes: number; errors: string[] }

function readOrigin(value: unknown): OriginBudget {
  object(value, '文档原始来源')
  fields(value, ['format', 'synthesized', 'resources', 'resourceErrors'])
  if ((value.format !== 'doc' && value.format !== 'odt') || value.synthesized !== (value.format === 'doc')) {
    throw new TypeError('文档原始来源格式或合成标记无效。')
  }
  if (!Array.isArray(value.resources) || value.resources.length > RESOURCE_COUNT_LIMIT) throw new TypeError('文档原始来源资源列表无效。')
  const errors = resourceErrors(value.resourceErrors)
  let rawBytes = 0, binaryBytes = 0
  for (const resource of value.resources) {
    object(resource, '文档原始资源')
    if (resource.kind === 'xml' && value.format === 'odt') {
      fields(resource, ['kind', 'partName', 'xml'])
      name(resource.partName)
      rawBytes += xmlBytes(resource.xml)
    } else if (resource.kind === 'binary' && value.format === 'doc') {
      fields(resource, ['kind', 'streamName', 'offset', 'dataHex'])
      name(resource.streamName)
      if (typeof resource.offset !== 'number' || !Number.isSafeInteger(resource.offset) || resource.offset < 0 || resource.offset > U32_MAX) throw new RangeError('DOC 来源资源偏移无效。')
      if (typeof resource.dataHex !== 'string' || !resource.dataHex.length || resource.dataHex.length % 2 !== 0
        || resource.dataHex.length > BINARY_BYTE_LIMIT * 2 || !/^[0-9a-f]+$/i.test(resource.dataHex)) throw new TypeError('DOC 来源资源十六进制数据无效。')
      const bytes = resource.dataHex.length / 2
      if (resource.offset + bytes > U32_MAX) throw new RangeError('DOC 来源资源偏移和长度溢出。')
      rawBytes += bytes
      binaryBytes += bytes
    } else throw new TypeError('文档原始资源类型与格式不一致。')
    if (rawBytes > TOTAL_RESOURCE_BYTE_LIMIT || binaryBytes > BINARY_BYTE_LIMIT) throw new RangeError('文档原始来源资源超过字节预算。')
  }
  return { origin: value as unknown as DocumentOfficeSourceOriginDto, rawBytes, binaryBytes, errors: [...errors] }
}

function sameHex(left: string, right: string): boolean {
  if (left.length !== right.length) return false
  for (let index = 0; index < left.length; index++) if ((left.charCodeAt(index) | 32) !== (right.charCodeAt(index) | 32)) return false
  return true
}

function sameOrigin(left: DocumentOfficeSourceOriginDto, right: DocumentOfficeSourceOriginDto): boolean {
  return left.format === right.format && left.synthesized === right.synthesized && left.resources.length === right.resources.length
    && left.resources.every((a, index) => {
      const b = right.resources[index]!
      if (a.kind === 'xml' && b.kind === 'xml') return a.partName === b.partName && a.xml === b.xml
      return a.kind === 'binary' && b.kind === 'binary' && a.streamName === b.streamName && a.offset === b.offset && sameHex(a.dataHex, b.dataHex)
    })
}

function encodedSidecarBytes(origins: OriginBudget[]): number {
  let size = '{"version":1,"legacyOrigins":['.length + 2 + Math.max(0, origins.length - 1)
  const stringSize = (value: string) => stringBytes(value, JSON_BYTE_LIMIT, true)
  for (const { origin, errors } of origins) {
    size += '{"format":'.length + stringSize(origin.format) + ',"synthesized":'.length + String(origin.synthesized).length + ',"resources":['.length
    size += Math.max(0, origin.resources.length - 1)
    for (const resource of origin.resources) {
      if (resource.kind === 'xml') size += '{"kind":"xml","partName":'.length + stringSize(resource.partName) + ',"xml":'.length + stringSize(resource.xml) + 1
      else size += '{"kind":"binary","streamName":'.length + stringSize(resource.streamName) + ',"offset":'.length + String(resource.offset).length + ',"dataHex":'.length + resource.dataHex.length + 2 + 1
      if (size > JSON_BYTE_LIMIT) throw new RangeError('文档历史来源 JSON 超过 64 MiB。')
    }
    size += 1 // resources ]
    if (errors.length) size += ',"resourceErrors":['.length + errors.reduce((total, error) => total + stringSize(error), 0) + Math.max(0, errors.length - 1) + 1
    size += 1 // origin }
    if (size > JSON_BYTE_LIMIT) throw new RangeError('文档历史来源 JSON 超过 64 MiB。')
  }
  return size
}

/**
 * Validate provenance without changing the DTO retained in metadata.source.
 * Word parts remain current; the dedicated DOCX part carries legacy origins.
 * Unsupported source fields are rejected instead of lost during Rust export.
 */
export function validateDocumentOfficeSource(
  value: unknown,
  format: string,
): asserts value is DocumentOfficeSourceDto | null | undefined {
  if (value === undefined || value === null) return
  object(value, '文档来源')
  fields(value, ['version', 'styles', 'theme', 'numbering', 'resourceErrors', 'origin', 'legacyOrigins'])
  if (value.version !== 1) throw new TypeError('文档来源版本无效。')
  xmlPart(value.styles)
  xmlPart(value.theme)
  numberingPart(value.numbering)
  const errors = resourceErrors(value.resourceErrors)
  if (format === 'odt' && (value.styles !== undefined && value.styles !== null || value.theme !== undefined && value.theme !== null
    || value.numbering !== undefined && value.numbering !== null)) throw new TypeError('ODT 来源资源不能送入 Word 样式、主题或编号。')
  const candidates: OriginBudget[] = []
  if (value.origin === undefined) {
    if (format === 'doc' || format === 'odt') throw new TypeError('DOC/ODT 文档来源缺少格式标记。')
  } else {
    const current = readOrigin(value.origin)
    if (current.origin.format !== format) throw new TypeError('文档当前来源格式与解析格式不一致。')
    for (const error of errors) if (!current.errors.includes(error)) current.errors.push(error)
    candidates.push(current)
  }
  if (value.legacyOrigins !== undefined) {
    if (!Array.isArray(value.legacyOrigins) || value.legacyOrigins.length > ORIGIN_COUNT_LIMIT) throw new TypeError('文档历史来源列表无效。')
    for (const origin of value.legacyOrigins) candidates.push(readOrigin(origin))
  }
  const origins: OriginBudget[] = []
  for (const candidate of candidates) {
    const existing = origins.find(prior => sameOrigin(prior.origin, candidate.origin))
    if (existing) { for (const error of candidate.errors) if (!existing.errors.includes(error)) existing.errors.push(error) }
    else origins.push(candidate)
  }
  if (origins.length > ORIGIN_COUNT_LIMIT) throw new RangeError('文档来源合计超过 32 项。')
  let resources = 0, rawBytes = 0, binaryBytes = 0, errorCount = 0, errorBytes = 0
  for (const item of origins) {
    resources += item.origin.resources.length
    rawBytes += item.rawBytes
    binaryBytes += item.binaryBytes
    errorCount += item.errors.length
    errorBytes += item.errors.reduce((total, error) => total + stringBytes(error, ERROR_BYTE_LIMIT), 0)
    if (resources > RESOURCE_COUNT_LIMIT || rawBytes > TOTAL_RESOURCE_BYTE_LIMIT || binaryBytes > BINARY_BYTE_LIMIT
      || errorCount > ERROR_COUNT_LIMIT || errorBytes > TOTAL_ERROR_BYTE_LIMIT) throw new RangeError('文档来源合计超过资源或诊断预算。')
  }
  encodedSidecarBytes(origins)
}
