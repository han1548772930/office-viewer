import type { DocumentSectionBreakType, DocumentSectionSetup } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { validateDocumentSectionSetupPatch } from '../features/document/services/document-settings'
import { bodyBoolean, bodyNumber, bodyRecord, type UniverBodyRecord } from './document-body-style'
import { documentBodyArray } from './document-body-types'

// The public descriptor terminates its section; sectionType describes that
// section's beginning. A local sectionBreak block introduces the next section.
export const DOCUMENT_BODY_SECTION_TYPES: readonly DocumentSectionBreakType[] = [
  'unspecified', 'continuous', 'nextColumn', 'nextPage', 'evenPage', 'oddPage',
]

export function documentSectionTypeFromBody(value: unknown): DocumentSectionBreakType {
  const type = bodyNumber(value) ?? 0
  if (!Number.isSafeInteger(type) || !DOCUMENT_BODY_SECTION_TYPES[type]) throw new RangeError('文档分节类型无效。')
  return DOCUMENT_BODY_SECTION_TYPES[type]!
}

/** Public page/column geometry is already expressed in 96-DPI pixels. */
export function documentSectionSetupFromBody(value: unknown): DocumentSectionSetup {
  const raw = bodyRecord(value), result: DocumentSectionSetup = {}, size = bodyRecord(raw.pageSize)
  for (const key of ['width', 'height'] as const) {
    const current = bodyNumber(size[key]); if (current !== undefined) result[key] = current
  }
  for (const key of ['marginTop', 'marginRight', 'marginBottom', 'marginLeft', 'marginHeader', 'marginFooter', 'pageNumberStart'] as const) {
    const current = bodyNumber(raw[key]); if (current !== undefined) result[key] = current
  }
  if (raw.pageOrient !== undefined && raw.pageOrient !== null) {
    if (raw.pageOrient !== 0 && raw.pageOrient !== 1) throw new RangeError('文档页面方向无效。')
    result.orientation = raw.pageOrient === 1 ? 'landscape' : 'portrait'
  }
  if (raw.sectionType !== undefined) result.sectionStart = documentSectionTypeFromBody(raw.sectionType)
  if (raw.columnProperties !== undefined) {
    const columns = documentBodyArray(raw.columnProperties, '分节栏定义')
    if (!columns.length) throw new RangeError('分节栏定义不能为空。')
    for (const column of columns) {
      const width = bodyNumber(column.width), gap = bodyNumber(column.paddingEnd)
      if (width === undefined || width <= 0 || gap === undefined || gap < 0) throw new RangeError('分节栏宽和间距无效。')
    }
    result.columnCount = columns.length
    result.columnGap = Number(columns[0]!.paddingEnd)
  }
  if (raw.columnSeparatorType !== undefined && raw.columnSeparatorType !== null) {
    if (![0, 1, 2].includes(Number(raw.columnSeparatorType)) || typeof raw.columnSeparatorType !== 'number') throw new RangeError('分节栏分隔线无效。')
    result.columnSeparator = raw.columnSeparatorType === 2 ? 'line' : 'none'
  }
  if (raw.noteProperties !== undefined) result.noteProperties = cloneModel(bodyRecord(raw.noteProperties)) as DocumentSectionSetup['noteProperties']
  const suppress = bodyBoolean(raw.suppressEndnotes)
  if (suppress !== undefined) result.suppressEndnotes = suppress
  validateDocumentSectionSetupPatch(result)
  return result
}

export function documentSectionSetupToBody(setup: DocumentSectionSetup | undefined, source: unknown): UniverBodyRecord {
  const raw = cloneModel(bodyRecord(source))
  if (!setup) return raw
  validateDocumentSectionSetupPatch(setup)
  const size = bodyRecord(raw.pageSize)
  if (setup.width !== undefined || setup.height !== undefined) raw.pageSize = { ...size,
    ...(setup.width === undefined ? {} : { width: setup.width }), ...(setup.height === undefined ? {} : { height: setup.height }) }
  for (const key of ['marginTop', 'marginRight', 'marginBottom', 'marginLeft', 'marginHeader', 'marginFooter', 'pageNumberStart', 'noteProperties', 'suppressEndnotes'] as const) {
    if (setup[key] !== undefined) raw[key] = cloneModel(setup[key])
  }
  if (setup.orientation !== undefined) raw.pageOrient = setup.orientation === 'landscape' ? 1 : 0
  if (setup.sectionStart !== undefined) raw.sectionType = DOCUMENT_BODY_SECTION_TYPES.indexOf(setup.sectionStart)
  if (setup.columnSeparator !== undefined) raw.columnSeparatorType = setup.columnSeparator === 'line' ? 2 : 1
  if (setup.columnCount !== undefined) {
    const prior = documentBodyArray(raw.columnProperties, '分节栏定义'), gap = setup.columnGap ?? 18
    const width = setup.width ?? Number(size.width), left = setup.marginLeft ?? Number(raw.marginLeft ?? 0), right = setup.marginRight ?? Number(raw.marginRight ?? 0)
    const available = width - left - right - gap * (setup.columnCount - 1)
    if (Number.isFinite(available) && available <= 0) throw new RangeError('分节栏没有可用内容宽度。')
    const unchanged = prior.length === setup.columnCount && prior.slice(0, -1).every(column => column.paddingEnd === gap)
      && (setup.width === undefined || Number(size.width) === setup.width)
      && (setup.marginLeft === undefined || Number(bodyRecord(source).marginLeft ?? 0) === setup.marginLeft)
      && (setup.marginRight === undefined || Number(bodyRecord(source).marginRight ?? 0) === setup.marginRight)
    if (!unchanged && Number.isFinite(available)) raw.columnProperties = Array.from({ length: setup.columnCount }, (_, index) => ({
      ...prior[index], width: available / setup.columnCount, paddingEnd: index + 1 === setup.columnCount ? 0 : gap,
    }))
  }
  return raw
}
