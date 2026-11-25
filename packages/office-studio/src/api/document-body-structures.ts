import type { DocumentBlock, DocumentCalloutConfig, DocumentCodeConfig, DocumentColumnGroup } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { defaultDocumentCalloutConfig, defaultDocumentCodeConfig, validateDocumentCalloutConfig, validateDocumentCodeConfig,
  validateDocumentColumnGroup } from '../features/document/model/document-structural-blocks'
import { bodyBoolean, bodyLength, bodyNumber, bodyRecord, type UniverBodyRecord } from './document-body-style'
import { documentBodyIdentity, type DocumentBodyDecodeOptions } from './document-body-types'

const calloutSizes = ['borderWidth', 'borderRadius', 'paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft', 'iconSize', 'iconGap'] as const
const optionalColumnLength = (value: unknown, fontSize: number): number => {
  const length = bodyLength(value, fontSize)
  if (length === undefined) throw new TypeError('多栏组提供的长度必须为数字或有效 NumberUnit。')
  return length
}

export function documentCalloutFromBody(value: unknown): DocumentCalloutConfig {
  const raw = bodyRecord(value), config = defaultDocumentCalloutConfig()
  for (const key of ['icon', 'backgroundColor', 'borderColor'] as const) if (raw[key] !== undefined) {
    if (typeof raw[key] !== 'string') throw new TypeError('Callout 文本与颜色必须为字符串。')
    config[key] = raw[key]
  }
  config.showIcon = bodyBoolean(raw.showIcon) ?? config.showIcon
  for (const key of calloutSizes) config[key] = bodyNumber(raw[key]) ?? config[key]
  config.borderOpacity = bodyNumber(raw.borderOpacity) ?? config.borderOpacity
  if (raw.borderStyle !== undefined) {
    const styles: Record<string, DocumentCalloutConfig['borderStyle']> = { 0: 'solid', 1: 'solid', 2: 'dotted', 3: 'dashed',
      solid: 'solid', dotted: 'dotted', dashed: 'dashed' }
    const style = styles[String(raw.borderStyle)]
    if (!style) throw new TypeError('Callout 边框类型无效。')
    config.borderStyle = style
  }
  validateDocumentCalloutConfig(config)
  return config
}

export function documentCalloutToBody(config: DocumentCalloutConfig, source?: unknown): UniverBodyRecord {
  validateDocumentCalloutConfig(config)
  return { ...cloneModel(bodyRecord(source)), ...config, borderStyle: config.borderStyle === 'dotted' ? 2 : config.borderStyle === 'dashed' ? 3 : 1 }
}

export function documentCodeFromBody(value: unknown): DocumentCodeConfig {
  const raw = bodyRecord(value), config = defaultDocumentCodeConfig()
  if (raw.language !== undefined) {
    if (typeof raw.language !== 'string') throw new TypeError('代码语言必须为字符串。')
    config.language = raw.language
  }
  config.showLineNumbers = bodyBoolean(raw.showLineNumbers) ?? config.showLineNumbers
  config.wrap = bodyBoolean(raw.wrap) ?? config.wrap
  config.tabSize = bodyNumber(raw.tabSize) ?? config.tabSize
  validateDocumentCodeConfig(config)
  return config
}

export function documentColumnGroupFromBody(raw: UniverBodyRecord, children: DocumentBlock[][], options: DocumentBodyDecodeOptions): DocumentColumnGroup {
  if (raw.version !== undefined && raw.version !== 1) throw new TypeError('多栏组版本无效。')
  if (raw.columns !== undefined && (!Array.isArray(raw.columns) || raw.columns.length !== children.length)) throw new RangeError('多栏组的栏描述与内容标记数量不一致。')
  const groupId = documentBodyIdentity(raw.columnGroupId, '多栏组')
  const group: DocumentColumnGroup = { version: 1, gap: bodyLength(raw.gap, options.defaults.text.fontSize) ?? 16,
    ...(raw.horizontalPadding === undefined ? {} : { horizontalPadding: optionalColumnLength(raw.horizontalPadding, options.defaults.text.fontSize) }),
    ...(raw.minHeight === undefined ? {} : { minHeight: optionalColumnLength(raw.minHeight, options.defaults.text.fontSize) }),
    layout: raw.layout === undefined ? 'fixed' : raw.layout as DocumentColumnGroup['layout'],
    responsive: raw.responsive === undefined ? 'shrink' : raw.responsive as DocumentColumnGroup['responsive'],
    columns: children.map((blocks, index) => {
      const source = bodyRecord(Array.isArray(raw.columns) ? raw.columns[index] : undefined)
      const minWidth = bodyLength(source.minWidth, options.defaults.text.fontSize)
      return { id: documentBodyIdentity(source.columnId ?? `${groupId}:column:${index + 1}`, '栏'), widthRatio: bodyNumber(source.widthRatio) ?? 1,
        ...(minWidth === undefined ? {} : { minWidth }), blocks, metadata: { univerBodyColumnSource: cloneModel(source),
          univerBodyColumnBaseline: { minWidth, fontSize: options.defaults.text.fontSize } } }
    }) }
  validateDocumentColumnGroup(group)
  return group
}

/** Preserve an unchanged authored number unit; changed geometry is exported in pixels. */
export function documentColumnGroupToBody(block: DocumentBlock): UniverBodyRecord {
  const group = block.columns!, source = cloneModel(bodyRecord(block.metadata.univerBodyColumnGroupSource))
  validateDocumentColumnGroup(group)
  const baseline = bodyRecord(block.metadata.univerBodyColumnGroupBaseline)
  source.gap = baseline.gap === group.gap && source.gap !== undefined ? source.gap : { v: group.gap, u: 3 }
  for (const key of ['horizontalPadding', 'minHeight'] as const) {
    if (group[key] === undefined) delete source[key]
    else source[key] = baseline[key] === group[key] && source[key] !== undefined ? source[key] : { v: group[key], u: 3 }
  }
  source.layout = group.layout; source.responsive = group.responsive; source.version = group.version
  source.columns = group.columns.map(column => {
    const raw = cloneModel(bodyRecord(column.metadata?.univerBodyColumnSource)), baseline = bodyRecord(column.metadata?.univerBodyColumnBaseline)
    raw.columnId = column.id; raw.widthRatio = column.widthRatio
    if (column.minWidth === undefined) delete raw.minWidth
    else raw.minWidth = baseline.minWidth === column.minWidth && raw.minWidth !== undefined ? raw.minWidth : { v: column.minWidth, u: 3 }
    return raw
  })
  return source
}
