import type { DocumentBlock, DocumentCalloutConfig, DocumentCodeConfig, DocumentColumnGroup } from '../../../domain/model'
import type { DocumentTable } from './document-model'
import { documentTableCellContentContainers } from './document-table-continuations'

export type DocumentDecoratedBlockType = 'quote' | 'callout' | 'code'
export const DOCUMENT_DECORATED_BLOCK_TYPES: readonly DocumentDecoratedBlockType[] = ['quote', 'callout', 'code']
/** Fixed v1.6.0 docs-column public contract; section columns have separate bounds. */
export const DOCUMENT_COLUMN_GROUP_MIN_COLUMNS = 2
export const DOCUMENT_COLUMN_GROUP_MAX_COLUMNS = 5
export const DOCUMENT_COLUMN_HORIZONTAL_PADDING = 8
export const DOCUMENT_COLUMN_MIN_HEIGHT = 72

export function defaultDocumentCalloutConfig(): DocumentCalloutConfig {
  return { icon: '💡', showIcon: true, backgroundColor: '#eef6ff', borderColor: '#a8c7e8', borderWidth: 1,
    borderOpacity: 1, borderStyle: 'solid', borderRadius: 6, paddingTop: 12, paddingRight: 12, paddingBottom: 12,
    paddingLeft: 12, iconSize: 18, iconGap: 10 }
}

export function defaultDocumentCodeConfig(): DocumentCodeConfig {
  return { language: 'plain', showLineNumbers: true, tabSize: 4, wrap: true }
}

/** These arrays own their blocks; traversal never introduces a second tree. */
export function documentStructuredChildContainers(block: DocumentBlock): DocumentBlock[][] {
  if (block.type === 'columns') return block.columns?.columns.map(column => column.blocks) ?? []
  return block.children === undefined ? [] : [block.children]
}

export function documentBlockChildContainers(block: DocumentBlock): DocumentBlock[][] {
  const containers = documentStructuredChildContainers(block)
  const raw = block.metadata.table
  if (raw && typeof raw === 'object' && 'rows' in raw && Array.isArray(raw.rows)) {
    for (const row of (raw as DocumentTable).rows) for (const cell of row.cells) containers.push(...documentTableCellContentContainers(cell))
  }
  return containers
}

export function isDocumentStructuredBlock(block: DocumentBlock): boolean {
  return block.type === 'columns' || block.type === 'callout' || block.type === 'code' || block.type === 'quote' && block.children !== undefined
}

export function isDocumentTextBlock(block: DocumentBlock): boolean {
  return !isDocumentStructuredBlock(block) && !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)
}

export function visitDocumentBlocks(blocks: readonly DocumentBlock[], visitor: (block: DocumentBlock) => void): void {
  const visit = (blocks: readonly DocumentBlock[], depth: number) => {
    if (depth > 64) throw new RangeError('文档结构嵌套超过 64 层。')
    for (const block of blocks) {
      visitor(block)
      for (const children of documentBlockChildContainers(block)) visit(children, depth + 1)
    }
  }
  visit(blocks, 0)
}

const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
const finite = (value: unknown, low: number, high: number, label: string): number => {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < low || value > high) throw new RangeError(`${label} 超出允许范围。`)
  return value
}
const text = (value: unknown, maximum: number, label: string): string => {
  if (typeof value !== 'string' || value.length > maximum) throw new TypeError(`${label} 必须为有效字符串。`)
  return value
}

export function validateDocumentCalloutConfig(config: DocumentCalloutConfig): void {
  if (!record(config) || typeof config.showIcon !== 'boolean') throw new TypeError('Callout 图标配置无效。')
  text(config.icon, 256, 'Callout 图标')
  for (const key of ['backgroundColor', 'borderColor'] as const) text(config[key], 512, 'Callout 颜色')
  for (const key of ['borderWidth', 'borderRadius', 'paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft', 'iconSize', 'iconGap'] as const) finite(config[key], 0, 4096, 'Callout 尺寸')
  finite(config.borderOpacity, 0, 1, 'Callout 边框透明度')
  if (!['solid', 'dashed', 'dotted'].includes(config.borderStyle)) throw new TypeError('Callout 边框类型无效。')
}

export function validateDocumentCodeConfig(config: DocumentCodeConfig): void {
  if (!record(config) || typeof config.showLineNumbers !== 'boolean' || typeof config.wrap !== 'boolean') throw new TypeError('代码块配置无效。')
  if (!text(config.language, 256, '代码语言').trim()) throw new TypeError('代码语言不能为空。')
  finite(config.tabSize, 1, 16, '代码缩进')
  if (!Number.isInteger(config.tabSize)) throw new RangeError('代码缩进必须为整数。')
}

export function validateDocumentColumnGroup(config: DocumentColumnGroup): void {
  if (!record(config) || config.version !== 1 || !Array.isArray(config.columns) || config.columns.length < DOCUMENT_COLUMN_GROUP_MIN_COLUMNS || config.columns.length > DOCUMENT_COLUMN_GROUP_MAX_COLUMNS
    || !['fixed', 'auto'].includes(config.layout) || !['stack', 'shrink'].includes(config.responsive)) throw new TypeError('多栏组配置无效。')
  finite(config.gap, 0, 4096, '栏间距')
  if (config.horizontalPadding !== undefined) finite(config.horizontalPadding, 0, 4096, '栏水平内边距')
  if (config.minHeight !== undefined) finite(config.minHeight, 0, 1e6, '栏最小高度')
  const ids = new Set<string>()
  for (const column of config.columns) {
    if (!record(column) || !text(column.id, 4096, '栏标识').trim() || ids.has(column.id) || !Array.isArray(column.blocks)) throw new TypeError('多栏组必须有独立的栏标识与内容。')
    ids.add(column.id); finite(column.widthRatio, Number.MIN_VALUE, 1e12, '栏宽比例')
    if (column.minWidth !== undefined) finite(column.minWidth, 0, 1e6, '栏最小宽度')
    if (column.metadata !== undefined && !record(column.metadata)) throw new TypeError('栏来源数据无效。')
  }
}

export function validateDocumentStructuredBlock(block: DocumentBlock): void {
  const decorated = DOCUMENT_DECORATED_BLOCK_TYPES.some(type => type === block.type)
  if (block.children !== undefined && (!decorated || !Array.isArray(block.children))) throw new TypeError('高级块的子内容与块类型不匹配。')
  if ((block.type === 'callout' || block.type === 'code') && block.children === undefined) throw new TypeError('高级块缺少真实的子内容。')
  if (isDocumentStructuredBlock(block) && block.runs.length) throw new TypeError('结构容器不能同时持有段落文本。')
  if (block.type === 'callout') validateDocumentCalloutConfig(block.callout!)
  else if (block.callout !== undefined) throw new TypeError('Callout 属性与块类型不匹配。')
  if (block.type === 'code') validateDocumentCodeConfig(block.code!)
  else if (block.code !== undefined) throw new TypeError('代码属性与块类型不匹配。')
  if (block.type === 'columns') validateDocumentColumnGroup(block.columns!)
  else if (block.columns !== undefined) throw new TypeError('多栏属性与块类型不匹配。')
}
