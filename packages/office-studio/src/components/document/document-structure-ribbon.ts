import { CornerDownLeft, FileCog, ImagePlus, Layers2, Trash2 } from 'lucide-vue-next'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema'
import type { DocumentCodeConfig, DocumentColumnGroup } from '../../domain/model'
import type { DocumentStructureCommand } from '../../features/document/commands/document-structure-commands'
import { DOCUMENT_COLUMN_GROUP_MIN_COLUMNS, DOCUMENT_COLUMN_GROUP_MAX_COLUMNS } from '../../features/document/model/document-structural-blocks'
import { documentStructuredCodeLanguageOptions } from '../../features/document/code-blocks/document-code-languages'

export interface DocumentStructureRibbonContext {
  blockId: string
  type: 'quote' | 'callout' | 'code' | 'columns'
  disabled: boolean
  code?: DocumentCodeConfig
  columns?: DocumentColumnGroup
  columnId?: string
  fontFamily?: string
}

export function documentStructureContextItems(context: DocumentStructureRibbonContext): RibbonItem[] {
  const title = { quote: '引用块', callout: '标注块', code: '代码块', columns: '多栏组' }[context.type]
  const item = { type: 'button' as const, disabled: context.disabled }
  return [
    ...(context.code ? [{ ...item, type: 'menu' as const, id: 'doc-context-code-language', title: '代码语言',
      children: documentStructuredCodeLanguageOptions(context.code.language).map(language => ({ ...item, id: `doc-context-code-language-${language.value}`,
        title: language.label, activated: language.value === context.code!.language, commandId: 'doc.command.update-code-settings',
        params: { blockId: context.blockId, patch: { language: language.value } } })) },
    { ...item, id: 'doc-context-code-wrap', title: context.code.wrap ? '取消自动换行' : '自动换行', commandId: 'doc.command.update-code-settings',
      params: { blockId: context.blockId, patch: { wrap: !context.code.wrap } } },
    { ...item, id: 'doc-context-code-line-numbers', title: context.code.showLineNumbers ? '隐藏行号' : '显示行号', commandId: 'doc.command.update-code-settings',
      params: { blockId: context.blockId, patch: { showLineNumbers: !context.code.showLineNumbers } } },
    { ...item, id: 'doc-context-code-font', title: '等宽字体', commandId: 'doc.operation.edit-code-font', params: { blockId: context.blockId } },
    { ...item, id: 'doc-context-code-more', title: '缩进与换行设置', commandId: 'doc.operation.edit-code-settings', params: { blockId: context.blockId } }] : []),
    ...(context.columns && context.columnId ? [
      { ...item, id: 'doc-context-column-add', title: '在当前栏右侧添加', disabled: context.disabled || context.columns.columns.length >= DOCUMENT_COLUMN_GROUP_MAX_COLUMNS,
        commandId: 'doc.command.add-column', params: { blockId: context.blockId, columnId: context.columnId, position: 'after' } },
      { ...item, id: 'doc-context-column-delete', title: '删除当前栏和内容', disabled: context.disabled || context.columns.columns.length <= DOCUMENT_COLUMN_GROUP_MIN_COLUMNS,
        commandId: 'doc.command.remove-column', params: { blockId: context.blockId, columnId: context.columnId, keepContents: false } },
      { ...item, id: 'doc-context-column-equal', title: '各栏等宽', commandId: 'doc.command.update-column-group',
        params: { blockId: context.blockId, widthRatios: context.columns.columns.map(() => 1) } },
    ] : []),
    { ...item, id: 'doc-context-structured-properties', title: `${title}属性`, commandId: 'doc.operation.edit-structured-block', params: { blockId: context.blockId } },
    { ...item, id: 'doc-context-structured-unwrap', title: '保留内容并移除容器', commandId: 'doc.command.unwrap-structured-block', params: { blockId: context.blockId } },
    { ...item, id: 'doc-context-structured-exit', title: '在块后继续输入', commandId: 'doc.command.exit-structured-block', params: { blockId: context.blockId } },
    { ...item, id: 'doc-context-structured-remove', title: '删除块和内容', commandId: 'doc.command.remove-structured-block', params: { blockId: context.blockId } },
  ]
}

export const DOCUMENT_STRUCTURED_PARAGRAPH_MENU_MOUNTS = ['layout-below', 'content-insert', 'empty-paragraph', 'quick-insert', 'quick-edit', 'quick-insert-below', 'table-below'] as const
export type DocumentStructuredParagraphMenuMount = typeof DOCUMENT_STRUCTURED_PARAGRAPH_MENU_MOUNTS[number]
export interface DocumentStructuredParagraphMenuContext {
  paragraphId: string; afterBlockId?: string; disabled: boolean; inColumns?: boolean; belowInColumns?: boolean
  activeBlockId?: string; activeType?: DocumentStructureRibbonContext['type']; codeAllowed?: boolean
  columns?: DocumentColumnGroup; columnId?: string; columnGroupId?: string
}

/** The six fixed paragraph registrations and table-below registration share checked content commands. */
export function documentStructuredMountedMenuItems(context: DocumentStructuredParagraphMenuContext, mount: DocumentStructuredParagraphMenuMount): RibbonItem[] {
  const below = ['layout-below', 'content-insert', 'quick-insert-below', 'table-below'].includes(mount)
  const after = { afterBlockId: context.afterBlockId ?? context.activeBlockId ?? context.paragraphId }
  const quick = mount === 'quick-insert' || mount === 'quick-insert-below' || mount === 'quick-edit'
  const order = quick ? ['code', 'quote', 'callout'] as const : ['callout', 'quote', 'code'] as const
  const structures: RibbonItem[] = order.map(blockType => {
    const convert = !below && context.activeBlockId && context.activeType !== 'columns'
    const hidden = mount === 'quick-edit' && (context.activeType === 'callout' && blockType !== 'callout'
      || (context.activeType === 'code' || context.activeType === 'quote') && blockType === 'callout')
    return { id: `doc-context-pop-${mount}-${blockType}`, type: 'button', title: { callout: '标注块', quote: '引用块', code: '代码块' }[blockType],
      disabled: context.disabled || !below && blockType === 'code' && context.codeAllowed === false, hidden,
      activated: !below && context.activeType === blockType,
      commandId: below ? 'doc.command.insert-structured-block' : convert ? 'doc.command.convert-structured-block' : 'doc.command.wrap-structured-block',
      params: { blockType, ...(below ? after : convert ? { blockId: context.activeBlockId } : { blockIds: [context.paragraphId] }) } }
  })
  const columnDisabled = context.disabled || (below ? context.belowInColumns ?? context.inColumns : context.inColumns)
  const columns: RibbonItem = { id: `doc-context-pop-${mount}-columns`, type: 'menu', title: '分栏', disabled: columnDisabled,
    children: [2, 3, 4, 5].map(columnCount => ({ id: `doc-context-pop-${mount}-columns-${columnCount}`, type: 'button', title: `${columnCount} 栏`,
      disabled: columnDisabled, commandId: below ? 'doc.command.insert-column-group' : 'doc.command.wrap-column-group',
      params: { columnCount, ...(below ? after : { blockIds: [context.paragraphId] }) } })) }
  const imageMount = ['layout-below', 'content-insert', 'empty-paragraph', 'quick-insert', 'quick-insert-below'].includes(mount)
  const image: RibbonItem = { id: `doc-context-pop-${mount}-image`, type: 'button', title: '图片', icon: ImagePlus,
    order: quick ? 1 : 5, disabled: context.disabled, commandId: 'doc.command.insert-float-image',
    ...(below ? { params: { paragraphMenuPlacement: 'below' } } : {}) }
  if (mount !== 'quick-edit') return imageMount ? quick ? [columns, image, ...structures] : [columns, ...structures, image] : [columns, ...structures]
  const groupId = context.columnGroupId ?? (context.activeType === 'columns' ? context.activeBlockId : undefined)
  const operations = context.columns && context.columnId && groupId
    ? documentStructureContextItems({ blockId: groupId, type: 'columns', disabled: context.disabled, columns: context.columns, columnId: context.columnId }) : []
  const columnActions = ['doc-context-column-add', 'doc-context-column-delete', 'doc-context-structured-remove', 'doc-context-column-equal']
    .flatMap(id => { const item = operations.find(item => item.id === id); return item ? [{ ...item, ...(id === 'doc-context-structured-remove' ? { title: '删除多栏组和内容' } : {}) }] : [] })
  return [...structures, ...columnActions]
}

export function documentStructuredParagraphMenuItems(context: { paragraphId: string; disabled: boolean; inColumns?: boolean }): RibbonItem[] {
  const entries = (below: boolean): RibbonItem[] => {
    const suffix = below ? 'below' : 'current', after = below ? { afterBlockId: context.paragraphId } : {}
    return [
      ...(['callout', 'quote', 'code'] as const).map(blockType => ({ id: `doc-context-insert-${blockType}-${suffix}`, type: 'button' as const,
        title: { callout: '标注块', quote: '引用块', code: '代码块' }[blockType], disabled: context.disabled,
        commandId: below ? 'doc.command.insert-structured-block' : 'doc.command.wrap-structured-block', params: { blockType, ...after } })),
      { id: `doc-context-insert-columns-${suffix}`, type: 'menu', title: '分栏', disabled: context.disabled || context.inColumns,
        children: [2, 3, 4, 5].map(columnCount => ({ id: `doc-context-insert-columns-${suffix}-${columnCount}`, type: 'button', title: `${columnCount} 栏`,
          disabled: context.disabled || context.inColumns, commandId: below ? 'doc.command.insert-column-group' : 'doc.command.wrap-column-group', params: { columnCount, ...after } })) },
    ]
  }
  return [
    { id: 'doc-context-structured-insert-current', type: 'menu', title: '将选段设为', disabled: context.disabled, children: entries(false) },
    { id: 'doc-context-structured-insert-below', type: 'menu', title: '在段落下方插入', disabled: context.disabled, children: entries(true) },
  ]
}

/** Menus, floating controls and the Ribbon all enter the same checked commands. */
export function documentStructureCommandFromMenu(id: string, value: unknown): DocumentStructureCommand | undefined {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined
  const params = value as Record<string, unknown>, blockId = typeof params.blockId === 'string' && params.blockId.trim() ? params.blockId : undefined
  if (id === 'doc.command.convert-structured-block' && blockId && ['quote', 'callout', 'code'].includes(String(params.blockType))) {
    return { id, blockId, blockType: params.blockType as 'quote' | 'callout' | 'code' }
  }
  if ((id === 'doc.command.unwrap-structured-block' || id === 'doc.command.exit-structured-block' || id === 'doc.command.remove-structured-block') && blockId) return { id, blockId }
  if (id === 'doc.command.wrap-structured-block' || id === 'doc.command.insert-structured-block') {
    const blockType = params.blockType
    if (blockType !== 'quote' && blockType !== 'callout' && blockType !== 'code') return undefined
    if (id === 'doc.command.wrap-structured-block') {
      if (params.blockIds !== undefined && (!Array.isArray(params.blockIds) || !params.blockIds.length || params.blockIds.some(id => typeof id !== 'string' || !id.trim()))) return undefined
      return { id, blockType, ...(Array.isArray(params.blockIds) ? { blockIds: [...params.blockIds] } : {}) }
    }
    return typeof params.afterBlockId === 'string' && params.afterBlockId.trim() ? { id, blockType, afterBlockId: params.afterBlockId } : undefined
  }
  if (id === 'doc.command.wrap-column-group' || id === 'doc.command.insert-column-group') {
    const count = params.columnCount
    if (typeof count !== 'number' || !Number.isInteger(count) || count < DOCUMENT_COLUMN_GROUP_MIN_COLUMNS || count > DOCUMENT_COLUMN_GROUP_MAX_COLUMNS) return undefined
    if (id === 'doc.command.wrap-column-group') {
      if (params.blockIds !== undefined && (!Array.isArray(params.blockIds) || !params.blockIds.length || params.blockIds.some(id => typeof id !== 'string' || !id.trim()))) return undefined
      return { id, columnCount: count, ...(Array.isArray(params.blockIds) ? { blockIds: [...params.blockIds] } : {}) }
    }
    return typeof params.afterBlockId === 'string' && params.afterBlockId.trim() ? { id, columnCount: count, afterBlockId: params.afterBlockId } : undefined
  }
  if (!blockId) return undefined
  if (id === 'doc.command.set-code-font-family' && typeof params.fontFamily === 'string' && params.fontFamily.trim() && params.fontFamily.length <= 256) return { id, blockId, fontFamily: params.fontFamily }
  if (id === 'doc.command.update-code-settings' && params.patch && typeof params.patch === 'object' && !Array.isArray(params.patch)) {
    const patch = params.patch as Record<string, unknown>
    if (!Object.keys(patch).length || Object.keys(patch).some(key => !['language', 'wrap', 'showLineNumbers', 'tabSize'].includes(key))
      || patch.language !== undefined && (typeof patch.language !== 'string' || !patch.language.trim() || patch.language.length > 256)
      || patch.wrap !== undefined && typeof patch.wrap !== 'boolean' || patch.showLineNumbers !== undefined && typeof patch.showLineNumbers !== 'boolean'
      || patch.tabSize !== undefined && (typeof patch.tabSize !== 'number' || !Number.isInteger(patch.tabSize) || patch.tabSize < 1 || patch.tabSize > 16)) return undefined
    return { id, blockId, patch: { ...patch } as Partial<DocumentCodeConfig> }
  }
  if (id === 'doc.command.add-column' && typeof params.columnId === 'string' && (params.position === 'before' || params.position === 'after')) return { id, blockId, columnId: params.columnId, position: params.position }
  if (id === 'doc.command.remove-column' && typeof params.columnId === 'string' && typeof params.keepContents === 'boolean') return { id, blockId, columnId: params.columnId, keepContents: params.keepContents }
  if (id === 'doc.command.update-column-group' && Array.isArray(params.widthRatios) && params.widthRatios.length >= DOCUMENT_COLUMN_GROUP_MIN_COLUMNS
    && params.widthRatios.length <= DOCUMENT_COLUMN_GROUP_MAX_COLUMNS && params.widthRatios.every(ratio => typeof ratio === 'number' && Number.isFinite(ratio) && ratio > 0 && ratio <= 1e12)) {
    return { id, blockId, widthRatios: [...params.widthRatios] }
  }
  return undefined
}

export function createDocumentStructureRibbon(context: DocumentStructureRibbonContext): RibbonTab {
  const title = { quote: '引用块', callout: '标注块', code: '代码块', columns: '多栏组' }[context.type]
  const item = { type: 'button' as const, disabled: context.disabled, params: { blockId: context.blockId } }
  return { id: 'structured-block', title, order: 120, contextual: true, groups: [{ id: 'structured-block-actions', title: '操作', order: 0, items: [
    { ...item, id: 'doc-structured-properties', title: '属性', icon: FileCog,
      commandId: 'doc.operation.edit-structured-block', order: 0, gridLayout: { row: 1, column: 1, showLabel: true } },
    { ...item, id: 'doc-structured-unwrap', title: '保留内容并移除容器', icon: Layers2, commandId: 'doc.command.unwrap-structured-block', order: 1, gridLayout: { row: 1, column: 2, showLabel: true } },
    { ...item, id: 'doc-structured-exit', title: '在块后继续输入', icon: CornerDownLeft, commandId: 'doc.command.exit-structured-block', order: 2, gridLayout: { row: 2, column: 2, showLabel: true } },
    { ...item, id: 'doc-structured-remove', title: '删除块和内容', icon: Trash2, commandId: 'doc.command.remove-structured-block', order: 3, gridLayout: { row: 2, column: 1, showLabel: true } },
  ] }] }
}
