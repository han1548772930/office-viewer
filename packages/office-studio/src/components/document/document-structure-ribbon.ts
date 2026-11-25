import { CornerDownLeft, FileCog, Layers2, Trash2 } from 'lucide-vue-next'
import type { RibbonTab } from '../../platform/ribbon/ribbon-schema'

export interface DocumentStructureRibbonContext {
  blockId: string
  type: 'quote' | 'callout' | 'code' | 'columns'
  disabled: boolean
}

export function documentStructureContextItems(context: DocumentStructureRibbonContext) {
  const title = { quote: '引用块', callout: '标注块', code: '代码块', columns: '多栏组' }[context.type]
  return [
    { id: 'doc-context-structured-properties', title: `${title}属性`, commandId: 'doc.operation.edit-structured-block', disabled: context.disabled, params: { blockId: context.blockId } },
    { id: 'doc-context-structured-unwrap', title: '保留内容并移除容器', commandId: 'doc.command.unwrap-structured-block', disabled: context.disabled, params: { blockId: context.blockId } },
    { id: 'doc-context-structured-exit', title: '在块后继续输入', commandId: 'doc.command.exit-structured-block', disabled: context.disabled, params: { blockId: context.blockId } },
    { id: 'doc-context-structured-remove', title: '删除块和内容', commandId: 'doc.command.remove-structured-block', disabled: context.disabled, params: { blockId: context.blockId } },
  ]
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
