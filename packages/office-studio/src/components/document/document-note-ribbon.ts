import { ArrowLeft, ArrowRight, BookMarked, FileCog, NotebookPen, Superscript } from 'lucide-vue-next'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema'
import type { DocumentNoteActionState } from '../../features/document/commands/document-note-command-state'
import type { DocumentNoteKind } from '../../features/document/notes/document-note-types'

export interface DocumentNoteRibbonContext {
  visible: boolean; insert: DocumentNoteActionState; settings: DocumentNoteActionState
  adjacent: Record<DocumentNoteKind, { previous?: string; next?: string }>
  disabled?: boolean
}

/** Univer 1.6.0 docs-reference-ui: eight roots, References order 3, two groups in order 1/2. */
export function createDocumentNoteRibbon(context: DocumentNoteRibbonContext): RibbonTab | undefined {
  if (!context.visible) return undefined
  const group = (kind: DocumentNoteKind, order: number) => {
    const label = kind === 'footnote' ? '脚注' : '尾注'
    const items: RibbonItem[] = [
      { id: `docs.operation.insert-${kind}`, commandId: `docs.operation.insert-${kind}`, title: `插入${label}`, type: 'button',
        icon: kind === 'footnote' ? Superscript : BookMarked, order: 0, disabled: !context.insert.available, disabledReason: context.insert.reason,
        gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } },
      { id: `docs.operation.previous-${kind}`, commandId: `docs.operation.previous-${kind}`, title: '上一条', type: 'button', icon: ArrowLeft,
        order: 1, disabled: context.disabled || !context.adjacent[kind].previous, disabledReason: '当前位置之前没有此类注释。',
        gridLayout: { row: 1, column: 2, showLabel: true } },
      { id: `docs.operation.next-${kind}`, commandId: `docs.operation.next-${kind}`, title: '下一条', type: 'button', icon: ArrowRight,
        order: 2, disabled: context.disabled || !context.adjacent[kind].next, disabledReason: '当前位置之后没有此类注释。',
        gridLayout: { row: 1, column: 3, showLabel: true } },
      { id: `docs.operation.open-${kind}-settings`, commandId: `docs.operation.open-${kind}-settings`, title: `${label}设置`, type: 'button', icon: FileCog,
        order: 3, disabled: context.disabled, gridLayout: { row: 2, column: 2, columnSpan: 2, showLabel: true } },
    ]
    return { id: `docs-reference-ui.${kind}`, title: label, order, items }
  }
  return { id: 'references', title: '引用', order: 3, groups: [group('footnote', 1), group('endnote', 2),
    { id: 'doc-note-management', title: '管理', order: 100, items: [{ id: 'doc-notes-manage', type: 'button', title: '管理脚注和尾注',
      commandId: 'doc.operation.open-notes', icon: NotebookPen, disabled: context.disabled,
      gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }] }] }
}

export function documentNoteReturnContextItem(kind: DocumentNoteKind, noteId: string, disabled: boolean): RibbonItem {
  return { id: 'docs.operation.navigate-footnote', type: 'button', title: '返回引用', commandId: 'docs.operation.navigate-footnote',
    params: { kind, noteId, toReference: true }, disabled, order: 0 }
}
