import type { Component } from 'vue'
import {
  AlignCenter, AlignLeft, AlignRight, ArrowDownWideNarrow, ArrowUpWideNarrow, Columns3, Grid2X2, ListFilter,
  Merge, PaintBucket, Palette, PanelTop, Pin, Rows3, SlidersHorizontal, Split, Square, Table2, Trash2,
} from 'lucide-vue-next'
import type { DocumentModel } from '../../domain/model'
import { findDocumentBlock, tableForBlock, type DocumentSelection } from '../../features/document/model/document-model'
import { documentTableCellsInRange, documentTableGrid, type DocumentTableCellRange } from '../../features/document/model/document-table-grid'
import { documentTableCellBorder, documentTableCellFill, DOCUMENT_TABLE_BORDER_PRESETS, DOCUMENT_TABLE_DEFAULT_BORDER,
  type DocumentTableBorderPreset } from '../../features/document/model/document-table-style'
import { documentTableCommandState, documentTableSelection } from '../../features/document/commands/table-commands'
import type { DocumentTableCommand } from '../../features/document/commands/table-command-types'
import type { DocumentPermissionContext } from '../../features/document/permissions/document-permissions'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema'
import DocumentTableColorControl from './DocumentTableColorControl.vue'

export interface DocumentTableRibbonContext {
  model: DocumentModel; selection: DocumentSelection; permission: DocumentPermissionContext
  mode: 'traditional' | 'modern'; composing?: boolean; floatingToolbarOpen: boolean
  lastBackgroundColor?: string; lastBorderColor?: string; measuredRowHeights?: number[]
}
export const DOCUMENT_TABLE_BORDER_LABELS: Record<DocumentTableBorderPreset, string> = {
  all: '所有边框', inner: '内部边框', outer: '外侧边框', top: '上边框', bottom: '下边框', left: '左边框', right: '右边框',
  none: '无边框', innerHorizontal: '内部横边框', innerVertical: '内部竖边框',
}
const BORDER_WIDTHS_PT = [0, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]

function action(context: DocumentTableRibbonContext, id: string, title: string, icon: Component, command: DocumentTableCommand,
  row: number, column: number, order: number, showLabel = false): RibbonItem {
  const state = context.composing ? { available: false, reason: '请先完成当前输入法输入' }
    : documentTableCommandState(context.model, context.selection, command, context.permission)
  const { id: commandId, ...params } = command
  return { id, type: 'button', title, icon, commandId, params, order, disabled: !state.available, disabledReason: state.reason,
    gridLayout: { row, column, showLabel }, showLabel }
}

export function createDocumentTableRibbon(context: DocumentTableRibbonContext): RibbonTab | undefined {
  const selected = documentTableSelection(context.model, context.selection)
  if (!selected) return undefined
  const tableId = selected.blockId, block = findDocumentBlock(context.model, tableId)!, source = tableForBlock(block)!, grid = documentTableGrid(source)
  const range: DocumentTableCellRange = { startRow: selected.startRow, endRow: selected.endRow, startColumn: selected.startColumn, endColumn: selected.endColumn }
  const target = { tableId, range }, owners = documentTableCellsInRange(grid, range)
  const uniform = <T>(values: readonly T[]) => values.every(value => value === values[0]) ? values[0] : undefined
  const fills = owners.map(owner => documentTableCellFill(source, owner.cell))
  const borders = owners.flatMap(owner => (['top', 'right', 'bottom', 'left'] as const).map(side => documentTableCellBorder(grid, owner, side)))
  const background = uniform(fills), borderColor = uniform(borders.map(border => border.color)), borderStyle = uniform(borders.map(border => border.style))
  const borderWidth = uniform(borders.map(border => border.width)), pointWidth = borderWidth === undefined ? undefined : borderWidth * 0.75
  const button = (id: string, title: string, icon: Component, command: DocumentTableCommand, row: number, column: number, order: number, label = false) => action(context, id, title, icon, command, row, column, order, label)
  const menu = (id: string, title: string, icon: Component, row: number, column: number, order: number, children: RibbonItem[], label = false): RibbonItem => ({
    id, title, icon, type: 'menu', order, gridLayout: { row, column, showLabel: label }, showLabel: label, children,
    disabled: children.every(child => child.disabled), disabledReason: children.find(child => child.disabledReason)?.disabledReason,
  })
  const operation = (id: string, title: string, icon: Component, commandId: string, params: Record<string, unknown>, probe: DocumentTableCommand,
    row: number, column: number, order: number): RibbonItem => ({ ...button(id, title, icon, probe, row, column, order, true), commandId, params })
  const colors = (field: 'background' | 'border', row: number, column: number, order: number): RibbonItem => {
    const id = field === 'background' ? 'doc-table-background' : 'doc-table-border-color', title = field === 'background' ? '单元格背景' : '边框颜色'
    const command: DocumentTableCommand = field === 'background' ? { id: 'doc.command.table-set-background', ...target, color: background ?? '#ffffff' }
      : { id: 'doc.command.table-set-border-color', ...target, color: borderColor ?? DOCUMENT_TABLE_DEFAULT_BORDER.color }
    const item = button(id, title, field === 'background' ? PaintBucket : Palette, command, row, column, order)
    return { ...item, type: 'split', inlineComponent: DocumentTableColorControl, popupCommandId: command.id,
      popupProps: { title, field, icon: item.icon, defaultColor: field === 'background' ? '#ffffff' : DOCUMENT_TABLE_DEFAULT_BORDER.color,
        lastColor: field === 'background' ? context.lastBackgroundColor ?? background : context.lastBorderColor ?? borderColor,
        selectedColor: field === 'background' ? background : borderColor, mixed: field === 'background' ? background === undefined : borderColor === undefined, ...target },
      children: [button(`${id}-clear`, field === 'background' ? '无背景' : '自动边框颜色', item.icon!, { ...command, color: null }, 1, 1, 0)] }
  }
  const widths = button('doc-table-border-width', '边框宽度', Square, { id: 'doc.command.table-set-border-width', ...target, width: borderWidth ?? 1 }, 2, 2, 3)
  const styles = button('doc-table-border-style', '边框样式', Square, { id: 'doc.command.table-set-border-style', ...target, style: borderStyle ?? 'solid' }, 1, 3, 4)
  return { id: 'table', title: '表格', contextual: true, order: 100, groups: [
    { id: 'doc-table-style', title: '样式', order: 0, items: [
      colors('background', 1, 1, 0),
      menu('doc-table-border', '边框', Grid2X2, 2, 3, 1, DOCUMENT_TABLE_BORDER_PRESETS.map((preset, index) => button(`doc-table-border-${preset}`, DOCUMENT_TABLE_BORDER_LABELS[preset], Grid2X2,
        { id: 'doc.command.table-set-border', ...target, preset, border: { width: borderWidth ?? 1, color: borderColor ?? DOCUMENT_TABLE_DEFAULT_BORDER.color, style: borderStyle ?? 'solid' } }, 1, 1, index))),
      colors('border', 1, 2, 2),
      { ...widths, type: 'select', width: 70, value: BORDER_WIDTHS_PT.find(value => pointWidth !== undefined && Math.abs(value - pointWidth) < 1e-6) ?? null,
        placeholder: '混合', options: BORDER_WIDTHS_PT.map(value => ({ value, label: `${value} pt`, disabled: !documentTableCommandState(context.model, context.selection,
          { id: 'doc.command.table-set-border-width', ...target, width: value * 4 / 3 }, context.permission).available })), valueParams: value => ({ ...target, width: Number(value) * 4 / 3 }) },
      { ...styles, type: 'select', width: 72, value: borderStyle ?? null, placeholder: '混合', options: [{ value: 'solid', label: '实线' }, { value: 'dot', label: '点线' }, { value: 'dash', label: '虚线' }], valueParams: value => ({ ...target, style: value }) },
    ] },
    { id: 'doc-table-layout', title: '布局', order: 1, items: [
      menu('doc-table-insert', '插入', Table2, 1, 1, 0, [
        button('doc-table-row-before', '在上方插入行', Rows3, { id: 'doc.command.table-insert-row', ...target, position: 'before' }, 1, 1, 0),
        button('doc-table-row-after', '在下方插入行', Rows3, { id: 'doc.command.table-insert-row', ...target, position: 'after' }, 1, 1, 1),
        button('doc-table-column-before', '在左侧插入列', Columns3, { id: 'doc.command.table-insert-column', ...target, position: 'before' }, 1, 1, 2),
        button('doc-table-column-after', '在右侧插入列', Columns3, { id: 'doc.command.table-insert-column', ...target, position: 'after' }, 1, 1, 3),
      ]),
      menu('doc-table-align', '整表对齐', source.style?.align === 'center' ? AlignCenter : source.style?.align === 'right' ? AlignRight : AlignLeft, 2, 1, 1,
        ([['left', '靠左', AlignLeft], ['center', '居中', AlignCenter], ['right', '靠右', AlignRight]] as const).map(([value, title, icon], index) => ({
          ...button(`doc-table-align-${value}`, title, icon, { id: 'doc.command.table-set-align', ...target, value }, 1, 1, index), activated: (source.style?.align ?? 'left') === value }))),
      button('doc-table-merge', '合并单元格', Merge, { id: 'doc.command.table-merge-cells', ...target }, 1, 2, 2),
      button('doc-table-unmerge', '拆分单元格', Split, { id: 'doc.command.table-unmerge-cells', ...target }, 2, 2, 3),
      button('doc-table-distribute-rows', '平均分布行', Rows3, { id: 'doc.command.table-distribute-rows', ...target,
        ...(context.measuredRowHeights ? { heights: context.measuredRowHeights } : {}) }, 1, 3, 4),
      button('doc-table-distribute-columns', '平均分布列', Columns3, { id: 'doc.command.table-distribute-columns', ...target }, 2, 3, 5),
    ] },
    { id: 'doc-table-data', title: '数据', order: 2, items: [
      menu('doc-table-sort', '排序', ListFilter, 1, 1, 0, [
        button('doc-table-sort-ascending', '升序', ArrowUpWideNarrow, { id: 'doc.command.table-sort', ...target, direction: 'ascending' }, 1, 1, 0),
        button('doc-table-sort-descending', '降序', ArrowDownWideNarrow, { id: 'doc.command.table-sort', ...target, direction: 'descending' }, 1, 1, 1),
      ], true),
      menu('doc-table-column-type', '列类型', SlidersHorizontal, 2, 1, 1, ([['text', '文本'], ['number', '数字'], ['date', '日期'], ['checkbox', '复选框'], ['dropdown', '下拉列表']] as const).map(([type, label], index) => {
        const command: DocumentTableCommand = { id: 'doc.command.table-set-column-type', ...target, config: { type, ...(type === 'dropdown' ? { dropdownOptions: source.columnTypes?.[range.startColumn]?.dropdownOptions ?? [] } : {}) } }
        const item = button(`doc-table-type-${type}`, label, SlidersHorizontal, command, 1, 1, index)
        return { ...item, activated: (source.columnTypes?.[range.startColumn]?.type ?? 'text') === type,
          ...(type === 'dropdown' ? { commandId: 'doc.operation.open-table-properties', params: { ...target, kind: 'type', type } } : {}) }
      }), true),
    ] },
    { id: 'doc-table-options', title: '选项', order: 3, items: [
      { ...button('doc-table-title', '插入标题行', PanelTop, { id: 'doc.command.table-insert-title-row', ...target }, 1, 1, 0, true), hidden: context.mode === 'modern' },
      { ...operation('doc-table-header', '固定表头', Pin, 'doc.operation.open-table-properties', { ...target, kind: 'header' },
        { id: 'doc.command.table-set-header-row-count', ...target, count: source.headerRowCount ?? 0 }, 2, 1, 1), hidden: context.mode === 'modern', activated: (source.headerRowCount ?? 0) > 0 },
      menu('doc-table-delete', '删除', Trash2, 1, 2, 2, [
        button('doc-table-delete-rows', '删除行', Rows3, { id: 'doc.command.table-delete-rows', ...target }, 1, 1, 0),
        button('doc-table-delete-columns', '删除列', Columns3, { id: 'doc.command.table-delete-columns', ...target }, 1, 1, 1),
        button('doc-table-delete-table', '删除表格', Table2, { id: 'doc.command.table-delete', ...target }, 1, 1, 2),
      ], true),
      { id: 'doc-table-floating-toolbar', type: 'button', title: '浮动工具栏', icon: SlidersHorizontal, commandId: 'doc.operation.toggle-table-floating-toolbar',
        params: target, activated: context.floatingToolbarOpen, order: 3, showLabel: true, gridLayout: { row: 2, column: 2, showLabel: true } },
    ] },
  ] }
}

/** Additional properties/move entry points are outside the 17 upstream Ribbon roots. */
export function documentTableExtraContextItems(context: DocumentTableRibbonContext): RibbonItem[] {
  const selected = documentTableSelection(context.model, context.selection)
  if (!selected) return []
  const target = { tableId: selected.blockId, range: { startRow: selected.startRow, endRow: selected.endRow, startColumn: selected.startColumn, endColumn: selected.endColumn } }
  return ([['table', '表格属性'], ['row', '行高与跨页'], ['column', '列宽'], ['cell', '单元格边距'], ['type', '列类型选项'], ['move', '移动表格或行列']] as const).map(([kind, title]) => ({
    id: `doc-table-properties-${kind}`, type: 'button', title, icon: SlidersHorizontal, commandId: 'doc.operation.open-table-properties', params: { ...target, kind },
    disabled: context.composing, disabledReason: context.composing ? '请先完成当前输入法输入' : undefined,
  }))
}
