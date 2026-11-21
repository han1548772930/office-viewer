import type { GridHit } from '../../rendering/spreadsheet/layout.ts'

export interface SheetContextMenuItem {
  id: string
  label: string
  commandId: string
  params?: unknown
  disabled?: boolean
  separatorBefore?: boolean
}

export interface SheetContextMenuContext {
  sheetId: string
  canDeleteRow: boolean
  canDeleteColumn: boolean
  frozen: boolean
  hasHiddenRows: boolean
  hasHiddenColumns: boolean
}

function targetParams(hit: GridHit, sheetId: string, axis: 'row' | 'column', offset = 0) {
  const index = axis === 'row'
    ? hit.kind === 'row' || hit.kind === 'cell' ? hit.row : undefined
    : hit.kind === 'column' || hit.kind === 'cell' ? hit.column : undefined
  return index === undefined ? undefined : { sheetId, index: index + offset }
}

export function createSheetContextMenuItems(hit: GridHit, options: SheetContextMenuContext): SheetContextMenuItem[] {
  if (hit.kind === 'resize-row' || hit.kind === 'resize-column') return []

  const clipboard: SheetContextMenuItem[] = [
    { id: 'cut', label: '剪切', commandId: 'sheet.command.cut' },
    { id: 'copy', label: '复制', commandId: 'sheet.command.copy' },
    { id: 'copy-formulas', label: '仅复制公式', commandId: 'sheet.command.copy-formulas' },
    { id: 'paste', label: '粘贴', commandId: 'sheet.command.paste' },
    { id: 'paste-values', label: '仅粘贴值', commandId: 'sheet.command.paste-values' },
    { id: 'paste-formulas', label: '仅粘贴公式', commandId: 'sheet.command.paste-formulas' },
    { id: 'paste-formats', label: '仅粘贴格式', commandId: 'sheet.command.paste-formats' },
    { id: 'paste-column-widths', label: '仅粘贴列宽', commandId: 'sheet.command.paste-column-widths' },
    { id: 'paste-except-borders', label: '保留目标边框粘贴', commandId: 'sheet.command.paste-except-borders' },
    { id: 'clear-content', label: '清除内容', commandId: 'sheet.command.clear-content', separatorBefore: true },
  ]
  const freeze: SheetContextMenuItem[] = [
    { id: 'freeze-selection', label: '冻结到当前选区', commandId: 'sheet.command.freeze-selection', params: hit.kind === 'row' ? { sheetId: options.sheetId, axis: 'row', index: hit.row } : hit.kind === 'column' ? { sheetId: options.sheetId, axis: 'column', index: hit.column } : undefined },
    { id: 'freeze-first-row', label: '冻结首行', commandId: 'sheet.command.freeze-first-row' },
    { id: 'freeze-first-column', label: '冻结首列', commandId: 'sheet.command.freeze-first-column' },
    { id: 'unfreeze', label: '取消冻结', commandId: 'sheet.command.unfreeze', disabled: !options.frozen },
  ]
  const protect: SheetContextMenuItem = {
    id: 'protect-range', label: '保护范围', commandId: 'sheet.operation.open-range-protection',
    params: hit.kind === 'row' ? { sheetId: options.sheetId, range: `A${hit.row + 1}:${columnLabel(16383)}${hit.row + 1}` }
      : hit.kind === 'column' ? { sheetId: options.sheetId, range: `${columnLabel(hit.column)}1:${columnLabel(hit.column)}1048576` }
        : hit.kind === 'cell' ? { sheetId: options.sheetId, range: hit.key } : { sheetId: options.sheetId },
    separatorBefore: true,
  }
  const rowTarget = targetParams(hit, options.sheetId, 'row')
  const columnTarget = targetParams(hit, options.sheetId, 'column')

  if (hit.kind === 'row') return [
    { id: 'insert-row-above', label: '在上方插入行', commandId: 'sheet.command.insert-row', params: rowTarget },
    { id: 'insert-row-below', label: '在下方插入行', commandId: 'sheet.command.insert-row', params: targetParams(hit, options.sheetId, 'row', 1) },
    { id: 'delete-row', label: '删除行', commandId: 'sheet.command.delete-row', params: rowTarget, disabled: !options.canDeleteRow },
    { id: 'hide-row', label: '隐藏行', commandId: 'sheet.command.hide-row', params: rowTarget, separatorBefore: true },
    { id: 'show-hidden-rows', label: '显示隐藏行', commandId: 'sheet.command.show-hidden-rows', params: { sheetId: options.sheetId }, disabled: !options.hasHiddenRows },
    { id: 'auto-fit-row', label: '自动调整行高', commandId: 'sheet.command.auto-fit-row', params: rowTarget },
    ...freeze,
    protect,
  ]
  if (hit.kind === 'column') return [
    { id: 'insert-column-left', label: '在左侧插入列', commandId: 'sheet.command.insert-column', params: columnTarget },
    { id: 'insert-column-right', label: '在右侧插入列', commandId: 'sheet.command.insert-column', params: targetParams(hit, options.sheetId, 'column', 1) },
    { id: 'delete-column', label: '删除列', commandId: 'sheet.command.delete-column', params: columnTarget, disabled: !options.canDeleteColumn },
    { id: 'hide-column', label: '隐藏列', commandId: 'sheet.command.hide-column', params: columnTarget, separatorBefore: true },
    { id: 'show-hidden-columns', label: '显示隐藏列', commandId: 'sheet.command.show-hidden-columns', params: { sheetId: options.sheetId }, disabled: !options.hasHiddenColumns },
    { id: 'auto-fit-column', label: '自动调整列宽', commandId: 'sheet.command.auto-fit-column', params: columnTarget },
    ...freeze,
    protect,
  ]
  if (hit.kind === 'all') return [
    ...clipboard,
    ...freeze,
    protect,
  ]
  return [
    ...clipboard,
    { id: 'insert-row', label: '在此处插入行', commandId: 'sheet.command.insert-row', params: rowTarget, separatorBefore: true },
    { id: 'insert-column', label: '在此处插入列', commandId: 'sheet.command.insert-column', params: columnTarget },
    { id: 'delete-row', label: '删除此行', commandId: 'sheet.command.delete-row', params: rowTarget, disabled: !options.canDeleteRow },
    { id: 'delete-column', label: '删除此列', commandId: 'sheet.command.delete-column', params: columnTarget, disabled: !options.canDeleteColumn },
    ...freeze,
    protect,
  ]
}

function columnLabel(index: number) {
  let column = index + 1
  let label = ''
  while (column > 0) {
    const remainder = (column - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    column = Math.floor((column - 1) / 26)
  }
  return label
}
