import type { GridHit } from '../../rendering/spreadsheet/layout.ts'

export interface SheetContextMenuItem {
  id: string
  label: string
  commandId: string
  params?: unknown
  disabled?: boolean
  separatorBefore?: boolean
}

export function createSheetContextMenuItems(hit: GridHit, options: { canDeleteRow: boolean; canDeleteColumn: boolean; frozen: boolean }): SheetContextMenuItem[] {
  if (hit.kind === 'resize-row' || hit.kind === 'resize-column') return []

  const clipboard: SheetContextMenuItem[] = [
    { id: 'cut', label: '剪切', commandId: 'sheet.command.cut' },
    { id: 'copy', label: '复制', commandId: 'sheet.command.copy' },
    { id: 'paste', label: '粘贴', commandId: 'sheet.command.paste' },
    { id: 'paste-values', label: '仅粘贴值', commandId: 'sheet.command.paste-values' },
    { id: 'paste-formats', label: '仅粘贴格式', commandId: 'sheet.command.paste-formats' },
    { id: 'paste-column-widths', label: '仅粘贴列宽', commandId: 'sheet.command.paste-column-widths' },
    { id: 'paste-except-borders', label: '保留目标边框粘贴', commandId: 'sheet.command.paste-except-borders' },
    { id: 'clear-content', label: '清除内容', commandId: 'sheet.command.clear-content', separatorBefore: true },
  ]
  const freeze: SheetContextMenuItem[] = [
    { id: 'freeze-selection', label: '冻结到当前选区', commandId: 'sheet.command.freeze-selection' },
    { id: 'freeze-first-row', label: '冻结首行', commandId: 'sheet.command.freeze-first-row' },
    { id: 'freeze-first-column', label: '冻结首列', commandId: 'sheet.command.freeze-first-column' },
    { id: 'unfreeze', label: '取消冻结', commandId: 'sheet.command.unfreeze', disabled: !options.frozen },
  ]
  const protect: SheetContextMenuItem = { id: 'protect-range', label: '保护范围', commandId: 'sheet.operation.open-range-protection', separatorBefore: true }

  if (hit.kind === 'row') return [
    { id: 'insert-row', label: '在上方插入行', commandId: 'sheet.command.insert-row' },
    { id: 'delete-row', label: '删除行', commandId: 'sheet.command.delete-row', disabled: !options.canDeleteRow },
    ...freeze,
    protect,
  ]
  if (hit.kind === 'column') return [
    { id: 'insert-column', label: '在左侧插入列', commandId: 'sheet.command.insert-column' },
    { id: 'delete-column', label: '删除列', commandId: 'sheet.command.delete-column', disabled: !options.canDeleteColumn },
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
    { id: 'insert-row', label: '插入行', commandId: 'sheet.command.insert-row', separatorBefore: true },
    { id: 'insert-column', label: '插入列', commandId: 'sheet.command.insert-column' },
    { id: 'delete-row', label: '删除行', commandId: 'sheet.command.delete-row', disabled: !options.canDeleteRow },
    { id: 'delete-column', label: '删除列', commandId: 'sheet.command.delete-column', disabled: !options.canDeleteColumn },
    ...freeze,
    protect,
  ]
}
