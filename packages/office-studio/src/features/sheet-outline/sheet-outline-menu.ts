import { Group } from 'lucide-vue-next'
import type { OutlineAxis } from '../../domain/model.ts'
import type { RibbonItem } from '../../platform/ribbon/ribbon-schema.ts'
import { getSheetOutlinePromptChoices, type SheetOutlineCommandContext } from './sheet-outline-command.ts'
import { canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'

export const SHEET_OUTLINE_TOOLBAR_GROUP_MENU_ID = 'sheet.toolbar.dimension-outline.group'
export const SHEET_OUTLINE_CONTEXT_MENU_ID = 'sheet.contextMenu.dimension-outline'
export const SHEET_OUTLINE_SHORTCUTS = [
  { commandId: 'sheet.command.outline-selected-dimension-shortcut', key: 'ArrowRight', altKey: true, shiftKey: true, macKey: 'k', macPrimary: true },
  { commandId: 'sheet.command.clear-selected-dimension-outlines-shortcut', key: 'ArrowLeft', altKey: true, shiftKey: true, macKey: 'j', macPrimary: true },
] as const
export function createSheetOutlineRibbonItem(context: SheetOutlineCommandContext | undefined): RibbonItem {
  const groups = context ? getSheetOutlinePromptChoices(context, 'group') : [], clears = context ? getSheetOutlinePromptChoices(context, 'remove') : []
  const selection = context?.selections?.length === 1 ? context.selections[0] : undefined
  const available = (choices: typeof groups) => choices.some(choice => !choice.disabled
    && (selection?.rangeType !== 1 && selection?.rangeType !== 2 || choice.axis === (selection.rangeType === 1 ? 'row' : 'column')))
  return { id: SHEET_OUTLINE_TOOLBAR_GROUP_MENU_ID, type: 'select', title: '组合', tooltip: '组合与取消行列分组', icon: Group,
    order: 3, gridLayout: { row: 2, column: 2 }, value: null, placeholder: '组合', width: 96,
    hidden: !selection, disabled: !context || !canEditWorksheetProtection(context.sheet, context.userId), options: [
      { label: '组合', value: 'group', commandId: 'sheet.command.outline-selected-dimension-shortcut', params: {}, disabled: !available(groups) },
      { label: '取消组合', value: 'ungroup', commandId: 'sheet.command.clear-selected-dimension-outlines-shortcut', params: {}, disabled: !available(clears) },
    ] }
}
export function createSheetOutlineContextMenuItem(context: SheetOutlineCommandContext | undefined, axis: OutlineAxis): RibbonItem {
  const group = context ? getSheetOutlinePromptChoices(context, 'group').find(choice => choice.axis === axis) : undefined
  const clear = context ? getSheetOutlinePromptChoices(context, 'remove').find(choice => choice.axis === axis) : undefined
  const suffix = axis === 'row' ? 'rows' : 'columns'
  return { id: SHEET_OUTLINE_CONTEXT_MENU_ID, type: 'menu', title: '组合', icon: Group, order: 8,
    hidden: context?.selections?.length !== 1, disabled: !context || !canEditWorksheetProtection(context.sheet, context.userId), children: [
      { id: `sheet.command.outline-selected-${suffix}`, commandId: `sheet.command.outline-selected-${suffix}`, type: 'button',
        title: group?.label ?? (axis === 'row' ? '组合行' : '组合列'), params: { value: axis }, disabled: group?.disabled ?? true, disabledReason: group?.reason },
      { id: `sheet.command.clear-selected-${suffix}-outlines`, commandId: `sheet.command.clear-selected-${suffix}-outlines`, type: 'button',
        title: clear?.label ?? (axis === 'row' ? '取消行分组' : '取消列分组'), params: { value: axis }, disabled: clear?.disabled ?? true, disabledReason: clear?.reason },
    ] }
}
