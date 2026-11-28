import type { WorkbookModel } from '../../domain/model.ts'
import { isFormulaError } from '../formula/formula-types.ts'
import type { SheetCellInputDefaults, SheetCellInputPlan } from '../sheet-workbench/sheet-cell-input-command.ts'
import { sheetCheckboxRuleValues, sheetCheckboxState } from './validation-checkbox.ts'
import { sheetValidationCellValue, sheetValidationFormulaValue, sheetValidationSourceReader } from './validation-value-read.ts'
import { planSheetValidationWidgetValue, resolveSheetValidationWidgetTarget, sheetValidationWidgetTargetInput,
  validationWidgetRecord, type SheetValidationWidgetTargetInput } from './validation-widget-input.ts'

export const COMMIT_SHEET_CHECKBOX_INPUT_COMMAND = 'sheet.command.commit-checkbox-input'
export interface SheetCheckboxInput extends SheetValidationWidgetTargetInput { checked?: boolean }

export function sheetCheckboxInput(value: unknown): SheetCheckboxInput {
  const source = validationWidgetRecord(value, '复选框输入命令参数'), checked = source.checked
  if (checked !== undefined && typeof checked !== 'boolean') throw new TypeError('复选框状态必须是布尔值。')
  return { ...sheetValidationWidgetTargetInput(source), ...(checked === undefined ? {} : { checked }) }
}

/** A gesture changes one actual main cell. Custom values never pass through the ordinary text input parser. */
export function planSheetCheckboxInput(workbook: WorkbookModel, input: SheetCheckboxInput,
  defaults: SheetCellInputDefaults): SheetCellInputPlan {
  const target = resolveSheetValidationWidgetTarget(workbook, input, defaults)
  if ('result' in target) return target
  if (target.rule.type !== 'checkbox') throw new TypeError('此规则没有复选框控件。')
  const canReadSource = sheetValidationSourceReader(defaults.userId)
  const values = sheetCheckboxRuleValues(target.rule, (formula, anchor) =>
    sheetValidationFormulaValue(formula, workbook, target.sheet.id, anchor, canReadSource))
  const previous = target.sheet.cells[`${target.position.row}:${target.position.column}`]
  const actual = sheetValidationCellValue(workbook, target.sheet, target.position.row, target.position.column, canReadSource)
  const state = sheetCheckboxState(values, isFormulaError(actual) ? undefined : actual, !!previous?.richText?.length, !!previous?.image)
  if (!values || isFormulaError(actual) || !state.renderable) {
    return { position: target.position, result: { accepted: false, changed: false, issues: [],
      message: '当前内容或规则值不能通过复选框切换，请先编辑数据验证规则或单元格内容。' } }
  }
  const checked = input.checked ?? !state.checked
  if (checked === state.checked && state.valid) return { position: target.position, result: { accepted: true, changed: false, issues: [] } }
  const value = checked ? values.checked : values.unchecked
  return planSheetValidationWidgetValue(target, value, typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'text')
}
