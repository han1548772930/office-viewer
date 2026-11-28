import type { CellRange, SheetDataValidationOperator, SheetDataValidationRule, SheetDataValidationType } from '../../domain/model.ts'
import { parseSheetAddress } from '../sheet-workbench/sheet-workbench.ts'
import { createSheetDataValidationRule, parseValidationListItems, serializeValidationListItems } from './validation-engine.ts'
import { sheetOpaqueColorHex } from '../sheet-workbench/sheet-colors.ts'

export type ValidationRangeTarget = 'range' | 'list' | 'formula1' | 'formula2'

/** Rule coverage can be discontinuous; an operand must be one reference. */
export function validationPickedAddress(target: ValidationRangeTarget, ranges: readonly CellRange[]): string | undefined {
  if (!ranges.length || (target !== 'range' && ranges.length !== 1)) return undefined
  const addresses = ranges.map(range => formatValidationRange(range))
  return target === 'range' ? addresses.join(', ') : `=${addresses[0]}`
}
export interface ValidationListItem { id: string; label: string; color: string }
export function validationListItemsFromCsv(value: string, previous: readonly ValidationListItem[], createId: () => string = () => crypto.randomUUID()): ValidationListItem[] {
  const items = parseValidationListItems(`"${value}"`), used = new Set<string>()
  return items.flatMap(label => {
    const key = label.trim().toLocaleLowerCase()
    if (used.has(key)) return []
    used.add(key)
    const retained = previous.find(item => item.label.trim().toLocaleLowerCase() === key)
    return [{ id: retained?.id ?? createId(), label, color: retained?.color ?? '' }]
  })
}
export function moveValidationListItem(items: readonly ValidationListItem[], id: string, offset: -1 | 1): ValidationListItem[] {
  const result = items.map(item => ({ ...item })), index = result.findIndex(item => item.id === id), target = index + offset
  if (index < 0 || target < 0 || target >= result.length) return result
  const item = result.splice(index, 1)[0]!
  result.splice(target, 0, item)
  return result
}
export interface ValidationPanelSettings {
  type: SheetDataValidationType
  operator: SheetDataValidationOperator | ''
  formula1: string
  formula2: string
  allowBlank: boolean
  showDropDown: boolean
  showErrorMessage: boolean
  error?: string
  errorTitle?: string
  errorStyle?: SheetDataValidationRule['errorStyle']
  showInputMessage: boolean
  prompt?: string
  promptTitle?: string
  renderMode?: SheetDataValidationRule['renderMode']
  checkboxCustomValues: boolean
  listMode: 'custom' | 'reference'
  listItems: ValidationListItem[]
  listReference: string
}

const COMPARISON_TYPES: readonly SheetDataValidationType[] = ['whole', 'decimal', 'textLength', 'date', 'time']
export const VALIDATION_OPERATORS: readonly SheetDataValidationOperator[] = [
  'between', 'equal', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'notBetween', 'notEqual',
]

export function validationUsesOperator(type: SheetDataValidationType): boolean { return COMPARISON_TYPES.includes(type) }
export function validationUsesTwoFormulas(type: SheetDataValidationType, operator: ValidationPanelSettings['operator']): boolean {
  return validationUsesOperator(type) && (operator === 'between' || operator === 'notBetween')
}

export function validationPanelSettings(rule?: SheetDataValidationRule | null, initialType?: SheetDataValidationType): ValidationPanelSettings {
  const type = rule?.type ?? initialType ?? 'whole'
  const reference = rule?.formula1?.trim().startsWith('=') ?? false
  const labels = !reference && (type === 'list' || type === 'listMultiple') ? parseValidationListItems(rule?.formula1 ?? '') : []
  const colors = rule?.formula2?.split(',') ?? []
  return {
    type,
    operator: rule ? rule.operator ?? '' : validationUsesOperator(type) ? type === 'date' ? '' : 'between' : '',
    formula1: rule?.formula1 ?? '',
    formula2: rule?.formula2 ?? '',
    allowBlank: rule?.allowBlank ?? true,
    showDropDown: rule?.showDropDown ?? true,
    showErrorMessage: rule?.showErrorMessage ?? true,
    ...(rule?.error === undefined ? {} : { error: rule.error }),
    ...(rule?.errorTitle === undefined ? {} : { errorTitle: rule.errorTitle }),
    ...(rule?.errorStyle === undefined ? {} : { errorStyle: rule.errorStyle }),
    showInputMessage: rule?.showInputMessage ?? false,
    ...(rule?.prompt === undefined ? {} : { prompt: rule.prompt }),
    ...(rule?.promptTitle === undefined ? {} : { promptTitle: rule.promptTitle }),
    ...(rule?.renderMode === undefined ? {} : { renderMode: rule.renderMode }),
    checkboxCustomValues: type === 'checkbox' && (rule?.formula1 !== undefined || rule?.formula2 !== undefined),
    listMode: reference ? 'reference' : 'custom',
    listItems: labels.map((label, index) => ({ id: crypto.randomUUID(), label, color: colors[index] ?? '' })),
    listReference: reference ? rule?.formula1 ?? '=' : '=',
  }
}

export function changeValidationType(settings: ValidationPanelSettings, type: SheetDataValidationType): ValidationPanelSettings {
  if (settings.type === type) return settings
  const next = validationPanelSettings(null, type)
  if (settings.type.startsWith('list') && type.startsWith('list')) {
    return { ...settings, type, operator: '' }
  }
  return next
}

export function validationRuleFromPanel(
  settings: ValidationPanelSettings,
  ranges: readonly CellRange[],
  options: Partial<SheetDataValidationRule>,
  previous?: SheetDataValidationRule | null,
): SheetDataValidationRule {
  const draft: Omit<SheetDataValidationRule, 'id' | 'ranges'> & { id?: string } = {
    ...previous,
    allowBlank: settings.allowBlank,
    showDropDown: settings.showDropDown,
    showErrorMessage: settings.showErrorMessage,
    error: settings.error,
    errorTitle: settings.errorTitle,
    errorStyle: settings.errorStyle,
    showInputMessage: settings.showInputMessage,
    prompt: settings.prompt,
    promptTitle: settings.promptTitle,
    renderMode: settings.renderMode,
    ...options,
    type: settings.type,
    operator: undefined,
    formula1: undefined,
    formula2: undefined,
    formulaAnchor: undefined,
  }
  if (settings.type === 'list' || settings.type === 'listMultiple') {
    if (settings.listMode === 'reference') {
      draft.formula1 = settings.listReference.trim()
      draft.formula2 = settings.formula2 || (previous?.formula2 === '' ? '' : undefined)
    } else {
      const unique = new Map<string, ValidationListItem>()
      for (const item of settings.listItems) {
        const key = item.label.trim().toLocaleLowerCase()
        if (key && !unique.has(key)) unique.set(key, item)
      }
      const previousList = previous && (previous.type === 'list' || previous.type === 'listMultiple') ? previous : undefined
      const items = [...unique.values()], originalLabels = previousList && !previousList.formula1?.trim().startsWith('=') ? parseValidationListItems(previousList.formula1 ?? '') : []
      if (!items.length) throw new TypeError('请输入至少一个下拉选项。')
      const originalColors = previous?.formula2?.split(',') ?? []
      const sameLabels = previousList !== undefined
        && originalLabels.length === items.length && items.every((item, index) => item.label === originalLabels[index])
      const sameColors = sameLabels && items.every((item, index) => item.color === (originalColors[index] ?? ''))
      for (const item of items) {
        const originalIndex = originalLabels.findIndex(label => label.trim().toLocaleLowerCase() === item.label.trim().toLocaleLowerCase())
        if (item.color && item.color !== originalColors[originalIndex] && !sheetOpaqueColorHex(item.color)) throw new TypeError('列表颜色必须是有效的不透明颜色。')
      }
      draft.formula1 = sameLabels && previousList ? previousList.formula1 : serializeValidationListItems(items.map(item => item.label))
      draft.formula2 = sameColors && previousList ? previousList.formula2 : items.some(item => item.color) || previousList?.formula2 !== undefined
        ? items.map(item => item.color).join(',') : undefined
    }
  } else if (settings.type === 'checkbox') {
    if (settings.checkboxCustomValues) {
      draft.formula1 = settings.formula1
      draft.formula2 = settings.formula2
    }
  } else if (settings.type === 'custom') {
    draft.formula1 = settings.formula1.trim()
  } else if (validationUsesOperator(settings.type) && settings.operator) {
    draft.operator = settings.operator
    draft.formula1 = settings.formula1.trim()
    if (validationUsesTwoFormulas(settings.type, settings.operator)) draft.formula2 = settings.formula2.trim()
  }
  const relatedList = previous?.type.startsWith('list') && draft.type.startsWith('list')
  if (previous && (previous.type === draft.type || relatedList) && previous.formula1 === draft.formula1
    && (relatedList || previous.formula2 === draft.formula2)) draft.formulaAnchor = previous.formulaAnchor
  return createSheetDataValidationRule(draft, ranges, previous ?? undefined)
}

export function parseValidationRanges(value: string, sheetName?: string): CellRange[] {
  const tokens = value.match(/(?:'[^']*(?:''[^']*)*'|[^,;])+/g) ?? []
  const ranges = tokens.map(token => {
    const text = token.trim()
    const qualified = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(text)
    const qualifier = qualified?.[1]?.replaceAll("''", "'") ?? qualified?.[2]
    if (qualifier && qualifier !== sheetName) throw new RangeError('应用区域必须属于当前工作表。')
    const parsed = parseSheetAddress(qualified?.[3] ?? text, { rowCount: 1_048_576, columnCount: 16_384 })
    if (!parsed) throw new TypeError('请输入有效的单元格区域，例如 A1:B4、A:A 或 1:3。')
    return { startRow: parsed.range.startRow, startCol: parsed.range.startColumn, endRow: parsed.range.endRow, endCol: parsed.range.endColumn }
  })
  if (!ranges.length) throw new TypeError('请选择至少一个数据验证区域。')
  return ranges
}

export function formatValidationRange(range: CellRange, sheetName?: string): string {
  const first = `${columnLabel(range.startCol)}${range.startRow + 1}`
  const last = `${columnLabel(range.endCol)}${range.endRow + 1}`
  const address = first === last ? first : `${first}:${last}`
  return sheetName ? `'${sheetName.replaceAll("'", "''")}'!${address}` : address
}

function columnLabel(index: number): string {
  let number = index + 1
  let label = ''
  while (number > 0) {
    number -= 1
    label = String.fromCharCode(number % 26 + 65) + label
    number = Math.floor(number / 26)
  }
  return label
}
