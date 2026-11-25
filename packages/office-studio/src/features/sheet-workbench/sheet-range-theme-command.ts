import type { CellRange, SheetRangeThemeRule, SheetRangeThemeStyle } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { sheetRangeCommandInputRanges, type SheetRangeCommandInput } from './sheet-range-command.ts'
import { removeSheetRangeThemeRules, setSheetRangeThemeRules } from './sheet-range-theme-model.ts'
import type { SheetRangeThemeContext, SheetRangeThemeSheet } from './sheet-range-theme-resolver.ts'
import { assertSheetRangeThemeName, normalizeSheetRangeThemeStyle, RangeThemeStyle } from './sheet-range-theme-style.ts'
import { SHEET_BUILTIN_RANGE_THEME_NAMES } from './sheet-range-theme-presets.ts'

export const OPEN_SHEET_RANGE_THEME_OPERATION = 'sheet.operation.open-range-theme'
export const SHEET_RANGE_THEME_COMMANDS = Object.freeze({
  register: 'sheet.command.register-worksheet-range-theme-style', unregister: 'sheet.command.unregister-worksheet-range-theme-style',
  set: 'sheet.command.set-worksheet-range-theme-style', remove: 'sheet.command.remove-worksheet-range-theme-style',
})
export interface SheetRangeThemeCommandInput extends SheetRangeCommandInput {
  themeName?: string | null
  rangeThemeStyle?: RangeThemeStyle
  rangeThemeStyleJson?: unknown
  expectedRevision?: number
  expectedHistoryGeneration?: number
  /** Captured UI host. Facade calls may explicitly edit a background worksheet. */
  host?: { unitId: string; sheetId: string }
}
export interface SheetRangeThemeCommandPlan {
  rangeThemeRules: SheetRangeThemeRule[]
  affected: CellRange[]
  changed: boolean
}
export function sheetRangeThemeCommandInput(value: unknown): SheetRangeThemeCommandInput {
  if (value === undefined) return {}
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('范围主题命令参数必须是对象。')
  const input = value as SheetRangeThemeCommandInput, ranges = sheetRangeCommandInputRanges(input)
  if (input.themeName !== undefined && input.themeName !== null) assertSheetRangeThemeName(input.themeName)
  if (input.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0)) throw new TypeError('工作簿版本无效。')
  if (input.expectedHistoryGeneration !== undefined && (!Number.isSafeInteger(input.expectedHistoryGeneration) || input.expectedHistoryGeneration <= 0)) throw new TypeError('工作簿历史版本无效。')
  if (input.host !== undefined && (!input.host || typeof input.host !== 'object' || Array.isArray(input.host)
    || typeof input.host.unitId !== 'string' || typeof input.host.sheetId !== 'string')) throw new TypeError('范围主题宿主无效。')
  return { ...input, ...(ranges === undefined ? {} : { ranges }) }
}

export type SheetRangeThemeCommandKind = keyof typeof SHEET_RANGE_THEME_COMMANDS
export interface SheetRangeThemeWorkbook extends SheetRangeThemeContext {
  sheets: readonly (SheetRangeThemeSheet & { id: string })[]
}
export interface SheetRangeThemeWorkbookCommandPlan {
  sheetId: string
  rangeThemes?: Record<string, SheetRangeThemeStyle>
  rangeThemeRules?: SheetRangeThemeRule[]
  affected: Record<string, CellRange[]>
  catalogChanged: boolean
  rulesChanged: boolean
  changed: boolean
}

/** Registering an edited definition and applying it are one content transaction. */
export function planWorkbookSheetRangeThemeCommand(workbook: SheetRangeThemeWorkbook, sheetId: string,
  ranges: readonly CellRange[], kind: SheetRangeThemeCommandKind, input: SheetRangeThemeCommandInput): SheetRangeThemeWorkbookCommandPlan {
  const sheet = workbook.sheets.find(item => item.id === sheetId)
  if (!sheet) throw new Error('找不到目标工作表。')
  const hasDefinition = input.rangeThemeStyle !== undefined || input.rangeThemeStyleJson !== undefined
  if (hasDefinition && (kind === 'remove' || kind === 'unregister')) throw new TypeError('移除范围主题不能同时保存主题配置。')
  if (input.rangeThemeStyle !== undefined && !(input.rangeThemeStyle instanceof RangeThemeStyle)) throw new TypeError('范围主题构建器无效。')
  const definition = kind === 'register' || hasDefinition ? normalizeSheetRangeThemeStyle(input.rangeThemeStyle instanceof RangeThemeStyle
    ? input.rangeThemeStyle.toJson() : input.rangeThemeStyleJson) : undefined
  const name = definition?.name ?? input.themeName
  if (definition && input.themeName !== undefined && input.themeName !== definition.name) throw new TypeError('应用名称与范围主题配置不一致。')
  if (definition && kind === 'set' && SHEET_BUILTIN_RANGE_THEME_NAMES.includes(definition.name)) throw new TypeError('内置主题名称已保留，请使用自定义名称。')
  if (kind === 'unregister') assertSheetRangeThemeName(name)
  const previousCatalog = workbook.rangeThemes ?? sheet.rangeThemes
  let catalog = previousCatalog === undefined ? undefined : cloneModel({ ...previousCatalog })
  const canChangeCatalog = kind === 'register' || kind === 'unregister' || kind === 'set' && ranges.length > 0
  if (canChangeCatalog && definition) {
    const entries = Object.entries(catalog ?? {})
    catalog = Object.fromEntries(Object.hasOwn(catalog ?? {}, definition.name)
      ? entries.map(([key, value]) => [key, key === definition.name ? cloneModel(definition) : value])
      : [...entries, [definition.name, cloneModel(definition)]])
  }
  else if (kind === 'unregister' && catalog && Object.hasOwn(catalog, name!)) catalog = Object.fromEntries(Object.entries(catalog).filter(([key]) => key !== name))
  const catalogChanged = JSON.stringify(catalog) !== JSON.stringify(previousCatalog), affected: Record<string, CellRange[]> = Object.create(null)
  if (catalogChanged) for (const item of workbook.sheets) {
    const references = (item.rangeThemeRules ?? []).filter(rule => rule.themeName === name).map(rule => cloneModel(rule.range))
    if (references.length) affected[item.id] = references
  }
  const rules = (kind === 'set' || kind === 'remove') && ranges.length ? planSheetRangeThemeCommand(sheet, ranges,
    kind, name, { ...workbook, rangeThemes: catalog }) : undefined
  if (rules) affected[sheetId] = [...(affected[sheetId] ?? []), ...rules.affected]
  return { sheetId, rangeThemes: catalog, rangeThemeRules: rules?.rangeThemeRules, affected,
    catalogChanged, rulesChanged: rules?.changed ?? false, changed: catalogChanged || (rules?.changed ?? false) }
}

/** Domain worksheets omit view catalogs; detached views update every existing catalog projection. */
export function applySheetRangeThemeWorkbookPlan(workbook: { rangeThemes?: Record<string, SheetRangeThemeStyle>;
  sheets: Array<{ id: string; rangeThemeRules?: SheetRangeThemeRule[]; rangeThemes?: Record<string, SheetRangeThemeStyle> }> },
  plan: SheetRangeThemeWorkbookCommandPlan): void {
  if (plan.catalogChanged) {
    workbook.rangeThemes = plan.rangeThemes === undefined ? undefined : cloneModel(plan.rangeThemes)
    for (const sheet of workbook.sheets) if (Object.hasOwn(sheet, 'rangeThemes')) sheet.rangeThemes = plan.rangeThemes === undefined ? undefined : cloneModel(plan.rangeThemes)
  }
  if (plan.rulesChanged) workbook.sheets.find(sheet => sheet.id === plan.sheetId)!.rangeThemeRules = cloneModel(plan.rangeThemeRules!)
}
/** API and detached workbenches prepare the same rule changes and whole removed-rule permission scope. */
export function planSheetRangeThemeCommand(sheet: SheetRangeThemeSheet, ranges: readonly CellRange[],
  action: 'set' | 'remove', themeName: string | null | undefined, context: SheetRangeThemeContext = {}): SheetRangeThemeCommandPlan {
  const previous = sheet.rangeThemeRules ?? [], candidate = { rangeThemeRules: cloneModel([...previous]) }
  if (action === 'set') setSheetRangeThemeRules(candidate, ranges, themeName ?? undefined, { ...context, rangeThemes: context.rangeThemes ?? sheet.rangeThemes })
  else { assertSheetRangeThemeName(themeName); removeSheetRangeThemeRules(candidate, ranges, themeName) }
  const remaining = new Set(candidate.rangeThemeRules.map(rule => rule.id))
  return { rangeThemeRules: candidate.rangeThemeRules,
    affected: [...ranges.map(range => ({ ...range })), ...previous.filter(rule => !remaining.has(rule.id)).map(rule => ({ ...rule.range }))],
    changed: JSON.stringify(candidate.rangeThemeRules) !== JSON.stringify(previous) }
}
