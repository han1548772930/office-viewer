import type { CellRange, SheetRangeThemeRule, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { SheetAxisEdit } from '../../domain/sheet-axis-structure.ts'
import { assertSheetRangeThemeName, normalizeSheetRangeThemeStyle } from './sheet-range-theme-style.ts'
import { sheetRangeThemeByName, sheetRangeThemeRangesIntersect, sheetUsedRangeThemeRule, type SheetRangeThemeContext, type SheetRangeThemeRegionSnapshot } from './sheet-range-theme-resolver.ts'
import { canonicalSheetAxisSpans } from '../../domain/sheet-axis.ts'
import { assertSheetRangeThemeStyleRange, planSheetRangeThemeAxisEdit } from './sheet-range-theme-structure.ts'

const object = (input: unknown): input is Record<string, unknown> => Boolean(input) && typeof input === 'object' && !Array.isArray(input)
export const SHEET_RANGE_THEME_MODEL_PLUGIN = 'SHEET_RANGE_THEME_MODEL_PLUGIN'
export function assertSheetRangeThemeRange(input: unknown): asserts input is CellRange {
  if (!object(input) || ![input.startRow, input.endRow, input.startCol, input.endCol].every(Number.isSafeInteger)
    || Number(input.startRow) < 0 || Number(input.endRow) < Number(input.startRow) || Number(input.endRow) >= 1_048_576
    || Number(input.startCol) < 0 || Number(input.endCol) < Number(input.startCol) || Number(input.endCol) >= 16_384) throw new RangeError('范围主题区域越界。')
}
function normalizeRule(input: unknown): SheetRangeThemeRule {
  if (!object(input) || typeof input.id !== 'string' || !input.id.trim()) throw new TypeError('范围主题规则身份无效。')
  assertSheetRangeThemeName(input.themeName); assertSheetRangeThemeRange(input.range)
  if (input.styleRange !== undefined) assertSheetRangeThemeStyleRange(input.styleRange, input.range)
  if (input.metadata !== undefined && !object(input.metadata)) throw new TypeError('范围主题规则扩展无效。')
  return { id: input.id, themeName: input.themeName, range: cloneModel(input.range),
    ...(input.styleRange !== undefined ? { styleRange: cloneModel(input.styleRange) } : {}),
    ...(input.metadata ? { metadata: cloneModel(input.metadata as Record<string, unknown>) } : {}) }
}
export function assertSheetRangeThemeRegionSnapshot(input: SheetRangeThemeRegionSnapshot, rowCount: number, columnCount: number): void {
  if (!Number.isSafeInteger(rowCount) || !Number.isSafeInteger(columnCount) || rowCount < 1 || columnCount < 1) throw new RangeError('区域范围主题快照尺寸无效。')
  if (!object(input) || !Array.isArray(input.rules) || !Array.isArray(input.hiddenRowSpans)) throw new TypeError('区域范围主题快照无效。')
  for (const [axis, length, maximum] of [[input.rows, rowCount, 1_048_576], [input.columns, columnCount, 16_384]] as const) {
    if (Array.isArray(axis)) {
      if (axis.length !== length || axis.some((value, index) => !Number.isSafeInteger(value) || value < 0 || value >= maximum || index > 0 && value <= axis[index - 1]!)) throw new RangeError('范围主题来源坐标无效。')
    } else if (!object(axis) || !Number.isSafeInteger(axis.start) || !Number.isSafeInteger(axis.end)
      || Number(axis.start) < 0 || Number(axis.end) >= maximum || Number(axis.end) - Number(axis.start) + 1 !== length) throw new RangeError('范围主题来源坐标无效。')
  }
  for (const item of input.rules) {
    if (!object(item)) throw new TypeError('范围主题采样规则无效。')
    assertSheetRangeThemeRange(item.range)
    if (item.styleRange !== undefined) assertSheetRangeThemeStyleRange(item.styleRange, item.range)
    if (item.theme !== null) normalizeSheetRangeThemeStyle(item.theme)
  }
  for (const span of input.hiddenRowSpans) if (!object(span) || typeof span.value !== 'boolean' || Number(span.end) >= 1_048_576) throw new RangeError('范围主题来源隐藏范围无效。')
  const canonical = canonicalSheetAxisSpans(input.hiddenRowSpans)
  if (canonical.length !== input.hiddenRowSpans.length || canonical.some((span, index) => {
    const source = input.hiddenRowSpans[index]!
    return span.start !== source.start || span.end !== source.end || span.value !== source.value
  })) throw new RangeError('范围主题来源隐藏范围必须有序且互不重叠。')
}
/** Initial model/resource data is detached and validated before the workbook is published. */
export function normalizeWorkbookRangeThemes(workbook: WorkbookModel): void {
  let themes = workbook.rangeThemes
  if (workbook.rangeThemes !== undefined) {
    if (!object(workbook.rangeThemes)) throw new TypeError('范围主题目录必须是对象。')
    themes = Object.fromEntries(Object.entries(workbook.rangeThemes).map(([name, input]) => {
      const theme = normalizeSheetRangeThemeStyle(input)
      if (theme.name !== name) throw new TypeError('范围主题目录名称不一致。')
      return [name, theme]
    }))
  }
  const candidates = new Map<SpreadsheetSheet, SheetRangeThemeRule[]>()
  for (const sheet of workbook.sheets) if (sheet.rangeThemeRules !== undefined) {
    if (!Array.isArray(sheet.rangeThemeRules)) throw new TypeError('范围主题规则必须是数组。')
    const rules = sheet.rangeThemeRules.map(normalizeRule)
    if (new Set(rules.map(rule => rule.id)).size !== rules.length) throw new TypeError('范围主题规则身份重复。')
    candidates.set(sheet, rules)
  }
  if (themes !== undefined) workbook.rangeThemes = themes
  for (const [sheet, rules] of candidates) sheet.rangeThemeRules = rules
}
export function importSheetRangeThemeResources(workbook: WorkbookModel, resources: unknown): void {
  if (resources === undefined) return
  if (!Array.isArray(resources)) throw new TypeError('工作簿资源必须是数组。')
  const resource = resources.find(item => object(item) && item.name === SHEET_RANGE_THEME_MODEL_PLUGIN)
  if (!resource) return
  const raw: unknown = typeof resource.data === 'string' ? JSON.parse(resource.data) : resource.data
  if (!object(raw)) throw new TypeError('公开范围主题资源无效。')
  const styles = raw.rangeThemeStyleMapJson ?? {}, rules = raw.rangeThemeStyleRuleMap ?? {}
  if (!object(styles) || !object(rules)) throw new TypeError('公开范围主题资源目录无效。')
  const themes = Object.fromEntries(Object.entries(styles).map(([name, style]) => {
    const theme = normalizeSheetRangeThemeStyle(style)
    if (theme.name !== name) throw new TypeError('范围主题目录名称不一致。')
    return [name, theme]
  }))
  const candidates = new Map<string, SheetRangeThemeRule[]>()
  for (const [id, value] of Object.entries(rules)) {
    if (!object(value) || !object(value.rangeInfo) || !object(value.rangeInfo.range) || typeof value.rangeInfo.subUnitId !== 'string') throw new TypeError('公开范围主题规则无效。')
    const sheetId = value.rangeInfo.subUnitId
    if (!workbook.sheets.some(sheet => sheet.id === sheetId)) throw new Error('公开范围主题资源引用了不存在的工作表。')
    const range = value.rangeInfo.range
    if (value.officeRuleMetadata !== undefined && !object(value.officeRuleMetadata)) throw new TypeError('公开范围主题扩展无效。')
    const rule = normalizeRule({ id, themeName: value.themeName, range: { startRow: range.startRow, endRow: range.endRow,
      startCol: range.startColumn, endCol: range.endColumn }, ...(value.officeStyleRange !== undefined ? { styleRange: value.officeStyleRange } : {}),
      metadata: { ...cloneModel(value.officeRuleMetadata ?? {}), sourceRule: cloneModel(value) } })
    candidates.set(sheetId, [...(candidates.get(sheetId) ?? []), rule])
  }
  const candidate = { ...workbook, rangeThemes: { ...themes, ...workbook.rangeThemes }, sheets: workbook.sheets.map(sheet => {
    const liveIds = new Set((sheet.rangeThemeRules ?? []).map(rule => rule.id))
    return { ...sheet, ...(candidates.has(sheet.id) ? { rangeThemeRules: [
      ...candidates.get(sheet.id)!.filter(rule => !liveIds.has(rule.id)), ...(sheet.rangeThemeRules ?? []),
    ] } : {}) }
  }) }
  normalizeWorkbookRangeThemes(candidate)
  workbook.rangeThemes = candidate.rangeThemes
  for (const sheet of workbook.sheets) {
    const rules = candidate.sheets.find(item => item.id === sheet.id)?.rangeThemeRules
    if (rules !== undefined) sheet.rangeThemeRules = rules
  }
}
/** Export the current directory and rules once; imported resource copies never override live edits. */
export function exportSheetRangeThemeResource(workbook: WorkbookModel, prior?: { name: string; data: string }): { name: string; data: string } | undefined {
  if (!prior && workbook.rangeThemes === undefined && !workbook.sheets.some(sheet => sheet.rangeThemeRules?.length)) return undefined
  const saved: unknown = prior?.data ? JSON.parse(prior.data) : {}
  if (!object(saved)) throw new TypeError('公开范围主题资源无效。')
  const ruleMap: Record<string, Record<string, unknown>> = {}
  for (const sheet of workbook.sheets) for (const rule of sheet.rangeThemeRules ?? []) {
    if (Object.hasOwn(ruleMap, rule.id)) throw new TypeError('公开范围主题规则身份跨工作表重复。')
    const source = object(rule.metadata?.sourceRule) ? cloneModel(rule.metadata.sourceRule) : {}
    const rangeInfo = object(source.rangeInfo) ? source.rangeInfo : {}, sourceRange = object(rangeInfo.range) ? rangeInfo.range : {}
    const metadata = cloneModel(rule.metadata ?? {})
    delete metadata.sourceRule
    const output: Record<string, unknown> = { ...source, themeName: rule.themeName,
      rangeInfo: { ...rangeInfo, unitId: workbook.id, subUnitId: sheet.id, range: { ...sourceRange,
        startRow: rule.range.startRow, endRow: rule.range.endRow, startColumn: rule.range.startCol, endColumn: rule.range.endCol } },
      officeRuleMetadata: metadata }
    delete output.officeStyleRange
    if (rule.styleRange) output.officeStyleRange = cloneModel(rule.styleRange)
    ruleMap[rule.id] = output
  }
  return { name: SHEET_RANGE_THEME_MODEL_PLUGIN, data: JSON.stringify({ ...saved,
    rangeThemeStyleMapJson: cloneModel(workbook.rangeThemes ?? {}), rangeThemeStyleRuleMap: ruleMap }) }
}
export function setSheetRangeThemeRules(sheet: Pick<SpreadsheetSheet, 'rangeThemeRules'>, ranges: readonly CellRange[], themeName: string | undefined, context: SheetRangeThemeContext): void {
  if (themeName !== undefined) {
    assertSheetRangeThemeName(themeName)
    if (!sheetRangeThemeByName(themeName, context)) throw new Error(`找不到范围主题：${themeName}`)
  }
  const rules = [...(sheet.rangeThemeRules ?? [])]
  for (const range of ranges) {
    assertSheetRangeThemeRange(range)
    if (themeName === undefined) {
      const current = sheetUsedRangeThemeRule({ rangeThemeRules: rules }, range)
      if (current) rules.splice(rules.findIndex(rule => rule.id === current.id), 1)
    } else {
      const equal = (first: CellRange, second: CellRange) => first.startRow === second.startRow && first.endRow === second.endRow && first.startCol === second.startCol && first.endCol === second.endCol
      const current = rules.find(rule => equal(rule.range, range))
      if (current?.themeName === themeName && !current.styleRange) continue
      if (current) rules.splice(rules.indexOf(current), 1)
      rules.push({ id: current?.id ?? `range-theme-${crypto.randomUUID()}`, themeName, range: cloneModel(range), ...(current?.metadata ? { metadata: cloneModel(current.metadata) } : {}) })
    }
  }
  sheet.rangeThemeRules = rules
}
export function removeSheetRangeThemeRules(sheet: Pick<SpreadsheetSheet, 'rangeThemeRules'>, ranges: readonly CellRange[], themeName: string): void {
  assertSheetRangeThemeName(themeName)
  const rules = [...(sheet.rangeThemeRules ?? [])]
  for (const range of ranges) {
    assertSheetRangeThemeRange(range)
    const current = rules.findLastIndex(rule => rule.themeName === themeName && sheetRangeThemeRangesIntersect(rule.range, range))
    if (current >= 0) rules.splice(current, 1)
  }
  sheet.rangeThemeRules = rules
}
export function transformSheetRangeThemeRules(sheet: SpreadsheetSheet, edit: SheetAxisEdit): void {
  if (!sheet.rangeThemeRules) return
  sheet.rangeThemeRules = planSheetRangeThemeAxisEdit(sheet.rangeThemeRules, edit)
}
