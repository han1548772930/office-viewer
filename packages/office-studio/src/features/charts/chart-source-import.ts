import { randomId } from '../../api/model-utils'
import type { CellRange, ChartModel, PivotTableModel, SheetStyleSource, SheetTableModel, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model'
import { cloneModel } from '../../domain/serialization'
import { mapFormulaReferenceAreas, mapFormulaSheetQualifiers, type FormulaSheetQualifier } from '../formula/formula-reference-areas'
import { mapFormulaIdentifiers, visitFormulaIdentifiers, type FormulaIdentifier } from '../formula/formula-identifiers'
import { readFormulaDefinedNames, type FormulaDefinedName as SourceName, type FormulaDefinedNameKey as NameKey } from '../formula/formula-defined-names'
import { composeFormulaImportAliases, formulaCellUsesCalendar, readFormulaCalendarValue, readFormulaImportAliases, readFormulaImportSource } from '../formula/formula-import-context'
import { canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection'
import { assertSheetCommandRanges, sheetRangeContainsCell, sheetRangesIntersect } from '../sheet-workbench/sheet-range-command'
import { composeSheetCellStyles, resolveSheetCellStyle, resolveSheetStyleSource } from '../sheet-workbench/sheet-style-inheritance'
import { createWorksheetCopyName } from '../sheet-workbench/sheet-worksheet-copy'
import { copyPivotReportCache } from '../slicer/slicer-connections-lifecycle'
import { prepareCopiedPivotSourceReport } from '../pivot/services/pivot-source-lifecycle'
import { CHART_IMAGE_RESOURCE_KEY, readChartImageResources, normalizeChartImageResource, resolveChartImageResource } from './chart-image-resources'
import { assertChartReadPermission } from './services/chart-statistics-service'
import type { ChartObjectClipboard } from './chart-object-clipboard'
import { assertChartSourceImportState, chartSourceImportCatalog, type ChartSourceImportState } from './chart-source-import-state'
import { consumeChartPreparation, consumeChartPreparationSync, type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { chartSourceDateDelta, pivotFieldUsesDates, remapChartSourceAxisDates, remapChartSourceDateCell,
  remapChartSourceDateFilter, remapChartSourcePivotDates } from './chart-source-import-dates'

export const CHART_SOURCE_IMPORT_MAX_BYTES = 64 * 1024 * 1024
export const CHART_SOURCE_IMPORT_MAX_CELLS = 250_000
const IMPORT_CHUNK = 256, MAX_RANGES = 8192, MAX_NAMES = 8192
const fullRange: CellRange = { startRow: 0, startCol: 0, endRow: 1_048_575, endCol: 16_383 }
const contains = (outer: CellRange, inner: CellRange): boolean => sheetRangeContainsCell(outer, inner.startRow, inner.startCol)
  && sheetRangeContainsCell(outer, inner.endRow, inner.endCol)
const quoteSheet = (name: string): string => `'${name.replace(/'/g, "''")}'!`
export interface ChartSourceImportPlan {
  readonly version: 1
  readonly sourceUnitId: string
  readonly targetUnitId: string
  readonly sourceChartId: string
  readonly sourceSnapshot: string
  readonly sourceState: ChartSourceImportState
  readonly targetSheetId: string
  readonly chart: ChartModel
  readonly sheets: SpreadsheetSheet[]
  readonly tables: SheetTableModel[]
  readonly pivotTables: PivotTableModel[]
  readonly pivotCaches: NonNullable<WorkbookModel['pivotCaches']>
  readonly styles: NonNullable<WorkbookModel['styles']>
  readonly rangeThemes: NonNullable<WorkbookModel['rangeThemes']>
  readonly names: Partial<Record<NameKey, Record<string, unknown>[]>>
  readonly images: ReturnType<typeof readChartImageResources>
  readonly sourceSheetIds: readonly string[]
  readonly sourceRanges: ReadonlyArray<{ sheetId: string; range: CellRange }>
  readonly diagnostics: readonly string[]
  readonly estimatedBytes: number
  readonly copiedCells: number
}
export interface ChartSourceImportOptions extends ChartComputePreparationContext { userId: string; targetSheetId: string }

function sourceNames(workbook: WorkbookModel): SourceName[] {
  return readFormulaDefinedNames(workbook, MAX_NAMES)
}
function uniqueName(name: string, used: Set<string>): string {
  let candidate = name, suffix = 2
  while (used.has(candidate.toLowerCase())) candidate = `${name.slice(0, 240)}_${suffix++}`
  used.add(candidate.toLowerCase()); return candidate
}
function qualifierSheets(workbook: WorkbookModel, qualifier: Readonly<FormulaSheetQualifier>): SpreadsheetSheet[] {
  if (qualifier.external) return []
  if (!qualifier.threeDimensional) {
    const sheet = workbook.sheets.find(sheet => sheet.name.toLowerCase() === qualifier.sheetName.toLowerCase())
    return sheet ? [sheet] : []
  }
  const [first, last] = qualifier.sheetName.split(':'), start = workbook.sheets.findIndex(sheet => sheet.name.toLowerCase() === first?.toLowerCase()),
    end = workbook.sheets.findIndex(sheet => sheet.name.toLowerCase() === last?.toLowerCase())
  return start < 0 || end < 0 ? [] : workbook.sheets.slice(Math.min(start, end), Math.max(start, end) + 1)
}

/** The sole planner is stepped for UI calls and consumed synchronously only by
 * small-source/model contracts. It imports a sparse dependency closure, never
 * constructs a dense worksheet or publishes part of the candidate. */
function* importSteps(source: WorkbookModel, target: WorkbookModel, clipboard: ChartObjectClipboard,
  options: ChartSourceImportOptions): Generator<void, ChartSourceImportPlan, void> {
  if (typeof options.userId !== 'string' || !options.userId.trim()) throw new TypeError('图表依赖导入需要明确的用户权限上下文。')
  if (!source.id || !target.id || source.id === target.id || clipboard.source.unitId !== source.id) throw new TypeError('图表依赖导入的源与目标工作簿无效。')
  if (!target.sheets.some(sheet => sheet.id === options.targetSheetId)) throw new Error('粘贴图表的目标工作表不存在。')
  const sourceChart = source.charts?.find(chart => chart.id === clipboard.source.chartId && chart.sheetId === clipboard.source.sheetId)
  if (!sourceChart || JSON.stringify(sourceChart) !== clipboard.source.snapshot || JSON.stringify(clipboard.chart) !== clipboard.source.snapshot) throw new Error('复制源图表已更改或删除，请重新复制。')
  assertChartReadPermission(source, sourceChart, options.userId)
  const sourceImages: Record<string, string> = Object.create(null)
  for (const [id, image] of Object.entries(clipboard.images)) {
    const current = resolveChartImageResource(source, sourceChart, id)
    if (!current || JSON.stringify(current) !== JSON.stringify(image)) throw new Error('复制源图片资源已更改，请重新复制图表。')
    sourceImages[id] = JSON.stringify(current)
  }
  let estimatedBytes = 0, copiedCells = 0, work = 0
  const budget = Math.min(options.memoryBudgetBytes ?? CHART_SOURCE_IMPORT_MAX_BYTES, CHART_SOURCE_IMPORT_MAX_BYTES)
  const charge = (value: unknown, overhead = 0) => {
    const text = JSON.stringify(value)
    estimatedBytes += (text?.length ?? 0) * 2 + overhead
    if (!Number.isSafeInteger(estimatedBytes) || estimatedBytes > budget) throw new RangeError('图表依赖导入超出 64 MiB 内存预算。')
  }
  const sourceState: ChartSourceImportState = { catalog: chartSourceImportCatalog(source), cells: Object.create(null), images: sourceImages }
  charge(sourceState.catalog, 256)
  charge(sourceImages, 128)
  const queue: Array<{ sheetId: string; range: CellRange }> = [], ranges = new Map<string, CellRange[]>(), cells = new Map<string, Record<string, SpreadsheetCell>>()
  const formulas: Array<{ formula: string; owner: string }> = []
  const selectedTables = new Map<string, SheetTableModel>(), selectedPivots = new Map<string, PivotTableModel>(), selectedNames = new Map<string, SourceName>()
  const diagnostics = new Set<string>(), names = sourceNames(source), dateRanges = new Map<string, CellRange[]>()
  const sheetById = new Map(source.sheets.map(sheet => [sheet.id, sheet]))
  const namesByIdentity = new Map<string, SourceName>()
  for (const name of names) {
    const previous = namesByIdentity.get(name.identity)
    if (previous && previous.expression !== name.expression) throw new TypeError('图表来源的同一名称作用域存在冲突。')
    if (!previous) namesByIdentity.set(name.identity, name)
  }
  const addScopeSheet = (sheetId: string): SpreadsheetSheet => {
    const sheet = sheetById.get(sheetId)
    if (!sheet) throw new Error('图表的公式、名称或透视来源工作表已经丢失。')
    if (!canViewWorksheetProtection(sheet, options.userId)) throw new Error('没有查看图表定义名称所属工作表的权限。')
    if (!cells.has(sheetId)) {
      cells.set(sheetId, Object.create(null)); ranges.set(sheetId, [])
      sourceState.cells[sheetId] = Object.create(null)
      charge({ id: sheetId }, 256)
    }
    return sheet
  }
  const addFormula = (formula: string, owner: string) => {
    if (formula.length > 32_768) throw new RangeError('图表依赖的公式表达式超出预算。')
    charge(undefined, 64); formulas.push({ formula, owner })
  }
  const addRange = (sheetId: string, requested: CellRange) => {
    assertSheetCommandRanges([requested])
    const sheet = sheetById.get(sheetId)
    if (!sheet) throw new Error('图表的公式、名称或透视来源工作表已经丢失。')
    const range = { ...requested, endRow: Math.min(requested.endRow, (sheet.rowCount ?? 1_048_576) - 1),
      endCol: Math.min(requested.endCol, (sheet.columnCount ?? 16_384) - 1) }
    if (range.startRow > range.endRow || range.startCol > range.endCol) return
    const previous = ranges.get(sheetId) ?? []
    if (previous.some(item => contains(item, range))) return
    if (!canViewWorksheetProtection(sheet, options.userId) || !canViewProtectedRanges(sheet, [range], options.userId)) throw new Error('没有查看图表全部数据依赖的权限。')
    if (queue.length >= MAX_RANGES) throw new RangeError('图表数据依赖的范围数量超出预算。')
    addScopeSheet(sheetId)
    previous.push(range); ranges.set(sheetId, previous); queue.push({ sheetId, range }); charge(range, 128)
  }
  const addDateRange = (sheetId: string, range: CellRange) => {
    const current = dateRanges.get(sheetId) ?? []; current.push({ ...range }); dateRanges.set(sheetId, current)
  }
  const addPivot = (pivot: PivotTableModel) => {
    if (selectedPivots.has(pivot.id)) return
    if (pivot.source.unitId !== source.id || pivot.anchor.unitId !== source.id) throw new Error('透视图的来源或输出属于未解析的外部工作簿。')
    selectedPivots.set(pivot.id, pivot); addRange(pivot.source.subUnitId, pivot.source.range)
    addRange(pivot.anchor.subUnitId, pivot.outputRange ?? { startRow: pivot.anchor.row, endRow: pivot.anchor.row, startCol: pivot.anchor.col, endCol: pivot.anchor.col })
    for (const field of pivot.fields) if (pivotFieldUsesDates(field)) addDateRange(pivot.source.subUnitId,
      { ...pivot.source.range, startRow: pivot.source.range.startRow + 1, startCol: field.sourceColumn, endCol: field.sourceColumn })
  }
  const resolveName = (identifier: Readonly<FormulaIdentifier>, owner: string): SourceName | undefined => {
    const scoped = identifier.sheetName ? qualifierSheets(source, identifier)[0]?.id : owner
    if (identifier.sheetName && !scoped) return undefined
    return (scoped ? namesByIdentity.get(`${scoped}:${identifier.name.toLowerCase()}`) : undefined)
      ?? namesByIdentity.get(`*:${identifier.name.toLowerCase()}`)
  }
  const selectName = (name: SourceName): void => {
    if (selectedNames.has(name.identity)) return
    if (name.expression === undefined) throw new TypeError('图表依赖的定义名称缺少可保留的引用表达式。')
    selectedNames.set(name.identity, name)
    addScopeSheet(name.defaultOwner)
    charge(name.record, 128); addFormula(name.expression, name.defaultOwner)
  }
  const inspectFormula = (formula: string, owner: string): void => {
    mapFormulaReferenceAreas(formula, area => {
      if (area.external) { diagnostics.add('external-formula-reference'); return { ...area } }
      const hosts = area.sheetName ? qualifierSheets(source, { ...area, sheetName: area.sheetName }) : [sheetById.get(owner)!]
      if (!hosts.length) { diagnostics.add('unresolved-formula-sheet'); return { ...area } }
      if (area.endSheetName && area.endSheetName.toLowerCase() !== area.sheetName?.toLowerCase()) {
        for (const sheet of qualifierSheets(source, { sheetName: area.endSheetName, external: false, threeDimensional: false })) hosts.push(sheet)
      }
      for (const sheet of hosts) addRange(sheet.id, { startRow: Math.min(area.startRow, area.endRow), endRow: Math.max(area.startRow, area.endRow),
        startCol: Math.min(area.startColumn, area.endColumn), endCol: Math.max(area.startColumn, area.endColumn) })
      return { ...area }
    })
    visitFormulaIdentifiers(formula, identifier => {
      if (identifier.external) { diagnostics.add('external-formula-name'); return }
      if (identifier.kind === 'function') {
        if (/^(?:_xlfn\.)?(?:INDIRECT|OFFSET)$/i.test(identifier.name)) {
          diagnostics.add('dynamic-formula-reference')
          // Dynamic references retain every possible local cell. Exact string
          // evaluation remains the existing formula engine's responsibility.
          for (const sheet of source.sheets) addRange(sheet.id, fullRange)
          if (/INDIRECT$/i.test(identifier.name)) for (const name of names) selectName(name)
        }
        return
      }
      if (identifier.kind === 'table') {
        const table = source.tables?.find(table => table.name.toLowerCase() === identifier.name.toLowerCase())
        if (table) { selectedTables.set(table.id, table); addRange(table.sheetId, table.range) }
        else diagnostics.add('unresolved-structured-table')
        return
      }
      if (/^(?:TRUE|FALSE)$/i.test(identifier.name)) return
      if (identifier.sheetName) {
        const scoped = qualifierSheets(source, identifier)
        if (!scoped.length) diagnostics.add('unresolved-formula-sheet')
        for (const sheet of scoped) addScopeSheet(sheet.id)
      }
      const name = resolveName(identifier, owner)
      if (!name) {
        const table = source.tables?.find(table => table.name.toLowerCase() === identifier.name.toLowerCase())
        if (table) { selectedTables.set(table.id, table); addRange(table.sheetId, table.range) }
        else diagnostics.add('unresolved-formula-name')
        return
      }
      selectName(name)
    })
  }
  if (clipboard.chart.pivotTableId) {
    const pivot = source.pivotTables?.find(pivot => pivot.id === clipboard.chart.pivotTableId)
    if (!pivot) throw new Error('复制图表的透视报告已经丢失。')
    addPivot(pivot)
  } else if (clipboard.chart.ownedPivot) {
    const pivot = clipboard.chart.ownedPivot
    if (pivot.source.unitId !== source.id) throw new Error('独立透视图表的来源不属于源工作簿。')
    addRange(pivot.source.subUnitId, pivot.source.range)
    for (const field of pivot.fields) if (pivotFieldUsesDates(field)) addDateRange(pivot.source.subUnitId,
      { ...pivot.source.range, startRow: pivot.source.range.startRow + 1, startCol: field.sourceColumn, endCol: field.sourceColumn })
  } else {
    const sheetId = clipboard.chart.source.sheetId ?? clipboard.chart.sheetId, range = clipboard.chart.source.range
    addRange(sheetId, range)
    if (clipboard.chart.options?.mapping?.useDateAxis) {
      const category = clipboard.chart.categoryField ?? 0, columns = clipboard.chart.source.orientation === 'columns'
      addDateRange(sheetId, columns ? { ...range, startRow: range.startRow + 1, startCol: range.startCol + category, endCol: range.startCol + category }
        : { ...range, startCol: range.startCol + 1, startRow: range.startRow + category, endRow: range.startRow + category })
    }
  }
  let rangeIndex = 0, formulaIndex = 0
  while (rangeIndex < queue.length || formulaIndex < formulas.length) {
    // Names are queued, including cycles. A long dependency chain never
    // consumes the JavaScript call stack or bypasses cooperative cancellation.
    if (formulaIndex < formulas.length) {
      const entry = formulas[formulaIndex++]!
      inspectFormula(entry.formula, entry.owner)
      if (++work % IMPORT_CHUNK === 0) yield
      continue
    }
    const { sheetId, range } = queue[rangeIndex++]!, sheet = sheetById.get(sheetId)!, selected = cells.get(sheetId)!
    for (const table of source.tables ?? []) if (table.sheetId === sheetId && sheetRangesIntersect(range, table.range)) {
      selectedTables.set(table.id, table); addRange(sheetId, table.range)
    }
    for (const merged of sheet.mergedRanges) if (sheetRangesIntersect(range, merged)) addRange(sheetId, merged)
    if (sheet.autoFilter && sheetRangesIntersect(range, sheet.autoFilter.range)) addRange(sheetId, sheet.autoFilter.range)
    for (const key in sheet.cells) {
      if (++work % IMPORT_CHUNK === 0) yield
      if (!Object.hasOwn(sheet.cells, key) || Object.hasOwn(selected, key)) continue
      const cell = sheet.cells[key]!
      if (!sheetRangeContainsCell(range, cell.row, cell.col)) continue
      if (++copiedCells > CHART_SOURCE_IMPORT_MAX_CELLS) throw new RangeError('图表源依赖单元格数量超出 250000 格预算。')
      charge(cell, 256)
      const snapshot = JSON.stringify(cell)
      charge(snapshot, 64); sourceState.cells[sheetId]![key] = snapshot
      selected[key] = cloneModel(cell)
      if (cell.formula) {
        const sourceContext = readFormulaImportSource(cell)
        if (sourceContext?.aliasSheetId) addScopeSheet(sourceContext.aliasSheetId)
        addFormula(cell.formula, sheetId)
      }
      const pivotId = cell.metadata?.pivotTableId
      if (typeof pivotId === 'string') {
        const pivot = source.pivotTables?.find(pivot => pivot.id === pivotId)
        if (!pivot) throw new Error('图表依赖单元格的透视报告已经丢失。')
        addPivot(pivot)
      }
    }
  }
  const sheetIds = new Map<string, string>(), sheetNames = new Map<string, string>(), mappedSheetList = [...target.sheets]
  for (const sheet of source.sheets) if (cells.has(sheet.id)) {
    const name = mappedSheetList.some(item => item.name.toLowerCase() === sheet.name.toLowerCase()) ? createWorksheetCopyName(mappedSheetList, sheet.name) : sheet.name
    const id = randomId('chart-data-sheet'); sheetIds.set(sheet.id, id); sheetNames.set(sheet.name.toLowerCase(), name)
    mappedSheetList.push({ ...sheet, id, name })
  }
  const targetNames = sourceNames(target)
  const namespace = new Set([...(target.tables ?? []).map(table => table.name.toLowerCase()), ...targetNames.map(name => name.name.toLowerCase()),
    ...[...selectedNames.values()].filter(name => name.owner).map(name => name.name.toLowerCase())])
  const reserveFormulaNames = (formula: string): void => visitFormulaIdentifiers(formula, identifier => {
    if (identifier.kind !== 'function') namespace.add(identifier.name.toLowerCase())
  })
  // Existing unresolved target names must remain unresolved after import.
  // Stored ref_text values also reserve their namespace for INDIRECT users.
  for (const sheet of target.sheets) for (const cell of Object.values(sheet.cells)) {
    if (cell.formula) reserveFormulaNames(cell.formula)
    if (typeof cell.value === 'string' && /^[\p{L}_\\][\p{L}\p{M}\p{N}_.\\]*$/u.test(cell.value)) namespace.add(cell.value.toLowerCase())
    if (++work % IMPORT_CHUNK === 0) yield
  }
  for (const name of targetNames) if (name.expression) reserveFormulaNames(name.expression)
  const tableNames = new Map<string, string>(), globalNames = new Map<string, string>(), unresolvedNames = new Map<string, string>()
  const tables = [...selectedTables.values()].map(table => {
    const name = uniqueName(table.name, namespace), id = randomId('chart-data-table')
    tableNames.set(table.name.toLowerCase(), name)
    return { ...cloneModel(table), id, name, sheetId: sheetIds.get(table.sheetId)!, columns: table.columns.map(column => ({ ...column, id: randomId('table-column') })) }
  })
  for (const name of selectedNames.values()) if (!name.owner) globalNames.set(name.name.toLowerCase(), uniqueName(name.name, namespace))
  const remapFormula = (formula: string, owner: string, qualify = false): string => {
    let next = mapFormulaIdentifiers(formula, identifier => {
      if (identifier.external || identifier.kind === 'function') return undefined
      if (identifier.kind === 'table') return tableNames.get(identifier.name.toLowerCase())
      const local = resolveName(identifier, owner)
      if (local?.owner && qualify && !identifier.sheetName) return { name: local.name, sheetName: sheetById.get(local.owner)!.name }
      if (local) return local.owner ? undefined : { name: globalNames.get(local.name.toLowerCase())!, sheetName: sheetById.get(local.defaultOwner)!.name }
      const table = tableNames.get(identifier.name.toLowerCase())
      if (table) return table
      if (/^(?:TRUE|FALSE)$/i.test(identifier.name) || identifier.sheetName || !namespace.has(identifier.name.toLowerCase())) return undefined
      // A source #NAME? must not silently bind to an unrelated destination name.
      const key = `${owner}:${identifier.name.toLowerCase()}`
      if (!unresolvedNames.has(key)) unresolvedNames.set(key, uniqueName(`${identifier.name}_unresolved`, namespace))
      return unresolvedNames.get(key)
    })
    next = mapFormulaSheetQualifiers(next, qualifier => {
      if (qualifier.external) return qualifier.sheetName
      if (qualifier.threeDimensional) {
        const mapped = qualifier.sheetName.split(':').map(name => sheetNames.get(name.toLowerCase()))
        return mapped.every(name => name !== undefined) ? mapped.join(':') : null
      }
      return sheetNames.get(qualifier.sheetName.toLowerCase()) ?? null
    })
    if (qualify) next = mapFormulaReferenceAreas(next, area => !area.qualifier ? { ...area, qualifier: quoteSheet(sheetNames.get(sheetById.get(owner)!.name.toLowerCase())!),
      sheetName: sheetNames.get(sheetById.get(owner)!.name.toLowerCase())! } : { ...area })
    return next
  }
  const neededStyles = new Set<string>(), neededThemes = new Set<string>()
  const needStyle = (value: SheetStyleSource | undefined): void => { if (typeof value === 'string') neededStyles.add(value) }
  needStyle(source.defaultStyle)
  for (const sheet of source.sheets) if (cells.has(sheet.id)) {
    needStyle(sheet.defaultStyle)
    for (const styles of [sheet.rowStyles, sheet.columnStyles]) for (const value of Object.values(styles ?? {})) needStyle(value)
    for (const spans of [sheet.rowStyleSpans, sheet.columnStyleSpans]) for (const span of spans ?? []) needStyle(span.value)
    for (const cell of Object.values(cells.get(sheet.id)!)) {
      needStyle(cell.authoredStyleReference); needStyle(cell.themeStyle)
      if (++work % IMPORT_CHUNK === 0) yield
    }
    for (const rule of sheet.rangeThemeRules ?? []) if (ranges.get(sheet.id)!.some(range => sheetRangesIntersect(range, rule.range))) neededThemes.add(rule.themeName)
  }
  const styles: NonNullable<WorkbookModel['styles']> = {}, styleIds = new Map<string, string>()
  for (const id of neededStyles) {
    const style = source.styles?.[id]
    if (!style) throw new Error(`图表数据依赖的样式已经丢失：${id}`)
    charge(style, 128)
    const mapped = target.styles?.[id] && JSON.stringify(target.styles[id]) !== JSON.stringify(style) ? randomId('chart-data-style') : id
    styles[mapped] = cloneModel(style); styleIds.set(id, mapped)
    if (++work % IMPORT_CHUNK === 0) yield
  }
  const rangeThemes: NonNullable<WorkbookModel['rangeThemes']> = {}, themeNames = new Map<string, string>(), usedThemes = new Set(Object.keys(target.rangeThemes ?? {}).map(name => name.toLowerCase()))
  for (const name of neededThemes) {
    const theme = source.rangeThemes?.[name]
    if (!theme) continue // Built-in themes are resolved by the shared catalog.
    charge(theme, 128)
    const mapped = uniqueName(name, usedThemes); themeNames.set(name, mapped); rangeThemes[mapped] = { ...cloneModel(theme), name: mapped }
    if (++work % IMPORT_CHUNK === 0) yield
  }
  const styleSource = (value: SheetStyleSource): SheetStyleSource => typeof value === 'string' ? styleIds.get(value) ?? value : value
  const importedNames: ChartSourceImportPlan['names'] = {}, localNames = new Map<string, Partial<Record<NameKey, Record<string, unknown>[]>>>()
  for (const name of selectedNames.values()) {
    const owner = sheetIds.get(name.owner ?? name.defaultOwner)
    if (!owner) throw new Error('定义名称的本地工作表未进入图表依赖闭环。')
    const mapped: Record<string, unknown> = { ...cloneModel(name.record), name: name.owner ? name.name : globalNames.get(name.name.toLowerCase())!,
      ...(name.record.id === undefined ? {} : { id: randomId('defined-name') }),
      refersTo: remapFormula(name.expression!, name.defaultOwner, true) }
    if (name.record.formulaOrRefString !== undefined) mapped.formulaOrRefString = mapped.refersTo
    mapped.sheetId = owner
    if (name.record.subUnitId !== undefined) mapped.subUnitId = owner
    if (name.record.sheet !== undefined) mapped.sheet = mappedSheetList.find(sheet => sheet.id === owner)!.name
    if (name.record.localSheetId !== undefined) mapped.localSheetId = typeof name.record.localSheetId === 'string' ? owner : mappedSheetList.findIndex(sheet => sheet.id === owner)
    const metadata = localNames.get(owner) ?? {}; (metadata[name.key] ??= []).push(mapped); localNames.set(owner, metadata)
    if (++work % IMPORT_CHUNK === 0) yield
  }
  const delta = chartSourceDateDelta(source.date1904, target.date1904), pivotIds = new Map([...selectedPivots.keys()].map(id => [id, randomId('chart-data-pivot')]))
  const sheets: SpreadsheetSheet[] = [], convertedDateColumns = new Map<string, Set<number>>()
  for (const sourceSheet of source.sheets) if (cells.has(sourceSheet.id)) {
    const sheetId = sheetIds.get(sourceSheet.id)!, sourceRanges = ranges.get(sourceSheet.id)!
    const projection = { ...sourceSheet, cells: cells.get(sourceSheet.id)!, images: undefined, shapes: undefined, drawingState: undefined, backgroundImage: undefined,
      dataValidations: undefined, conditionalFormats: undefined, rangeProtectionRules: undefined, worksheetProtection: undefined, metadata: undefined }
    // Clone the small axis/format descriptor independently of the already
    // detached sparse cells; no full source-sheet clone is allocated.
    const descriptor = { ...projection, cells: {} }
    charge(descriptor, 1024)
    const sheet: SpreadsheetSheet = { ...cloneModel(descriptor), cells: projection.cells, id: sheetId,
      name: sheetNames.get(sourceSheet.name.toLowerCase())!, visible: false }
    sheet.defaultStyle = composeSheetCellStyles(resolveSheetStyleSource(source.defaultStyle, source), resolveSheetStyleSource(sourceSheet.defaultStyle, source))
    if (sheet.rowStyles) sheet.rowStyles = Object.fromEntries(Object.entries(sheet.rowStyles).map(([row, style]) => [row, styleSource(style)]))
    if (sheet.columnStyles) sheet.columnStyles = Object.fromEntries(Object.entries(sheet.columnStyles).map(([column, style]) => [column, styleSource(style)]))
    if (sheet.rowStyleSpans) sheet.rowStyleSpans = sheet.rowStyleSpans.map(span => ({ ...span, value: styleSource(span.value) }))
    if (sheet.columnStyleSpans) sheet.columnStyleSpans = sheet.columnStyleSpans.map(span => ({ ...span, value: styleSource(span.value) }))
    sheet.mergedRanges = sheet.mergedRanges.filter(range => sourceRanges.some(source => contains(source, range)))
    sheet.rangeThemeRules = sheet.rangeThemeRules?.filter(rule => sourceRanges.some(range => sheetRangesIntersect(range, rule.range)))
      .map(rule => ({ ...rule, id: randomId('range-theme'), themeName: themeNames.get(rule.themeName) ?? rule.themeName }))
    if (sheet.outline) sheet.outline.groups = sheet.outline.groups.map(group => ({ ...group, id: randomId('outline') }))
    const provenance: Record<string, unknown> = { version: 1, sourceUnitId: source.id,
      sourceSheetId: sourceSheet.id, sourceSheetName: sourceSheet.name, sourceChartId: clipboard.source.chartId, sourceDate1904: source.date1904,
      ranges: cloneModel(sourceRanges), ...(sourceSheet.metadata ? { sourceMetadata: cloneModel(sourceSheet.metadata) } : {}) }
    sheet.metadata = { hiddenState: 'hidden', ...localNames.get(sheetId), chartDependencyImport: provenance }
    const aliasNames: Record<string, string> = Object.fromEntries(globalNames)
    const nameSheets: Record<string, string> = {}
    for (const name of selectedNames.values()) if (!name.owner) {
      const mapped = globalNames.get(name.name.toLowerCase())!, owner = sheetNames.get(sheetById.get(name.defaultOwner)!.name.toLowerCase())!
      nameSheets[name.name.toLowerCase()] = owner; nameSheets[mapped.toLowerCase()] = owner
    }
    for (const name of selectedNames.values()) if (name.owner) {
      const nameSheet = sheetById.get(name.owner)!
      const sourceKey = `${nameSheet.name.toLowerCase()}!${name.name.toLowerCase()}`
      const mappedOwner = sheetNames.get(nameSheet.name.toLowerCase())!
      aliasNames[sourceKey] = name.name; nameSheets[sourceKey] = mappedOwner
      const mappedKey = `${mappedOwner.toLowerCase()}!${name.name.toLowerCase()}`
      aliasNames[mappedKey] = name.name; nameSheets[mappedKey] = mappedOwner
      if (name.owner === sourceSheet.id) { aliasNames[name.name.toLowerCase()] = name.name; nameSheets[name.name.toLowerCase()] = mappedOwner }
    }
    const formulaAliases = composeFormulaImportAliases({ sheets: Object.fromEntries(sheetNames), names: aliasNames, tables: Object.fromEntries(tableNames), nameSheets }, readFormulaImportAliases(sourceSheet))
    provenance.formulaAliases = formulaAliases
    charge(sheet.metadata)
    const dateColumns = new Set<number>()
    for (const cell of Object.values(sheet.cells)) {
      if (++work % IMPORT_CHUNK === 0) yield
      cell.style = resolveSheetCellStyle(sourceSheet, cell.row, cell.col, cell, source)
      const sourceFormulaContext = readFormulaImportSource(cell)
      const forced = dateRanges.get(sourceSheet.id)?.some(range => sheetRangeContainsCell(range, cell.row, cell.col)) === true
      const sourceNumeric = typeof cell.value === 'number' ? cell.value
        : forced && typeof cell.value === 'string' && cell.value.trim() ? Number(cell.value) : undefined
      const sourceDateContext = sourceNumeric === undefined ? undefined : readFormulaCalendarValue(cell, sourceNumeric, source.date1904)
      const sourceFormula = sourceSheet.cells[`${cell.row}:${cell.col}`]?.formula
      const previous = cell.value
      // Validate the old date marker before changing its expected expression.
      remapChartSourceDateCell(cell, delta, target.date1904, forced)
      if (cell.formula) cell.formula = remapFormula(cell.formula, sourceSheet.id)
      if (cell.authoredStyleReference) cell.authoredStyleReference = styleIds.get(cell.authoredStyleReference) ?? cell.authoredStyleReference
      if (cell.themeStyle !== undefined) cell.themeStyle = styleSource(cell.themeStyle)
      if (cell.formula) {
        cell.metadata = { ...cell.metadata, chartDependencyFormulaSource: { version: 1, formula: cell.formula,
          aliasSheetId: sheetIds.get(sourceFormulaContext?.aliasSheetId ?? sourceSheet.id)!,
          date1904: sourceFormulaContext?.date1904 ?? source.date1904,
          dateResult: sourceFormulaContext?.dateResult === true || forced || formulaCellUsesCalendar(cell) } }
        charge(cell.metadata.chartDependencyFormulaSource, 64)
      }
      if (cell.value !== previous || sourceDateContext) {
        if (cell.value !== previous) dateColumns.add(cell.col)
        cell.metadata = { ...cell.metadata, chartDependencyDateSource: { version: 1,
          date1904: sourceDateContext?.date1904 ?? source.date1904, value: sourceDateContext?.value ?? previous,
          targetDate1904: target.date1904, targetValue: cell.value, ...(cell.formula ? { targetFormula: cell.formula } : {}),
          ...(sourceFormula ? { formula: sourceFormula } : {}) } }
        if (sourceFormula) diagnostics.add('date-formula-source-calendar-retained')
        if (typeof cell.value === 'number' && (cell.value < 0 || !source.date1904 && Number(previous) >= 60 && Number(previous) < 61)) {
          diagnostics.add('unrepresentable-calendar-date')
        }
        charge(cell.metadata.chartDependencyDateSource, 64)
      }
      // Floating drawings and cell-image/annotation authoring belong to their
      // own clipboard. These hidden sheets retain only chart data dependencies.
      delete cell.image; delete cell.sparkline; delete cell.hyperlink
      if (cell.richText) cell.richText = cell.richText.map(run => { const next = { ...run }; delete next.hyperlink; return next })
      if (typeof cell.metadata?.pivotTableId === 'string') cell.metadata = { ...cell.metadata, pivotTableId: pivotIds.get(cell.metadata.pivotTableId)! }
    }
    if (sheet.autoFilter && !sourceRanges.some(range => contains(range, sheet.autoFilter!.range))) delete sheet.autoFilter
    if (sheet.autoFilter) sheet.autoFilter.columns = Object.fromEntries(Object.entries(sheet.autoFilter.columns)
      .map(([column, rule]) => [column, remapChartSourceDateFilter(rule, delta, target.date1904, dateColumns.has(Number(column)))]))
    for (const table of tables.filter(table => table.sheetId === sheetId)) if (table.autoFilter) table.autoFilter.columns = Object.fromEntries(Object.entries(table.autoFilter.columns)
      .map(([column, rule]) => [column, remapChartSourceDateFilter(rule, delta, target.date1904, dateColumns.has(Number(column)))]))
    convertedDateColumns.set(sourceSheet.id, dateColumns)
    sheets.push(sheet)
  }
  const resourceWorkbook: WorkbookModel = { ...target, sheets: [...target.sheets, ...sheets], pivotCaches: [] }, cacheIds = new Map<string, string>()
  const copyReport = (original: PivotTableModel, owned = false): PivotTableModel => {
    const next = cloneModel(original), sourceSheetId = sheetIds.get(original.source.subUnitId)
    if (!sourceSheetId) throw new Error('透视图的原始数据没有进入依赖导入。')
    next.id = owned ? randomId('owned-pivot') : pivotIds.get(original.id)!
    next.revision = 0
    next.source = { ...next.source, unitId: target.id!, subUnitId: sourceSheetId, sheetName: sheets.find(sheet => sheet.id === sourceSheetId)!.name }
    next.anchor = { ...next.anchor, unitId: target.id!, subUnitId: owned ? options.targetSheetId : sheetIds.get(original.anchor.subUnitId)! }
    copyPivotReportCache(source, resourceWorkbook, original, next, cacheIds)
    prepareCopiedPivotSourceReport(next, owned); remapChartSourcePivotDates(next, delta, convertedDateColumns.get(original.source.subUnitId))
    return next
  }
  const pivotTables = [...selectedPivots.values()].map(pivot => copyReport(pivot)), chart = cloneModel(clipboard.chart)
  chart.sheetId = options.targetSheetId
  if (chart.pivotTableId) chart.pivotTableId = pivotIds.get(chart.pivotTableId)!
  if (chart.ownedPivot) chart.ownedPivot = copyReport(chart.ownedPivot, true)
  const report = chart.pivotTableId ? pivotTables.find(pivot => pivot.id === chart.pivotTableId) : chart.ownedPivot
  chart.source = { ...chart.source, range: cloneModel(chart.source.range), sheetId: report
    ? chart.ownedPivot ? report.source.subUnitId : report.anchor.subUnitId : sheetIds.get(chart.source.sheetId ?? clipboard.chart.sheetId)! }
  remapChartSourceAxisDates(chart, delta)
  const images: ReturnType<typeof readChartImageResources> = Object.create(null), registry = readChartImageResources(target)
  for (const [id, image] of Object.entries(clipboard.images)) {
    charge(image, 256)
    const mapped = registry[id] && JSON.stringify(registry[id]) !== JSON.stringify(image) ? randomId('chart-image') : id
    images[mapped] = normalizeChartImageResource({ ...image, id: mapped })
    const cloud = chart.options?.wordCloud
    if (cloud?.maskImageId === id) cloud.maskImageId = mapped
    if (cloud?.maskImage === id) cloud.maskImage = mapped
    if (cloud?.maskImageResource?.id === id) cloud.maskImageResource = { ...cloud.maskImageResource, id: mapped }
  }
  readChartImageResources({ metadata: { [CHART_IMAGE_RESOURCE_KEY]: { ...registry, ...images } } })
  charge(chart, 1024); charge(tables); charge(pivotTables); charge(resourceWorkbook.pivotCaches); charge(importedNames)
  const plan: ChartSourceImportPlan = { version: 1, sourceUnitId: source.id, targetUnitId: target.id, sourceChartId: clipboard.source.chartId, sourceSnapshot: clipboard.source.snapshot, sourceState,
    targetSheetId: options.targetSheetId, chart, sheets, tables, pivotTables, pivotCaches: resourceWorkbook.pivotCaches ?? [], styles, rangeThemes,
    names: importedNames, images, sourceSheetIds: [...cells.keys()], sourceRanges: queue, diagnostics: [...diagnostics], estimatedBytes, copiedCells }
  assertChartSourceImportState(source, plan)
  return plan
}
export function planChartSourceImport(source: WorkbookModel, target: WorkbookModel, clipboard: ChartObjectClipboard, options: ChartSourceImportOptions): ChartSourceImportPlan {
  return consumeChartPreparationSync(importSteps(source, target, clipboard, options), options)
}
export function planChartSourceImportAsync(source: WorkbookModel, target: WorkbookModel, clipboard: ChartObjectClipboard, options: ChartSourceImportOptions): Promise<ChartSourceImportPlan> {
  return consumeChartPreparation(importSteps(source, target, clipboard, options), options)
}
/** Recheck dependency permissions after cooperative preparation and before publication. */
export function assertChartSourceImportSource(source: WorkbookModel, plan: ChartSourceImportPlan, userId: string): void {
  if (source.id !== plan.sourceUnitId) throw new TypeError('图表导入计划的来源工作簿无效。')
  const chart = source.charts?.find(chart => chart.id === plan.sourceChartId)
  if (!chart || JSON.stringify(chart) !== plan.sourceSnapshot) throw new Error('复制源图表已更改或删除，请重新复制。')
  assertChartReadPermission(source, chart, userId)
  assertChartSourceImportState(source, plan)
  for (const sheetId of plan.sourceSheetIds) {
    const sheet = source.sheets.find(sheet => sheet.id === sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, userId)) throw new Error('图表定义名称所属工作表已关闭或不可查看。')
  }
  for (const scope of plan.sourceRanges) {
    const sheet = source.sheets.find(sheet => sheet.id === scope.sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, userId) || !canViewProtectedRanges(sheet, [scope.range], userId)) {
      throw new Error('没有查看图表全部数据依赖的权限。')
    }
  }
}
export { applyChartSourceImport } from './chart-source-import-apply'
