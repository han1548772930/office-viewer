import type { CellRange, ChartModel, ChartType, PivotChartControls, PivotTableModel, WorkbookModel } from '../../../domain/model.ts'
import type { ChartConfig } from '../../charts/chart-engine.ts'
import { chartColumnName, chartConfigFromDraft, chartRangeNotation, createChartSettingsDraft, parseChartSourceRange, type ChartSettingsDraft } from '../../charts/chart-settings-draft.ts'
import { chartCellPosition, chartPositionCell, type ChartPlacementGeometry } from '../../charts/chart-placement'
import { cloneModel } from '../../../domain/serialization.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../sheet-protection/range-protection.ts'
import { sheetDrawingRefAffectedRanges } from '../../sheet-drawings/sheet-drawing-host.ts'
import { createDetachedPivotReport, getPivotSourceFields, type PivotSourceFieldInfo, type PivotTableSourceInput } from './pivot-engine.ts'
import type { PivotChartFieldsConfig } from './pivot-chart-engine.ts'
import { getPivotChartFieldItems } from './pivot-chart-engine.ts'
import type { PivotChartConfigureInput, PivotChartInfo, PivotChartInsertOptions } from './pivot-chart-api-types.ts'
import { normalizePivotChartControls } from './pivot-chart-controls.ts'

export type PivotChartCreationMode = 'chartOnly' | 'withPivotTable'
export const PIVOT_CHART_CREATION_MODES: ReadonlyArray<{ value: PivotChartCreationMode; label: string }> = [
  { value: 'chartOnly', label: '仅创建透视图' }, { value: 'withPivotTable', label: '同时创建透视表和透视图' },
]

export interface PivotChartDraft {
  chartId: string
  pivotTableId: string
  type: ChartType
  title: string
  position: string
  legendVisible: boolean
  legendPosition: 'top' | 'right' | 'bottom' | 'left'
  settings: ChartSettingsDraft
  initialPositionText: string
  sourceKind: 'range' | 'pivotTable'
  sourceSheetId: string
  sourceRange: string
  fieldsConfig: PivotChartFieldsConfig
  initialFieldsConfig: PivotChartFieldsConfig
  createMode: PivotChartCreationMode
  pivotTableSheetId: string
  pivotTableAddress: string
  controls: PivotChartControls
}

export type PivotChartCreationDraft = Pick<PivotChartDraft, 'chartId' | 'sourceKind' | 'sourceSheetId' | 'sourceRange'
  | 'createMode' | 'pivotTableSheetId' | 'pivotTableAddress'>

export function createPivotChartDraft(chart?: ChartModel, pivotTableId = '', geometry?: ChartPlacementGeometry, report?: PivotTableModel): PivotChartDraft {
  const position = chart ? chartPositionCell(chart.position, geometry) : geometry
    ? `${chartColumnName(Math.min(5, geometry.columns.count - 1))}${Math.min(1, geometry.rows.count - 1) + 1}` : 'F2'
  const settings = createChartSettingsDraft(chart ?? {
    type: 'bar', source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, orientation: 'columns' },
    pivotTableId, position: chartCellPosition(position, geometry), size: { width: 480, height: 300 },
  })
  const fieldsConfig = pivotChartFieldsFromReport(report ?? chart?.ownedPivot)
  return {
    chartId: chart?.id ?? '',
    pivotTableId: chart?.pivotTableId ?? pivotTableId,
    settings,
    initialPositionText: position,
    position,
    sourceKind: chart?.ownedPivot || !chart?.pivotTableId && !pivotTableId ? 'range' : 'pivotTable',
    sourceSheetId: report?.source.subUnitId ?? chart?.ownedPivot?.source.subUnitId ?? chart?.source.sheetId ?? chart?.sheetId ?? '',
    sourceRange: report || chart?.ownedPivot ? chartRangeNotation((report ?? chart!.ownedPivot!).source.range) : '',
    fieldsConfig,
    initialFieldsConfig: cloneModel(fieldsConfig),
    createMode: 'chartOnly',
    pivotTableSheetId: report?.anchor.subUnitId ?? chart?.sheetId ?? '',
    pivotTableAddress: report ? chartRangeNotation({ startRow: report.anchor.row, endRow: report.anchor.row,
      startCol: report.anchor.col, endCol: report.anchor.col }) : 'A1',
    get controls() { return normalizePivotChartControls(this.settings.config.controls ?? {}) },
    set controls(value: PivotChartControls) { this.settings.config.controls = normalizePivotChartControls(value) },
    get type() { return this.settings.config.type },
    set type(value: ChartType) { this.settings.config.type = value },
    get title() { return this.settings.config.title ?? '' },
    set title(value: string) { this.settings.config.title = value },
    get legendVisible() { return this.settings.config.legend?.visible ?? true },
    set legendVisible(value: boolean) { this.settings.config.legend = { ...this.settings.config.legend, visible: value } },
    get legendPosition() { return this.settings.config.legend?.position ?? 'bottom' },
    set legendPosition(value: 'top' | 'right' | 'bottom' | 'left') { this.settings.config.legend = { ...this.settings.config.legend, position: value } },
  }
}

export function pivotChartConfig(draft: PivotChartDraft, geometry?: ChartPlacementGeometry): ChartConfig & { pivotTableId: string } {
  if (!draft.pivotTableId.trim()) throw new TypeError('Select a pivot table for this chart')
  const config = chartConfigFromDraft(draft.settings)
  // The address is a placement shortcut. Leaving it unchanged preserves the
  // precise dragged position, source, size and full authored chart settings.
  return { ...config, pivotTableId: draft.pivotTableId,
    position: draft.position.trim().toUpperCase() === draft.initialPositionText ? { ...config.position } : chartCellPosition(draft.position, geometry) }
}

export function pivotChartPatch(draft: PivotChartDraft, geometry?: ChartPlacementGeometry): Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>> {
  return pivotChartConfig(draft, geometry)
}

export function addressToPosition(address: string, geometry?: ChartPlacementGeometry): { x: number; y: number } { return chartCellPosition(address, geometry) }

export interface PivotChartPanelOpenInput {
  chartId?: string
  initialFieldId?: string
  pivotTableId?: string
  sourceKind?: 'range' | 'pivotTable'
  sourceSheetId?: string
  sourceRange?: CellRange
  createMode?: PivotChartCreationMode
  pivotTableSheetId?: string
  pivotTableAddress?: string
}

export interface PivotChartPanelSource {
  report: PivotTableModel
  fields: PivotSourceFieldInfo[]
}

export function pivotChartFieldsFromReport(report?: PivotTableModel): PivotChartFieldsConfig {
  return { fields: cloneModel(report?.fields ?? []), ...(report ? { options: cloneModel(report.options), dateGroupingMode: report.dateGroupingMode ?? 'derived',
    ...(report.collapseInfo ? { collapseInfo: cloneModel(report.collapseInfo) } : {}) } : {}) }
}

/** Only the UI draft owns this report. Metadata reads never reserve or write worksheet cells. */
export function readPivotChartPanelSource(workbook: WorkbookModel, draft: PivotChartDraft, hostSheetId: string, userId = ''): PivotChartPanelSource {
  let report: PivotTableModel
  const chart = draft.chartId ? workbook.charts?.find(item => item.id === draft.chartId) : undefined
  if (draft.chartId && (!chart || !chart.pivotTableId && !chart.ownedPivot)) throw new Error('所选透视图已被删除或来源已失效。')
  if (chart && (Boolean(chart.ownedPivot) !== (draft.sourceKind === 'range') || chart.pivotTableId && chart.pivotTableId !== draft.pivotTableId)) {
    throw new Error('已有透视图的来源归属不能通过设置草稿更换。')
  }
  if (draft.sourceKind === 'pivotTable') {
    const linked = workbook.pivotTables?.find(item => item.id === draft.pivotTableId)
    if (!linked) throw new Error('请选择仍然存在的数据透视表。')
    report = cloneModel(linked)
  } else if (chart?.ownedPivot) report = cloneModel(chart.ownedPivot)
  else {
    const source = pivotChartPanelRangeSource(workbook, draft)
    const sheet = workbook.sheets.find(item => item.id === source.subUnitId)!
    const range = { startRow: source.range.startRow, endRow: source.range.endRow, startCol: source.range.startColumn, endCol: source.range.endColumn }
    if (!canViewWorksheetProtection(sheet, userId) || !canViewProtectedRanges(sheet, [range], userId)) {
      throw new Error('当前用户无法访问透视图的来源区域。')
    }
    report = createDetachedPivotReport(workbook, source, { unitId: workbook.id!, subUnitId: hostSheetId, row: 0, col: 0 }, 'pivot-chart-panel-source')
  }
  if (!canViewPivotChartPanelSource(workbook, report, hostSheetId, chart, userId, draft.sourceKind === 'pivotTable')) {
    throw new Error('当前用户无法访问透视图的来源、报告或放置工作表。')
  }
  return { report, fields: getPivotSourceFields(workbook, report) }
}

export function canViewPivotChartPanelSource(workbook: WorkbookModel, report: PivotTableModel, hostSheetId: string, chart: ChartModel | undefined,
  userId: string, linked: boolean): boolean {
  const required = new Set([hostSheetId, chart?.sheetId ?? hostSheetId, report.source.subUnitId,
    ...(linked ? [report.anchor.subUnitId] : [])])
  if (report.source.unitId !== workbook.id || [...required].some(id => {
    const sheet = workbook.sheets.find(item => item.id === id)
    return !sheet || !canViewWorksheetProtection(sheet, userId)
  })) return false
  const source = workbook.sheets.find(item => item.id === report.source.subUnitId)!
  if (!canViewProtectedRanges(source, [report.source.range], userId)) return false
  if (linked) {
    const output = workbook.sheets.find(item => item.id === report.anchor.subUnitId)!
    if (!canViewProtectedRanges(output, [pivotChartPanelOutputRange(report)], userId)) return false
  }
  if (chart) {
    const host = workbook.sheets.find(item => item.id === chart.sheetId)!
    if (!canViewProtectedRanges(host, sheetDrawingRefAffectedRanges(workbook, chart.sheetId, [{ kind: 'chart', id: chart.id }]), userId)) return false
  }
  return true
}

export function canConfigurePivotChartPanel(workbook: WorkbookModel, report: PivotTableModel | undefined, hostSheetId: string,
  chart: ChartModel | undefined, userId: string, linked: boolean): boolean {
  const host = workbook.sheets.find(item => item.id === hostSheetId)
  if (!host || !canViewWorksheetProtection(host, userId) || !canEditWorksheetProtection(host, userId)) return false
  if (chart && !canEditProtectedRanges(host, sheetDrawingRefAffectedRanges(workbook, hostSheetId, [{ kind: 'chart', id: chart.id }]), userId)) return false
  if (linked && report) {
    const output = workbook.sheets.find(item => item.id === report.anchor.subUnitId)
    if (!output || !canEditWorksheetProtection(output, userId) || !canEditProtectedRanges(output, [pivotChartPanelOutputRange(report)], userId)) return false
  }
  return !report || canViewPivotChartPanelSource(workbook, report, hostSheetId, chart, userId, linked)
}

function pivotChartPanelOutputRange(report: PivotTableModel): CellRange {
  return report.outputRange ?? { startRow: report.anchor.row, endRow: report.anchor.row, startCol: report.anchor.col, endCol: report.anchor.col }
}

/** Navigation keeps the current source range, including edits made after the panel was opened. */
export function pivotChartPanelInputFromDraft(draft: PivotChartDraft): PivotChartPanelOpenInput {
  const range = parseChartSourceRange(draft.sourceRange)
  return { chartId: draft.chartId || undefined, pivotTableId: draft.pivotTableId || undefined,
    sourceKind: draft.sourceKind, sourceSheetId: draft.sourceSheetId, sourceRange: range ? { ...range } : undefined,
    createMode: draft.createMode, pivotTableSheetId: draft.pivotTableSheetId, pivotTableAddress: draft.pivotTableAddress }
}

/** Source-specific author mappings cannot be reused after choosing a different report. */
export function resetPivotChartDraftMapping(draft: PivotChartDraft): void {
  const settings = cloneModel(draft.settings)
  settings.config.valueFields = []; settings.config.categoryField = 0; settings.config.seriesToCategory = {}
  if (settings.config.options) delete settings.config.options.mapping
  draft.settings = settings
}

export function pivotChartPanelRangeSource(workbook: WorkbookModel, draft: Pick<PivotChartDraft, 'sourceSheetId' | 'sourceRange'>): PivotTableSourceInput {
  const sheet = workbook.sheets.find(item => item.id === draft.sourceSheetId), range = parseChartSourceRange(draft.sourceRange)
  if (!workbook.id || !sheet) throw new Error('请选择当前工作簿中的来源工作表。')
  if (!range || range.endRow <= range.startRow) throw new RangeError('来源区域需要包含字段标题和至少一行数据，例如 A1:C20。')
  return { unitId: workbook.id, subUnitId: sheet.id, sheetName: sheet.name,
    range: { startRow: range.startRow, startColumn: range.startCol, endRow: range.endRow, endColumn: range.endCol } }
}

export function createPivotChartPanelDraft(workbook: WorkbookModel, hostSheetId: string, input: PivotChartPanelOpenInput = {}, geometry?: ChartPlacementGeometry): PivotChartDraft {
  const chart = input.chartId ? workbook.charts?.find(item => item.id === input.chartId) : undefined
  if (input.chartId && !chart) throw new Error('所选透视图已经不存在。')
  const pivotId = chart?.pivotTableId ?? input.pivotTableId ?? ''
  const report = chart?.ownedPivot ?? workbook.pivotTables?.find(item => item.id === pivotId)
  const draft = createPivotChartDraft(chart, pivotId, geometry, report)
  if (!chart) {
    draft.sourceKind = input.sourceKind ?? (pivotId ? 'pivotTable' : 'range')
    draft.sourceSheetId = input.sourceSheetId ?? hostSheetId
    draft.sourceRange = input.sourceRange ? chartRangeNotation(input.sourceRange) : ''
    draft.createMode = draft.sourceKind === 'range' ? input.createMode ?? 'chartOnly' : 'chartOnly'
    draft.pivotTableSheetId = input.pivotTableSheetId ?? hostSheetId
    draft.pivotTableAddress = input.pivotTableAddress ?? pivotChartDefaultTableAddress(draft, hostSheetId)
  }
  return draft
}

function pivotChartDefaultTableAddress(draft: Pick<PivotChartDraft, 'sourceSheetId' | 'sourceRange'>, hostSheetId: string): string {
  const range = parseChartSourceRange(draft.sourceRange)
  if (!range || draft.sourceSheetId !== hostSheetId) return 'A1'
  return range.endCol + 2 < 16_384 ? `${chartColumnName(range.endCol + 2)}${range.startRow + 1}` : `A${range.endRow + 3}`
}

/** Destination is command metadata. The content command checks the final report extent atomically. */
export function pivotChartCreateOptions(draft: PivotChartCreationDraft, workbook: WorkbookModel, hostSheetId: string,
  userId = ''): PivotChartInsertOptions | undefined {
  if (!PIVOT_CHART_CREATION_MODES.some(option => option.value === draft.createMode)) throw new TypeError('请选择有效的透视图创建方式。')
  if (draft.createMode === 'chartOnly') return undefined
  if (draft.chartId || draft.sourceKind !== 'range') throw new Error('同时创建透视表和透视图只适用于尚未创建的区域来源。')
  if (!workbook.id) throw new Error('工作簿没有稳定标识。')
  const sheetId = draft.pivotTableSheetId || hostSheetId, sheet = workbook.sheets.find(sheet => sheet.id === sheetId)
  const range = parseChartSourceRange(draft.pivotTableAddress)
  if (!sheet || !canViewWorksheetProtection(sheet, userId) || !canEditWorksheetProtection(sheet, userId)) throw new Error('当前用户无法在所选工作表创建透视表。')
  if (!range || range.startRow !== range.endRow || range.startCol !== range.endCol) throw new Error('透视表起始位置必须是有效的单个单元格，例如 A1。')
  if (!canViewProtectedRanges(sheet, [range], userId) || !canEditProtectedRanges(sheet, [range], userId)) throw new Error('当前用户无法在所选单元格创建透视表。')
  const source = pivotChartPanelRangeSource(workbook, draft)
  if (sheetId === source.subUnitId && range.startRow >= source.range.startRow && range.startRow <= source.range.endRow
    && range.startCol >= source.range.startColumn && range.startCol <= source.range.endColumn) throw new Error('透视表起始位置不能与来源区域重叠。')
  return { pivotTable: { anchor: { unitId: workbook.id, subUnitId: sheetId, row: range.startRow, col: range.startCol } } }
}

export function pivotChartCreateOptionsError(draft: PivotChartCreationDraft, workbook: WorkbookModel | undefined, hostSheetId: string,
  userId = ''): string {
  if (!workbook) return '工作簿已经关闭。'
  try { pivotChartCreateOptions(draft, workbook, hostSheetId, userId); return '' }
  catch (reason) { return reason instanceof Error ? reason.message : String(reason) }
}

/** Controls are stored once in author config; unchanged implicit defaults do not add content. */
export function setPivotChartDraftControls(draft: PivotChartDraft, patch: Partial<PivotChartControls>): void {
  const before = draft.controls, next = normalizePivotChartControls({ ...before, ...patch })
  if (before.showFieldButtons === next.showFieldButtons && before.showExpandCollapseButtons === next.showExpandCollapseButtons) return
  draft.controls = next
}

export function pivotChartCreateInfo(draft: PivotChartDraft, workbook: WorkbookModel, hostSheetId: string, geometry?: ChartPlacementGeometry, userId?: string): PivotChartInfo {
  const { report } = readPivotChartPanelSource(workbook, draft, hostSheetId, userId)
  const { source: _source, pivotTableId: _pivotId, ownedPivot: _owned, ...appearance } = chartConfigFromDraft(draft.settings)
  const position = draft.position.trim().toUpperCase() === draft.initialPositionText ? { ...appearance.position } : chartCellPosition(draft.position, geometry)
  return { ...appearance, position,
    source: draft.sourceKind === 'range' ? { kind: 'range', sourceRangeInfo: pivotChartPanelRangeSource(workbook, draft) }
      : { kind: 'pivotTable', pivotTableRef: { unitId: workbook.id!, subUnitId: report.anchor.subUnitId, pivotTableId: report.id } },
    initialFieldsConfig: cloneModel(draft.fieldsConfig) }
}

/** Configure author options and report fields together; source/ownership is derived by the content command. */
export function pivotChartConfigureInput(draft: PivotChartDraft, geometry?: ChartPlacementGeometry): PivotChartConfigureInput {
  const { source: _source, pivotTableId: _pivotId, ownedPivot: _owned, position: initialPosition, ...appearance } = chartConfigFromDraft(draft.settings)
  return { patch: { ...appearance,
    legend: { visible: true, position: 'bottom', ...appearance.legend },
    xAxis: { scaleType: 'linear', ...appearance.xAxis }, yAxis: { scaleType: 'linear', ...appearance.yAxis },
    position: draft.position.trim().toUpperCase() === draft.initialPositionText ? { ...initialPosition } : chartCellPosition(draft.position, geometry) },
    fieldsConfig: cloneModel(draft.fieldsConfig) }
}

/** A preview is an unregistered data-only report, calculated by the same chart/Pivot service as apply. */
export function pivotChartPreviewConfig(config: ChartConfig, fieldsConfig: PivotChartFieldsConfig, source: PivotChartPanelSource, hostSheetId: string,
  pairedTable?: PivotChartInsertOptions['pivotTable']): ChartConfig {
  const report = cloneModel(source.report)
  report.anchor = pairedTable ? { ...pairedTable.anchor } : { ...report.anchor, subUnitId: hostSheetId, row: 0, col: 0 }
  report.fields = cloneModel(fieldsConfig.fields)
  report.options = { ...report.options, ...fieldsConfig.options }
  report.dateGroupingMode = fieldsConfig.dateGroupingMode ?? report.dateGroupingMode
  if (Object.hasOwn(fieldsConfig, 'collapseInfo')) {
    if (fieldsConfig.collapseInfo == null) delete report.collapseInfo
    else report.collapseInfo = cloneModel(fieldsConfig.collapseInfo)
  }
  delete report.outputRange; delete report.generatedCellKeys; delete report.drillDownMap; delete report.axisOutput
  const { pivotTableId: _pivotId, ownedPivot: _owned, ...appearance } = config
  return { ...appearance, source: { range: cloneModel(report.source.range), sheetId: report.source.subUnitId, orientation: 'columns' }, ownedPivot: report }
}

export function readPivotChartDraftFieldItems(workbook: WorkbookModel, draft: PivotChartDraft, source: PivotChartPanelSource, hostSheetId: string, fieldId: string): string[] {
  const config = pivotChartPreviewConfig(chartConfigFromDraft(draft.settings), draft.fieldsConfig, source, hostSheetId)
  return getPivotChartFieldItems(workbook, { ...config, id: 'pivot-chart-panel-preview', sheetId: hostSheetId, revision: 0,
    categoryField: config.categoryField ?? 0, valueFields: config.valueFields ?? [],
    legend: { visible: true, position: 'bottom', ...config.legend },
    xAxis: { scaleType: 'linear', ...config.xAxis }, yAxis: { scaleType: 'linear', ...config.yAxis } }, fieldId)
}
