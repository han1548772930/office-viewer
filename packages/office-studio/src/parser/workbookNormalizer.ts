import { defaultCellStyle, type CellScalar, type ChartModel, type PivotTableModel, type SheetTableModel, type SheetConditionalFormatRule, type SheetRangeProtectionRule, type SheetWorksheetProtectionRule, type SparklineSpec, type SpreadsheetCell, type SpreadsheetSheet, type WorkbookModel } from '../domain/model.ts'
import { conditionalFormatRuleTypes, normalizeConditionalFormatOfficeSource } from '../features/conditional-formatting/conditional-format-office-source.ts'
import { assertWorkbookConditionalFormatBudget } from '../features/conditional-formatting/conditional-format-workbook-budget.ts'
import type { ParseEnvelope } from './contracts.ts'
import { cloneModel } from '../domain/serialization.ts'
import { array, finite, positive, record, text } from './validation.ts'
import { normalizeWorkbookAxisDimensions } from './workbookAxisDimensions.ts'
import { normalizeWorkbookAutoFilter } from './workbookAutoFilter.ts'
import { normalizeWorkbookDataValidations } from './workbookDataValidation.ts'
import { refreshAutoFilter } from '../features/sheet-sort-filter/sort-filter-engine.ts'

function normalizeCell(value: unknown): SpreadsheetCell {
  const raw = record(value)
  const row = Math.max(0, Math.floor(finite(raw.row)))
  const col = Math.max(0, Math.floor(finite(raw.col)))
  const style = record(raw.style)
  const font = record(style.font)
  const sourceRotation = finite(style.rotation)
  const wrap = style.wrap === true
  const base = defaultCellStyle()
  const scalar: CellScalar = raw.value === null || typeof raw.value === 'string' ||
    typeof raw.value === 'boolean' || (typeof raw.value === 'number' && Number.isFinite(raw.value)) ? raw.value : null
  const formula = typeof raw.formula === 'string' && raw.formula.length ? (raw.formula.startsWith('=') ? raw.formula : '='+raw.formula) : null
  return {
    id: row+':'+col, row, col, value: scalar, formula,
    display: text(raw.display, scalar === null ? '' : String(scalar)),
    dataType: text(raw.dataType),
    cachedDataType: typeof raw.cachedDataType === 'string' ? raw.cachedDataType : null,
    richText: array(raw.richText).map(record),
    hyperlink: typeof raw.hyperlink === 'string' ? raw.hyperlink : null,
    sparkline: normalizeSparkline(raw.sparkline),
    style: {
      ...base,
      font: {
        family: text(font.name, base.font.family), size: positive(font.sizePt, base.font.size),
        bold: font.bold === true, italic: font.italic === true, underline: font.underline === true,
        strike: font.strikethrough === true, color: text(font.color) || null,
      },
      fill: text(style.fill) || null,
      numberFormat: text(style.numberFormat, 'General'),
      horizontal: style.horizontal === 'left' || style.horizontal === 'center' || style.horizontal === 'right' || style.horizontal === 'justify' ? style.horizontal : 'general',
      vertical: style.vertical === 'top' || style.vertical === 'bottom' ? style.vertical : 'middle',
      wrap, rotation: sourceRotation === 255 ? 0 : sourceRotation,
      rotationVertical: style.rotationVertical === true || sourceRotation === 255, indent: finite(style.indent),
      shrinkToFit: !wrap && style.shrinkToFit === true, border: record(style.border), metadata: style,
    },
    metadata: { source: raw },
  }
}

function normalizeSparkline(value: unknown): SparklineSpec | undefined {
  const raw = record(value)
  const type = raw.type
  if (type !== 'line' && type !== 'column' && type !== 'winloss') return undefined
  const color = text(raw.color)
  const sourceRange = text(raw.sourceRange)
  if (!/^#[0-9a-f]{6}$/i.test(color) || !/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/i.test(sourceRange)) return undefined
  const optionalColor = (candidate: unknown) => typeof candidate === 'string' && /^#[0-9a-f]{6}$/i.test(candidate) ? candidate : undefined
  const emptyCells = raw.emptyCells === 'zero' || raw.emptyCells === 'connect' ? raw.emptyCells : 'gap'
  const axis = raw.axis === 'custom' ? 'custom' : 'individual'
  const minimum = finite(raw.min, Number.NaN)
  const maximum = finite(raw.max, Number.NaN)
  return {
    type, sourceRange: sourceRange.toUpperCase(), color,
    negativeColor: optionalColor(raw.negativeColor), highPointColor: optionalColor(raw.highPointColor), lowPointColor: optionalColor(raw.lowPointColor),
    emptyCells, axis,
    ...(axis === 'custom' && Number.isFinite(minimum) && Number.isFinite(maximum) && minimum < maximum ? { min: minimum, max: maximum } : {}),
  }
}

function normalizeSheet(value: unknown, index: number, id: string): SpreadsheetSheet {
  const raw = record(value)
  const axis = normalizeWorkbookAxisDimensions(raw)
  const autoFilter = normalizeWorkbookAutoFilter(raw.autoFilter)
  const cells: Record<string, SpreadsheetCell> = {}
  let rowCount = axis.rowCount ?? 100
  let columnCount = axis.columnCount ?? 26
  for (const source of array(raw.cells)) {
    const cell = normalizeCell(source)
    cells[cell.id] = cell
    rowCount = Math.max(rowCount, cell.row + 1)
    columnCount = Math.max(columnCount, cell.col + 1)
  }
  const dimensions = record(raw.dimensions)
  const frozen = record(raw.freezePanes)
  const dataValidations = normalizeWorkbookDataValidations(raw.dataValidations)
  // The workbook guard has counted this exact property representation before
  // any clone. Materialize it before projecting fields so pure toJSON keys do
  // not change when an individual rule or extension is cloned at another root.
  const conditionalInput: unknown[] = Array.isArray(raw.conditionalFormats)
    ? cloneModel({ conditionalFormats: raw.conditionalFormats }).conditionalFormats : []
  const conditionalFormats = conditionalInput.flatMap((value): SheetConditionalFormatRule[] => {
    const item = record(value)
    if (typeof item.id !== 'string' || typeof item.type !== 'string') return []
    if (!conditionalFormatRuleTypes.includes(item.type as SheetConditionalFormatRule['type'])) return []
    const officeSource = normalizeConditionalFormatOfficeSource(item.officeSource)
    if (item.type === 'office' && !officeSource) throw new TypeError('源文件条件格式缺少往返来源。')
    const operators: Record<string, SheetConditionalFormatRule['operator']> = {
      equal: 'equal', notequal: 'notEqual', greaterthan: 'greaterThan', greaterthanorequal: 'greaterThanOrEqual',
      lessthan: 'lessThan', lessthanorequal: 'lessThanOrEqual', between: 'between', notbetween: 'notBetween',
    }
    const operator = typeof item.operator === 'string' ? operators[item.operator.toLowerCase()] ?? item.operator : undefined
    if (!Array.isArray(item.ranges) || !item.ranges.length) throw new TypeError('条件格式范围无效。')
    const ranges = array(item.ranges).flatMap(rangeValue => {
      const range = record(rangeValue)
      const startRow = finite(range.startRow, -1), startCol = finite(range.startCol, -1)
      const endRow = finite(range.endRow, -1), endCol = finite(range.endCol, -1)
      return Number.isInteger(startRow) && Number.isInteger(startCol) && Number.isInteger(endRow) && Number.isInteger(endCol) &&
        startRow >= 0 && startCol >= 0 && endRow >= startRow && endCol >= startCol && endRow <= 1_048_575 && endCol <= 16_383
        ? [{ startRow, startCol, endRow, endCol }]
        : []
    })
    if (ranges.length !== array(item.ranges).length) throw new TypeError('条件格式范围无效。')
    if (!ranges.length || (item.type === 'cellIs' && (!operator || typeof item.formula1 !== 'string'))) return []
    const formulaAnchor = record(item.formulaAnchor)
    if (item.formulaAnchor !== undefined && (typeof formulaAnchor.row !== 'number' || typeof formulaAnchor.col !== 'number'
      || !Number.isSafeInteger(formulaAnchor.row) || !Number.isSafeInteger(formulaAnchor.col)
      || formulaAnchor.row < 0 || formulaAnchor.row > 1_048_575 || formulaAnchor.col < 0 || formulaAnchor.col > 16_383)) {
      throw new TypeError('条件格式公式基准无效。')
    }
    const style = record(item.style)
    const config = record(item.config)
    const extensions = (source: Record<string, unknown>, known: readonly string[]): Record<string, unknown> =>
      cloneModel(Object.fromEntries(Object.entries(source).filter(([key]) => !known.includes(key))))
    const normalizeValueConfig = (input: unknown): { type: string; value?: number | string } | undefined => {
      const value = record(input)
      if (typeof value.type !== 'string') return undefined
      return { ...extensions(value, ['type', 'value']), type: value.type, ...(typeof value.value === 'string' || typeof value.value === 'number' ? { value: value.value } : {}) }
    }
    const thresholds = array(config.thresholds).flatMap(entry => {
      const threshold = record(entry), value = normalizeValueConfig(threshold.value)
      return value && typeof threshold.color === 'string' ? [{ ...extensions(threshold, ['value', 'color']), value, color: threshold.color }] : []
    })
    const icons = array(config.icons).flatMap(entry => {
      const icon = record(entry), value = normalizeValueConfig(icon.value)
      return value ? [{ ...extensions(icon, ['value', 'operator', 'iconType', 'iconId']), value, ...(typeof icon.operator === 'string' ? { operator: icon.operator } : {}), ...(typeof icon.iconType === 'string' ? { iconType: icon.iconType } : {}), ...(typeof icon.iconId === 'string' || typeof icon.iconId === 'number' ? { iconId: String(icon.iconId) } : {}) }] : []
    })
    return [{
      ...extensions(item, ['id', 'type', 'subType', 'ranges', 'formulaAnchor', 'operator', 'formula1', 'formula2', 'value', 'isBottom', 'isPercent', 'config', 'style', 'priority', 'stopIfTrue', 'metadata', 'officeSource']),
      id: item.id, type: item.type as SheetConditionalFormatRule['type'], ranges,
      ...(typeof formulaAnchor.row === 'number' && typeof formulaAnchor.col === 'number' ? { formulaAnchor: { row: formulaAnchor.row, col: formulaAnchor.col } } : {}),
      ...(typeof item.subType === 'string' ? { subType: item.subType as SheetConditionalFormatRule['subType'] } : {}),
      operator: operator ?? '', formula1: typeof item.formula1 === 'string' ? item.formula1 : '',
      ...(typeof item.formula2 === 'string' ? { formula2: item.formula2 } : {}),
      ...(typeof item.value === 'string' || typeof item.value === 'number' || Array.isArray(item.value) ? { value: cloneModel(item.value) as string | number | [number, number] } : {}),
      ...(typeof item.isBottom === 'boolean' ? { isBottom: item.isBottom } : {}),
      ...(typeof item.isPercent === 'boolean' ? { isPercent: item.isPercent } : {}),
      ...(Object.keys(config).length ? { config: {
        ...extensions(config, ['min', 'max', 'isGradient', 'positiveColor', 'nativeColor', 'thresholds', 'icons', 'isShowValue', 'iconSet', 'reverse', 'minLength', 'maxLength', 'border', 'borderColor', 'negativeBorderColor', 'axisColor', 'negativeBarColorSameAsPositive', 'negativeBarBorderColorSameAsPositive', 'direction', 'axisPosition']),
        ...(normalizeValueConfig(config.min) ? { min: normalizeValueConfig(config.min) } : {}),
        ...(normalizeValueConfig(config.max) ? { max: normalizeValueConfig(config.max) } : {}),
        ...(typeof config.isGradient === 'boolean' ? { isGradient: config.isGradient } : {}),
        ...(typeof config.positiveColor === 'string' ? { positiveColor: config.positiveColor } : {}),
        ...(typeof config.nativeColor === 'string' ? { nativeColor: config.nativeColor } : {}),
        ...(thresholds.length ? { thresholds } : {}), ...(icons.length ? { icons } : {}),
        ...(typeof config.isShowValue === 'boolean' ? { isShowValue: config.isShowValue } : {}),
        ...(typeof config.iconSet === 'string' ? { iconSet: config.iconSet } : {}),
        ...(typeof config.reverse === 'boolean' ? { reverse: config.reverse } : {}),
        ...(typeof config.minLength === 'number' ? { minLength: config.minLength } : {}),
        ...(typeof config.maxLength === 'number' ? { maxLength: config.maxLength } : {}),
        ...(typeof config.border === 'boolean' ? { border: config.border } : {}),
        ...(typeof config.borderColor === 'string' ? { borderColor: config.borderColor } : {}),
        ...(typeof config.negativeBorderColor === 'string' ? { negativeBorderColor: config.negativeBorderColor } : {}),
        ...(typeof config.axisColor === 'string' ? { axisColor: config.axisColor } : {}),
        ...(typeof config.negativeBarColorSameAsPositive === 'boolean' ? { negativeBarColorSameAsPositive: config.negativeBarColorSameAsPositive } : {}),
        ...(typeof config.negativeBarBorderColorSameAsPositive === 'boolean' ? { negativeBarBorderColorSameAsPositive: config.negativeBarBorderColorSameAsPositive } : {}),
        ...(typeof config.direction === 'string' && ['context', 'leftToRight', 'rightToLeft'].includes(config.direction) ? { direction: config.direction as NonNullable<SheetConditionalFormatRule['config']>['direction'] } : {}),
        ...(typeof config.axisPosition === 'string' && ['automatic', 'middle', 'none'].includes(config.axisPosition) ? { axisPosition: config.axisPosition as NonNullable<SheetConditionalFormatRule['config']>['axisPosition'] } : {}),
      } } : {}),
      style: {
        ...extensions(style, ['fill', 'fontColor', 'bold', 'italic', 'underline', 'strike']),
        ...(typeof style.fill === 'string' ? { fill: style.fill } : {}),
        ...(typeof style.fontColor === 'string' ? { fontColor: style.fontColor } : {}),
        ...(typeof style.bold === 'boolean' ? { bold: style.bold } : {}),
        ...(typeof style.italic === 'boolean' ? { italic: style.italic } : {}),
        ...(typeof style.underline === 'boolean' ? { underline: style.underline } : {}),
        ...(typeof style.strike === 'boolean' ? { strike: style.strike } : {}),
      },
      priority: Math.max(1, Math.floor(finite(item.priority, 1))), stopIfTrue: item.stopIfTrue === true,
      ...(Object.keys(record(item.metadata)).length ? { metadata: cloneModel(record(item.metadata)) } : {}),
      ...(officeSource ? { officeSource } : {}),
    }]
  })
  const rangeProtectionRules = array(raw.rangeProtectionRules).flatMap((value): SheetRangeProtectionRule[] => {
    const item = record(value)
    const id = typeof item.id === 'string' ? item.id : ''
    const permissionId = typeof item.permissionId === 'string' ? item.permissionId : id
    const ownerUserId = typeof item.ownerUserId === 'string' ? item.ownerUserId : ''
    const ranges = array(item.ranges).flatMap(rangeValue => {
      const range = record(rangeValue)
      const startRow = finite(range.startRow, -1), startCol = finite(range.startCol, -1)
      const endRow = finite(range.endRow, -1), endCol = finite(range.endCol, -1)
      return Number.isInteger(startRow) && Number.isInteger(startCol) && Number.isInteger(endRow) && Number.isInteger(endCol) &&
        startRow >= 0 && startCol >= 0 && endRow >= startRow && endCol >= startCol && endRow <= 1_048_575 && endCol <= 16_383
        ? [{ startRow, startCol, endRow, endCol }]
        : []
    })
    if (!id || !permissionId || !ownerUserId || !ranges.length) return []
    return [{
      id, permissionId, ownerUserId, ranges,
      description: typeof item.description === 'string' ? item.description : undefined,
      allowedUserIds: array(item.allowedUserIds).filter((userId): userId is string => typeof userId === 'string'),
      editState: item.editState === 'designedUserCanEdit' ? 'designedUserCanEdit' : 'onlyMe',
      viewState: item.viewState === 'noOneElseCanView' ? 'noOneElseCanView' : 'othersCanView',
      metadata: record(item.metadata),
    }]
  })
  const worksheetProtectionInput = record(raw.worksheetProtection)
  const worksheetProtection: SheetWorksheetProtectionRule | undefined = typeof worksheetProtectionInput.id === 'string' &&
    typeof worksheetProtectionInput.permissionId === 'string' &&
    typeof worksheetProtectionInput.ownerUserId === 'string'
    ? {
      id: worksheetProtectionInput.id,
      permissionId: worksheetProtectionInput.permissionId,
      unitType: 'worksheet' as const,
      description: typeof worksheetProtectionInput.description === 'string' ? worksheetProtectionInput.description : undefined,
      ownerUserId: worksheetProtectionInput.ownerUserId,
      allowedUserIds: array(worksheetProtectionInput.allowedUserIds).filter((userId): userId is string => typeof userId === 'string'),
      editState: worksheetProtectionInput.editState === 'designedUserCanEdit' ? 'designedUserCanEdit' as const : 'onlyMe' as const,
      viewState: worksheetProtectionInput.viewState === 'noOneElseCanView' ? 'noOneElseCanView' as const : 'othersCanView' as const,
    }
    : undefined
  const normalized: SpreadsheetSheet = {
    id: id+'-sheet-'+index, name: text(raw.name, 'Sheet'+(index+1)),
    visible: raw.visible !== 'hidden' && raw.visible !== 'veryHidden' && raw.visible !== false,
    cells,
    mergedRanges: array(raw.mergedRanges).map((value) => {
      const range = record(value)
      return { startRow: finite(range.firstRow), startCol: finite(range.firstCol), endRow: finite(range.lastRow), endCol: finite(range.lastCol) }
    }),
    axisUnit: axis.axisUnit, columnWidths: axis.columnWidths, rowHeights: axis.rowHeights,
    rowHeightSpans: axis.rowHeightSpans, columnWidthSpans: axis.columnWidthSpans,
    rowAutoHeightSpans: axis.rowAutoHeightSpans, rowAutoHeights: axis.rowAutoHeights,
    frozenRows: Math.max(0, Math.floor(finite(frozen.rows))),
    frozenCols: Math.max(0, Math.floor(finite(frozen.columns))),
    showGridlines: raw.gridlines !== false,
    hiddenRows: axis.hiddenRows,
    hiddenColumns: axis.hiddenColumns,
    rowHiddenSpans: axis.rowHiddenSpans,
    columnHiddenSpans: axis.columnHiddenSpans,
    autoFilter,
    dataValidations,
    conditionalFormats,
    rangeProtectionRules,
    worksheetProtection,
    defaultColumnWidth: axis.defaultColumnWidth, defaultRowHeight: axis.defaultRowHeight, defaultRowAutoHeight: axis.defaultRowAutoHeight,
    rowCount: Math.max(rowCount, axis.rowExtent, finite(dimensions.lastRow, -1) + 1, autoFilter ? autoFilter.range.endRow + 1 : 0),
    columnCount: Math.max(columnCount, axis.columnExtent, finite(dimensions.lastCol, -1) + 1, autoFilter ? autoFilter.range.endCol + 1 : 0),
    zoom: positive(raw.zoom, 100) / 100,
    metadata: { source: raw, axisDimensions: { sourceUnit: axis.compatibility, valuesKeptForLegacyView: axis.compatibility === 'legacyUnspecified' } },
  }
  const filterExtension = record(record(raw.sourceAxisDimensions).projectAutoFilterExtension)
  const axesExtension = record(record(raw.sourceAxisDimensions).projectAxisExtension)
  if (autoFilter && raw.sourceAutoFilter && filterExtension.restored !== true) {
    // Standard XLSX has one hidden bit for multiple causes. Criteria identify
    // failed rows; exact project geometry retains manual overlaps separately.
    const filteredRows = new Set(refreshAutoFilter(normalized))
    if (axesExtension.restored !== true) normalized.hiddenRows = normalized.hiddenRows?.filter(row => !filteredRows.has(row))
  }
  return normalized
}

function normalizePivotCharts(rawSheets: unknown, sheets: SpreadsheetSheet[], pivots: PivotTableModel[]): ChartModel[] {
  return array(rawSheets).flatMap((sheetValue, sheetIndex) => {
    const rawSheet = record(sheetValue)
    const sheetId = sheets[sheetIndex]?.id
    if (!sheetId) return []
    return array(rawSheet.charts).flatMap((chartValue, chartIndex): ChartModel[] => {
      const raw = record(chartValue)
      const pivotTableId = typeof raw.pivotTableId === 'string' ? raw.pivotTableId : ''
      const pivot = pivots.find(candidate => candidate.id === pivotTableId)
      if (!pivot || !pivot.outputRange) return []
      const type = normalizeChartKind(raw.kind)
      if (!type) return []
      const drawing = record(raw.drawing)
      const fromOffset = Array.isArray(drawing.fromOffsetEmu) ? drawing.fromOffsetEmu : []
      const sizeEmu = Array.isArray(drawing.absoluteSizeEmu) ? drawing.absoluteSizeEmu : []
      return [{
        id: `pivot-chart-${sheetIndex + 1}-${chartIndex + 1}`, sheetId, type,
        source: { range: { ...pivot.outputRange }, orientation: 'columns', sheetId: pivot.anchor.subUnitId },
        position: { x: Math.max(0, finite(fromOffset[0]) / 9525), y: Math.max(0, finite(fromOffset[1]) / 9525) },
        size: { width: Math.max(1, finite(sizeEmu[0], 480 * 9525) / 9525), height: Math.max(1, finite(sizeEmu[1], 320 * 9525) / 9525) },
        ...(typeof raw.title === 'string' ? { title: raw.title } : {}),
        legend: { visible: raw.legend !== false, position: normalizeChartLegendPosition(raw.legendPosition) }, categoryField: 0, valueFields: [], pivotTableId,
        xAxis: normalizeChartAxis(raw.xAxis, raw.xAxisTitle), yAxis: normalizeChartAxis(raw.yAxis, raw.yAxisTitle), revision: 0,
      }]
    })
  })
}

function normalizeRangeCharts(rawSheets: unknown, sheets: SpreadsheetSheet[]): ChartModel[] {
  return array(rawSheets).flatMap((sheetValue, sheetIndex) => {
    const sheet = sheets[sheetIndex]
    if (!sheet) return []
    return array(record(sheetValue).charts).flatMap((chartValue, chartIndex): ChartModel[] => {
      const raw = record(chartValue)
      const type = normalizeChartKind(raw.kind)
      if (!type || raw.pivotTableId) return []
      const vectors = array(raw.series).map((value) => {
        const series = record(value)
        return {
          x: parseChartVector(series.categories, sheets, sheet),
          y: parseChartVector(series.values, sheets, sheet),
        }
      })
      if (!vectors.length || vectors.some((vector) => !vector.x || !vector.y)) return []
      const first = vectors[0]!.x!
      const orientation = first.startRow === first.endRow ? 'rows' : 'columns'
      const ranges = vectors.flatMap((vector) => [vector.x!, vector.y!])
      if (ranges.some(range => range.sheetId !== first.sheetId)) return []
      const aligned = orientation === 'columns'
        ? ranges.every((range) => range.startCol === range.endCol && range.startRow === first.startRow && range.endRow === first.endRow)
        : ranges.every((range) => range.startRow === range.endRow && range.startCol === first.startCol && range.endCol === first.endCol)
      if (!aligned || (orientation === 'columns' ? first.startRow === 0 : first.startCol === 0)) return []
      const range = {
        startRow: Math.min(...ranges.map((value) => value.startRow)) - (orientation === 'columns' ? 1 : 0),
        startCol: Math.min(...ranges.map((value) => value.startCol)) - (orientation === 'rows' ? 1 : 0),
        endRow: Math.max(...ranges.map((value) => value.endRow)),
        endCol: Math.max(...ranges.map((value) => value.endCol)),
      }
      const field = (vector: NonNullable<ReturnType<typeof parseChartVector>>) => orientation === 'columns' ? vector.startCol - range.startCol : vector.startRow - range.startRow
      if (type !== 'scatter' && vectors.some(vector => field(vector.x!) !== field(first))) return []
      const drawing = record(raw.drawing)
      const offset = array(drawing.fromOffsetEmu), size = array(drawing.absoluteSizeEmu)
      return [{
        id: `${type}-chart-${sheetIndex + 1}-${chartIndex + 1}`, sheetId: sheet.id, type,
        source: { range, orientation, sheetId: first.sheetId }, categoryField: field(first), valueFields: vectors.map((vector) => field(vector.y!)),
        ...(type === 'scatter' ? { seriesToCategory: Object.fromEntries(vectors.map((vector) => [field(vector.y!), [field(vector.x!)]])) } : {}),
        position: { x: Math.max(0, finite(offset[0]) / 9525), y: Math.max(0, finite(offset[1]) / 9525) },
        size: { width: Math.max(1, finite(size[0], 480 * 9525) / 9525), height: Math.max(1, finite(size[1], 300 * 9525) / 9525) },
        ...(typeof raw.title === 'string' ? { title: raw.title } : {}),
        legend: {
          visible: raw.legend !== false,
          position: normalizeChartLegendPosition(raw.legendPosition),
        },
        xAxis: normalizeChartAxis(raw.xAxis, raw.xAxisTitle), yAxis: normalizeChartAxis(raw.yAxis, raw.yAxisTitle), revision: 0,
      }]
    })
  })
}

function normalizeChartKind(value: unknown): ChartModel['type'] | undefined {
  return value === 'bar' || value === 'line' || value === 'area' || value === 'pie' || value === 'doughnut'
    || value === 'stacked-column' || value === 'scatter' || value === 'radar' ? value : undefined
}

function normalizeChartLegendPosition(value: unknown): ChartModel['legend']['position'] {
  return value === 'top' || value === 'right' || value === 'left' ? value : 'bottom'
}

function normalizeChartAxis(value: unknown, title: unknown): ChartModel['xAxis'] {
  const raw = record(value)
  const logBase = finite(raw.logBase, 10)
  return {
    ...(typeof title === 'string' ? { title } : {}), scaleType: raw.scaleType === 'log' ? 'log' : 'linear',
    ...(raw.scaleType === 'log' && logBase >= 2 && logBase <= 1000 ? { logBase } : {}),
  }
}

function parseChartVector(value: unknown, sheets: readonly SpreadsheetSheet[], host: SpreadsheetSheet):
  { sheetId: string; startRow: number; startCol: number; endRow: number; endCol: number } | undefined {
  if (typeof value !== 'string') return undefined
  const match = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/i.exec(value.trim().replace(/^=/, ''))
  if (!match) return undefined
  const sheetName = match[1]?.replaceAll("''", "'") ?? match[2]
  const source = sheetName === undefined ? host : sheets.find(sheet => sheet.name.toLowerCase() === sheetName.toLowerCase())
  if (!source) return undefined
  const column = (label: string) => [...label.toUpperCase()].reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1
  const range = { sheetId: source.id, startRow: Number(match[4]) - 1, startCol: column(match[3]!), endRow: Number(match[6] ?? match[4]) - 1, endCol: column(match[5] ?? match[3]!) }
  return [range.startRow, range.startCol, range.endRow, range.endCol].every(Number.isSafeInteger)
    && range.startRow <= range.endRow && range.startCol <= range.endCol && range.endRow < 1_048_576 && range.endCol < 16_384 ? range : undefined
}

function normalizePivotTables(value: unknown, sheets: SpreadsheetSheet[], unitId: string): PivotTableModel[] {
  const sheetIdByName = new Map(sheets.map(sheet => [sheet.name, sheet.id]))
  return array(value).flatMap((value, index): PivotTableModel[] => {
    const raw = record(value)
    const source = record(raw.source)
    const anchor = record(raw.anchor)
    const sourceSheetId = sheetIdByName.get(text(source.sheetName))
    const targetSheetId = sheetIdByName.get(text(anchor.subUnitId))
    if (!sourceSheetId || !targetSheetId) return []

    const sourceRange = record(source.range)
    const outputRange = record(raw.outputRange)
    const startRow = finite(sourceRange.startRow, -1)
    const startCol = finite(sourceRange.startCol, -1)
    const endRow = finite(sourceRange.endRow, -1)
    const endCol = finite(sourceRange.endCol, -1)
    const outputStartRow = finite(outputRange.startRow, -1)
    const outputStartCol = finite(outputRange.startCol, -1)
    const outputEndRow = finite(outputRange.endRow, -1)
    const outputEndCol = finite(outputRange.endCol, -1)
    if (startRow < 0 || startCol < 0 || endRow < startRow || endCol < startCol || outputStartRow < 0 || outputStartCol < 0 || outputEndRow < outputStartRow || outputEndCol < outputStartCol) return []

    const fields = array(raw.fields).flatMap((fieldValue): PivotTableModel['fields'] => {
      const field = record(fieldValue)
      const area = field.area
      if (area !== 'row' && area !== 'column' && area !== 'filter' && area !== 'value') return []
      const aggregation = field.aggregation
      return [{
        id: text(field.id), sourceColumn: finite(field.sourceColumn), name: text(field.name), area, position: finite(field.position),
        ...(typeof field.displayName === 'string' ? { displayName: field.displayName } : {}),
        ...(aggregation === 'average' || aggregation === 'count' || aggregation === 'countNums' || aggregation === 'max' || aggregation === 'min' || aggregation === 'product' || aggregation === 'stdDev' || aggregation === 'stdDevp' || aggregation === 'sum' || aggregation === 'var' || aggregation === 'varp' ? { aggregation } : {}),
      }]
    })
    if (!fields.length) return []
    const options = record(raw.options)
    return [{
      id: text(raw.id, `pivot-${index + 1}`),
      source: { unitId, subUnitId: sourceSheetId, sheetName: text(source.sheetName), range: { startRow, startCol, endRow, endCol } },
      anchor: { unitId, subUnitId: targetSheetId, row: outputStartRow, col: outputStartCol },
      outputRange: { startRow: outputStartRow, startCol: outputStartCol, endRow: outputEndRow, endCol: outputEndCol },
      fields,
      options: {
        repeatRowLabels: options.repeatRowLabels === true,
        showRowGrandTotal: options.showRowGrandTotal !== false,
        showColGrandTotal: options.showColGrandTotal !== false,
        showRowSubtotals: options.showRowSubtotals !== false,
        showColSubtotals: options.showColSubtotals !== false,
        valuePosition: options.valuePosition === 'row' ? 'row' : 'column',
      },
      revision: finite(raw.revision),
    }]
  })
}

function normalizeTables(value: unknown, sheets: SpreadsheetSheet[]): SheetTableModel[] {
  return array(value).flatMap((entry): SheetTableModel[] => {
    const raw = record(entry)
    const sheet = sheets[finite(raw.sheetIndex)] ?? sheets.find(candidate => candidate.name === text(raw.sheetName))
    const rangeValue = record(raw.range)
    const range = {
      startRow: finite(rangeValue.startRow), startCol: finite(rangeValue.startCol),
      endRow: finite(rangeValue.endRow), endCol: finite(rangeValue.endCol),
    }
    const columns = array(raw.columns).flatMap((column, index) => {
      const item = record(column)
      const name = text(item.name)
      return name ? [{ id: text(item.id) || `${text(raw.id)}-column-${index + 1}`, name }] : []
    })
    if (!sheet || range.endCol < range.startCol || range.endRow < range.startRow || !text(raw.name) || !columns.length) return []
    return [{
      id: text(raw.id) || text(raw.name), name: text(raw.name), sheetId: sheet.id, range, columns,
      showHeaderRow: raw.showHeaderRow !== false, showTotalsRow: raw.showTotalsRow === true,
      styleName: text(raw.styleName) || 'TableStyleMedium2',
    }]
  })
}

export function normalizeWorkbook(envelope: ParseEnvelope): WorkbookModel {
  const data = envelope.payload.data
  assertWorkbookConditionalFormatBudget(data.sheets)
  const id = crypto.randomUUID()
  const sheets = array(data.sheets).map((sheet, index) => normalizeSheet(sheet, index, id))
  assertWorkbookConditionalFormatBudget(sheets)
  const pivotTables = normalizePivotTables(data.pivotTables, sheets, id)
  if (!sheets.length) throw new Error('解析结果不包含工作表，未用空白工作簿替代原文件')
  return {
    id, kind: 'workbook', name: envelope.fileName, sheets,
    activeSheetId: sheets[finite(data.activeSheet)]?.id ?? sheets[0]!.id,
    date1904: data.date1904 === true, modified: false,
    pivotTables,
    tables: normalizeTables(data.tables, sheets),
    charts: [...normalizePivotCharts(data.sheets, sheets, pivotTables), ...normalizeRangeCharts(data.sheets, sheets)],
    metadata: {
      format: envelope.format, parserVersion: envelope.parserVersion,
      definedNames: array(data.definedNames), localDefinedNames: array(data.localDefinedNames),
      properties: record(data.properties), structureProtected: data.structureProtected, textTruncated: data.textTruncated,
    },
  }
}
