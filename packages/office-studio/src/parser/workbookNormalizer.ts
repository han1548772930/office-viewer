import { defaultCellStyle, type CellScalar, type ChartModel, type PivotTableModel, type SheetTableModel, type SheetConditionalFormatRule, type SheetRangeProtectionRule, type SheetWorksheetProtectionRule, type SparklineSpec, type SpreadsheetCell, type SpreadsheetSheet, type WorkbookModel, type SheetDataValidationRule } from '../domain/model.ts'
import type { ParseEnvelope } from './contracts.ts'
import { array, finite, positive, record, text } from './validation.ts'

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
      horizontal: style.horizontal === 'center' || style.horizontal === 'right' || style.horizontal === 'justify' ? style.horizontal : 'left',
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
  const cells: Record<string, SpreadsheetCell> = {}
  let rowCount = 100
  let columnCount = 26
  for (const source of array(raw.cells)) {
    const cell = normalizeCell(source)
    cells[cell.id] = cell
    rowCount = Math.max(rowCount, cell.row + 1)
    columnCount = Math.max(columnCount, cell.col + 1)
  }
  const dimensions = record(raw.dimensions)
  const frozen = record(raw.freezePanes)
  const sizes = (source: unknown): Record<number, number> => {
    const result: Record<number, number> = {}
    for (const item of array(source)) {
      const axis = record(item)
      if (finite(axis.index, -1) >= 0) result[Math.floor(finite(axis.index))] = positive(axis.size, 20)
    }
    return result
  }
  const dataValidations = array(raw.dataValidations).flatMap((value): SheetDataValidationRule[] => {
    const wrapped = record(value)
    const item = record(wrapped.rule) ? record(wrapped.rule) : wrapped
    const typeNames: Record<string, SheetDataValidationRule['type']> = {
      whole: 'whole', integer: 'whole', decimal: 'decimal', textlength: 'textLength', date: 'date', time: 'time',
      checkbox: 'checkbox', list: 'list', listmultiple: 'listMultiple', custom: 'custom',
    }
    const sourceType = typeof item.type === 'string' ? item.type.toLowerCase() : ''
    const type = typeNames[sourceType]
    const idValue = typeof item.uid === 'string' ? item.uid : typeof item.id === 'string' ? item.id : ''
    if (!type || !idValue) return []
    const ranges = array(item.ranges).flatMap(rangeValue => {
      const range = record(rangeValue)
      const startRow = finite(range.startRow ?? range.firstRow, -1)
      const startCol = finite(range.startColumn ?? range.startCol ?? range.firstColumn ?? range.firstCol, -1)
      const endRow = finite(range.endRow ?? range.lastRow, startRow)
      const endCol = finite(range.endColumn ?? range.endCol ?? range.lastColumn ?? range.lastCol, startCol)
      return startRow >= 0 && startCol >= 0 && endRow >= startRow && endCol >= startCol
        ? [{ startRow, startCol, endRow, endCol }]
        : []
    })
    if (!ranges.length) return []
    const operatorNames: Record<string, NonNullable<SheetDataValidationRule['operator']>> = {
      between: 'between', equal: 'equal', greaterthan: 'greaterThan', greaterthanorequal: 'greaterThanOrEqual',
      lessthan: 'lessThan', lessthanorequal: 'lessThanOrEqual', notbetween: 'notBetween', notequal: 'notEqual',
    }
    const errorStyle = typeof item.errorStyle === 'string' ? ({ stop: 'stop', warning: 'warning', info: 'info' } as const)[item.errorStyle.toLowerCase() as 'stop' | 'warning' | 'info'] : undefined
    return [{
      id: idValue, type, ranges,
      formula1: typeof item.formula1 === 'string' ? item.formula1 : undefined,
      formula2: typeof item.formula2 === 'string' ? item.formula2 : undefined,
      operator: typeof item.operator === 'string' ? operatorNames[item.operator.toLowerCase()] : undefined,
      allowBlank: typeof item.allowBlank === 'boolean' ? item.allowBlank : undefined,
      showDropDown: typeof item.showDropDown === 'boolean' ? item.showDropDown : undefined,
      showErrorMessage: typeof item.showErrorMessage === 'boolean' ? item.showErrorMessage : undefined,
      error: typeof item.error === 'string' ? item.error : undefined,
      errorTitle: typeof item.errorTitle === 'string' ? item.errorTitle : undefined,
      errorStyle,
      showInputMessage: typeof item.showInputMessage === 'boolean' ? item.showInputMessage : undefined,
      prompt: typeof item.prompt === 'string' ? item.prompt : undefined,
      promptTitle: typeof item.promptTitle === 'string' ? item.promptTitle : undefined,
      renderMode: typeof item.renderMode === 'string' && ['text', 'arrow', 'custom'].includes(item.renderMode.toLowerCase())
        ? item.renderMode.toLowerCase() as SheetDataValidationRule['renderMode'] : undefined,
    }]
  })
  const conditionalFormats = array(raw.conditionalFormats).flatMap((value): SheetConditionalFormatRule[] => {
    const item = record(value)
    if (typeof item.id !== 'string' || typeof item.type !== 'string') return []
    const supportedTypes: SheetConditionalFormatRule['type'][] = ['cellIs', 'highlightCell', 'top10', 'average', 'duplicateValues', 'uniqueValues', 'timePeriod', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'expression', 'dataBar', 'colorScale', 'iconSet']
    if (!supportedTypes.includes(item.type as SheetConditionalFormatRule['type'])) return []
    const operators: Record<string, SheetConditionalFormatRule['operator']> = {
      equal: 'equal', notequal: 'notEqual', greaterthan: 'greaterThan', greaterthanorequal: 'greaterThanOrEqual',
      lessthan: 'lessThan', lessthanorequal: 'lessThanOrEqual', between: 'between', notbetween: 'notBetween',
    }
    const operator = typeof item.operator === 'string' ? operators[item.operator.toLowerCase()] ?? item.operator : undefined
    const ranges = array(item.ranges).flatMap(rangeValue => {
      const range = record(rangeValue)
      const startRow = finite(range.startRow, -1), startCol = finite(range.startCol, -1)
      const endRow = finite(range.endRow, -1), endCol = finite(range.endCol, -1)
      return Number.isInteger(startRow) && Number.isInteger(startCol) && Number.isInteger(endRow) && Number.isInteger(endCol) &&
        startRow >= 0 && startCol >= 0 && endRow >= startRow && endCol >= startCol
        ? [{ startRow, startCol, endRow, endCol }]
        : []
    })
    if (!ranges.length || (item.type === 'cellIs' && (!operator || typeof item.formula1 !== 'string'))) return []
    const style = record(item.style)
    const config = record(item.config)
    const normalizeValueConfig = (input: unknown): { type: string; value?: number | string } | undefined => {
      const value = record(input)
      if (typeof value.type !== 'string') return undefined
      return { type: value.type, ...(typeof value.value === 'string' || typeof value.value === 'number' ? { value: value.value } : {}) }
    }
    const thresholds = array(config.thresholds).flatMap(entry => {
      const threshold = record(entry), value = normalizeValueConfig(threshold.value)
      return value && typeof threshold.color === 'string' ? [{ value, color: threshold.color }] : []
    })
    const icons = array(config.icons).flatMap(entry => {
      const icon = record(entry), value = normalizeValueConfig(icon.value)
      return value ? [{ value, ...(typeof icon.operator === 'string' ? { operator: icon.operator } : {}), ...(typeof icon.iconType === 'string' ? { iconType: icon.iconType } : {}), ...(typeof icon.iconId === 'string' || typeof icon.iconId === 'number' ? { iconId: String(icon.iconId) } : {}) }] : []
    })
    return [{
      id: item.id, type: item.type as SheetConditionalFormatRule['type'], ranges,
      ...(typeof item.subType === 'string' ? { subType: item.subType as SheetConditionalFormatRule['subType'] } : {}),
      operator: operator ?? '', formula1: typeof item.formula1 === 'string' ? item.formula1 : '',
      ...(typeof item.formula2 === 'string' ? { formula2: item.formula2 } : {}),
      ...(typeof item.value === 'string' || typeof item.value === 'number' || Array.isArray(item.value) ? { value: item.value as string | number | [number, number] } : {}),
      ...(typeof item.isBottom === 'boolean' ? { isBottom: item.isBottom } : {}),
      ...(typeof item.isPercent === 'boolean' ? { isPercent: item.isPercent } : {}),
      ...(Object.keys(config).length ? { config: {
        ...(normalizeValueConfig(config.min) ? { min: normalizeValueConfig(config.min) } : {}),
        ...(normalizeValueConfig(config.max) ? { max: normalizeValueConfig(config.max) } : {}),
        ...(typeof config.isGradient === 'boolean' ? { isGradient: config.isGradient } : {}),
        ...(typeof config.positiveColor === 'string' ? { positiveColor: config.positiveColor } : {}),
        ...(typeof config.nativeColor === 'string' ? { nativeColor: config.nativeColor } : {}),
        ...(thresholds.length ? { thresholds } : {}), ...(icons.length ? { icons } : {}),
        ...(typeof config.isShowValue === 'boolean' ? { isShowValue: config.isShowValue } : {}),
        ...(typeof config.iconSet === 'string' ? { iconSet: config.iconSet } : {}),
        ...(typeof config.reverse === 'boolean' ? { reverse: config.reverse } : {}),
      } } : {}),
      style: {
        ...(typeof style.fill === 'string' ? { fill: style.fill } : {}),
        ...(typeof style.fontColor === 'string' ? { fontColor: style.fontColor } : {}),
        ...(typeof style.bold === 'boolean' ? { bold: style.bold } : {}),
        ...(typeof style.italic === 'boolean' ? { italic: style.italic } : {}),
        ...(typeof style.underline === 'boolean' ? { underline: style.underline } : {}),
        ...(typeof style.strike === 'boolean' ? { strike: style.strike } : {}),
      },
      priority: Math.max(1, Math.floor(finite(item.priority, 1))), stopIfTrue: item.stopIfTrue === true,
      ...(Object.keys(record(item.metadata)).length ? { metadata: record(item.metadata) } : {}),
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
  return {
    id: id+'-sheet-'+index, name: text(raw.name, 'Sheet'+(index+1)),
    visible: raw.visible !== 'hidden' && raw.visible !== 'veryHidden' && raw.visible !== false,
    cells,
    mergedRanges: array(raw.mergedRanges).map((value) => {
      const range = record(value)
      return { startRow: finite(range.firstRow), startCol: finite(range.firstCol), endRow: finite(range.lastRow), endCol: finite(range.lastCol) }
    }),
    columnWidths: sizes(raw.columnWidths), rowHeights: sizes(raw.rowHeights),
    frozenRows: Math.max(0, Math.floor(finite(frozen.rows))),
    frozenCols: Math.max(0, Math.floor(finite(frozen.columns))),
    showGridlines: raw.gridlines !== false,
    hiddenRows: array(raw.hiddenRows).filter((n): n is number => typeof n === 'number' && Number.isFinite(n)),
    hiddenColumns: array(raw.hiddenColumns).filter((n): n is number => typeof n === 'number' && Number.isFinite(n)),
    dataValidations,
    conditionalFormats,
    rangeProtectionRules,
    worksheetProtection,
    defaultColumnWidth: positive(raw.defaultColumnWidth, 96), defaultRowHeight: positive(raw.defaultRowHeight, 24),
    rowCount: Math.max(rowCount, finite(dimensions.lastRow, -1) + 1),
    columnCount: Math.max(columnCount, finite(dimensions.lastCol, -1) + 1),
    zoom: positive(raw.zoom, 100) / 100,
    metadata: { source: raw },
  }
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
      const type = raw.kind
      if (type !== 'bar' && type !== 'line' && type !== 'area' && type !== 'pie' && type !== 'doughnut' && type !== 'stacked-column') return []
      const drawing = record(raw.drawing)
      const fromOffset = Array.isArray(drawing.fromOffsetEmu) ? drawing.fromOffsetEmu : []
      const sizeEmu = Array.isArray(drawing.absoluteSizeEmu) ? drawing.absoluteSizeEmu : []
      return [{
        id: `pivot-chart-${sheetIndex + 1}-${chartIndex + 1}`, sheetId, type,
        source: { range: { ...pivot.outputRange }, orientation: 'columns' },
        position: { x: Math.max(0, finite(fromOffset[0]) / 9525), y: Math.max(0, finite(fromOffset[1]) / 9525) },
        size: { width: Math.max(1, finite(sizeEmu[0], 480 * 9525) / 9525), height: Math.max(1, finite(sizeEmu[1], 320 * 9525) / 9525) },
        ...(typeof raw.title === 'string' ? { title: raw.title } : {}),
        legend: { visible: raw.legend !== false, position: 'bottom' }, categoryField: 0, valueFields: [], pivotTableId,
        xAxis: { title: typeof raw.xAxisTitle === 'string' ? raw.xAxisTitle : undefined, scaleType: 'linear' },
        yAxis: { title: typeof raw.yAxisTitle === 'string' ? raw.yAxisTitle : undefined, scaleType: 'linear' }, revision: 0,
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
      const type = raw.kind
      if ((type !== 'scatter' && type !== 'radar') || raw.pivotTableId) return []
      const vectors = array(raw.series).map((value) => {
        const series = record(value)
        return {
          x: parseChartVector(series.categories, sheet.name),
          y: parseChartVector(series.values, sheet.name),
        }
      })
      if (!vectors.length || vectors.some((vector) => !vector.x || !vector.y)) return []
      const first = vectors[0]!.x!
      const orientation = first.startRow === first.endRow ? 'rows' : 'columns'
      const ranges = vectors.flatMap((vector) => [vector.x!, vector.y!])
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
      if (type === 'radar' && vectors.some(vector => field(vector.x!) !== field(first))) return []
      const drawing = record(raw.drawing)
      const offset = array(drawing.fromOffsetEmu), size = array(drawing.absoluteSizeEmu)
      const legendPosition = raw.legendPosition
      return [{
        id: `${type}-chart-${sheetIndex + 1}-${chartIndex + 1}`, sheetId: sheet.id, type,
        source: { range, orientation }, categoryField: field(first), valueFields: vectors.map((vector) => field(vector.y!)),
        ...(type === 'scatter' ? { seriesToCategory: Object.fromEntries(vectors.map((vector) => [field(vector.y!), [field(vector.x!)]])) } : {}),
        position: { x: Math.max(0, finite(offset[0]) / 9525), y: Math.max(0, finite(offset[1]) / 9525) },
        size: { width: Math.max(1, finite(size[0], 480 * 9525) / 9525), height: Math.max(1, finite(size[1], 300 * 9525) / 9525) },
        ...(typeof raw.title === 'string' ? { title: raw.title } : {}),
        legend: {
          visible: raw.legend !== false,
          position: legendPosition === 'top' || legendPosition === 'right' || legendPosition === 'left' ? legendPosition : 'bottom',
        },
        xAxis: normalizeChartAxis(raw.xAxis, raw.xAxisTitle), yAxis: normalizeChartAxis(raw.yAxis, raw.yAxisTitle), revision: 0,
      }]
    })
  })
}

function normalizeChartAxis(value: unknown, title: unknown): ChartModel['xAxis'] {
  const raw = record(value)
  const logBase = finite(raw.logBase, 10)
  return {
    ...(typeof title === 'string' ? { title } : {}), scaleType: raw.scaleType === 'log' ? 'log' : 'linear',
    ...(raw.scaleType === 'log' && logBase >= 2 && logBase <= 1000 ? { logBase } : {}),
  }
}

function parseChartVector(value: unknown, sheetName: string): { startRow: number; startCol: number; endRow: number; endCol: number } | undefined {
  if (typeof value !== 'string') return undefined
  const match = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/i.exec(value)
  if (!match || (match[1]?.replaceAll("''", "'") ?? match[2]) !== sheetName) return undefined
  const column = (label: string) => [...label.toUpperCase()].reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1
  const range = { startRow: Number(match[4]) - 1, startCol: column(match[3]!), endRow: Number(match[6] ?? match[4]) - 1, endCol: column(match[5] ?? match[3]!) }
  return range.startRow <= range.endRow && range.startCol <= range.endCol ? range : undefined
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
  const id = crypto.randomUUID()
  const sheets = array(data.sheets).map((sheet, index) => normalizeSheet(sheet, index, id))
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
