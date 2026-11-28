import { DEFAULT_SHEET_ROW_COUNT, DEFAULT_SHEET_COLUMN_COUNT, defaultCellStyle, type CellScalar, type PivotTableModel, type SheetTableModel, type SheetConditionalFormatRule, type SheetRangeProtectionRule, type SheetWorksheetProtectionRule, type SpreadsheetCell, type SpreadsheetSheet, type WorkbookModel } from '../domain/model.ts'
import { normalizeWorkbookSparkline } from './workbookSparkline.ts'
import { conditionalFormatRuleTypes, normalizeConditionalFormatOfficeSource } from '../features/conditional-formatting/conditional-format-office-source.ts'
import { assertWorkbookConditionalFormatBudget } from '../features/conditional-formatting/conditional-format-workbook-budget.ts'
import type { ParseEnvelope } from './contracts.ts'
import { cloneModel } from '../domain/serialization.ts'
import { array, finite, positive, record, text } from './validation.ts'
import { normalizeWorkbookAxisDimensions } from './workbookAxisDimensions.ts'
import { normalizeWorkbookAutoFilter } from './workbookAutoFilter.ts'
import { normalizeWorkbookDataValidations } from './workbookDataValidation.ts'
import { normalizeWorkbookCharts } from './workbookCharts.ts'
import { refreshAutoFilter } from '../features/sheet-sort-filter/sort-filter-engine.ts'
import { normalizeSlicerCaches } from '../features/slicer/slicer-connections.ts'
import { importSlicerStyleResources } from '../features/slicer/slicer-style-resources.ts'
import { applySheetOutlineImport, reconcileSheetOutlineImportedVisibility } from '../features/sheet-outline/sheet-outline-import.ts'
import { applyWorkbookFormulaCalendarSource, normalizeWorkbookFormulaCalendarSource } from './workbookFormulaCalendar.ts'
import { conditionalFormatUtf8Size } from '../features/conditional-formatting/conditional-format-source-budget.ts'

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
    sparkline: normalizeWorkbookSparkline(raw.sparkline, row, col),
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

function normalizeSheet(value: unknown, index: number, id: string, reconcileFilters: Array<(workbook: WorkbookModel) => void>): SpreadsheetSheet {
  const raw = record(value)
  const axis = normalizeWorkbookAxisDimensions(raw)
  const autoFilter = normalizeWorkbookAutoFilter(raw.autoFilter)
  const cells: Record<string, SpreadsheetCell> = {}
  let rowCount = axis.rowCount ?? DEFAULT_SHEET_ROW_COUNT
  let columnCount = axis.columnCount ?? DEFAULT_SHEET_COLUMN_COUNT
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
    visible: raw.visible !== false && !['hidden', 'veryhidden'].includes(text(raw.visible).toLowerCase()),
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
    metadata: { source: raw, ...(text(raw.visible).toLowerCase() === 'veryhidden' ? { hiddenState: 'veryHidden' } : {}),
      axisDimensions: { sourceUnit: axis.compatibility, valuesKeptForLegacyView: axis.compatibility === 'legacyUnspecified' } },
  }
  const outlineImport = applySheetOutlineImport(normalized, raw)
  const filterExtension = record(record(raw.sourceAxisDimensions).projectAutoFilterExtension)
  const axesExtension = record(record(raw.sourceAxisDimensions).projectAxisExtension)
  if (autoFilter && raw.sourceAutoFilter && filterExtension.restored !== true) {
    // Standard XLSX has one hidden bit for multiple causes. Criteria identify
    // failed rows; exact project geometry retains manual overlaps separately.
    // Defer until all sheets, names and the date system exist. A color rule can
    // depend on inherited styles or a conditional formula in another sheet.
    reconcileFilters.push(workbook => {
      if (!workbook.sheets.includes(normalized)) return
      const filteredRows = new Set(refreshAutoFilter(normalized, workbook))
      if (axesExtension.restored !== true && !(outlineImport.nativeSource && outlineImport.hiddenReasonsExact)) normalized.hiddenRows = normalized.hiddenRows?.filter(row => !filteredRows.has(row))
      reconcileSheetOutlineImportedVisibility(normalized, outlineImport, [...filteredRows])
    })
  }
  return normalized
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
        id: text(field.id), sourceFieldId: text(field.sourceFieldId, String(finite(field.sourceColumn) - startCol)), sourceColumn: finite(field.sourceColumn), name: text(field.name), area, position: finite(field.position),
        ...(typeof field.displayName === 'string' ? { displayName: field.displayName } : {}),
        ...(aggregation === 'average' || aggregation === 'count' || aggregation === 'countNums' || aggregation === 'max' || aggregation === 'min' || aggregation === 'product' || aggregation === 'stdDev' || aggregation === 'stdDevp' || aggregation === 'sum' || aggregation === 'var' || aggregation === 'varp' ? { aggregation } : {}),
      }]
    })
    if (!fields.length) return []
    const options = record(raw.options)
    return [{
      id: text(raw.id, `pivot-${index + 1}`),
      ...(text(raw.pivotCacheId || source.pivotCacheId) ? { pivotCacheId: text(raw.pivotCacheId || source.pivotCacheId) } : {}),
      source: { unitId, subUnitId: sourceSheetId, sheetName: text(source.sheetName), range: { startRow, startCol, endRow, endCol },
        ...(typeof source.pivotCacheId === 'string' && source.pivotCacheId.length ? { pivotCacheId: source.pivotCacheId } : {}) },
      anchor: { unitId, subUnitId: targetSheetId, row: outputStartRow, col: outputStartCol },
      outputRange: { startRow: outputStartRow, startCol: outputStartCol, endRow: outputEndRow, endCol: outputEndCol },
      fields,
      options: {
        repeatRowLabels: options.repeatRowLabels === true,
        showRowGrandTotal: options.showRowGrandTotal !== false,
        showColGrandTotal: options.showColGrandTotal !== false,
        showRowSubtotals: options.showRowSubtotals !== false,
        showColSubtotals: options.showColSubtotals !== false,
        valuePosition: options.valuePosition === 'row' ? 'row' : options.valuePosition === 'none' ? 'none' : 'column',
        ...(typeof options.valueIndex === 'number' && Number.isSafeInteger(options.valueIndex) && options.valueIndex >= 0 ? { valueIndex: options.valueIndex } : {}),
      },
      revision: finite(raw.revision),
    }]
  })
}

function normalizeTables(value: unknown, sheets: SpreadsheetSheet[]): SheetTableModel[] {
  let sourceBytes = 0
  const positiveId = (value: unknown): number | undefined => typeof value === 'number' && Number.isSafeInteger(value) && value > 0 && value <= 0xffff_ffff ? value : undefined
  const optionalText = (raw: Record<string, unknown>, key: string): string | null | undefined => {
    const value = raw[key]
    if (value !== undefined && value !== null && typeof value !== 'string') throw new TypeError(`TABLE_SOURCE_INVALID: ${key} 必须是文本或空值。`)
    return value
  }
  const optionalBoolean = (raw: Record<string, unknown>, key: string): boolean | undefined => {
    const value = raw[key]
    if (value !== undefined && typeof value !== 'boolean') throw new TypeError(`TABLE_SOURCE_INVALID: ${key} 必须是布尔值。`)
    return value
  }
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
      const officeColumnId = positiveId(item.officeColumnId)
      return name ? [{ id: text(item.id) || `${text(raw.id)}-column-${index + 1}`, name, ...(officeColumnId ? { officeColumnId } : {}),
        ...(item.formula !== undefined ? { formula: optionalText(item, 'formula') } : {}),
        ...(item.formulaIsArray !== undefined ? { formulaIsArray: optionalBoolean(item, 'formulaIsArray') } : {}),
        ...(item.totalsRowFormula !== undefined ? { totalsRowFormula: optionalText(item, 'totalsRowFormula') } : {}),
        ...(item.totalsRowFunction !== undefined ? { totalsRowFunction: optionalText(item, 'totalsRowFunction') } : {}),
        ...(item.totalsRowLabel !== undefined ? { totalsRowLabel: optionalText(item, 'totalsRowLabel') } : {}),
        ...(item.showFilterButton !== undefined ? { showFilterButton: optionalBoolean(item, 'showFilterButton') } : {}),
      }] : []
    })
    if (!sheet || range.endCol < range.startCol || range.endRow < range.startRow || !text(raw.name) || !columns.length) return []
    const source = record(raw.officeSource)
    let officeSource: SheetTableModel['officeSource']
    if (raw.officeSource !== undefined && raw.officeSource !== null) {
      if (typeof source.partName !== 'string' || !source.partName || typeof source.xml !== 'string' || !source.xml) throw new TypeError('TABLE_SOURCE_INVALID: 结构化表格来源无效。')
      const bytes = conditionalFormatUtf8Size(source.xml, 16 * 1024 * 1024)
      if (bytes > 16 * 1024 * 1024 || sourceBytes + bytes > 64 * 1024 * 1024) throw new RangeError('TABLE_SOURCE_LIMIT: 结构化表格来源超过字节预算。')
      sourceBytes += bytes
      officeSource = { partName: source.partName, xml: source.xml }
    }
    const officeTableId = positiveId(raw.officeTableId), officeSheetId = positiveId(raw.officeSheetId)
    return [{
      id: text(raw.id) || text(raw.name), name: text(raw.name), sheetId: sheet.id, range, columns,
      showHeaderRow: raw.showHeaderRow !== false, showTotalsRow: raw.showTotalsRow === true,
      styleName: text(raw.styleName) || 'TableStyleMedium2',
      ...(raw.showFirstColumn !== undefined ? { showFirstColumn: optionalBoolean(raw, 'showFirstColumn') } : {}),
      ...(raw.showLastColumn !== undefined ? { showLastColumn: optionalBoolean(raw, 'showLastColumn') } : {}),
      ...(raw.showRowStripes !== undefined ? { showRowStripes: optionalBoolean(raw, 'showRowStripes') } : {}),
      ...(raw.showColumnStripes !== undefined ? { showColumnStripes: optionalBoolean(raw, 'showColumnStripes') } : {}),
      ...(raw.autoFilter === null ? { autoFilter: null } : raw.autoFilter !== undefined ? { autoFilter: normalizeWorkbookAutoFilter(raw.autoFilter) } : {}),
      ...(officeTableId ? { officeTableId } : {}), ...(officeSheetId ? { officeSheetId } : {}), ...(officeSource ? { officeSource } : {}),
    }]
  })
}

export function normalizeWorkbook(envelope: ParseEnvelope): WorkbookModel {
  const data = envelope.payload.data
  assertWorkbookConditionalFormatBudget(data.sheets)
  const formulaCalendarSource = normalizeWorkbookFormulaCalendarSource(data.formulaCalendarSource)
  const id = crypto.randomUUID()
  const reconcileFilters: Array<(workbook: WorkbookModel) => void> = []
  const sheets = array(data.sheets).map((sheet, index) => normalizeSheet(sheet, index, id, reconcileFilters))
  assertWorkbookConditionalFormatBudget(sheets)
  const pivotTables = normalizePivotTables(data.pivotTables, sheets, id)
  const chartImport = normalizeWorkbookCharts(data.sheets, sheets, pivotTables)
  if (!sheets.length) throw new Error('解析结果不包含工作表，未用空白工作簿替代原文件')
  const workbook: WorkbookModel = {
    id, kind: 'workbook', name: envelope.fileName, sheets,
    activeSheetId: sheets[finite(data.activeSheet)]?.id ?? sheets[0]!.id,
    date1904: data.date1904 === true, modified: false,
    pivotTables,
    tables: normalizeTables(data.tables, sheets),
    charts: chartImport.charts,
    metadata: {
      format: envelope.format, parserVersion: envelope.parserVersion,
      definedNames: array(data.definedNames), localDefinedNames: array(data.localDefinedNames),
      properties: record(data.properties), structureProtected: data.structureProtected, textTruncated: data.textTruncated,
      ...(chartImport.unresolved.length ? { unresolvedOfficeCharts: chartImport.unresolved } : {}),
    },
  }
  const styleImport = importSlicerStyleResources(data)
  if (styleImport.slicerStyles) workbook.slicerStyles = styleImport.slicerStyles
  if (styleImport.timelineStyles) workbook.timelineStyles = styleImport.timelineStyles
  if (styleImport.metadata) Object.assign(workbook.metadata!, styleImport.metadata)
  applyWorkbookFormulaCalendarSource(workbook, data.sheets, data.tables, formulaCalendarSource)
  normalizeSlicerCaches(workbook)
  for (const reconcile of reconcileFilters) reconcile(workbook)
  return workbook
}
