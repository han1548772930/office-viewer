import type { SparklineSpec, WorkbookModel } from '../../domain/model.ts'
import { createCell } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { OfficeSparklineSource } from '../../domain/sparkline-office-source.ts'
import { normalizeOfficeSparklineSource } from './sparkline-office-source.ts'
import { createSparklineSpec, readSparklineGroups, sparklineSettingsInput, resolveSparklineSettings, assertSparklineGroupConsistency } from './sparkline-model.ts'
import { sparklineConfigFromPublic, sparklineConfigToPublic, type ISparklineGroupConfig } from './sparkline-api-types.ts'
import { parseSparklineReference, sparklineCellAddress, sparklineRangeString, sparklineSheetQualifier } from './sparkline-reference.ts'
import { normalizeSparklineSourceReference, parseSparklineSource, sparklineSourceToPublicRange, type SparklinePublicSourceRange } from './sparkline-source.ts'

export const SHEET_SPARKLINE_PLUGIN = 'SHEET_SPARKLINE_PLUGIN'
export type SparklineResourceRange = SparklinePublicSourceRange & { officeStudioOfficeSource?: OfficeSparklineSource }
export interface SparklineResourceGroup { config: ISparklineGroupConfig; sparklines: Record<number, Record<number, SparklineResourceRange>> }
export type SparklineResource = Record<string, Record<string, Record<string, SparklineResourceGroup>>>
const object = (value: unknown): value is Record<string, unknown> => Boolean(value && typeof value === 'object' && !Array.isArray(value))

export function serializeWorkbookSparklineResource(workbook: WorkbookModel): SparklineResource {
  if (!workbook.id) throw new TypeError('迷你图资源需要工作簿 ID。')
  const sheets: SparklineResource[string] = {}
  for (const sheet of workbook.sheets) {
    const groups: Record<string, SparklineResourceGroup> = {}
    const specs = new Map(Object.values(sheet.cells).filter(cell => cell.sparkline).map(cell => [sparklineCellAddress(cell.row, cell.col), cell.sparkline!]))
    for (const group of readSparklineGroups(sheet)) {
      const sparklines: SparklineResourceGroup['sparklines'] = {}
      for (const member of group.members) {
        const target = parseSparklineReference(member.address), spec = specs.get(member.address)
        if (!target || !spec) throw new TypeError('迷你图资源包含失效目标引用。')
        const row = target.range.startRow, col = target.range.startCol
        ;(sparklines[row] ??= {})[col] = { ...sparklineSourceToPublicRange(spec, sheet, workbook),
          ...(spec.officeSource === undefined ? {} : { officeStudioOfficeSource: cloneModel(spec.officeSource) }) }
      }
      groups[group.groupId] = { config: sparklineConfigToPublic(group.settings), sparklines }
    }
    if (Object.keys(groups).length) sheets[sheet.id] = groups
  }
  return { [workbook.id]: sheets }
}

export function importWorkbookSparklineResources(workbook: WorkbookModel, resources: unknown): void {
  if (!Array.isArray(resources)) return
  const matches = resources.filter(value => object(value) && value.name === SHEET_SPARKLINE_PLUGIN)
  if (!matches.length) return
  if (matches.length !== 1 || !object(matches[0]) || typeof matches[0].data !== 'string') throw new TypeError('迷你图插件资源重复或缺少 JSON 数据。')
  let value: unknown
  try { value = JSON.parse(matches[0].data) } catch { throw new TypeError('迷你图插件资源不是有效 JSON。') }
  const patches = deserializeWorkbookSparklineResource(workbook, value)
  // All sheets/groups are validated before applying the detached candidate.
  for (const sheet of workbook.sheets) {
    for (const cell of Object.values(sheet.cells)) delete cell.sparkline
    for (const [address, spec] of Object.entries(patches[sheet.id] ?? {})) {
      const target = parseSparklineReference(address)!, row = target.range.startRow, col = target.range.startCol, key = `${row}:${col}`
      const cell = sheet.cells[key] ?? createCell(row, col)
      cell.sparkline = cloneModel(spec); sheet.cells[key] = cell
      sheet.rowCount = Math.max(sheet.rowCount, row + 1); sheet.columnCount = Math.max(sheet.columnCount, col + 1)
    }
    assertSparklineGroupConsistency(sheet)
  }
}

export function deserializeWorkbookSparklineResource(workbook: WorkbookModel, input: unknown): Record<string, Record<string, SparklineSpec>> {
  if (!object(input)) throw new TypeError('迷你图插件资源必须是工作簿映射。')
  const units = Object.keys(input), sourceUnit = workbook.id && Object.hasOwn(input, workbook.id) ? workbook.id : units.length === 1 ? units[0] : undefined
  if (sourceUnit === undefined) { if (!units.length) return {}; throw new TypeError('迷你图资源无法确定所属工作簿。') }
  const rawSheets = input[sourceUnit]
  if (!object(rawSheets)) throw new TypeError('迷你图工作表资源无效。')
  const patches: Record<string, Record<string, SparklineSpec>> = {}
  for (const [sheetId, rawGroups] of Object.entries(rawSheets)) {
    const sheet = workbook.sheets.find(sheet => sheet.id === sheetId)
    if (!sheet || !object(rawGroups)) throw new TypeError('迷你图资源引用了不存在的工作表。')
    const targets: Record<string, SparklineSpec> = {}
    for (const [groupId, rawGroup] of Object.entries(rawGroups)) {
      if (!object(rawGroup) || !object(rawGroup.config) || !object(rawGroup.sparklines)) throw new TypeError('迷你图组资源无效。')
      const settings = resolveSparklineSettings(sparklineConfigFromPublic(rawGroup.config as ISparklineGroupConfig))
      for (const [rowKey, rowValues] of Object.entries(rawGroup.sparklines)) {
        if (!/^\d+$/.test(rowKey) || !object(rowValues)) throw new TypeError('迷你图行矩阵无效。')
        for (const [colKey, rawRange] of Object.entries(rowValues)) {
          if (!/^\d+$/.test(colKey) || !object(rawRange)) throw new TypeError('迷你图列矩阵无效。')
          const row = Number(rowKey), col = Number(colKey), address = sparklineCellAddress(row, col)
          if (Object.hasOwn(targets, address)) throw new TypeError('多个迷你图组使用了同一目标单元格。')
          const coordinates = [rawRange.startRow, rawRange.endRow, rawRange.startColumn, rawRange.endColumn]
          const sentinel = coordinates.every(value => value === -9999)
          if (!coordinates.every(value => typeof value === 'number' && Number.isSafeInteger(value))
            || typeof rawRange.startRow !== 'number' || typeof rawRange.endRow !== 'number' || typeof rawRange.startColumn !== 'number' || typeof rawRange.endColumn !== 'number'
            || !sentinel && (rawRange.startRow < 0 || rawRange.startColumn < 0 || rawRange.endRow < rawRange.startRow || rawRange.endColumn < rawRange.startColumn
              || rawRange.endRow > 1_048_575 || rawRange.endColumn > 16_383)) throw new TypeError('迷你图资源源范围无效。')
          for (const key of ['sheetId', 'unitId'] as const) {
            const identity = rawRange[key]
            if (identity !== undefined && (typeof identity !== 'string' || !identity.trim() || identity.length > 4096 || /[\u0000-\u001f\u007f]/.test(identity))) throw new TypeError('迷你图资源来源标识无效。')
          }
          if (rawRange.rangeType !== undefined && rawRange.rangeType !== 0 && rawRange.rangeType !== 1 && rawRange.rangeType !== 2 && rawRange.rangeType !== 3) throw new TypeError('迷你图资源来源类型无效。')
          const qualifier = rawRange.sheetId === undefined ? undefined : typeof rawRange.sheetId === 'string'
            ? workbook.sheets.find(sheet => sheet.id === rawRange.sheetId || sheet.name === rawRange.sheetId)?.name ?? rawRange.sheetId : undefined
          if (rawRange.sheetId !== undefined && qualifier === undefined) throw new TypeError('迷你图资源源工作表无效。')
          const reference = rawRange.officeStudioSourceReference
          if (reference !== undefined && typeof reference !== 'string') throw new TypeError('迷你图资源来源原文无效。')
          const foreignUnit = typeof rawRange.unitId === 'string' && rawRange.unitId !== workbook.id && rawRange.unitId !== sourceUnit
          if (foreignUnit && reference === undefined && !sentinel && qualifier === undefined) throw new TypeError('迷你图外部来源缺少工作表或原文。')
          const sourceRange = typeof reference === 'string' ? reference : sentinel ? '#REF!'
            : foreignUnit ? `${sparklineSheetQualifier(`[${rawRange.unitId}]${qualifier}`)}${sparklineRangeString({ startRow: rawRange.startRow, endRow: rawRange.endRow, startCol: rawRange.startColumn, endCol: rawRange.endColumn })}`
              : sparklineRangeString({ startRow: rawRange.startRow, endRow: rawRange.endRow, startCol: rawRange.startColumn, endCol: rawRange.endColumn }, qualifier)
          const parsed = parseSparklineSource(sourceRange)
          const preserved = normalizeSparklineSourceReference(rawRange.officeStudioSourceMetadata)
          if (preserved && (!sentinel || preserved.invalidRange.sheetId !== rawRange.sheetId || preserved.invalidRange.unitId !== rawRange.unitId
            || preserved.invalidRange.rangeType !== rawRange.rangeType)) throw new TypeError('迷你图断开来源的保留数据与公共范围不一致。')
          const sourceReference = preserved ?? (sentinel && (reference === undefined || parsed.kind === 'a1'
            || parsed.kind === 'broken' && (rawRange.sheetId !== undefined || rawRange.unitId !== undefined || rawRange.rangeType !== 0))
            ? normalizeSparklineSourceReference({ version: 1, invalidRange: { startRow: -9999, endRow: -9999, startColumn: -9999, endColumn: -9999,
              ...(rawRange.sheetId === undefined ? {} : { sheetId: rawRange.sheetId }), ...(rawRange.unitId === undefined ? {} : { unitId: rawRange.unitId }),
              ...(rawRange.rangeType === undefined ? {} : { rangeType: rawRange.rangeType }) } }) : undefined)
          const parsedRange = parsed.kind === 'a1' ? parsed.reference : (parsed.kind === 'external' || parsed.kind === 'three-dimensional') && parsed.addressText ? parseSparklineReference(parsed.addressText) : undefined
          if (!sentinel && parsedRange && (parsedRange.range.startRow !== rawRange.startRow || parsedRange.range.endRow !== rawRange.endRow
            || parsedRange.range.startCol !== rawRange.startColumn || parsedRange.range.endCol !== rawRange.endColumn)) throw new TypeError('迷你图资源保留引用与源范围不一致。')
          if (!sentinel && parsed.kind === 'a1') {
            const sourceSheet = parsed.reference.sheetName === undefined ? sheet : workbook.sheets.find(candidate => candidate.id === parsed.reference.sheetName || candidate.name.toLocaleLowerCase() === parsed.reference.sheetName!.toLocaleLowerCase())
            if (foreignUnit || rawRange.sheetId !== undefined && sourceSheet?.id !== rawRange.sheetId && sourceSheet?.name !== rawRange.sheetId
              && parsed.reference.sheetName !== rawRange.sheetId) throw new TypeError('迷你图资源来源工作表与保留原文不一致。')
          }
          targets[address] = createSparklineSpec({ ...sparklineSettingsInput(settings), type: settings.type, color: settings.color, targetRange: address, sourceRange,
            ...(sourceReference === undefined ? {} : { sourceReference }),
            ...(rawRange.officeStudioOfficeSource === undefined ? {} : { officeSource: normalizeOfficeSparklineSource(rawRange.officeStudioOfficeSource) }),
            ...(groupId === `ungrouped:${address}` ? {} : { groupId }) })
        }
      }
    }
    assertSparklineGroupConsistency({ ...sheet, cells: Object.fromEntries(Object.entries(targets).map(([address, sparkline]) => {
      const position = parseSparklineReference(address)!.range, cell = createCell(position.startRow, position.startCol)
      cell.sparkline = sparkline; return [cell.id, cell]
    })) })
    patches[sheetId] = targets
  }
  return patches
}
