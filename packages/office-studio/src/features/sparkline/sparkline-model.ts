import type { CellRange, SparklineEmptyCells, SparklineSpec, SparklineType, SpreadsheetSheet } from '../../domain/model.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import type { OfficeSparklineSource } from '../../domain/sparkline-office-source.ts'
import type { SparklineSourceReference } from '../../domain/sparkline-source.ts'
import { normalizeOfficeSparklineSource } from './sparkline-office-source.ts'
import { parseSparklineReference, sparklineCellAddress, sparklineRangeString, type SparklineReference } from './sparkline-reference.ts'
import { normalizeSparklineSourceReference, normalizeSparklineSourceText, parseSparklineSource, resolveSparklineSource, shiftSparklineSource, type SparklineSourceWorkbook } from './sparkline-source.ts'

export const SPARKLINE_POINT_ORDER = ['highPoint', 'firstPoint', 'lowPoint', 'lastPoint', 'negativePoint', 'markersPoint'] as const
export type SparklinePointName = typeof SPARKLINE_POINT_ORDER[number]
export type SparklineTheme = 'standard' | 'specialty' | 'distinctive' | 'soft' | 'fresh' | 'custom'
export type SparklineExtremumMode = 'individual' | 'group' | 'custom'
export const SPARKLINE_THEMES: Readonly<Record<Exclude<SparklineTheme, 'custom'>, readonly [string, string]>> = {
  standard: ['#364ef4', '#ec64a9'], specialty: ['#316af6', '#1eb1bd'], distinctive: ['#7c4293', '#fc753c'],
  soft: ['#5066f8', '#f468f9'], fresh: ['#2ebb7a', '#f8be15'],
}
export type SparklineSettingsInput = Partial<Omit<SparklineSpec, 'sourceRange' | 'sourceReference' | 'groupId' | 'officeSource' | 'extremumMin' | 'extremumMax'>> & {
  extremumMin?: Partial<NonNullable<SparklineSpec['extremumMin']>>
  extremumMax?: Partial<NonNullable<SparklineSpec['extremumMax']>>
}
export interface ResolvedSparklineSettings {
  type: SparklineType
  color: string
  themeType: SparklineTheme
  emptyCells: SparklineEmptyCells
  nonNumericCells: SparklineEmptyCells
  containHiddenCells: boolean
  axisOptions: { visible: boolean; reverse: boolean; color: string }
  extremumMin: { type: SparklineExtremumMode; value?: number }
  extremumMax: { type: SparklineExtremumMode; value?: number }
  points: Record<SparklinePointName, { visible: boolean; color: string }>
  lineWidth: number
  showGradient: boolean
  showRadius: boolean
}
export interface SparklineGroupRequest extends SparklineSettingsInput {
  targetRange: string
  sourceRange: string
  type: SparklineType
  color: string
  groupId?: string
  officeSource?: OfficeSparklineSource
  sourceReference?: SparklineSourceReference
}
export interface SparklinePlacementRequest {
  sourceRanges: readonly string[]
  targetRanges: readonly string[]
  settings?: SparklineSettingsInput
  groupId: string
}
export interface SparklineGroupInfo {
  groupId: string
  members: Array<{ address: string; sourceRange: string; sourceReference?: SparklineSourceReference }>
  sourceRanges: string[]
  targetRanges: string[]
  selectedAddresses: string[]
  settings: ResolvedSparklineSettings
  consistent: boolean
}
export type SparklineMutationRequest =
  | { kind: 'insert'; sourceRanges: readonly string[]; targetRanges: readonly string[]; settings?: SparklineSettingsInput; groupId?: string }
  | { kind: 'settings'; ranges: readonly string[]; settings: SparklineSettingsInput }
  | { kind: 'change-source'; ranges: readonly string[]; resetType: 'item' | 'group'; sourceRanges: readonly string[]; targetRanges: readonly string[]; primary?: string; settings?: SparklineSettingsInput }
  | { kind: 'combine'; ranges: readonly string[]; groupId?: string; settings?: SparklineSettingsInput }
  | { kind: 'uncombine'; ranges: readonly string[] }
  | { kind: 'clear'; ranges: readonly string[]; scope: 'item' | 'group' }
export interface SparklineMutationPlan {
  patches: Record<string, SparklineSpec | null>
  affectedRanges: CellRange[]
  sourceReferences: string[]
  changed: boolean
}
export interface SparklineMutationOptions {
  sheets?: readonly SpreadsheetSheet[]
  workbook?: SparklineSourceWorkbook<SpreadsheetSheet>
  createGroupId?: () => string
}

function validColor(value: unknown): value is string {
  return typeof value === 'string' && (/^#(?:[0-9a-f]{3}|[0-9a-f]{4}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value)
    || value === 'transparent' || /^rgba?\(\s*\d+(?:\.\d+)?\s*,\s*\d+(?:\.\d+)?\s*,\s*\d+(?:\.\d+)?(?:\s*,\s*(?:0|1|0?\.\d+))?\s*\)$/i.test(value))
}
function checkColor(value: unknown, fallback: string): string {
  if (value === undefined) return fallback
  if (!validColor(value)) throw new Error('迷你图颜色格式无效。')
  return value
}
function checkBoolean(value: unknown, fallback: boolean): boolean {
  if (value === undefined) return fallback
  if (typeof value !== 'boolean') throw new Error('迷你图显隐设置必须是布尔值。')
  return value
}
function checkMode(value: unknown, fallback: SparklineEmptyCells): SparklineEmptyCells {
  if (value === undefined) return fallback
  if (value !== 'gap' && value !== 'zero' && value !== 'connect') throw new Error('不支持的空白或非数字处理方式。')
  return value
}
function resolveExtremum(input: SparklineSettingsInput['extremumMin'], legacyAxis: SparklineSettingsInput['axis'], legacyValue: number | undefined) {
  const type = input?.type ?? (legacyAxis === 'custom' ? 'custom' : 'individual')
  if (type !== 'individual' && type !== 'group' && type !== 'custom') throw new Error('迷你图坐标轴缩放方式无效。')
  const value = input?.value ?? (legacyAxis === 'custom' ? legacyValue : undefined)
  if (value !== undefined && !Number.isFinite(value)) throw new Error('迷你图坐标轴数值必须是有限数值。')
  if (type === 'custom' && !Number.isFinite(value)) throw new Error('自定义迷你图坐标轴必须提供有限数值。')
  return { type, ...(value === undefined ? {} : { value }) }
}

function checkSettingsShape(input: SparklineSettingsInput): void {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new Error('迷你图设置必须是配置对象。')
  for (const field of ['points', 'axisOptions', 'extremumMin', 'extremumMax'] as const) {
    const value = input[field]
    if (value !== undefined && (!value || typeof value !== 'object' || Array.isArray(value))) throw new Error('迷你图分组、轴或标记配置格式无效。')
  }
  for (const name of SPARKLINE_POINT_ORDER) {
    const value = input.points?.[name]
    if (value !== undefined && (!value || typeof value !== 'object' || Array.isArray(value))) throw new Error('迷你图标记配置格式无效。')
  }
}

/** One effective configuration interprets retained Office fields and independent public settings. */
export function resolveSparklineSettings(input: SparklineSettingsInput = {}): ResolvedSparklineSettings {
  checkSettingsShape(input)
  const type = input.type ?? 'line'
  if (type !== 'line' && type !== 'column' && type !== 'winloss' && type !== 'pie') throw new Error('不支持的迷你图类型。')
  const themeType = input.themeType ?? 'standard'
  if (themeType !== 'custom' && !Object.prototype.hasOwnProperty.call(SPARKLINE_THEMES, themeType)) throw new Error('迷你图主题无效。')
  const palette = themeType === 'custom' ? SPARKLINE_THEMES.standard : SPARKLINE_THEMES[themeType]
  if (input.axis !== undefined && input.axis !== 'individual' && input.axis !== 'custom') throw new Error('迷你图坐标轴缩放方式无效。')
  const extremumMin = resolveExtremum(input.extremumMin, input.axis, input.min)
  const extremumMax = resolveExtremum(input.extremumMax, input.axis, input.max)
  if (extremumMin.type === 'custom' && extremumMax.type === 'custom' && extremumMin.value! >= extremumMax.value!) {
    throw new Error('自定义迷你图坐标轴必须提供递增的最小值和最大值。')
  }
  const point = (name: SparklinePointName, legacy?: string) => ({
    visible: checkBoolean(input.points?.[name]?.visible, legacy !== undefined), color: checkColor(input.points?.[name]?.color ?? legacy, palette[1]),
  })
  const points = { highPoint: point('highPoint', input.highPointColor), firstPoint: point('firstPoint'), lowPoint: point('lowPoint', input.lowPointColor),
    lastPoint: point('lastPoint'), negativePoint: point('negativePoint', input.negativeColor), markersPoint: point('markersPoint') }
  const lineWidth = input.lineWidth ?? 2
  if (!Number.isFinite(lineWidth) || lineWidth < 0 || lineWidth > 2112) throw new Error('迷你图线宽超出 Office 读写的有效范围。')
  return {
    type, themeType, color: checkColor(input.color, palette[0]), points, extremumMin, extremumMax,
    emptyCells: checkMode(input.emptyCells, 'gap'), nonNumericCells: checkMode(input.nonNumericCells, 'gap'),
    containHiddenCells: checkBoolean(input.containHiddenCells, false),
    axisOptions: { visible: checkBoolean(input.axisOptions?.visible, false), reverse: checkBoolean(input.axisOptions?.reverse, false), color: checkColor(input.axisOptions?.color, '#000000') },
    lineWidth, showGradient: checkBoolean(input.showGradient, false), showRadius: checkBoolean(input.showRadius, false),
  }
}

export function sparklineSettingsInput(settings: ResolvedSparklineSettings): SparklineSettingsInput {
  return { ...settings, points: Object.fromEntries(SPARKLINE_POINT_ORDER.map(name => [name, { ...settings.points[name] }])),
    axisOptions: { ...settings.axisOptions }, extremumMin: { ...settings.extremumMin }, extremumMax: { ...settings.extremumMax } }
}

export function mergeSparklineSettings(current: SparklineSpec | ResolvedSparklineSettings, patch: SparklineSettingsInput): ResolvedSparklineSettings {
  checkSettingsShape(patch)
  const previous = resolveSparklineSettings(current), next = { ...sparklineSettingsInput(previous), ...patch,
    axisOptions: { ...previous.axisOptions, ...patch.axisOptions },
    extremumMin: { ...previous.extremumMin, ...patch.extremumMin, type: patch.extremumMin?.type ?? previous.extremumMin.type },
    extremumMax: { ...previous.extremumMax, ...patch.extremumMax, type: patch.extremumMax?.type ?? previous.extremumMax.type },
    points: Object.fromEntries(SPARKLINE_POINT_ORDER.map(name => [name, { ...previous.points[name], ...patch.points?.[name] }])),
  }
  for (const [name, value] of [['negativePoint', patch.negativeColor], ['highPoint', patch.highPointColor], ['lowPoint', patch.lowPointColor]] as const) {
    if (value !== undefined && patch.points?.[name] === undefined) next.points[name] = { visible: true, color: value }
  }
  if (patch.axis !== undefined && patch.extremumMin === undefined && patch.extremumMax === undefined) {
    next.extremumMin = resolveExtremum(undefined, patch.axis, patch.min)
    next.extremumMax = resolveExtremum(undefined, patch.axis, patch.max)
  }
  if (patch.themeType !== undefined && patch.themeType !== 'custom') {
    const palette = SPARKLINE_THEMES[patch.themeType]
    if (!palette) throw new Error('迷你图主题无效。')
    next.color = palette[0]
    for (const name of SPARKLINE_POINT_ORDER) next.points[name]!.color = palette[1]
  } else if (patch.color !== undefined || SPARKLINE_POINT_ORDER.some(name => patch.points?.[name]?.color !== undefined)) next.themeType = 'custom'
  return resolveSparklineSettings(next)
}

function normalizeGroupId(value: string): string {
  if (typeof value !== 'string' || !value.trim() || value !== value.trim() || value.length > 200 || value.startsWith('ungrouped:')) throw new Error('迷你图分组标识无效。')
  return value
}
export function createSparklineGroupId(): string { return `sparkline-${globalThis.crypto.randomUUID()}` }

/** Retained imports may contain rectangular sources; the placement planner validates new series counts. */
export function createSparklineSpec(request: SparklineGroupRequest): SparklineSpec {
  const sourceRange = normalizeSparklineSourceText(request.sourceRange), target = parseSparklineReference(request.targetRange)
  if (!target) throw new Error('迷你图目标区域必须是有效单元格范围。')
  const settings = resolveSparklineSettings(request)
  const custom = settings.extremumMin.type === 'custom' && settings.extremumMax.type === 'custom'
  const advanced = request.points !== undefined || request.extremumMin !== undefined || request.extremumMax !== undefined
    || request.axisOptions !== undefined || request.themeType !== undefined || request.nonNumericCells !== undefined
    || request.containHiddenCells !== undefined || request.lineWidth !== undefined || request.showGradient !== undefined || request.showRadius !== undefined
  return {
    type: settings.type, sourceRange, color: settings.color, emptyCells: settings.emptyCells, axis: custom ? 'custom' : 'individual',
    ...(settings.points.negativePoint.visible ? { negativeColor: settings.points.negativePoint.color } : {}),
    ...(settings.points.highPoint.visible ? { highPointColor: settings.points.highPoint.color } : {}),
    ...(settings.points.lowPoint.visible ? { lowPointColor: settings.points.lowPoint.color } : {}),
    ...(custom ? { min: settings.extremumMin.value, max: settings.extremumMax.value } : {}),
    ...(request.groupId === undefined ? {} : { groupId: normalizeGroupId(request.groupId) }),
    ...(advanced ? settings : {}),
    ...(request.officeSource === undefined ? {} : { officeSource: normalizeOfficeSparklineSource(request.officeSource) }),
    ...(request.sourceReference === undefined ? {} : { sourceReference: normalizeSparklineSourceReference(request.sourceReference) }),
  }
}
export function normalizeSparklineSpec(spec: SparklineSpec, targetAddress: string): SparklineSpec {
  return createSparklineSpec({ ...spec, targetRange: targetAddress })
}

function requiredReferences(ranges: readonly string[], label: string): SparklineReference[] {
  if (!Array.isArray(ranges) || !ranges.length) throw new Error(`请指定迷你图${label}区域。`)
  return ranges.map(value => {
    const reference = parseSparklineReference(value)
    if (!reference) throw new Error(`迷你图${label}区域必须是有效 A1 范围。`)
    return reference
  })
}
function targetCount(references: readonly SparklineReference[]): number {
  let count = 0
  for (const { range } of references) count += (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1)
  return count
}
function vector(reference: SparklineReference): boolean { return reference.range.startRow === reference.range.endRow || reference.range.startCol === reference.range.endCol }
function canonicalTargets(references: readonly SparklineReference[], host?: Pick<SpreadsheetSheet, 'name' | 'id'>): SparklineReference[] {
  return references.map(reference => {
    if (reference.sheetName !== undefined && (!host || reference.sheetName.toLocaleLowerCase() !== host.name.toLocaleLowerCase() && reference.sheetName !== host.id)) {
      throw new Error('迷你图放置区域必须属于当前工作表。')
    }
    return { ...reference, normalized: sparklineRangeString(reference.range), sheetName: undefined }
  })
}

/** v1.6.0 N / Se: source rows win an ambiguous square; target orientation does not choose series direction. */
export function createSparklinePlacementPatches(request: SparklinePlacementRequest, host?: Pick<SpreadsheetSheet, 'name' | 'id'>,
  workbook?: SparklineSourceWorkbook): Record<string, SparklineSpec> {
  if (!Array.isArray(request.sourceRanges) || !request.sourceRanges.length) throw new Error('请指定迷你图源数据区域。')
  const sources = request.sourceRanges.map(value => {
    const expression = parseSparklineSource(value)
    if (expression.kind === 'broken') throw new Error('请先修复迷你图断开的来源再放置。')
    const resolved = host && workbook ? resolveSparklineSource(value, host, workbook) : undefined
    return { expression, reference: expression.kind === 'a1' ? expression.reference : resolved?.status === 'resolved' ? resolved.reference : undefined }
  }), targets = canonicalTargets(requiredReferences(request.targetRanges, '放置'), host)
  if (targets.length > 1 && targets.some(reference => targetCount([reference]) !== 1) || targets.length === 1 && !vector(targets[0]!)) {
    throw new Error('迷你图目标区域必须为单行或单列；多个目标区域必须各为单格。')
  }
  const count = targetCount(targets), source = sources[0]!, reference = source.reference
  const rows = reference ? reference.range.endRow - reference.range.startRow + 1 : 1, columns = reference ? reference.range.endCol - reference.range.startCol + 1 : 1
  if (sources.length > 1 ? sources.length !== count || sources.some(source => source.reference && !vector(source.reference))
    : source.expression.kind === 'a1' ? count !== rows && count !== columns : count !== 1) {
    throw new Error('迷你图目标数必须与源数据系列数一致。')
  }
  const groupId = normalizeGroupId(request.groupId), settings = resolveSparklineSettings(request.settings), patches: Record<string, SparklineSpec> = {}
  let offset = 0
  for (const { range } of targets) for (let row = range.startRow; row <= range.endRow; row += 1) for (let col = range.startCol; col <= range.endCol; col += 1) {
    const address = sparklineCellAddress(row, col)
    if (patches[address]) throw new Error('迷你图放置区域不能包含重复单元格。')
    const itemSource = sources.length > 1 ? normalizeSparklineSourceText(sources[offset]!.expression.text) : count === 1 ? normalizeSparklineSourceText(source.expression.text)
      : slicedSource(reference!, count === rows ? 'row' : 'column', offset)
    patches[address] = createSparklineSpec({ ...sparklineSettingsInput(settings), type: settings.type, color: settings.color, sourceRange: itemSource, targetRange: address, groupId })
    offset += 1
  }
  return patches
}
function slicedSource(source: SparklineReference, direction: 'row' | 'column', offset: number): string {
  const first = { ...source.first }, last = { ...source.last }
  if (direction === 'row') first.row = last.row = source.range.startRow + offset
  else first.column = last.column = source.range.startCol + offset
  const format = (address: typeof first) => {
    const parts = /^([A-Z]+)(\d+)$/.exec(sparklineCellAddress(address.row, address.column))!
    return `${address.absoluteColumn ? '$' : ''}${parts[1]}${address.absoluteRow ? '$' : ''}${parts[2]}`
  }
  const firstText = format(first), lastText = format(last), qualifier = source.normalized.slice(0, source.normalized.lastIndexOf('!') + 1)
  return `${qualifier}${firstText}${firstText === lastText ? '' : `:${lastText}`}`
}

export function createSparklineGroupPatches(request: SparklineGroupRequest): Record<string, SparklineSpec> {
  return createSparklinePlacementPatches({ sourceRanges: [request.sourceRange], targetRanges: [request.targetRange], settings: request, groupId: request.groupId ?? createSparklineGroupId() })
}
export function translateSparklineForMove(spec: SparklineSpec, rowDelta: number, columnDelta: number): SparklineSpec {
  return { ...spec, sourceRange: spec.sourceReference ? spec.sourceRange : shiftSparklineSource(spec.sourceRange, rowDelta, columnDelta) }
}

function groupKey(address: string, spec: SparklineSpec): string { return spec.groupId ?? `ungrouped:${address}` }
function addressInRanges(address: string, ranges: readonly SparklineReference[]): boolean {
  const reference = parseSparklineReference(address)!
  return ranges.some(({ range }) => reference.range.startRow >= range.startRow && reference.range.startRow <= range.endRow
    && reference.range.startCol >= range.startCol && reference.range.startCol <= range.endCol)
}
function sheetEntries(sheet: SpreadsheetSheet): Array<{ address: string; spec: SparklineSpec; row: number; col: number }> {
  return Object.values(sheet.cells).filter(cell => cell.sparkline !== undefined).map(cell =>
    ({ address: sparklineCellAddress(cell.row, cell.col), spec: cell.sparkline!, row: cell.row, col: cell.col })).sort((a, b) => a.row - b.row || a.col - b.col)
}

export function readSparklineGroups(sheet: SpreadsheetSheet, ranges?: readonly string[]): SparklineGroupInfo[] {
  const selectedRanges = ranges === undefined ? undefined : canonicalTargets(requiredReferences(ranges, '选择'), sheet)
  const grouped = new Map<string, SparklineGroupInfo>()
  for (const { address, spec } of sheetEntries(sheet)) {
    const key = groupKey(address, spec), settings = resolveSparklineSettings(spec)
    let group = grouped.get(key)
    if (!group) { group = { groupId: key, settings, members: [], sourceRanges: [], targetRanges: [], selectedAddresses: [], consistent: true }; grouped.set(key, group) }
    group.members.push({ address, sourceRange: spec.sourceRange, ...(spec.sourceReference === undefined ? {} : { sourceReference: normalizeSparklineSourceReference(spec.sourceReference) }) }); group.sourceRanges.push(spec.sourceRange); group.targetRanges.push(address)
    if (!selectedRanges || addressInRanges(address, selectedRanges)) group.selectedAddresses.push(address)
    if (!equivalentModelData(settings, group.settings)) group.consistent = false
  }
  return [...grouped.values()].filter(group => group.selectedAddresses.length > 0)
}

export function assertSparklineGroupConsistency(sheet: SpreadsheetSheet, patches: Readonly<Record<string, SparklineSpec | null>> = {}): void {
  const settingsByGroup = new Map<string, ResolvedSparklineSettings>(), entries = new Map(sheetEntries(sheet).map(entry => [entry.address, entry.spec]))
  for (const [address, spec] of Object.entries(patches)) { if (spec === null) entries.delete(address); else entries.set(address, spec) }
  for (const [address, spec] of entries) {
    if (!spec.groupId) continue
    normalizeGroupId(spec.groupId)
    const settings = resolveSparklineSettings(spec), previous = settingsByGroup.get(spec.groupId)
    if (previous && !equivalentModelData(previous, settings)) throw new Error(`迷你图组 ${spec.groupId} 的设置不一致，请先统一整组设置。`)
    settingsByGroup.set(spec.groupId, settings)
    if (!parseSparklineReference(address)) throw new Error('迷你图目标单元格无效。')
  }
}

function checkSources(patches: Readonly<Record<string, SparklineSpec | null>>, sheet: SpreadsheetSheet, workbook: SparklineSourceWorkbook<SpreadsheetSheet>): string[] {
  const sources = new Set<string>()
  for (const [address, spec] of Object.entries(patches)) {
    if (!spec) continue
    const resolved = resolveSparklineSource(spec, sheet, workbook)
    if (resolved.status !== 'resolved') {
      const target = parseSparklineReference(address)!.range, previous = sheet.cells[`${target.startRow}:${target.startCol}`]?.sparkline
      const retained = previous !== undefined && previous.sourceRange === spec.sourceRange && equivalentModelData(previous.sourceReference, spec.sourceReference)
      if (resolved.status === 'broken' && (resolved.source.kind === 'a1' || resolved.source.kind === 'broken' || spec.sourceReference) && !retained) {
        throw new Error('迷你图来源已断开，请通过修改数据来源修复。')
      }
      continue
    }
    sources.add(`${resolved.sheet.id}!${sparklineRangeString(resolved.range)}`)
  }
  return [...sources]
}

/** Read-only planner. Permission/version acceptance and one history commit belong to the shared command. */
export function planSparklineMutation(sheet: SpreadsheetSheet, request: SparklineMutationRequest, options: SparklineMutationOptions = {}): SparklineMutationPlan {
  const patches: Record<string, SparklineSpec | null> = {}, entries = sheetEntries(sheet), existingIds = new Set(entries.flatMap(entry => entry.spec.groupId ? [entry.spec.groupId] : []))
  const workbook = options.workbook ?? { sheets: options.sheets ?? [sheet] }
  const groupSizes = new Map<string, number>()
  for (const entry of entries) { const key = groupKey(entry.address, entry.spec); groupSizes.set(key, (groupSizes.get(key) ?? 0) + 1) }
  const newId = (requested?: string): string => {
    const id = normalizeGroupId(requested ?? (options.createGroupId ?? createSparklineGroupId)())
    if (existingIds.has(id)) throw new Error('迷你图新分组标识已被使用。')
    existingIds.add(id); return id
  }
  if (request.kind === 'insert') {
    Object.assign(patches, createSparklinePlacementPatches({ ...request, groupId: newId(request.groupId) }, sheet, workbook))
  } else {
    const ranges = canonicalTargets(requiredReferences(request.ranges, '选择'), sheet), selected = entries.filter(entry => addressInRanges(entry.address, ranges))
    if (!selected.length) throw new Error('所选区域没有迷你图。')
    const keys = new Set(selected.map(entry => groupKey(entry.address, entry.spec))), reached = entries.filter(entry => keys.has(groupKey(entry.address, entry.spec)))
    if (request.kind === 'settings') {
      for (const entry of reached) {
        const settings = mergeSparklineSettings(entry.spec, request.settings)
        patches[entry.address] = createSparklineSpec({ ...sparklineSettingsInput(settings), type: settings.type, color: settings.color,
          targetRange: entry.address, sourceRange: entry.spec.sourceRange, officeSource: entry.spec.officeSource, sourceReference: entry.spec.sourceReference, ...(entry.spec.groupId ? { groupId: entry.spec.groupId } : {}) })
      }
    } else if (request.kind === 'combine') {
      if (selected.length < 2) throw new Error('请至少选择两个迷你图后合并组。')
      const groupId = newId(request.groupId), base = selected[0]!.spec, settings = request.settings ? mergeSparklineSettings(base, request.settings) : resolveSparklineSettings(base)
      for (const entry of selected) patches[entry.address] = createSparklineSpec({ ...sparklineSettingsInput(settings), type: settings.type, color: settings.color,
        sourceRange: entry.spec.sourceRange, targetRange: entry.address, groupId, officeSource: entry.spec.officeSource, sourceReference: entry.spec.sourceReference })
    } else if (request.kind === 'uncombine') {
      for (const entry of selected) {
        if ((groupSizes.get(groupKey(entry.address, entry.spec)) ?? 0) < 2) continue
        patches[entry.address] = { ...entry.spec, groupId: newId() }
      }
    } else if (request.kind === 'clear') {
      if (request.scope !== 'item' && request.scope !== 'group') throw new Error('迷你图清除范围无效。')
      for (const entry of request.scope === 'group' ? reached : selected) patches[entry.address] = null
    } else if (request.kind === 'change-source') {
      if (request.resetType !== 'item' && request.resetType !== 'group') throw new Error('迷你图数据重设方式无效。')
      if (request.resetType === 'group' && keys.size !== 1) throw new Error('整组修改数据时只能选择一个迷你图组。')
      const primaryReference = request.primary ? parseSparklineReference(request.primary) : undefined
      if (primaryReference && targetCount([primaryReference]) !== 1) throw new Error('迷你图主单元格必须是单格。')
      const primaryAddress = primaryReference ? sparklineCellAddress(primaryReference.range.startRow, primaryReference.range.startCol) : selected[0]!.address
      const primary = selected.find(entry => entry.address === primaryAddress)
      if (!primary) throw new Error('迷你图主单元格必须属于当前选择。')
      if (!Array.isArray(request.sourceRanges) || !request.sourceRanges.length) throw new Error('请指定迷你图源数据区域。')
      const sourceReferences = request.sourceRanges.map(parseSparklineSource), targetReferences = canonicalTargets(requiredReferences(request.targetRanges, '放置'), sheet)
      if (request.resetType === 'item' && (sourceReferences.length !== 1 || targetCount(targetReferences) !== 1)) throw new Error('单项迷你图修改需要一个源区域和一个放置单元格。')
      const groupId = primary.spec.groupId ?? newId(), settings = request.settings ? mergeSparklineSettings(primary.spec, request.settings) : resolveSparklineSettings(primary.spec)
      const replacement = createSparklinePlacementPatches({ sourceRanges: request.sourceRanges, targetRanges: request.targetRanges, groupId, settings: sparklineSettingsInput(settings) }, sheet, workbook)
      const moving = request.resetType === 'group' ? reached : [primary], movingAddresses = new Set(moving.map(entry => entry.address))
      for (const address of Object.keys(replacement)) {
        if (entries.some(entry => entry.address === address) && !movingAddresses.has(address)) throw new Error('迷你图新放置区域与其他现有迷你图冲突。')
      }
      for (const entry of moving) patches[entry.address] = null
      if (request.settings && request.resetType === 'item') {
        for (const entry of entries.filter(candidate => groupKey(candidate.address, candidate.spec) === groupKey(primary.address, primary.spec))) {
          if (movingAddresses.has(entry.address)) continue
          patches[entry.address] = createSparklineSpec({ ...sparklineSettingsInput(settings), type: settings.type, color: settings.color,
            sourceRange: entry.spec.sourceRange, targetRange: entry.address, officeSource: entry.spec.officeSource, sourceReference: entry.spec.sourceReference, ...(entry.spec.groupId ? { groupId: entry.spec.groupId } : {}) })
        }
      }
      for (const [index, [address, spec]] of Object.entries(replacement).entries()) {
        const officeSource = (request.resetType === 'item' ? primary : moving[index])?.spec.officeSource ?? primary.spec.officeSource
        patches[address] = officeSource === undefined ? spec : { ...spec, officeSource: normalizeOfficeSparklineSource(officeSource) }
      }
    } else throw new Error('不支持的迷你图操作。')
  }
  const before = new Map(entries.map(entry => [entry.address, entry.spec]))
  for (const [address, spec] of Object.entries(patches)) if (equivalentModelData(before.get(address) ?? null, spec)) delete patches[address]
  const sourceReferences = checkSources(patches, sheet, workbook)
  assertSparklineGroupConsistency(sheet, patches)
  return { patches, sourceReferences, affectedRanges: Object.keys(patches).map(address => parseSparklineReference(address)!.range), changed: Object.keys(patches).length > 0 }
}
