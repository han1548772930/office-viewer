import { FOffice } from '../../api/facade.ts'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { CellRange, SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { planSparklineMutation, resolveSparklineSettings, sparklineSettingsInput, SPARKLINE_POINT_ORDER, SPARKLINE_THEMES,
  type ResolvedSparklineSettings, type SparklineGroupInfo, type SparklineMutationRequest, type SparklineSettingsInput } from './sparkline-model.ts'
import { parseSparklineRangeList, parseSparklineReference } from './sparkline-reference.ts'

export interface SparklinePanelTarget { workbookId: string; sheetId: string }
export interface SparklinePanelOwner {
  studio: OfficeStudio
  target: SparklinePanelTarget
  revision: number
  historyGeneration: number
  userId: string
}
export interface SparklinePanelSnapshot {
  owner: SparklinePanelOwner
  ranges: string[]
  groups: SparklineGroupInfo[]
  selectedGroups: SparklineGroupInfo[]
}
export interface SparklinePanelDraft {
  settings: ResolvedSparklineSettings
  sourceText: string
  targetText: string
  dataOperation: 'settings' | 'item' | 'group'
  primary: string
}

export function captureSparklinePanelOwner(studio: OfficeStudio, target: SparklinePanelTarget): SparklinePanelOwner {
  const owner: SparklinePanelOwner = { studio, target: { ...target }, revision: studio.getWorkbookRevision(target.workbookId),
    historyGeneration: studio.editor.history.captureSavepoint(target.workbookId).generation, userId: studio.currentUserId }
  if (!sparklinePanelOwnerCurrent(owner)) throw new Error('迷你图工作台已更改，请重新打开面板。')
  return owner
}
export function sparklinePanelOwnerCurrent(owner: SparklinePanelOwner, published = false): boolean {
  const { studio, target } = owner, workbook = studio.getWorkbookModel(target.workbookId)
  if (studio.isDisposed || !workbook || studio.activeWorkbookId !== target.workbookId || studio.editor.state.activeKind !== 'workbook'
    || workbook.activeSheetId !== target.sheetId || !workbook.sheets.some(sheet => sheet.id === target.sheetId) || studio.currentUserId !== owner.userId) return false
  const revision = studio.getWorkbookRevision(target.workbookId)
  return (revision === owner.revision || published && revision === owner.revision + 1)
    && studio.editor.history.captureSavepoint(target.workbookId).generation === owner.historyGeneration
}
export function sparklinePanelOwnerContext(owner: SparklinePanelOwner) {
  if (!sparklinePanelOwnerCurrent(owner)) throw new Error('迷你图内容或历史已更改，请重新读取后再应用。')
  return { host: 'workbench' as const, expectedRevision: owner.revision, historyGeneration: owner.historyGeneration, expectedUserId: owner.userId }
}
function currentSheet(owner: SparklinePanelOwner): SpreadsheetSheet {
  const sheet = owner.studio.getWorkbookModel(owner.target.workbookId)?.sheets.find(candidate => candidate.id === owner.target.sheetId)
  if (!sheet) throw new Error('迷你图所在工作表已关闭。')
  return sheet
}

export function readSparklinePanelState(studio: OfficeStudio, target: SparklinePanelTarget, ranges: readonly string[]): SparklinePanelSnapshot {
  const owner = captureSparklinePanelOwner(studio, target), sheet = currentSheet(owner)
  if (!canViewWorksheetProtection(sheet, owner.userId)) throw new Error('当前用户无权查看迷你图所在工作表。')
  const facade = FOffice.newAPI(studio).getWorkbook(target.workbookId)?.getSheetBySheetId(target.sheetId)
  if (!facade) throw new Error('迷你图所在工作表已关闭。')
  return { owner, ranges: [...ranges], groups: facade.getSparklineGroups(), selectedGroups: facade.getSparklineGroups(ranges) }
}

export function createSparklinePanelDraft(snapshot: SparklinePanelSnapshot, groupId?: string, primary?: string,
  operation: SparklinePanelDraft['dataOperation'] = 'settings'): SparklinePanelDraft {
  const group = groupId ? snapshot.groups.find(candidate => candidate.groupId === groupId)
    : snapshot.selectedGroups.find(candidate => candidate.members.some(member => member.address === primary)) ?? snapshot.selectedGroups[0]
  const member = group?.members.find(candidate => candidate.address === primary) ?? group?.members.find(candidate => group.selectedAddresses.includes(candidate.address)) ?? group?.members[0]
  return { settings: cloneModel(group?.settings ?? resolveSparklineSettings()), primary: member?.address ?? '', dataOperation: operation,
    sourceText: operation === 'group' && group ? group.sourceRanges.join('; ') : member?.sourceRange ?? 'B1:E1',
    targetText: operation === 'group' && group ? group.targetRanges.join('; ') : member?.address ?? snapshot.ranges.join('; ') }
}

/** A draft may temporarily hold an empty custom number; publication performs the strict validation. */
export function editSparklineSettingsDraft(current: ResolvedSparklineSettings, patch: SparklineSettingsInput): ResolvedSparklineSettings {
  const next = cloneModel(current)
  for (const key of ['type', 'color', 'themeType', 'emptyCells', 'nonNumericCells', 'containHiddenCells', 'lineWidth', 'showGradient', 'showRadius'] as const) {
    if (patch[key] !== undefined) Object.assign(next, { [key]: patch[key] })
  }
  if (patch.axisOptions) next.axisOptions = { ...next.axisOptions, ...patch.axisOptions }
  if (patch.extremumMin) next.extremumMin = { ...next.extremumMin, ...patch.extremumMin, type: patch.extremumMin.type ?? next.extremumMin.type }
  if (patch.extremumMax) next.extremumMax = { ...next.extremumMax, ...patch.extremumMax, type: patch.extremumMax.type ?? next.extremumMax.type }
  for (const name of SPARKLINE_POINT_ORDER) if (patch.points?.[name]) next.points[name] = { ...next.points[name], ...patch.points[name] }
  if (patch.themeType !== undefined && patch.themeType !== 'custom') {
    const palette = SPARKLINE_THEMES[patch.themeType]
    next.color = palette[0]
    for (const name of SPARKLINE_POINT_ORDER) next.points[name].color = palette[1]
  } else if (patch.color !== undefined || SPARKLINE_POINT_ORDER.some(name => patch.points?.[name]?.color !== undefined)) next.themeType = 'custom'
  return next
}

/** Changed nested fields alone update mixed groups without replacing untouched settings. */
export function buildSparklineSettingsPatch(current: ResolvedSparklineSettings, original: ResolvedSparklineSettings): SparklineSettingsInput {
  resolveSparklineSettings(current)
  const patch: SparklineSettingsInput = {}
  for (const key of ['type', 'color', 'themeType', 'emptyCells', 'nonNumericCells', 'containHiddenCells', 'lineWidth', 'showGradient', 'showRadius'] as const) {
    if (!Object.is(current[key], original[key])) Object.assign(patch, { [key]: current[key] })
  }
  for (const key of ['visible', 'reverse', 'color'] as const) if (!Object.is(current.axisOptions[key], original.axisOptions[key])) {
    patch.axisOptions ??= {}; Object.assign(patch.axisOptions, { [key]: current.axisOptions[key] })
  }
  for (const end of ['extremumMin', 'extremumMax'] as const) for (const key of ['type', 'value'] as const) if (!Object.is(current[end][key], original[end][key])) {
    patch[end] ??= {}; Object.assign(patch[end]!, { [key]: current[end][key] })
  }
  for (const name of SPARKLINE_POINT_ORDER) for (const key of ['visible', 'color'] as const) if (!Object.is(current.points[name][key], original.points[name][key])) {
    patch.points ??= {}; patch.points[name] ??= {}; Object.assign(patch.points[name]!, { [key]: current.points[name][key] })
  }
  return patch
}
export function sparklinePanelDraftDirty(current: SparklinePanelDraft, original: SparklinePanelDraft): boolean {
  return !equivalentModelData(current, original)
}
export function buildSparklinePanelRequest(draft: SparklinePanelDraft, original: SparklinePanelDraft, ranges: readonly string[], existing: boolean): SparklineMutationRequest {
  const settings = buildSparklineSettingsPatch(draft.settings, original.settings)
  if (!existing) return { kind: 'insert', sourceRanges: parseSparklineRangeList(draft.sourceText), targetRanges: parseSparklineRangeList(draft.targetText), settings: sparklineSettingsInput(draft.settings) }
  if (draft.dataOperation === 'settings') return { kind: 'settings', ranges, settings }
  return { kind: 'change-source', ranges: [draft.primary], resetType: draft.dataOperation, primary: draft.primary,
    sourceRanges: parseSparklineRangeList(draft.sourceText), targetRanges: parseSparklineRangeList(draft.targetText), ...(Object.keys(settings).length ? { settings } : {}) }
}

export function canEditSparklinePanelRanges(owner: SparklinePanelOwner, ranges: readonly string[]): boolean {
  const sheet = currentSheet(owner), parsed: CellRange[] = []
  for (const range of ranges) {
    const reference = parseSparklineReference(range)
    if (!reference || reference.sheetName !== undefined && reference.sheetName !== sheet.id && reference.sheetName.toLocaleLowerCase() !== sheet.name.toLocaleLowerCase()) return false
    parsed.push(reference.range)
  }
  return parsed.length > 0 && canViewWorksheetProtection(sheet, owner.userId) && canEditWorksheetProtection(sheet, owner.userId)
    && canViewProtectedRanges(sheet, parsed, owner.userId) && canEditProtectedRanges(sheet, parsed, owner.userId)
}

/** Same full target and source footprint as the command, rechecked immediately before publication. */
export function assertSparklinePanelMutation(owner: SparklinePanelOwner, request: SparklineMutationRequest): void {
  sparklinePanelOwnerContext(owner)
  const workbook = owner.studio.getWorkbookModel(owner.target.workbookId)!, sheet = currentSheet(owner)
  const plan = planSparklineMutation(sheet, request, { sheets: workbook.sheets })
  if (!canViewWorksheetProtection(sheet, owner.userId) || !canEditWorksheetProtection(sheet, owner.userId)
    || !canViewProtectedRanges(sheet, plan.affectedRanges, owner.userId) || !canEditProtectedRanges(sheet, plan.affectedRanges, owner.userId)) {
    throw new Error('迷你图操作涉及受保护的目标或整组成员。')
  }
  for (const spec of Object.values(plan.patches)) {
    if (!spec) continue
    const reference = parseSparklineReference(spec.sourceRange)!
    const source = reference.sheetName === undefined ? sheet : workbook.sheets.find(candidate => candidate.id === reference.sheetName || candidate.name.toLocaleLowerCase() === reference.sheetName!.toLocaleLowerCase())
    if (!source || !canViewWorksheetProtection(source, owner.userId) || !canViewProtectedRanges(source, [reference.range], owner.userId)) throw new Error('迷你图源数据已关闭或当前用户无权查看。')
  }
  sparklinePanelOwnerContext(owner)
}
