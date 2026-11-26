import type { OfficeStudio } from './office-studio.ts'
import type { ApiEventMap, Disposable } from './events.ts'
import type { CellRange, WorkbookModel } from '../domain/model.ts'
import { SHEET_DRAWING_ALIGN_TYPES, type SheetDrawingAlignType, type SheetDrawingAffineGeometry, type SheetDrawingKind, type SheetDrawingMatrix, type SheetDrawingMoveDirection, type SheetDrawingRef } from '../domain/sheet-drawings.ts'
import { drawingAffineFromFrame, multiplyDrawingAffine, normalizeDrawingAffineGeometry } from '../domain/drawing-affine.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../rendering/spreadsheet/layout.ts'
import { canViewProtectedRanges } from '../features/sheet-protection/range-protection.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { cloneDrawingRef, drawingGraph, drawingRefKey, normalizeDrawingRef, planDrawingAlign, planDrawingArrange, planDrawingGroup, planDrawingGroupTransform,
  planDrawingMove, planDrawingSelectionTransform, planDrawingUngroup, SHEET_DRAWING_GROUP_COMMAND_IDS, SheetDrawingError } from '../features/drawing-groups/index.ts'
import type { SheetDrawingArrangeType, SheetDrawingPlanResult, SheetDrawingScene, SheetDrawingTransaction } from '../features/drawing-groups/index.ts'
import { applySheetDrawingTransaction, createSheetDrawingScene, removeSheetDrawings, sheetDrawingRefAffectedRanges,
  sheetDrawingTransactionAffectedRanges } from '../features/sheet-drawings/sheet-drawing-host.ts'
import { clearSheetDrawingSelections, getSheetDrawingSelection, getSheetDrawingSelectionGeneration, setSheetDrawingSelection } from './sheet-drawing-selection.ts'
import { sheetDrawingMoveExtent } from '../features/sheet-drawings/sheet-drawing-move.ts'
import { SheetDrawingMoveHistory, type SheetDrawingMoveHistoryContext } from '../features/drawing-groups/drawing-move-history.ts'
import type { HistorySavepoint } from '../platform/history/history-types.ts'
import { assertChartObjectAction } from '../features/charts/chart-object-commands.ts'
import type { ChartObjectTarget } from '../features/charts/chart-object-commands.ts'

export type SheetDrawingCommandId = typeof SHEET_DRAWING_GROUP_COMMAND_IDS[keyof typeof SHEET_DRAWING_GROUP_COMMAND_IDS]
export interface SheetDrawingPublicRef { unitId?: string; subUnitId?: string; drawingId: string; drawingType?: number | SheetDrawingKind }
export interface SheetDrawingPublicGroupUpdate {
  parent: SheetDrawingPublicRef & { transform?: Record<string, unknown> }
  children: SheetDrawingPublicRef[]
}
export interface SheetDrawingCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  selection?: SheetDrawingRef[]
  expectedSelection?: SheetDrawingRef[]
  drawings?: SheetDrawingPublicRef[]
  drawingIds?: string[]
  imageIds?: string[]
  imageId?: string
  groupId?: string
  frame?: SheetDrawingAffineGeometry
  delta?: SheetDrawingMatrix
  arrangeType?: SheetDrawingArrangeType | 0 | 1 | 2 | 3
  alignType?: SheetDrawingAlignType
  direction?: SheetDrawingMoveDirection
  step?: 1 | 10
  zOrder?: number
  expectedRevision?: number
  expectedHistoryGeneration?: number
  host?: { unitId: string; sheetId: string }
  signal?: AbortSignal
}
export type SheetDrawingCommandParams = SheetDrawingCommandInput | readonly SheetDrawingPublicGroupUpdate[]

const kinds: Record<number, SheetDrawingKind> = { 0: 'image', 1: 'shape', 2: 'chart', 6: 'group', 10: 'slicer', 11: 'timeline' }
const labels: Record<SheetDrawingCommandId, string> = {
  [SHEET_DRAWING_GROUP_COMMAND_IDS.group]: '组合绘图', [SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup]: '取消组合',
  [SHEET_DRAWING_GROUP_COMMAND_IDS.transform]: '变换绘图', [SHEET_DRAWING_GROUP_COMMAND_IDS.arrange]: '排列绘图',
  [SHEET_DRAWING_GROUP_COMMAND_IDS.remove]: '删除绘图',
  [SHEET_DRAWING_GROUP_COMMAND_IDS.align]: '对齐绘图',
  [SHEET_DRAWING_GROUP_COMMAND_IDS.move]: '移动绘图',
}
interface SelectionSnapshot { sheetId: string; refs: SheetDrawingRef[] }
const selectionHistory = new WeakMap<OfficeStudio, Map<string, Map<string, SelectionSnapshot>>>()
const selectionKey = (generation: number, revision: number) => generation + ':' + revision
const refSet = (refs: readonly SheetDrawingRef[]) => JSON.stringify(refs.map(drawingRefKey).sort())
const moveHistories = new WeakMap<OfficeStudio, SheetDrawingMoveHistory>()
function moveHistoryFor(api: OfficeStudio): SheetDrawingMoveHistory {
  const history = moveHistories.get(api) ?? new SheetDrawingMoveHistory()
  moveHistories.set(api, history); return history
}

function valueRecord(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(label + '必须是对象。')
  return value as Record<string, unknown>
}
function id(value: unknown): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 1024 || /[\u0000-\u001f]/.test(value)) throw new TypeError('绘图标识无效。')
  return value
}
function refs(value: unknown): SheetDrawingRef[] {
  if (!Array.isArray(value)) throw new TypeError('绘图选择必须是数组。')
  const result = value.map(normalizeDrawingRef), keys = result.map(drawingRefKey)
  if (new Set(keys).size !== result.length) throw new TypeError('绘图选择不能包含重复对象。')
  return result
}
function publicRef(value: unknown): SheetDrawingPublicRef {
  const raw = valueRecord(value, '公开绘图引用'), result: SheetDrawingPublicRef = { drawingId: id(raw.drawingId) }
  if (raw.unitId !== undefined) result.unitId = id(raw.unitId)
  if (raw.subUnitId !== undefined) result.subUnitId = id(raw.subUnitId)
  if (raw.drawingType !== undefined) {
    const kind = typeof raw.drawingType === 'number' ? kinds[raw.drawingType] : raw.drawingType
    normalizeDrawingRef({ kind, id: result.drawingId }); result.drawingType = raw.drawingType as number | SheetDrawingKind
  }
  return result
}
export function sheetDrawingCommandInput(value: unknown): SheetDrawingCommandInput {
  if (value === undefined) return {}
  const raw = valueRecord(value, '绘图命令'), input: SheetDrawingCommandInput = {}
  for (const key of ['unitId', 'sheetId', 'subUnitId', 'imageId', 'groupId'] as const) if (raw[key] !== undefined) input[key] = id(raw[key])
  if (input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('绘图命令的工作表标识不一致。')
  for (const key of ['selection', 'expectedSelection'] as const) if (raw[key] !== undefined) input[key] = refs(raw[key])
  for (const key of ['drawingIds', 'imageIds'] as const) if (raw[key] !== undefined) {
    if (!Array.isArray(raw[key])) throw new TypeError('绘图标识批次必须是数组。')
    input[key] = (raw[key] as unknown[]).map(id)
  }
  if (raw.drawings !== undefined) {
    if (!Array.isArray(raw.drawings)) throw new TypeError('公开绘图批次必须是数组。')
    input.drawings = raw.drawings.map(publicRef)
    input.unitId ??= input.drawings[0]?.unitId
    input.subUnitId ??= input.drawings[0]?.subUnitId
    if (input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('绘图命令的工作表标识不一致。')
  }
  for (const key of ['expectedRevision', 'expectedHistoryGeneration'] as const) if (raw[key] !== undefined) {
    if (typeof raw[key] !== 'number' || !Number.isSafeInteger(raw[key]) || (raw[key] as number) < 0) throw new TypeError('绘图内容或历史版本无效。')
    input[key] = raw[key] as number
  }
  if (raw.frame !== undefined) input.frame = normalizeDrawingAffineGeometry(raw.frame)
  if (raw.delta !== undefined) input.delta = normalizeDrawingAffineGeometry({ width: 1, height: 1, matrix: raw.delta }).matrix
  if (raw.arrangeType !== undefined) {
    if (!['front', 'forward', 'back', 'backward', 0, 1, 2, 3].includes(raw.arrangeType as string | number)) throw new TypeError('绘图排列方式无效。')
    input.arrangeType = raw.arrangeType as SheetDrawingCommandInput['arrangeType']
  }
  if (raw.alignType !== undefined) {
    if (!(Object.values(SHEET_DRAWING_ALIGN_TYPES) as readonly unknown[]).includes(raw.alignType)) throw new TypeError('绘图对齐方式无效。')
    input.alignType = raw.alignType as SheetDrawingAlignType
  }
  if (raw.direction !== undefined) {
    if (raw.direction !== 0 && raw.direction !== 1 && raw.direction !== 2 && raw.direction !== 3) throw new TypeError('绘图移动方向无效。')
    input.direction = raw.direction
  }
  if (raw.step !== undefined) {
    if (raw.step !== 1 && raw.step !== 10) throw new TypeError('绘图移动步长无效。')
    input.step = raw.step
  }
  if (raw.zOrder !== undefined) { if (typeof raw.zOrder !== 'number') throw new TypeError('绘图层级必须是数值。'); input.zOrder = raw.zOrder }
  if (raw.host !== undefined) { const host = valueRecord(raw.host, '绘图宿主'); input.host = { unitId: id(host.unitId), sheetId: id(host.sheetId) } }
  if (raw.signal !== undefined) {
    if (!raw.signal || typeof raw.signal !== 'object' || typeof (raw.signal as AbortSignal).aborted !== 'boolean') throw new TypeError('绘图取消信号无效。')
    input.signal = raw.signal as AbortSignal
  }
  return input
}

function publicGeometry(value: Record<string, unknown>): SheetDrawingAffineGeometry {
  const raw = valueRecord(value, '公开组合变换')
  for (const key of ['flipX', 'flipY'] as const) if (raw[key] !== undefined && typeof raw[key] !== 'boolean') throw new TypeError('公开组合变换的 ' + key + ' 必须是布尔值。')
  if (raw.matrix !== undefined) return normalizeDrawingAffineGeometry(raw)
  const number = (key: string, fallback?: number) => {
    const result = raw[key] ?? fallback
    if (typeof result !== 'number' || !Number.isFinite(result)) throw new TypeError('公开组合变换的 ' + key + ' 无效。')
    return result
  }
  const width = number('width'), height = number('height'), left = number('left', 0) - GRID_HEADER_WIDTH, top = number('top', 0) - GRID_HEADER_HEIGHT
  const base = drawingAffineFromFrame({ left, top, width, height }, number('angle', 0))
  const scale = { a: number('scaleX', 1) * (raw.flipX ? -1 : 1), b: 0, c: 0, d: number('scaleY', 1) * (raw.flipY ? -1 : 1), e: 0, f: 0 }
  const skew = multiplyDrawingAffine({ a: 1, b: 0, c: Math.tan(number('skewX', 0) * Math.PI / 180), d: 1, e: 0, f: 0 },
    { a: 1, b: Math.tan(number('skewY', 0) * Math.PI / 180), c: 0, d: 1, e: 0, f: 0 })
  const matrix = multiplyDrawingAffine(base.matrix, multiplyDrawingAffine(scale, skew))
  matrix.e = left + width / 2 - matrix.a * width / 2 - matrix.c * height / 2
  matrix.f = top + height / 2 - matrix.b * width / 2 - matrix.d * height / 2
  return normalizeDrawingAffineGeometry({ width, height, matrix })
}

function resolvePublicRef(scene: SheetDrawingScene, drawing: SheetDrawingPublicRef): SheetDrawingRef {
  const graph = drawingGraph(scene), kind = typeof drawing.drawingType === 'number' ? kinds[drawing.drawingType] : drawing.drawingType
  if (kind !== undefined) {
    const ref = normalizeDrawingRef({ kind, id: drawing.drawingId })
    if (!graph.refs.has(drawingRefKey(ref))) throw new SheetDrawingError('not-found', '公开绘图引用不存在。', ref)
    return ref
  }
  const candidates = [...graph.refs.values()].filter(ref => ref.id === drawing.drawingId)
  if (candidates.length !== 1) throw new SheetDrawingError(candidates.length ? 'invalid-input' : 'not-found', '绘图 ID 不唯一或不存在，请提供 kind / drawingType。')
  return cloneDrawingRef(candidates[0]!)
}
function selectedRefs(api: OfficeStudio, unitId: string, scene: SheetDrawingScene, input: SheetDrawingCommandInput, command: SheetDrawingCommandId): SheetDrawingRef[] {
  if (input.selection !== undefined) return input.selection.map(cloneDrawingRef)
  if (input.drawings !== undefined) return input.drawings.map(drawing => resolvePublicRef(scene, drawing))
  if (input.drawingIds !== undefined) return input.drawingIds.map(drawingId => resolvePublicRef(scene, { drawingId }))
  if (input.imageIds !== undefined || input.imageId !== undefined) return (input.imageIds ?? [input.imageId!]).map(id => ({ kind: 'image', id }))
  if (input.groupId !== undefined && command !== SHEET_DRAWING_GROUP_COMMAND_IDS.group) return [{ kind: 'group', id: input.groupId }]
  return getSheetDrawingSelection(api, unitId, scene.sheetId)
}
function uniqueRanges(ranges: readonly CellRange[]): CellRange[] {
  return [...new Map(ranges.map(range => [[range.startRow, range.startCol, range.endRow, range.endCol].join(':'), range])).values()]
}
function pruneSelectionHistory(api: OfficeStudio, unitId: string): Map<string, SelectionSnapshot> | undefined {
  const snapshots = selectionHistory.get(api)?.get(unitId), history = api.editor.history.units[unitId]
  if (!snapshots) return undefined
  if (!history) { selectionHistory.get(api)?.delete(unitId); return undefined }
  const retained = new Set([selectionKey(history.generation, history.revision)])
  for (const entry of [...history.undo, ...history.redo]) {
    retained.add(selectionKey(history.generation, entry.before)); retained.add(selectionKey(history.generation, entry.after))
  }
  for (const key of snapshots.keys()) if (!retained.has(key)) snapshots.delete(key)
  return snapshots
}
function rememberSelection(api: OfficeStudio, unitId: string, before: SheetDrawingScene, after: SheetDrawingScene,
  beforeRefs: readonly SheetDrawingRef[], afterRefs: readonly SheetDrawingRef[], beforeHistoryRevision: number, committed: HistorySavepoint): void {
  const units = selectionHistory.get(api) ?? new Map<string, Map<string, SelectionSnapshot>>(), snapshots = units.get(unitId) ?? new Map<string, SelectionSnapshot>()
  selectionHistory.set(api, units); units.set(unitId, snapshots)
  snapshots.set(selectionKey(committed.generation, beforeHistoryRevision), { sheetId: before.sheetId, refs: beforeRefs.map(cloneDrawingRef) })
  snapshots.set(selectionKey(committed.generation, committed.revision), { sheetId: after.sheetId, refs: afterRefs.map(cloneDrawingRef) })
  // Match the real content history's retained revisions, including its 100-entry cap and discarded redo branches.
  pruneSelectionHistory(api, unitId)
}

/** Shared drawing edits, keyboard moves and timed history use the same guarded native-owner runner. */
export function runSheetDrawingCommand(api: OfficeStudio, command: SheetDrawingCommandId, params?: SheetDrawingCommandParams): boolean {
  if (!Object.values(SHEET_DRAWING_GROUP_COMMAND_IDS).includes(command)) throw new TypeError('未知绘图命令。')
  const moveHistory = moveHistoryFor(api)
  if (command !== SHEET_DRAWING_GROUP_COMMAND_IDS.move) moveHistory.clear()
  const publicBatch = Array.isArray(params) ? params.map(value => {
    const raw = valueRecord(value, '组合更新'), parent = publicRef(raw.parent)
    if (parent.drawingType !== undefined && parent.drawingType !== 'group' && parent.drawingType !== 6) throw new TypeError('组合父项的 drawingType 必须是 group。')
    if (!Array.isArray(raw.children)) throw new TypeError('组合子项必须是数组。')
    const parentRaw = valueRecord(raw.parent, '组合父项')
    return { parent, children: raw.children.map(publicRef), frame: parentRaw.transform === undefined ? undefined : publicGeometry(parentRaw.transform as Record<string, unknown>) }
  }) : undefined
  if (publicBatch && command !== SHEET_DRAWING_GROUP_COMMAND_IDS.group && command !== SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup) throw new TypeError('parent/children 批次仅适用于组合或取消组合。')
  if (publicBatch && !publicBatch.length) return false
  const firstUpdate = publicBatch?.[0]
  const input = publicBatch ? sheetDrawingCommandInput({ unitId: firstUpdate?.parent.unitId ?? firstUpdate?.children[0]?.unitId,
    subUnitId: firstUpdate?.parent.subUnitId ?? firstUpdate?.children[0]?.subUnitId }) : sheetDrawingCommandInput(params)
  if (input.frame !== undefined && input.delta !== undefined) throw new TypeError('需要且只能指定目标 frame 或世界 delta。')
  if (input.alignType !== undefined && command !== SHEET_DRAWING_GROUP_COMMAND_IDS.align) throw new TypeError('alignType 仅适用于对齐绘图。')
  if ((input.direction !== undefined || input.step !== undefined) && command !== SHEET_DRAWING_GROUP_COMMAND_IDS.move) throw new TypeError('direction / step 仅适用于移动绘图。')
  if ((input.frame !== undefined || input.delta !== undefined) && command !== SHEET_DRAWING_GROUP_COMMAND_IDS.group && command !== SHEET_DRAWING_GROUP_COMMAND_IDS.transform) {
    throw new TypeError('frame / delta 仅适用于组合或变换绘图。')
  }
  const target = resolveSheetRangeCommandTarget(api, input), revision = api.getWorkbookRevision(target.unitId), userId = api.currentUserId
  const historySavepoint = api.editor.history.captureSavepoint(target.unitId), historyGeneration = historySavepoint.generation
  const publicRefs = [...(input.drawings ?? []), ...(publicBatch ?? []).flatMap(update => [update.parent, ...update.children])]
  if (publicRefs.some(ref => ref.unitId !== undefined && ref.unitId !== target.unitId || ref.subUnitId !== undefined && ref.subUnitId !== target.sheetId)) {
    throw new TypeError('绘图操作需要同一工作簿及工作表。')
  }
  let permissionRanges: CellRange[] | undefined
  const chartTargets: ChartObjectTarget[] = []
  let chartAction: 'edit' | 'delete' = 'edit'
  const assertCurrent = () => {
    if (input.signal?.aborted) throw new DOMException('绘图操作已取消。', 'AbortError')
    if (api.isDisposed || api.getWorkbookModel(target.unitId) !== target.workbook || api.getWorkbookRevision(target.unitId) !== revision
      || api.editor.history.captureSavepoint(target.unitId).generation !== historyGeneration || api.currentUserId !== userId
      || input.expectedRevision !== undefined && input.expectedRevision !== revision
      || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== historyGeneration) throw new Error('绘图目标内容或历史已更改。')
    if (input.host && (input.host.unitId !== target.unitId || input.host.sheetId !== target.sheetId || api.activeWorkbookId !== target.unitId
      || api.editor.state.activeKind !== 'workbook' || target.workbook.activeSheetId !== target.sheetId)) throw new Error('绘图宿主已切换。')
    if (input.expectedSelection !== undefined && refSet(input.expectedSelection) !== refSet(getSheetDrawingSelection(api, target.unitId, target.sheetId))) throw new Error('绘图选择已更改。')
    api.assertCanViewWorksheet(target.unitId, target.sheetId); api.assertCanEditWorksheet(target.unitId, target.sheetId)
    if (permissionRanges) {
      if (!canViewProtectedRanges(target.sheet, permissionRanges, userId)) throw new Error('没有查看绘图完整占用范围的权限。')
      for (const range of permissionRanges) api.assertCanEditRange(target.unitId, target.sheetId, range)
    }
    for (const chart of chartTargets) assertChartObjectAction(api, chart, chartAction)
  }
  assertCurrent()
  const before = createSheetDrawingScene(target.workbook, target.sheetId), beforeGraph = drawingGraph(before)
  const newGroupIds = new Set(publicBatch?.map(update => update.parent.drawingId))
  const supplied = publicBatch ? command === SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup
    ? publicBatch.map(update => resolvePublicRef(before, { ...update.parent, drawingType: 'group' }))
    : publicBatch.flatMap(update => update.children).flatMap(child => {
      const kind = typeof child.drawingType === 'number' ? kinds[child.drawingType] : child.drawingType
      if (newGroupIds.has(child.drawingId) && (kind === undefined || kind === 'group') && !beforeGraph.refs.has(drawingRefKey({ kind: 'group', id: child.drawingId }))) return []
      return [resolvePublicRef(before, child)]
    }) : undefined
  const initialSelection = supplied ? [...new Map(supplied.map(ref => [drawingRefKey(ref), ref])).values()] : selectedRefs(api, target.unitId, before, input, command)
  const moveContext = (): SheetDrawingMoveHistoryContext => ({ unitId: target.unitId, sheetId: target.sheetId, userId,
    selectionKey: refSet(initialSelection), selectionGeneration: getSheetDrawingSelectionGeneration(api, target.unitId),
    contentRevision: api.getWorkbookRevision(target.unitId), savepoint: api.editor.history.captureSavepoint(target.unitId),
    savedRevision: api.editor.history.units[target.unitId]!.savedRevision, now: Date.now() })
  const amendment = command === SHEET_DRAWING_GROUP_COMMAND_IDS.move
    ? moveHistory.amendment(moveContext(), receipt => api.editor.history.canAmend(target.unitId, receipt)) : undefined
  if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.transform && input.groupId !== undefined && refSet(initialSelection) !== refSet([{ kind: 'group', id: input.groupId }])) {
    throw new TypeError('groupId 与实际绘图选择不一致。')
  }
  const candidate: WorkbookModel = { ...target.workbook, sheets: target.workbook.sheets.map(sheet => sheet.id === target.sheetId ? { ...sheet } : sheet) }
  const transactions: SheetDrawingTransaction[] = [], ranges: CellRange[] = []
  let selection = initialSelection, removed = false, removedRefs: SheetDrawingRef[] = []
  const apply = (result: SheetDrawingPlanResult) => {
    if (result.status === 'rejected') throw new SheetDrawingError(result.reason, result.message, result.ref)
    if (result.status === 'noop') return false
    ranges.push(...sheetDrawingTransactionAffectedRanges(candidate, target.sheetId, result.transaction))
    selection = applySheetDrawingTransaction(candidate, target.sheetId, result.transaction).selection
    transactions.push(result.transaction)
    return true
  }
  if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.group) {
    if (publicBatch) {
      const created: SheetDrawingRef[] = []
      for (const update of publicBatch) {
        const scene = createSheetDrawingScene(candidate, target.sheetId), members = update.children.map(child => resolvePublicRef(scene, child))
        if (!apply(planDrawingGroup(scene, { selection: members, groupId: update.parent.drawingId }))) continue
        created.push({ kind: 'group', id: update.parent.drawingId })
        if (update.frame) apply(planDrawingGroupTransform(createSheetDrawingScene(candidate, target.sheetId), { groupId: update.parent.drawingId, frame: update.frame }))
      }
      const graph = drawingGraph(createSheetDrawingScene(candidate, target.sheetId)), createdIds = new Set(created.map(ref => ref.id))
      selection = created.filter(ref => {
        let parent = graph.parents.get(drawingRefKey(ref)) ?? null
        while (parent !== null) {
          if (createdIds.has(parent)) return false
          parent = graph.parents.get(drawingRefKey({ kind: 'group', id: parent })) ?? null
        }
        return true
      })
    } else {
      const groupId = input.groupId ?? crypto.randomUUID()
      if (apply(planDrawingGroup(before, { selection, groupId })) && (input.frame !== undefined || input.delta !== undefined)) {
        const transform = input.frame === undefined ? { delta: input.delta! } : { frame: input.frame }
        apply(planDrawingGroupTransform(createSheetDrawingScene(candidate, target.sheetId), { groupId, ...transform }))
      }
    }
  } else if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup) {
    const groups = publicBatch ? publicBatch.map(update => resolvePublicRef(before, { ...update.parent, drawingType: 'group' })) : selection
    if (publicBatch) for (const update of publicBatch) {
      const group = beforeGraph.groups.get(update.parent.drawingId), supplied = update.children.map(child => resolvePublicRef(before, child))
      if (!group || refSet(group.members) !== refSet(supplied)) throw new SheetDrawingError('invalid-input', '取消组合的子项与当前模型不一致。')
    }
    apply(planDrawingUngroup(before, { selection: groups }))
  } else if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.arrange) {
    const arrangeType = typeof input.arrangeType === 'number' ? (['forward', 'backward', 'front', 'back'] as const)[input.arrangeType] : input.arrangeType
    if (input.zOrder !== undefined && arrangeType !== undefined || input.zOrder === undefined && arrangeType === undefined) throw new TypeError('需要且只能指定一个排列方式或层级。')
    apply(planDrawingArrange(before, input.zOrder === undefined ? { selection, arrangeType: arrangeType! } : { selection, zOrder: input.zOrder }))
  } else if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.align) {
    if (input.alignType === undefined) throw new TypeError('需要指定绘图对齐方式。')
    apply(planDrawingAlign(before, { selection, alignType: input.alignType }))
  } else if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.move) {
    if (input.direction === undefined) throw new TypeError('需要指定绘图移动方向。')
    apply(planDrawingMove(before, { selection, direction: input.direction, step: input.step, extent: sheetDrawingMoveExtent(target.workbook, target.sheetId) }))
  } else if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.transform) {
    if ((input.frame === undefined) === (input.delta === undefined)) throw new TypeError('需要且只能指定目标 frame 或世界 delta。')
    const transform = input.frame === undefined ? { delta: input.delta! } : { frame: input.frame }
    apply(input.groupId !== undefined ? planDrawingGroupTransform(before, { groupId: input.groupId, ...transform })
      : planDrawingSelectionTransform(before, { selection, ...transform }))
  } else if (selection.length) {
    const result = removeSheetDrawings(candidate, target.sheetId, selection)
    ranges.push(...sheetDrawingRefAffectedRanges(target.workbook, target.sheetId, result.affectedRefs), ...sheetDrawingRefAffectedRanges(candidate, target.sheetId, result.affectedRefs))
    selection = result.selection; removedRefs = result.removedRefs; removed = removedRefs.length > 0
  }
  if (!transactions.length && !removed) return false
  const affectedRanges = uniqueRanges(ranges)
  permissionRanges = affectedRanges
  const affected = [...new Map([...transactions.flatMap(transaction => transaction.affectedRefs), ...removedRefs].map(ref => [drawingRefKey(ref), ref])).values()]
  chartAction = removed ? 'delete' : 'edit'
  for (const ref of affected) if (ref.kind === 'chart') chartTargets.push({ unitId: target.unitId, sheetId: target.sheetId, chartId: ref.id,
    expectedRevision: revision, snapshot: JSON.stringify(target.workbook.charts?.find(chart => chart.sheetId === target.sheetId && chart.id === ref.id)) })
  assertCurrent()
  const removedSlicers = (target.workbook.slicers ?? []).filter(slicer => slicer.hostSheetId === target.sheetId && removedRefs.some(ref => ref.id === slicer.id && ref.kind === slicer.type))
  const changedSlicerKeys = new Set(transactions.flatMap(transaction => transaction.geometryChanges)
    .filter(change => change.ref.kind === 'slicer' || change.ref.kind === 'timeline').map(change => drawingRefKey(change.ref))),
    changedSlicers = (target.workbook.slicers ?? []).filter(slicer => slicer.hostSheetId === target.sheetId && changedSlicerKeys.has(drawingRefKey({ kind: slicer.type, id: slicer.id })))
  for (const slicer of [...removedSlicers, ...changedSlicers]) {
    const event: ApiEventMap['BeforeSlicerChange'] = { unitId: target.unitId, hostSheetId: target.sheetId, slicerId: slicer.id,
      commandId: removedSlicers.includes(slicer) ? 'sheet.command.remove-slicer' : command }
    api.events.emit('BeforeSlicerChange', event)
    if (event.cancel) throw new Error('切片器绘图修改已取消。')
    assertCurrent()
  }
  assertCurrent()
  const after = createSheetDrawingScene(candidate, target.sheetId)
  let committed: HistorySavepoint | undefined
  let slicerEvents: ApiEventMap['SlicerChanged'][] = []
  const changed = api.commitWorkbook(target.unitId, labels[command], model => {
    assertCurrent()
    if (JSON.stringify(createSheetDrawingScene(model, target.sheetId)) !== JSON.stringify(before)) throw new SheetDrawingError('stale-object', '绘图模型在发布前已更改。')
    if (removed) removeSheetDrawings(model, target.sheetId, initialSelection)
    else for (const transaction of transactions) applySheetDrawingTransaction(model, target.sheetId, transaction)
  }, { preserveStyleInheritance: true, recalculate: false, expectedRevision: revision, expectedHistoryGeneration: historyGeneration,
    ...(amendment ? { amend: amendment } : {}), onHistoryCommitted: receipt => { committed = { ...receipt } },
    prepareView: model => {
      assertCurrent(); toExcelWorkbookView(model)
      slicerEvents = (model.slicers ?? []).filter(slicer => slicer.hostSheetId === target.sheetId
        && changedSlicerKeys.has(drawingRefKey({ kind: slicer.type, id: slicer.id })))
        .map(slicer => ({ unitId: target.unitId, hostSheetId: target.sheetId, slicerId: slicer.id, revision: slicer.revision }))
    } })
  const publishedCurrent = () => !!committed && !api.isDisposed && api.getWorkbookRevision(target.unitId) === revision + 1
    && !input.signal?.aborted && api.currentUserId === userId
    && (!input.host || api.activeWorkbookId === target.unitId && api.editor.state.activeKind === 'workbook'
      && api.getWorkbookModel(target.unitId)?.activeSheetId === target.sheetId)
    && api.editor.history.captureSavepoint(target.unitId).generation === committed.generation
    && api.editor.history.captureSavepoint(target.unitId).revision === committed.revision
  if (changed && committed) {
    rememberSelection(api, target.unitId, before, after, initialSelection, selection, historySavepoint.revision, committed)
    if (publishedCurrent()) setSheetDrawingSelection(api, target.unitId, target.sheetId, selection)
    if (command === SHEET_DRAWING_GROUP_COMMAND_IDS.move && publishedCurrent()
      && refSet(getSheetDrawingSelection(api, target.unitId, target.sheetId)) === refSet(selection)) moveHistory.remember(moveContext())
    for (const slicer of removedSlicers) api.publishObserverEvent('SlicerRemoved', () => ({ unitId: target.unitId, hostSheetId: target.sheetId, slicerId: slicer.id }))
    for (const event of slicerEvents) api.publishObserverEvent('SlicerChanged', () => ({ ...event }))
  }
  return changed
}

/** Arrangement is already registered by the image host; it delegates to runSheetDrawingCommand. */
export function registerSheetDrawingCommands(api: OfficeStudio): Disposable {
  const registrations = [SHEET_DRAWING_GROUP_COMMAND_IDS.group, SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup,
    SHEET_DRAWING_GROUP_COMMAND_IDS.transform, SHEET_DRAWING_GROUP_COMMAND_IDS.remove, SHEET_DRAWING_GROUP_COMMAND_IDS.align, SHEET_DRAWING_GROUP_COMMAND_IDS.move]
    .map(command => api.registerCommand<SheetDrawingCommandParams | undefined, boolean>(command, params => {
      try { return runSheetDrawingCommand(api, command, params) }
      catch (error) { moveHistories.get(api)?.clear(); throw error }
    }))
  const restore = (event: { unitId: string; kind: 'workbook' | 'document'; label?: string }) => {
    if (event.kind !== 'workbook' || !Object.values(labels).includes(event.label ?? '')) return
    const workbook = api.getWorkbookModel(event.unitId)
    if (!workbook) return
    const savepoint = api.editor.history.captureSavepoint(event.unitId), snapshot = pruneSelectionHistory(api, event.unitId)?.get(selectionKey(savepoint.generation, savepoint.revision))
    if (snapshot && workbook.sheets.some(sheet => sheet.id === snapshot.sheetId)) {
      const graph = drawingGraph(createSheetDrawingScene(workbook, snapshot.sheetId))
      setSheetDrawingSelection(api, event.unitId, snapshot.sheetId, snapshot.refs.filter(ref => graph.refs.has(drawingRefKey(ref))))
    }
  }
  const undo = api.events.addEvent('Undo', restore), redo = api.events.addEvent('Redo', restore)
  const commands = api.events.addEvent('CommandExecuted', event => { if (event.id !== SHEET_DRAWING_GROUP_COMMAND_IDS.move) moveHistories.get(api)?.clear() })
  const beforeUndo = api.events.addEvent('BeforeUndo', () => moveHistories.get(api)?.clear())
  const beforeRedo = api.events.addEvent('BeforeRedo', () => moveHistories.get(api)?.clear())
  const disposed = api.events.addEvent('WorkbookDisposed', event => {
    clearSheetDrawingSelections(api, event.workbookId); selectionHistory.get(api)?.delete(event.workbookId)
    moveHistories.get(api)?.clear()
  })
  return { dispose: () => {
    registrations.forEach(registration => registration.dispose()); undo.dispose(); redo.dispose(); commands.dispose(); beforeUndo.dispose(); beforeRedo.dispose(); disposed.dispose()
    moveHistories.delete(api)
    clearSheetDrawingSelections(api); selectionHistory.delete(api)
  } }
}
