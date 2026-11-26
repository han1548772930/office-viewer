import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import type { SheetDrawingAffineGeometry, SheetDrawingRef, SheetDrawingShapeDTO, SheetDrawingState } from '../../domain/sheet-drawings.ts'
import { drawingAffineBounds, drawingAffineFromFrame, normalizeDrawingAffineGeometry } from '../../domain/drawing-affine.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../rendering/spreadsheet/layout.ts'
import { sheetClipboardObjectAxes, sheetClipboardObjectRange } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'
import { sheetImageDrawingGeometry, sheetImageFrame, sheetImageWithDrawingGeometry } from '../sheet-images/sheet-image-placement.ts'
import { resolveSlicerBounds, slicerDrawingGeometry, slicerWithDrawingGeometry } from '../slicer/slicer-drawing.ts'
import { isPrimaryChartObject } from '../charts/chart-object-commands.ts'
import { pruneChartImageResources } from '../charts/chart-image-resources.ts'
import { applyDrawingTransaction, drawingGraph, drawingRefKey, normalizeDrawingState, reconcileDrawingState, SheetDrawingError } from '../drawing-groups/index.ts'
import type { SheetDrawingApplyResult, SheetDrawingObject, SheetDrawingReconcileChanges, SheetDrawingScene, SheetDrawingTransaction } from '../drawing-groups/index.ts'
import { assertDrawingEditable, drawingDescendants, drawingSelection } from '../drawing-groups/drawing-group-model.ts'

type ShapeSheet = SpreadsheetSheet & { shapes?: SheetDrawingShapeDTO[] }
export interface SheetDrawingHostOptions {
  canEdit?: (ref: SheetDrawingRef, geometry: SheetDrawingAffineGeometry) => boolean
  available?: (ref: SheetDrawingRef) => boolean
}

function drawingSheet(workbook: WorkbookModel, sheetId: string): ShapeSheet {
  const sheet = workbook.sheets.find(sheet => sheet.id === sheetId)
  if (!sheet) throw new SheetDrawingError('not-found', '绘图所在工作表已移除。')
  return sheet
}

/** Images use anchor-projected geometry; legacy chart/slicer coordinates include the grid headers. */
export function createSheetDrawingScene(workbook: WorkbookModel, sheetId: string, options: SheetDrawingHostOptions = {}): SheetDrawingScene {
  const sheet = drawingSheet(workbook, sheetId), axes = sheetClipboardObjectAxes(workbook, sheet), objects: SheetDrawingObject[] = []
  const add = (object: SheetDrawingObject) => {
    if (options.canEdit) object.editable = options.canEdit(object.ref, object.geometry)
    if (options.available) object.available = object.available !== false && options.available(object.ref)
    objects.push(object)
  }
  for (const image of sheet.images ?? []) {
    const frame = sheetImageFrame(image, axes), collapsed = image.anchorType === '1' && (frame.width === 0 || frame.height === 0)
    // Both anchors may project to the same point while rows/columns are hidden. This geometry is only
    // a reversible state placeholder; rendering still uses the unchanged native image and its anchors.
    const geometry = collapsed && !image.drawingGeometry ? drawingAffineFromFrame({ ...frame, width: Math.max(1, frame.width), height: Math.max(1, frame.height) }, image.rotation)
      : sheetImageDrawingGeometry(image, axes)
    add({ ref: { kind: 'image', id: image.id }, sheetId, geometry, groupable: true, transformable: true,
      constrainMovement: image.anchorType !== '2',
      available: !collapsed && Boolean(image.resource && typeof image.resource.source === 'string' && image.resource.source.length),
    })
  }
  for (const chart of workbook.charts ?? []) if (chart.sheetId === sheetId) add({ ref: { kind: 'chart', id: chart.id }, sheetId, revision: chart.revision,
    geometry: chart.drawingGeometry ? normalizeDrawingAffineGeometry(chart.drawingGeometry) : drawingAffineFromFrame({
      left: chart.position.x - GRID_HEADER_WIDTH, top: chart.position.y - GRID_HEADER_HEIGHT, ...chart.size,
    }), groupable: true, transformable: true,
  })
  for (const shape of sheet.shapes ?? []) {
    if (shape.sheetId !== sheetId) throw new SheetDrawingError('invalid-input', '形状与宿主工作表不一致。', { kind: 'shape', id: shape.id })
    add({ ref: { kind: 'shape', id: shape.id }, sheetId, geometry: normalizeDrawingAffineGeometry(shape.geometry), groupable: true, transformable: true })
  }
  for (const slicer of workbook.slicers ?? []) if (slicer.hostSheetId === sheetId) {
    const bounds = resolveSlicerBounds(slicer, axes), collapsed = Number.isFinite(bounds.width) && bounds.width >= 0
      && Number.isFinite(bounds.height) && bounds.height >= 0 && (bounds.width === 0 || bounds.height === 0)
    add({ ref: { kind: slicer.type, id: slicer.id }, sheetId, revision: slicer.revision,
      geometry: collapsed && !slicer.drawingGeometry ? drawingAffineFromFrame({ ...bounds, left: bounds.left - GRID_HEADER_WIDTH, top: bounds.top - GRID_HEADER_HEIGHT,
        width: Math.max(1, bounds.width), height: Math.max(1, bounds.height) }) : slicerDrawingGeometry(slicer, axes),
      groupable: true, transformable: !slicer.lockedPosition, constrainMovement: slicer.drawing?.anchorType !== 'none',
      ...(collapsed ? { available: false } : {}),
    })
  }
  return { sheetId, objects, state: normalizeDrawingState({ sheetId, state: sheet.drawingState, objects }) }
}

/** Stage every leaf update before assigning anything to the caller's detached workbook. */
export function applySheetDrawingTransaction(workbook: WorkbookModel, sheetId: string, transaction: SheetDrawingTransaction,
  direction: 'forward' | 'backward' = 'forward', options: SheetDrawingHostOptions = {}): SheetDrawingApplyResult {
  const sheet = drawingSheet(workbook, sheetId), current = createSheetDrawingScene(workbook, sheetId, options)
  const applied = applyDrawingTransaction(current, transaction, { direction }), axes = sheetClipboardObjectAxes(workbook, sheet)
  const patches = new Map(transaction.geometryChanges.map(change => [drawingRefKey(change.ref), direction === 'backward' ? change.before : change.after]))
  const guards = new Map(transaction.expectedObjects.map(guard => [drawingRefKey(guard.ref), guard]))
  const images = (sheet.images ?? []).map(image => {
    const geometry = patches.get(drawingRefKey({ kind: 'image', id: image.id }))
    return geometry ? sheetImageWithDrawingGeometry(image, geometry, axes) : image
  })
  const charts = (workbook.charts ?? []).map(chart => {
    if (chart.sheetId !== sheetId) return chart
    const geometry = patches.get(drawingRefKey({ kind: 'chart', id: chart.id }))
    if (!geometry) return chart
    const originalRevision = guards.get(drawingRefKey({ kind: 'chart', id: chart.id }))?.revision
    if (direction === 'backward' && originalRevision !== undefined && chart.revision !== originalRevision + 1) {
      throw new SheetDrawingError('stale-object', '图表在变换后已经更改，不能覆盖其新内容。', { kind: 'chart', id: chart.id })
    }
    const bounds = drawingAffineBounds(geometry)
    return { ...cloneModel(chart), drawingGeometry: normalizeDrawingAffineGeometry(geometry),
      position: { x: GRID_HEADER_WIDTH + bounds.left, y: GRID_HEADER_HEIGHT + bounds.top },
      size: { ...chart.size }, revision: direction === 'backward' && originalRevision !== undefined ? originalRevision : chart.revision + 1,
    }
  })
  const shapes = (sheet.shapes ?? []).map(shape => {
    const geometry = patches.get(drawingRefKey({ kind: 'shape', id: shape.id }))
    return geometry ? { ...cloneModel(shape), geometry: normalizeDrawingAffineGeometry(geometry) } : shape
  })
  const slicers = (workbook.slicers ?? []).map(slicer => {
    if (slicer.hostSheetId !== sheetId) return slicer
    const ref = { kind: slicer.type, id: slicer.id }, geometry = patches.get(drawingRefKey(ref))
    if (!geometry) return slicer
    const originalRevision = guards.get(drawingRefKey(ref))?.revision
    if (direction === 'backward' && originalRevision !== undefined && slicer.revision !== originalRevision + 1) {
      throw new SheetDrawingError('stale-object', '切片器在变换后已经更改，不能覆盖其新内容。', ref)
    }
    const next = slicerWithDrawingGeometry(workbook, slicer, geometry)
    next.revision = direction === 'backward' && originalRevision !== undefined ? originalRevision : slicer.revision + 1
    return next
  })
  const writesImages = transaction.geometryChanges.some(change => change.ref.kind === 'image'),
    writesCharts = transaction.geometryChanges.some(change => change.ref.kind === 'chart'),
    writesShapes = transaction.geometryChanges.some(change => change.ref.kind === 'shape'),
    writesSlicers = transaction.geometryChanges.some(change => change.ref.kind === 'slicer' || change.ref.kind === 'timeline')
  const nextSheet = { ...sheet, ...(writesImages ? { images } : {}), ...(writesShapes ? { shapes } : {}), drawingState: applied.scene.state }
  const candidate = { ...workbook, ...(writesCharts ? { charts } : {}), ...(writesSlicers ? { slicers } : {}),
    sheets: workbook.sheets.map(candidate => candidate.id === sheetId ? nextSheet : candidate) }
  const nextScene = createSheetDrawingScene(candidate, sheetId, options)
  if (writesImages) sheet.images = images
  if (writesCharts) workbook.charts = charts
  if (writesShapes) sheet.shapes = shapes
  if (writesSlicers) workbook.slicers = slicers
  sheet.drawingState = nextScene.state
  return { scene: nextScene, selection: applied.selection }
}

/** Before/after affine AABBs include changed groups and all their children, then use the editor's actual axis indexes. */
export function sheetDrawingTransactionAffectedRanges(workbook: WorkbookModel, sheetId: string, transaction: SheetDrawingTransaction): CellRange[] {
  const before = createSheetDrawingScene(workbook, sheetId), after = applyDrawingTransaction(before, transaction).scene
  const ranges = new Map<string, CellRange>()
  for (const scene of [before, after]) {
    const graph = drawingGraph(scene)
    for (const ref of transaction.affectedRefs) {
      const geometry = ref.kind === 'group' ? graph.groups.get(ref.id)?.frame : graph.objects.get(drawingRefKey(ref))?.geometry
      if (!geometry) continue // A created/deleted group exists in only one side; leaf owners never disappear here.
      const bounds = drawingAffineBounds(geometry), range = sheetClipboardObjectRange(workbook, sheetId,
        { ...bounds, left: GRID_HEADER_WIDTH + bounds.left, top: GRID_HEADER_HEIGHT + bounds.top })
      ranges.set([range.startRow, range.startCol, range.endRow, range.endCol].join(':'), range)
    }
  }
  return [...ranges.values()]
}

/** Call after native leaf changes, passing explicit removed references and complete copy identity mappings. */
export function reconcileSheetDrawingState(workbook: WorkbookModel, sheetId: string, changes: SheetDrawingReconcileChanges = {}): SheetDrawingState {
  const sheet = drawingSheet(workbook, sheetId)
  // Scene construction normally checks all owners; explicit deletion must first inspect the previous structural state.
  const previous = sheet.drawingState
  const probe = { ...workbook, sheets: workbook.sheets.map(candidate => candidate.id === sheetId ? { ...candidate, drawingState: undefined } : candidate) }
  const scene = createSheetDrawingScene(probe, sheetId)
  const state = reconcileDrawingState({ ...scene, state: previous }, changes)
  sheet.drawingState = state
  return state
}

export function sheetDrawingRefAffectedRanges(workbook: WorkbookModel, sheetId: string, refs: readonly SheetDrawingRef[]): CellRange[] {
  const graph = drawingGraph(createSheetDrawingScene(workbook, sheetId)), ranges = new Map<string, CellRange>()
  for (const ref of refs) {
    const geometry = ref.kind === 'group' ? graph.groups.get(ref.id)?.frame : graph.objects.get(drawingRefKey(ref))?.geometry
    if (!geometry) continue // Explicitly deleted refs exist only in the before side of a removal.
    const bounds = drawingAffineBounds(geometry), range = sheetClipboardObjectRange(workbook, sheetId,
      { ...bounds, left: GRID_HEADER_WIDTH + bounds.left, top: GRID_HEADER_HEIGHT + bounds.top })
    ranges.set([range.startRow, range.startCol, range.endRow, range.endCol].join(':'), range)
  }
  return [...ranges.values()]
}

export interface SheetDrawingRemoveResult extends SheetDrawingApplyResult { removedRefs: SheetDrawingRef[]; affectedRefs: SheetDrawingRef[] }
/** Delete only selected drawing owners. Pivot/Table sources and media referenced by other owners stay intact. */
export function removeSheetDrawings(workbook: WorkbookModel, sheetId: string, selection: readonly SheetDrawingRef[],
  options: SheetDrawingHostOptions = {}): SheetDrawingRemoveResult {
  const sheet = drawingSheet(workbook, sheetId), graph = drawingGraph(createSheetDrawingScene(workbook, sheetId, options))
  const selected = drawingSelection(graph, selection), removedRefs = drawingDescendants(graph, selected), removed = new Set(removedRefs.map(drawingRefKey))
  if (!selected.length) return { scene: graph.scene, selection: [], removedRefs: [], affectedRefs: [] }
  assertDrawingEditable(graph, selected)
  for (const ref of removedRefs) if (ref.kind === 'chart') {
    const chart = workbook.charts?.find(chart => chart.sheetId === sheetId && chart.id === ref.id)
    if (chart && isPrimaryChartObject(workbook, chart)) throw new SheetDrawingError('read-only', '图表工作表的主图不能独立删除。', ref)
  }
  const affected = new Map(removedRefs.map(ref => [drawingRefKey(ref), cloneModel(ref)]))
  for (const ref of selected) {
    let parent = graph.parents.get(drawingRefKey(ref)) ?? null
    while (parent !== null) {
      const ancestor = { kind: 'group' as const, id: parent }; affected.set(drawingRefKey(ancestor), ancestor)
      parent = graph.parents.get(drawingRefKey(ancestor)) ?? null
    }
  }
  const nextSheet: ShapeSheet = { ...sheet, drawingState: graph.scene.state,
    ...(sheet.images === undefined ? {} : { images: sheet.images.filter(image => !removed.has(drawingRefKey({ kind: 'image', id: image.id }))) }),
    ...(sheet.shapes === undefined ? {} : { shapes: sheet.shapes.filter(shape => !removed.has(drawingRefKey({ kind: 'shape', id: shape.id }))) }),
  }
  const candidate: WorkbookModel = { ...workbook, sheets: workbook.sheets.map(candidate => candidate.id === sheetId ? nextSheet : candidate),
    ...(workbook.charts === undefined ? {} : { charts: workbook.charts.filter(chart => chart.sheetId !== sheetId || !removed.has(drawingRefKey({ kind: 'chart', id: chart.id }))) }),
    ...(workbook.slicers === undefined ? {} : { slicers: workbook.slicers.filter(slicer => slicer.hostSheetId !== sheetId || !removed.has(drawingRefKey({ kind: slicer.type, id: slicer.id }))) }),
  }
  const removesCharts = removedRefs.some(ref => ref.kind === 'chart')
  if (removesCharts) pruneChartImageResources(candidate)
  reconcileSheetDrawingState(candidate, sheetId, { removedRefs: selected })
  const scene = createSheetDrawingScene(candidate, sheetId, options)
  if (sheet.images !== undefined) sheet.images = nextSheet.images
  if (sheet.shapes !== undefined) sheet.shapes = nextSheet.shapes
  if (workbook.charts !== undefined) workbook.charts = candidate.charts
  if (workbook.slicers !== undefined) workbook.slicers = candidate.slicers
  if (removesCharts) workbook.metadata = candidate.metadata
  sheet.drawingState = scene.state
  return { scene, selection: [], removedRefs, affectedRefs: [...affected.values()] }
}
