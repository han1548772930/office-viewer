import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import type { SheetDrawingRef } from '../domain/sheet-drawings.ts'
import { cloneDrawingRef, drawingRefKey, normalizeDrawingRef } from '../features/drawing-groups/index.ts'

export interface SheetDrawingSelection { sheetId: string; refs: readonly SheetDrawingRef[] }
interface SelectionOwner {
  read: () => SheetDrawingSelection
  write?: (selection: SheetDrawingSelection) => void
}
const owners = new WeakMap<OfficeStudio, Map<string, SelectionOwner>>()
const published = new WeakMap<OfficeStudio, Map<string, Map<string, SheetDrawingSelection>>>()
const generations = new WeakMap<OfficeStudio, Map<string, number>>()
const selectionKey = (refs: readonly SheetDrawingRef[]) => JSON.stringify(refs.map(drawingRefKey).sort())

function advanceSelectionGeneration(api: OfficeStudio, unitId: string): void {
  const units = generations.get(api) ?? new Map<string, number>()
  generations.set(api, units); units.set(unitId, (units.get(unitId) ?? 0) + 1)
}

/** A changed selection or replaced input owner separates two keyboard-move bursts. */
export function getSheetDrawingSelectionGeneration(api: OfficeStudio, unitId: string): number {
  return generations.get(api)?.get(unitId) ?? 0
}

function selectionRefs(refs: readonly SheetDrawingRef[]): SheetDrawingRef[] {
  if (!Array.isArray(refs)) throw new TypeError('绘图选择必须是数组。')
  const result = new Map<string, SheetDrawingRef>()
  for (const value of refs) { const ref = normalizeDrawingRef(value); result.set(drawingRefKey(ref), ref) }
  return [...result.values()]
}

export function bindSheetDrawingSelection(api: OfficeStudio, unitId: string, read: () => SheetDrawingSelection,
  write?: (selection: SheetDrawingSelection) => void): Disposable {
  const units = owners.get(api) ?? new Map<string, SelectionOwner>(), owner = { read, write }
  owners.set(api, units); units.set(unitId, owner)
  advanceSelectionGeneration(api, unitId)
  return { dispose: () => {
    if (units.get(unitId) === owner) { units.delete(unitId); advanceSelectionGeneration(api, unitId) }
    if (!units.size && owners.get(api) === units) owners.delete(api)
  } }
}

export function getSheetDrawingSelection(api: OfficeStudio, unitId: string, sheetId: string): SheetDrawingRef[] {
  if (api.isDisposed) return []
  const visible = api.activeWorkbookId === unitId && api.editor.state.activeKind === 'workbook'
  const owned = visible ? owners.get(api)?.get(unitId)?.read() : undefined
  const selection = owned?.sheetId === sheetId ? owned : published.get(api)?.get(unitId)?.get(sheetId)
  return selection?.sheetId === sheetId ? selectionRefs(selection.refs) : []
}

/** Distinguishes a deliberately empty API selection from a host that has never published one. */
export function hasSheetDrawingSelection(api: OfficeStudio, unitId: string, sheetId: string): boolean {
  return !api.isDisposed && published.get(api)?.get(unitId)?.has(sheetId) === true
}

/** Selection publication belongs to this API instance/workbook; an inactive sheet does not overwrite its visible owner. */
export function setSheetDrawingSelection(api: OfficeStudio, unitId: string, sheetId: string, refs: readonly SheetDrawingRef[]): void {
  if (api.isDisposed) return
  const selection = { sheetId, refs: selectionRefs(refs) }, units = published.get(api) ?? new Map<string, Map<string, SheetDrawingSelection>>()
  const sheets = units.get(unitId) ?? new Map<string, SheetDrawingSelection>()
  const previous = sheets.get(sheetId)
  if (!previous || selectionKey(previous.refs) !== selectionKey(selection.refs)) advanceSelectionGeneration(api, unitId)
  published.set(api, units); units.set(unitId, sheets); sheets.set(sheetId, selection)
  const owner = owners.get(api)?.get(unitId)
  if (api.activeWorkbookId === unitId && api.editor.state.activeKind === 'workbook' && owner?.read().sheetId === sheetId) {
    owner.write?.({ sheetId, refs: selection.refs.map(cloneDrawingRef) })
  }
}

export function clearSheetDrawingSelections(api: OfficeStudio, unitId?: string): void {
  if (unitId === undefined) {
    for (const id of generations.get(api)?.keys() ?? []) advanceSelectionGeneration(api, id)
    owners.delete(api); published.delete(api); return
  }
  advanceSelectionGeneration(api, unitId)
  for (const collection of [owners, published]) {
    const units = collection.get(api); units?.delete(unitId)
    if (!units?.size) collection.delete(api)
  }
}
