import type { OfficeStudio } from './office-studio.ts'
interface Selection { sheetId: string; imageIds: readonly string[] }
const owners = new WeakMap<OfficeStudio, Map<string, () => Selection>>()
export function bindSheetImageSelection(api: OfficeStudio, unitId: string, selection: () => Selection) {
  const units = owners.get(api) ?? new Map<string, () => Selection>()
  owners.set(api, units); units.set(unitId, selection)
  return { dispose: () => { if (units.get(unitId) === selection) units.delete(unitId); if (!units.size && owners.get(api) === units) owners.delete(api) } }
}
export function getSheetImageSelection(api: OfficeStudio, unitId: string, sheetId: string): string[] {
  if (api.isDisposed) return []
  const selection = owners.get(api)?.get(unitId)?.()
  return selection?.sheetId === sheetId ? [...selection.imageIds] : []
}
export function clearSheetImageSelections(api: OfficeStudio, unitId?: string): void {
  if (unitId === undefined) owners.delete(api)
  else { const units = owners.get(api); units?.delete(unitId); if (!units?.size) owners.delete(api) }
}
