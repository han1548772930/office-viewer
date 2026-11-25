import { defineStore } from 'pinia'
import { ref, shallowRef } from 'vue'
import { cloneModel } from '../../domain/serialization.ts'
import type { SheetClipboardLocation, SheetClipboardRecentPaste, SheetClipboardSnapshot } from './sheet-clipboard-types.ts'
import type { SheetImageObjectClipboard } from '../sheet-images/sheet-image-object-clipboard.ts'
import { assertSheetImageObjectClipboard } from '../sheet-images/sheet-image-object-clipboard.ts'

/** Copy marks and optional paste are transient; content history remains in the shared history store. */
export const useSheetClipboardStore = defineStore('sheet-clipboard', () => {
  const clipboard = shallowRef<SheetClipboardSnapshot>()
  const imageClipboard = shallowRef<SheetImageObjectClipboard>()
  const copyMark = shallowRef<SheetClipboardLocation>()
  const recentPaste = shallowRef<SheetClipboardRecentPaste>()
  const clipboardGeneration = ref(0)
  function setClipboard(snapshot: SheetClipboardSnapshot): void {
    clipboard.value = cloneModel(snapshot)
    imageClipboard.value = undefined
    copyMark.value = snapshot.source ? cloneModel({ unitId: snapshot.source.unitId, sheetId: snapshot.source.sheetId, range: snapshot.source.range }) : undefined
    recentPaste.value = undefined
    clipboardGeneration.value++
  }
  function setImageClipboard(snapshot: SheetImageObjectClipboard): void {
    assertSheetImageObjectClipboard(snapshot)
    imageClipboard.value = cloneModel(snapshot)
    clipboard.value = undefined; copyMark.value = undefined; recentPaste.value = undefined; clipboardGeneration.value++
  }
  function getImageClipboard(id?: string): SheetImageObjectClipboard | undefined {
    const current = imageClipboard.value
    return current && (id === undefined || current.id === id) ? cloneModel(current) : undefined
  }
  function clearImageClipboard(): void { if (imageClipboard.value) { imageClipboard.value = undefined; clipboardGeneration.value++ } }
  function getClipboard(id?: string): SheetClipboardSnapshot | undefined {
    const current = clipboard.value
    return current && (id === undefined || id === current.id) ? cloneModel(current) : undefined
  }
  function consumeCut(id: string): void {
    if (clipboard.value?.id === id && clipboard.value.copyType === 'cut') { clipboard.value = undefined; copyMark.value = undefined; clipboardGeneration.value++ }
  }
  function setRecentPaste(value?: SheetClipboardRecentPaste): void { recentPaste.value = value ? cloneModel(value) : undefined }
  function getRecentPaste(): SheetClipboardRecentPaste | undefined { return recentPaste.value ? cloneModel(recentPaste.value) : undefined }
  function cancelCopyMark(): void {
    copyMark.value = undefined
    if (clipboard.value?.copyType === 'cut') clipboard.value = undefined
    recentPaste.value = undefined
    imageClipboard.value = undefined
    clipboardGeneration.value++
  }
  function clearUnit(unitId: string): void {
    if (imageClipboard.value?.source.unitId === unitId) imageClipboard.value = undefined
    if (copyMark.value?.unitId === unitId) copyMark.value = undefined
    if (clipboard.value?.source?.unitId === unitId && clipboard.value.copyType === 'cut') clipboard.value = undefined
    if (recentPaste.value?.target.unitId === unitId || recentPaste.value?.clipboard.source?.unitId === unitId) recentPaste.value = undefined
    clipboardGeneration.value++
  }
  function clear(): void { clipboard.value = undefined; imageClipboard.value = undefined; copyMark.value = undefined; recentPaste.value = undefined; clipboardGeneration.value++ }
  return { clipboard, imageClipboard, copyMark, recentPaste, clipboardGeneration, setClipboard, getClipboard, setImageClipboard, getImageClipboard, clearImageClipboard,
    consumeCut, setRecentPaste, getRecentPaste, cancelCopyMark, clearUnit, clear }
})
