import { computed, onActivated, onDeactivated, onUnmounted, ref, shallowRef, watch } from 'vue'
import type { SheetFloatingImage } from '../../domain/sheet-images.ts'
import type { SheetDrawingAffineGeometry } from '../../domain/sheet-drawings.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { drawingAffinePoint } from '../../domain/drawing-affine.ts'
import type { GridViewport, SpreadsheetGeometry, SpreadsheetRenderSheet } from '../../rendering/spreadsheet/layout.ts'
import { hitSheetFloatingImage, moveSheetImage, rotateSheetImage, sheetImageDrawCommands, sheetImageGeometryRange, sheetImagePaneLayouts, visibleSheetImageResources,
  type SheetImageInteractionHandle } from '../../features/sheet-images/sheet-image-layout.ts'
import { SheetImageRenderCache } from '../../features/sheet-images/sheet-image-render-cache.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { cropSheetImageToRatio, SHEET_IMAGE_CROP_PRESETS } from '../../features/sheet-images/sheet-image-crop.ts'
import { dragSheetImageCrop, sheetImageUncroppedFrame, sheetImageUncroppedDrawingGeometry, SHEET_IMAGE_CROP_OPERATIONS,
  type SheetImageCropOperation, type SheetImageCropResult } from '../../features/sheet-images/sheet-image-crop-session.ts'
import { sheetImageFrame, sheetImageWithFrame, sheetImageWithDrawingGeometry } from '../../features/sheet-images/sheet-image-placement.ts'
import type { SheetImageCommandInput } from '../../features/sheet-images/sheet-image-commands.ts'
import { sheetImageSelectionBounds, transformSheetImageSelection } from '../../features/sheet-images/sheet-image-selection.ts'
import type { SheetDrawingMoveDirection } from '../../domain/sheet-drawings.ts'
import { sheetDrawingMoveDirectionForKey } from '../../features/drawing-groups/drawing-group-move.ts'

interface ImageCanvasContext {
  geometry(): SpreadsheetGeometry
  viewport(): GridViewport
  sheet(): SpreadsheetRenderSheet
  selectedId(): string | undefined
  selectedIds?(): readonly string[]
  drawingOrder?(): readonly string[]
  drawingPreview?(id: string): SheetDrawingAffineGeometry | undefined
  selectionControls?(): boolean
  hitAllowed?(id: string, position: { x: number; y: number }): boolean
  revision(): number | undefined
  userId(): string
  editing(): boolean
  composing(): boolean
  painter(): boolean
  color(): string
  background(): string
  select(id?: string, additive?: boolean, preserve?: boolean): void
  edit(id: string, row?: number, column?: number): void
  update(id: string, image: SheetFloatingImage, expectedRevision?: number): void
  updateBatch(images: SheetFloatingImage[], expectedRevision?: number): void
  nudge?(ids: readonly string[], direction: SheetDrawingMoveDirection, step: 1 | 10, expectedRevision?: number): boolean
  remove(id: string): void
  removeMany(ids: string[]): void
  crop(id: SheetImageCropOperation, input: SheetImageCommandInput): void
  render(): void
}
interface ImageGesture {
  image: SheetFloatingImage
  preview: SheetFloatingImage
  pointerId: number
  start: { x: number; y: number }
  handle?: SheetImageInteractionHandle
  center: { x: number; y: number }
  sheetId: string
  sheet: SpreadsheetRenderSheet
  revision?: number
  userId: string
  signature: string
  zoom: number
  viewport: GridViewport
  geometry: SpreadsheetGeometry
  cropping?: boolean
  images?: SheetFloatingImage[]
  previews?: SheetFloatingImage[]
  signatures?: Map<string, string>
}
interface ImageCropSession {
  original: SheetFloatingImage; baseline: SheetFloatingImage; preview: SheetFloatingImage; signature: string
  sheet: SpreadsheetRenderSheet; revision?: number; userId: string; changed: boolean
}
export function useSheetImageCanvas(context: ImageCanvasContext) {
  const imageError = ref(''), imageVersion = ref(0), gesture = shallowRef<ImageGesture>(), cropSession = shallowRef<ImageCropSession>()
  let active = true, disposed = false
  const cache = new SheetImageRenderCache(() => { if (active && !disposed) { imageVersion.value++; context.render() } },
    message => { if (active && !disposed) imageError.value = message })
  const visible = (row: number, column: number) => canViewWorksheetProtection(context.sheet(), context.userId()) && canViewProtectedRanges(context.sheet(),
    [{ startRow: row, endRow: row, startCol: column, endCol: column }], context.userId())
  const floatingVisible = (image: SheetFloatingImage) => canViewWorksheetProtection(context.sheet(), context.userId())
    && canViewProtectedRanges(context.sheet(), [sheetImageGeometryRange(context.geometry(), image)], context.userId())
  const selectedIds = computed(() => context.selectedIds?.() ?? (context.selectedId() ? [context.selectedId()!] : []))
  const selectedImages = computed(() => selectedIds.value.flatMap(id => {
    const image = context.sheet().images?.find(image => image.id === id)
    return image && floatingVisible(image) ? [image] : []
  }))
  const resources = computed(() => [...visibleSheetImageResources(context.geometry(), context.viewport(), visible, canViewWorksheetProtection(context.sheet(), context.userId()), floatingVisible),
    ...(cropSession.value && floatingVisible(cropSession.value.original) ? [cropSession.value.preview.resource] : [])])
  const backgroundConfig = computed(() => canViewWorksheetProtection(context.sheet(), context.userId()) ? context.sheet().backgroundImage : undefined)
  watch([resources, backgroundConfig], ([value, background]) => { if (active && !disposed) { imageError.value = ''; cache.sync(value, background) } }, { immediate: true })
  const layouts = computed(() => {
    const order = context.drawingOrder?.(), ranks = new Map(order?.map((id, index) => [id, index]) ?? [])
    const images = (context.sheet().images ?? []).filter(floatingVisible).map(image => {
      const external = context.drawingPreview?.(image.id)
      return gesture.value?.previews?.find(preview => preview.id === image.id) ?? (gesture.value?.image.id === image.id ? gesture.value.preview
        : cropSession.value?.original.id === image.id ? cropSession.value.preview : external ? sheetImageWithDrawingGeometry(image, external, context.geometry()) : image)
    })
    if (order) images.sort((left, right) => (ranks.get(left.id) ?? order.length) - (ranks.get(right.id) ?? order.length))
    return sheetImagePaneLayouts(context.geometry(), context.viewport(), images)
  })
  const selectionLayouts = computed(() => {
    if (selectedImages.value.length < 2) return []
    const images = selectedImages.value.map(image => gesture.value?.previews?.find(preview => preview.id === image.id) ?? image),
      bounds = sheetImageSelectionBounds(images, context.geometry()), selection: SheetFloatingImage = { ...images[0]!,
        id: '__sheet_image_selection__', anchorType: '2', absolutePosition: { left: bounds.left, top: bounds.top },
        size: { width: bounds.width, height: bounds.height }, crop: undefined, drawingGeometry: undefined, rotation: 0, flipX: false, flipY: false }
    return sheetImagePaneLayouts(context.geometry(), context.viewport(), [selection])
  })
  const sceneOptions = computed(() => {
    void imageVersion.value
    const viewport = context.viewport()
    const crop = cropSession.value && isCropCurrent(cropSession.value) ? cropSession.value : undefined, controls = context.selectionControls?.() !== false,
      multiple = controls && selectedImages.value.length > 1,
      imageCommands = sheetImageDrawCommands(layouts.value, resource => cache.get(resource), crop || !controls ? undefined : context.selectedId(), context.color(), context.background(),
        { selectedIds: multiple ? selectedIds.value : undefined, outlineOnly: multiple })
    if (multiple && !crop) imageCommands.push(...sheetImageDrawCommands(selectionLayouts.value, () => undefined, '__sheet_image_selection__', context.color(), context.background()))
    if (crop) {
      const frame = sheetImageUncroppedFrame(crop.preview, context.geometry()),
        sourceImage: SheetFloatingImage = { ...crop.preview, anchorType: '2', absolutePosition: { left: frame.left, top: frame.top },
          size: { width: frame.width, height: frame.height }, crop: undefined,
          drawingGeometry: sheetImageUncroppedDrawingGeometry(crop.preview, context.geometry()), opacity: 0.28 },
        sourceLayouts = sheetImagePaneLayouts(context.geometry(), viewport, [sourceImage]),
        cropLayouts = layouts.value.filter(layout => layout.image.id === crop.original.id)
      imageCommands.push(...sheetImageDrawCommands(sourceLayouts, resource => cache.get(resource)),
        ...sheetImageDrawCommands(cropLayouts, resource => cache.get(resource), crop.original.id, context.color(), context.background(), { crop: true }))
    }
    return { imageSource: (resource: Parameters<SheetImageRenderCache['get']>[0]) => cache.get(resource),
      imageVisible: visible,
      backgroundImageSource: cache.backgroundSource(context.geometry(), viewport, canViewWorksheetProtection(context.sheet(), context.userId()) ? context.sheet().backgroundImage : undefined),
      backgroundImageVersion: cache.backgroundVersion,
      imageCommands }
  })
  const menuPosition = computed(() => {
    if (context.editing() || context.composing() || context.painter() || context.selectionControls?.() === false) return undefined
    const id = context.selectedId(), viewport = context.viewport(), layout = selectionLayouts.value[0] ?? layouts.value.find(item => item.image.id === id)
    if (!layout) return undefined
    return { left: Math.max(44, Math.min(viewport.width - 300, Math.max(layout.bounds.x, layout.clip.x))),
      top: Math.max(26, Math.min(viewport.height - 36, Math.max(layout.bounds.y, layout.clip.y) - 34)) }
  })
  const selectedEditable = computed(() => {
    const images = selectedImages.value
    return images.length > 0 && images.length === selectedIds.value.length && canEditWorksheetProtection(context.sheet(), context.userId())
      && images.every(image => floatingVisible(image) && canEditProtectedRanges(context.sheet(), [sheetImageGeometryRange(context.geometry(), image)], context.userId()))
  })
  function isCurrent(state: ImageGesture): boolean {
    const current = context.sheet().images?.find(image => image.id === state.image.id)
    const viewport = context.viewport()
    return active && !disposed && !context.editing() && !context.composing() && !context.painter()
      && context.sheet() === state.sheet && context.sheet().id === state.sheetId && context.revision() === state.revision
      && context.userId() === state.userId && context.geometry() === state.geometry
      && (['zoom', 'scrollLeft', 'scrollTop', 'width', 'height'] as const).every(key => viewport[key] === state.viewport[key])
      && JSON.stringify(current) === state.signature
      && (!state.images || state.images.every(image => JSON.stringify(context.sheet().images?.find(current => current.id === image.id)) === state.signatures?.get(image.id))
        && state.images.length === selectedIds.value.length && state.images.every(image => selectedIds.value.includes(image.id)))
      && selectedEditable.value
  }
  function cancel(): void {
    const state = gesture.value
    if (state?.cropping && cropSession.value) cropSession.value = { ...cropSession.value, preview: state.image,
      changed: JSON.stringify(state.image) !== JSON.stringify(cropSession.value.baseline) }
    gesture.value = undefined; context.render()
  }
  function cancelCrop(): void { cropSession.value = undefined; if (gesture.value?.cropping) gesture.value = undefined; context.render() }
  function isCropCurrent(state: ImageCropSession): boolean {
    const image = context.sheet().images?.find(image => image.id === state.original.id)
    return active && !disposed && !context.editing() && !context.composing() && !context.painter()
      && context.sheet() === state.sheet && context.revision() === state.revision && context.userId() === state.userId
      && selectedIds.value.length <= 1 && JSON.stringify(image) === state.signature && floatingVisible(state.original)
      && canEditWorksheetProtection(context.sheet(), context.userId()) && canEditProtectedRanges(context.sheet(), [sheetImageGeometryRange(context.geometry(), state.original)], context.userId())
  }
  function operateImageCrop(id: SheetImageCropOperation, input: SheetImageCommandInput = {}): SheetImageCropResult {
    if (input.sheetId !== undefined && input.sheetId !== context.sheet().id || input.subUnitId !== undefined && input.subUnitId !== context.sheet().id) return { handled: false }
    if (id === SHEET_IMAGE_CROP_OPERATIONS.close) {
      const state = cropSession.value
      if (!state) return { handled: false }
      if (input.isCancel) { cancelCrop(); return { handled: true } }
      const current = isCropCurrent(state); cancelCrop()
      return current ? { handled: true, image: state.changed ? cloneModel(state.preview) : undefined, expectedRevision: state.revision } : { handled: false }
    }
    if (id !== SHEET_IMAGE_CROP_OPERATIONS.open && id !== SHEET_IMAGE_CROP_OPERATIONS.auto) return { handled: false }
    const imageId = input.imageId ?? input.drawingId ?? context.selectedId(), original = context.sheet().images?.find(image => image.id === imageId)
    if (!original || selectedIds.value.length > 1 || !active || disposed || context.editing() || context.composing() || context.painter()) return { handled: false }
    const range = sheetImageGeometryRange(context.geometry(), original)
    if (!floatingVisible(original) || !canEditWorksheetProtection(context.sheet(), context.userId()) || !canEditProtectedRanges(context.sheet(), [range], context.userId())
      || input.expectedRevision !== undefined && input.expectedRevision !== context.revision()) return { handled: false }
    const existing = cropSession.value
    if (existing && (!isCropCurrent(existing) || existing.original.id !== original.id)) cancelCrop()
    const baseline = sheetImageWithFrame(original, sheetImageFrame(original, context.geometry()), context.geometry()),
      state: ImageCropSession = cropSession.value ?? { original: cloneModel(original), baseline, preview: baseline, signature: JSON.stringify(original),
        sheet: context.sheet(), revision: context.revision(), userId: context.userId(), changed: false }
    if (id === SHEET_IMAGE_CROP_OPERATIONS.auto) {
      const preset = SHEET_IMAGE_CROP_PRESETS.find(preset => preset.value === input.cropType)
      if (!preset) return { handled: false }
      state.preview = preset.ratio ? cropSheetImageToRatio(state.preview, context.geometry(), preset.ratio) : state.preview
      state.changed = JSON.stringify(state.preview) !== JSON.stringify(state.baseline)
    }
    cropSession.value = { ...state }; context.select(original.id); context.render()
    return { handled: true }
  }
  function begin(position: { x: number; y: number }, event: PointerEvent): boolean {
    if (!active || disposed || context.editing() || context.composing() || context.painter() || event.button !== 0 && event.button !== 2) return false
    const cropping = cropSession.value
    if (cropping && !isCropCurrent(cropping)) cancelCrop()
    const selectionHit = !cropSession.value && selectedImages.value.length > 1 ? hitSheetFloatingImage(selectionLayouts.value, position.x, position.y, '__sheet_image_selection__') : undefined
    const hit = selectionHit?.handle ? selectionHit : hitSheetFloatingImage(cropSession.value ? layouts.value.filter(layout => layout.image.id === cropSession.value!.original.id) : layouts.value,
      position.x, position.y, selectedImages.value.length > 1 && !cropSession.value ? undefined : context.selectedId(), !cropSession.value)
    if (cropSession.value && !hit) {
      context.crop(SHEET_IMAGE_CROP_OPERATIONS.close, {}); event.preventDefault(); event.stopPropagation(); return true
    }
    if (!hit) { if (context.selectedId()) context.select(); return false }
    const selectedControl = hit.layout.image.id === '__sheet_image_selection__', picked = selectedControl
      ? selectedImages.value.find(image => image.id === context.selectedId()) ?? selectedImages.value[0]! : hit.layout.image,
      range = sheetImageGeometryRange(context.geometry(), picked)
    if (!hit.handle && !cropSession.value && context.hitAllowed?.(picked.id, position) === false) return false
    if (!canViewProtectedRanges(context.sheet(), [range], context.userId())) return true
    const additive = !cropSession.value && (event.shiftKey || event.ctrlKey || event.metaKey)
    context.select(picked.id, additive, !additive && selectedIds.value.includes(picked.id))
    event.preventDefault(); event.stopPropagation()
    if (additive) return true
    if (event.button !== 0 || !canEditWorksheetProtection(context.sheet(), context.userId()) || !canEditProtectedRanges(context.sheet(), [range], context.userId())) return true
    const image = cloneModel(picked), images = !cropSession.value && selectedIds.value.length > 1 && selectedIds.value.includes(picked.id)
      ? selectedImages.value.map(cloneModel) : undefined
    if (images && (!selectedEditable.value || images.length !== selectedIds.value.length)) return true
    gesture.value = { image, preview: image, pointerId: event.pointerId, start: position, handle: hit.handle,
      center: hit.layout.matrix ? drawingAffinePoint(hit.layout.matrix, hit.layout.rect.x + hit.layout.rect.width / 2, hit.layout.rect.y + hit.layout.rect.height / 2)
        : { x: hit.layout.rect.x + hit.layout.rect.width / 2, y: hit.layout.rect.y + hit.layout.rect.height / 2 },
      sheetId: context.sheet().id, sheet: context.sheet(), revision: context.revision(), userId: context.userId(),
      signature: cropSession.value?.signature ?? JSON.stringify(picked), zoom: context.viewport().zoom,
      viewport: { ...context.viewport() }, geometry: context.geometry(), cropping: Boolean(cropSession.value), images,
      signatures: images ? new Map(images.map(image => [image.id, JSON.stringify(image)])) : undefined }
    return true
  }
  function preview(state: ImageGesture, position: { x: number; y: number }, event: PointerEvent): SheetFloatingImage {
    if (state.cropping) return dragSheetImageCrop(state.image, context.geometry(), (position.x - state.start.x) / state.zoom,
      (position.y - state.start.y) / state.zoom, state.handle === 'rotate' ? undefined : state.handle)
    if (state.handle === 'rotate') {
      const start = Math.atan2(state.start.y - state.center.y, state.start.x - state.center.x),
        end = Math.atan2(position.y - state.center.y, position.x - state.center.x), delta = ((end - start) * 180 / Math.PI + 540) % 360 - 180
      return rotateSheetImage(context.geometry(), state.image, delta, event.shiftKey)
    }
    return moveSheetImage(context.geometry(), state.image, (position.x - state.start.x) / state.zoom, (position.y - state.start.y) / state.zoom, state.handle)
  }
  function move(position: { x: number; y: number }, event: PointerEvent): boolean {
    const state = gesture.value
    if (!state || state.pointerId !== event.pointerId) return false
    if (!isCurrent(state)) { cancel(); return true }
    const previews = selectionPreview(state, position, event), next = previews?.find(image => image.id === state.image.id) ?? preview(state, position, event)
    gesture.value = { ...state, preview: next, previews }
    if (state.cropping && cropSession.value) cropSession.value = { ...cropSession.value, preview: gesture.value.preview,
      changed: JSON.stringify(gesture.value.preview) !== JSON.stringify(cropSession.value.baseline) }
    event.preventDefault(); context.render(); return true
  }
  function end(position: { x: number; y: number }, event: PointerEvent): boolean {
    const state = gesture.value
    if (!state || state.pointerId !== event.pointerId) return false
    const current = isCurrent(state)
    gesture.value = undefined
    if (event.type !== 'pointerup' && state.cropping && cropSession.value) cropSession.value = { ...cropSession.value, preview: state.image,
      changed: JSON.stringify(state.image) !== JSON.stringify(cropSession.value.baseline) }
    if (event.type === 'pointerup' && current && Math.hypot(position.x - state.start.x, position.y - state.start.y) >= 2) {
      const previews = selectionPreview(state, position, event)
      const next = previews?.find(image => image.id === state.image.id) ?? preview(state, position, event)
      if (state.cropping && cropSession.value) cropSession.value = { ...cropSession.value, preview: next,
        changed: JSON.stringify(next) !== JSON.stringify(cropSession.value.baseline) }
      else if (previews) { if (JSON.stringify(previews) !== JSON.stringify(state.images)) context.updateBatch(previews, state.revision) }
      else if (JSON.stringify(next) !== JSON.stringify(state.image)) context.update(state.image.id, next, state.revision)
    }
    context.render(); return true
  }
  function editAt(position: { x: number; y: number }, floatingOnly = false): boolean {
    if (!active || disposed || context.editing() || context.composing() || context.painter()) return false
    if (cropSession.value) { context.crop(SHEET_IMAGE_CROP_OPERATIONS.close, {}); return true }
    const hit = hitSheetFloatingImage(layouts.value, position.x, position.y, context.selectedId())
    if (hit && context.hitAllowed?.(hit.layout.image.id, position) !== false) { context.select(hit.layout.image.id); context.edit(hit.layout.image.id); return true }
    const cell = !floatingOnly && context.geometry().hitTest(position.x, position.y, context.viewport())
    if (cell && cell.kind === 'cell') {
      const image = context.sheet().cells[cell.key]?.image
      if (image && visible(cell.row, cell.column)) { context.edit(image.id, cell.row, cell.column); return true }
    }
    return false
  }
  function selectAt(position: { x: number; y: number }): boolean {
    if (!active || disposed || context.editing() || context.composing() || context.painter()) return false
    if (cropSession.value) { context.crop(SHEET_IMAGE_CROP_OPERATIONS.close, {}); return true }
    const hit = hitSheetFloatingImage(layouts.value, position.x, position.y, context.selectedId())
    if (!hit || context.hitAllowed?.(hit.layout.image.id, position) === false) return false
    context.select(hit.layout.image.id, false, selectedIds.value.includes(hit.layout.image.id)); return true
  }
  function selectionPreview(state: ImageGesture, position: { x: number; y: number }, event: PointerEvent): SheetFloatingImage[] | undefined {
    if (!state.images) return undefined
    if (state.handle === 'rotate') {
      const start = Math.atan2(state.start.y - state.center.y, state.start.x - state.center.x), end = Math.atan2(position.y - state.center.y, position.x - state.center.x),
        delta = ((end - start) * 180 / Math.PI + 540) % 360 - 180
      return transformSheetImageSelection(state.images, context.geometry(), 0, 0, undefined, event.shiftKey ? Math.round(delta / 15) * 15 : delta)
    }
    return transformSheetImageSelection(state.images, context.geometry(), (position.x - state.start.x) / state.zoom, (position.y - state.start.y) / state.zoom, state.handle)
  }
  function keydown(event: KeyboardEvent): boolean {
    if (!active || disposed || event.defaultPrevented || context.editing() || context.composing() || event.isComposing || event.keyCode === 229 ||
      event.target instanceof Element && event.target.closest('input,textarea,button,[contenteditable="true"]')) return false
    const id = context.selectedId(), image = context.sheet().images?.find(item => item.id === id)
    if (!id || !image) return false
    const crop = cropSession.value
    if (crop) {
      if (event.key === 'Escape' || event.key === 'Enter') context.crop(SHEET_IMAGE_CROP_OPERATIONS.close, { isCancel: event.key === 'Escape' })
      else if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key) && !event.ctrlKey && !event.metaKey && !event.altKey) {
        if (!isCropCurrent(crop)) cancelCrop()
        else {
          const step = event.shiftKey ? 10 : 1, next = dragSheetImageCrop(crop.preview, context.geometry(),
            event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0, event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0)
          cropSession.value = { ...crop, preview: next, changed: JSON.stringify(next) !== JSON.stringify(crop.baseline) }; context.render()
        }
      } else return false
      event.preventDefault(); event.stopPropagation(); return true
    }
    if (event.key === 'Escape') { cancel(); context.select() }
    else if (event.key === 'Delete' || event.key === 'Backspace') {
      if (selectedEditable.value) { if (selectedImages.value.length > 1) context.removeMany([...selectedIds.value]); else context.remove(id) }
    }
    else if (event.key === 'Enter') { if (selectedEditable.value && selectedImages.value.length === 1) context.edit(id) }
    else if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key) && !event.ctrlKey && !event.metaKey && !event.altKey) {
      if (!selectedEditable.value) { event.preventDefault(); return true }
      if (gesture.value) { event.preventDefault(); event.stopPropagation(); return true }
      const step = event.shiftKey ? 10 : 1, direction = sheetDrawingMoveDirectionForKey(event.key)
      if (direction !== undefined && context.nudge?.(selectedIds.value, direction, step, context.revision())) {
        event.preventDefault(); event.stopPropagation(); return true
      }
      const dx = event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0, dy = event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0
      if (selectedImages.value.length > 1) context.updateBatch(transformSheetImageSelection(selectedImages.value, context.geometry(), dx, dy), context.revision())
      else context.update(id, moveSheetImage(context.geometry(), image, dx, dy), context.revision())
    } else return false
    event.preventDefault(); event.stopPropagation(); return true
  }
  watch([() => context.sheet(), () => context.revision(), () => context.userId(), () => context.editing(), () => context.composing(), () => context.painter(),
    () => context.viewport().scrollLeft, () => context.viewport().scrollTop, () => context.viewport().zoom, () => context.viewport().width, () => context.viewport().height], () => {
    if (gesture.value && !isCurrent(gesture.value)) cancel()
    if (cropSession.value && !isCropCurrent(cropSession.value)) cancelCrop()
  })
  watch(() => context.selectedId(), id => { if (cropSession.value && id !== cropSession.value.original.id) cancelCrop() })
  watch(selectedIds, () => {
    if (gesture.value && !isCurrent(gesture.value)) cancel()
    if (cropSession.value && !isCropCurrent(cropSession.value)) cancelCrop()
  })
  onDeactivated(() => { active = false; cancel(); cancelCrop(); cache.clear() })
  onActivated(() => { active = true; cache.sync(resources.value, backgroundConfig.value); imageVersion.value++ })
  onUnmounted(() => { disposed = true; cancel(); cancelCrop(); cache.dispose() })
  return { imageError, imageVersion, sceneOptions, menuPosition, selectedEditable, begin, move, end, editAt, selectAt, keydown, cancel,
    hasControlAt: (position: { x: number; y: number }) => context.selectionControls?.() !== false && Boolean(hitSheetFloatingImage(
      selectedImages.value.length > 1 ? selectionLayouts.value : layouts.value, position.x, position.y,
      selectedImages.value.length > 1 ? '__sheet_image_selection__' : context.selectedId())?.handle),
    operateImageCrop, isCropping: computed(() => Boolean(cropSession.value)),
    hasGesture: () => Boolean(gesture.value) }
}
