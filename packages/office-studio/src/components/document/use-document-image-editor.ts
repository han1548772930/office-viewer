import { computed, nextTick, onDeactivated, onUnmounted, ref, shallowRef, watch } from 'vue'
import type { FDocument } from '../../api/facade'
import type { DocumentModel } from '../../domain/model'
import { cloneModel } from '../../domain/serialization'
import type { DocumentSelection } from '../../features/document/model/document-model'
import { documentId } from '../../features/document/model/document-model'
import type { DocumentAuxiliaryBlockHost } from '../../features/document/model/document-block-host'
import { sameDocumentBlockHost } from '../../features/document/model/document-block-host'
import type { DocumentInputSession } from '../../features/document/model/document-input-session'
import { documentImageEntries, type DocumentImagePatch } from '../../features/document/images/document-image-model'
import { assertDocumentImageEditable } from '../../features/document/commands/document-image-commands'
import type { DocumentCommand } from '../../features/document/commands/document-commands'
import { permissionTargetsForImage, permissionTargetsForSelection, resolveDocumentPermission, type DocumentPermissionContext } from '../../features/document/permissions/document-permissions'
import { documentImageFromResource, prepareDocumentImageSource } from '../../api/document-images'
import { resolveDocumentImageEditorTarget } from '../../api/document-image-commands'
import type { DocumentImageEditorTarget, DocumentImageEditorMode } from '../../api/document-image-editors'
import { trackDocumentImageTask } from '../../api/document-image-editors'

export interface DocumentImageEditorEnvironment {
  source: () => DocumentModel; model: () => DocumentModel; selection: () => DocumentSelection; host: () => DocumentAuxiliaryBlockHost | undefined
  api: () => FDocument | undefined; session: () => DocumentInputSession | undefined; permission: () => DocumentPermissionContext; inactive: () => boolean
  execute: (command: DocumentCommand, restoreFocus?: boolean, feedback?: { selectionCurrent: boolean; error?: string }) => boolean
  focus: () => void; notify: (message: string) => void
}
export function useDocumentImageEditor(environment: DocumentImageEditorEnvironment) {
  const selectedId = ref<string>(), open = ref(false), mode = ref<DocumentImageEditorMode>('insert'), editorKey = ref(0), busy = ref(false), error = ref(''), composing = ref(false)
  const target = shallowRef<DocumentImageEditorTarget>(), selected = computed(() => documentImageEntries(environment.model()).find(entry => entry.imageId === selectedId.value))
  const image = computed(() => target.value?.imageId ? documentImageEntries(environment.model()).find(entry => entry.imageId === target.value!.imageId)?.image : undefined)
  let controller: AbortController | undefined, disposed = false, applying = false, closeTarget: DocumentImageEditorTarget | undefined
  function current(value = target.value): boolean {
    if (!value || disposed || environment.inactive() || value.source !== environment.source() || !sameDocumentBlockHost(value.host, environment.host())
      || JSON.stringify(value.selection) !== JSON.stringify(environment.selection())) return false
    const api = environment.api(), session = environment.session()
    return !api || api.id === value.unitId && api.getModel() === value.source && api.getModelRevision() === value.revision
      && api.getHistoryGeneration() === value.historyGeneration && session?.valid === true && session.generation === value.inputGeneration
  }
  function editable(imageId?: string): boolean {
    if (environment.inactive()) return false
    const model = environment.model()
    const targets = imageId ? permissionTargetsForImage(model, imageId) : permissionTargetsForSelection(model, environment.selection())
    const rules = environment.source().permissionRules ?? [], context = environment.permission()
    if (!resolveDocumentPermission(rules, context, 'edit', targets).allowed || !resolveDocumentPermission(rules, context, 'view', targets).allowed) return false
    try { if (imageId) assertDocumentImageEditable(model, imageId); return true } catch { return false }
  }
  const readOnly = computed(() => !current() || !editable(target.value?.imageId))
  function close(force = false): boolean {
    if (!force && composing.value) return false
    controller?.abort(); controller = undefined; busy.value = false; composing.value = false
    closeTarget = target.value; open.value = false
    return true
  }
  function accept(value: DocumentImageEditorTarget): boolean {
    const imageId = value.imageId ?? (value.mode === 'insert' ? undefined : selectedId.value)
    if (value.mode !== 'insert' && !imageId || !current(value) || !editable(imageId) || composing.value) return false
    close(true); closeTarget = undefined
    target.value = { ...value, imageId }; mode.value = value.mode; error.value = ''; editorKey.value++; open.value = true
    return true
  }
  function show(nextMode: DocumentImageEditorMode = 'insert', imageId?: string, paragraphMenuPlacement?: 'below'): boolean {
    if (!editable(imageId) || environment.inactive()) return false
    const api = environment.api(), session = environment.session(), source = environment.source()
    const value: DocumentImageEditorTarget = api ? resolveDocumentImageEditorTarget(api, { imageId, paragraphMenuPlacement }, nextMode)
      : { source, unitId: source.id ?? '', revision: 0, historyGeneration: 0, inputGeneration: session?.generation, segmentId: '',
        selection: cloneModel(environment.selection()), host: cloneModel(environment.host()), imageId, paragraphMenuPlacement, mode: nextMode }
    return accept(value)
  }
  function apply(patch: DocumentImagePatch): boolean {
    const value = target.value
    if (!value?.imageId || !current(value) || !editable(value.imageId) || composing.value) { error.value = '图片或输入位置已变化，请重新打开。'; return false }
    const feedback = { selectionCurrent: true, error: undefined as string | undefined }
    applying = true
    try {
      const changed = environment.execute({ id: 'doc.command.update-image', imageId: value.imageId, patch, ...(value.host ? { host: value.host } : {}) }, false, feedback)
      if (!changed) { error.value = feedback.error ?? '无法应用图片格式。'; return false }
      closeTarget = feedback.selectionCurrent ? { ...value, source: environment.source(), revision: environment.api()?.getModelRevision() ?? 0,
        historyGeneration: environment.api()?.getHistoryGeneration() ?? 0,
        inputGeneration: environment.session()?.generation, selection: cloneModel(environment.selection()) } : undefined
      open.value = false
      return true
    } finally { applying = false }
  }
  async function choose(source: string | File, alt: string): Promise<boolean> {
    const value = target.value
    if (!value || busy.value || !current(value) || !editable(value.imageId) || composing.value) return false
    const active = new AbortController(); controller?.abort(); controller = active; busy.value = true; error.value = ''
    const release = trackDocumentImageTask(value.source, active)
    try {
      const resource = await prepareDocumentImageSource(source, active.signal, () => open.value && target.value === value && current(value))
      if (active.signal.aborted || target.value !== value || !current(value)) return false
      const image = documentImageFromResource(resource, { alt, width: Math.min(500, resource.width) }, typeof source === 'string' ? source : undefined)
      const imageId = value.mode === 'replace' && value.imageId ? value.imageId : documentId('image')
      const command: DocumentCommand = value.mode === 'replace' && value.imageId ? { id: 'doc.command.replace-image', imageId: value.imageId, image }
        : { id: 'doc.command.insert-image', imageId, image, paragraphMenuPlacement: value.paragraphMenuPlacement }
      applying = true
      const feedback = { selectionCurrent: true, error: undefined as string | undefined }, changed = environment.execute({ ...command, ...(value.host ? { host: value.host } : {}) }, false, feedback)
      if (!changed) { error.value = feedback.error ?? '无法插入图片。'; return false }
      if (feedback.selectionCurrent) selectedId.value = imageId
      closeTarget = feedback.selectionCurrent ? { ...value, source: environment.source(), revision: environment.api()?.getModelRevision() ?? 0,
        historyGeneration: environment.api()?.getHistoryGeneration() ?? 0,
        inputGeneration: environment.session()?.generation, selection: cloneModel(environment.selection()) } : undefined
      open.value = false
      return true
    } catch (reason) {
      if (!active.signal.aborted && target.value === value && open.value) error.value = reason instanceof Error ? reason.message : String(reason)
      return false
    } finally { release(); applying = false; if (controller === active) { controller = undefined; busy.value = false } }
  }
  function closed() {
    const value = closeTarget ?? target.value
    closeTarget = undefined
    if (value && current(value)) void nextTick(() => { if (!open.value && current(value)) environment.focus() })
  }
  watch(() => [environment.source(), environment.host(), environment.selection(), environment.session()?.generation, environment.inactive()], () => {
    if (applying) return
    if (open.value && !current()) { controller?.abort(); busy.value = false; error.value = '文档、图片或输入位置已变化，请重新打开。' }
    if (selectedId.value && !selected.value) selectedId.value = undefined
  })
  watch(() => environment.api(), (api, _previous, cleanup) => {
    if (!api) return
    const owner = api.bindImageEditor({ open: accept, close })
    cleanup(() => owner.dispose())
  }, { immediate: true })
  onDeactivated(() => { close(true); selectedId.value = undefined })
  onUnmounted(() => { disposed = true; close(true) })
  return { selectedId, selected, open, mode, image, editorKey, busy, error, composing, readOnly, editable, current, accept, show, close, apply, choose, closed }
}
