import { nextTick, ref, watch } from 'vue'
import type { DocumentModel } from '../../domain/model'
import type { DocumentLayout } from '../../features/document/layout/document-layout'
import { cloneDocument } from '../../features/document/model/document-model'
import { isDocumentTextBoundary } from '../../features/document/model/document-story'
import { applySectionHeaderFooter, resolveSectionHeaderFooter } from '../../features/document/services/section-header-footer'
import { editStoryBuffer, type StoryLocation } from '../../features/document/services/story-editor'

interface StoryRange { start: number; end: number }
interface StoryInputProps {
  model: DocumentModel
  story?: StoryLocation
  storyValue?: string
  storySelection?: StoryRange
  storyReadonly?: boolean
}
interface StoryInputHost {
  element: () => HTMLTextAreaElement | null | undefined
  layout: () => DocumentLayout
  active: () => boolean
  edit: (value: string, caret: number) => void
  select: (range: StoryRange) => void
  composing: (value: boolean) => void
  preview: (model: DocumentModel | null) => void
  draw: () => void
  notify: (message: string) => void
}

/** TextArea owns only the current native input draft; the document remains authoritative. */
export function useDocumentStoryInput(props: StoryInputProps, host: StoryInputHost) {
  const inputValue = ref(props.storyValue ?? '')
  const composing = ref(false)
  const bridgeVersion = ref(0)
  let composition: { scope: string; model: DocumentModel; value: string; element: HTMLTextAreaElement } | undefined
  let blockedElement: HTMLTextAreaElement | null | undefined
  let inputRangeBeforeEdit: StoryRange | undefined
  let submitted: { scope: string; value: string; range?: StoryRange } | undefined
  let committed: { scope: string; value: string } | undefined
  let pending: { scope: string; range?: StoryRange; focus: boolean } | undefined
  let scheduled = false
  let epoch = 0

  function scope(location = props.story): string {
    if (!location) return ''
    const page = host.layout().pageHeaderFooters[location.page]
    return `${props.model.id}\u0000${location.page}\u0000${location.area}\u0000${page?.sectionId ?? ''}\u0000${location.area === 'header' ? page?.headerStory ?? '' : page?.footerStory ?? ''}`
  }
  function storyRange(): StoryRange {
    const element = host.element()
    return { start: element?.selectionStart ?? props.storySelection?.start ?? 0, end: element?.selectionEnd ?? props.storySelection?.end ?? 0 }
  }
  function storyEventIsCurrent(event?: Event): boolean {
    const element = host.element()
    return host.active() && !!props.story && !!element && element !== blockedElement && (!event || event.target === element)
  }
  function syncStorySelection(event?: Event) {
    if (!storyEventIsCurrent(event)) return
    const range = storyRange(), previous = props.storySelection
    if (!previous || previous.start !== range.start || previous.end !== range.end) host.select(range)
    host.draw()
  }
  function restoreValue(element: HTMLTextAreaElement) {
    const value = props.storyValue ?? ''
    inputValue.value = value
    if (element.value !== value) element.value = value
  }
  function selectionFor(value: string, range: StoryRange): StoryRange {
    const clamp = (offset: number) => Math.max(0, Math.min(value.length, Number.isSafeInteger(offset) ? offset : 0))
    let start = clamp(Math.min(range.start, range.end)), end = clamp(Math.max(range.start, range.end))
    if (!isDocumentTextBoundary(value, start)) start--
    if (!isDocumentTextBoundary(value, end)) end++
    return { start, end }
  }
  function scheduleRestore(focus: boolean, range?: StoryRange, location = props.story) {
    if (!location) return
    const target = scope(location), previous = pending?.scope === target ? pending : undefined
    pending = { scope: target, focus: focus || previous?.focus === true, range: range ?? previous?.range }
    if (scheduled) return
    scheduled = true
    const request = epoch
    void nextTick(() => {
      if (request !== epoch) return
      scheduled = false
      const task = pending
      pending = undefined
      if (!task || !host.active() || !props.story || task.scope !== scope() || composing.value) return
      const element = host.element()
      if (!element) return
      restoreValue(element)
      const wasRejected = submitted?.scope === task.scope && submitted.value !== element.value
      const rejected = wasRejected ? submitted?.range : undefined
      const range = selectionFor(element.value, task.range ?? rejected ?? props.storySelection ?? storyRange())
      submitted = undefined
      if (wasRejected) committed = undefined
      if (task.focus && document.activeElement !== element) element.focus({ preventScroll: true })
      element.setSelectionRange(range.start, range.end)
      syncStorySelection()
    })
  }
  function focusStory(offset?: number, end = offset, location = props.story) {
    scheduleRestore(true, offset === undefined ? undefined : { start: offset, end: end ?? offset }, location)
  }
  function cancelStoryInput() {
    epoch++
    scheduled = false
    pending = undefined
    committed = undefined
    submitted = undefined
    inputRangeBeforeEdit = undefined
    const hadComposition = composing.value || composition !== undefined
    composition = undefined
    composing.value = false
    if (hadComposition) { host.composing(false); host.preview(null) }
    inputValue.value = props.storyValue ?? ''
    const element = host.element()
    if (hadComposition) { blockedElement = element; bridgeVersion.value++ }
    if (element) restoreValue(element)
    if (hadComposition && host.active() && props.story) scheduleRestore(true)
  }
  function compositionIsCurrent(): boolean {
    return !!composition && host.active() && !props.storyReadonly && composition.scope === scope()
      && composition.model === props.model && composition.value === (props.storyValue ?? '') && composition.element === host.element()
  }
  function commit(value: string) {
    if (!host.active() || !props.story || props.storyReadonly || host.element() === blockedElement) { scheduleRestore(false); return }
    const key = scope()
    if (value === (props.storyValue ?? '') || committed?.scope === key && committed.value === value) return
    committed = { scope: key, value }
    submitted = { scope: key, value, range: inputRangeBeforeEdit }
    inputRangeBeforeEdit = undefined
    inputValue.value = value
    try { host.edit(value, storyRange().start) }
    catch (error) { host.notify(error instanceof Error ? error.message : String(error)) }
    // A command can be rejected without changing the prop. Restore both the
    // controlled draft and native buffer after the parent has handled it.
    scheduleRestore(false)
  }
  function storyValueChanged(value: unknown) {
    syncStorySelection()
    const text = String(value ?? '')
    if (!composing.value) { commit(text); return }
    if (!compositionIsCurrent()) { cancelStoryInput(); return }
    inputValue.value = text
    const location = props.story, page = location && host.layout().pageHeaderFooters[location.page]
    if (!location || !page) return
    try {
      const model = cloneDocument(props.model), key = location.area === 'header' ? page.headerStory : page.footerStory
      applySectionHeaderFooter(model, page.sectionId, editStoryBuffer(resolveSectionHeaderFooter(model, page.sectionId), key, text, storyRange().start, model))
      host.preview(model)
    } catch (error) {
      host.preview(null)
      host.notify(error instanceof Error ? error.message : String(error))
    }
  }
  function storyInputChanged(event: Event) {
    if (!(event.target instanceof HTMLTextAreaElement) || !storyEventIsCurrent(event)) return
    storyValueChanged(event.target.value)
  }
  function captureStoryInput(event: InputEvent) {
    if (!storyEventIsCurrent(event) || props.storyReadonly) return
    if (!composing.value && !event.isComposing && !event.inputType.startsWith('history')) { inputRangeBeforeEdit = storyRange(); committed = undefined }
  }
  function storyCompositionStart(event?: CompositionEvent) {
    const element = host.element()
    if (!storyEventIsCurrent(event) || props.storyReadonly || !element) return
    epoch++
    pending = undefined
    scheduled = false
    committed = undefined
    submitted = undefined
    inputRangeBeforeEdit = storyRange()
    composition = { scope: scope(), model: props.model, value: props.storyValue ?? '', element }
    composing.value = true
    host.composing(true)
  }
  function storyCompositionEnd(event?: CompositionEvent) {
    if (!storyEventIsCurrent(event)) return
    if (!composing.value) { scheduleRestore(false); return }
    if (!compositionIsCurrent()) { cancelStoryInput(); return }
    const value = host.element()?.value ?? inputValue.value
    composition = undefined
    composing.value = false
    host.composing(false)
    host.preview(null)
    syncStorySelection()
    commit(value)
  }
  function storyFocused() {
    if (!storyEventIsCurrent()) return
    if (!composing.value) {
      const element = host.element()
      if (element) restoreValue(element)
    }
    syncStorySelection()
  }

  watch([() => props.model, () => scope(), () => props.storyReadonly, () => props.storyValue], ([model, key, readonly, value], [previousModel, previousKey]) => {
    if (composition && !compositionIsCurrent()) cancelStoryInput()
    if (key !== previousKey || readonly || model !== previousModel && committed?.value !== value) committed = undefined
    if (!composing.value) { inputValue.value = props.storyValue ?? ''; scheduleRestore(false) }
    host.draw()
  }, { flush: 'sync' })
  watch(() => props.storySelection, () => { if (!composing.value) scheduleRestore(false); host.draw() })

  return { inputValue, composing, bridgeVersion, storyRange, storyEventIsCurrent, syncStorySelection, focusStory, storyInputChanged, captureStoryInput, storyCompositionStart, storyCompositionEnd, storyFocused, cancelStoryInput }
}
