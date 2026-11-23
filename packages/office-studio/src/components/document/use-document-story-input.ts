import { nextTick, ref, watch } from 'vue'
import type { DocumentModel } from '../../domain/model'
import type { DocumentLayout } from '../../features/document/layout/document-layout'
import { cloneDocument, type RunPatch } from '../../features/document/model/document-model'
import { cloneDocumentInputMarks, type DocumentInputSession, type DocumentStoryInputCommit } from '../../features/document/model/document-input-session'
import { clampDocumentTextRange } from '../../features/document/model/document-story'
import { applySectionHeaderFooter, resolveSectionHeaderFooter } from '../../features/document/services/section-header-footer'
import { editStoryBuffer, type StoryLocation } from '../../features/document/services/story-editor'

interface StoryRange { start: number; end: number }
interface StoryInputProps {
  model: DocumentModel
  story?: StoryLocation
  storyValue?: string
  storySelection?: StoryRange
  storyReadonly?: boolean
  typingMarks?: RunPatch
  captureInputSession?: (sourceModel?: DocumentModel) => DocumentInputSession | undefined
}
interface StoryInputHost {
  element: () => HTMLTextAreaElement | null | undefined
  layout: () => DocumentLayout
  active: () => boolean
  edit: (value: string, caret: number, marks?: RunPatch, input?: DocumentStoryInputCommit, sourceRange?: StoryRange) => void
  select: (range: StoryRange, preserveInputStyle?: boolean) => void
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
  let composition: { scope: string; model: DocumentModel; value: string; element: HTMLTextAreaElement; range: StoryRange; marks: RunPatch; input?: Omit<DocumentStoryInputCommit, 'caret'> } | undefined
  let blockedElement: HTMLTextAreaElement | null | undefined
  let inputRangeBeforeEdit: StoryRange | undefined
  let inputMarksBeforeEdit: RunPatch | undefined
  let inputSessionBeforeEdit: Omit<DocumentStoryInputCommit, 'caret'> | undefined
  let submitted: { scope: string; value: string; range?: StoryRange } | undefined
  let committed: { scope: string; value: string } | undefined
  let pending: { scope: string; range?: StoryRange; focus: boolean } | undefined
  let scheduled = false
  let epoch = 0

  function scope(location = props.story): string {
    if (!location) return ''
    const page = host.layout().pageHeaderFooters[location.page]
    if (!page || page.isParityBlank) return ''
    return `${props.model.id}\u0000${location.page}\u0000${location.area}\u0000${page?.sectionId ?? ''}\u0000${location.area === 'header' ? page?.headerStory ?? '' : page?.footerStory ?? ''}`
  }
  function storyRange(): StoryRange {
    const element = host.element()
    return { start: element?.selectionStart ?? props.storySelection?.start ?? 0, end: element?.selectionEnd ?? props.storySelection?.end ?? 0 }
  }
  function captureInput(): { marks: RunPatch; input?: Omit<DocumentStoryInputCommit, 'caret'> } {
    const location = props.story
    const page = location && host.layout().pageHeaderFooters[location.page]
    if (!page || page.isParityBlank) return { marks: cloneDocumentInputMarks(props.typingMarks ?? {}) }
    const state = props.captureInputSession?.(props.model)
    const key = location?.area === 'header' ? page?.headerStory : page?.footerStory
    const range = storyRange()
    if (state?.valid && state.location.kind === 'story' && page?.sectionId === state.location.range.sectionId && key === state.location.range.story
      && range.start === state.location.range.start && range.end === state.location.range.end) {
      return { marks: cloneDocumentInputMarks(state.marks), input: { range: { ...state.location.range }, modelRevision: state.modelRevision, generation: state.generation } }
    }
    return { marks: cloneDocumentInputMarks(props.typingMarks ?? {}) }
  }
  function inputSessionIsCurrent(input: Omit<DocumentStoryInputCommit, 'caret'> | undefined): boolean {
    if (!props.captureInputSession) return true
    const state = props.captureInputSession()
    if (!input) return state === undefined
    return state?.valid === true && state.modelRevision === input.modelRevision && state.generation === input.generation
      && state.location.kind === 'story' && state.location.range.sectionId === input.range.sectionId && state.location.range.story === input.range.story
      && state.location.range.start === input.range.start && state.location.range.end === input.range.end
  }
  function storyEventIsCurrent(event?: Event): boolean {
    const element = host.element()
    return host.active() && !!scope() && !!element && element !== blockedElement && (!event || event.target === element)
  }
  function syncStorySelection(event?: Event, preserveInputStyle = false) {
    if (!storyEventIsCurrent(event)) return
    const range = storyRange(), previous = props.storySelection
    if (!previous || previous.start !== range.start || previous.end !== range.end) host.select(range, preserveInputStyle || composing.value || inputRangeBeforeEdit !== undefined || submitted?.scope === scope())
    host.draw()
  }
  function restoreValue(element: HTMLTextAreaElement) {
    const value = props.storyValue ?? ''
    inputValue.value = value
    if (element.value !== value) element.value = value
  }
  function selectionFor(value: string, range: StoryRange): StoryRange {
    return clampDocumentTextRange(value, range)
  }
  function scheduleRestore(focus: boolean, range?: StoryRange, location = props.story) {
    if (!location) return
    const target = scope(location), previous = pending?.scope === target ? pending : undefined
    if (!target) return
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
      syncStorySelection(undefined, true)
    })
  }
  function focusStory(offset?: number, end = offset, location = props.story) {
    scheduleRestore(true, offset === undefined ? undefined : { start: offset, end: end ?? offset }, location)
  }
  function cancelStoryInput(refocus = true) {
    epoch++
    scheduled = false
    pending = undefined
    committed = undefined
    submitted = undefined
    inputRangeBeforeEdit = undefined
    inputMarksBeforeEdit = undefined
    inputSessionBeforeEdit = undefined
    const hadComposition = composing.value || composition !== undefined
    composition = undefined
    composing.value = false
    if (hadComposition) { host.composing(false); host.preview(null) }
    inputValue.value = props.storyValue ?? ''
    const element = host.element()
    if (hadComposition) { blockedElement = element; bridgeVersion.value++ }
    if (element) restoreValue(element)
    if (refocus && hadComposition && host.active() && props.story) scheduleRestore(true)
  }
  function compositionIsCurrent(): boolean {
    return !!composition && host.active() && !props.storyReadonly && composition.scope === scope()
      && composition.model === props.model && composition.value === (props.storyValue ?? '') && composition.element === host.element()
      && inputSessionIsCurrent(composition.input)
  }
  function commit(value: string, marks = inputMarksBeforeEdit ?? props.typingMarks ?? {}, source = inputSessionBeforeEdit, sourceRange = inputRangeBeforeEdit) {
    if (!host.active() || !scope() || props.storyReadonly || host.element() === blockedElement) { scheduleRestore(false); return }
    const key = scope()
    if ((value === (props.storyValue ?? '') && (!sourceRange || sourceRange.start === sourceRange.end)) || committed?.scope === key && committed.value === value) {
      inputRangeBeforeEdit = undefined; inputMarksBeforeEdit = undefined; inputSessionBeforeEdit = undefined
      return
    }
    committed = { scope: key, value }
    submitted = { scope: key, value, range: sourceRange }
    inputRangeBeforeEdit = undefined
    inputMarksBeforeEdit = undefined
    inputSessionBeforeEdit = undefined
    inputValue.value = value
    const caret = storyRange().start
    try { host.edit(value, caret, marks, source ? { ...source, range: { ...source.range }, caret } : undefined, sourceRange) }
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
    if (!location || !page || page.isParityBlank) return
    try {
      const model = cloneDocument(props.model), key = location.area === 'header' ? page.headerStory : page.footerStory
      applySectionHeaderFooter(model, page.sectionId, editStoryBuffer(resolveSectionHeaderFooter(model, page.sectionId), key, text, storyRange().start, model, composition!.marks, composition!.range))
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
    if (!composing.value && !event.isComposing && !event.inputType.startsWith('history')) {
      syncStorySelection(event)
      const captured = captureInput()
      inputRangeBeforeEdit = storyRange(); inputMarksBeforeEdit = captured.marks; inputSessionBeforeEdit = captured.input; committed = undefined
    }
  }
  function storyCompositionStart(event?: CompositionEvent) {
    const element = host.element()
    if (!storyEventIsCurrent(event) || props.storyReadonly || !element) return
    epoch++
    pending = undefined
    scheduled = false
    committed = undefined
    submitted = undefined
    inputRangeBeforeEdit = undefined
    inputMarksBeforeEdit = undefined
    inputSessionBeforeEdit = undefined
    syncStorySelection(event)
    const range = storyRange()
    const captured = captureInput()
    if ((props.captureInputSession && !captured.input) || !inputSessionIsCurrent(captured.input)) {
      blockedElement = element
      bridgeVersion.value++
      restoreValue(element)
      return
    }
    inputRangeBeforeEdit = range
    composition = { scope: scope(), model: props.model, value: props.storyValue ?? '', element, range: { ...range }, ...captured }
    composing.value = true
    host.composing(true)
  }
  function storyCompositionEnd(event?: CompositionEvent) {
    if (!storyEventIsCurrent(event)) return
    if (!composing.value) { scheduleRestore(false); return }
    if (!compositionIsCurrent()) { cancelStoryInput(); return }
    const value = host.element()?.value ?? inputValue.value
    const marks = composition!.marks
    const input = composition!.input
    const range = composition!.range
    composition = undefined
    composing.value = false
    host.composing(false)
    host.preview(null)
    syncStorySelection()
    commit(value, marks, input, range)
  }
  function storyFocused(event?: FocusEvent) {
    if (!storyEventIsCurrent(event)) return
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
  watch(() => props.captureInputSession?.(), () => {
    if (composition && !compositionIsCurrent()) cancelStoryInput(false)
  }, { flush: 'sync' })

  return { inputValue, composing, bridgeVersion, storyRange, storyEventIsCurrent, syncStorySelection, focusStory, storyInputChanged, captureStoryInput, storyCompositionStart, storyCompositionEnd, storyFocused, cancelStoryInput }
}
