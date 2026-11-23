import { replaceSheetCellInputSelection, type SheetCellInputSelection } from './sheet-cell-edit-input.ts'
import type { SheetCellNativeInputEvent } from './sheet-cell-input-trace.ts'

interface NativeBinding { sessionId: string; hostId: string; value: string }
export interface SheetCellInputBridgeContext { sessionId?: string; value: string }
let nextNativeHost = 0

/** Capture native events around library-owned controls; this does not add a second DOM input. */
export function createSheetCellNativeInputBridge(options: {
  host: 'cell' | 'formula'
  sessionId: () => string | undefined
  value: () => string
  contains: (element: HTMLInputElement | HTMLTextAreaElement) => boolean
  emit: (event: SheetCellNativeInputEvent) => void
}) {
  const bindings = new WeakMap<HTMLInputElement | HTMLTextAreaElement, NativeBinding>()
  const elementFor = (target: EventTarget | null): HTMLInputElement | HTMLTextAreaElement | undefined => {
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) || !options.contains(target)) return undefined
    return target
  }
  const bind = (element: HTMLInputElement | HTMLTextAreaElement): NativeBinding | undefined => {
    const sessionId = options.sessionId()
    if (!sessionId || !options.contains(element)) return undefined
    const current = bindings.get(element)
    // Session-keyed controls are replaced. A queued event from a former control cannot rebind itself.
    if (current) return current.sessionId === sessionId ? current : undefined
    const value = options.value()
    const binding = { sessionId, hostId: `sheet-input-${options.host}-${++nextNativeHost}`, value }
    bindings.set(element, binding)
    options.emit({ sessionId, hostId: binding.hostId, host: options.host, phase: 'bind', value,
      start: Math.min(value.length, element.selectionStart ?? value.length), end: Math.min(value.length, element.selectionEnd ?? value.length) })
    return binding
  }
  const capture = (event: Event, phase: SheetCellNativeInputEvent['phase']): void => {
    const element = elementFor(event.target)
    if (!element) return
    const binding = bind(element)
    if (!binding) return
    const input = event as InputEvent
    const impliedType = event instanceof KeyboardEvent ? event.key === 'Backspace' ? 'deleteContentBackward'
      : event.key === 'Delete' ? 'deleteContentForward' : undefined
      : event instanceof ClipboardEvent && event.type === 'paste' ? 'insertFromPaste' : undefined
    const value = element.value
    options.emit({ sessionId: binding.sessionId, hostId: binding.hostId, host: options.host, phase, value,
      start: element.selectionStart ?? value.length, end: element.selectionEnd ?? value.length,
      ...(typeof input.inputType === 'string' && input.inputType ? { inputType: input.inputType } : impliedType ? { inputType: impliedType } : {}),
      ...('data' in event ? { data: (event as InputEvent | CompositionEvent).data } : {}),
      ...('isComposing' in event ? { composing: input.isComposing } : {}) })
    binding.value = value
  }
  const replace = (element: HTMLInputElement | HTMLTextAreaElement, text: string,
    selection?: SheetCellInputSelection, inputType = 'insertReplacementText', before?: string) => {
    const binding = bind(element)
    if (!binding) return undefined
    const value = before ?? options.value()
    const start = Math.max(0, Math.min(value.length, selection?.start ?? 0))
    const end = Math.max(start, Math.min(value.length, selection?.end ?? value.length))
    const insertion = text.replace(/\r\n?/gu, '\n')
    const replacement = replaceSheetCellInputSelection(value, insertion, { start, end })
    options.emit({ sessionId: binding.sessionId, hostId: binding.hostId, host: options.host, phase: 'replace',
      value, start, end, text: insertion, inputType })
    binding.value = replacement.value
    return replacement
  }
  return { capture, replace }
}
