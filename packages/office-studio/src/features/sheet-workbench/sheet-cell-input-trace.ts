/** All offsets refer to the native antdv-next input's UTF-16 value before this operation. */
export interface SheetCellTextEditOperation {
  start: number
  end: number
  text: string
  deletedText: string
  inputType?: string
  origin?: 'beforeinput' | 'composition' | 'programmatic' | 'input-caret' | 'compatibility'
}

export interface SheetCellInputEditTrace {
  version: 1
  sessionId: string
  source: string
  operations: SheetCellTextEditOperation[]
}

export interface SheetCellNativeInputSnapshot {
  value: string
  start: number
  end: number
  inputType?: string
  data?: string | null
  composing?: boolean
}

export interface SheetCellNativeInputEvent extends SheetCellNativeInputSnapshot {
  sessionId: string
  hostId: string
  host: 'cell' | 'formula'
  phase: 'bind' | 'selection' | 'beforeinput' | 'input' | 'compositionstart' | 'compositionend' | 'replace'
  text?: string
}

export function isSheetCellTextBoundary(value: string, offset: number): boolean {
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > value.length) return false
  return offset === 0 || offset === value.length
    || !(value.charCodeAt(offset - 1) >= 0xd800 && value.charCodeAt(offset - 1) <= 0xdbff
      && value.charCodeAt(offset) >= 0xdc00 && value.charCodeAt(offset) <= 0xdfff)
}

export function applySheetCellTextEdit(value: string, operation: SheetCellTextEditOperation): string | undefined {
  if (!isSheetCellTextBoundary(value, operation.start) || !isSheetCellTextBoundary(value, operation.end)
    || operation.end < operation.start || typeof operation.text !== 'string' || typeof operation.deletedText !== 'string'
    || value.slice(operation.start, operation.end) !== operation.deletedText) return undefined
  return value.slice(0, operation.start) + operation.text + value.slice(operation.end)
}

/** A trace is untrusted command input: clone only its declared fields and validate every source interval. */
export function readSheetCellInputEditTrace(value: unknown): SheetCellInputEditTrace {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('输入范围轨迹必须是对象。')
  const trace = value as Record<string, unknown>
  if (trace.version !== 1 || typeof trace.sessionId !== 'string' || !trace.sessionId.trim()
    || typeof trace.source !== 'string' || !Array.isArray(trace.operations)) throw new TypeError('输入范围轨迹无效。')
  let current = trace.source
  const operations: SheetCellTextEditOperation[] = []
  for (const item of trace.operations) {
    if (!item || typeof item !== 'object' || Array.isArray(item)) throw new TypeError('输入范围操作无效。')
    const source = item as Record<string, unknown>
    if (typeof source.start !== 'number' || typeof source.end !== 'number' || typeof source.text !== 'string'
      || typeof source.deletedText !== 'string' || source.inputType !== undefined && typeof source.inputType !== 'string'
      || source.origin !== undefined && (typeof source.origin !== 'string'
        || !['beforeinput', 'composition', 'programmatic', 'input-caret', 'compatibility'].includes(source.origin))) {
      throw new TypeError('输入范围操作无效。')
    }
    const operation: SheetCellTextEditOperation = { start: source.start, end: source.end,
      text: source.text, deletedText: source.deletedText,
      ...(source.inputType === undefined ? {} : { inputType: source.inputType as string }),
      ...(source.origin === undefined ? {} : { origin: source.origin as SheetCellTextEditOperation['origin'] }) }
    const next = applySheetCellTextEdit(current, operation)
    if (next === undefined) throw new TypeError('输入范围与源文本不一致。')
    operations.push(operation); current = next
  }
  return { version: 1, sessionId: trace.sessionId, source: trace.source, operations }
}

export function replaySheetCellInputEditTrace(trace: SheetCellInputEditTrace): string | undefined {
  let value = trace.source
  for (const operation of trace.operations) {
    const next = applySheetCellTextEdit(value, operation)
    if (next === undefined) return undefined
    value = next
  }
  return value
}

/** Map the editor's optional apostrophe escape into raw cell intervals without a final-string diff. */
export function projectSheetCellInputEditTrace(trace: SheetCellInputEditTrace, before: string, after: string,
  textFormat: boolean): SheetCellTextEditOperation[] | undefined {
  let editor = trace.source
  let prefix = editor === before ? 0 : !textFormat && editor === "'" + before ? 1 : -1
  if (prefix < 0) return undefined
  let raw = before
  const operations: SheetCellTextEditOperation[] = []
  const append = (start: number, end: number, text: string, source: SheetCellTextEditOperation): boolean => {
    const operation: SheetCellTextEditOperation = { ...source, start, end, text, deletedText: raw.slice(start, end) }
    const next = applySheetCellTextEdit(raw, operation)
    if (next === undefined) return false
    // Removing only the editor escape has no effect on the rich text's typing affinity.
    if (start !== end || text !== '') operations.push(operation)
    raw = next
    return true
  }
  for (const source of trace.operations) {
    const nextEditor = applySheetCellTextEdit(editor, source)
    if (nextEditor === undefined) return undefined
    let retainedPrefix = prefix
    if (prefix && source.start === 0) {
      // Inserting before the old escape turns it into an actual character; replacing it consumes it.
      if (!append(0, Math.max(0, source.end - prefix), source.text + (source.end === 0 ? "'" : ''), source)) return undefined
      retainedPrefix = 0
    } else if (!append(Math.max(0, source.start - prefix), Math.max(0, source.end - prefix), source.text, source)) return undefined
    const nextPrefix = !textFormat && nextEditor.startsWith("'") ? 1 : 0
    if (nextPrefix > retainedPrefix && !append(0, 1, '', source)) return undefined
    if (nextPrefix < retainedPrefix) return undefined
    prefix = nextPrefix; editor = nextEditor
    if (raw !== editor.slice(prefix)) return undefined
  }
  return raw === after ? operations : undefined
}

function rangeOperation(before: string, after: string, start: number, end: number,
  inputType: string | undefined, origin: SheetCellTextEditOperation['origin']): SheetCellTextEditOperation | undefined {
  if (end < start || !isSheetCellTextBoundary(before, start) || !isSheetCellTextBoundary(before, end)) return undefined
  const insertedLength = after.length - before.length + end - start
  if (insertedLength < 0 || before.slice(0, start) !== after.slice(0, start)
    || before.slice(end) !== after.slice(start + insertedLength)) return undefined
  const operation: SheetCellTextEditOperation = { start, end, text: after.slice(start, start + insertedLength),
    deletedText: before.slice(start, end), inputType, origin }
  return applySheetCellTextEdit(before, operation) === after ? operation : undefined
}

/** Compatibility input has no source selection; its ambiguity remains explicit in the exported trace. */
export function inferSheetCellTextEdit(before: string, after: string, inputType?: string): SheetCellTextEditOperation {
  const oldPoints = [...before], newPoints = [...after]
  let first = 0, last = 0
  while (first < oldPoints.length && first < newPoints.length && oldPoints[first] === newPoints[first]) first++
  while (last < oldPoints.length - first && last < newPoints.length - first
    && oldPoints[oldPoints.length - last - 1] === newPoints[newPoints.length - last - 1]) last++
  const start = oldPoints.slice(0, first).join('').length
  const end = before.length - oldPoints.slice(oldPoints.length - last).join('').length
  return { start, end, text: newPoints.slice(first, newPoints.length - last).join(''),
    deletedText: before.slice(start, end), inputType, origin: 'compatibility' }
}

function nativeOperation(before: SheetCellNativeInputSnapshot, after: SheetCellNativeInputSnapshot,
  origin: SheetCellTextEditOperation['origin']): SheetCellTextEditOperation | undefined {
  const inputType = before.inputType ?? after.inputType
  if (before.start !== before.end) return rangeOperation(before.value, after.value, before.start, before.end, inputType, origin)
  const removed = before.value.length - after.value.length
  if (inputType?.startsWith('delete') && removed > 0) {
    if (inputType.endsWith('Backward')) return rangeOperation(before.value, after.value, before.start - removed, before.start, inputType, origin)
    if (inputType.endsWith('Forward')) return rangeOperation(before.value, after.value, before.start, before.start + removed, inputType, origin)
  }
  const insertion = rangeOperation(before.value, after.value, before.start, before.end, inputType, origin)
  if (insertion) return insertion
  // Spellcheck/autofill can replace a native target interval larger than selectionStart/selectionEnd.
  const data = after.data ?? before.data
  if (typeof data === 'string' && after.start === after.end) {
    const start = after.start - data.length, end = start + before.value.length - after.value.length + data.length
    const replacement = rangeOperation(before.value, after.value, start, end, inputType, 'input-caret')
    if (replacement?.text === data && start <= before.start && end >= before.end) return replacement
  }
  return undefined
}

interface CompositionDraft {
  hostId: string
  before: string
  start: number
  end: number
  journal: SheetCellTextEditOperation[]
  cursor: number
  origin?: SheetCellTextEditOperation['origin']
}

/** One shared draft owns both input hosts; native undo revisits recorded ranges instead of guessing an inverse splice. */
export class SheetCellInputEditRecorder {
  readonly sessionId: string
  readonly source: string
  private current: string
  private journal: SheetCellTextEditOperation[] = []
  private cursor = 0
  private hosts = new Map<string, SheetCellNativeInputEvent['host']>()
  private pending: { hostId: string; snapshot: SheetCellNativeInputSnapshot; origin: SheetCellTextEditOperation['origin'] } | undefined
  private selections = new Map<string, SheetCellNativeInputSnapshot>()
  private composition: CompositionDraft | undefined
  private completedComposition: CompositionDraft | undefined

  constructor(sessionId: string, source: string) {
    this.sessionId = sessionId; this.source = source; this.current = source
  }

  get value(): string { return this.current }
  get composing(): boolean { return this.composition !== undefined }
  get compositionHost(): SheetCellNativeInputEvent['host'] | undefined {
    return this.composition ? this.hosts.get(this.composition.hostId) : undefined
  }
  get trace(): SheetCellInputEditTrace {
    return { version: 1, sessionId: this.sessionId, source: this.source,
      operations: this.journal.slice(0, this.cursor).map(operation => ({ ...operation })) }
  }

  handle(event: SheetCellNativeInputEvent): boolean {
    if (event.sessionId !== this.sessionId || !event.hostId) return false
    if (typeof event.value !== 'string' || !isSheetCellTextBoundary(event.value, event.start)
      || !isSheetCellTextBoundary(event.value, event.end) || event.start > event.end) return false
    if (event.phase === 'bind') {
      if (event.value !== this.current) return false
      this.hosts.set(event.hostId, event.host); this.selections.set(event.hostId, { ...event })
      return true
    }
    if (this.hosts.get(event.hostId) !== event.host) return false
    if (event.phase === 'selection' || event.phase === 'beforeinput') {
      if (event.value !== this.current) return false
      this.selections.set(event.hostId, { ...event })
      if (event.phase === 'beforeinput') this.pending = { hostId: event.hostId, snapshot: { ...event }, origin: 'beforeinput' }
      return true
    }
    if (event.phase === 'compositionstart') {
      if (event.value !== this.current || this.composition) return false
      this.completedComposition = undefined
      this.composition = { hostId: event.hostId, before: this.current, start: event.start, end: event.end, journal: this.journal, cursor: this.cursor }
      this.pending = undefined
      return true
    }
    if (event.phase === 'replace') {
      if (event.value !== this.current || this.composition || typeof event.text !== 'string') return false
      const operation: SheetCellTextEditOperation = { start: event.start, end: event.end, text: event.text,
        deletedText: this.current.slice(event.start, event.end), inputType: event.inputType, origin: 'programmatic' }
      if (!this.append(operation)) return false
      this.pending = undefined; this.completedComposition = undefined
      return true
    }
    if (event.phase !== 'input' && event.phase !== 'compositionend') return false
    if (this.composition && event.hostId !== this.composition.hostId) return false
    const pending = this.pending?.hostId === event.hostId ? this.pending : undefined
    this.pending = undefined
    const inputType = event.inputType ?? pending?.snapshot.inputType
    if (inputType === 'historyUndo' || inputType === 'historyRedo') {
      if ((pending?.snapshot ?? this.selections.get(event.hostId))?.value !== this.current) return false
      this.composition = undefined; this.completedComposition = undefined
      const restored = this.restoreNativeHistory(event.value, inputType === 'historyUndo' ? -1 : 1)
      if (restored) this.selections.set(event.hostId, { ...event })
      return restored
    }
    if (!this.composition && event.phase === 'input' && (event.composing || pending?.snapshot.composing)) {
      const before = pending?.snapshot ?? this.selections.get(event.hostId)
      if (before?.value !== this.current) return false
      this.completedComposition = undefined
      this.composition = { hostId: event.hostId, before: this.current, start: before.start, end: before.end,
        journal: this.journal, cursor: this.cursor }
    }
    if (this.composition) {
      const draft = this.composition
      if (event.hostId !== draft.hostId) return false
      // Cancellation restores both the original text and its rich-run identity path.
      if (event.phase === 'compositionend' && event.data === '' && event.value === draft.before) {
        this.journal = draft.journal; this.cursor = draft.cursor; this.current = draft.before
        this.composition = undefined; this.completedComposition = undefined
        return true
      }
      if (this.cursor === draft.cursor && pending?.snapshot.value === draft.before
        && !rangeOperation(draft.before, event.value, draft.start, draft.end, inputType, 'composition')) {
        const located = nativeOperation(pending.snapshot, event, 'composition')
        if (located) { draft.start = located.start; draft.end = located.end; draft.origin = located.origin }
      }
      if (!this.updateComposition(draft, event.value)) return false
      if (event.phase === 'compositionend') { this.composition = undefined; this.completedComposition = draft }
      this.selections.set(event.hostId, { ...event })
      return true
    }
    if (event.phase === 'compositionend') return event.value === this.current
    if (this.completedComposition?.hostId === event.hostId && (inputType === 'insertFromComposition' || inputType === 'insertCompositionText'
      || inputType === 'insertText' && !pending && event.data != null && event.value === this.current)) {
      if (event.value === this.current) return true // Firefox may deliver a duplicate trailing input.
      return this.updateComposition(this.completedComposition, event.value)
    }
    this.completedComposition = undefined
    const before = pending?.snapshot ?? this.selections.get(event.hostId)
    if (before && before.value !== this.current) return false
    if (event.value === this.current && !pending) return true
    const operation = before ? nativeOperation(before, event, pending?.origin ?? 'input-caret') : undefined
    if (!this.append(operation ?? inferSheetCellTextEdit(this.current, event.value, inputType))) return false
    this.selections.set(event.hostId, { ...event })
    return true
  }

  private append(operation: SheetCellTextEditOperation): boolean {
    const next = applySheetCellTextEdit(this.current, operation)
    if (next === undefined) return false
    this.journal = [...this.journal.slice(0, this.cursor), operation]
    this.cursor++; this.current = next
    return true
  }

  private updateComposition(draft: CompositionDraft, value: string): boolean {
    const located = rangeOperation(draft.before, value, draft.start, draft.end, 'insertCompositionText', draft.origin ?? 'composition')
    const operation = located ?? inferSheetCellTextEdit(draft.before, value, 'insertCompositionText')
    if (!located) { draft.start = operation.start; draft.end = operation.end; draft.origin = 'compatibility' }
    this.journal = [...draft.journal.slice(0, draft.cursor), operation]
    this.cursor = draft.cursor + 1; this.current = value
    return true
  }

  private restoreNativeHistory(value: string, direction: -1 | 1): boolean {
    let candidate = this.source
    let match: number | undefined = direction < 0 && value === candidate && this.cursor > 0 ? 0 : undefined
    for (let index = 0; index < this.journal.length; index++) {
      const next = applySheetCellTextEdit(candidate, this.journal[index]!)
      if (next === undefined) return false
      candidate = next
      const cursor = index + 1
      if (candidate === value && (direction < 0 ? cursor < this.cursor : cursor > this.cursor)) {
        match = cursor
        if (direction > 0) break
      }
    }
    if (match === undefined) {
      // A native browser history predating this session is not an edit against its captured source.
      return value === this.current
    }
    this.cursor = match; this.current = value
    return true
  }
}
