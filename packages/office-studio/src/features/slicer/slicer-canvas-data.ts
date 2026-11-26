import type { ISlicerDisplayMemberWindow, ISlicerSourceMetadata, SlicerDisplayMember } from './slicer-member-query.ts'
import { assertSlicerCanvasWindow, copySlicerCanvasMember, copySlicerCanvasMetadata, readSlicerCanvasMetadata,
  readSlicerCanvasWindow, SlicerCanvasDataError } from './slicer-canvas-data-validation.ts'

export interface SlicerCanvasDataOwner {
  instance: object
  model?: object
  unitId: string
  sheetId: string
  revision: number
  historyGeneration: number
  userId: string
  locale: string
  assertCurrent(): void
}
export interface SlicerCanvasDataLoader {
  members(start: number, count: number, signal: AbortSignal): Promise<ISlicerDisplayMemberWindow>
  metadata(signal: AbortSignal): Promise<ISlicerSourceMetadata>
  assertCurrent?(): void
}
export interface SlicerCanvasMemberWindow { start: number; total: number; items: SlicerDisplayMember[] }
export interface SlicerCanvasDataState {
  total?: number
  metadata?: ISlicerSourceMetadata
  pending: boolean
  error?: string
  errorCode?: string
}
interface WindowRequest { start: number; count: number }
interface Page { items: SlicerDisplayMember[]; used: number; bytes: number }
interface Entry {
  readonly id: string
  readonly loader: SlicerCanvasDataLoader
  readonly kind: 'slicer' | 'timeline'
  pages: Map<number, Page>
  captionPages: Map<string, number>
  total?: number
  metadata?: ISlicerSourceMetadata
  metadataBytes: number
  window?: WindowRequest
  pending: Map<string, Task>
  error?: string
  errorCode?: string
}
interface Task {
  readonly entry: Entry
  readonly owner: SlicerCanvasDataOwner
  readonly loader: SlicerCanvasDataLoader
  readonly token: string
  readonly generation: number
  readonly controller: AbortController
  readonly start?: number
  active: boolean
  timer?: ReturnType<typeof setTimeout>
}
export interface SlicerCanvasDataOptions {
  pageSize?: number
  maxConcurrent?: number
  maxCachedMembers?: number
  memoryBudgetBytes?: number
  maxQueuedTasks?: number
  timeoutMs?: number
  onChange?(): void
}

export function sameSlicerCanvasDataOwner(first?: SlicerCanvasDataOwner, second?: SlicerCanvasDataOwner): boolean {
  return first === second || Boolean(first && second && first.instance === second.instance && first.model === second.model
    && first.unitId === second.unitId && first.sheetId === second.sheetId && first.revision === second.revision
    && first.historyGeneration === second.historyGeneration && first.userId === second.userId && first.locale === second.locale)
}

/** Presentation DTOs only. Filtering and membership remain authoritative in the shared Worker/Rust query service. */
export class SlicerCanvasDataCache {
  private readonly entries = new Map<string, Entry>()
  private readonly queue: Task[] = []
  private readonly running = new Set<Task>()
  private readonly pageSize: number
  private readonly maxConcurrent: number
  private readonly maxCachedMembers: number
  private readonly memoryBudgetBytes: number
  private readonly maxQueuedTasks: number
  private readonly timeoutMs: number
  private readonly onChange?: () => void
  private owner?: SlicerCanvasDataOwner
  private generation = 0
  private clock = 0
  private cachedMembers = 0
  private cachedBytes = 0
  private disposed = false

  constructor(options: SlicerCanvasDataOptions = {}) {
    this.pageSize = options.pageSize ?? 200
    this.maxConcurrent = options.maxConcurrent ?? 2
    this.maxCachedMembers = options.maxCachedMembers ?? 20_000
    this.memoryBudgetBytes = options.memoryBudgetBytes ?? 16 * 1024 * 1024
    this.maxQueuedTasks = options.maxQueuedTasks ?? Math.ceil(this.maxCachedMembers / this.pageSize) + this.maxConcurrent
    this.timeoutMs = options.timeoutMs ?? 30_000
    this.onChange = options.onChange
    if (![this.pageSize, this.maxConcurrent, this.maxCachedMembers, this.memoryBudgetBytes, this.maxQueuedTasks].every(value => Number.isSafeInteger(value) && value > 0)
      || this.pageSize > 200 || this.maxConcurrent > 32 || this.maxCachedMembers < this.pageSize
      || this.memoryBudgetBytes > 0xffff_ffff || !Number.isFinite(this.timeoutMs) || this.timeoutMs <= 0 || this.timeoutMs > 2_147_483_647) {
      throw new RangeError('切片器画布缓存、队列或超时预算无效。')
    }
  }

  reset(owner?: SlicerCanvasDataOwner): void {
    if (this.disposed) return
    if (owner && (!Number.isSafeInteger(owner.revision) || owner.revision < 0 || !Number.isSafeInteger(owner.historyGeneration)
      || owner.historyGeneration < 0 || !owner.unitId || !owner.sheetId || !owner.locale || typeof owner.assertCurrent !== 'function')) {
      throw new TypeError('切片器画布须绑定完整的文档、历史代次和来源上下文。')
    }
    const captured = owner ? { ...owner } : undefined
    if (sameSlicerCanvasDataOwner(captured, this.owner)) { this.owner = captured; return }
    this.generation += 1
    for (const entry of [...this.entries.values()]) this.removeEntry(entry)
    this.owner = captured; this.onChange?.()
  }

  retain(ids: ReadonlySet<string>): void {
    let changed = false
    for (const entry of [...this.entries.values()]) if (!ids.has(entry.id)) { this.removeEntry(entry); changed = true }
    if (changed) { this.onChange?.(); this.pump() }
  }
  bind(id: string, kind: 'slicer' | 'timeline', loader: SlicerCanvasDataLoader): void {
    if (!this.isCurrent()) return
    if (!id || (kind !== 'slicer' && kind !== 'timeline') || typeof loader.members !== 'function' || typeof loader.metadata !== 'function') {
      throw new TypeError('切片器画布来源绑定无效。')
    }
    const previous = this.entries.get(id)
    if (previous?.kind === kind && previous.loader === loader) return
    if (previous) this.removeEntry(previous)
    this.entries.set(id, { id, kind, loader, pages: new Map(), captionPages: new Map(), metadataBytes: 0, pending: new Map() })
    this.onChange?.(); this.pump()
  }
  state(id: string): SlicerCanvasDataState {
    if (!this.isCurrent()) return { pending: false }
    const entry = this.entries.get(id)
    if (!entry) return { pending: false }
    this.checkReadable(entry)
    return { total: entry.total, ...(entry.metadata ? { metadata: copySlicerCanvasMetadata(entry.metadata) } : {}),
      pending: entry.pending.size > 0, error: entry.error, errorCode: entry.errorCode }
  }

  read(id: string, start: number, count: number): SlicerCanvasMemberWindow | undefined {
    assertSlicerCanvasWindow(start, count, this.maxCachedMembers)
    if (!this.isCurrent()) return undefined
    const entry = this.entries.get(id)
    if (!entry || !this.checkReadable(entry) || entry.error || entry.total === undefined) return undefined
    const end = Math.min(entry.total, start + count), items: SlicerDisplayMember[] = []
    for (let index = start; index < end; index++) {
      const page = entry.pages.get(Math.floor(index / this.pageSize) * this.pageSize), item = page?.items[index % this.pageSize]
      if (!page || !item) return undefined
      page.used = ++this.clock; items.push(copySlicerCanvasMember(item))
    }
    return { start, total: entry.total, items }
  }

  request(id: string, start = 0, count = this.pageSize): void {
    if (!this.isCurrent()) return
    const entry = this.entries.get(id)
    if (!entry || !this.checkReadable(entry) || entry.error) return
    if (entry.kind === 'timeline') {
      if (!entry.metadata && !entry.pending.has('metadata')) this.schedule(entry, 'metadata')
      this.pump(); return
    }
    try { assertSlicerCanvasWindow(start, count, this.maxCachedMembers) }
    catch (error) { this.failEntry(entry, error); return }
    const end = entry.total === undefined ? start + count : Math.min(entry.total, start + count)
    if (Math.ceil((Math.max(0, end - start) + start % this.pageSize) / this.pageSize) * this.pageSize > this.maxCachedMembers) {
      this.failEntry(entry, new SlicerCanvasDataError('SLICER_CANVAS_MEMORY_LIMIT', '当前切片器可见成员数量超过画布缓存预算，请调整按钮高度或列数。')); return
    }
    const replaced = entry.window?.start !== start || entry.window?.count !== count, beforePending = entry.pending.size
    entry.window = { start, count }
    this.cancelOutsideWindow(entry)
    for (let offset = Math.floor(start / this.pageSize) * this.pageSize; offset < end; offset += this.pageSize) {
      const token = `members:${offset}`
      const page = entry.pages.get(offset)
      if (page) { page.used = ++this.clock; continue }
      if (!entry.pending.has(token) && !this.schedule(entry, token, offset)) break
    }
    if (replaced) {
      const foreground = this.queue.filter(task => task.entry === entry), background = this.queue.filter(task => task.entry !== entry)
      this.queue.splice(0, this.queue.length, ...foreground, ...background)
    }
    if (replaced || beforePending !== entry.pending.size) this.onChange?.()
    this.pump()
  }

  /** Retain the last explicit window; a new entry identity rejects old retry results with the same page token. */
  retry(id: string): void {
    if (!this.isCurrent()) return
    const entry = this.entries.get(id)
    if (!entry) return
    const { kind, loader, window } = entry
    this.removeEntry(entry); this.bind(id, kind, loader)
    this.request(id, window?.start, window?.count)
  }
  dispose(): void { if (this.disposed) return; this.reset(); this.disposed = true }

  private removeEntry(entry: Entry): void {
    this.entries.delete(entry.id)
    for (const task of [...entry.pending.values()]) this.cancel(task)
    this.dropData(entry)
  }
  private dropData(entry: Entry): void {
    for (const page of entry.pages.values()) { this.cachedMembers -= page.items.length; this.cachedBytes -= page.bytes }
    entry.pages.clear(); entry.captionPages.clear(); this.cachedBytes -= entry.metadataBytes
    entry.metadata = undefined; entry.metadataBytes = 0; entry.total = undefined
  }
  private cancel(task: Task): void {
    if (!task.active) return
    this.detach(task)
    task.controller.abort(new SlicerCanvasDataError('SLICER_CANVAS_CANCELLED', '切片器画布查询已取消。'))
  }
  private detach(task: Task): void {
    task.active = false; clearTimeout(task.timer); task.timer = undefined
    this.running.delete(task)
    const index = this.queue.indexOf(task)
    if (index >= 0) this.queue.splice(index, 1)
    if (task.entry.pending.get(task.token) === task) task.entry.pending.delete(task.token)
  }
  private wanted(entry: Entry, start: number): boolean {
    const window = entry.window
    if (!window) return false
    const end = Math.min(window.start + window.count, entry.total ?? Number.MAX_SAFE_INTEGER)
    return start < end && start + this.pageSize > window.start
  }
  private cancelOutsideWindow(entry: Entry): void {
    for (const task of [...entry.pending.values()]) if (task.start !== undefined && !this.wanted(entry, task.start)) this.cancel(task)
  }
  private check(task: Task): void {
    if (!task.active || task.controller.signal.aborted || task.generation !== this.generation || this.disposed
      || !this.owner || this.entries.get(task.entry.id) !== task.entry || task.entry.pending.get(task.token) !== task
      || !sameSlicerCanvasDataOwner(task.owner, this.owner)) {
      throw new SlicerCanvasDataError('SLICER_CANVAS_CANCELLED', '切片器画布查询已取消。')
    }
    task.owner.assertCurrent(); task.loader.assertCurrent?.()
  }
  private checkReadable(entry: Entry): boolean {
    try { entry.loader.assertCurrent?.(); return true }
    catch (error) { this.failEntry(entry, error); return false }
  }
  private isCurrent(): boolean {
    if (this.disposed || !this.owner) return false
    try { this.owner.assertCurrent(); return true }
    catch { this.reset(); return false }
  }
  private reserve(members: number, bytes: number): void {
    if (bytes > this.memoryBudgetBytes || members > this.maxCachedMembers) {
      throw new SlicerCanvasDataError('SLICER_CANVAS_MEMORY_LIMIT', '切片器窗口超过画布缓存预算。')
    }
    const candidates = [...this.entries.values()].flatMap(entry => [...entry.pages].filter(([start]) => !this.wanted(entry, start))
      .map(([start, page]) => ({ entry, start, page }))).sort((first, second) => first.page.used - second.page.used)
    let totalMembers = this.cachedMembers + members, totalBytes = this.cachedBytes + bytes
    const evictions: typeof candidates = []
    for (const item of candidates) {
      if (totalMembers <= this.maxCachedMembers && totalBytes <= this.memoryBudgetBytes) break
      evictions.push(item); totalMembers -= item.page.items.length; totalBytes -= item.page.bytes
    }
    if (totalMembers > this.maxCachedMembers || totalBytes > this.memoryBudgetBytes) {
      throw new SlicerCanvasDataError('SLICER_CANVAS_MEMORY_LIMIT', '当前切片器可见窗口和时间线信息超过共享画布缓存预算。')
    }
    for (const { entry, start, page } of evictions) {
      entry.pages.delete(start); this.cachedMembers -= page.items.length; this.cachedBytes -= page.bytes
      for (const item of page.items) entry.captionPages.delete(item.label)
    }
  }
  private schedule(entry: Entry, token: string, start?: number): boolean {
    if (!this.owner) return false
    if (start !== undefined && !Number.isSafeInteger(start + this.pageSize)) {
      this.failEntry(entry, new SlicerCanvasDataError('SLICER_CANVAS_INVALID_WINDOW', '切片器画布分页位置超过安全整数范围。')); return false
    }
    if (this.queue.length >= this.maxQueuedTasks) {
      this.failEntry(entry, new SlicerCanvasDataError('SLICER_CANVAS_QUEUE_FULL', '切片器成员加载队列已满，请重试。')); return false
    }
    const loader: SlicerCanvasDataLoader = { members: entry.loader.members.bind(entry.loader), metadata: entry.loader.metadata.bind(entry.loader),
      ...(entry.loader.assertCurrent ? { assertCurrent: entry.loader.assertCurrent.bind(entry.loader) } : {}) }
    const task: Task = { entry, loader, token, start, generation: this.generation, owner: this.owner, controller: new AbortController(), active: true }
    entry.pending.set(token, task); this.queue.push(task)
    task.timer = setTimeout(() => {
      if (!task.active || task.generation !== this.generation || this.entries.get(entry.id) !== entry) return
      this.failEntry(entry, new SlicerCanvasDataError('SLICER_CANVAS_TIMEOUT', '切片器画布查询已超时，请重试。'))
    }, this.timeoutMs)
    this.onChange?.(); return true
  }
  private async load(task: Task): Promise<void> {
    this.check(task)
    const entry = task.entry, signal = task.controller.signal
    if (task.start === undefined) {
      const raw = await task.loader.metadata(signal)
      this.check(task)
      const { metadata, bytes } = readSlicerCanvasMetadata(raw, this.memoryBudgetBytes)
      if (!metadata.timeline) throw new SlicerCanvasDataError('SLICER_CANVAS_INVALID_RESULT', '切片器时间线来源信息缺少日期范围状态。')
      this.reserve(0, bytes); this.check(task)
      entry.metadata = metadata; entry.metadataBytes = bytes; this.cachedBytes += bytes
    } else {
      const raw = await task.loader.members(task.start, this.pageSize, signal)
      this.check(task)
      const { window, bytes } = readSlicerCanvasWindow(raw, task.start, this.pageSize, this.memoryBudgetBytes)
      if (entry.total !== undefined && entry.total !== window.total) {
        throw new SlicerCanvasDataError('SLICER_CANVAS_INVALID_RESULT', '切片器成员来源在查询期间已改变，请重新读取。')
      }
      if (window.items.some(item => entry.captionPages.has(item.label))) {
        throw new SlicerCanvasDataError('SLICER_CANVAS_INVALID_RESULT', '切片器成员在不同窗口中重复，请重新读取。')
      }
      this.reserve(window.items.length, bytes); this.check(task)
      entry.total = window.total; entry.pages.set(task.start, { items: window.items, used: ++this.clock, bytes })
      for (const item of window.items) entry.captionPages.set(item.label, task.start)
      this.cachedMembers += window.items.length; this.cachedBytes += bytes
      this.cancelOutsideWindow(entry)
    }
    this.onChange?.()
  }
  private failEntry(entry: Entry, error: unknown): void {
    if (this.entries.get(entry.id) !== entry) return
    const message = error instanceof Error ? error.message : '切片器画布数据加载失败。'
    const code = error instanceof SlicerCanvasDataError ? error.code
      : error !== null && typeof error === 'object' && 'code' in error && typeof error.code === 'string' ? error.code : undefined
    const changed = entry.error !== message || entry.errorCode !== code || entry.pending.size > 0 || entry.pages.size > 0 || Boolean(entry.metadata)
    entry.error = message; entry.errorCode = code
    for (const task of [...entry.pending.values()]) this.cancel(task)
    this.dropData(entry)
    if (changed) this.onChange?.()
    this.pump()
  }
  private pump(): void {
    while (this.isCurrent() && this.running.size < this.maxConcurrent && this.queue.length) {
      const task = this.queue.shift()!
      if (!task.active || task.generation !== this.generation || this.entries.get(task.entry.id) !== task.entry) { this.cancel(task); continue }
      this.running.add(task)
      void this.load(task).catch(error => {
        if (task.active && task.generation === this.generation && this.entries.get(task.entry.id) === task.entry && this.isCurrent()) {
          this.failEntry(task.entry, error)
        }
      }).finally(() => {
        if (task.active) { this.detach(task); this.onChange?.() }
        this.pump()
      })
    }
  }
}
