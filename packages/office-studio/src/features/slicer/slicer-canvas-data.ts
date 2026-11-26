import type { ISlicerDisplayMemberWindow, ISlicerSourceMetadata, SlicerDisplayMember } from './slicer-member-query.ts'

export interface SlicerCanvasDataOwner {
  instance: object
  model?: object
  unitId: string
  sheetId: string
  revision: number
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
}
interface Page { items: SlicerDisplayMember[]; used: number; bytes: number }
interface Entry {
  id: string
  loader: SlicerCanvasDataLoader
  kind: 'slicer' | 'timeline'
  pages: Map<number, Page>
  total?: number
  metadata?: ISlicerSourceMetadata
  metadataReady: boolean
  pending: Set<string>
  error?: string
}
interface Task { id: string; token: string; generation: number; controller: AbortController; work(): Promise<void> }
export interface SlicerCanvasDataOptions { pageSize?: number; maxConcurrent?: number; maxCachedMembers?: number; memoryBudgetBytes?: number; onChange?(): void }

/** Derived windows have an owner/revision, a bounded queue, and independent cancellation; they never author content or history. */
export class SlicerCanvasDataCache {
  private readonly entries = new Map<string, Entry>()
  private readonly queue: Task[] = []
  private readonly running = new Set<Task>()
  private readonly pageSize: number
  private readonly maxConcurrent: number
  private readonly maxCachedMembers: number
  private readonly memoryBudgetBytes: number
  private readonly onChange?: () => void
  private owner?: SlicerCanvasDataOwner
  private generation = 0
  private clock = 0
  private disposed = false

  constructor(options: SlicerCanvasDataOptions = {}) {
    this.pageSize = options.pageSize ?? 200
    this.maxConcurrent = options.maxConcurrent ?? 2
    this.maxCachedMembers = options.maxCachedMembers ?? 20_000
    this.memoryBudgetBytes = options.memoryBudgetBytes ?? 16 * 1024 * 1024
    this.onChange = options.onChange
    if (![this.pageSize, this.maxConcurrent, this.maxCachedMembers, this.memoryBudgetBytes].every(value => Number.isSafeInteger(value) && value > 0)
      || this.pageSize > 200 || this.maxCachedMembers < this.pageSize) throw new RangeError('切片器画布缓存预算无效。')
  }

  reset(owner?: SlicerCanvasDataOwner): void {
    if (this.disposed) return
    if (owner && this.owner && owner.instance === this.owner.instance && owner.model === this.owner.model && owner.unitId === this.owner.unitId && owner.sheetId === this.owner.sheetId
      && owner.revision === this.owner.revision && owner.userId === this.owner.userId && owner.locale === this.owner.locale) {
      this.owner = owner; return
    }
    this.generation += 1
    for (const task of this.queue) task.controller.abort()
    this.queue.length = 0
    for (const task of this.running) task.controller.abort()
    this.entries.clear(); this.owner = owner; this.onChange?.()
  }

  retain(ids: ReadonlySet<string>): void {
    for (const id of this.entries.keys()) if (!ids.has(id)) this.remove(id)
  }
  bind(id: string, kind: 'slicer' | 'timeline', loader: SlicerCanvasDataLoader): void {
    if (!this.isCurrent()) return
    const existing = this.entries.get(id)
    if (existing && existing.kind !== kind) this.remove(id)
    const entry = this.entries.get(id)
    if (entry) entry.loader = loader
    else this.entries.set(id, { id, loader, kind, pages: new Map(), metadataReady: false, pending: new Set() })
  }
  state(id: string): SlicerCanvasDataState {
    if (!this.isCurrent()) return { pending: false }
    const entry = this.entries.get(id)
    if (entry) {
      try { entry.loader.assertCurrent?.() }
      catch (error) { return { pending: false, error: error instanceof Error ? error.message : '切片器来源已不可读取。' } }
    }
    return entry ? { total: entry.total, metadata: entry.metadata, pending: entry.pending.size > 0, error: entry.error } : { pending: false }
  }

  read(id: string, start: number, count: number): SlicerCanvasMemberWindow | undefined {
    if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(count) || count <= 0 || !Number.isSafeInteger(start + count)) {
      throw new RangeError('切片器画布读取窗口无效。')
    }
    if (!this.isCurrent()) return undefined
    const entry = this.entries.get(id)
    if (!entry || entry.total === undefined) return undefined
    try { entry.loader.assertCurrent?.() }
    catch { return undefined }
    const end = Math.min(entry.total, start + count), items: SlicerDisplayMember[] = []
    for (let index = start; index < end; index++) {
      const page = entry.pages.get(Math.floor(index / this.pageSize) * this.pageSize), item = page?.items[index % this.pageSize]
      if (!page || !item) return undefined
      page.used = ++this.clock; items.push(item)
    }
    return { start, total: entry.total, items }
  }

  request(id: string, start = 0, count = this.pageSize): void {
    const entry = this.entries.get(id)
    if (!entry || !this.isCurrent() || entry.error) return
    if (entry.kind === 'timeline') {
      if (!entry.metadataReady && !entry.pending.has('metadata')) this.schedule(id, entry, 'metadata', async signal => {
        const metadata = await entry.loader.metadata(signal)
        this.check(entry, signal); entry.metadata = metadata; entry.metadataReady = true
      })
      return
    }
    if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(count) || count <= 0
      || !Number.isSafeInteger(start + count) || count > this.maxCachedMembers) {
      entry.error = '当前切片器可见成员数量超过画布缓存预算，请调整按钮高度或列数。'; this.onChange?.(); return
    }
    const end = entry.total === undefined ? start + count : Math.min(entry.total, start + count)
    if (Math.ceil((Math.max(0, end - start) + start % this.pageSize) / this.pageSize) * this.pageSize > this.maxCachedMembers) {
      entry.error = '当前切片器可见成员数量超过画布缓存预算，请调整按钮高度或列数。'; this.onChange?.(); return
    }
    for (let offset = Math.floor(start / this.pageSize) * this.pageSize; offset < end; offset += this.pageSize) {
      const token = `members:${offset}`
      if (entry.pages.has(offset) || entry.pending.has(token)) continue
      this.schedule(id, entry, token, async signal => {
        const window = await entry.loader.members(offset, this.pageSize, signal)
        this.check(entry, signal)
        if (!Number.isSafeInteger(window.total) || window.total < 0 || window.items.length > this.pageSize
          || window.items.length !== Math.max(0, Math.min(this.pageSize, window.total - offset))) throw new Error('切片器成员窗口不完整。')
        if (entry.total !== undefined && entry.total !== window.total) throw new Error('切片器成员来源在查询期间已改变。')
        const bytes = window.items.reduce((sum, item) => sum + 64 + item.key.length * 2 + item.label.length * 2
          + (item.value.type === 'string' || item.value.type === 'error' ? item.value.value.length * 2 : 8), 0)
        this.reserve(window.items.length, bytes)
        entry.total = window.total; entry.pages.set(offset, { items: window.items, used: ++this.clock, bytes })
      })
    }
  }

  retry(id: string): void {
    const entry = this.entries.get(id)
    if (!entry) return
    const { kind, loader } = entry
    this.remove(id); this.bind(id, kind, loader); this.onChange?.()
  }
  dispose(): void { if (this.disposed) return; this.reset(); this.disposed = true }

  private remove(id: string): void {
    for (let index = this.queue.length - 1; index >= 0; index--) if (this.queue[index]!.id === id) {
      this.queue[index]!.controller.abort(); this.queue.splice(index, 1)
    }
    for (const task of this.running) if (task.id === id) task.controller.abort()
    this.entries.delete(id)
  }
  private check(entry: Entry, signal: AbortSignal): void {
    if (this.disposed || signal.aborted || !this.owner || this.entries.get(entry.id) !== entry) throw new Error('切片器画布查询已取消。')
    this.owner.assertCurrent()
    entry.loader.assertCurrent?.()
  }
  private isCurrent(): boolean {
    if (this.disposed || !this.owner) return false
    try { this.owner.assertCurrent(); return true }
    catch { return false }
  }
  private reserve(members: number, bytes: number): void {
    if (bytes > this.memoryBudgetBytes || members > this.maxCachedMembers) throw new Error('切片器成员窗口超过画布缓存预算。')
    const pages = [...this.entries.values()].flatMap(entry => [...entry.pages].map(([start, page]) => ({ entry, start, page }))),
      totals = pages.reduce((value, item) => ({ members: value.members + item.page.items.length, bytes: value.bytes + item.page.bytes }), { members, bytes })
    pages.sort((first, second) => first.page.used - second.page.used)
    for (const item of pages) {
      if (totals.members <= this.maxCachedMembers && totals.bytes <= this.memoryBudgetBytes) break
      item.entry.pages.delete(item.start); totals.members -= item.page.items.length; totals.bytes -= item.page.bytes
    }
  }
  private schedule(id: string, entry: Entry, token: string, work: (signal: AbortSignal) => Promise<void>): void {
    if (this.queue.length >= Math.ceil(this.maxCachedMembers / this.pageSize) + this.entries.size) {
      entry.error = '切片器成员加载队列已满，请重试。'; this.onChange?.(); return
    }
    const generation = this.generation, controller = new AbortController()
    entry.pending.add(token)
    const task: Task = { id, token, generation, controller, work: async () => {
      this.check(entry, controller.signal); await work(controller.signal); this.check(entry, controller.signal)
    } }
    this.queue.push(task); this.onChange?.(); this.pump()
  }
  private pump(): void {
    while (!this.disposed && this.running.size < this.maxConcurrent && this.queue.length) {
      const task = this.queue.shift()!, entry = this.entries.get(task.id)
      if (!entry || task.generation !== this.generation || task.controller.signal.aborted) continue
      this.running.add(task)
      void task.work().catch(error => {
        if (!task.controller.signal.aborted && task.generation === this.generation && this.entries.get(task.id) === entry)
          entry.error = error instanceof Error ? error.message : '切片器成员加载失败。'
      }).finally(() => {
        this.running.delete(task)
        if (task.generation === this.generation && this.entries.get(task.id) === entry) { entry.pending.delete(task.token); this.onChange?.() }
        this.pump()
      })
    }
  }
}
