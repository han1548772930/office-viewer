import type { SheetImageResource } from '../../domain/sheet-images.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { prepareSheetImageFiles, prepareSheetImageResources, prepareSheetImageUrl } from './sheet-image-io.ts'
import { normalizeSheetImageResource, SHEET_IMAGE_MAX_BATCH_BYTES, SHEET_IMAGE_MAX_COUNT, sheetImageRecord } from './sheet-image-model.ts'
import { SHEET_DRAWING_PLUGIN, type SheetImageSourceType } from './sheet-image-public.ts'

export interface SheetImageInputPreparationOptions {
  signal?: AbortSignal
  current?: () => boolean
  progress?: (completed: number, total: number) => void
  /** UUIDs require actual media bytes from a registered resource owner. */
  resolveResource?: (source: string, sourceType: SheetImageSourceType, signal?: AbortSignal) => Promise<SheetImageResource | File | string | null>
}
const object = (value: unknown): Record<string, unknown> | undefined => value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : undefined
interface ImageJob { source: string; sourceType: SheetImageSourceType; name: string; resource?: SheetImageResource; apply(resource: SheetImageResource): void }

/** Reads external sources before creation. The caller's snapshot is never modified or partially published. */
export async function prepareWorkbookSheetImageInput<T extends Record<string, unknown>>(input: T, options: SheetImageInputPreparationOptions = {}): Promise<T> {
  const next = cloneModel(input), jobs: ImageJob[] = [], resourceTables: Array<{ entry: Record<string, unknown>; table: Record<string, unknown> }> = []
  const assertCurrent = () => {
    if (options.signal?.aborted || options.current && !options.current()) throw new DOMException('工作簿图片读取已取消。', 'AbortError')
  }
  assertCurrent()
  function enqueueResource(owner: Record<string, unknown>): void {
    const raw = sheetImageRecord(owner.resource, '图片资源')
    if (typeof raw.source !== 'string') throw new TypeError('图片资源缺少来源。')
    const resource = /^data:/i.test(raw.source) ? normalizeSheetImageResource(raw) : undefined
    jobs.push({ source: raw.source, sourceType: /^data:/i.test(raw.source) ? 'BASE64' : 'URL', name: typeof raw.name === 'string' ? raw.name : 'image', resource,
      apply: actual => { owner.resource = { ...cloneModel(raw), ...actual, id: typeof raw.id === 'string' ? raw.id : actual.id, metadata: cloneModel(object(raw.metadata) ?? {}) } } })
  }
  function enqueuePublic(owner: Record<string, unknown>, background = false, internal = false): void {
    const extension = object(owner.officeSheetImage), rawResource = object(extension?.resource ?? owner.officeSheetImageResource)
    const source = typeof owner.source === 'string' ? owner.source : rawResource?.source
    if (typeof source !== 'string') throw new TypeError('公开图片缺少来源。')
    const sourceType = owner.imageSourceType === 'UUID' ? 'UUID' : /^data:/i.test(source) ? 'BASE64' : 'URL'
    const resource = rawResource && rawResource.source === source && /^data:/i.test(source) ? normalizeSheetImageResource(rawResource) : undefined
    jobs.push({ source, sourceType, name: typeof owner.title === 'string' ? owner.title : background ? 'background' : 'image', resource,
      apply: actual => {
        owner.source = actual.source; owner.imageSourceType = 'BASE64'
        if (background && internal) {
          owner.resource = actual
          delete owner.source; delete owner.imageSourceType; delete owner.officeSheetImageResource
        } else if (background) owner.officeSheetImageResource = actual
        else owner.officeSheetImage = { ...extension, resource: actual }
      } })
  }
  function bodyImages(raw: unknown): void {
    const body = object(raw), drawings = object(body?.drawings)
    for (const drawing of Object.values(drawings ?? {})) if (object(drawing)?.drawingType === 0) enqueuePublic(object(drawing)!)
  }
  const sheets = Array.isArray(next.sheets) ? next.sheets : Object.values(object(next.sheets) ?? {})
  for (const value of sheets) {
    const sheet = sheetImageRecord(value, '工作表'), background = object(sheet.backgroundImage)
    if (background) background.resource === undefined ? enqueuePublic(background, true, next.kind === 'workbook') : enqueueResource(background)
    for (const image of Array.isArray(sheet.images) ? sheet.images : []) enqueueResource(sheetImageRecord(image))
    for (const cell of Object.values(object(sheet.cells) ?? {})) {
      const image = object(object(cell)?.image)
      if (image) enqueueResource(image)
    }
    for (const columns of Object.values(object(sheet.cellData) ?? {})) for (const cell of Object.values(object(columns) ?? {})) bodyImages(object(cell)?.p)
  }
  for (const value of Array.isArray(next.resources) ? next.resources : []) {
    const entry = sheetImageRecord(value, '工作簿资源')
    if (entry.name !== SHEET_DRAWING_PLUGIN) continue
    if (typeof entry.data !== 'string' || entry.data.length > SHEET_IMAGE_MAX_BATCH_BYTES * 3) throw new RangeError('工作簿绘图资源超过读取预算或格式无效。')
    const table = sheetImageRecord(entry.data ? JSON.parse(entry.data) : {}, '公开绘图资源')
    resourceTables.push({ entry, table })
    for (const unit of Object.values(table)) for (const drawing of Object.values(object(object(unit)?.data) ?? {})) {
      if (object(drawing)?.drawingType === 0) enqueuePublic(object(drawing)!)
    }
  }
  if (jobs.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('一次最多准备 1000 张图片。')
  const fetched = new Map<string, SheetImageResource>()
  let bytes = 0, completed = 0
  for (const job of jobs) {
    assertCurrent()
    const key = `${job.sourceType}\0${job.source}`
    let actual = fetched.get(key)
    if (!actual) {
      if (job.resource) actual = (await prepareSheetImageResources([job.resource], { signal: options.signal, current: options.current }))[0]!
      else if (job.sourceType === 'UUID') {
        const resolved = await options.resolveResource?.(job.source, job.sourceType, options.signal)
        if (!resolved) throw new Error('找不到 UUID 图片的实际媒体资源。')
        actual = typeof resolved === 'string' ? await prepareSheetImageUrl(resolved, { signal: options.signal, name: job.name })
          : resolved instanceof File ? (await prepareSheetImageFiles([resolved], options))[0]!
            : (await prepareSheetImageResources([resolved], options))[0]!
      } else actual = await prepareSheetImageUrl(job.source, { signal: options.signal, name: job.name })
      fetched.set(key, actual)
    }
    assertCurrent()
    bytes += actual.byteLength
    if (bytes > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('工作簿图片批次超过 32 MiB。')
    job.apply(cloneModel(actual))
    options.progress?.(++completed, jobs.length)
  }
  assertCurrent()
  for (const { entry, table } of resourceTables) entry.data = JSON.stringify(table)
  return next
}
