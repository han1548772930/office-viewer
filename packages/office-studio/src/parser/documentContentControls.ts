import type { DocumentBlock, DocumentContentControl, DocumentModel } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import type { NativeDocumentContentControlDto } from './contracts'
import { array, record } from './validation'
import { documentStoredBlockHosts, sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { documentTableCellContentContainers } from '../features/document/model/document-table-continuations'
import { runText, runsForTarget, tableForBlock, textTargets } from '../features/document/model/document-model'
import { isDocumentTextBoundary } from '../features/document/model/document-story'
import { documentContentControlContains, documentContentControlOwnedRowBlockIds, documentContentControlTableRowBlockIds, resolveDocumentContentControlTableRows } from '../features/document/model/content-control-table-rows'
import { assertDocumentContentControlBlockRange, documentBlockContentControlContains, documentContentControlBlockTextTargets, documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from '../features/document/model/content-control-block-range'
import { validateContentControlPropertyPatch } from '../features/document/services/content-control-properties'
import type { ContentControlPatch } from '../features/document/services/content-controls'
import { parsedDocumentTextProjection } from './documentTextProjection'

interface BlockAddress { block: DocumentBlock; scoped: DocumentModel; container: DocumentBlock[]; position: number; host?: DocumentAuxiliaryBlockHost }
export interface NativeContentControlHydration {
  version: 1
  hydratedCount: number
  /** Exact native source remains persisted when a structural domain mapping is still pending. */
  unmapped: Array<{ reason: string; control: NativeDocumentContentControlDto }>
}

function assertDto(value: unknown): asserts value is NativeDocumentContentControlDto {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('原生内容控件协议必须是对象。')
  const dto = record(value), ranges = dto.textRanges, rows = dto.tableRowRange
  for (const key of ['id', 'partName', 'story', 'sourceXml']) if (typeof dto[key] !== 'string') throw new TypeError('原生内容控件来源字段无效。')
  if (typeof dto.id !== 'string' || !dto.id.trim() || dto.id.length > 4096 || dto.parentId !== undefined && (typeof dto.parentId !== 'string' || !dto.parentId.trim())) throw new TypeError('原生内容控件身份无效。')
  if (!dto.properties || typeof dto.properties !== 'object' || Array.isArray(dto.properties) || !Array.isArray(ranges)) throw new TypeError('原生内容控件属性或文字范围无效。')
  for (const value of ranges) {
    const range = record(value)
    if (typeof range.blockId !== 'string' || !range.blockId.trim() || typeof range.startOffset !== 'number' || !Number.isSafeInteger(range.startOffset)
      || typeof range.endOffset !== 'number' || !Number.isSafeInteger(range.endOffset) || range.startOffset < 0 || range.endOffset < range.startOffset) throw new RangeError('原生控件 UTF-16 范围无效。')
    for (const key of ['startNoteReferenceCount', 'endNoteReferenceCount'] as const) if (range[key] !== undefined && (typeof range[key] !== 'number' || !Number.isSafeInteger(range[key]) || Number(range[key]) < 0)) throw new RangeError('原生控件注释边界序号无效。')
  }
  if (rows !== undefined) {
    const range = record(rows)
    if (typeof range.tableBlockId !== 'string' || !range.tableBlockId.trim() || !Array.isArray(range.rowIndexes) || !range.rowIndexes.length
      || range.rowIndexes.length > 10_000 || range.rowIndexes.some((row, index, indexes) => typeof row !== 'number' || !Number.isSafeInteger(row)
        || row < 0 || index > 0 && row !== indexes[index - 1] + 1)) throw new RangeError('原生控件物理行序列无效。')
  }
  if (dto.blockIds !== undefined) assertDocumentContentControlBlockRange({ blockIds: dto.blockIds })
  if (dto.hasObjectContent !== undefined && typeof dto.hasObjectContent !== 'boolean') throw new TypeError('原生控件对象内容标记无效。')
  if (dto.objectContentCount !== undefined && (!Number.isSafeInteger(dto.objectContentCount) || dto.objectContentCount < 0)) throw new RangeError('原生控件对象数量无效。')
  if (dto.hasObjectContent === false && Number(dto.objectContentCount ?? 0) > 0 || dto.hasObjectContent === true && dto.objectContentCount === 0) throw new TypeError('原生控件对象标记与数量不一致。')
  if (dto.blockIds !== undefined && rows !== undefined) throw new TypeError('原生控件不能同时拥有完整块和物理行范围。')
  if (dto.noteId !== undefined && (typeof dto.noteId !== 'number' || !Number.isSafeInteger(dto.noteId))) throw new TypeError('原生控件 note 身份无效。')
}

function addresses(model: DocumentModel): { source: Map<string, BlockAddress[]>; byBlockId: Map<string, BlockAddress> } {
  const source = new Map<string, BlockAddress[]>(), byBlockId = new Map<string, BlockAddress>()
  const visit = (blocks: DocumentBlock[], scoped: DocumentModel, host?: DocumentAuxiliaryBlockHost) => {
    for (const [position, block] of blocks.entries()) {
      const rawSource = record(block.metadata.source), rawId = typeof rawSource.id === 'string' ? rawSource.id : block.id
      const entry = { block, scoped, container: blocks, position, host }, candidates = source.get(rawId) ?? []
      if (byBlockId.has(block.id)) throw new TypeError('导入内容区的领域块标识重复。')
      byBlockId.set(block.id, entry)
      if (!candidates.some(candidate => candidate.block.id === block.id && sameDocumentBlockHost(candidate.host, host))) candidates.push(entry)
      source.set(rawId, candidates)
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) visit(content, scoped, host)
    }
  }
  visit(model.blocks, model)
  for (const entry of documentStoredBlockHosts(model)) visit(entry.blocks, { ...model, blocks: entry.blocks }, entry.host)
  return { source, byBlockId }
}

function resolveAddress(index: Map<string, BlockAddress[]>, id: string, dto: NativeDocumentContentControlDto): BlockAddress | undefined {
  const candidates = index.get(id) ?? []
  const matching = candidates.filter(candidate => {
    if (dto.story === 'body' && candidate.host) return false
    if (dto.noteId !== undefined) {
      const host = candidate.host
      if (!host || host.kind === 'story' || host.noteKind !== dto.story) return false
      if (host.kind === 'note' && host.noteId !== String(dto.noteId)) {
        const note = (host.noteKind === 'footnote' ? candidate.scoped.footnotes : candidate.scoped.endnotes)?.[host.noteId]
        if (record(note?.metadata?.officeNoteSource).nativeId !== dto.noteId) return false
      } else if (host.kind === 'noteDecoration') {
        const settings = candidate.scoped.noteSettings?.[host.noteKind]
        const content = host.dormantIndex === undefined ? settings?.[host.decoration]
          : array(record(settings?.metadata?.officeDormantNoteDecorations).contents)[host.dormantIndex]
        if (record(record(content).metadata).decorationSource === undefined || record(record(record(content).metadata).decorationSource).noteId !== dto.noteId) return false
      }
    }
    const references = [...array(candidate.block.metadata.sourceRefs), ...array(record(candidate.block.metadata.source).sourceRefs)]
    if (!references.length) return candidates.length === 1
    return references.some(reference => {
    const source = record(reference)
    return source.partName === dto.partName && source.story === dto.story
    })
  })
  return matching.length === 1 ? matching[0] : undefined
}

function scalarValue(control: DocumentContentControl, text: string): string {
  if (control.showingPlaceholder) return ''
  if (control.type === 'checkbox') return (control.checkbox?.checked ?? control.checked) ? 'true' : 'false'
  if (control.type === 'date') return control.dateValue ?? text
  if (control.type === 'dropDownList' || control.type === 'comboBox') return control.items?.find(item => item.displayText === text)?.value ?? text
  return text
}

function physicalParagraphBlocks(address: BlockAddress): DocumentBlock[] {
  const rawId = record(address.block.metadata.source).id ?? address.block.id
  const blocks = [address.block]
  for (let index = address.position + 1; index < address.container.length; index++) {
    const block = address.container[index]!
    if (block.metadata.officeSourceParagraphId !== rawId) break
    blocks.push(block)
  }
  return blocks
}

/** No UI controls or second editor model are constructed during import. */
export function normalizeNativeDocumentContentControls(model: DocumentModel, values: unknown): NativeContentControlHydration | undefined {
  if (values === undefined) return undefined
  if (!Array.isArray(values) || values.length > 100_000) throw new RangeError('原生内容控件数量超过协议预算。')
  const { source: index, byBlockId } = addresses(model), result: NativeContentControlHydration = { version: 1, hydratedCount: 0, unmapped: [] }
  const ids = new Set<string>(), pending = new Map<string, string>(), contexts = new Map<string, BlockAddress>()
  const sources = new Map<string, NativeDocumentContentControlDto>()
  let sourceBytes = 0
  const controls: DocumentContentControl[] = []
  for (const value of values) {
    assertDto(value)
    const dto = cloneModel(value)
    if (ids.has(dto.id)) throw new TypeError('原生内容控件 ID 重复。')
    ids.add(dto.id)
    sources.set(dto.id, dto)
    sourceBytes += dto.sourceXml.length * 2 + JSON.stringify(dto.properties).length * 2
    if (sourceBytes > 64 * 1024 * 1024) throw new RangeError('原生内容控件保留数据超过归一化预算。')
    const { value: _authoredValue, blockRange: _authoredBlockRange, tableRowRange: _authoredRowRange, checked, dateValue, ...rawProperties } = dto.properties
    // One explicit DTO/domain assertion follows the shared runtime property
    // validator. Scalar value/checked/date fields are separately validated below.
    const properties = rawProperties as ContentControlPatch
    validateContentControlPropertyPatch(properties)
    if (!properties.type) throw new TypeError('原生内容控件缺少类型。')
    if (checked !== undefined && typeof checked !== 'boolean' || dateValue !== undefined && typeof dateValue !== 'string') throw new TypeError('原生控件值字段无效。')
    let address: BlockAddress | undefined
    let range: { blockId: string; startOffset: number; endOffset: number } | undefined
    let rowRange: DocumentContentControl['tableRowRange']
    let blockRange: DocumentContentControl['blockRange']
    let reason: string | undefined
    if (dto.tableRowRange) {
      address = resolveAddress(index, dto.tableRowRange.tableBlockId, dto)
      const table = address && tableForBlock(address.block)
      if (!address || !table || dto.tableRowRange.rowIndexes.some(row => row >= table.rows.length)) reason = '原生 row SDT 尚未映射到同一真实表格的完整行序列。'
      else {
        rowRange = { tableId: address.block.id, rowIds: dto.tableRowRange.rowIndexes.map(row => table.rows[row]!.id) }
        const rows = resolveDocumentContentControlTableRows(address.scoped, rowRange), owned = documentContentControlTableRowBlockIds(rows)
        const target = textTargets(address.scoped).find(target => owned.has(target.cellBlockId ?? target.blockId))
        if (!target) reason = 'row SDT 没有可用的真实文字输入锚点。'
        else range = { blockId: target.cellBlockId ?? target.blockId, startOffset: 0, endOffset: runText(runsForTarget(address.scoped, target)).length }
      }
    } else if (dto.blockIds) {
      const blocks = dto.blockIds.map(id => resolveAddress(index, id, dto))
      address = blocks[0]
      if (!address || blocks.some(block => !block || !sameDocumentBlockHost(block.host, address?.host))) reason = '原生 block SDT 尚未映射到同一完整内容区。'
      else {
        blockRange = { blockIds: blocks.flatMap(block => physicalParagraphBlocks(block!).map(block => block.id)) }
        resolveDocumentContentControlBlocks(address.scoped, blockRange)
        const target = documentContentControlBlockTextTargets(address.scoped, blockRange)[0]
        if (!target) reason = 'block SDT 没有可用的真实文字输入锚点。'
        else range = { blockId: target.cellBlockId ?? target.blockId, startOffset: 0, endOffset: runText(runsForTarget(address.scoped, target)).length }
      }
    } else if (dto.textRanges.length === 1) {
      const source = dto.textRanges[0]!
      address = resolveAddress(index, source.blockId, dto)
      if (address && textTargets(address.scoped).some(target => (target.cellBlockId ?? target.blockId) === address!.block.id)) {
        const text = runText(address.block.runs)
        const projection = parsedDocumentTextProjection(address.block)
        const startOffset = projection.mapBoundary(source.startOffset, source.startNoteReferenceCount), endOffset = projection.mapBoundary(source.endOffset, source.endNoteReferenceCount)
        if (startOffset === undefined || endOffset === undefined) {
          if (!projection.changed) throw new RangeError('原生控件文字范围超过真实段落或拆分代理项。')
          reason = '原生控件边界缺少有效的公式或注释引用来源映射；完整来源保留。'
        } else if (!isDocumentTextBoundary(text, startOffset) || !isDocumentTextBoundary(text, endOffset)) throw new RangeError('原生控件公式投影后的 UTF-16 范围无效。')
        else range = { blockId: address.block.id, startOffset, endOffset }
      } else reason = '原生控件文字宿主尚未映射为可输入段落。'
    } else reason = dto.textRanges.length ? '跨段或跨单元格的原生 SDT 结构范围仍需接入统一控件服务。' : '原生 SDT 尚无经过核实的文字或物理行范围。'
    let picture: DocumentBlock | undefined
    if (!reason && address && range && (dto.hasObjectContent || Number(dto.objectContentCount ?? 0) > 0 || properties.type === 'picture' && properties.showingPlaceholder !== true)) {
      const owned = blockRange ? documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(address.scoped, blockRange)) : rowRange
        ? documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(address.scoped, rowRange)) : undefined
      const images: DocumentBlock[] = []
      for (const id of owned ?? []) {
        const candidate = byBlockId.get(id)
        if (candidate && sameDocumentBlockHost(candidate.host, address.host) && candidate.block.type === 'image') images.push(candidate.block)
      }
      if (!owned || !images.length || images.some(image => !String(record(image.metadata.image).src ?? image.metadata.src ?? ''))
        || dto.objectContentCount !== undefined && images.length !== dto.objectContentCount
        || properties.type === 'picture' && images.length !== 1) reason = '原生对象 SDT 尚未映射到完整、唯一且真实的可编辑对象范围；完整来源保留。'
      else if (properties.type === 'picture') picture = images[0]
    }
    if (reason || !address || !range) {
      pending.set(dto.id, reason ?? '原生 SDT 范围未映射。')
      result.unmapped.push({ reason: pending.get(dto.id)!, control: dto })
      continue
    }
    const control: DocumentContentControl = { ...properties, id: dto.id, type: properties.type, ...range,
      ...(dto.parentId ? { parentId: dto.parentId } : {}), ...(rowRange ? { placement: 'row', tableRowRange: rowRange } : {}),
      ...(blockRange ? { placement: 'block', blockRange } : {}),
      ...(typeof checked === 'boolean' ? { checked } : {}), ...(typeof dateValue === 'string' ? { dateValue } : {}),
      metadata: { ...properties.metadata, ...(picture ? { pictureBlockId: picture.id, pictureAnchorBlockId: range.blockId } : {}),
        nativeSdtSource: { version: 1, partName: dto.partName, story: dto.story, noteId: dto.noteId } } }
    if (control.type === 'checkbox') { control.checked = control.checkbox?.checked ?? control.checked ?? false; control.checkbox = { ...control.checkbox, checked: control.checked } }
    const block = address.scoped.blocks && (rowRange ? textTargets(address.scoped).find(target => (target.cellBlockId ?? target.blockId) === control.blockId) : undefined)
    const text = blockRange ? documentContentControlBlockTextTargets(address.scoped, blockRange).map(target => runText(runsForTarget(address.scoped, target))).join('\n')
      : rowRange && block ? runText(runsForTarget(address.scoped, block)) : runText(address.block.runs).slice(control.startOffset, control.endOffset)
    control.value = picture && !control.showingPlaceholder ? String(record(picture.metadata.image).src ?? picture.metadata.src ?? '') : scalarValue(control, text)
    if (control.dataBinding) control.dataBinding = { ...control.dataBinding, cachedResult: control.value, dirty: false }
    controls.push(control); contexts.set(control.id, address)
  }
  // Keep a whole native subtree together until its parent's structural model is
  // implemented; stripping parentId would change locking/repeating semantics.
  let changed = true
  while (changed) {
    changed = false
    for (const control of controls) if (!pending.has(control.id) && control.parentId && pending.has(control.parentId)) { pending.set(control.id, '父控件的原生结构范围尚未映射。'); changed = true }
  }
  const byId = new Map(controls.map(control => [control.id, control]))
  for (const control of controls.filter(control => !pending.has(control.id))) {
    const address = contexts.get(control.id)!
    if (control.parentId) {
      const parent = byId.get(control.parentId), parentAddress = contexts.get(control.parentId)
      const contains = parent?.blockRange ? documentBlockContentControlContains(address.scoped, parent, control)
        : parent?.tableRowRange && control.blockRange ? control.blockRange.blockIds.every(id => documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(address.scoped, parent.tableRowRange!)).has(id))
          : parent && documentContentControlContains(address.scoped, parent, control)
      if (!parent || !parentAddress || !sameDocumentBlockHost(address.host, parentAddress.host) || !contains) throw new RangeError('原生内容控件父子范围或内容区不一致。')
    }
    let parent = control.parentId; const ancestors = new Set([control.id])
    while (parent) { if (ancestors.has(parent)) throw new RangeError('原生内容控件存在循环父级。'); ancestors.add(parent); parent = byId.get(parent)?.parentId }
  }
  for (const control of controls.filter(control => pending.has(control.id))) {
    const source = sources.get(control.id)
    if (source) result.unmapped.push({ reason: pending.get(control.id)!, control: cloneModel(source) })
  }
  model.contentControls = controls.filter(control => !pending.has(control.id))
  // Decorate only fully hydrated, uniquely owned pictures. Structural wrapper
  // controls never replace the scalar picture association on a shared image.
  const pictureOwners = new Map<string, DocumentContentControl[]>()
  for (const control of model.contentControls) if (control.type === 'picture' && typeof control.metadata?.pictureBlockId === 'string') {
    const owners = pictureOwners.get(control.metadata.pictureBlockId) ?? []
    owners.push(control); pictureOwners.set(control.metadata.pictureBlockId, owners)
  }
  for (const [id, owners] of pictureOwners) if (owners.length === 1) {
    const image = byBlockId.get(id)?.block
    if (image) {
      const owner = owners[0]!
      image.metadata.contentControlId = owner.id
      image.metadata.anchorBlockId = owner.blockId
      image.metadata.anchorOffset = owner.startOffset
      image.metadata.pictureInputAnchor = true
    }
  }
  result.hydratedCount = model.contentControls.length
  return result
}
