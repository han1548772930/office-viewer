import type { DocumentContentControl, DocumentCustomXmlPart, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { DocumentCustomXmlError, normalizeDocumentCustomXmlParts } from '../../../domain/document-custom-xml'
import { browserDocumentCustomXmlRuntime, type DocumentCustomXmlRuntime, type DocumentCustomXmlRuntimeDocument,
  type DocumentCustomXmlStructureDocument } from '../../../platform/dom/custom-xml-runtime'
import { textTargets } from '../model/document-model'
import { documentBlockHostForId, documentEditingBlockHost, sameDocumentBlockHost } from '../model/document-block-host'
import { documentModelForBlockHost } from './document-block-host'
import { assertDocumentPermission, permissionTargetsForBlock, type DocumentPermissionContext } from '../permissions/document-permissions'
import { assertDocumentCustomXmlBindings } from './document-custom-xml-validation'

type Binding = NonNullable<DocumentContentControl['dataBinding']>
const scalarTypes = new Set<DocumentContentControl['type']>(['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox'])

export interface DocumentRepeatingBindingStructureRequest {
  sectionId: string
  sourceItemId: string
  action: 'duplicate' | 'delete'
  resultItemId?: string
  identityMap: { control: ReadonlyMap<string, string> }
  permissionContext?: DocumentPermissionContext
}

export interface DocumentRepeatingBindingRemap {
  controlId: string
  before?: DocumentContentControl['dataBinding']
  after: Binding
}

export interface DocumentRepeatingBindingStructurePlan {
  changedParts: readonly DocumentCustomXmlPart[]
  bindingRemaps: readonly DocumentRepeatingBindingRemap[]
  scalarRefreshControlIds: readonly string[]
  /** Apply XML and bindings to this detached candidate; its shared candidate
   * boundary must refresh scalars/permissions before the content transaction. */
  commit(): void
}

interface CapturedBinding { control: DocumentContentControl; binding: Binding; keys: string[]; collection: boolean; implicit?: boolean }

function fail(code: string, message: string): never { throw new DocumentCustomXmlError(code, message) }
const active = (binding: DocumentContentControl['dataBinding']): binding is Binding & { storeItemID: string; xpath: string } => !!binding?.storeItemID && !!binding.xpath
const sameKeys = (a: readonly string[], b: readonly string[]) => a.length === b.length && a.every((value, index) => value === b[index])

function owningSection(controls: ReadonlyMap<string, DocumentContentControl>, item: DocumentContentControl): DocumentContentControl | undefined {
  let parent = item.parentId ? controls.get(item.parentId) : undefined
  const visited = new Set([item.id])
  while (parent) {
    if (visited.has(parent.id)) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复节的控件父关系循环。')
    visited.add(parent.id)
    if (parent.type === 'repeatingSection') return parent
    if (parent.parentId && !controls.has(parent.parentId)) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复节的父控件不存在。')
    parent = parent.parentId ? controls.get(parent.parentId) : undefined
  }
  return undefined
}

function orderedItems(model: DocumentModel, controls: ReadonlyMap<string, DocumentContentControl>, section: DocumentContentControl): DocumentContentControl[] {
  const editingHost = documentEditingBlockHost(model)
  const sectionHost = documentBlockHostForId(model, section.blockId) ?? editingHost
  const scoped = sectionHost && !sameDocumentBlockHost(editingHost, sectionHost) ? documentModelForBlockHost(model, sectionHost) : model
  const order = new Map(textTargets(scoped).map((target, index) => [target.cellBlockId ?? target.blockId, index]))
  return Array.from(controls.values()).filter(control => control.type === 'repeatingSectionItem' && owningSection(controls, control)?.id === section.id)
    .map((control, index) => {
      const position = order.get(control.blockId)
      const itemHost = documentBlockHostForId(model, control.blockId) ?? editingHost
      if (position === undefined || !sameDocumentBlockHost(sectionHost, itemHost)) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复项目的文字输入锚点不在同一真实文档宿主中。')
      return { control, position, index }
    }).sort((a, b) => a.position - b.position || a.control.startOffset - b.control.startOffset || a.index - b.index).map(item => item.control)
}

function controlSubtreeIds(controls: ReadonlyMap<string, DocumentContentControl>, rootId: string): Set<string> {
  const children = new Map<string, string[]>()
  for (const control of controls.values()) if (control.parentId) {
    const ids = children.get(control.parentId) ?? []
    ids.push(control.id); children.set(control.parentId, ids)
  }
  const ids = new Set([rootId]), pending = [rootId]
  for (let index = 0; index < pending.length; index++) for (const id of children.get(pending[index]!) ?? []) {
    if (ids.has(id)) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复节的控件父关系循环。')
    ids.add(id); pending.push(id)
  }
  return ids
}

function resolveKeys(document: DocumentCustomXmlRuntimeDocument, structure: DocumentCustomXmlStructureDocument,
  binding: Binding, collection: boolean): string[] {
  if (!active(binding)) return []
  return collection ? structure.resolveElements(binding.xpath, binding.prefixMappings).map(node => node.key)
    : (() => { const value = document.resolve(binding.xpath, binding.prefixMappings); return value ? [value.key] : [] })()
}

/** Real XML clone/remove with captured node identities. Positional XPath aliases
 * are relocated after insertion/removal; copies of descendants select cloned
 * nodes. All browser handles are disposed before returning the DTO-only plan. */
export function prepareRepeatingBindingStructure(before: DocumentModel, candidate: DocumentModel,
  request: DocumentRepeatingBindingStructureRequest,
  runtime: DocumentCustomXmlRuntime = browserDocumentCustomXmlRuntime): DocumentRepeatingBindingStructurePlan | undefined {
  if (before.id !== candidate.id) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复节的 XML 候选不属于同一文档。')
  const controls = new Map((before.contentControls ?? []).map(control => [control.id, control]))
  const section = controls.get(request.sectionId), item = controls.get(request.sourceItemId)
  if (!section || section.type !== 'repeatingSection' || !item || item.type !== 'repeatingSectionItem'
    || owningSection(controls, item)?.id !== section.id) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', 'XML 重复项目必须属于指定的真实重复节。')
  if (!['duplicate', 'delete'].includes(request.action)) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', 'XML 重复节操作无效。')
  const seed = active(item.dataBinding) ? item.dataBinding : active(section.dataBinding) ? section.dataBinding : undefined
  // The ordinary repeating planner owns unbound structure semantics, including
  // removing its empty section after the last item. No XML plan is needed there.
  if (before.customXmlParts === undefined || !seed) return undefined
  const candidateControls = new Map((candidate.contentControls ?? []).map(control => [control.id, control]))
  if (request.action === 'duplicate') {
    if (!request.resultItemId || request.identityMap.control.get(item.id) !== request.resultItemId || controls.has(request.resultItemId)
      || candidateControls.get(request.resultItemId)?.type !== 'repeatingSectionItem'
      || owningSection(candidateControls, candidateControls.get(request.resultItemId)!)?.id !== section.id) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', 'XML 重复项目的复制身份与候选不一致。')
  } else if (request.resultItemId || request.identityMap.control.size || candidateControls.has(item.id)) {
    fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', 'XML 重复项目的删除计划与候选不一致。')
  }
  const candidateSection = candidateControls.get(section.id)
  const siblings = orderedItems(before, controls, section)
  if (candidateSection && candidateSection.type !== 'repeatingSection') fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', 'XML 重复节在候选中改变了控件类型。')
  if (!candidateSection) {
    const removedIds = controlSubtreeIds(controls, section.id)
    if (request.action !== 'delete' || siblings.length !== 1 || siblings[0]!.id !== item.id
      || Array.from(candidateControls.values()).some(control => removedIds.has(control.id) || !!control.parentId && removedIds.has(control.parentId))) {
      fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '只有删除末项及完整控件子树时才可以移除空重复节。')
    }
  }
  const nextSiblings = candidateSection ? orderedItems(candidate, candidateControls, candidateSection) : []
  const sourceIndex = siblings.findIndex(sibling => sibling.id === item.id)
  const expectedOrder = siblings.flatMap(sibling => sibling.id === item.id ? request.action === 'duplicate'
    ? [sibling.id, request.resultItemId!] : [] : [sibling.id])
  if (sourceIndex < 0 || !sameKeys(expectedOrder, nextSiblings.map(sibling => sibling.id))) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', 'XML 结构操作必须保持真实项目的复制位置或删除顺序。')
  assertDocumentCustomXmlBindings(before); assertDocumentCustomXmlBindings(candidate)
  if (section.repeatingSection?.doNotAllowInsertDelete) fail('DOC_CUSTOM_XML_STRUCTURE_LOCKED', '此重复节禁止增加或删除项目。')
  const store = seed.storeItemID.toLowerCase()
  if (active(item.dataBinding) && active(section.dataBinding) && item.dataBinding.storeItemID.toLowerCase() !== section.dataBinding.storeItemID.toLowerCase()) {
    fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复节与项目不能绑定不同的 XML 数据源。')
  }
  const parts = normalizeDocumentCustomXmlParts(candidate.customXmlParts), part = parts.find(part => part.storeItemID?.toLowerCase() === store)
  const originalPart = before.customXmlParts.find(part => part.storeItemID?.toLowerCase() === store)
  if (!part || !originalPart) fail('DOC_CUSTOM_XML_STORE_NOT_FOUND', '重复节的数据源不存在。')
  if (part.xml !== originalPart.xml || part.partName !== originalPart.partName) fail('DOC_CUSTOM_XML_BINDING_CHANGED', '重复项目的源 XML 在结构计划前已经变化。')
  const targetIds = new Set<string>()
  for (const [source, target] of request.identityMap.control) {
    if (!controls.has(source) || !candidateControls.has(target) || controls.has(target) || targetIds.has(target)) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复项目的控件复制映射不是唯一的新身份。')
    const original = controls.get(source)!, copied = candidateControls.get(target)!
    let ancestor: DocumentContentControl | undefined = original
    const visited = new Set<string>()
    while (ancestor && ancestor.id !== item.id) {
      if (visited.has(ancestor.id)) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复项目的复制源存在循环父级。')
      visited.add(ancestor.id)
      ancestor = ancestor.parentId ? controls.get(ancestor.parentId) : undefined
    }
    if (!ancestor || original.type !== copied.type || copied.parentId !== (original.id === item.id ? original.parentId
      : original.parentId && request.identityMap.control.get(original.parentId))) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', 'XML 复制映射必须保持项目内真实控件类型和父子关系。')
    targetIds.add(target)
  }
  const document = runtime.open(part.xml)
  try {
    const structure = document.structure
    if (!structure) fail('DOC_CUSTOM_XML_PLATFORM_UNAVAILABLE', '当前 XML 平台没有真实重复节结构操作。')
    const captures = new Map<string, CapturedBinding>()
    for (const control of controls.values()) {
      if (!active(control.dataBinding) || control.dataBinding.storeItemID.toLowerCase() !== store || control.metadata?.officeStoryLockSource) continue
      const collection = control.type === 'repeatingSection'
      const keys = resolveKeys(document, structure, control.dataBinding, collection)
      if (keys.length) captures.set(control.id, { control, binding: { ...control.dataBinding }, keys, collection })
    }
    let root: string
    const itemBinding = item.dataBinding
    if (active(itemBinding)) {
      const nodes = structure.resolveElements(itemBinding.xpath, itemBinding.prefixMappings)
      if (nodes.length !== 1) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复项目的 XPath 必须唯一选择其元素。')
      root = nodes[0]!.key
      captures.set(item.id, { control: item, binding: { ...itemBinding }, keys: [root], collection: false })
    } else {
      const collection = structure.resolveElements(seed.xpath, seed.prefixMappings)
      if (collection.length !== siblings.length) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复节的 XML 元素集合与真实项目数量不一致。')
      for (const [index, sibling] of siblings.entries()) {
        if (active(sibling.dataBinding)) {
          const keys = structure.resolveElements(sibling.dataBinding.xpath, sibling.dataBinding.prefixMappings).map(node => node.key)
          if (!sameKeys(keys, [collection[index]!.key])) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复项目的 XML 绑定与实际项目顺序不一致。')
        } else captures.set(sibling.id, { control: sibling, binding: { ...sibling.dataBinding, storeItemID: seed.storeItemID },
          keys: [collection[index]!.key], collection: false, implicit: true })
      }
      const index = siblings.findIndex(sibling => sibling.id === item.id)
      if (index < 0) fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复项目不在当前重复节中。')
      root = collection[index]!.key
    }
    const sectionCapture = captures.get(section.id)
    if (active(section.dataBinding) && (!sectionCapture || !sectionCapture.keys.some(key => structure.contains(key, root)))) {
      fail('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复项目的 XML 元素不在重复节绑定中。')
    }
    const sectionCollection = !!sectionCapture?.keys.includes(root)
    const tracked = new Set<string>([root])
    for (const capture of captures.values()) for (const key of capture.keys) if (structure.contains(root, key)) tracked.add(key)
    if (request.action === 'delete') {
      for (const capture of captures.values()) {
        if (capture.control.id === section.id && sectionCollection || !candidateControls.has(capture.control.id)) continue
        if (capture.keys.some(key => structure.contains(root, key))) fail('DOC_CUSTOM_XML_REFERENCED_NODE', '删除重复项目会移除仍被其他控件引用的 XML 节点，请先解除对应绑定。')
      }
    }
    const copied = request.action === 'duplicate' ? structure.cloneElementAfter(root, Array.from(tracked)) : undefined
    if (request.action === 'delete') structure.removeElement(root)
    const remaps = new Map<string, DocumentRepeatingBindingRemap>()
    const propose = (controlId: string, capture: CapturedBinding, keys: string[], forcePath = false) => {
      const control = candidateControls.get(controlId)
      if (!control) return
      const prior = control.dataBinding
      const resolved = !forcePath && prior ? resolveKeys(document, structure, prior, capture.collection) : []
      const retained = prior && sameKeys(keys, resolved)
      const path = !keys.length ? { xpath: '/*[false()]', prefixMappings: undefined }
        : retained ? { xpath: prior?.xpath, prefixMappings: prior?.prefixMappings }
        : capture.collection ? structure.bindingForElements(keys) : structure.bindingFor(keys[0]!)
      const after: Binding = { ...capture.binding, ...prior, storeItemID: seed.storeItemID,
        xpath: path.xpath, prefixMappings: path.prefixMappings, cachedResult: keys.map(key => document.read(key)).join(''), dirty: false }
      if (JSON.stringify(prior) !== JSON.stringify(after)) remaps.set(controlId, { controlId, ...(prior ? { before: cloneModel(prior) } : {}), after })
    }
    for (const capture of captures.values()) {
      if (!candidateControls.has(capture.control.id)) continue
      let keys = capture.keys
      if (capture.control.id === section.id && sectionCollection) {
        keys = capture.keys.flatMap(key => key === root ? copied ? [key, copied.key] : [] : [key])
      }
      propose(capture.control.id, capture, keys, capture.implicit)
    }
    if (copied) for (const [source, target] of request.identityMap.control) {
      const capture = captures.get(source)
      if (!capture) continue
      const keys = capture.keys.map(key => copied.copiedKeys.get(key) ?? key)
      propose(target, capture, keys, true)
    }
    const xml = document.serialize()
    part.xml = xml
    const nextParts = normalizeDocumentCustomXmlParts(parts)
    const preparedRemaps = Array.from(remaps.values()).map(remap => cloneModel(remap))
    const bindingRemaps = preparedRemaps.map(remap => cloneModel(remap))
    if (request.permissionContext) {
      const affected = new Set<string>([item.blockId, section.blockId])
      for (const control of controls.values()) if (active(control.dataBinding) && control.dataBinding.storeItemID.toLowerCase() === store) {
        for (const id of [control.blockId, ...(control.blockRange?.blockIds ?? []), ...(control.tableRowRange ? [control.tableRowRange.tableId] : [])]) affected.add(id)
      }
      for (const remap of bindingRemaps) {
        const original = controls.get(remap.controlId) ?? controls.get(Array.from(request.identityMap.control).find(([, target]) => target === remap.controlId)?.[0] ?? '')
        if (original) for (const id of [original.blockId, ...(original.blockRange?.blockIds ?? []), ...(original.tableRowRange ? [original.tableRowRange.tableId] : [])]) affected.add(id)
      }
      const targets = Array.from(affected).flatMap(id => permissionTargetsForBlock(before, id))
      assertDocumentPermission(before.permissionRules ?? [], request.permissionContext, 'edit', targets)
      if (copied) assertDocumentPermission(before.permissionRules ?? [], request.permissionContext, 'copy', targets)
    }
    const scalarRefreshControlIds = Array.from(candidateControls.values()).filter(control => scalarTypes.has(control.type)
      && active(control.dataBinding) && control.dataBinding.storeItemID.toLowerCase() === store).map(control => control.id)
    const expectedParts = JSON.stringify(candidate.customXmlParts)
    const expectedBindings = new Map(bindingRemaps.map(remap => [remap.controlId, JSON.stringify(candidateControls.get(remap.controlId)!.dataBinding)]))
    let committed = false
    return {
      changedParts: xml !== originalPart.xml ? [cloneModel(part)] : [], bindingRemaps, scalarRefreshControlIds,
      commit() {
        if (committed) fail('DOC_CUSTOM_XML_FAILED_SESSION', '重复节 XML 结构计划已经应用。')
        if (JSON.stringify(candidate.customXmlParts) !== expectedParts) fail('DOC_CUSTOM_XML_BINDING_CHANGED', '重复节 XML 候选在应用前已经变化。')
        const current = new Map((candidate.contentControls ?? []).map(control => [control.id, control]))
        for (const [id, binding] of expectedBindings) if (!current.has(id) || JSON.stringify(current.get(id)!.dataBinding) !== binding) {
          fail('DOC_CUSTOM_XML_BINDING_CHANGED', '重复节控件的绑定在结构应用前已经变化。')
        }
        candidate.customXmlParts = cloneModel(nextParts)
        for (const remap of preparedRemaps) current.get(remap.controlId)!.dataBinding = cloneModel(remap.after)
        committed = true
      },
    }
  } finally { document.dispose() }
}
