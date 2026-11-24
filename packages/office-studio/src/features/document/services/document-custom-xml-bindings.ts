import type { DocumentContentControl, DocumentCustomXmlPart, DocumentModel } from '../../../domain/model'
import { browserDocumentCustomXmlRuntime, type DocumentCustomXmlRuntime, type DocumentCustomXmlRuntimeDocument } from '../../../platform/dom/custom-xml-runtime'
import { DocumentCustomXmlError, normalizeDocumentCustomXmlParts } from '../../../domain/document-custom-xml'
import type { DocumentContentControlBindingSource } from './content-control-bindings'

export interface DocumentCustomXmlControlBinding {
  controlId: string
  key: string
  value: string
}

export interface DocumentCustomXmlBindingSession {
  /** Resolve and atomically plan changed scalars; keys identify actual selected XML nodes. */
  applySources(sources: readonly DocumentContentControlBindingSource[]): readonly DocumentContentControlBindingSource[]
  /** Current values for every scalar control in a touched data store, including XPath aliases. */
  bindings(): readonly DocumentCustomXmlControlBinding[]
  /** Explicit XML-source edits refresh all mappings, without rewriting source XML from stale UI. */
  refresh(): readonly DocumentCustomXmlControlBinding[]
  /** Publish only into the caller's isolated candidate, after all peer/permission checks. */
  commit(): void
  dispose(): void
}

const scalarTypes = new Set<DocumentContentControl['type']>(['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox'])

interface OpenStore { part: DocumentCustomXmlPart; document: DocumentCustomXmlRuntimeDocument; assigned: Map<string, string> }
interface PendingSource { source: DocumentContentControlBindingSource; store?: string; node?: string; entry?: OpenStore }

function booleanEcho(source: DocumentContentControlBindingSource, previous: string): boolean {
  if (source.control.type !== 'checkbox') return false
  const boolean = (value: string): boolean | undefined => value === 'true' || value === '1' ? true : value === 'false' || value === '0' || value === '' ? false : undefined
  const value = boolean(previous)
  return value !== undefined && value === boolean(source.value)
}

/** No data store in a legacy snapshot means peer synchronization only, never a claimed XML write. */
export function createDocumentCustomXmlBindingSession(candidate: DocumentModel,
  runtime: DocumentCustomXmlRuntime = browserDocumentCustomXmlRuntime): DocumentCustomXmlBindingSession | undefined {
  if (candidate.customXmlParts === undefined) return undefined
  const parts = normalizeDocumentCustomXmlParts(candidate.customXmlParts)
  const stores = new Map(parts.filter(part => part.storeItemID).map(part => [part.storeItemID!.toLowerCase(), part]))
  const opened = new Map<string, OpenStore>(), sourceNodes = new Map<string, { store: string; node: string }>()
  const selectedPeers = new Map<string, { store: string; node: string }>()
  let disposed = false, failed = false
  const assertLive = () => {
    if (disposed) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_DISPOSED', '自定义 XML 数据源会话已经释放。')
    if (failed) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_FAILED_SESSION', '失败的 XML 绑定会话不能继续或发布部分结果。')
  }
  const canonicalKey = (store: string, node: string) => 'xml:' + JSON.stringify([store, node])
  const open = (store: string): OpenStore => {
    assertLive()
    let entry = opened.get(store)
    if (!entry) {
      const part = stores.get(store)
      if (!part) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_STORE_NOT_FOUND', '内容控件引用的自定义 XML 数据源不存在：' + store)
      entry = { part, document: runtime.open(part.xml), assigned: new Map() }; opened.set(store, entry)
    }
    return entry
  }
  const currentBindings = (): DocumentCustomXmlControlBinding[] => {
    assertLive()
    const bindings: DocumentCustomXmlControlBinding[] = []
    for (const control of candidate.contentControls ?? []) {
      const binding = control.dataBinding, store = binding?.storeItemID?.toLowerCase()
      if (!store || !binding?.xpath || !opened.has(store) || !scalarTypes.has(control.type) || control.metadata?.officeStoryLockSource) continue
      const entry = opened.get(store)!
      const resolved = entry.document.resolve(binding.xpath, binding.prefixMappings)
      if (resolved) bindings.push({ controlId: control.id, key: canonicalKey(store, resolved.key), value: resolved.value })
      else {
        // A value predicate may cease matching after writing its selected node.
        // Retain the selected source node for this transaction; don't invent a new node.
        const source = sourceNodes.get(control.id) ?? selectedPeers.get(control.id)
        if (source?.store === store) bindings.push({ controlId: control.id, key: canonicalKey(store, source.node), value: entry.document.read(source.node) })
      }
    }
    return bindings
  }
  return {
    applySources(sources) {
      assertLive()
      try {
      const pending: PendingSource[] = sources.map(source => {
        const binding = source.control.dataBinding, store = binding?.storeItemID?.toLowerCase()
        if (!store || !binding?.xpath) return { source }
        const entry = open(store), prior = sourceNodes.get(source.control.id)
        const resolved = entry.document.resolve(binding.xpath, binding.prefixMappings)
        const node = prior?.store === store ? prior.node : resolved?.key
        if (!node) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_NODE_NOT_FOUND', '内容控件 XPath 没有匹配可写节点：' + source.control.id)
        if (prior && (prior.store !== store || resolved && resolved.key !== prior.node)) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_BINDING_CHANGED', '同一事务中的内容控件 XML 绑定目标发生变化。')
        sourceNodes.set(source.control.id, { store, node })
        return { source, store, node, entry }
      })
      // Capture aliases before an element write removes its text/CDATA node.
      // Missing predicates may still resolve to a different node after writing;
      // a remembered node is only used when the new XPath has no match.
      for (const control of candidate.contentControls ?? []) {
        const binding = control.dataBinding, store = binding?.storeItemID?.toLowerCase()
        if (!store || !binding?.xpath || !opened.has(store) || !scalarTypes.has(control.type) || control.metadata?.officeStoryLockSource) continue
        const resolved = opened.get(store)!.document.resolve(binding.xpath, binding.prefixMappings)
        if (resolved && !selectedPeers.has(control.id)) selectedPeers.set(control.id, { store, node: resolved.key })
      }
      // Plan every source before changing any node, so conflicts never depend on source order.
      for (const item of pending) {
        if (!item.entry || !item.node) continue
        const previous = item.entry.assigned.get(item.node)
        if (previous !== undefined && previous !== item.source.value && !booleanEcho(item.source, previous)) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_VALUE_CONFLICT', '同一编辑事务为同一 XML 节点提供了不同标量。')
        for (const other of pending) {
          if (other === item || other.entry !== item.entry || !other.node) continue
          if (other.node === item.node && other.source.value !== item.source.value) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_VALUE_CONFLICT', '同一编辑事务为同一 XML 节点提供了不同标量。')
          if (other.node !== item.node && item.entry.document.overlaps(item.node, other.node)) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_NODE_CONFLICT', '同一编辑事务不能覆盖相互包含的 XML 节点。')
        }
        for (const node of item.entry.assigned.keys()) if (node !== item.node && item.entry.document.overlaps(item.node, node)) throw new DocumentCustomXmlError('DOC_CUSTOM_XML_NODE_CONFLICT', 'XML 绑定级联覆盖了之前的源节点。')
      }
      return pending.map(item => {
        if (!item.entry || !item.node || !item.store) return item.source
        const previous = item.entry.assigned.get(item.node)
        const value = previous !== undefined && booleanEcho(item.source, previous) ? previous : item.source.value
        item.entry.document.write(item.node, value)
        item.entry.assigned.set(item.node, value)
        return { ...item.source, key: canonicalKey(item.store, item.node), value }
      })
      } catch (error) { failed = true; throw error }
    },
    bindings() {
      assertLive()
      try { return currentBindings() }
      catch (error) { failed = true; throw error }
    },
    refresh() {
      assertLive()
      try {
        for (const control of candidate.contentControls ?? []) {
          const store = control.dataBinding?.storeItemID?.toLowerCase()
          if (store && control.dataBinding?.xpath && scalarTypes.has(control.type) && !control.metadata?.officeStoryLockSource) open(store)
        }
        return currentBindings()
      } catch (error) { failed = true; throw error }
    },
    commit() {
      assertLive()
      try {
        for (const entry of opened.values()) entry.part.xml = entry.document.serialize()
        const normalized = normalizeDocumentCustomXmlParts(parts), bindings = currentBindings()
        const updates = bindings.flatMap(binding => {
          const control = candidate.contentControls?.find(control => control.id === binding.controlId)
          return control?.dataBinding ? [{ control, dataBinding: { ...control.dataBinding, cachedResult: binding.value, dirty: false } }] : []
        })
        candidate.customXmlParts = normalized
        for (const update of updates) update.control.dataBinding = update.dataBinding
      } catch (error) { failed = true; throw error }
    },
    dispose() { if (disposed) return; disposed = true; for (const entry of opened.values()) entry.document.dispose(); opened.clear(); sourceNodes.clear(); selectedPeers.clear() },
  }
}
