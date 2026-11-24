import { DocumentCustomXmlError, DOCUMENT_CUSTOM_XML_MAX_DEPTH, DOCUMENT_CUSTOM_XML_MAX_ELEMENTS,
  DOCUMENT_CUSTOM_XML_MAX_PART_BYTES, DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH, documentCustomXmlUtf8Bytes,
  documentCustomXmlResourceByteLength } from '../../domain/document-custom-xml'

export interface DocumentCustomXmlNodeValue { key: string; value: string }
export interface DocumentCustomXmlNodeBinding { xpath: string; prefixMappings?: string }
export interface DocumentCustomXmlElementCopy { key: string; copiedKeys: ReadonlyMap<string, string> }

/** Structural operations share the same detached XML document and XPath engine.
 * Only opaque identities and serializable bindings leave this platform layer. */
export interface DocumentCustomXmlStructureDocument {
  resolveElements(xpath: string, prefixMappings?: string): readonly DocumentCustomXmlNodeValue[]
  contains(parent: string, child: string): boolean
  cloneElementAfter(key: string, trackedKeys: readonly string[]): DocumentCustomXmlElementCopy
  removeElement(key: string): void
  bindingFor(key: string): DocumentCustomXmlNodeBinding
  bindingForElements(keys: readonly string[]): DocumentCustomXmlNodeBinding
}

/** Opaque node identity keeps browser handles out of commands and persistent models. */
export interface DocumentCustomXmlRuntimeDocument {
  resolve(xpath: string, prefixMappings?: string): DocumentCustomXmlNodeValue | undefined
  read(key: string): string
  write(key: string, value: string): void
  overlaps(left: string, right: string): boolean
  structure?: DocumentCustomXmlStructureDocument
  serialize(): string
  dispose(): void
}

export interface DocumentCustomXmlRuntime {
  open(xml: string): DocumentCustomXmlRuntimeDocument
}

export interface DocumentCustomXmlPackageRelationship {
  id: string
  type: string
  target: string
  external: boolean
}

/** Inspection returns DTOs only; DOM nodes remain owned by this platform layer. */
export interface DocumentCustomXmlPackageRuntime {
  validate(xml: string): void
  propertiesStoreItemID(xml: string): string
  relationships(xml: string): readonly DocumentCustomXmlPackageRelationship[]
  decodeXmlResource(dataBase64: string): string
}

const CUSTOM_XML_NAMESPACES = ['http://schemas.openxmlformats.org/officeDocument/2006/customXml', 'http://purl.oclc.org/ooxml/officeDocument/customXml']
const RELATIONSHIP_NAMESPACES = ['http://schemas.openxmlformats.org/package/2006/relationships', 'http://purl.oclc.org/ooxml/package/relationships']

const XMLNS = 'http://www.w3.org/2000/xmlns/'
const XML = 'http://www.w3.org/XML/1998/namespace'

function invalid(code: string, message: string, error?: unknown): never {
  throw new DocumentCustomXmlError(code, message, error instanceof Error ? error.message : undefined)
}

/** Inspect actual declarations before DOMParser, preserving literal text in comments/CDATA/PIs. */
function assertNoDocumentType(value: string): void {
  let cursor = 0
  while (cursor < value.length) {
    const start = value.indexOf('<', cursor)
    if (start < 0) return
    const skip = (opener: string, closer: string): boolean => {
      if (!value.startsWith(opener, start)) return false
      const end = value.indexOf(closer, start + opener.length)
      cursor = end < 0 ? value.length : end + closer.length
      return true
    }
    if (skip('<!--', '-->') || skip('<![CDATA[', ']]>') || skip('<?', '?>')) continue
    if (value.startsWith('<!', start)) invalid('DOC_CUSTOM_XML_INVALID_XML', '自定义 XML 不能包含文档类型或实体声明。')
    // Quoted attribute values belong to this tag; malformed markup remains DOMParser's responsibility.
    let quote = ''
    cursor = start + 1
    for (; cursor < value.length; cursor++) {
      const character = value[cursor]
      if (quote) { if (character === quote) quote = '' }
      else if (character === '"' || character === "'") quote = character
      else if (character === '>') { cursor++; break }
    }
  }
}

function parseXml(value: string): XMLDocument {
  if (documentCustomXmlUtf8Bytes(value) > DOCUMENT_CUSTOM_XML_MAX_PART_BYTES) invalid('DOC_CUSTOM_XML_RESOURCE_LIMIT', '自定义 XML 部件超过 8 MiB 限额。')
  assertNoDocumentType(value)
  if (typeof DOMParser === 'undefined') invalid('DOC_CUSTOM_XML_PLATFORM_UNAVAILABLE', '当前环境没有 XML/XPath 平台适配。')
  const document = new DOMParser().parseFromString(value, 'application/xml')
  if (!document.documentElement || document.getElementsByTagNameNS('http://www.mozilla.org/newlayout/xml/parsererror.xml', 'parsererror').length
    || document.documentElement.localName === 'parsererror' && document.documentElement.namespaceURI === 'http://www.w3.org/1999/xhtml') {
    invalid('DOC_CUSTOM_XML_INVALID_XML', '自定义 XML 部件不是有效的 XML。')
  }
  let elements = 0
  const pending: Array<{ node: Element; depth: number }> = [{ node: document.documentElement, depth: 1 }]
  while (pending.length) {
    const { node, depth } = pending.pop()!
    if (++elements > DOCUMENT_CUSTOM_XML_MAX_ELEMENTS || depth > DOCUMENT_CUSTOM_XML_MAX_DEPTH) invalid('DOC_CUSTOM_XML_RESOURCE_LIMIT', '自定义 XML 的节点数量或嵌套深度超过限额。')
    for (let child = node.lastElementChild; child; child = child.previousElementSibling) pending.push({ node: child, depth: depth + 1 })
  }
  return document
}

function namespaces(value: string | undefined): Map<string, string> {
  const result = new Map<string, string>([['xml', XML]])
  if (!value) return result
  if (value.length > DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH) invalid('DOC_CUSTOM_XML_INVALID_NAMESPACE', 'XML 命名空间映射超过限额。')
  const root = parseXml(`<binding ${value}/>`).documentElement
  for (const attribute of root.attributes) {
    if (attribute.namespaceURI !== XMLNS) invalid('DOC_CUSTOM_XML_INVALID_NAMESPACE', 'XML 命名空间映射只能声明 xmlns 属性。')
    result.set(attribute.name === 'xmlns' ? '' : attribute.localName, attribute.value)
  }
  return result
}

export const browserDocumentCustomXmlPackageRuntime: DocumentCustomXmlPackageRuntime = {
  validate(xml) { parseXml(xml) },
  propertiesStoreItemID(xml) {
    const root = parseXml(xml).documentElement
    if (root.localName !== 'datastoreItem' || !CUSTOM_XML_NAMESPACES.includes(root.namespaceURI ?? '')) invalid('DOC_CUSTOM_XML_INVALID_PROPERTIES', '自定义 XML properties 必须使用 datastoreItem 根节点。')
    const id = CUSTOM_XML_NAMESPACES.map(namespace => root.getAttributeNS(namespace, 'itemID')).find(value => value !== null)
    if (!id || id.length > 128 || /[\x00-\x1f\x7f-\x9f]/.test(id)) invalid('DOC_CUSTOM_XML_INVALID_PROPERTIES', '自定义 XML properties 的数据源 ID 无效。')
    return id
  },
  relationships(xml) {
    const root = parseXml(xml).documentElement
    if (root.localName !== 'Relationships' || !RELATIONSHIP_NAMESPACES.includes(root.namespaceURI ?? '')) invalid('DOC_CUSTOM_XML_INVALID_RELATIONSHIPS', '自定义 XML 关系必须使用 OPC Relationships 根节点。')
    for (const node of root.childNodes) if ((node.nodeType === 3 || node.nodeType === 4) && node.nodeValue?.trim()) invalid('DOC_CUSTOM_XML_INVALID_RELATIONSHIPS', 'OPC Relationships 根节点不能包含正文。')
    const ids = new Set<string>()
    return Array.from(root.children, element => {
      if (element.localName !== 'Relationship' || element.namespaceURI !== root.namespaceURI || element.children.length || element.textContent?.trim()) invalid('DOC_CUSTOM_XML_INVALID_RELATIONSHIPS', 'OPC 关系包含无效的子节点或正文。')
      const attribute = (name: string): string => {
        const value = element.getAttributeNS(null, name)
        if (!value) invalid('DOC_CUSTOM_XML_INVALID_RELATIONSHIPS', 'OPC 关系缺少属性：' + name)
        return value
      }
      const id = attribute('Id'), mode = element.getAttributeNS(null, 'TargetMode')
      if (ids.has(id)) invalid('DOC_CUSTOM_XML_INVALID_RELATIONSHIPS', 'OPC 关系 ID 重复：' + id)
      ids.add(id)
      if (mode !== null && mode !== 'Internal' && mode !== 'External') invalid('DOC_CUSTOM_XML_INVALID_RELATIONSHIPS', 'OPC 关系 TargetMode 无效。')
      return { id, type: attribute('Type'), target: attribute('Target'), external: mode === 'External' }
    })
  },
  decodeXmlResource(dataBase64) {
    const length = documentCustomXmlResourceByteLength(dataBase64), bytes = new Uint8Array(length)
    const digit = (code: number): number => code >= 65 && code <= 90 ? code - 65 : code >= 97 && code <= 122 ? code - 71
      : code >= 48 && code <= 57 ? code + 4 : code === 43 ? 62 : code === 47 ? 63 : 0
    for (let index = 0, offset = 0; index < dataBase64.length; index += 4) {
      const value = digit(dataBase64.charCodeAt(index)) * 262144 + digit(dataBase64.charCodeAt(index + 1)) * 4096
        + digit(dataBase64.charCodeAt(index + 2)) * 64 + digit(dataBase64.charCodeAt(index + 3))
      if (offset < length) bytes[offset++] = value >>> 16
      if (offset < length) bytes[offset++] = value >>> 8
      if (offset < length) bytes[offset++] = value
    }
    const encoding = bytes[0] === 255 && bytes[1] === 254 || bytes[0] === 60 && bytes[1] === 0 ? 'utf-16le'
      : bytes[0] === 254 && bytes[1] === 255 || bytes[0] === 0 && bytes[1] === 60 ? 'utf-16be' : 'utf-8'
    try { return new TextDecoder(encoding, { fatal: true }).decode(bytes) }
    catch (error) { invalid('DOC_CUSTOM_XML_INVALID_ENCODING', '自定义 XML 关系或 properties 资源的字节编码无效。', error) }
  },
}

/**
 * Browser-native XPath 1.0 is the single current query engine. Detached XML documents
 * never become UI DOM. This synchronous platform baseline is bounded; Worker migration
 * and main-thread timing still require the deferred same-dataset benchmarks.
 */
export const browserDocumentCustomXmlRuntime: DocumentCustomXmlRuntime = {
  open(xml) {
    let document: XMLDocument | undefined = parseXml(xml)
    const ids = new WeakMap<Node, string>(), nodes = new Map<string, Node>()
    const mappingCache = new Map<string, Map<string, string>>()
    let changed = false, serial = 0
    const liveDocument = () => {
      if (!document) invalid('DOC_CUSTOM_XML_DISPOSED', '自定义 XML 数据源会话已经释放。')
      return document
    }
    // /value and /value/text() are aliases for the same scalar in a leaf element.
    const scalarNode = (node: Node): Node => {
      if ((node.nodeType === 3 || node.nodeType === 4) && node.parentNode?.nodeType === 1) {
        const parent = node.parentNode
        if (parent.childNodes.length === 1) return parent
      }
      return node
    }
    const nodeFor = (key: string): Node => {
      liveDocument()
      const node = nodes.get(key)
      if (!node) invalid('DOC_CUSTOM_XML_INVALID_NODE', '自定义 XML 节点身份不存在。')
      if (node.nodeType === 2 ? !(node as Attr).ownerElement?.isConnected : !node.isConnected) invalid('DOC_CUSTOM_XML_DETACHED_NODE', '自定义 XML 节点已经移出数据源。')
      return node
    }
    const valueFor = (node: Node) => node.nodeType === 1 ? node.textContent ?? '' : node.nodeValue ?? ''
    const keyFor = (node: Node): string => {
      let key = ids.get(node)
      if (!key) { key = `xml-node-${++serial}`; ids.set(node, key); nodes.set(key, node) }
      return key
    }
    const evaluate = (xpath: string, prefixMappings: string | undefined, snapshot: boolean): Node[] => {
      const current = liveDocument()
      if (typeof xpath !== 'string' || !xpath || xpath.length > DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH) invalid('DOC_CUSTOM_XML_INVALID_XPATH', '内容控件的 XPath 无效或超过限额。')
      const mappingKey = prefixMappings ?? ''
      let mapping = mappingCache.get(mappingKey)
      if (!mapping) { mapping = namespaces(prefixMappings); mappingCache.set(mappingKey, mapping) }
      try {
        const result = current.evaluate(xpath, current, prefix => mapping!.get(prefix ?? '') ?? current.documentElement.lookupNamespaceURI(prefix), snapshot ? 7 : 9, null)
        return snapshot ? Array.from({ length: result.snapshotLength }, (_, index) => result.snapshotItem(index)!)
          : result.singleNodeValue ? [result.singleNodeValue] : []
      } catch (error) { invalid('DOC_CUSTOM_XML_INVALID_XPATH', '内容控件的 XPath 或命名空间映射无效。', error) }
    }
    const contains = (parent: Node, child: Node): boolean => {
      const target = child.nodeType === 2 ? (child as Attr).ownerElement : child
      return !!target && (parent === target || parent.contains(target))
    }
    const elementFor = (key: string): Element => {
      const node = nodeFor(key)
      if (node.nodeType !== 1) invalid('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复节的数据绑定必须选择 XML 元素。')
      return node as Element
    }
    const bindingForNodes = (selection: readonly Node[], collection: boolean): DocumentCustomXmlNodeBinding => {
      if (!selection.length) invalid('DOC_CUSTOM_XML_INVALID_STRUCTURE', '不能为不存在的 XML 节点生成绑定。')
      const prefixByNamespace = new Map<string, string>()
      const qualified = (node: Element | Attr): string => {
        const namespace = node.namespaceURI
        if (!namespace) return node.localName
        if (namespace === XML) return `xml:${node.localName}`
        let prefix = prefixByNamespace.get(namespace)
        if (!prefix) { prefix = `ox${prefixByNamespace.size}`; prefixByNamespace.set(namespace, prefix) }
        return `${prefix}:${node.localName}`
      }
      const sameName = (a: Element, b: Element) => a.localName === b.localName && a.namespaceURI === b.namespaceURI
      const path = (node: Node): string => {
        if (node.nodeType === 2) {
          const attribute = node as Attr
          if (!attribute.ownerElement) invalid('DOC_CUSTOM_XML_DETACHED_NODE', 'XML 属性已经移出数据源。')
          return `${path(attribute.ownerElement)}/@${qualified(attribute)}`
        }
        const parent = node.parentNode
        if (!parent) invalid('DOC_CUSTOM_XML_DETACHED_NODE', 'XML 节点已经移出数据源。')
        const root = parent.nodeType === 9 ? '' : path(parent)
        if (node.nodeType === 1) {
          const element = node as Element
          let index = 1
          for (let sibling = element.previousElementSibling; sibling; sibling = sibling.previousElementSibling) if (sameName(element, sibling)) index++
          return `${root}/${qualified(element)}[${index}]`
        }
        const test = node.nodeType === 3 || node.nodeType === 4 ? 'text()' : node.nodeType === 7 ? 'processing-instruction()' : node.nodeType === 8 ? 'comment()' : undefined
        if (!test) invalid('DOC_CUSTOM_XML_INVALID_NODE', '无法为此 XML 节点生成可写绑定。')
        let index = 1
        for (let sibling = node.previousSibling; sibling; sibling = sibling.previousSibling) {
          if (test === 'text()' ? sibling.nodeType === 3 || sibling.nodeType === 4 : sibling.nodeType === node.nodeType) index++
        }
        return `${root}/${test}[${index}]`
      }
      const first = selection[0]!
      let xpath: string
      if (collection && first.nodeType === 1 && selection.every(node => node.nodeType === 1 && node.parentNode === first.parentNode && sameName(first as Element, node as Element))) {
        const siblings = Array.from((first.parentNode as Element | Document).children).filter(node => sameName(first as Element, node))
        const selected = new Set(selection)
        xpath = siblings.length === selected.size && siblings.every(node => selected.has(node))
          ? `${first.parentNode!.nodeType === 9 ? '' : path(first.parentNode!)}/${qualified(first as Element)}`
          : selection.map(path).join(' | ')
      } else xpath = selection.map(path).join(' | ')
      const attribute = (value: string) => value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
      const prefixMappings = Array.from(prefixByNamespace, ([namespace, prefix]) => `xmlns:${prefix}="${attribute(namespace)}"`).join(' ')
      if (xpath.length > DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH || prefixMappings.length > DOCUMENT_CUSTOM_XML_MAX_XPATH_LENGTH) invalid('DOC_CUSTOM_XML_RESOURCE_LIMIT', '重定位后的 XML 绑定超过 XPath 或命名空间限额。')
      return { xpath, ...(prefixMappings ? { prefixMappings } : {}) }
    }
    const structure: DocumentCustomXmlStructureDocument = {
      resolveElements(xpath, prefixMappings) {
        return evaluate(xpath, prefixMappings, true).map(node => {
          if (node.nodeType !== 1) invalid('DOC_CUSTOM_XML_INVALID_STRUCTURE', '重复节 XPath 必须选择元素集合。')
          return { key: keyFor(node), value: valueFor(node) }
        })
      },
      contains(parent, child) { return contains(elementFor(parent), nodeFor(child)) },
      cloneElementAfter(key, trackedKeys) {
        const source = elementFor(key), parent = source.parentNode
        if (parent?.nodeType !== 1) invalid('DOC_CUSTOM_XML_INVALID_STRUCTURE', '不能复制数据源的根元素作为重复项目。')
        const tracked = Array.from(new Set([key, ...trackedKeys]), original => {
          const node = nodeFor(original)
          if (!contains(source, node)) invalid('DOC_CUSTOM_XML_INVALID_STRUCTURE', '复制映射包含重复项目之外的 XML 节点。')
          return { original, node }
        })
        const copy = source.cloneNode(true) as Element, copiedKeys = new Map<string, string>()
        parent.insertBefore(copy, source.nextSibling)
        changed = true
        for (const { original, node } of tracked) {
          const attribute = node.nodeType === 2 ? node as Attr : undefined
          let cursor: Node = attribute?.ownerElement ?? node
          const positions: number[] = []
          while (cursor !== source) {
            const owner = cursor.parentNode
            if (!owner) invalid('DOC_CUSTOM_XML_DETACHED_NODE', '复制过程中 XML 源节点已经失效。')
            positions.unshift(Array.prototype.indexOf.call(owner.childNodes, cursor))
            cursor = owner
          }
          cursor = copy
          for (const position of positions) {
            const child = cursor.childNodes[position]
            if (!child) invalid('DOC_CUSTOM_XML_INVALID_STRUCTURE', '复制的 XML 节点结构与来源不一致。')
            cursor = child
          }
          if (attribute) {
            const result = (cursor as Element).getAttributeNodeNS(attribute.namespaceURI, attribute.localName)
            if (!result) invalid('DOC_CUSTOM_XML_INVALID_STRUCTURE', '复制的 XML 属性与来源不一致。')
            cursor = result
          }
          copiedKeys.set(original, keyFor(cursor))
        }
        return { key: keyFor(copy), copiedKeys }
      },
      removeElement(key) {
        const node = elementFor(key), parent = node.parentNode
        if (parent?.nodeType !== 1) invalid('DOC_CUSTOM_XML_INVALID_STRUCTURE', '不能删除数据源根元素作为重复项目。')
        parent.removeChild(node); changed = true
      },
      bindingFor(key) { return bindingForNodes([nodeFor(key)], false) },
      bindingForElements(keys) { return bindingForNodes(Array.from(new Set(keys), elementFor), true) },
    }
    return {
      resolve(xpath, prefixMappings) {
        let node: Node | undefined = evaluate(xpath, prefixMappings, false)[0]
        if (!node) return undefined
        node = scalarNode(node)
        if (![1, 2, 3, 4, 7, 8].includes(node.nodeType)) invalid('DOC_CUSTOM_XML_INVALID_NODE', '内容控件 XPath 必须选择可写的 XML 内容节点。')
        return { key: keyFor(node), value: valueFor(node) }
      },
      read(key) { return valueFor(nodeFor(key)) },
      write(key, value) {
        const node = nodeFor(key)
        if (typeof value !== 'string' || /[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value)) invalid('DOC_CUSTOM_XML_INVALID_VALUE', '绑定值包含 XML 不能保存的字符。')
        if (documentCustomXmlUtf8Bytes(value) > DOCUMENT_CUSTOM_XML_MAX_PART_BYTES) invalid('DOC_CUSTOM_XML_RESOURCE_LIMIT', 'XML 绑定标量超过数据部件限额。')
        if (valueFor(node) === value) return
        if (node.nodeType === 1) node.textContent = value
        else node.nodeValue = value
        changed = true
      },
      overlaps(left, right) {
        const a = nodeFor(left), b = nodeFor(right)
        if (a === b) return true
        const contains = (parent: Node, target: Node) => parent.nodeType === 1 && (target.nodeType === 2
          ? (target as Attr).ownerElement !== parent && parent.contains((target as Attr).ownerElement)
          : parent.contains(target))
        return contains(a, b) || contains(b, a)
      },
      structure,
      serialize() {
        const current = liveDocument()
        if (!changed) return xml
        const serialized = new XMLSerializer().serializeToString(current)
        parseXml(serialized)
        return serialized
      },
      dispose() { document = undefined; nodes.clear(); mappingCache.clear() },
    }
  },
}
