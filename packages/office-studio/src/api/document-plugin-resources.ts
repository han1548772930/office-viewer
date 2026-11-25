import { cloneModel } from '../domain/serialization'
import { bodyRecord, type UniverBodyRecord } from './document-body-style'
import type { UniverDocumentBodyData, UniverDocumentBodyResources } from './document-body-types'

const PLUGINS = [['DOC_CALLOUT_PLUGIN', 'callouts'], ['DOC_CODE_PLUGIN', 'codes']] as const

/** Fixed public plugins serialize their configuration in named JSON resources. */
export function readDocumentPluginResources(value: unknown): Pick<UniverDocumentBodyResources, 'callouts' | 'codes'> {
  if (value === undefined) return {}
  if (!Array.isArray(value)) throw new TypeError('文档插件资源必须为数组。')
  const result: Pick<UniverDocumentBodyResources, 'callouts' | 'codes'> = {}, seen = new Set<string>()
  for (const item of value) {
    const raw = bodyRecord(item), plugin = PLUGINS.find(([name]) => name === raw.name)
    if (!plugin) continue
    if (seen.has(plugin[0])) throw new TypeError('文档高级块插件资源重复：' + plugin[0])
    seen.add(plugin[0])
    let data: unknown = raw.data
    if (typeof data === 'string') {
      if (data.length > 16 * 1024 * 1024) throw new RangeError('文档插件资源超过单次文本预算。')
      try { data = JSON.parse(data) } catch { throw new TypeError('文档高级块插件资源不是合法 JSON。') }
    }
    if (!data || typeof data !== 'object' || Array.isArray(data)) throw new TypeError('文档高级块插件资源缺少配置对象。')
    const configs = Reflect.get(data, plugin[1])
    if (!configs || typeof configs !== 'object' || Array.isArray(configs)) throw new TypeError('文档高级块配置映射无效。')
    for (const [id, config] of Object.entries(configs)) if (!id.trim() || !config || typeof config !== 'object' || Array.isArray(config)) throw new TypeError('文档高级块配置必须以稳定标识存储。')
    result[plugin[1]] = cloneModel(configs as Record<string, UniverBodyRecord>)
  }
  return result
}

/** Only live ranges contribute modeled configs; deleting a block removes its plugin entry. */
export function writeDocumentPluginResources(source: UniverBodyRecord, bodies: readonly UniverDocumentBodyData[]): UniverBodyRecord[] | undefined {
  const original = readDocumentPluginResources(source.resources), modeled = { callouts: new Set<string>(), codes: new Set<string>() }
  const sourceBodies = [source, ...['headers', 'footers', 'notes', 'footnotes', 'endnotes'].flatMap(key => Object.values(bodyRecord(source[key])).map(bodyRecord))]
  for (const value of sourceBodies) for (const item of Array.isArray(bodyRecord(value.body).blockRanges) ? bodyRecord(value.body).blockRanges as unknown[] : []) {
    const raw = bodyRecord(item)
    if (typeof raw.blockId === 'string' && (raw.blockType === 'callout' || raw.blockType === 'code')) modeled[raw.blockType === 'callout' ? 'callouts' : 'codes'].add(raw.blockId)
  }
  const configs = { callouts: { ...original.callouts }, codes: { ...original.codes } }
  for (const key of ['callouts', 'codes'] as const) for (const id of modeled[key]) delete configs[key][id]
  for (const data of bodies) for (const range of data.body.blockRanges ?? []) {
    const key = range.blockType === 'callout' ? 'callouts' : range.blockType === 'code' ? 'codes' : undefined
    if (!key) continue
    const config = data[key]?.[range.blockId]
    if (!config) throw new RangeError('高级块缺少要保存的插件配置：' + range.blockId)
    const previous = configs[key][range.blockId]
    if (previous && JSON.stringify(previous) !== JSON.stringify(config)) throw new RangeError('同一高级块身份对应不同配置：' + range.blockId)
    configs[key][range.blockId] = cloneModel(config)
  }
  const resources = (Array.isArray(source.resources) ? source.resources : []).map(bodyRecord)
    .filter(item => !PLUGINS.some(([name]) => name === item.name)).map(cloneModel)
  for (const [name, key] of PLUGINS) if (Object.keys(configs[key]).length) resources.push({ name, data: JSON.stringify({ [key]: configs[key] }) })
  return resources.length ? resources : undefined
}
