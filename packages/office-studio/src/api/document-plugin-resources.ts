import { cloneModel } from '../domain/serialization'
import { bodyRecord, type UniverBodyRecord } from './document-body-style'
import type { UniverDocumentBodyData, UniverDocumentBodyResources } from './document-body-types'
import { isDocumentBodyLatexRange } from './document-body-latex'

const PLUGINS = [['DOC_CALLOUT_PLUGIN', 'callouts', 'callouts'], ['DOC_CODE_PLUGIN', 'codes', 'codes'],
  ['DOC_LATEX_PLUGIN', 'latexFormulas', 'formulas']] as const
type PluginKey = typeof PLUGINS[number][1]
type PluginResources = Pick<UniverDocumentBodyResources, PluginKey>

/** Fixed public plugins serialize their configuration in named JSON resources. */
export function readDocumentPluginResources(value: unknown): PluginResources {
  if (value === undefined) return {}
  if (!Array.isArray(value)) throw new TypeError('文档插件资源必须为数组。')
  const result: PluginResources = {}, seen = new Set<string>()
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
    const configs = Reflect.get(data, plugin[2])
    if (!configs || typeof configs !== 'object' || Array.isArray(configs)) throw new TypeError('文档高级块配置映射无效。')
    for (const [id, config] of Object.entries(configs)) if (!id.trim() || !config || typeof config !== 'object' || Array.isArray(config)) throw new TypeError('文档高级块配置必须以稳定标识存储。')
    result[plugin[1]] = cloneModel(configs as Record<string, UniverBodyRecord>)
  }
  return result
}

/** Only live ranges contribute modeled configs; deleting a block removes its plugin entry. */
export function writeDocumentPluginResources(source: UniverBodyRecord, bodies: readonly UniverDocumentBodyData[]): UniverBodyRecord[] | undefined {
  const sourceBodies = [source, ...['headers', 'footers', 'notes', 'footnotes', 'endnotes'].flatMap(key => Object.values(bodyRecord(source[key])).map(bodyRecord))]
  const configs: Record<PluginKey, Record<string, UniverBodyRecord>> = { callouts: {}, codes: {}, latexFormulas: {} }
  const modeled = { callouts: new Set<string>(), codes: new Set<string>(), latexFormulas: new Set<string>() }
  for (const value of sourceBodies) {
    const original = readDocumentPluginResources(value.resources)
    for (const [, key] of PLUGINS) for (const [id, config] of Object.entries(original[key] ?? {})) {
      const previous = configs[key][id]
      if (previous && JSON.stringify(previous) !== JSON.stringify(config)) throw new RangeError('同一插件身份对应不同来源配置：' + id)
      configs[key][id] = cloneModel(config)
    }
    for (const item of Array.isArray(bodyRecord(value.body).customRanges) ? bodyRecord(value.body).customRanges as unknown[] : []) {
      const range = bodyRecord(item)
      if (isDocumentBodyLatexRange(range) && typeof range.rangeId === 'string') modeled.latexFormulas.add(range.rangeId)
    }
  }
  for (const value of sourceBodies) for (const item of Array.isArray(bodyRecord(value.body).blockRanges) ? bodyRecord(value.body).blockRanges as unknown[] : []) {
    const raw = bodyRecord(item)
    if (typeof raw.blockId === 'string' && (raw.blockType === 'callout' || raw.blockType === 'code')) modeled[raw.blockType === 'callout' ? 'callouts' : 'codes'].add(raw.blockId)
  }
  for (const [, key] of PLUGINS) for (const id of modeled[key]) delete configs[key][id]
  for (const data of bodies) for (const range of data.body.blockRanges ?? []) {
    const key = range.blockType === 'callout' ? 'callouts' : range.blockType === 'code' ? 'codes' : undefined
    if (!key) continue
    const config = data[key]?.[range.blockId]
    if (!config) throw new RangeError('高级块缺少要保存的插件配置：' + range.blockId)
    const previous = configs[key][range.blockId]
    if (previous && JSON.stringify(previous) !== JSON.stringify(config)) throw new RangeError('同一高级块身份对应不同配置：' + range.blockId)
    configs[key][range.blockId] = cloneModel(config)
  }
  for (const data of bodies) for (const range of data.body.customRanges ?? []) {
    if (!isDocumentBodyLatexRange(range)) continue
    const config = data.latexFormulas?.[range.rangeId]
    if (!config) throw new RangeError('公式缺少要保存的插件配置：' + range.rangeId)
    const previous = configs.latexFormulas[range.rangeId]
    if (previous && JSON.stringify(previous) !== JSON.stringify(config)) throw new RangeError('同一公式身份对应不同配置：' + range.rangeId)
    configs.latexFormulas[range.rangeId] = cloneModel(config)
  }
  const resources = (Array.isArray(source.resources) ? source.resources : []).map(bodyRecord)
    .filter(item => !PLUGINS.some(([name]) => name === item.name)).map(cloneModel)
  for (const [name, key, jsonKey] of PLUGINS) {
    const raw = sourceBodies.flatMap(value => Array.isArray(value.resources) ? value.resources.map(bodyRecord) : []).find(value => value.name === name)
    const wrapper = raw && typeof raw.data === 'string' ? bodyRecord(JSON.parse(raw.data)) : bodyRecord(raw?.data)
    if (!Object.keys(configs[key]).length && !Object.keys(wrapper).some(field => field !== jsonKey)
      && !Object.keys(raw ?? {}).some(field => field !== 'name' && field !== 'data')) continue
    resources.push({ ...cloneModel(raw ?? {}), name, data: JSON.stringify({ ...wrapper, [jsonKey]: configs[key] }) })
  }
  return resources.length ? resources : undefined
}

/** Retain every nested wrapper's unknown fields while replacing its stale config map. */
export function refreshDocumentLatexPluginWrapper(value: UniverBodyRecord, live: Record<string, UniverBodyRecord>): void {
  if (!Array.isArray(value.resources)) return
  value.resources = value.resources.map(bodyRecord).map(raw => {
    if (raw.name !== 'DOC_LATEX_PLUGIN') return raw
    const wrapper = typeof raw.data === 'string' ? bodyRecord(JSON.parse(raw.data)) : bodyRecord(raw.data)
    const ids = new Set(Object.keys(bodyRecord(wrapper.formulas)))
    for (const entry of Array.isArray(bodyRecord(value.body).customRanges) ? bodyRecord(value.body).customRanges as unknown[] : []) {
      const range = bodyRecord(entry)
      if (isDocumentBodyLatexRange(range) && typeof range.rangeId === 'string') ids.add(range.rangeId)
    }
    const formulas = Object.fromEntries([...ids].filter(id => live[id] !== undefined).map(id => [id, cloneModel(live[id])]))
    return { ...cloneModel(raw), data: JSON.stringify({ ...cloneModel(wrapper), formulas }) }
  })
}
