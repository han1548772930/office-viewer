import { cloneModel } from '../domain/serialization'
import type { DocumentBlock } from '../domain/model'
import { blockHasNumbering, blockNumberingDefinition, blockNumberingId } from '../features/document/model/document-numbering'
import { blockListPreset } from '../features/document/model/document-lists'
import { tableForBlock } from '../features/document/model/document-model'
import { documentTableCellContentContainers } from '../features/document/model/document-table-continuations'
import type { DocumentBodyListResource } from './document-body-paragraph'
import { bodyRecord, type UniverBodyRecord } from './document-body-style'

/** Old adjacent lists without IDs still describe one sequence in the editor.
 * Assign its first stable paragraph ID once for public save/reopen. */
export function documentBodyListInstanceIds(blocks: readonly DocumentBlock[]): ReadonlyMap<string, string> {
  const result = new Map<string, string>()
  const visit = (container: readonly DocumentBlock[]) => {
    let previous: string | undefined, listId = ''
    for (const block of container) {
      const table = tableForBlock(block)
      if (table) {
        previous = undefined
        for (const row of table.rows) for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) visit(content)
        continue
      }
      if (!blockHasNumbering(block)) { previous = undefined; continue }
      const id = blockNumberingId(block)
      if (id) { result.set(block.id, id); previous = undefined; continue }
      const style = String(blockNumberingDefinition(block)?.abstract.abstractNumId ?? blockListPreset(block) ?? block.type)
      if (previous !== style) listId = `office-legacy:${block.id}`
      previous = style; result.set(block.id, listId)
    }
  }
  visit(blocks)
  return result
}

/** Current list resources merge all instances instead of letting the last
 * paragraph replace earlier restart offsets and shared-counter membership. */
export class DocumentBodyListResources {
  readonly values: Record<string, UniverBodyRecord> = {}
  private readonly groups = new Map<string, { types: Set<string>; instances: Set<string> }>()

  add(resource: DocumentBodyListResource): void {
    const previous = this.values[resource.id], current = cloneModel(resource.value)
    if (previous) {
      const starts = { ...bodyRecord(previous.counterStarts) }
      for (const [id, value] of Object.entries(bodyRecord(current.counterStarts))) {
        const before = starts[id]
        if (before !== undefined && !sameValue(before, value)) throw new RangeError('同一列表实例的重启编号互相冲突：' + id)
        starts[id] = cloneModel(value)
      }
      const withoutCounters = (value: UniverBodyRecord) => {
        const { counterIds: _ids, counterStarts: _starts, ...definition } = value
        return definition
      }
      if (!sameValue(withoutCounters(previous), withoutCounters(current))) throw new RangeError('同一公开列表类型对应不同层级定义：' + resource.id)
      if (Object.keys(starts).length) current.counterStarts = starts
    }
    this.values[resource.id] = current
    if (resource.counterId !== undefined) {
      const group = this.groups.get(resource.counterId) ?? { types: new Set<string>(), instances: new Set<string>() }
      group.types.add(resource.id); group.instances.add(resource.listId)
      this.groups.set(resource.counterId, group)
    }
  }

  finish(): Record<string, UniverBodyRecord> {
    const owners = new Map<string, string>()
    for (const [id, group] of this.groups) {
      const counterIds = [id, ...[...group.instances].filter(instance => instance !== id)]
      for (const type of group.types) {
        const owner = owners.get(type)
        if (owner !== undefined && owner !== id) throw new RangeError('不同计数器组不能覆盖同一公开列表类型：' + type)
        owners.set(type, id); this.values[type]!.counterIds = counterIds
      }
    }
    return this.values
  }
}

function sameValue(left: unknown, right: unknown): boolean {
  if (left === right) return true
  if (Array.isArray(left) || Array.isArray(right)) return Array.isArray(left) && Array.isArray(right)
    && left.length === right.length && left.every((value, index) => sameValue(value, right[index]))
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object') return false
  const a = bodyRecord(left), b = bodyRecord(right)
  const keys = Object.keys(a).filter(key => a[key] !== undefined), other = Object.keys(b).filter(key => b[key] !== undefined)
  return keys.length === other.length && keys.every(key => Object.hasOwn(b, key) && sameValue(a[key], b[key]))
}
