import type { DocumentBlock, DocumentContentControl, DocumentStoryContent } from '../domain/model'
import { tableForBlock } from '../features/document/model/document-model'
import { documentTextFields } from '../features/document/model/document-text-fields'
import { documentBlockChildContainers } from '../features/document/model/document-structural-blocks'
import { documentBodyIdentity } from './document-body-types'

/** Namespaces match the current domain's copy maps and addressable identities. */
export function documentBodyContentIdentities(blocks: readonly DocumentBlock[], controls: readonly DocumentContentControl[]): string[] {
  const ids: string[] = []
  const add = (kind: string, id: string) => ids.push(`${kind}:${documentBodyIdentity(id, kind)}`)
  const visit = (blocks: readonly DocumentBlock[]) => {
    for (const block of blocks) {
      add('block', block.id)
      for (const run of block.runs) add('run', run.id)
      for (const field of documentTextFields(block)) add('field', field.id)
      for (const row of tableForBlock(block)?.rows ?? []) {
        add('row', row.id)
        for (const cell of row.cells) {
          add('cell', cell.id)
        }
      }
      for (const column of block.columns?.columns ?? []) add('column', column.id)
      for (const children of documentBlockChildContainers(block)) visit(children)
    }
  }
  visit(blocks)
  for (const control of controls) add('control', control.id)
  return ids
}
export function assertDocumentBodyContentIdentities(content: DocumentStoryContent, controls: readonly DocumentContentControl[]): void {
  const ids = documentBodyContentIdentities(content.blocks ?? [], controls), own = new Set<string>()
  for (const id of ids) {
    if (own.has(id)) throw new TypeError('文档内容区身份重复：' + id)
    own.add(id)
  }
}
