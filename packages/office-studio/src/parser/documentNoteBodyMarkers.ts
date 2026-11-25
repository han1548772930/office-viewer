import type { DocumentBlock } from '../domain/model'
import { runText, tableForBlock } from '../features/document/model/document-model'
import { documentTableCellContentContainers } from '../features/document/model/document-table-continuations'
import { record } from './validation'

/** Remove only annotated custom-marker fragments. Each parsed inline keeps
 * its source index so fields, SDTs and body markers share one projection. */
export function removeParsedDocumentNoteBodyMarkerText(blocks: readonly DocumentBlock[], kind: 'footnote' | 'endnote'): void {
  for (const block of blocks) {
    const grouped = new Map<number, DocumentBlock['runs']>()
    for (const run of block.runs) {
      const source = record(run.metadata?.officeParsedRunSource), marker = record(run.metadata?.noteBodyMarker)
      if (!Object.keys(marker).length) continue
      if (marker.version !== 1 || marker.kind !== kind || typeof marker.text !== 'string' || marker.text.length > 4096
        || marker.id !== undefined && (typeof marker.id !== 'string' || !marker.id.trim())
        || source.version !== 1 || typeof source.index !== 'number' || !Number.isSafeInteger(source.index) || source.index < 0) {
        throw new TypeError('原生注释正文标记来源无效。')
      }
      const group = grouped.get(source.index) ?? []
      group.push(run); grouped.set(source.index, group)
    }
    for (const runs of grouped.values()) {
      const marker = record(runs[0]!.metadata?.noteBodyMarker), text = marker.text as string
      if (runs.some(run => record(run.metadata?.noteBodyMarker).text !== text || record(run.metadata?.noteBodyMarker).id !== marker.id)) {
        throw new TypeError('同一原生 inline 的正文标记来源不一致。')
      }
      if (runs[0]!.metadata?.officeNoteBodyMarkerSource !== undefined) {
        const removed = record(runs[0]!.metadata.officeNoteBodyMarkerSource)
        if (removed.version !== 1 || removed.text !== text || removed.length !== text.length) throw new TypeError('已投影的原生正文标记来源不一致。')
        continue
      }
      if (!runText(runs).startsWith(text)) throw new RangeError('原生注释正文标记与来源文字不一致。')
      let remaining = text.length
      for (const run of runs) {
        const consumed = Math.min(remaining, run.text.length)
        run.text = run.text.slice(consumed); remaining -= consumed
      }
      runs[0]!.metadata = { ...runs[0]!.metadata, officeNoteBodyMarkerSource: { version: 1, text, length: text.length,
        ...(marker.id === undefined ? {} : { id: marker.id }) } }
    }
    for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) {
      for (const content of documentTableCellContentContainers(cell)) removeParsedDocumentNoteBodyMarkerText(content, kind)
    }
  }
}
