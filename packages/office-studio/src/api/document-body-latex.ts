import { cloneModel } from '../domain/serialization'
import { assertLatexFormulaSource, DOCUMENT_LATEX_RANGE_KIND, type LatexFormula, type LatexFormulaMode } from '../features/document/latex/latex-formula'
import { isDocumentTextBoundary } from '../features/document/model/document-story'
import { bodyRecord, type UniverBodyRecord } from './document-body-style'
import { documentBodyArray, documentBodyIdentity, documentBodyOffset, type UniverDocumentBodyResources } from './document-body-types'

export interface DocumentBodyLatexRange {
  start: number
  end: number
  rangeId: string
  formula: LatexFormula
  source: UniverBodyRecord
  config: UniverBodyRecord
}

export function isDocumentBodyLatexRange(raw: UniverBodyRecord): boolean {
  return raw.rangeType === 5 && bodyRecord(raw.properties).kind === DOCUMENT_LATEX_RANGE_KIND
}

export function documentBodyLatexConfig(latex: string, source?: unknown): UniverBodyRecord {
  return { ...cloneModel(bodyRecord(source)), latex, kind: 'inline', displayMode: 'inline' }
}

/** Source text remains public UTF-16; the editor represents the range as one atom. */
export function readDocumentBodyLatexRanges(value: UniverBodyRecord, resources: UniverDocumentBodyResources): Map<number, DocumentBodyLatexRange> {
  const stream = typeof value.dataStream === 'string' ? value.dataStream : '', ranges: DocumentBodyLatexRange[] = [], ids = new Set<string>()
  for (const raw of documentBodyArray(value.customRanges, 'customRanges')) {
    if (!isDocumentBodyLatexRange(raw)) continue
    const start = documentBodyOffset(raw.startIndex, stream.length, 'LaTeX.startIndex')
    const end = documentBodyOffset(raw.endIndex, stream.length - 1, 'LaTeX.endIndex') + 1
    if (end <= start || !isDocumentTextBoundary(stream, start) || !isDocumentTextBoundary(stream, end)) throw new RangeError('公式范围必须是完整的 UTF-16 文本区间。')
    const rangeId = documentBodyIdentity(raw.rangeId, '公式范围')
    if (ids.has(rangeId)) throw new TypeError('公式范围标识重复：' + rangeId)
    ids.add(rangeId)
    if (raw.wholeEntity !== true) throw new TypeError('公式范围必须声明 wholeEntity。')
    const latex = stream.slice(start, end), config = bodyRecord(resources.latexFormulas?.[rangeId])
    if (typeof config.latex === 'string' && config.latex !== latex) throw new RangeError('公式源文本与插件配置不一致：' + rangeId)
    if (config.latex !== undefined && typeof config.latex !== 'string'
      || config.kind !== undefined && config.kind !== 'inline' || config.displayMode !== undefined && config.displayMode !== 'inline') throw new TypeError('固定版本的公式公开配置只接受 inline。')
    assertLatexFormulaSource(latex)
    const extension = bodyRecord(raw.officeStudio)
    const mode: LatexFormulaMode = extension.latexMode === 'block' ? 'block' : 'inline'
    if (extension.latexMode !== undefined && extension.latexMode !== 'inline' && extension.latexMode !== 'block') throw new TypeError('公式本地显示形式无效。')
    ranges.push({ start, end, rangeId, formula: { latex, mode, version: 1 }, source: cloneModel(raw), config: documentBodyLatexConfig(latex, config) })
  }
  ranges.sort((a, b) => a.start - b.start)
  for (let index = 1; index < ranges.length; index++) if (ranges[index]!.start < ranges[index - 1]!.end) throw new RangeError('公式实体范围不能重叠。')
  return new Map(ranges.map(range => [range.start, range]))
}
