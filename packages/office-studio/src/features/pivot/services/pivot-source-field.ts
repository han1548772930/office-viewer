import type { PivotCacheModel, PivotFieldModel, PivotTableModel } from '../../../domain/model.ts'

/** The retained cache owns source-field identity. Report-local instance IDs
 * and copied physical columns cannot replace an existing canonical identity. */
export function pivotSourceFieldIdentity(report: PivotTableModel, field: PivotFieldModel, cache?: PivotCacheModel): string {
  const columnFields = cache?.fields?.filter(candidate => candidate.sourceColumn === field.sourceColumn) ?? []
  if (columnFields.length > 1) throw new TypeError('透视缓存中的同一来源列存在不同字段身份。')
  const canonical = columnFields[0]
  if (field.sourceFieldId !== undefined) {
    const retained = cache?.fields?.find(candidate => candidate.id === field.sourceFieldId)
    if (retained && retained.sourceColumn !== field.sourceColumn || canonical && canonical.id !== field.sourceFieldId) {
      throw new TypeError('透视报告字段与缓存来源列不一致。')
    }
    return field.sourceFieldId
  }
  return canonical?.id ?? String(field.sourceColumn - report.source.range.startCol)
}
