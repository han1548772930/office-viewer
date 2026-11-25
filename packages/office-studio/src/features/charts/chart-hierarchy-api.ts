/** Shared field contract of the fixed Treemap/Sunburst hierarchy builder. */
export function chartHierarchyFields(indexes: readonly number[]): number[] {
  if (!Array.isArray(indexes) || !indexes.length || new Set(indexes).size !== indexes.length) {
    throw new RangeError('层级字段须为非空、有序且不重复的非负字段索引。')
  }
  for (let position = 0; position < indexes.length; position++) {
    const index = indexes[position]
    if (!Object.hasOwn(indexes, position) || index === undefined || !Number.isSafeInteger(index) || index < 0 || index > 1_048_575) {
      throw new RangeError('层级字段须为非空、有序且不重复的非负字段索引。')
    }
  }
  return [...indexes]
}

export function chartHierarchyWeightField(index: number): number {
  if (!Number.isSafeInteger(index) || index < 0 || index > 1_048_575) throw new RangeError('层级权重字段需要非负字段索引。')
  return index
}
