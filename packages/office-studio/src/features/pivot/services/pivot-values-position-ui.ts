import type { PivotFieldArea, PivotFieldModel, PivotTableOptions, PivotValuePosition } from '../../../domain/model.ts'
import { cloneModel } from '../../../domain/serialization.ts'
import type { PivotChartFieldsConfig } from './pivot-chart-engine.ts'
import type { PivotPanelDraft } from './pivot-panel-state.ts'
import { adjustPivotValuesIndexForFieldEdit, getPivotValuesLayout, getPivotValuesPosition,
  normalizePivotValuesPosition, retainPivotValuesIndexAfterRemoval, type PivotValuesPositionState } from './pivot-values-position.ts'

type FieldOrder = Pick<PivotFieldModel, 'id' | 'area' | 'position'>
type ValuesOptions = Pick<PivotTableOptions, 'valuePosition' | 'valueIndex'>
export interface PivotValuesUiState extends PivotValuesPositionState {
  area: 'row' | 'column'
  effectiveIndex: number
  dimensionCount: number
  measureCount: number
  visible: boolean
}
export type PivotValuesUiEntry<T> = { kind: 'field'; field: T } | { kind: 'values'; values: PivotValuesUiState }

/** Layout/fallback/clamping belong to the common engine; this projection only inserts the UI marker. */
export function readPivotValuesUiState(options: ValuesOptions, fields: readonly Pick<PivotFieldModel, 'area'>[]): PivotValuesUiState {
  const state = getPivotValuesPosition({ options }), layout = getPivotValuesLayout({ options, fields })
  const measureCount = fields.filter(field => field.area === 'value').length
  return { ...state, area: layout.area, effectiveIndex: layout.index,
    dimensionCount: fields.filter(field => field.area === layout.area).length, measureCount,
    visible: state.position !== 'none' && measureCount > 1 }
}

export function pivotValuesUiEntries<T>(fields: readonly T[], state: PivotValuesUiState, area: PivotFieldArea): PivotValuesUiEntry<T>[] {
  const entries: PivotValuesUiEntry<T>[] = fields.map(field => ({ kind: 'field', field }))
  if (state.visible && area === state.position) entries.splice(state.effectiveIndex, 0, { kind: 'values', values: state })
  return entries
}

export function pivotValuesUiIndexOptions(state: PivotValuesUiState): Array<{ value: number; label: string }> {
  const entries = [{ value: -1, label: '自动末尾' }, ...Array.from({ length: state.dimensionCount + 1 }, (_, index) =>
    ({ value: index, label: index === 0 ? '第 1 个位置（首位）' : index === state.dimensionCount ? `第 ${index + 1} 个位置（末位）` : `第 ${index + 1} 个位置` }))]
  if (state.index > state.dimensionCount) entries.push({ value: state.index, label: `指定位置 ${state.index + 1}（当前末位）` })
  return entries
}

/** Top-level public aliases and options must agree when a chart draft is submitted. */
export function pivotChartValuesUiOptions(config: PivotChartFieldsConfig): ValuesOptions {
  if (config.valuePosition !== undefined && config.options?.valuePosition !== undefined && config.valuePosition !== config.options.valuePosition
    || config.valueIndex !== undefined && config.options?.valueIndex !== undefined && config.valueIndex !== config.options.valueIndex) {
    throw new TypeError('透视图值字段位置草稿有冲突。')
  }
  const state = normalizePivotValuesPosition(config.valuePosition ?? config.options?.valuePosition ?? 'column',
    config.valueIndex ?? config.options?.valueIndex ?? -1)
  return { valuePosition: state.position, valueIndex: state.index }
}

export function setPivotChartValuesUiDraft(config: PivotChartFieldsConfig, position: PivotValuePosition, index = -1): PivotChartFieldsConfig {
  const state = normalizePivotValuesPosition(position, index), next = cloneModel(config)
  next.options = { ...next.options, valuePosition: state.position, valueIndex: state.index }
  next.valuePosition = state.position; next.valueIndex = state.index
  return next
}

export function setPivotTableValuesUiDraft(draft: PivotPanelDraft, position: PivotValuePosition, index = -1): void {
  const state = normalizePivotValuesPosition(position, index)
  draft.options = { ...draft.options, valuePosition: state.position, valueIndex: state.index }
}

export function pivotTableValuesUiFieldOrder(draft: Pick<PivotPanelDraft, 'fields'>): FieldOrder[] {
  const positions: Record<PivotFieldArea, number> = { row: 0, column: 0, filter: 0, value: 0 }
  return draft.fields.flatMap(field => field.area === 'unused' ? [] :
    [{ id: field.fieldId, area: field.area, position: positions[field.area]++ }])
}

/** Draft field moves use the same gap adjustment as committed fields, rather than another axis algorithm. */
export function adjustPivotTableValuesUiFieldEdit(draft: PivotPanelDraft, previous: readonly FieldOrder[], fieldId?: string): void {
  const state = getPivotValuesPosition({ options: draft.options })
  if (state.position === 'none' || state.index === -1) return
  const projected = { fields: pivotTableValuesUiFieldOrder(draft), options: { ...draft.options } }
  if (fieldId) adjustPivotValuesIndexForFieldEdit(projected, previous, fieldId)
  else retainPivotValuesIndexAfterRemoval(projected, previous)
  if (projected.options.valueIndex !== draft.options.valueIndex) draft.options = { ...draft.options, valueIndex: projected.options.valueIndex }
}

export function adjustPivotChartValuesUiFieldEdit(config: PivotChartFieldsConfig, previous: readonly FieldOrder[], fieldId?: string): PivotChartFieldsConfig {
  const options = pivotChartValuesUiOptions(config), state = getPivotValuesPosition({ options })
  if (state.position === 'none' || state.index === -1) return config
  const projected = { fields: config.fields, options: { ...options } }
  if (fieldId) adjustPivotValuesIndexForFieldEdit(projected, previous, fieldId)
  else retainPivotValuesIndexAfterRemoval(projected, previous)
  return projected.options.valueIndex === state.index ? config : setPivotChartValuesUiDraft(config, state.position, projected.options.valueIndex)
}
