<script setup lang="ts">
import { computed, reactive, shallowRef, watch } from 'vue'
import { Button, Checkbox, ColorPicker, Form, FormItem, Input, InputNumber, Modal, Select, Switch } from 'antdv-next'
import { Plus, Trash2 } from 'lucide-vue-next'
import type { DocumentModel } from '../../domain/model'
import { documentId, findDocumentBlock, tableForBlock, type DocumentSelection, type DocumentTableCellMargin,
  type DocumentTableColumnTypeConfig, type DocumentTableBorder } from '../../features/document/model/document-model'
import { documentTableCellsInRange, documentTableGrid, expandDocumentTableRange, type DocumentTableCellRange } from '../../features/document/model/document-table-grid'
import { DOCUMENT_TABLE_BORDER_SIDES, documentTableCellFill, documentTableCellMargin, documentTableColumnWidths,
  documentTableCommonValue, documentTableRowProperties } from '../../features/document/model/document-table-style'
import { documentTableCellBorder, type DocumentTableBorderPreset } from '../../features/document/model/document-table-style'
import { DOCUMENT_TABLE_BORDER_LABELS } from './document-table-ribbon'
import { documentTableReferenceWidth } from '../../features/document/commands/table-mutations'
import { documentTableCommandState } from '../../features/document/commands/table-commands'
import type { DocumentTableCommand, DocumentTableStylePatch } from '../../features/document/commands/table-command-types'
import type { DocumentPermissionContext } from '../../features/document/permissions/document-permissions'
import type { DocumentLayout } from '../../features/document/layout/document-layout-types'

export type DocumentTablePropertiesKind = 'table' | 'row' | 'column' | 'cell' | 'header' | 'type' | 'move' | 'background' | 'border'
const props = defineProps<{
  open: boolean; model: DocumentModel; selection: DocumentSelection; permission: DocumentPermissionContext
  tableId: string; range: DocumentTableCellRange; kind: DocumentTablePropertiesKind; type?: DocumentTableColumnTypeConfig['type']; composing?: boolean
  layout?: DocumentLayout
}>()
const emit = defineEmits<{ 'update:open': [value: boolean]; submit: [command: DocumentTableCommand]; closed: [] }>()
const requestModel = shallowRef<DocumentModel>()
const requestRange = shallowRef<DocumentTableCellRange>()
const requestWidths = shallowRef<number[]>()
const hasMeasuredWidths = shallowRef(false)
const titles: Record<DocumentTablePropertiesKind, string> = { table: '表格属性', row: '行高与跨页', column: '列宽', cell: '单元格边距', header: '固定表头', type: '列类型选项', move: '移动表格或行列', background: '单元格背景', border: '单元格边框' }
const modern = computed(() => props.model.metadata?.documentFlavor === 'modern' || props.model.metadata?.documentFlavor === 2 || props.model.metadata?.mode === 'modern')
const source = computed(() => { const block = findDocumentBlock(props.model, props.tableId); return block && tableForBlock(block) })
const columnOptions = computed(() => source.value ? Array.from({ length: documentTableGrid(source.value).columnCount }, (_, column) => ({ value: column, label: `第 ${column + 1} 列` })) : [])
type PropertyNumber = number | undefined | null
interface TablePropertiesForm {
  width: PropertyNumber; height: PropertyNumber; align: 'left' | 'center' | 'right'; indent: number; offsetY: number; layout: 'auto' | 'fixed'
  heightRule: 'auto' | 'atLeast' | 'exact' | undefined; allowSplit: boolean | undefined; inheritMargin: boolean | undefined; margin: Record<keyof DocumentTableCellMargin, PropertyNumber>
  changeBackground: boolean; background: string | undefined; clearBackground: boolean; changeBorder: boolean; borderColor: string | undefined; borderWidth: PropertyNumber; borderStyle: 'solid' | 'dot' | 'dash' | undefined; borderPreset: DocumentTableBorderPreset
  headerCount: number; type: DocumentTableColumnTypeConfig['type']; column: number; options: Array<{ id: string; label: string; color?: string }>
  moveKind: 'physical' | 'offset' | 'row' | 'column'; deltaX: number; deltaY: number; targetOffset: number; targetIndex: number; position: 'before' | 'after'
}
const form = reactive<TablePropertiesForm>({ width: 300, height: undefined, align: 'left', indent: 0, offsetY: 0, layout: 'fixed', heightRule: 'auto', allowSplit: true,
  inheritMargin: true, margin: { top: 8, right: 8, bottom: 8, left: 8 }, changeBackground: false, background: '#ffffff', clearBackground: false,
  changeBorder: false, borderColor: '#aab3bf', borderWidth: 0.75, borderStyle: 'solid', borderPreset: 'all', headerCount: 0, type: 'text', column: 0, options: [],
  moveKind: 'physical', deltaX: 0, deltaY: 0, targetOffset: 0, targetIndex: 1, position: 'before' })
const requestForm = shallowRef<TablePropertiesForm>()
const mixed = reactive({ width: false, height: false, heightRule: false, allowSplit: false, inheritMargin: false, background: false,
  borderColor: false, borderWidth: false, borderStyle: false, margin: { top: false, right: false, bottom: false, left: false } })
const changed = <K extends keyof TablePropertiesForm>(key: K) => !!requestForm.value && !Object.is(form[key], requestForm.value[key])
const marginChanged = (side: keyof DocumentTableCellMargin) => !!requestForm.value && !Object.is(form.margin[side], requestForm.value.margin[side])

watch(() => [props.open, props.tableId, props.kind, props.range.startRow, props.range.endRow, props.range.startColumn, props.range.endColumn] as const, () => {
  if (!props.open || !source.value) return
  const table = source.value, grid = documentTableGrid(table), column = props.range.startColumn
  const range = expandDocumentTableRange(grid, props.range), owners = documentTableCellsInRange(grid, range), config = table.columnTypes?.[column]
  const fragment = props.layout?.tableFragments?.find(fragment => fragment.tableId === props.tableId && !fragment.repeatedHeader)
  const measured = fragment?.columnWidths
  hasMeasuredWidths.value = !!measured && measured.length === grid.columnCount && measured.every(width => Number.isFinite(width) && width > 0)
  const widths = hasMeasuredWidths.value ? [...measured!] : documentTableColumnWidths(table, grid.columnCount, documentTableReferenceWidth(props.model, props.tableId))
  requestWidths.value = [...widths]
  requestModel.value = props.model; requestRange.value = { ...props.range }
  const width = documentTableCommonValue(widths.slice(range.startColumn, range.endColumn + 1))
  mixed.width = props.kind === 'column' && width.mixed
  form.width = props.kind === 'column' ? width.value : widths.reduce((sum, width) => sum + width, 0)
  const rows = Array.from({ length: range.endRow - range.startRow + 1 }, (_, offset) => documentTableRowProperties(table, range.startRow + offset))
  const height = documentTableCommonValue(rows.map(row => row.height)), rule = documentTableCommonValue(rows.map(row => row.heightRule)), split = documentTableCommonValue(rows.map(row => row.allowSplit))
  form.height = height.value; form.heightRule = rule.value; form.allowSplit = split.value
  mixed.height = height.mixed; mixed.heightRule = rule.mixed; mixed.allowSplit = split.mixed
  form.align = table.style?.align ?? 'left'; form.indent = table.style?.indent ?? 0; form.offsetY = table.style?.offsetY ?? 0; form.layout = table.style?.layout ?? 'fixed'
  const inherited = documentTableCommonValue(props.kind === 'cell' ? owners.map(owner => !owner.cell.style.margin) : [!table.style?.cellMargin])
  form.inheritMargin = inherited.value; mixed.inheritMargin = inherited.mixed
  const margins = props.kind === 'cell' ? owners.map(owner => documentTableCellMargin(table, owner.cell)) : [documentTableCellMargin(table)]
  for (const side of DOCUMENT_TABLE_BORDER_SIDES) { const value = documentTableCommonValue(margins.map(margin => margin[side])); form.margin[side] = value.value; mixed.margin[side] = value.mixed }
  form.headerCount = table.headerRowCount ?? 0; form.column = column; form.type = props.type ?? config?.type ?? 'text'
  form.options = (config?.dropdownOptions ?? []).map(option => ({ ...option }))
  const background = documentTableCommonValue(props.kind === 'background' ? owners.map(owner => documentTableCellFill(table, owner.cell)) : [table.style?.fill ?? table.style?.background ?? '#ffffff'])
  form.changeBackground = false; form.background = background.value; mixed.background = background.mixed; form.clearBackground = false
  const borders = (props.kind === 'table' ? grid.anchors : owners).flatMap(owner => DOCUMENT_TABLE_BORDER_SIDES.map(side => documentTableCellBorder(grid, owner, side)))
  const borderColor = documentTableCommonValue(borders.map(border => border.color)), borderWidth = documentTableCommonValue(borders.map(border => border.width)), borderStyle = documentTableCommonValue(borders.map(border => border.style))
  form.changeBorder = false; form.borderColor = borderColor.value; form.borderWidth = borderWidth.value === undefined ? undefined : borderWidth.value * 0.75; form.borderStyle = borderStyle.value; form.borderPreset = 'all'
  mixed.borderColor = borderColor.mixed; mixed.borderWidth = borderWidth.mixed; mixed.borderStyle = borderStyle.mixed
  form.moveKind = 'physical'; form.deltaX = 0; form.deltaY = 0; form.targetOffset = 0; form.targetIndex = 1; form.position = 'before'
  requestForm.value = JSON.parse(JSON.stringify(form)) as TablePropertiesForm
}, { immediate: true })
watch(() => form.column, column => {
  if (!props.open || props.kind !== 'type') return
  const config = source.value?.columnTypes?.[column]
  form.type = column === props.range.startColumn && props.type ? props.type : config?.type ?? 'text'; form.options = (config?.dropdownOptions ?? []).map(option => ({ ...option }))
})
watch(() => props.layout, layout => {
  if (!props.open || !layout || requestModel.value !== props.model || !source.value || source.value.style?.layout !== 'auto') return
  const grid = documentTableGrid(source.value), widths = layout.tableFragments?.find(fragment => fragment.tableId === props.tableId && !fragment.repeatedHeader)?.columnWidths
  if (!widths || widths.length !== grid.columnCount || widths.some(width => !Number.isFinite(width) || width <= 0)) return
  requestWidths.value = [...widths]; hasMeasuredWidths.value = true
  if (!changed('width') && requestForm.value) {
    const common = documentTableCommonValue(widths.slice(props.range.startColumn, props.range.endColumn + 1))
    const width = props.kind === 'column' ? common.value : widths.reduce((sum, width) => sum + width, 0)
    mixed.width = props.kind === 'column' && common.mixed
    form.width = width; requestForm.value = { ...requestForm.value, width }
  }
})

function selectedBorderPatch(): Partial<DocumentTableBorder> | undefined {
  const patch: Partial<DocumentTableBorder> = {}, choosing = changed('borderPreset') || props.kind === 'table' && form.changeBorder
  if ((changed('borderColor') || choosing && !mixed.borderColor) && form.borderColor !== undefined) patch.color = form.borderColor
  if ((changed('borderWidth') || choosing && !mixed.borderWidth) && typeof form.borderWidth === 'number') patch.width = form.borderWidth * 4 / 3
  if ((changed('borderStyle') || choosing && !mixed.borderStyle) && form.borderStyle !== undefined) patch.style = form.borderStyle
  if (form.borderPreset === 'none') patch.width = 0
  return Object.keys(patch).length ? patch : undefined
}
function selectedMarginPatch(): Partial<DocumentTableCellMargin> {
  const patch: Partial<DocumentTableCellMargin> = {}
  for (const side of DOCUMENT_TABLE_BORDER_SIDES) if ((marginChanged(side) || changed('inheritMargin') && form.inheritMargin === false) && typeof form.margin[side] === 'number') patch[side] = form.margin[side] as number
  return patch
}

const inputIssue = computed(() => {
  if (changed('width') && typeof form.width !== 'number') return '请输入要应用的宽度'
  if (changed('height') && typeof form.height !== 'number') return '请输入要应用的行高'
  if (DOCUMENT_TABLE_BORDER_SIDES.some(side => marginChanged(side) && typeof form.margin[side] !== 'number')) return '请输入要应用的单元格边距'
  if ((props.kind === 'background' || form.changeBackground) && changed('background') && !form.clearBackground && !form.background) return '请选择要应用的背景色'
  if ((props.kind === 'border' || form.changeBorder) && (changed('borderColor') && !form.borderColor || changed('borderWidth') && typeof form.borderWidth !== 'number' || changed('borderStyle') && !form.borderStyle)) return '请输入要应用的边框属性'
  if (source.value?.style?.layout === 'auto' && !hasMeasuredWidths.value
    && (props.kind === 'column' && changed('width') || props.kind === 'table' && (changed('width') || changed('layout') && form.layout === 'fixed'))) return '等待画布计算当前列宽后再应用尺寸'
  return undefined
})

const command = computed<DocumentTableCommand | undefined>(() => {
  const range = requestRange.value
  if (!range) return undefined
  const target = { tableId: props.tableId, range }
  switch (props.kind) {
    case 'table': {
      const patch: DocumentTableStylePatch = {}
      if (changed('width') && typeof form.width === 'number') patch.width = form.width
      for (const property of ['indent', 'offsetY', 'layout'] as const) if (changed(property)) Object.assign(patch, { [property]: form[property] })
      if (changed('align') && !modern.value) patch.align = form.align
      if (source.value?.style?.layout === 'auto' && (patch.width !== undefined || patch.layout === 'fixed') && hasMeasuredWidths.value) patch.columnWidths = [...requestWidths.value!]
      if (changed('inheritMargin') && form.inheritMargin === true) patch.cellMargin = null
      else if (changed('inheritMargin') || DOCUMENT_TABLE_BORDER_SIDES.some(marginChanged)) {
        const margin = { ...documentTableCellMargin(source.value!), ...selectedMarginPatch() }
        patch.cellMargin = margin
      }
      if (form.changeBackground && (changed('background') || form.clearBackground) && (form.clearBackground || form.background !== undefined)) patch.backgroundColor = form.clearBackground ? null : form.background
      const border = form.changeBorder && selectedBorderPatch()
      if (border) patch.border = { preset: 'all', patch: border }
      return Object.keys(patch).length ? { id: 'doc.command.table-set-style', ...target, patch } : undefined
    }
    case 'row': {
      const patch = { ...(changed('height') && typeof form.height === 'number' ? { height: form.height } : {}),
        ...(changed('heightRule') && form.heightRule !== undefined ? { heightRule: form.heightRule } : {}), ...(changed('allowSplit') && form.allowSplit !== undefined ? { allowSplit: form.allowSplit } : {}) }
      return Object.keys(patch).length ? { id: 'doc.command.table-resize-row', ...target, ...patch, preserveHeightRule: true } : undefined
    }
    case 'column': return changed('width') && typeof form.width === 'number' ? { id: 'doc.command.table-resize-column', ...target, width: form.width, widths: [...requestWidths.value!] } : undefined
    case 'cell': {
      if (changed('inheritMargin') && form.inheritMargin === true) return { id: 'doc.command.table-set-cell-margin', ...target, margin: null }
      const marginPatch = selectedMarginPatch()
      return Object.keys(marginPatch).length ? { id: 'doc.command.table-set-cell-margin', ...target, marginPatch } : undefined
    }
    case 'background': return (changed('background') || form.clearBackground) && (form.clearBackground || form.background !== undefined)
      ? { id: 'doc.command.table-set-background', ...target, color: form.clearBackground ? null : form.background! } : undefined
    case 'border': { const borderPatch = selectedBorderPatch(); return borderPatch ? { id: 'doc.command.table-set-border', ...target, preset: form.borderPreset, borderPatch } : undefined }
    case 'header': return { id: 'doc.command.table-set-header-row-count', ...target, count: form.headerCount }
    case 'type': return { id: 'doc.command.table-set-column-type', ...target, column: form.column, config: { type: form.type,
      ...(form.type === 'dropdown' ? { dropdownOptions: form.options.map(option => ({ ...option })) } : {}) } }
    case 'move':
      if (form.moveKind === 'offset') return { id: 'doc.command.table-move', ...target, targetOffset: form.targetOffset }
      if (form.moveKind === 'row') return { id: 'doc.command.table-move-rows', ...target, fromRow: range.startRow, count: range.endRow - range.startRow + 1, toRow: form.targetIndex - 1, position: form.position }
      if (form.moveKind === 'column') return { id: 'doc.command.table-move-columns', ...target, fromColumn: range.startColumn, count: range.endColumn - range.startColumn + 1, toColumn: form.targetIndex - 1, position: form.position }
      return { id: 'doc.command.table-move', ...target, deltaX: form.deltaX, deltaY: form.deltaY }
  }
})
const state = computed(() => props.composing ? { available: false, reason: '请先完成当前输入法输入' }
  : requestModel.value !== props.model ? { available: false, reason: '表格内容已更改，请重新打开属性' }
  : inputIssue.value ? { available: false, reason: inputIssue.value }
  : command.value ? documentTableCommandState(props.model, props.selection, command.value, props.permission) : { available: false, reason: source.value ? '请更改需要应用的属性' : '表格已不存在' })
function submit() { if (state.value.available && command.value) emit('submit', command.value) }
function addOption() { form.options.push({ id: documentId('table-option'), label: '' }) }
</script>

<template>
  <Modal :open="open" :title="titles[kind]" :width="kind === 'table' || kind === 'type' ? 560 : 440" :destroy-on-hidden="true"
    ok-text="应用" cancel-text="取消" :ok-button-props="{ disabled: !state.available }" @ok="submit" @cancel="emit('update:open', false)" @after-close="emit('closed')">
    <Form :model="form" layout="vertical" :disabled="composing" @finish="submit">
      <template v-if="kind === 'table'">
        <div class="table-property-grid">
          <FormItem name="width" label="表格宽度（px）"><InputNumber v-model:value="form.width" :min="1" /></FormItem>
          <FormItem v-if="!modern" name="align" label="整表对齐"><Select v-model:value="form.align" :options="[{ value: 'left', label: '靠左' }, { value: 'center', label: '居中' }, { value: 'right', label: '靠右' }]" /></FormItem>
          <FormItem name="indent" label="水平偏移（px）"><InputNumber v-model:value="form.indent" /></FormItem>
          <FormItem name="offsetY" label="垂直偏移（px）"><InputNumber v-model:value="form.offsetY" /></FormItem>
          <FormItem name="layout" label="列宽布局"><Select v-model:value="form.layout" :options="[{ value: 'fixed', label: '固定列宽' }, { value: 'auto', label: '按内容布局' }]" /></FormItem>
        </div>
        <FormItem name="changeBackground" label="更改表格背景"><Switch v-model:checked="form.changeBackground" /></FormItem>
        <div v-if="form.changeBackground" class="table-property-grid">
          <FormItem name="background" label="背景色"><ColorPicker v-model:value="form.background" value-format="hex" :disabled-alpha="true" :disabled="form.clearBackground" /></FormItem>
          <FormItem name="clearBackground" label="清除背景"><Switch v-model:checked="form.clearBackground" /></FormItem>
        </div>
        <FormItem name="changeBorder" label="更改所有边框"><Switch v-model:checked="form.changeBorder" /></FormItem>
        <div v-if="form.changeBorder" class="table-property-grid">
          <FormItem name="borderColor" label="边框颜色"><ColorPicker v-model:value="form.borderColor" value-format="hex" :disabled-alpha="true" /><span v-if="mixed.borderColor && form.borderColor === undefined" class="table-property-note">混合值（保持）</span></FormItem>
          <FormItem name="borderWidth" label="边框宽度（pt）"><InputNumber v-model:value="form.borderWidth" :placeholder="mixed.borderWidth ? '混合值（保持）' : undefined" :min="0" :step="0.25" /></FormItem>
          <FormItem name="borderStyle" label="边框样式"><Select v-model:value="form.borderStyle" :placeholder="mixed.borderStyle ? '混合值（保持）' : undefined" :options="[{ value: 'solid', label: '实线' }, { value: 'dot', label: '点线' }, { value: 'dash', label: '虚线' }]" /></FormItem>
        </div>
      </template>
      <template v-if="kind === 'row'">
        <FormItem name="heightRule" label="行高规则"><Select v-model:value="form.heightRule" :placeholder="mixed.heightRule ? '混合值（保持）' : undefined" :options="[{ value: 'auto', label: '自动' }, { value: 'atLeast', label: '最小值' }, { value: 'exact', label: '固定值' }]" /></FormItem>
        <FormItem name="height" label="行高（px）"><InputNumber v-model:value="form.height" :placeholder="mixed.height ? '混合值（保持）' : form.height === undefined ? '自动' : undefined" :min="1" :disabled="form.heightRule === 'auto'" /></FormItem>
        <FormItem name="allowSplit" label="允许跨页拆分同一行"><Checkbox v-model:checked="form.allowSplit" :indeterminate="form.allowSplit === undefined">允许</Checkbox></FormItem>
      </template>
      <FormItem v-if="kind === 'column'" name="width" label="所选列宽（px）"><InputNumber v-model:value="form.width" :placeholder="mixed.width ? '混合值（保持）' : undefined" :min="1" /></FormItem>
      <template v-if="kind === 'background'">
        <FormItem name="background" label="背景色"><ColorPicker v-model:value="form.background" value-format="hex" :disabled-alpha="true" :disabled="form.clearBackground" /><span v-if="mixed.background && form.background === undefined && !form.clearBackground" class="table-property-note">混合值（保持）</span></FormItem>
        <FormItem name="clearBackground" label="清除单元格背景"><Switch v-model:checked="form.clearBackground" /></FormItem>
      </template>
      <template v-if="kind === 'border'">
        <FormItem name="borderPreset" label="应用边框"><Select v-model:value="form.borderPreset" :options="Object.entries(DOCUMENT_TABLE_BORDER_LABELS).map(([value, label]) => ({ value, label }))" /></FormItem>
        <FormItem name="borderColor" label="边框颜色"><ColorPicker v-model:value="form.borderColor" value-format="hex" :disabled-alpha="true" /><span v-if="mixed.borderColor && form.borderColor === undefined" class="table-property-note">混合值（保持）</span></FormItem>
        <FormItem name="borderWidth" label="边框宽度（pt）"><InputNumber v-model:value="form.borderWidth" :placeholder="mixed.borderWidth ? '混合值（保持）' : undefined" :min="0" :step="0.25" /></FormItem>
        <FormItem name="borderStyle" label="边框样式"><Select v-model:value="form.borderStyle" :placeholder="mixed.borderStyle ? '混合值（保持）' : undefined" :options="[{ value: 'solid', label: '实线' }, { value: 'dot', label: '点线' }, { value: 'dash', label: '虚线' }]" /></FormItem>
      </template>
      <template v-if="kind === 'table' || kind === 'cell'">
        <FormItem name="inheritMargin" :label="kind === 'cell' ? '继承表格默认内边距' : '使用默认内边距'"><Checkbox v-if="kind === 'cell'" v-model:checked="form.inheritMargin" :indeterminate="form.inheritMargin === undefined">继承</Checkbox><Switch v-else v-model:checked="form.inheritMargin" /></FormItem>
        <div class="table-property-grid">
          <FormItem v-for="side in (['top', 'right', 'bottom', 'left'] as const)" :key="side" :name="['margin', side]" :label="({ top: '上', right: '右', bottom: '下', left: '左' })[side] + '内边距（px）'">
            <InputNumber v-model:value="form.margin[side]" :placeholder="mixed.margin[side] ? '混合值（保持）' : undefined" :min="0" :disabled="form.inheritMargin === true" />
          </FormItem>
        </div>
      </template>
      <template v-if="kind === 'header'">
        <FormItem name="headerCount" label="在每页重复的表头行数"><InputNumber v-model:value="form.headerCount" :min="0" :max="source?.rows.length ?? 0" :precision="0" /></FormItem>
        <p class="table-property-note">表头包含表格最上方的行；合并单元格必须完整包含在表头内。</p>
      </template>
      <template v-if="kind === 'type'">
        <div class="table-property-grid">
          <FormItem name="column" label="列"><Select v-model:value="form.column" :options="columnOptions" /></FormItem>
          <FormItem name="type" label="值类型"><Select v-model:value="form.type" :options="[{ value: 'text', label: '文本' }, { value: 'number', label: '数字' }, { value: 'date', label: '日期' }, { value: 'checkbox', label: '复选框' }, { value: 'dropdown', label: '下拉列表' }]" /></FormItem>
        </div>
        <template v-if="form.type === 'dropdown'">
          <div v-for="(option, index) in form.options" :key="option.id" class="table-option-row">
            <FormItem :name="['options', index, 'label']" :label="`选项 ${index + 1}`"><Input v-model:value="option.label" /></FormItem>
            <FormItem :name="['options', index, 'color']" label="颜色"><ColorPicker v-model:value="option.color" value-format="hex" :disabled-alpha="true" allow-clear /></FormItem>
            <Button type="text" :aria-label="`删除选项 ${index + 1}`" @click="form.options.splice(index, 1)"><Trash2 :size="16" /></Button>
          </div>
          <Button @click="addOption"><Plus :size="16" />添加选项</Button>
        </template>
      </template>
      <template v-if="kind === 'move'">
        <FormItem name="moveKind" label="移动方式"><Select v-model:value="form.moveKind" :options="[{ value: 'physical', label: '调整表格位置' }, { value: 'offset', label: '移至正文位置' }, { value: 'row', label: '移动所选行' }, { value: 'column', label: '移动所选列' }]" /></FormItem>
        <div v-if="form.moveKind === 'physical'" class="table-property-grid">
          <FormItem name="deltaX" label="水平移动（px）"><InputNumber v-model:value="form.deltaX" /></FormItem>
          <FormItem name="deltaY" label="垂直移动（px）"><InputNumber v-model:value="form.deltaY" /></FormItem>
        </div>
        <FormItem v-else-if="form.moveKind === 'offset'" name="targetOffset" label="正文字符位置"><InputNumber v-model:value="form.targetOffset" :min="0" :precision="0" /></FormItem>
        <template v-else>
          <FormItem name="targetIndex" :label="form.moveKind === 'row' ? '目标行' : '目标列'"><InputNumber v-model:value="form.targetIndex" :min="1" :max="form.moveKind === 'row' ? source?.rows.length : columnOptions.length" :precision="0" /></FormItem>
          <FormItem name="position" label="相对位置"><Select v-model:value="form.position" :options="[{ value: 'before', label: form.moveKind === 'row' ? '之前（上方）' : '之前（左侧）' }, { value: 'after', label: form.moveKind === 'row' ? '之后（下方）' : '之后（右侧）' }]" /></FormItem>
        </template>
      </template>
      <p v-if="!state.available" role="status" class="table-property-note">{{ state.reason }}</p>
    </Form>
  </Modal>
</template>

<style scoped>
.table-property-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 0 16px; }
.table-property-grid :deep(.ant-input-number) { width: 100%; }
.table-option-row { display: grid; grid-template-columns: minmax(0, 1fr) auto auto; gap: 12px; align-items: center; }
.table-property-note { color: var(--ant-color-text-secondary); font-size: 12px; line-height: 1.5; }
@media (max-width: 480px) { .table-property-grid { grid-template-columns: minmax(0, 1fr); } }
</style>
