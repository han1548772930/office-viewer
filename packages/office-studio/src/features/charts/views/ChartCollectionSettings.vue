<script setup lang="ts">
import { Button, ColorPicker, FormItem, InputNumber, Select } from 'antdv-next'
import type { ChartConfig } from '../chart-engine'
import { cloneChartDraftConfig, type ChartSettingsField } from '../chart-settings-draft'
import { chartPalette } from '../chart-options'
import { optionGroup } from '../chart-settings-controls'

const props = defineProps<{ config: ChartConfig; seriesFields: ChartSettingsField[]; categories: string[]; disabled?: boolean }>()
const emit = defineEmits<{ 'update:config': [value: ChartConfig] }>()
function mutate(action: (config: ChartConfig) => void): void { const config = cloneChartDraftConfig(props.config); action(config); emit('update:config', config) }
function setPalette(index: number, value: unknown): void { if (typeof value === 'string') mutate(config => { (config.palette ??= [])[index] = value }) }
function addColor(): void { mutate(config => { (config.palette ??= []).push('#2563eb') }) }
function customizePalette(): void { mutate(config => { config.palette = chartPalette(config) }) }
function resetPalette(): void { mutate(config => { config.palette = [] }) }
function removeColor(index: number): void { mutate(config => { config.palette?.splice(index, 1) }) }
function addRange(): void { mutate(config => { const gauge = optionGroup(config, 'gauge'); (gauge.ranges ??= []).push({ from: gauge.min ?? 0, to: gauge.max ?? 100, color: '#059669' }) }) }
function rangeNumber(index: number, key: 'from' | 'to', value: unknown): void { if (typeof value === 'number' && Number.isFinite(value)) mutate(config => { const range = config.options?.gauge?.ranges?.[index]; if (range) range[key] = value }) }
function rangeColor(index: number, value: unknown): void { if (typeof value === 'string') mutate(config => { const range = config.options?.gauge?.ranges?.[index]; if (range) range.color = value }) }
function removeRange(index: number): void { mutate(config => { config.options?.gauge?.ranges?.splice(index, 1) }) }
function addRole(): void { mutate(config => { const roles = optionGroup(config, 'waterfall').pointRoles ??= [], seriesIndex = props.seriesFields[0]?.value ?? 1; let dataPointIndex = 0; while (roles.some(role => role.seriesIndex === seriesIndex && role.dataPointIndex === dataPointIndex)) dataPointIndex++; roles.push({ seriesIndex, dataPointIndex, role: 'total' }) }) }
function roleNumber(index: number, key: 'seriesIndex' | 'dataPointIndex', value: unknown): void { if (typeof value === 'number' && Number.isSafeInteger(value)) mutate(config => { const role = config.options?.waterfall?.pointRoles?.[index]; if (role) role[key] = value }) }
function roleType(index: number, value: unknown): void { if (value === 'total' || value === 'change') mutate(config => { const role = config.options?.waterfall?.pointRoles?.[index]; if (role) role.role = value }) }
function removeRole(index: number): void { mutate(config => { config.options?.waterfall?.pointRoles?.splice(index, 1) }) }
</script>

<template>
  <div class="chart-collections">
    <FormItem label="自定义配色" extra="清空后由当前主题提供配色。">
      <div class="chart-collection-actions">
        <Button v-if="!config.palette?.length" :disabled="disabled" @click="customizePalette">使用默认配色编辑</Button>
        <Button :disabled="disabled" @click="addColor">添加颜色</Button>
        <Button :disabled="disabled" @click="resetPalette">恢复主题配色</Button>
      </div>
    </FormItem>
    <div v-for="(color, index) in config.palette ?? []" :key="`palette-${index}`" class="chart-collection-row">
      <FormItem :label="`配色 ${index + 1}`"><ColorPicker :value="color" value-format="hex" format="hex" show-text :disabled="disabled" :aria-label="`配色 ${index + 1}`" @update:value="setPalette(index, $event)" /></FormItem>
      <Button danger :disabled="disabled" @click="removeColor(index)">移除颜色</Button>
    </div>
    <template v-if="config.type === 'gauge'">
      <FormItem label="仪表分段颜色"><Button :disabled="disabled" @click="addRange">添加刻度区间</Button></FormItem>
      <div v-for="(range, index) in config.options?.gauge?.ranges ?? []" :key="`range-${index}`" class="chart-collection-row chart-range-row">
        <FormItem label="起点"><InputNumber :value="range.from" aria-label="仪表区间起点" :disabled="disabled" @update:value="rangeNumber(index, 'from', $event)" /></FormItem>
        <FormItem label="终点"><InputNumber :value="range.to" aria-label="仪表区间终点" :disabled="disabled" @update:value="rangeNumber(index, 'to', $event)" /></FormItem>
        <FormItem label="颜色"><ColorPicker :value="range.color" format="hex" value-format="hex" aria-label="仪表区间颜色" :disabled="disabled" @update:value="rangeColor(index, $event)" /></FormItem>
        <Button danger :disabled="disabled" @click="removeRange(index)">移除区间</Button>
      </div>
    </template>
    <template v-if="config.type === 'waterfall'">
      <FormItem label="累计与合计项目"><Button :disabled="disabled" @click="addRole">设置合计项目</Button></FormItem>
      <div v-for="(role, index) in config.options?.waterfall?.pointRoles ?? []" :key="`role-${index}`" class="chart-collection-row chart-role-row">
        <FormItem label="系列"><Select :value="role.seriesIndex" :options="seriesFields" aria-label="瀑布图项目系列" :disabled="disabled" @update:value="roleNumber(index, 'seriesIndex', $event)" /></FormItem>
        <FormItem label="项目"><Select :value="role.dataPointIndex" :options="categories.map((label, value) => ({ value, label: `${value + 1} · ${label}` }))" aria-label="瀑布图项目" :disabled="disabled" @update:value="roleNumber(index, 'dataPointIndex', $event)" /></FormItem>
        <FormItem label="含义"><Select :value="role.role" :options="[{ value: 'change', label: '增减值' }, { value: 'total', label: '合计值' }]" aria-label="瀑布图项目含义" :disabled="disabled" @update:value="roleType(index, $event)" /></FormItem>
        <Button danger :disabled="disabled" @click="removeRole(index)">移除设置</Button>
      </div>
    </template>
  </div>
</template>

<style scoped>
.chart-collections { display: grid; gap: 8px; }
.chart-collection-actions, .chart-collection-row { display: flex; align-items: flex-end; gap: 8px; flex-wrap: wrap; }
.chart-collection-row :deep(.ant-form-item) { flex: 1; min-width: 90px; margin-bottom: 0; }
.chart-range-row :deep(.ant-input-number), .chart-role-row :deep(.ant-select) { width: 100%; }
.chart-collection-row > .ant-btn { margin-bottom: 0; }
</style>
