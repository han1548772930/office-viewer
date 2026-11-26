<script setup lang="ts">
import { computed, inject } from 'vue'
import { Button, Checkbox, ColorPicker, Divider, FormItem, InputNumber, Select, Tooltip } from 'antdv-next'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { SPARKLINE_POINT_ORDER, type ResolvedSparklineSettings, type SparklineExtremumMode, type SparklinePointName,
  type SparklineSettingsInput, type SparklineTheme } from './sparkline-model.ts'

const props = defineProps<{ settings: ResolvedSparklineSettings; tab: 'basic' | 'advanced'; disabled: boolean; textDisabled: boolean;
  existing: boolean; canManage: boolean; canCombine: boolean; canUncombine: boolean }>()
const emit = defineEmits<{ change: [patch: SparklineSettingsInput]; clear: [scope: 'item' | 'group']; combine: []; uncombine: [] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const types = computed(() => [{ label: t('sheet.sparkline.type.line', '折线图'), value: 'line' }, { label: t('sheet.sparkline.type.column', '柱形图'), value: 'column' },
  { label: t('sheet.sparkline.type.winloss', '胜负图'), value: 'winloss' }, { label: t('sheet.sparkline.type.pie', '饼图'), value: 'pie' }])
const themes = computed<Array<{ label: string; value: SparklineTheme }>>(() => [
  { label: t('sheet.sparkline.theme.standard', '标准'), value: 'standard' }, { label: t('sheet.sparkline.theme.specialty', '专用'), value: 'specialty' },
  { label: t('sheet.sparkline.theme.distinctive', '鲜明'), value: 'distinctive' }, { label: t('sheet.sparkline.theme.soft', '柔和'), value: 'soft' },
  { label: t('sheet.sparkline.theme.fresh', '清新'), value: 'fresh' }, { label: t('sheet.sparkline.theme.custom', '自定义'), value: 'custom' },
])
const emptyModes = computed(() => [{ label: t('sheet.sparkline.empty-gap', '留空隙'), value: 'gap' }, { label: t('sheet.sparkline.empty-zero', '按零值绘制'), value: 'zero' },
  { label: t('sheet.sparkline.empty-connect', '连接相邻数据点'), value: 'connect' }])
const pointLabels = computed<Record<SparklinePointName, string>>(() => ({ highPoint: t('sheet.sparkline.high-point', '高值点'), firstPoint: t('sheet.sparkline.first-point', '首值点'),
  lowPoint: t('sheet.sparkline.low-point', '低值点'), lastPoint: t('sheet.sparkline.last-point', '末值点'), negativePoint: t('sheet.sparkline.negative-point', '负值点'), markersPoint: t('sheet.sparkline.markers', '所有标记') }))
const axes = computed<Array<{ label: string; value: SparklineExtremumMode }>>(() => [{ label: t('sheet.sparkline.axis-individual', '每个迷你图单独计算'), value: 'individual' },
  { label: t('sheet.sparkline.axis-group', '整组统一计算'), value: 'group' }, { label: t('sheet.sparkline.axis-custom', '自定义'), value: 'custom' }])
const extremumFields = computed(() => [{ key: 'extremumMin' as const, label: t('sheet.sparkline.minimum', '下限') }, { key: 'extremumMax' as const, label: t('sheet.sparkline.maximum', '上限') }])
const lineWidths = computed(() => [...new Set([1, 2, 3, 4, props.settings.lineWidth])].sort((a, b) => a - b).map(value => ({ label: `${value} px`, value })))
function pointVisible(name: SparklinePointName, value: boolean): void { emit('change', { points: { [name]: { visible: value } } }) }
function pointColor(name: SparklinePointName, value: unknown): void { if (typeof value === 'string') emit('change', { points: { [name]: { color: value } }, themeType: 'custom' }) }
function changeType(value: unknown): void { if (value === 'line' || value === 'column' || value === 'winloss' || value === 'pie') emit('change', { type: value }) }
function changeTheme(value: unknown): void { if (themes.value.some(theme => theme.value === value)) emit('change', { themeType: value as SparklineTheme }) }
function changeColor(value: unknown): void { if (typeof value === 'string') emit('change', { color: value, themeType: 'custom' }) }
function changeLineWidth(value: unknown): void { if (typeof value === 'number') emit('change', { lineWidth: value }) }
function changeAxisMode(key: 'extremumMin' | 'extremumMax', value: unknown): void {
  if (value !== 'individual' && value !== 'group' && value !== 'custom') return
  const current = props.settings[key]
  emit('change', { [key]: { type: value, ...(value === 'custom' ? { value: current.value ?? (key === 'extremumMin' ? 0 : 100) } : {}) } })
}
function changeAxisNumber(key: 'extremumMin' | 'extremumMax', value: unknown): void { emit('change', { [key]: { type: 'custom', value: typeof value === 'number' ? value : undefined } }) }
function changeEmpty(key: 'emptyCells' | 'nonNumericCells', value: unknown): void { if (value === 'gap' || value === 'zero' || value === 'connect') emit('change', { [key]: value }) }
function changeAxisColor(value: unknown): void { if (typeof value === 'string') emit('change', { axisOptions: { color: value } }) }
</script>

<template>
  <template v-if="tab === 'basic'">
    <FormItem :name="['settings', 'type']" :label="t('sheet.sparkline.type', '类型')"><Select :value="settings.type" :options="types" :disabled="disabled" aria-label="迷你图类型" @update:value="changeType" /></FormItem>
    <template v-if="settings.type !== 'pie'">
      <FormItem :name="['settings', 'themeType']" :label="t('sheet.sparkline.theme', '主题')"><Select :value="settings.themeType" :options="themes" :disabled="disabled" aria-label="迷你图主题" @update:value="changeTheme" /></FormItem>
      <Divider>{{ t('sheet.sparkline.adjust-style', '调整样式') }}</Divider>
      <div class="sparkline-style-grid">
        <FormItem :name="['settings', 'color']" :label="t('sheet.sparkline.series-color', '系列颜色')"><ColorPicker :value="settings.color" value-format="hex" show-text :disabled="disabled" aria-label="迷你图系列颜色" @update:value="changeColor" /></FormItem>
        <FormItem v-if="settings.type === 'line'" :name="['settings', 'lineWidth']" :label="t('sheet.sparkline.line-width', '线宽')"><Select :value="settings.lineWidth" :options="lineWidths" :disabled="disabled" aria-label="迷你图线宽" @update:value="changeLineWidth" /></FormItem>
        <FormItem v-if="settings.type === 'line'" :name="['settings', 'showGradient']"><Checkbox :checked="settings.showGradient" :disabled="disabled" aria-label="迷你图渐变" @update:checked="value => emit('change', { showGradient: value })">{{ t('sheet.sparkline.gradient', '渐变填充') }}</Checkbox></FormItem>
        <FormItem v-else :name="['settings', 'showRadius']"><Checkbox :checked="settings.showRadius" :disabled="disabled" aria-label="迷你图圆角" @update:checked="value => emit('change', { showRadius: value })">{{ t('sheet.sparkline.radius', '圆角') }}</Checkbox></FormItem>
      </div>
      <Divider>{{ t('sheet.sparkline.points', '标记') }}</Divider>
      <div class="sparkline-points-grid">
        <div v-for="name in SPARKLINE_POINT_ORDER" :key="name" class="sparkline-point-field">
          <FormItem :name="['settings', 'points', name, 'visible']"><Tooltip :title="name === 'markersPoint' && settings.type !== 'line' ? '所有标记仅用于折线迷你图' : undefined"><Checkbox :checked="settings.points[name].visible" :disabled="disabled || name === 'markersPoint' && settings.type !== 'line'" :aria-label="`迷你图${pointLabels[name]}`" @update:checked="value => pointVisible(name, value)">{{ pointLabels[name] }}</Checkbox></Tooltip></FormItem>
          <FormItem v-if="settings.points[name].visible" :name="['settings', 'points', name, 'color']"><ColorPicker :value="settings.points[name].color" value-format="hex" :disabled="disabled || name === 'markersPoint' && settings.type !== 'line'" :aria-label="`迷你图${pointLabels[name]}颜色`" @update:value="value => pointColor(name, value)" /></FormItem>
        </div>
      </div>
    </template>
    <template v-if="existing">
      <Divider>{{ t('sheet.sparkline.manager', '管理') }}</Divider>
      <div class="sparkline-manage-actions"><Button danger :disabled="disabled || !canManage" @click="emit('clear', 'item')">{{ t('sheet.sparkline.remove-item', '删除当前迷你图') }}</Button><Button danger :disabled="disabled || !canManage" @click="emit('clear', 'group')">{{ t('sheet.sparkline.remove-group', '删除整组迷你图') }}</Button></div>
    </template>
  </template>
  <template v-else-if="settings.type !== 'pie'">
    <Divider>{{ t('sheet.sparkline.compose-manager', '分组管理') }}</Divider>
    <div class="sparkline-manage-actions"><Button :disabled="disabled || !existing || !canCombine" @click="emit('combine')">{{ t('sheet.sparkline.combine', '合并组') }}</Button><Button :disabled="disabled || !existing || !canUncombine" @click="emit('uncombine')">{{ t('sheet.sparkline.uncombine', '拆分组') }}</Button></div>
    <Divider>{{ t('sheet.sparkline.horizontal-axis', '水平轴') }}</Divider>
    <FormItem :name="['settings', 'axisOptions', 'visible']"><Checkbox :checked="settings.axisOptions.visible" :disabled="disabled" aria-label="迷你图显示零轴" @update:checked="value => emit('change', { axisOptions: { visible: value } })">{{ t('sheet.sparkline.show-zero-axis', '显示零轴') }}</Checkbox></FormItem>
    <FormItem v-if="settings.axisOptions.visible" :name="['settings', 'axisOptions', 'color']" :label="t('sheet.sparkline.axis-color', '轴颜色')"><ColorPicker :value="settings.axisOptions.color" value-format="hex" show-text :disabled="disabled" aria-label="迷你图轴颜色" @update:value="changeAxisColor" /></FormItem>
    <FormItem :name="['settings', 'axisOptions', 'reverse']"><Checkbox :checked="settings.axisOptions.reverse" :disabled="disabled" aria-label="迷你图从右到左" @update:checked="value => emit('change', { axisOptions: { reverse: value } })">{{ t('sheet.sparkline.reverse', '从右到左绘制') }}</Checkbox></FormItem>
    <Divider>{{ t('sheet.sparkline.vertical-axis', '垂直轴') }}</Divider>
    <div class="sparkline-style-grid">
      <div v-for="field in extremumFields" :key="field.key">
        <FormItem :name="['settings', field.key, 'type']" :label="field.label"><Select :value="settings[field.key].type" :options="axes" :disabled="disabled" :aria-label="`迷你图${field.label}方式`" @update:value="value => changeAxisMode(field.key, value)" /></FormItem>
        <FormItem v-if="settings[field.key].type === 'custom'" :name="['settings', field.key, 'value']" :label="`自定义${field.label}`"><InputNumber :value="settings[field.key].value ?? null" :disabled="textDisabled" :aria-label="`迷你图自定义${field.label}`" @update:value="value => changeAxisNumber(field.key, value)" /></FormItem>
      </div>
    </div>
    <Divider>{{ t('sheet.sparkline.more', '更多设置') }}</Divider>
    <FormItem :name="['settings', 'containHiddenCells']"><Checkbox :checked="settings.containHiddenCells" :disabled="disabled" aria-label="迷你图包含隐藏单元格" @update:checked="value => emit('change', { containHiddenCells: value })">{{ t('sheet.sparkline.include-hidden', '包含隐藏单元格') }}</Checkbox></FormItem>
    <FormItem :name="['settings', 'emptyCells']" :label="t('sheet.sparkline.empty-cells', '空单元格')"><Select :value="settings.emptyCells" :options="emptyModes" :disabled="disabled" aria-label="迷你图空单元格处理" @update:value="value => changeEmpty('emptyCells', value)" /></FormItem>
    <FormItem :name="['settings', 'nonNumericCells']" :label="t('sheet.sparkline.non-numeric-cells', '非数字单元格')"><Select :value="settings.nonNumericCells" :options="emptyModes" :disabled="disabled" aria-label="迷你图非数字单元格处理" @update:value="value => changeEmpty('nonNumericCells', value)" /></FormItem>
  </template>
</template>

<style scoped>
.sparkline-style-grid, .sparkline-points-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px 16px; }
.sparkline-point-field { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-width: 0; }
.sparkline-point-field :deep(.ant-form-item) { margin-bottom: 0; }
.sparkline-style-grid :deep(.ant-input-number), .sparkline-style-grid :deep(.ant-color-picker-trigger) { width: 100%; }
.sparkline-manage-actions { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 12px; }
@media (max-width: 480px) { .sparkline-style-grid, .sparkline-points-grid { grid-template-columns: minmax(0, 1fr); } }
</style>
