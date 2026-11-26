<script setup lang="ts">
import { Button, Dropdown, Tooltip, theme } from 'antdv-next'
import { ChevronDown, GripVertical } from 'lucide-vue-next'
import { computed } from 'vue'
import type { PivotValuePosition } from '../../../domain/model.ts'
import { pivotValuesUiIndexOptions, type PivotValuesUiState } from '../services/pivot-values-position-ui.ts'

const props = withDefaults(defineProps<{ state: PivotValuesUiState; subjectLabel: '透视表' | '透视图'; disabled?: boolean; dense?: boolean; testId?: string; menuOpen?: boolean }>(), { menuOpen: undefined })
const emit = defineEmits<{ change: [position: PivotValuePosition, index: number]; 'drag-start': [event: DragEvent]; 'drag-end': []; 'menu-open': [open: boolean] }>()
const { token } = theme.useToken()
const appearance = computed(() => ({ color: token.value.colorText, background: token.value.colorBgContainer,
  borderColor: token.value.colorBorderSecondary, '--values-grip': token.value.colorTextTertiary }))
const menu = computed(() => ({
  selectedKeys: [`position:${props.state.position}`, `index:${props.state.index}`],
  items: [
    { key: 'position:row', label: '放到行区域', disabled: props.disabled },
    { key: 'position:column', label: '放到列区域', disabled: props.disabled },
    { key: 'position:none', label: '无 ΣValues 位置', disabled: props.disabled },
    { type: 'divider' as const },
    { key: 'up', label: '上移 ΣValues', disabled: props.disabled || props.state.effectiveIndex === 0 },
    { key: 'down', label: '下移 ΣValues', disabled: props.disabled || props.state.effectiveIndex >= props.state.dimensionCount },
    { key: 'automatic', label: 'ΣValues 自动末尾', disabled: props.disabled || props.state.index === -1 },
    { key: 'indices', label: 'ΣValues 插入位置', disabled: props.disabled,
      children: pivotValuesUiIndexOptions(props.state).map(option => ({ key: `index:${option.value}`, label: option.label, disabled: props.disabled })) },
  ],
  onClick: ({ key }: { key: string | number }) => {
    if (props.disabled) return
    if (key === 'position:row' || key === 'position:column' || key === 'position:none') emit('change', String(key).slice(9) as PivotValuePosition, props.state.index)
    else if (key === 'up' && props.state.effectiveIndex > 0) emit('change', props.state.position, props.state.effectiveIndex - 1)
    else if (key === 'down' && props.state.effectiveIndex < props.state.dimensionCount) emit('change', props.state.position, props.state.effectiveIndex + 1)
    else if (key === 'automatic') emit('change', props.state.position, -1)
    else if (String(key).startsWith('index:')) {
      const index = Number(String(key).slice(6))
      if (pivotValuesUiIndexOptions(props.state).some(option => option.value === index)) emit('change', props.state.position, index)
    }
  },
}))
function startDrag(event: DragEvent): void { if (props.disabled) event.preventDefault(); else emit('drag-start', event) }
</script>

<template>
  <div v-if="state.visible" class="pivot-values-field" :class="{ 'is-dense': dense }" :style="appearance"
    :data-testid="testId ?? (subjectLabel === '透视表' ? 'pivot-table-values-field' : 'pivot-chart-values-field')"
    data-pivot-synthetic="values" :data-values-position="state.position" :data-values-index="state.index" :data-values-effective-index="state.effectiveIndex"
    :draggable="!disabled" @dragstart="startDrag" @dragend="emit('drag-end')">
    <GripVertical v-if="!dense" class="values-grip" :size="16" aria-hidden="true" />
    <Dropdown :open="menuOpen" :trigger="['click']" :menu="menu" :disabled="disabled" placement="bottomLeft" @open-change="emit('menu-open', $event)">
      <Tooltip :title="`${state.measureCount} 个汇总字段的排列位置`"><Button :type="dense ? 'default' : 'text'" size="small" :disabled="disabled"
        :aria-label="`${subjectLabel} ΣValues 位置菜单`"><span>ΣValues</span><ChevronDown :size="14" aria-hidden="true" /></Button></Tooltip>
    </Dropdown>
  </div>
</template>

<style scoped>
.pivot-values-field { display: flex; align-items: center; gap: 5px; min-height: 28px; padding: 2px 4px; border: 1px solid; border-radius: 4px; }
.pivot-values-field :deep(.ant-btn) { display: inline-flex; align-items: center; gap: 4px; }
.values-grip { color: var(--values-grip); }
.is-dense { min-height: 0; padding: 0; border: 0; background: transparent !important; }
</style>
