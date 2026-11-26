<script setup lang="ts">
import { Button, Space, Tooltip } from 'antdv-next'
import type { PivotAxisState } from '../services/pivot-axis-state.ts'

defineProps<{ state: PivotAxisState; disabled: boolean; busy?: boolean; step(direction: 'expand' | 'collapse'): Promise<boolean> }>()
</script>

<template>
  <Space :size="2" class="pivot-chart-hierarchy-controls" aria-label="透视图类别层级">
    <Tooltip title="展开一级">
      <Button type="default" size="small" aria-label="展开透视图类别一级" :disabled="disabled || !state.canExpand" :loading="busy"
        @pointerdown.stop @click.stop="void step('expand')">+</Button>
    </Tooltip>
    <Tooltip title="折叠一级">
      <Button type="default" size="small" aria-label="折叠透视图类别一级" :disabled="disabled || !state.canCollapse" :loading="busy"
        @pointerdown.stop @click.stop="void step('collapse')">−</Button>
    </Tooltip>
  </Space>
</template>

<style scoped>
.pivot-chart-hierarchy-controls { position: absolute; right: 8px; bottom: 8px; pointer-events: auto; }
</style>
