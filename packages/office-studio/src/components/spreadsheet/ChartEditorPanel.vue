<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Alert, Button, Drawer, Form } from 'antdv-next'
import type { ChartModel } from '../../domain/model'
import type { SpreadsheetRenderSheet } from '../../rendering/spreadsheet/layout'
import type { ChartData } from '../../features/charts/chart-engine'
import { createChartSettingsDraft, chartSettingsPatch, type ChartSettingsDraft, type ChartSettingsSourceSheet } from '../../features/charts/chart-settings-draft'
import ChartSettingsFields from '../../features/charts/views/ChartSettingsFields.vue'

type ChartPatch = Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>
const props = defineProps<{ open: boolean; chart: ChartModel | null; sheet: SpreadsheetRenderSheet | undefined; sheets?: ChartSettingsSourceSheet[]; data?: ChartData; busy?: boolean; error?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [patch: ChartPatch]; closed: [] }>()
const draft = ref<ChartSettingsDraft>()
const localError = ref('')
const sourceSheets = computed(() => props.sheets ?? (props.sheet ? [props.sheet] : []))

watch(() => [props.open, props.chart?.id, props.chart?.revision] as const, ([open]) => {
  if (!open || !props.chart) return
  draft.value = createChartSettingsDraft(props.chart)
  localError.value = ''
}, { immediate: true })

function apply(): void {
  if (!props.chart || !draft.value || props.busy) return
  try {
    const patch = chartSettingsPatch(draft.value)
    localError.value = ''
    emit('apply', patch)
  } catch (error) { localError.value = error instanceof Error ? error.message : String(error) }
}
function close(): void { if (!props.busy) emit('update:open', false) }
function opened(value: boolean): void { if (!value) emit('closed') }
</script>

<template>
  <Drawer :open="open" title="图表设置" placement="right" :size="560" :footer="null" :mask-closable="!busy" @close="close" @after-open-change="opened">
    <Form v-if="draft" layout="vertical" class="chart-editor-form" @submit.prevent="apply">
      <ChartSettingsFields v-model:draft="draft" :sheets="sourceSheets" :host-sheet-id="chart?.sheetId" :data="data" :disabled="busy" />
      <Alert v-if="localError || error" type="error" :message="localError || error" show-icon />
      <div class="chart-editor-actions">
        <Button :disabled="busy" @click="close">关闭</Button>
        <Button type="primary" :loading="busy" @click="apply">应用</Button>
      </div>
    </Form>
  </Drawer>
</template>

<style scoped>
.chart-editor-form { display: grid; gap: 14px; min-width: 0; }
.chart-editor-actions { display: flex; justify-content: flex-end; gap: 8px; padding-top: 12px; }
</style>
