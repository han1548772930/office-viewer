<script setup lang="ts">
import { ref, watch } from 'vue'
import { Alert, Button, Form, Modal } from 'antdv-next'
import type { ChartConfig } from '../chart-engine'
import { chartConfigFromDraft, createChartSettingsDraft, type ChartSettingsDraft, type ChartSettingsSourceSheet } from '../chart-settings-draft'
import type { ChartSettingsSourceLoader, ChartSettingsSourceRevision, ChartSettingsStatisticsLoader } from '../chart-settings-data'
import { useChartSettingsSource } from './use-chart-settings-source'
import ChartSettingsFields from './ChartSettingsFields.vue'

const props = defineProps<{ open: boolean; config: ChartConfig | null; sheets?: ChartSettingsSourceSheet[]; hostSheetId?: string; date1904?: boolean; sourceRevision?: ChartSettingsSourceRevision; loadSource?: ChartSettingsSourceLoader; loadStatistics?: ChartSettingsStatisticsLoader; busy?: boolean; error?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [config: ChartConfig]; closed: [] }>()
const draft = ref<ChartSettingsDraft>(), localError = ref('')
watch(() => [props.open, props.config] as const, ([open, config]) => { if (open && config) { draft.value = createChartSettingsDraft(config); localError.value = '' } }, { immediate: true })
const { projection, loading: sourceLoading, assertReady, reload } = useChartSettingsSource({
  active: () => props.open, draft: () => draft.value, sheets: () => props.sheets, hostSheetId: () => props.hostSheetId,
  date1904: () => props.date1904, sourceRevision: () => props.sourceRevision, load: () => props.loadSource, loadStatistics: () => props.loadStatistics,
})
function apply(): void {
  if (!draft.value || props.busy) return
  try {
    const config = chartConfigFromDraft(draft.value)
    assertReady()
    localError.value = ''; emit('apply', config)
  }
  catch (error) { localError.value = error instanceof Error ? error.message : String(error) }
}
function close(): void { if (!props.busy) emit('update:open', false) }
function opened(value: boolean): void { if (!value && !props.open) emit('closed') }
</script>

<template>
  <Modal :open="open" title="插入图表" ok-text="插入" cancel-text="取消" :width="760" :confirm-loading="busy" :ok-button-props="{ disabled: sourceLoading }" :mask-closable="!busy" :keyboard="!busy" :closable="!busy" :focus-trigger-after-close="false" @ok="apply" @cancel="close" @after-open-change="opened">
    <Form v-if="draft" layout="vertical" class="chart-insert-form" @submit.prevent="apply">
      <ChartSettingsFields v-model:draft="draft" :sheets="sheets" :host-sheet-id="hostSheetId" :projection="projection" :date1904="date1904 ?? projection.data?.date1904" :disabled="busy" />
      <Alert v-if="sourceLoading" type="info" message="正在读取图表数据源…" show-icon role="status" />
      <Button v-if="projection.error && (loadSource || loadStatistics)" :disabled="busy || sourceLoading" @click="reload">重新读取数据源</Button>
      <Alert v-if="localError || error" type="error" :message="localError || error" show-icon />
    </Form>
  </Modal>
</template>

<style scoped>
.chart-insert-form { display: grid; gap: 14px; max-height: 70vh; min-width: 0; overflow-y: auto; padding-right: 8px; }
</style>
