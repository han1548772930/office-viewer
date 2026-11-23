<script setup lang="ts">
import { ref, watch } from 'vue'
import { Alert, Form, Modal } from 'antdv-next'
import type { ChartConfig } from '../chart-engine'
import { chartConfigFromDraft, createChartSettingsDraft, type ChartSettingsDraft, type ChartSettingsSourceSheet } from '../chart-settings-draft'
import ChartSettingsFields from './ChartSettingsFields.vue'

const props = defineProps<{ open: boolean; config: ChartConfig | null; sheets?: ChartSettingsSourceSheet[]; hostSheetId?: string; busy?: boolean; error?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [config: ChartConfig]; closed: [] }>()
const draft = ref<ChartSettingsDraft>(), localError = ref('')
watch(() => props.open, open => { if (open && props.config) { draft.value = createChartSettingsDraft(props.config); localError.value = '' } }, { immediate: true })
function apply(): void {
  if (!draft.value || props.busy) return
  try { const config = chartConfigFromDraft(draft.value); localError.value = ''; emit('apply', config) }
  catch (error) { localError.value = error instanceof Error ? error.message : String(error) }
}
function close(): void { if (!props.busy) emit('update:open', false) }
function opened(value: boolean): void { if (!value) emit('closed') }
</script>

<template>
  <Modal :open="open" title="插入图表" ok-text="插入" cancel-text="取消" :width="760" :confirm-loading="busy" :mask-closable="!busy" @ok="apply" @cancel="close" @after-open-change="opened">
    <Form v-if="draft" layout="vertical" class="chart-insert-form" @submit.prevent="apply">
      <ChartSettingsFields v-model:draft="draft" :sheets="sheets" :host-sheet-id="hostSheetId" :disabled="busy" />
      <Alert v-if="localError || error" type="error" :message="localError || error" show-icon />
    </Form>
  </Modal>
</template>

<style scoped>
.chart-insert-form { display: grid; gap: 14px; max-height: 70vh; min-width: 0; overflow-y: auto; padding-right: 8px; }
</style>
