<script setup lang="ts">
import { computed, ref, shallowRef, watch } from 'vue'
import { Alert, Button, Drawer, Form, Tabs } from 'antdv-next'
import type { ChartModel } from '../../domain/model'
import type { ChartData } from '../../features/charts/chart-engine'
import { createChartSettingsDraft, chartSettingsPatch, type ChartSettingsDraft, type ChartSettingsSourceSheet } from '../../features/charts/chart-settings-draft'
import type { ChartSettingsProvidedData, ChartSettingsSourceLoader, ChartSettingsSourceRevision, ChartSettingsStatisticsLoader } from '../../features/charts/chart-settings-data'
import { useChartSettingsSource } from '../../features/charts/views/use-chart-settings-source'
import ChartSettingsFields from '../../features/charts/views/ChartSettingsFields.vue'
import { normalizeChartEditPanelRoute, planChartEditPanelRoute, type ChartEditPanelNavigation, type ChartEditPanelRoute, type ChartEditPanelTab } from '../../features/charts/chart-edit-panel-route'

type ChartPatch = Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>
const props = defineProps<{ open: boolean; chart: ChartModel | null; sheet?: ChartSettingsSourceSheet; sheets?: ChartSettingsSourceSheet[]; data?: ChartData | ChartSettingsProvidedData; date1904?: boolean; sourceRevision?: ChartSettingsSourceRevision; loadSource?: ChartSettingsSourceLoader; loadStatistics?: ChartSettingsStatisticsLoader; route?: ChartEditPanelRoute; routeKey?: number; navigationCurrent?: () => boolean; busy?: boolean; error?: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [patch: ChartPatch]; closed: [] }>()
const draft = ref<ChartSettingsDraft>()
const imageSession = ref(0), imageLoading = ref(false)
const localError = ref('')
const routeError = ref(''), composing = ref(false)
const sourceSheets = computed(() => props.sheets ?? (props.sheet ? [props.sheet] : []))
const activeTab = ref<ChartEditPanelTab>('style'), drawerReady = ref(false), navigationKey = ref(0)
const navigation = shallowRef<ChartEditPanelNavigation>()
const tabs = [{ key: 'style', label: '样式' }, { key: 'data', label: '数据' }]
let plannedRouteKey: string | undefined

watch([() => props.open, () => props.chart?.id, () => props.chart?.revision], ([open]) => {
  if (!open || !props.chart) return
  draft.value = createChartSettingsDraft(props.chart)
  imageSession.value++; imageLoading.value = false
  localError.value = ''; routeError.value = ''; composing.value = false
}, { immediate: true })

const { projection, loading: sourceLoading, assertReady, reload } = useChartSettingsSource({
  active: () => props.open, draft: () => draft.value, sheets: () => sourceSheets.value,
  hostSheetId: () => props.chart?.sheetId, provided: () => props.data, date1904: () => props.date1904,
  sourceRevision: () => props.sourceRevision, load: () => props.loadSource, loadStatistics: () => props.loadStatistics,
})

watch([() => props.open, () => props.chart?.id, () => props.routeKey, () => props.route, () => projection.value.data, () => projection.value.prepared,
  () => props.navigationCurrent?.() !== false, composing], ([open]) => {
  if (!open) { drawerReady.value = false; plannedRouteKey = undefined; composing.value = false; return }
  if (!props.chart || !draft.value || composing.value || props.navigationCurrent?.() === false) return
  const key = `${props.chart.id}:${props.routeKey ?? 0}:${JSON.stringify(props.route ?? null)}`
  if (plannedRouteKey === key) return
  try {
    const route = props.route === undefined ? undefined : normalizeChartEditPanelRoute(props.route)
    if (route?.seriesId && (!projection.value.prepared || !projection.value.data)) { navigation.value = undefined; activeTab.value = route.tab; return }
    const plan = route ? planChartEditPanelRoute(draft.value.config, route, projection.value.data) : undefined
    navigation.value = plan; navigationKey.value++; activeTab.value = plan?.route.tab ?? 'style'; plannedRouteKey = key; routeError.value = ''
  } catch (reason) {
    navigation.value = undefined; plannedRouteKey = key
    routeError.value = reason instanceof Error ? reason.message : '图表面板路由无法定位。'
  }
}, { immediate: true, flush: 'post' })

function apply(): void {
  if (!props.chart || !draft.value || props.busy || imageLoading.value) return
  try {
    const patch = chartSettingsPatch(draft.value)
    assertReady()
    localError.value = ''
    emit('apply', patch)
  } catch (error) { localError.value = error instanceof Error ? error.message : String(error) }
}
function close(): void { if (!props.busy) emit('update:open', false) }
function opened(value: boolean): void {
  drawerReady.value = value && props.open
  if (!value && !props.open) emit('closed')
}
</script>

<template>
  <Drawer :open="open" title="图表设置" placement="right" :size="560" :footer="null" :mask-closable="!busy" :keyboard="!busy" :closable="{ disabled: busy }" :after-open-change="opened" @close="close">
    <Form v-if="draft" layout="vertical" class="chart-editor-form" @submit.prevent="apply" @compositionstart.capture="composing = true" @compositionend.capture="composing = false">
      <Tabs v-model:active-key="activeTab" :items="tabs" :animated="false" destroy-on-hidden aria-label="图表设置标签页">
        <template #contentRender="{ item }">
          <ChartSettingsFields v-if="item.key === activeTab" :key="activeTab" v-model:draft="draft" :sheets="sourceSheets" :host-sheet-id="chart?.sheetId" :projection="projection" :date1904="date1904 ?? projection.data?.date1904" :active="open" :session-key="imageSession" :panel-tab="activeTab" :navigation="navigation" :navigation-key="navigationKey" :navigation-ready="drawerReady" :navigation-current="navigationCurrent" :disabled="busy" @resource-loading="imageLoading = $event" />
        </template>
      </Tabs>
      <Alert v-if="sourceLoading" type="info" message="正在读取图表数据源…" show-icon role="status" />
      <Button v-if="projection.error && (loadSource || loadStatistics)" :disabled="busy || sourceLoading" @click="reload">重新读取数据源</Button>
      <Alert v-if="localError || routeError || error" type="error" :message="localError || routeError || error" show-icon />
      <div class="chart-editor-actions">
        <Button :disabled="busy" @click="close">关闭</Button>
        <Button type="primary" :loading="busy" :disabled="sourceLoading || imageLoading" @click="apply">应用</Button>
      </div>
    </Form>
  </Drawer>
</template>

<style scoped>
.chart-editor-form { display: grid; gap: 14px; min-width: 0; }
.chart-editor-actions { display: flex; justify-content: flex-end; gap: 8px; padding-top: 12px; }
</style>
