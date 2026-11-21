<script setup lang="ts">
import { Button, Checkbox, InputNumber, Select } from 'antdv-next'
import { computed, ref } from 'vue'

export type HeaderFooterRibbonKind = 'insertion' | 'navigation' | 'options' | 'position'
export interface HeaderFooterRibbonState {
  visible: boolean
  area: 'header' | 'footer'
  disabled: boolean
  canLinkToPrevious: boolean
  linkedToPrevious: boolean
  differentFirstPage: boolean
  differentOddEven: boolean
  marginHeader: number
  marginFooter: number
}

type HeaderFooterAction =
  | { action: 'insert-field'; field: 'PAGE' | 'NUMPAGES' }
  | { action: 'insert-date-time'; format: 'short' | 'long' | 'time'; automatic: boolean }
  | { action: 'refresh-fields' }
  | { action: 'navigate'; area: 'header' | 'footer' }
  | { action: 'set-options'; patch: { linkToPrevious?: boolean; differentFirstPage?: boolean; differentOddEven?: boolean; marginHeader?: number; marginFooter?: number } }

const props = defineProps<{ kind: HeaderFooterRibbonKind; state: HeaderFooterRibbonState }>()
const emit = defineEmits<{ insert: [value: string] }>()
const dateFormat = ref<'short' | 'long' | 'time'>('short')
const automaticDateTime = ref(true)
const dateFormatOptions = [
  { value: 'short', label: '短日期' },
  { value: 'long', label: '长日期' },
  { value: 'time', label: '时间' },
]
const area = computed(() => props.state.area)

function run(action: HeaderFooterAction): void {
  if (props.state.disabled) return
  emit('insert', JSON.stringify(action))
}

function setOptions(patch: Extract<HeaderFooterAction, { action: 'set-options' }>['patch']): void {
  run({ action: 'set-options', patch })
}
</script>

<template>
  <div class="header-footer-ribbon-control" :data-control-kind="kind">
    <template v-if="kind === 'insertion'">
      <Button type="text" size="small" :disabled="state.disabled" @click="run({ action: 'insert-field', field: 'PAGE' })">插入页码</Button>
      <Button type="text" size="small" :disabled="state.disabled" @click="run({ action: 'insert-field', field: 'NUMPAGES' })">插入总页数</Button>
      <div class="header-footer-ribbon-control__row">
        <Select v-model:value="dateFormat" size="small" :options="dateFormatOptions" aria-label="日期和时间格式" />
        <Checkbox v-model:checked="automaticDateTime">自动更新</Checkbox>
        <Button size="small" :disabled="state.disabled" @click="run({ action: 'insert-date-time', format: dateFormat, automatic: automaticDateTime })">插入</Button>
      </div>
      <Button type="text" size="small" :disabled="state.disabled" @click="run({ action: 'refresh-fields' })">更新字段</Button>
    </template>

    <template v-else-if="kind === 'navigation'">
      <Button size="small" :disabled="state.disabled || area === 'header'" @click="run({ action: 'navigate', area: 'header' })">转到页眉</Button>
      <Button size="small" :disabled="state.disabled || area === 'footer'" @click="run({ action: 'navigate', area: 'footer' })">转到页脚</Button>
      <Checkbox :checked="state.linkedToPrevious" :disabled="state.disabled || !state.canLinkToPrevious" @change="setOptions({ linkToPrevious: Boolean($event) })">链接到前一节</Checkbox>
    </template>

    <template v-else-if="kind === 'options'">
      <Checkbox :checked="state.differentFirstPage" :disabled="state.disabled" @change="setOptions({ differentFirstPage: Boolean($event) })">首页不同</Checkbox>
      <Checkbox :checked="state.differentOddEven" :disabled="state.disabled" @change="setOptions({ differentOddEven: Boolean($event) })">奇偶页不同</Checkbox>
    </template>

    <template v-else>
      <label class="header-footer-ribbon-control__field">
        <span>页眉距顶端</span>
        <InputNumber :value="state.marginHeader" :min="0" :max="200" :precision="1" size="small" :disabled="state.disabled" @change="setOptions({ marginHeader: Number($event ?? 0) })" />
      </label>
      <label class="header-footer-ribbon-control__field">
        <span>页脚距底端</span>
        <InputNumber :value="state.marginFooter" :min="0" :max="200" :precision="1" size="small" :disabled="state.disabled" @change="setOptions({ marginFooter: Number($event ?? 0) })" />
      </label>
    </template>
  </div>
</template>

<style scoped>
.header-footer-ribbon-control { display: grid; min-width: 208px; gap: 4px; padding: 4px; }
.header-footer-ribbon-control__row { display: flex; align-items: center; gap: 6px; }
.header-footer-ribbon-control__row :deep(.ant-select) { min-width: 92px; }
.header-footer-ribbon-control__field { display: grid; grid-template-columns: minmax(0, 1fr) 92px; align-items: center; gap: 12px; white-space: nowrap; }
</style>
