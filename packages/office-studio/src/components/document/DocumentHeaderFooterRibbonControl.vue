<script setup lang="ts">
import { Button, Checkbox, Form, FormItem, InputNumber, Select, Popover, Modal, Menu } from 'antdv-next'
import { CalendarClock, ChevronDown, Hash } from 'lucide-vue-next'
import { computed, ref } from 'vue'

export type HeaderFooterRibbonKind = 'insertion' | 'navigation' | 'options' | 'position'
export interface HeaderFooterRibbonState {
  visible: boolean
  area: 'header' | 'footer'
  disabled: boolean
  contentDisabled?: boolean
  canLinkToPrevious: boolean
  linkedToPrevious: boolean
  differentFirstPage: boolean
  differentOddEven: boolean
  marginHeader: number
  marginFooter: number
}

export type HeaderFooterAction =
  | { action: 'insert-field'; field: 'PAGE' | 'NUMPAGES' }
  | { action: 'insert-date-time'; format: 'short' | 'long' | 'time'; automatic: boolean }
  | { action: 'refresh-fields' }
  | { action: 'navigate'; area: 'header' | 'footer' }
  | { action: 'set-options'; patch: { linkToPrevious?: boolean; differentFirstPage?: boolean; differentOddEven?: boolean; marginHeader?: number; marginFooter?: number } }

const props = defineProps<{ kind: HeaderFooterRibbonKind; state: HeaderFooterRibbonState; classic?: boolean }>()
const emit = defineEmits<{ insert: [value: string] }>()
const dateFormat = ref<'short' | 'long' | 'time'>('short')
const automaticDateTime = ref(true)
const dateFormatOptions = [
  { value: 'short', label: '短日期' },
  { value: 'long', label: '长日期' },
  { value: 'time', label: '时间' },
]
const area = computed(() => props.state.area)
const pageOpen = ref(false), dateOpen = ref(false)

function run(action: HeaderFooterAction): void {
  if (props.state.disabled || props.state.contentDisabled && ['insert-field', 'insert-date-time', 'refresh-fields'].includes(action.action)) return
  pageOpen.value = dateOpen.value = false
  emit('insert', JSON.stringify(action))
}

function setOptions(patch: Extract<HeaderFooterAction, { action: 'set-options' }>['patch']): void {
  run({ action: 'set-options', patch })
}

function checkboxValue(value: boolean | { target?: { checked?: boolean } }): boolean {
  return typeof value === 'boolean' ? value : Boolean(value.target?.checked)
}
</script>

<template>
  <div class="header-footer-ribbon-control" :class="{ 'is-classic': classic }" :data-control-kind="kind">
    <fieldset v-if="kind === 'insertion'" :disabled="state.disabled || state.contentDisabled" class="header-footer-ribbon-control__fields">
      <Popover v-model:open="pageOpen" placement="bottomLeft" trigger="click">
        <Button type="text" size="small" aria-label="页码" @mousedown.prevent><Hash :size="14" />页码<ChevronDown :size="12" /></Button>
        <template #content><Menu :selectable="false" :items="[{ key: 'PAGE', label: '当前页码' }, { key: 'NUMPAGES', label: '总页数' }, { type: 'divider' }, { key: 'refresh', label: '更新字段' }]" @click="({ key }) => key === 'refresh' ? run({ action: 'refresh-fields' }) : run({ action: 'insert-field', field: key === 'PAGE' ? 'PAGE' : 'NUMPAGES' })" /></template>
      </Popover>
      <Button type="text" size="small" @mousedown.prevent @click="dateOpen = true"><CalendarClock :size="14" />日期和时间</Button>
      <Modal v-model:open="dateOpen" title="日期和时间" ok-text="插入" @ok="run({ action: 'insert-date-time', format: dateFormat, automatic: automaticDateTime })">
        <Form layout="vertical"><FormItem label="日期和时间格式"><Select v-model:value="dateFormat" :options="dateFormatOptions" /></FormItem><FormItem><Checkbox v-model:checked="automaticDateTime">自动更新</Checkbox></FormItem></Form>
      </Modal>
    </fieldset>

    <template v-else-if="kind === 'navigation'">
      <Button size="small" :disabled="state.disabled || area === 'header'" @click="run({ action: 'navigate', area: 'header' })">转到页眉</Button>
      <Button size="small" :disabled="state.disabled || area === 'footer'" @click="run({ action: 'navigate', area: 'footer' })">转到页脚</Button>
      <Checkbox :checked="state.linkedToPrevious" :disabled="state.disabled || !state.canLinkToPrevious" @change="setOptions({ linkToPrevious: checkboxValue($event) })">链接到前一节</Checkbox>
    </template>

    <template v-else-if="kind === 'options'">
      <Checkbox :checked="state.differentFirstPage" :disabled="state.disabled" @change="setOptions({ differentFirstPage: checkboxValue($event) })">首页不同</Checkbox>
      <Checkbox :checked="state.differentOddEven" :disabled="state.disabled" @change="setOptions({ differentOddEven: checkboxValue($event) })">奇偶页不同</Checkbox>
    </template>

    <Form v-else layout="vertical" class="header-footer-ribbon-control__form" @submit.prevent>
      <FormItem label="页眉距顶端">
        <InputNumber :value="state.marginHeader" :min="0" :max="200" :precision="1" size="small" :disabled="state.disabled" @change="setOptions({ marginHeader: Number($event ?? 0) })" />
      </FormItem>
      <FormItem label="页脚距底端">
        <InputNumber :value="state.marginFooter" :min="0" :max="200" :precision="1" size="small" :disabled="state.disabled" @change="setOptions({ marginFooter: Number($event ?? 0) })" />
      </FormItem>
    </Form>
  </div>
</template>

<style scoped>
.header-footer-ribbon-control { display: grid; width: max-content; align-items: center; gap: 2px 8px; white-space: nowrap; }
.header-footer-ribbon-control[data-control-kind="navigation"] { grid-template-rows: repeat(2, auto); grid-auto-flow: column; }
.header-footer-ribbon-control.is-classic, .is-classic .header-footer-ribbon-control__fields, .is-classic .header-footer-ribbon-control__form { display: flex; align-items: center; }
.header-footer-ribbon-control__fields { display: grid; gap: 4px; border: none; padding: 0; margin: 0; }
.header-footer-ribbon-control__row { display: flex; align-items: center; gap: 6px; }
.header-footer-ribbon-control__row :deep(.ant-select) { min-width: 92px; }
.header-footer-ribbon-control__form { display: grid; gap: 2px; }
.header-footer-ribbon-control__form :deep(.ant-form-item) { display: grid; grid-template-columns: minmax(0, 1fr) 92px; align-items: center; gap: 12px; margin-bottom: 0; white-space: nowrap; }
.header-footer-ribbon-control__form :deep(.ant-form-item-label), .header-footer-ribbon-control__form :deep(.ant-form-item-control) { min-width: 0; padding: 0; }
</style>
