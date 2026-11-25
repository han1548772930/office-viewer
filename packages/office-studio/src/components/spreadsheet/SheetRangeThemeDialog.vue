<script setup lang="ts">
import { computed, inject, nextTick, ref, watch } from 'vue'
import { Alert, Button, Form, FormItem, Input, Modal, Select, theme } from 'antdv-next'
import type { SheetRangeThemeStyle } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { SHEET_BUILTIN_RANGE_THEME_NAMES } from '../../features/sheet-workbench/sheet-range-theme-presets.ts'
import { formatValidationRange } from '../../features/sheet-validation/validation-panel-state.ts'
import { prepareSheetRangeThemeEditorDraft, uniqueSheetRangeThemeDraftName } from '../../features/sheet-workbench/sheet-range-theme-editor-state.ts'
import type { SheetRangeThemeStyleJson } from '../../features/sheet-workbench/sheet-range-theme-style.ts'
import type { SheetRangeThemeRequest } from './use-sheet-range-theme.ts'
import SheetRangeThemeStyleEditor from './SheetRangeThemeStyleEditor.vue'

const props = defineProps<{ request: SheetRangeThemeRequest | null; error: string }>()
const emit = defineEmits<{ confirm: [action: 'apply' | 'remove', themeName?: string, draft?: SheetRangeThemeStyleJson]; cancel: []; closed: [] }>()
const tr = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const selected = ref<string>(), select = ref<InstanceType<typeof Select> | null>(null), invalid = ref(''), composing = ref(false)
const draft = ref<SheetRangeThemeStyle | null>(null), originalDraft = ref<SheetRangeThemeStyle | null>(null)
const draftName = ref(''), editingName = ref<string>(), nameInput = ref<InstanceType<typeof Input> | null>(null)
const builtin = new Set(SHEET_BUILTIN_RANGE_THEME_NAMES)
const options = computed(() => (props.request?.themeNames ?? []).map(name => ({ value: name,
  label: builtin.has(name) ? name : `${tr('sheet.range-theme.custom', '自定义')} · ${name}` })))
const rangeLabel = computed(() => props.request?.ranges.map(range => formatValidationRange(range)).join('、') ?? '')
const errorMessage = computed(() => invalid.value || props.error)
const formModel = computed(() => ({ themeName: selected.value, ...(draft.value ?? {}), name: draftName.value }))
const editingExisting = computed(() => editingName.value !== undefined && draftName.value.trim() === editingName.value)
watch(() => props.request, request => {
  selected.value = request?.initialTheme; invalid.value = ''; composing.value = false; discardDraft()
}, { immediate: true })
function discardDraft(): void { draft.value = null; originalDraft.value = null; draftName.value = ''; editingName.value = undefined; invalid.value = '' }
function startDraft(source?: SheetRangeThemeStyle): void {
  if (!props.request || composing.value) return
  if (source && !builtin.has(source.name)) { editingName.value = source.name; draftName.value = source.name }
  else {
    editingName.value = undefined
    draftName.value = uniqueSheetRangeThemeDraftName(source ? `${source.name} ${tr('sheet.range-theme.custom', '自定义')}`
      : tr('sheet.range-theme.new-name', '自定义主题'), props.request.themeNames)
  }
  draft.value = source ? { ...cloneModel(source), name: draftName.value } : { name: draftName.value }
  originalDraft.value = cloneModel(draft.value); invalid.value = ''
  const request = props.request, current = draft.value
  void nextTick(() => { if (props.request === request && draft.value === current && !composing.value) nameInput.value?.focus() })
}
function editSelected(): void {
  const source = selected.value ? props.request?.themeStyles[selected.value] : undefined
  if (source) startDraft(source)
}
function resetDraft(): void {
  if (!originalDraft.value) return
  draft.value = cloneModel(originalDraft.value); draftName.value = originalDraft.value.name; invalid.value = ''
}
function apply(): void {
  if (composing.value) return
  if (draft.value && props.request) {
    try {
      const author = prepareSheetRangeThemeEditorDraft(draft.value, draftName.value, props.request.themeNames, editingName.value)
      emit('confirm', 'apply', author.name, author)
    } catch (error: unknown) { invalid.value = error instanceof Error ? error.message : tr('sheet.range-theme.invalid-style', '范围主题样式无效。') }
    return
  }
  if (!selected.value || !props.request?.themeNames.includes(selected.value)) { invalid.value = tr('sheet.range-theme.choose', '请选择可用的范围主题。'); return }
  emit('confirm', 'apply', selected.value)
}
function afterOpenChange(open: boolean): void {
  if (!open || !props.request || composing.value) return
  if (draft.value) nameInput.value?.focus(); else select.value?.focus()
}
</script>

<template>
  <Modal :open="Boolean(request)" :title="tr('sheet.range-theme.title', '范围主题')" :width="draft ? 760 : 480"
    :mask="{ closable: false }" :focus-trigger-after-close="false" :after-open-change="afterOpenChange"
    @cancel="emit('cancel')" @after-close="emit('closed')">
    <div data-sheet-range-theme-dialog :style="{ maxHeight: '70vh', overflowY: 'auto' }" @compositionstart="composing = true" @compositionend="composing = false">
      <p :style="{ color: token.colorTextSecondary }">{{ tr('sheet.range-theme.description', '范围主题包含表头、末行和行列条纹。单元格格式优先于范围主题。') }}</p>
      <Form layout="vertical" :model="formModel" @submit.prevent>
        <FormItem :label="tr('sheet.range-theme.ranges', '应用区域')"><span>{{ rangeLabel }}</span></FormItem>
        <FormItem name="themeName" :label="tr('sheet.range-theme.theme', '主题')" :validate-status="!draft && errorMessage ? 'error' : undefined" :help="!draft && errorMessage || undefined">
          <Select ref="select" v-model:value="selected" :options="options" show-search option-filter-prop="label"
            :aria-label="tr('sheet.range-theme.theme', '主题')" :disabled="Boolean(draft)" @change="invalid = ''" />
        </FormItem>
        <div v-if="!draft" :style="{ display: 'flex', gap: `${token.marginXS}px`, marginBottom: `${token.marginSM}px` }">
          <Button :disabled="!selected || composing" @click="editSelected">{{ selected && builtin.has(selected) ? tr('sheet.range-theme.edit-copy', '编辑副本') : tr('sheet.range-theme.edit', '编辑主题') }}</Button>
          <Button :disabled="composing" @click="startDraft()">{{ tr('sheet.range-theme.new', '新建主题') }}</Button>
        </div>
        <template v-else>
          <FormItem name="name" :label="tr('sheet.range-theme.name', '主题名称')" :validate-status="errorMessage ? 'error' : undefined" :help="errorMessage || undefined">
            <Input ref="nameInput" v-model:value="draftName" :aria-label="tr('sheet.range-theme.name', '主题名称')" @update:value="invalid = ''" />
          </FormItem>
          <p :style="{ color: token.colorTextSecondary }">{{ editingExisting ? tr('sheet.range-theme.edit-scope', '保存会更新所有使用此主题的区域，并将主题应用到当前选区。') : tr('sheet.range-theme.create-scope', '保存会创建自定义主题并应用到当前选区。') }}</p>
          <SheetRangeThemeStyleEditor v-model="draft" @error="invalid = $event" />
          <div :style="{ display: 'flex', gap: `${token.marginXS}px`, marginBottom: `${token.marginSM}px` }">
            <Button :disabled="composing" @click="resetDraft">{{ tr('sheet.range-theme.restore-draft', '恢复初始配置') }}</Button>
            <Button :disabled="composing" @click="discardDraft">{{ tr('sheet.range-theme.discard-draft', '放弃编辑') }}</Button>
          </div>
        </template>
      </Form>
      <Alert v-if="request?.mixed" type="info" show-icon :title="tr('sheet.range-theme.mixed', '选区使用了不同的范围主题。')" />
      <p v-if="request?.canRemove" :style="{ color: token.colorTextSecondary }">{{ tr('sheet.range-theme.remove-scope', '移除主题会移除与选区相交的主题规则及其整个应用区域。') }}</p>
    </div>
    <template #footer>
      <div :style="{ display: 'flex', justifyContent: 'space-between', gap: `${token.marginXS}px` }">
        <Button :disabled="!request?.canRemove || composing" @click="emit('confirm', 'remove')">{{ tr('sheet.range-theme.remove', '移除主题') }}</Button>
        <div :style="{ display: 'flex', gap: `${token.marginXS}px` }">
          <Button @click="emit('cancel')">{{ tr('sheet.range-theme.cancel', '取消') }}</Button>
          <Button type="primary" :disabled="composing" @click="apply">{{ draft ? tr('sheet.range-theme.save-apply', '保存并应用') : tr('sheet.range-theme.apply', '应用') }}</Button>
        </div>
      </div>
    </template>
  </Modal>
</template>
