<script setup lang="ts">
import { Alert, AutoComplete, Button, Checkbox, ColorPicker, Form, FormItem, Input, Modal, Select, Tooltip, theme } from 'antdv-next'
import { computed, inject, nextTick, reactive, ref, shallowRef, watch } from 'vue'
import type { ColorValueType, FormInstance } from 'antdv-next'
import { ArrowDown, ArrowUp, Plus, Trash2 } from 'lucide-vue-next'
import type { DocumentContentControl, DocumentContentControlType } from '../../domain/model'
import type { FDocument } from '../../api/facade'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'
import { cloneModel } from '../../domain/serialization'
import type { ContentControlPatch } from '../../features/document/services/content-controls'
import { assertContentControlPropertyDraftCurrent, contentControlPropertyDraftPatch, validateContentControlPropertyPatch } from '../../features/document/services/content-control-properties'
import { CONTENT_CONTROL_CHECKBOX_STYLES, CONTENT_CONTROL_DATE_FORMATS, contentControlCheckboxStyle, isContentControlCheckboxStyle, validateContentControlCheckboxState, validateContentControlDateProperties, validateContentControlListItems } from '../../features/document/services/content-control-presentation'

type ContentControlProperties = Partial<DocumentContentControl> & {
  date?: { format?: string; locale?: string; calendar?: string; storage?: string }
}

const props = defineProps<{
  open: boolean
  selectedControlId?: string
  documentApi?: FDocument
  readOnly?: boolean
}>()
const emit = defineEmits<{
  'update:open': [value: boolean]
  error: [message: string]
}>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const { token } = theme.useToken()
const form = ref<FormInstance>()
const baseline = shallowRef<DocumentContentControl>()
const saving = ref(false)
const propertyComposing = ref(false)
const customCheckboxEditing = ref(false)
const customCheckboxInput = ref<{ focus: () => void }>()
const validationError = ref('')
const listKeys = ref<string[]>([])
let nextListKey = 0
const themeStyle = computed(() => ({ '--control-properties-border': token.value.colorBorderSecondary }))

const properties = reactive<ContentControlProperties>({})
const selected = computed(() => {
  if (!props.open || !props.selectedControlId || !props.documentApi) return undefined
  return props.documentApi.getContentControl(props.selectedControlId)?.getInfo()
})
const contentControlTypes = computed<Array<{ label: string; value: DocumentContentControlType }>>(() => [
  { label: t('doc.content-control.type.richText', '富文本'), value: 'richText' },
  { label: t('doc.content-control.type.text', '纯文本'), value: 'text' },
  { label: t('doc.content-control.type.checkbox', '复选框'), value: 'checkbox' },
  { label: t('doc.content-control.type.date', '日期'), value: 'date' },
  { label: t('doc.content-control.type.dropDownList', '下拉列表'), value: 'dropDownList' },
  { label: t('doc.content-control.type.comboBox', '组合框'), value: 'comboBox' },
  { label: t('doc.content-control.type.picture', '图片'), value: 'picture' },
  { label: t('doc.content-control.type.group', '分组'), value: 'group' },
  { label: t('doc.content-control.type.repeatingSection', '重复节'), value: 'repeatingSection' },
  { label: t('doc.content-control.type.repeatingSectionItem', '重复节项'), value: 'repeatingSectionItem' },
  { label: t('doc.content-control.type.docPartObj', '文档部件'), value: 'docPartObj' },
])
const appearances = computed(() => [
  { label: t('doc.content-control.appearance.boundingBox', '边框'), value: 'boundingBox' },
  { label: t('doc.content-control.appearance.tags', '标签'), value: 'tags' },
  { label: t('doc.content-control.appearance.hidden', '隐藏'), value: 'hidden' },
])
const locks = computed(() => [
  { label: t('doc.content-control.lock.unlocked', '不锁定'), value: 'unlocked' },
  { label: t('doc.content-control.lock.sdtLocked', '锁定控件'), value: 'sdtLocked' },
  { label: t('doc.content-control.lock.contentLocked', '锁定内容'), value: 'contentLocked' },
  { label: t('doc.content-control.lock.sdtContentLocked', '内容和控件都锁定'), value: 'sdtContentLocked' },
])
const dateStorageOptions = computed(() => [
  { label: t('doc.content-control.date-storage.text', '文本'), value: 'text' },
  { label: t('doc.content-control.date-storage.date', '日期'), value: 'date' },
  { label: t('doc.content-control.date-storage.dateTime', '日期和时间'), value: 'dateTime' },
])
const hasPlaceholder = computed(() => selected.value != null && ![
  'checkbox', 'repeatingSection', 'repeatingSectionItem',
].includes(selected.value.type))
const hasDataBinding = computed(() => selected.value?.dataBinding != null)
const isDate = computed(() => selected.value?.type === 'date')
const isList = computed(() => selected.value?.type === 'dropDownList' || selected.value?.type === 'comboBox')
const isRepeatingSection = computed(() => selected.value?.type === 'repeatingSection')
const isDocPart = computed(() => selected.value?.type === 'docPartObj')
const isText = computed(() => selected.value?.type === 'text' || selected.value?.type === 'richText')
const isPlainText = computed(() => selected.value?.type === 'text')
const isCheckbox = computed(() => selected.value?.type === 'checkbox')
const contentLocked = computed(() => properties.lock === 'contentLocked' || properties.lock === 'sdtContentLocked')
const colorValue = computed(() => properties.color && /^#?[\da-f]{6}$/i.test(properties.color) ? `#${properties.color.replace(/^#/, '')}` : token.value.colorPrimary)
const checkboxStyle = computed(() => {
  if (customCheckboxEditing.value) return 'Custom'
  const control = selected.value
  return control ? contentControlCheckboxStyle({ ...control, checkbox: properties.checkbox, metadata: properties.metadata }) ?? 'Custom' : 'Custom'
})
const checkboxStyles = [
  { label: '叉号', value: 'Cross' }, { label: '勾号', value: 'Check' }, { label: '简式', value: 'Minimal' }, { label: '圆点', value: 'Round' }, { label: '自定义', value: 'Custom' },
]
const dateFormats = CONTENT_CONTROL_DATE_FORMATS.map(value => ({ value, label: value }))

watch(() => [props.open, props.selectedControlId, props.documentApi] as const, () => {
  const control = selected.value
  for (const key of Object.keys(properties) as Array<keyof ContentControlProperties>) delete properties[key]
  baseline.value = control ? cloneModel(control) : undefined
  validationError.value = ''; saving.value = false; propertyComposing.value = false; customCheckboxEditing.value = false
  if (control) {
    Object.assign(properties, cloneModel(control))
    properties.date = {
      format: control.dateFormat,
      locale: control.dateLocale,
      calendar: control.dateCalendar,
      storage: control.dateStorage,
    }
    listKeys.value = (control.items ?? []).map(() => `control-list-${nextListKey++}`)
  } else listKeys.value = []
  void nextTick(() => form.value?.clearValidate())
}, { immediate: true })

function patch(next: Partial<DocumentContentControl>): void {
  if (props.readOnly) return
  Object.assign(properties, next)
}

function patchDate(patch: NonNullable<ContentControlProperties['date']>): void {
  if (props.readOnly) return
  properties.date = { ...properties.date, ...patch }
}

function checkboxValue(value: boolean | { target?: { checked?: boolean } }): boolean {
  return typeof value === 'boolean' ? value : Boolean(value.target?.checked)
}

function addListItem(): void {
  if (props.readOnly) return
  properties.items = [...(properties.items ?? []), { displayText: '', value: '' }]
  listKeys.value.push(`control-list-${nextListKey++}`)
}
function updateListItem(index: number, key: 'displayText' | 'value', value: string): void {
  if (props.readOnly) return
  properties.items = properties.items?.map((item, itemIndex) => itemIndex === index ? { ...item, [key]: value } : item)
}
function removeListItem(index: number): void {
  if (props.readOnly) return
  properties.items = properties.items?.filter((_item, itemIndex) => itemIndex !== index)
  listKeys.value.splice(index, 1)
  void nextTick(() => form.value?.clearValidate())
}
function moveListItem(index: number, direction: number): void {
  if (props.readOnly) return
  const items = properties.items, target = index + direction
  if (!items || target < 0 || target >= items.length) return
  const next = [...items], keys = [...listKeys.value]
  ;[next[index], next[target]] = [next[target]!, next[index]!]
  ;[keys[index], keys[target]] = [keys[target]!, keys[index]!]
  properties.items = next; listKeys.value = keys
  void nextTick(() => form.value?.clearValidate())
}
const listRules = [{ validator: async () => validateContentControlListItems(properties.items ?? []) }]
function listValueRules(index: number) {
  return [{ validator: async (_rule: unknown, value: unknown) => {
    if (typeof value !== 'string' || !value.trim()) throw new Error('选项值不能为空。')
    if (properties.items?.some((item, itemIndex) => itemIndex !== index && item.value === value)) throw new Error('选项值不能重复。')
  } }]
}
const colorRules = [{ validator: async () => {
  if (properties.color !== baseline.value?.color) validateContentControlPropertyPatch({ color: properties.color })
} }]
const dateRules = [{ validator: async () => {
  const date = properties.date, original = baseline.value
  if (date?.format !== original?.dateFormat || date?.locale !== original?.dateLocale || date?.calendar !== original?.dateCalendar) {
    validateContentControlDateProperties(date?.format, date?.locale, date?.calendar)
  }
} }]
function glyphRules(key: 'checkedState' | 'uncheckedState') {
  return [{ validator: async () => validateContentControlCheckboxState(properties.checkbox?.[key] ?? {}) }]
}
function setCheckboxStyle(value: unknown): void {
  if (props.readOnly) return
  if (value === 'Custom') { customCheckboxEditing.value = true; void nextTick(() => customCheckboxInput.value?.focus()); return }
  if (typeof value !== 'string' || !isContentControlCheckboxStyle(value)) return
  customCheckboxEditing.value = false
  patch({ checkbox: { ...properties.checkbox, checked: properties.checkbox?.checked ?? selected.value?.checked ?? false, ...cloneModel(CONTENT_CONTROL_CHECKBOX_STYLES[value]) },
    metadata: { ...properties.metadata, checkboxStyle: value } })
}
function patchCheckboxState(key: 'checkedState' | 'uncheckedState', field: 'value' | 'font', value: string): void {
  if (props.readOnly) return
  customCheckboxEditing.value = true
  const metadata = { ...properties.metadata }
  delete metadata.checkboxStyle
  patch({ checkbox: { ...properties.checkbox, checked: properties.checkbox?.checked ?? selected.value?.checked ?? false,
    [key]: { ...properties.checkbox?.[key], [field]: value } }, metadata })
}
function updateLock(value: unknown): void {
  if (value === 'unlocked' || value === 'controlLocked' || value === 'sdtLocked' || value === 'contentLocked' || value === 'sdtContentLocked') patch({ lock: value })
}

function updateColor(value: ColorValueType): void {
  if (value === null || Array.isArray(value)) return
  const color = typeof value === 'string' ? value : value.toHexString()
  patch({ color: color.replace(/^#/, '').toUpperCase() })
}

function close(): void {
  propertyComposing.value = false
  emit('update:open', false)
}

async function save(): Promise<void> {
  const control = baseline.value, api = props.documentApi
  if (!control || !api || props.readOnly || saving.value || propertyComposing.value) return
  saving.value = true; validationError.value = ''
  try {
    await form.value?.validateFields()
    if (!props.open || props.readOnly || propertyComposing.value || props.documentApi !== api || props.selectedControlId !== control.id || baseline.value !== control) return
    const date = properties.date
    const draft: ContentControlPatch = {
      alias: properties.alias,
      tag: properties.tag,
      appearance: properties.appearance,
      lock: properties.lock,
      color: properties.color,
      placeholder: properties.placeholder,
      temporary: properties.temporary,
      multiLine: properties.multiLine,
      dataBinding: properties.dataBinding,
      metadata: properties.metadata,
    }
    if (control.type === 'date') Object.assign(draft, { dateFormat: date?.format, dateLocale: date?.locale, dateCalendar: date?.calendar, dateStorage: date?.storage })
    if (control.type === 'dropDownList' || control.type === 'comboBox') draft.items = properties.items
    if (control.type === 'repeatingSection') draft.repeatingSection = properties.repeatingSection
    if (control.type === 'docPartObj') draft.docPart = properties.docPart
    if (control.type === 'checkbox') draft.checkbox = properties.checkbox
    const submitted = contentControlPropertyDraftPatch(control, draft)
    const current = api.getContentControl(control.id)?.getInfo()
    if (!current) throw new Error('内容控件已删除，请重新选择。')
    assertContentControlPropertyDraftCurrent(control, current, submitted)
    validateContentControlPropertyPatch(submitted, current)
    if (Object.keys(submitted).length) {
      const changed = api.executeContentControlCommand('doc.command.update-content-control', { controlId: control.id, patch: submitted })
      if (!changed) throw new Error('内容控件属性修改未应用。')
    }
    close()
  } catch (error) {
    if (error instanceof Error) { validationError.value = error.message; emit('error', error.message) }
  } finally {
    saving.value = false
  }
}
</script>

<template>
  <Modal
    :open="open && selected != null"
    :title="t('doc.content-control.properties', '内容控件属性')"
    :width="640"
    :footer="null"
    destroy-on-hidden
    @cancel="close"
  >
    <Form v-if="selected" ref="form" :model="properties" :disabled="readOnly" layout="vertical" class="content-control-properties" :style="themeStyle"
      @submit.prevent="save" @compositionstart="propertyComposing = true" @compositionend="propertyComposing = false">
      <Alert v-if="validationError" type="error" :message="validationError" show-icon />
      <section class="properties-grid">
        <FormItem :label="t('doc.content-control.type', '类型')">
          <Select :value="selected.type" :options="contentControlTypes" disabled />
        </FormItem>
        <FormItem :label="t('doc.content-control.ooxml-id', 'OOXML 标识')">
          <Input :value="String(selected.metadata?.ooxmlId ?? '')" disabled />
        </FormItem>
        <FormItem name="alias" :label="t('doc.content-control.label-title', '标题')">
          <Input :value="properties.alias ?? selected.title ?? ''" @update:value="value => patch({ alias: value })" />
        </FormItem>
        <FormItem name="tag" :label="t('doc.content-control.tag', '标记')">
          <Input :value="properties.tag ?? ''" @update:value="value => patch({ tag: value })" />
        </FormItem>
        <FormItem name="appearance" :label="t('doc.content-control.appearance', '外观')">
          <Select :value="properties.appearance ?? 'boundingBox'" :options="appearances" @change="value => patch({ appearance: String(value) })" />
        </FormItem>
        <FormItem name="lock" :label="t('doc.content-control.lock', '锁定')">
          <Select :value="properties.lock === 'controlLocked' ? 'sdtLocked' : properties.lock ?? 'unlocked'" :options="locks" @change="updateLock" />
        </FormItem>
        <FormItem name="color" :label="t('doc.content-control.color', '颜色')" :rules="colorRules">
          <div class="color-input">
            <ColorPicker
              :value="colorValue"
              value-format="hex"
              mode="single"
              :aria-label="t('doc.content-control.color', '颜色')"
              @update:value="updateColor"
            />
            <Input :value="properties.color ?? ''" :placeholder="token.colorPrimary" @update:value="value => patch({ color: value.replace(/^#/, '').toUpperCase() })" />
          </div>
        </FormItem>
        <FormItem v-if="hasPlaceholder" name="placeholder" :label="t('doc.content-control.placeholder', '占位文本')">
          <Input :value="properties.placeholder ?? ''" :disabled="readOnly || (contentLocked && selected.showingPlaceholder)" @update:value="value => patch({ placeholder: value })" />
        </FormItem>
      </section>

      <section v-if="isText" class="properties-section">
        <FormItem v-if="isPlainText" name="multiLine"><Checkbox :checked="properties.multiLine ?? false" @change="value => patch({ multiLine: checkboxValue(value) })">{{ t('doc.content-control.multiline', '允许多行') }}</Checkbox></FormItem>
        <FormItem name="temporary"><Checkbox :checked="properties.temporary ?? false" @change="value => patch({ temporary: checkboxValue(value) })">{{ t('doc.content-control.temporary', '临时控件') }}</Checkbox></FormItem>
      </section>

      <section v-if="isCheckbox" class="properties-grid properties-section">
        <FormItem :label="t('doc.content-control.checkbox-style', '复选框样式')">
          <Select :value="checkboxStyle" :options="checkboxStyles" :disabled="readOnly || contentLocked" @change="setCheckboxStyle" />
        </FormItem>
        <FormItem :name="['checkbox', 'checkedState', 'value']" :label="'选中符号（十六进制）'" :rules="glyphRules('checkedState')">
          <Input ref="customCheckboxInput" :value="properties.checkbox?.checkedState?.value ?? '2612'" :disabled="readOnly || contentLocked" @update:value="value => patchCheckboxState('checkedState', 'value', value)" />
        </FormItem>
        <FormItem :name="['checkbox', 'checkedState', 'font']" :label="'选中符号字体'" :rules="glyphRules('checkedState')">
          <Input :value="properties.checkbox?.checkedState?.font ?? 'Segoe UI Symbol'" :disabled="readOnly || contentLocked" @update:value="value => patchCheckboxState('checkedState', 'font', value)" />
        </FormItem>
        <FormItem :name="['checkbox', 'uncheckedState', 'value']" :label="'未选中符号（十六进制）'" :rules="glyphRules('uncheckedState')">
          <Input :value="properties.checkbox?.uncheckedState?.value ?? '2610'" :disabled="readOnly || contentLocked" @update:value="value => patchCheckboxState('uncheckedState', 'value', value)" />
        </FormItem>
        <FormItem :name="['checkbox', 'uncheckedState', 'font']" :label="'未选中符号字体'" :rules="glyphRules('uncheckedState')">
          <Input :value="properties.checkbox?.uncheckedState?.font ?? 'Segoe UI Symbol'" :disabled="readOnly || contentLocked" @update:value="value => patchCheckboxState('uncheckedState', 'font', value)" />
        </FormItem>
      </section>

      <section v-if="isDate" class="properties-grid properties-section">
        <FormItem :name="['date', 'format']" :label="t('doc.content-control.date-format', '日期格式')" :rules="dateRules">
          <AutoComplete :value="properties.date?.format ?? ''" :options="dateFormats" :disabled="readOnly || (contentLocked && !selected.showingPlaceholder && !!selected.dateValue)"
            @update:value="value => { if (typeof value === 'string') patchDate({ format: value }) }" />
        </FormItem>
        <FormItem :name="['date', 'locale']" :label="t('doc.content-control.date-locale', '区域设置')" :rules="dateRules">
          <Input :value="properties.date?.locale ?? ''" :disabled="readOnly || (contentLocked && !selected.showingPlaceholder && !!selected.dateValue)" @update:value="value => patchDate({ locale: value })" />
        </FormItem>
        <FormItem :name="['date', 'calendar']" :label="t('doc.content-control.date-calendar', '日历')" :rules="dateRules">
          <Input :value="properties.date?.calendar ?? ''" :disabled="readOnly || (contentLocked && !selected.showingPlaceholder && !!selected.dateValue)" @update:value="value => patchDate({ calendar: value })" />
        </FormItem>
        <FormItem :name="['date', 'storage']" :label="t('doc.content-control.date-storage', '存储格式')">
          <Select :value="properties.date?.storage ?? 'text'" :options="dateStorageOptions" @change="value => patchDate({ storage: String(value) })" />
        </FormItem>
      </section>

      <FormItem v-if="isList" name="items" class="list-items" :label="t('doc.content-control.list-items', '列表项')" :rules="listRules">
        <div v-for="(item, index) in properties.items ?? []" :key="listKeys[index]" class="list-item-row">
          <FormItem :name="['items', index, 'displayText']" :label="'显示文字'">
            <Input :value="item.displayText ?? ''" :disabled="readOnly || contentLocked" @update:value="value => updateListItem(index, 'displayText', value)" />
          </FormItem>
          <FormItem :name="['items', index, 'value']" :label="'选项值'" :rules="listValueRules(index)">
            <Input :value="item.value" :disabled="readOnly || contentLocked" @update:value="value => updateListItem(index, 'value', value)" />
          </FormItem>
          <div class="list-item-actions">
            <Tooltip title="上移"><Button html-type="button" type="text" size="small" :disabled="readOnly || contentLocked || index === 0" aria-label="上移列表项" @click="moveListItem(index, -1)"><ArrowUp :size="14" /></Button></Tooltip>
            <Tooltip title="下移"><Button html-type="button" type="text" size="small" :disabled="readOnly || contentLocked || index === (properties.items?.length ?? 0) - 1" aria-label="下移列表项" @click="moveListItem(index, 1)"><ArrowDown :size="14" /></Button></Tooltip>
            <Tooltip title="删除"><Button html-type="button" type="text" size="small" :disabled="readOnly || contentLocked" aria-label="删除列表项" @click="removeListItem(index)"><Trash2 :size="14" /></Button></Tooltip>
          </div>
        </div>
        <Button html-type="button" :disabled="readOnly || contentLocked" @click="addListItem"><Plus :size="14" />新增选项</Button>
      </FormItem>

      <section v-if="isRepeatingSection" class="properties-grid properties-section">
        <FormItem :name="['repeatingSection', 'sectionTitle']" :label="t('doc.content-control.section-title', '重复节标题')">
          <Input :value="properties.repeatingSection?.sectionTitle ?? ''" @update:value="value => patch({ repeatingSection: { ...properties.repeatingSection, sectionTitle: value } })" />
        </FormItem>
        <FormItem :name="['repeatingSection', 'doNotAllowInsertDelete']"><Checkbox
          :checked="properties.repeatingSection?.doNotAllowInsertDelete ?? false"
          @change="value => patch({ repeatingSection: { ...properties.repeatingSection, doNotAllowInsertDelete: checkboxValue(value) } })"
        >
          {{ t('doc.content-control.prevent-repeat-changes', '禁止插入或删除重复项') }}
        </Checkbox></FormItem>
      </section>

      <section v-if="isDocPart" class="properties-grid properties-section">
        <FormItem :name="['docPart', 'gallery']" :label="t('doc.content-control.gallery', '文档部件库')">
          <Input :value="properties.docPart?.gallery ?? ''" @update:value="value => patch({ docPart: { ...properties.docPart, gallery: value } })" />
        </FormItem>
        <FormItem :name="['docPart', 'category']" :label="t('doc.content-control.category', '类别')">
          <Input :value="properties.docPart?.category ?? ''" @update:value="value => patch({ docPart: { ...properties.docPart, category: value } })" />
        </FormItem>
        <FormItem :name="['docPart', 'unique']"><Checkbox
          :checked="properties.docPart?.unique ?? false"
          @change="value => patch({ docPart: { ...properties.docPart, unique: checkboxValue(value) } })"
        >
          {{ t('doc.content-control.unique', '唯一部件') }}
        </Checkbox></FormItem>
      </section>

      <section v-if="hasDataBinding" class="properties-grid properties-section">
        <FormItem :name="['dataBinding', 'xpath']" :label="t('doc.content-control.binding-xpath', '绑定 XPath')">
          <Input :value="properties.dataBinding?.xpath ?? ''" @update:value="value => patch({ dataBinding: { ...properties.dataBinding, xpath: value, dirty: true } })" />
        </FormItem>
        <FormItem :name="['dataBinding', 'storeItemID']" :label="t('doc.content-control.binding-store', '数据存储 ID')">
          <Input :value="properties.dataBinding?.storeItemID ?? ''" @update:value="value => patch({ dataBinding: { ...properties.dataBinding, storeItemID: value, dirty: true } })" />
        </FormItem>
      </section>

      <footer class="properties-footer">
        <Button html-type="button" :disabled="false" @click="close">{{ t('common.cancel', '取消') }}</Button>
        <Button type="primary" html-type="submit" :loading="saving" :disabled="readOnly || propertyComposing">{{ t('common.save', '保存') }}</Button>
      </footer>
    </Form>
  </Modal>
</template>

<style scoped>
.content-control-properties { display: flex; max-height: 72vh; flex-direction: column; gap: 16px; overflow: auto; padding: 4px; }
.properties-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
.properties-grid :deep(.ant-form-item), .list-items { min-width: 0; margin: 0; }
.properties-grid :deep(.ant-form-item-label), .list-items :deep(.ant-form-item-label) { padding: 0 0 6px; }
.color-input { display: flex; min-width: 0; align-items: center; gap: 8px; }
.properties-section { border-top: 1px solid var(--control-properties-border); padding-top: 14px; }
.properties-footer { display: flex; justify-content: flex-end; gap: 8px; border-top: 1px solid var(--control-properties-border); padding-top: 14px; }
.list-item-row { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) auto; align-items: start; gap: 8px; margin-bottom: 10px; }
.list-item-row :deep(.ant-form-item) { margin-bottom: 0; }
.list-item-actions { display: flex; align-items: center; gap: 2px; padding-top: 28px; }
@media (max-width: 560px) { .properties-grid { grid-template-columns: 1fr; } }
</style>
