<script setup lang="ts">
import { Button, Form, FormItem, Input, Select } from 'antdv-next'
import { computed, ref, watch } from 'vue'

const props = defineProps<{ mode: 'symbol' | 'emoji' }>()
const emit = defineEmits<{ insert: [text: string] }>()
const category = ref('常用')
const custom = ref('')
const recent = ref<string[]>([])
const symbolGroups: Record<string, string> = {
  常用: '© ® ™ § ¶ † ‡ • … ‰ ° ℃ ℉ № ✓ ✗ ★ ☆ ♥ ♦ ♣ ♠',
  数学: '± × ÷ ≠ ≈ ≤ ≥ ∞ √ ∑ ∏ ∫ ∂ ∆ ∇ ∈ ∉ ⊂ ⊃ ∩ ∪ ∅ ∀ ∃',
  箭头: '← ↑ → ↓ ↔ ↕ ↖ ↗ ↘ ↙ ↩ ↪ ⇐ ⇑ ⇒ ⇓ ⇔ ⇕ ➜ ➤ ➔ ➟',
  标点: '“ ” ‘ ’ « » ‹ › 「 」 『 』 【 】 〔 〕 〈 〉 《 》 — – · …',
  货币: '¥ $ € £ ₩ ₽ ₹ ₫ ฿ ₱ ₺ ₴ ₦ ¢ ¤',
  希腊字母: 'α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ τ υ φ χ ψ ω Γ Δ Θ Λ Ξ Π Σ Φ Ψ Ω',
}
const emojiGroups: Record<string, string> = {
  常用: '😀 😊 😂 🙂 😍 🤔 😎 👍 👏 🙏 ❤️ 🎉 ✅ ⭐ 📌 💡 🚀 📅 📊',
  表情: '😀 😃 😄 😁 😆 😅 😂 🙂 🙃 😉 😊 😍 😘 😋 😎 🤔 😐 😴 😢 😭 😡 😮 🤯 🥳',
  手势与人物: '👍 👎 👏 🙌 🤝 🙏 👋 ✋ ✌️ 💪 👩 👨 👧 👦 👨‍👩‍👧‍👦 👩‍💻 👨‍💻',
  自然与食物: '🌞 🌙 ⭐ 🌈 ☁️ 🌧️ ❄️ 🌸 🌳 🌱 🐱 🐶 🐼 🦋 🍎 🍊 🍋 🍇 ☕ 🍵 🍰',
  活动与物品: '🎉 🎊 🎁 🏆 ⚽ 🏀 🎵 🎨 📚 📎 📌 ✏️ 📝 📅 💻 📱 💡 🔑 🔒 📈 📉 📊',
  标记与交通: '❤️ 🧡 💛 💚 💙 💜 ✅ ❎ ❗ ❓ ⚠️ ⏰ ⏳ ♻️ 🚀 ✈️ 🚗 🚲 🚆 🏠 🌍',
}
const groups = computed(() => props.mode === 'emoji' ? emojiGroups : symbolGroups)
const options = computed(() => [...(recent.value.length ? [{ label: '最近使用', value: '最近使用' }] : []), ...Object.keys(groups.value).map((value) => ({ label: value, value }))])
const characters = computed(() => category.value === '最近使用' ? recent.value : (groups.value[category.value] ?? groups.value.常用 ?? '').split(' ').filter(Boolean))
watch(() => props.mode, () => { category.value = '常用'; custom.value = '' })
function insert(value: string): void {
  if (!value.trim()) return
  recent.value = [value, ...recent.value.filter((item) => item !== value)].slice(0, 32)
  emit('insert', value)
}
</script>

<template>
  <div class="document-character-picker" :aria-label="mode === 'emoji' ? '表情选择器' : '符号选择器'">
      <Form layout="vertical" class="document-character-picker__form" @submit.prevent>
        <FormItem label="分类"><Select v-model:value="category" :options="options" aria-label="字符分类" /></FormItem>
      </Form>
      <div class="character-grid" :aria-label="mode === 'emoji' ? '表情列表' : '符号列表'">
        <Button v-for="character in characters" :key="character" :aria-label="`插入 ${character}`" :title="character" @click="insert(character)">{{ character }}</Button>
      </div>
      <Form layout="vertical" class="document-character-picker__form" @submit.prevent>
        <FormItem label="自定义字符"><Input v-model:value="custom" aria-label="自定义字符" placeholder="输入或粘贴其他符号、表情及组合字符" @press-enter="insert(custom)" /></FormItem>
      </Form>
      <Button type="primary" :disabled="!custom.trim()" @click="insert(custom)">插入字符</Button>
  </div>
</template>

<style scoped>
.document-character-picker { display: flex; flex-direction: column; align-items: stretch; gap: 14px; padding-top: 4px; }
.document-character-picker__form { margin: 0; }
.document-character-picker__form :deep(.ant-form-item) { margin-bottom: 0; }
.character-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(42px, 1fr)); gap: 6px; min-height: 132px; max-height: 292px; overflow-y: auto; padding: 2px; }
.character-grid :deep(.ant-btn) { height: 42px; padding: 0; font: 22px/1.2 'Segoe UI Emoji', 'Segoe UI Symbol', 'Microsoft YaHei', sans-serif; }
.document-character-picker > :deep(.ant-btn) { align-self: flex-end; }
</style>
