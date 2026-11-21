import { computed } from 'vue'
import { theme } from 'antdv-next'

/** Also applied to portaled content so ConfigProvider tokens reach every Ribbon surface. */
export function useRibbonTheme() {
  const { token } = theme.useToken()
  return computed(() => ({
    '--ribbon-background': token.value.colorBgContainer,
    '--ribbon-surface': token.value.colorFillAlter,
    '--ribbon-text': token.value.colorText,
    '--ribbon-text-secondary': token.value.colorTextSecondary,
    '--ribbon-disabled': token.value.colorTextDisabled,
    '--ribbon-border': token.value.colorBorderSecondary,
    '--ribbon-primary': token.value.colorPrimary,
    '--ribbon-active-background': token.value.colorPrimaryBg,
    '--ribbon-hover-background': token.value.colorFillTertiary,
  }))
}
