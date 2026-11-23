import type { ThemeConfig } from 'antdv-next'

/** Popup menus use their popup surface border instead of Menu's navigation divider. */
export const WORKBENCH_CONTROL_THEME = {
  components: {
    Menu: { activeBarBorderWidth: 0 },
  },
} satisfies ThemeConfig
