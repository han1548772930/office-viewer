import { inject, type InjectionKey, type Ref } from 'vue'

export const WORKBENCH_CHROME = ['full', 'no-ribbon', 'canvas-only'] as const
export type WorkbenchChrome = typeof WORKBENCH_CHROME[number]
export const WORKBENCH_ZOOMS = [50, 75, 100, 125, 150, 175, 200, 400] as const
export type WorkbenchZoom = typeof WORKBENCH_ZOOMS[number]
export const WORKBENCH_ACCENTS = {
  blue: '#2C53F1', green: '#047857', orange: '#EA580C', purple: '#7C3AED',
  red: '#E11D48', yellow: '#A16207', 'dark-blue': '#93C5FD',
} as const
export type WorkbenchAccent = keyof typeof WORKBENCH_ACCENTS
export function workbenchVisibility(mode: WorkbenchChrome) {
  return { toolbar: mode === 'full', header: mode !== 'canvas-only', footer: mode !== 'canvas-only', leftSidebar: mode !== 'canvas-only' }
}
export interface WorkbenchDisplay { chrome: WorkbenchChrome; zoom: WorkbenchZoom }
export const WorkbenchDisplayKey: InjectionKey<Readonly<Ref<WorkbenchDisplay>>> = Symbol('WorkbenchDisplay')
export function useWorkbenchDisplay() { return inject(WorkbenchDisplayKey, undefined) }
