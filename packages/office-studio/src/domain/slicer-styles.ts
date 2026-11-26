/** Public style identifiers, independent of UI, Canvas and the document's color theme. */
export const SlicerStyleType = {
  UnselectedItemWithData: 'unselectedItemWithData',
  SelectedItemWithData: 'selectedItemWithData',
  UnselectedItemWithNoData: 'unselectedItemWithNoData',
  SelectedItemWithNoData: 'selectedItemWithNoData',
  HoveredUnselectedItemWithData: 'hoveredUnselectedItemWithData',
  HoveredSelectedItemWithData: 'hoveredSelectedItemWithData',
  HoveredUnselectedItemWithNoData: 'hoveredUnselectedItemWithNoData',
  HoveredSelectedItemWithNoData: 'hoveredSelectedItemWithNoData',
} as const
export type SlicerStyleType = typeof SlicerStyleType[keyof typeof SlicerStyleType]

export const TimelineStyleType = {
  SelectionLabel: 'selectionLabel', TimeLevel: 'timeLevel', PeriodLabel1: 'periodLabel1', PeriodLabel2: 'periodLabel2',
  SelectedTimeBlock: 'selectedTimeBlock', UnselectedTimeBlock: 'unselectedTimeBlock', SelectedTimeBlockSpace: 'selectedTimeBlockSpace',
} as const
export type TimelineStyleType = typeof TimelineStyleType[keyof typeof TimelineStyleType]

/** Numeric theme references are authored values. The Canvas theme adapter resolves them at render time. */
export interface SlicerStyleColor { rgb?: string | null; th?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 }
export interface SlicerStyleTextDecoration { s: 0 | 1; c?: 0 | 1; t?: number; cl?: SlicerStyleColor }
export interface SlicerStyleBorder { s: number; cl: SlicerStyleColor }
export type SlicerStyleBorderSide = 't' | 'r' | 'b' | 'l' | 'tl_br' | 'tl_bc' | 'tl_mr' | 'bl_tr' | 'ml_tr' | 'bc_tr'
export type SlicerStyleBorders = Partial<Record<SlicerStyleBorderSide, SlicerStyleBorder | null>>

/** Sparse formatting retains explicit null resets and all published border/outline/alignment fields. */
export interface SlicerStyleFormatting {
  bg?: SlicerStyleColor | null
  ol?: SlicerStyleTextDecoration
  bd?: SlicerStyleBorders | null
  cl?: SlicerStyleColor | null
  ht?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | null
  vt?: 0 | 1 | 2 | 3 | null
  bl?: 0 | 1
}
export interface ISlicerStyleElement { type: SlicerStyleType; formatting?: SlicerStyleFormatting }
export interface ITimelineStyleElement { type: TimelineStyleType; formatting?: SlicerStyleFormatting }
export interface ISlicerStyleDefinition {
  name: string
  wholeStyle?: SlicerStyleFormatting
  headerRowStyle?: SlicerStyleFormatting
  slicerStyleElements: ISlicerStyleElement[]
}
export interface ITimelineStyleDefinition {
  name: string
  wholeStyle?: SlicerStyleFormatting
  headerRowStyle?: SlicerStyleFormatting
  timelineStyleElements: ITimelineStyleElement[]
}
export interface ISlicerStylesDefinition { defaultSlicerStyle: string; styles: ISlicerStyleDefinition[] }
export interface ITimelineStylesDefinition { defaultTimelineStyle: string; styles: ITimelineStyleDefinition[] }
export interface ISlicerItemStyleState { selected: boolean; hasData: boolean; hovered: boolean }
