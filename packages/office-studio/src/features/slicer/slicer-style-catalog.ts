import type { SlicerType, WorkbookModel } from '../../domain/model.ts'
import { SlicerStyleType, TimelineStyleType, type ISlicerItemStyleState, type ISlicerStyleDefinition,
  type ITimelineStyleDefinition, type SlicerStyleFormatting } from '../../domain/slicer-styles.ts'

export const BUILT_IN_SLICER_STYLE_NAMES = ['SlicerStyleLight1', 'SlicerStyleLight2', 'SlicerStyleLight3', 'SlicerStyleLight4',
  'SlicerStyleLight5', 'SlicerStyleLight6', 'SlicerStyleOther1', 'SlicerStyleOther2', 'SlicerStyleDark1', 'SlicerStyleDark2',
  'SlicerStyleDark3', 'SlicerStyleDark4', 'SlicerStyleDark5', 'SlicerStyleDark6'] as const
export const BUILT_IN_TIMELINE_STYLE_NAMES = ['TimelineStyleLight1', 'TimelineStyleLight2', 'TimelineStyleLight3', 'TimelineStyleLight4',
  'TimelineStyleLight5', 'TimelineStyleLight6', 'TimelineStyleDark1', 'TimelineStyleDark2', 'TimelineStyleDark3',
  'TimelineStyleDark4', 'TimelineStyleDark5', 'TimelineStyleDark6'] as const
export type BuiltInSlicerStyleName = typeof BUILT_IN_SLICER_STYLE_NAMES[number]
export type BuiltInTimelineStyleName = typeof BUILT_IN_TIMELINE_STYLE_NAMES[number]

/** Kept for existing UI consumers; document-aware galleries use listSlicerStyleNames. */
export const SLICER_STYLE_NAMES: Readonly<Record<SlicerType, readonly string[]>> = {
  slicer: BUILT_IN_SLICER_STYLE_NAMES, timeline: BUILT_IN_TIMELINE_STYLE_NAMES,
}
export const SLICER_ITEM_STYLE_TYPES = Object.values(SlicerStyleType)
export const TIMELINE_ELEMENT_STYLE_TYPES = Object.values(TimelineStyleType)
export type SlicerStyleOwner = Pick<WorkbookModel, 'slicerStyles' | 'timelineStyles'>

export function isBuiltInSlicerStyleName(value: string): value is BuiltInSlicerStyleName {
  return (BUILT_IN_SLICER_STYLE_NAMES as readonly string[]).includes(value)
}
export function isBuiltInTimelineStyleName(value: string): value is BuiltInTimelineStyleName {
  return (BUILT_IN_TIMELINE_STYLE_NAMES as readonly string[]).includes(value)
}
export function listSlicerStyleNames(workbook: SlicerStyleOwner, type: SlicerType = 'slicer'): string[] {
  const styles = type === 'slicer' ? workbook.slicerStyles?.styles : workbook.timelineStyles?.styles
  return [...new Set([...SLICER_STYLE_NAMES[type], ...(styles ?? []).map(style => style.name)])]
}
export function listTimelineStyleNames(workbook: SlicerStyleOwner): string[] { return listSlicerStyleNames(workbook, 'timeline') }
export function hasSlicerStyle(workbook: SlicerStyleOwner, type: SlicerType, name: string): boolean {
  return listSlicerStyleNames(workbook, type).includes(name)
}
export function hasTimelineStyle(workbook: SlicerStyleOwner, name: string): boolean { return hasSlicerStyle(workbook, 'timeline', name) }
export function getDefaultSlicerStyleName(workbook: SlicerStyleOwner, type: SlicerType = 'slicer'): string {
  const authored = type === 'slicer' ? workbook.slicerStyles?.defaultSlicerStyle : workbook.timelineStyles?.defaultTimelineStyle
  return authored !== undefined && hasSlicerStyle(workbook, type, authored) ? authored : SLICER_STYLE_NAMES[type][0]!
}

interface StyleReference { name: string; builtIn: boolean; usesDefault: boolean; unresolvedName?: string }
export type SlicerStyleDescriptor = StyleReference & ({ type: 'slicer'; definition?: ISlicerStyleDefinition }
  | { type: 'timeline'; definition?: ITimelineStyleDefinition })
export function getSlicerStyle(workbook: SlicerStyleOwner, type: 'slicer', name?: string): Extract<SlicerStyleDescriptor, { type: 'slicer' }>
export function getSlicerStyle(workbook: SlicerStyleOwner, type: 'timeline', name?: string): Extract<SlicerStyleDescriptor, { type: 'timeline' }>
export function getSlicerStyle(workbook: SlicerStyleOwner, type: SlicerType, name?: string): SlicerStyleDescriptor
export function getSlicerStyle(workbook: SlicerStyleOwner, type: SlicerType, name?: string): SlicerStyleDescriptor {
  const known = name !== undefined && hasSlicerStyle(workbook, type, name)
  const resolved = name !== undefined && known ? name : getDefaultSlicerStyleName(workbook, type)
  const reference: StyleReference = { name: resolved, builtIn: SLICER_STYLE_NAMES[type].includes(resolved), usesDefault: !known,
    ...(name !== undefined && !known ? { unresolvedName: name } : {}) }
  if (type === 'slicer') {
    const definition = workbook.slicerStyles?.styles.find(style => style.name === resolved)
    return { ...reference, type, ...(definition ? { definition: structuredClone(definition) } : {}) }
  }
  const definition = workbook.timelineStyles?.styles.find(style => style.name === resolved)
  return { ...reference, type, ...(definition ? { definition: structuredClone(definition) } : {}) }
}
export function getTimelineStyle(workbook: SlicerStyleOwner, name?: string): Extract<SlicerStyleDescriptor, { type: 'timeline' }> {
  return getSlicerStyle(workbook, 'timeline', name)
}

/** Selects one of the eight published states without changing member selection or availability. */
export function getSlicerStyleType(state: ISlicerItemStyleState): SlicerStyleType {
  return state.hovered ? state.selected ? state.hasData ? SlicerStyleType.HoveredSelectedItemWithData : SlicerStyleType.HoveredSelectedItemWithNoData
    : state.hasData ? SlicerStyleType.HoveredUnselectedItemWithData : SlicerStyleType.HoveredUnselectedItemWithNoData
    : state.selected ? state.hasData ? SlicerStyleType.SelectedItemWithData : SlicerStyleType.SelectedItemWithNoData
      : state.hasData ? SlicerStyleType.UnselectedItemWithData : SlicerStyleType.UnselectedItemWithNoData
}

export interface SlicerStyleFormattingProjection {
  reference: Extract<SlicerStyleDescriptor, { type: 'slicer' }>
  wholeStyle: SlicerStyleFormatting
  headerRowStyle: SlicerStyleFormatting
  items: Record<SlicerStyleType, SlicerStyleFormatting>
}
export interface TimelineStyleFormattingProjection {
  reference: Extract<SlicerStyleDescriptor, { type: 'timeline' }>
  wholeStyle: SlicerStyleFormatting
  headerRowStyle: SlicerStyleFormatting
  elements: Record<TimelineStyleType, SlicerStyleFormatting>
}
/** Author formatting remains sparse. The Canvas adapter supplies theme/palette defaults before applying each region. */
export function projectSlicerStyleFormatting(workbook: SlicerStyleOwner, name?: string): SlicerStyleFormattingProjection {
  const reference = getSlicerStyle(workbook, 'slicer', name), definition = reference.definition
  const items = Object.fromEntries(SLICER_ITEM_STYLE_TYPES.map(type => [type,
    structuredClone(definition?.slicerStyleElements.find(element => element.type === type)?.formatting ?? {})])) as Record<SlicerStyleType, SlicerStyleFormatting>
  return { reference, wholeStyle: structuredClone(definition?.wholeStyle ?? {}), headerRowStyle: structuredClone(definition?.headerRowStyle ?? {}), items }
}
export function projectTimelineStyleFormatting(workbook: SlicerStyleOwner, name?: string): TimelineStyleFormattingProjection {
  const reference = getSlicerStyle(workbook, 'timeline', name), definition = reference.definition
  const elements = Object.fromEntries(TIMELINE_ELEMENT_STYLE_TYPES.map(type => [type,
    structuredClone(definition?.timelineStyleElements.find(element => element.type === type)?.formatting ?? {})])) as Record<TimelineStyleType, SlicerStyleFormatting>
  return { reference, wholeStyle: structuredClone(definition?.wholeStyle ?? {}), headerRowStyle: structuredClone(definition?.headerRowStyle ?? {}), elements }
}
export function projectSlicerItemStyle(workbook: SlicerStyleOwner, name: string | undefined, state: ISlicerItemStyleState):
  { type: SlicerStyleType; formatting: SlicerStyleFormatting } {
  const type = getSlicerStyleType(state)
  return { type, formatting: projectSlicerStyleFormatting(workbook, name).items[type] }
}
