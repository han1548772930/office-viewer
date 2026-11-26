/** Published owner values. The retained model separately uses year/quarter/month/day grains. */
export const TimelinePublicLevel = { Year: 0, Quarter: 1, Month: 2, Day: 3 } as const
export type TimelinePublicLevel = typeof TimelinePublicLevel[keyof typeof TimelinePublicLevel]

/** Excel serials in the host workbook date system; filtering excludes endDate. */
export interface ITimelineRange { startDate: number; endDate: number; useWholeDay?: boolean }
export const TimelineDateFilterOperator = {
  Between: 'between', NotBetween: 'notBetween', Equal: 'equal', NotEqual: 'notEqual',
  NewerThan: 'newerThan', NewerThanOrEqual: 'newerThanOrEqual', OlderThan: 'olderThan', OlderThanOrEqual: 'olderThanOrEqual',
} as const
export type TimelineDateFilterOperator = typeof TimelineDateFilterOperator[keyof typeof TimelineDateFilterOperator]
export interface ITimelineDateCondition {
  operator: TimelineDateFilterOperator
  expected: number | [number, number]
  useWholeDay?: boolean
}
export type TimelineDateFilterState = { range: ITimelineRange; condition?: never }
  | { range?: never; condition: ITimelineDateCondition } | { range?: never; condition?: never }
