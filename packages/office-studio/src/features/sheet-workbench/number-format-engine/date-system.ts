/** Matches Office serial-date systems at the pure formatting boundary. */
export const DateSystem = { Date1900: 'date1900', Date1904: 'date1904' } as const
export type DateSystem = typeof DateSystem[keyof typeof DateSystem]
