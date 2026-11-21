import { FOffice, OfficeStudio } from '@office-studio/office-studio'

export const studio = new OfficeStudio()
export const officeAPI = FOffice.newAPI(studio)

if (import.meta.hot) import.meta.hot.dispose(() => studio.dispose())

