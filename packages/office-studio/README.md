# Office Studio

Vue components and a typed workbook/document facade for embedding the Office Studio editor.

```ts
import { FOffice, OfficeStudio } from '@office-studio/office-studio'

const studio = new OfficeStudio()
const officeAPI = FOffice.newAPI(studio)
const worksheet = officeAPI.getActiveWorkbook()?.getActiveSheet()
worksheet?.getRange('A1:B2').setValues([
  ['Name', 'Amount'],
  ['Apples', 12],
])
```

Install the stylesheet entry alongside the app's Vue and Pinia setup:

```ts
import '@office-studio/office-studio/style.css'
```

`OfficeStudio` owns a Pinia instance. Install `studio.pinia` on the embedding Vue app before rendering `DocumentWorkbench` or `ExcelWorkbench`. The `examples/workbench` app demonstrates this integration.
