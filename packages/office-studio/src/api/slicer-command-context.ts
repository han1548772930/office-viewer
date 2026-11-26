/** Native Facade calls can address inactive units; captured workbench input requires its active host. */
export interface SlicerCommandContext {
  expectedRevision?: number
  host?: 'workbench'
}
