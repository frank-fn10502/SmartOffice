export type HubPage = 'outlook' | 'swagger'

export type WorkspaceNavValue = string | number | boolean

export type WorkspaceNavOption<TValue extends WorkspaceNavValue = string> = {
  label: string
  value: TValue
  disabled?: boolean
}
