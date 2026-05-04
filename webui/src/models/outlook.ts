export interface FolderDto {
  name: string
  folderPath: string
  itemCount: number
  subFolders: FolderDto[]
}

export interface MailItemDto {
  id: string
  subject: string
  senderName: string
  senderEmail: string
  receivedTime: string
  body: string
  bodyHtml: string
  folderPath: string
  categories: string
  isRead: boolean
  isMarkedAsTask: boolean
  flagRequest: string
  flagInterval: string
  taskStartDate?: string
  taskDueDate?: string
  taskCompletedDate?: string
  importance: string
  sensitivity: string
}

export interface OutlookCategoryDto {
  name: string
  color: string
  shortcutKey: string
}

export interface MailPropertiesCommandRequest {
  mailId: string
  folderPath: string
  isRead: boolean
  flagInterval: string
  flagRequest: string
  taskStartDate?: string
  taskDueDate?: string
  taskCompletedDate?: string
  categories: string[]
  newCategories: OutlookCategoryDto[]
}

export interface CategoryCommandRequest {
  name: string
  color: string
  shortcutKey: string
}

export interface OutlookRuleDto {
  name: string
  enabled: boolean
  executionOrder: number
  ruleType: string
  conditions: string[]
  actions: string[]
  exceptions: string[]
}

export interface CalendarEventDto {
  id: string
  subject: string
  start: string
  end: string
  location: string
  organizer: string
  requiredAttendees: string
  isRecurring: boolean
  busyStatus: string
}

export interface ChatMessageDto {
  id?: string
  source: string
  text: string
  timestamp: string
}

export interface AddinStatusDto {
  connected: boolean
  lastPollTime?: string
  lastPushTime?: string
  lastCommand: string
}

export interface AddinLogEntry {
  level: 'info' | 'warn' | 'error' | string
  message: string
  timestamp: string
}

export interface OutlookSignalRTestClientInfo {
  clientName: string
  workstation: string
  version: string
}

export interface OutlookSignalRTestConnectionEvent extends OutlookSignalRTestClientInfo {
  connectionId: string
  timestamp: string
}

export interface OutlookSignalRTestCommand {
  id: string
  type: string
  payload: string
  createdAt: string
}

export interface OutlookSignalRTestMessage {
  connectionId: string
  source: string
  level: string
  text: string
  timestamp: string
}

export interface OutlookSignalRTestResult {
  connectionId: string
  commandId: string
  success: boolean
  message: string
  payload: string
  timestamp: string
}

export type AppView = 'normal' | 'outlook' | 'calendar' | 'admin' | 'signalr-test' | 'swagger'
export type SignalRState = 'connected' | 'reconnecting' | 'disconnected'
