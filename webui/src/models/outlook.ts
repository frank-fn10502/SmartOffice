export interface FolderDto {
  name: string
  folderPath: string
  parentFolderPath: string
  itemCount: number
  storeId: string
  isStoreRoot: boolean
}

export interface OutlookStoreDto {
  storeId: string
  displayName: string
  storeKind: string
  storeFilePath: string
  rootFolderPath: string
}

export interface FolderTreeNode extends FolderDto {
  subFolders: FolderTreeNode[]
}

export interface FolderSnapshotDto {
  stores: OutlookStoreDto[]
  folders: FolderDto[]
}

export interface FolderSyncBeginDto {
  syncId: string
  reset: boolean
  timestamp: string
}

export interface FolderSyncBatchDto {
  syncId: string
  sequence: number
  reset: boolean
  isFinal: boolean
  stores: OutlookStoreDto[]
  folders: FolderDto[]
}

export interface FolderSyncCompleteDto {
  syncId: string
  totalCount: number
  success: boolean
  message: string
  timestamp: string
}

export interface CommandDispatchResponse {
  commandId: string
  status: 'mocked' | 'dispatched' | 'addin_unavailable' | string
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

export type AppView = 'outlook' | 'chat' | 'calendar' | 'admin' | 'swagger'
export type SignalRState = 'connected' | 'reconnecting' | 'disconnected'
