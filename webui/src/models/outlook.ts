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

export interface SearchMailsRequest {
  searchId: string
  storeId: string
  scopeFolderPaths: string[]
  includeSubFolders: boolean
  keyword: string
  matchMode: 'contains' | 'exact' | 'fuzzy' | 'regex'
  fields: string[]
  receivedFrom?: string
  receivedTo?: string
  maxCount: number
}

export interface MailSearchSliceResultDto {
  searchId: string
  commandId: string
  parentCommandId: string
  sequence: number
  sliceIndex: number
  sliceCount: number
  reset: boolean
  isFinal: boolean
  isSliceComplete: boolean
  mails: MailItemDto[]
  message: string
}

export interface MailSearchCompleteDto {
  searchId: string
  commandId: string
  parentCommandId: string
  totalCount: number
  success: boolean
  message: string
  timestamp: string
}

export interface MailSearchProgressDto {
  searchId: string
  commandId: string
  status: string
  phase: string
  processedStores: number
  totalStores: number
  processedFolders: number
  totalFolders: number
  resultCount: number
  currentStoreId: string
  currentFolderPath: string
  message: string
  timestamp: string
  percent: number
}

export interface OutlookCommandResult {
  commandId: string
  success: boolean
  message: string
  payload: string
  timestamp: string
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
  attachmentCount: number
  attachmentNames: string
  flagRequest: string
  flagInterval: string
  taskStartDate?: string
  taskDueDate?: string
  taskCompletedDate?: string
  importance: string
  sensitivity: string
}

export interface MailBodyDto {
  mailId: string
  folderPath: string
  body: string
  bodyHtml: string
}

export interface MailAttachmentDto {
  mailId: string
  id: string
  attachmentId: string
  index: number
  fileName: string
  displayName: string
  name: string
  contentType: string
  size: number
  isExported: boolean
  exportedAttachmentId: string
  path: string
  localPath: string
  fullPath: string
  exportedPath: string
}

export interface MailAttachmentsDto {
  mailId: string
  folderPath: string
  attachments: MailAttachmentDto[]
}

export interface ExportedMailAttachmentDto {
  mailId: string
  folderPath: string
  id: string
  attachmentId: string
  index: number
  exportedAttachmentId: string
  fileName: string
  displayName: string
  name: string
  contentType: string
  size: number
  path: string
  localPath: string
  fullPath: string
  exportedPath: string
  exportedAt: string
}

export interface AttachmentExportSettingsDto {
  rootPath: string
  defaultRootPath: string
}

export interface OutlookCategoryDto {
  name: string
  color: string
  colorValue: number
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

export interface MailPropertiesDraft {
  isRead: boolean
  flagInterval: string
  flagRequest: string
  taskStartDate: string
  taskDueDate: string
  taskCompletedDate: string
  categories: string[]
}

export interface CategoryCommandRequest {
  name: string
  color: string
  colorValue: number
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

export type AppView = 'outlook' | 'search' | 'chat' | 'calendar' | 'admin' | 'swagger'
export type SignalRState = 'connected' | 'reconnecting' | 'disconnected'
