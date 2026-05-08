export type OutlookFolderType =
  | 'Unknown'
  | 'StoreRoot'
  | 'Mail'
  | 'Inbox'
  | 'Sent'
  | 'Drafts'
  | 'Deleted'
  | 'Junk'
  | 'Archive'
  | 'Outbox'
  | 'SyncIssues'
  | 'Conflicts'
  | 'LocalFailures'
  | 'ServerFailures'
  | 'Calendar'
  | 'Contacts'
  | 'Tasks'
  | 'Notes'
  | 'Journal'
  | 'RssFeeds'
  | 'ConversationHistory'
  | 'ConversationActionSettings'
  | 'OtherSystem'

export interface FolderDto {
  name: string
  entryId: string
  folderPath: string
  parentEntryId: string
  parentFolderPath: string
  itemCount: number
  storeId: string
  isStoreRoot: boolean
  folderType: OutlookFolderType
  defaultItemType: number
  isHidden: boolean
  isSystem: boolean
  hasChildren: boolean
  childrenLoaded: boolean
  discoveryState: 'partial' | 'loaded' | 'failed' | string
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

export interface OutlookRequestResponse<TData = Record<string, unknown>> {
  requestId: string
  request: string
  state: 'accepted' | 'running' | 'completed' | 'failed' | 'unavailable' | 'timeout' | string
  message?: string
  data: TData
}

export interface FetchResultRequest {
  requestId: string
  cursor?: string
  take?: number
}

export interface FetchResultNext {
  cursor: string
  hasMore: boolean
}

export interface FetchResultResponse<TData = Record<string, unknown>> {
  requestId: string
  request: string
  state: 'accepted' | 'running' | 'completed' | 'failed' | 'unavailable' | 'timeout' | string
  message: string
  next: FetchResultNext
  data: TData
}

export interface FolderDiscoveryRequest {
  syncId?: string
  storeId?: string
  parentEntryId?: string
  parentFolderPath?: string
  maxDepth?: number
  maxChildren?: number
  reset?: boolean
}

export interface SearchMailsRequest {
  searchId: string
  storeId: string
  scopeFolderPaths: string[]
  includeSubFolders: boolean
  keyword: string
  textFields: Array<'subject' | 'sender' | 'body'>
  categoryNames: string[]
  hasAttachments?: boolean
  flagState: 'any' | 'flagged' | 'unflagged'
  readState: 'any' | 'unread' | 'read'
  receivedFrom?: string
  receivedTo?: string
}

export interface FolderMailsRequest {
  folderPath: string
  includeSubFolders: boolean
  receivedFrom?: string
  receivedTo?: string
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

export interface OutlookRecipientDto {
  recipientKind: string
  displayName: string
  smtpAddress: string
  rawAddress: string
  addressType: string
  entryUserType: string
  isGroup: boolean
  isResolved: boolean
  members: OutlookRecipientDto[]
}

export interface MailItemDto {
  id: string
  subject: string
  sender: OutlookRecipientDto
  toRecipients: OutlookRecipientDto[]
  ccRecipients: OutlookRecipientDto[]
  bccRecipients: OutlookRecipientDto[]
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
  organizer: OutlookRecipientDto
  requiredAttendees: OutlookRecipientDto[]
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

export type AppView = 'outlook' | 'search' | 'chat' | 'calendar'
export type HubPage = 'outlook' | 'admin' | 'swagger'
export type SignalRState = 'connected' | 'reconnecting' | 'disconnected'
