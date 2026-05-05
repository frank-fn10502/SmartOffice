import type {
  AddinLogEntry,
  AddinStatusDto,
  AttachmentExportSettingsDto,
  CalendarEventDto,
  CategoryCommandRequest,
  ChatMessageDto,
  CommandDispatchResponse,
  FolderSnapshotDto,
  ExportedMailAttachmentDto,
  MailAttachmentDto,
  MailAttachmentsDto,
  MailBodyDto,
  MailSearchBatchDto,
  MailPropertiesCommandRequest,
  MailItemDto,
  OutlookCategoryDto,
  OutlookRuleDto,
  SearchMailsRequest,
} from '../models/outlook'
import { categoryColorValue, normalizeCategoryColor } from '../utils/categoryColors'

type LooseRecord = Record<string, unknown>

function readString(source: LooseRecord, camelName: string, pascalName: string, fallback = '') {
  return String(source[camelName] ?? source[pascalName] ?? fallback)
}

function readBoolean(source: LooseRecord, camelName: string, pascalName: string, fallback = false) {
  const value = source[camelName] ?? source[pascalName]
  return typeof value === 'boolean' ? value : fallback
}

function readDate(source: LooseRecord, camelName: string, pascalName: string) {
  const value = source[camelName] ?? source[pascalName]
  return typeof value === 'string' ? value : undefined
}

function readStringList(value: unknown) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item).trim()).filter(Boolean).join(', ')
  }
  return typeof value === 'string' ? value : ''
}

function readNumber(source: LooseRecord, camelName: string, pascalName: string, fallback = 0) {
  const value = source[camelName] ?? source[pascalName]
  return typeof value === 'number' ? value : typeof value === 'string' ? Number(value) || fallback : fallback
}

export function normalizeMailItem(item: unknown): MailItemDto {
  const source = (item ?? {}) as LooseRecord
  const flagInterval = readString(source, 'flagInterval', 'FlagInterval', 'none') || 'none'
  const flagRequest = readString(source, 'flagRequest', 'FlagRequest')
  const isMarkedAsTask = readBoolean(source, 'isMarkedAsTask', 'IsMarkedAsTask')
    || flagInterval !== 'none'
    || Boolean(flagRequest.trim())

  return {
    id: readString(source, 'id', 'Id'),
    subject: readString(source, 'subject', 'Subject'),
    senderName: readString(source, 'senderName', 'SenderName'),
    senderEmail: readString(source, 'senderEmail', 'SenderEmail'),
    receivedTime: readString(source, 'receivedTime', 'ReceivedTime'),
    body: readString(source, 'body', 'Body'),
    bodyHtml: readString(source, 'bodyHtml', 'BodyHtml'),
    folderPath: readString(source, 'folderPath', 'FolderPath'),
    categories: readStringList(source.categories ?? source.Categories),
    isRead: readBoolean(source, 'isRead', 'IsRead'),
    isMarkedAsTask,
    attachmentCount: readNumber(source, 'attachmentCount', 'AttachmentCount'),
    attachmentNames: readString(source, 'attachmentNames', 'AttachmentNames'),
    flagRequest,
    flagInterval,
    taskStartDate: readDate(source, 'taskStartDate', 'TaskStartDate'),
    taskDueDate: readDate(source, 'taskDueDate', 'TaskDueDate'),
    taskCompletedDate: readDate(source, 'taskCompletedDate', 'TaskCompletedDate'),
    importance: readString(source, 'importance', 'Importance', 'normal'),
    sensitivity: readString(source, 'sensitivity', 'Sensitivity', 'normal'),
  }
}

export function normalizeMailItems(items: unknown): MailItemDto[] {
  return Array.isArray(items) ? items.map(normalizeMailItem) : []
}

export function normalizeMailSearchBatch(item: unknown): MailSearchBatchDto {
  const source = (item ?? {}) as LooseRecord
  return {
    searchId: readString(source, 'searchId', 'SearchId'),
    sequence: readNumber(source, 'sequence', 'Sequence'),
    reset: readBoolean(source, 'reset', 'Reset'),
    isFinal: readBoolean(source, 'isFinal', 'IsFinal'),
    mails: normalizeMailItems(source.mails ?? source.Mails),
    message: readString(source, 'message', 'Message'),
  }
}

export function normalizeMailBody(item: unknown): MailBodyDto {
  const source = (item ?? {}) as LooseRecord
  return {
    mailId: readString(source, 'mailId', 'MailId'),
    folderPath: readString(source, 'folderPath', 'FolderPath'),
    body: readString(source, 'body', 'Body'),
    bodyHtml: readString(source, 'bodyHtml', 'BodyHtml'),
  }
}

export function normalizeMailAttachment(item: unknown): MailAttachmentDto {
  const source = (item ?? {}) as LooseRecord
  const attachmentId = readString(source, 'attachmentId', 'AttachmentId') || readString(source, 'id', 'Id') || readString(source, 'index', 'Index')
  const name = readString(source, 'name', 'Name') || readString(source, 'fileName', 'FileName') || readString(source, 'displayName', 'DisplayName')
  const exportedPath = readString(source, 'exportedPath', 'ExportedPath')
    || readString(source, 'localPath', 'LocalPath')
    || readString(source, 'fullPath', 'FullPath')
    || readString(source, 'path', 'Path')
  return {
    mailId: readString(source, 'mailId', 'MailId'),
    id: readString(source, 'id', 'Id') || attachmentId,
    attachmentId,
    index: readNumber(source, 'index', 'Index'),
    fileName: readString(source, 'fileName', 'FileName') || name,
    displayName: readString(source, 'displayName', 'DisplayName') || name,
    name,
    contentType: readString(source, 'contentType', 'ContentType'),
    size: readNumber(source, 'size', 'Size'),
    isExported: readBoolean(source, 'isExported', 'IsExported'),
    exportedAttachmentId: readString(source, 'exportedAttachmentId', 'ExportedAttachmentId'),
    path: readString(source, 'path', 'Path') || exportedPath,
    localPath: readString(source, 'localPath', 'LocalPath') || exportedPath,
    fullPath: readString(source, 'fullPath', 'FullPath') || exportedPath,
    exportedPath,
  }
}

export function normalizeMailAttachments(item: unknown): MailAttachmentsDto {
  const source = (item ?? {}) as LooseRecord
  const attachments = source.attachments ?? source.Attachments
  return {
    mailId: readString(source, 'mailId', 'MailId'),
    folderPath: readString(source, 'folderPath', 'FolderPath'),
    attachments: Array.isArray(attachments) ? attachments.map(normalizeMailAttachment) : [],
  }
}

export function normalizeExportedMailAttachment(item: unknown): ExportedMailAttachmentDto {
  const source = (item ?? {}) as LooseRecord
  const attachmentId = readString(source, 'attachmentId', 'AttachmentId') || readString(source, 'id', 'Id') || readString(source, 'index', 'Index')
  const name = readString(source, 'name', 'Name') || readString(source, 'fileName', 'FileName') || readString(source, 'displayName', 'DisplayName')
  const exportedPath = readString(source, 'exportedPath', 'ExportedPath')
    || readString(source, 'localPath', 'LocalPath')
    || readString(source, 'fullPath', 'FullPath')
    || readString(source, 'path', 'Path')
  return {
    mailId: readString(source, 'mailId', 'MailId'),
    folderPath: readString(source, 'folderPath', 'FolderPath'),
    id: readString(source, 'id', 'Id') || attachmentId,
    attachmentId,
    index: readNumber(source, 'index', 'Index'),
    exportedAttachmentId: readString(source, 'exportedAttachmentId', 'ExportedAttachmentId'),
    fileName: readString(source, 'fileName', 'FileName') || name,
    displayName: readString(source, 'displayName', 'DisplayName') || name,
    name,
    contentType: readString(source, 'contentType', 'ContentType'),
    size: readNumber(source, 'size', 'Size'),
    path: readString(source, 'path', 'Path') || exportedPath,
    localPath: readString(source, 'localPath', 'LocalPath') || exportedPath,
    fullPath: readString(source, 'fullPath', 'FullPath') || exportedPath,
    exportedPath,
    exportedAt: readString(source, 'exportedAt', 'ExportedAt'),
  }
}

export function normalizeOutlookCategory(item: unknown): OutlookCategoryDto {
  const source = (item ?? {}) as LooseRecord
  const color = normalizeCategoryColor(readString(source, 'color', 'Color'))
  const rawColorValue = source.colorValue ?? source.ColorValue
  const colorValue = typeof rawColorValue === 'number'
    ? rawColorValue
    : typeof rawColorValue === 'string' && rawColorValue.trim()
      ? Number(rawColorValue)
      : categoryColorValue(color)
  return {
    name: readString(source, 'name', 'Name'),
    color,
    colorValue: Number.isFinite(colorValue) ? colorValue : categoryColorValue(color),
    shortcutKey: readString(source, 'shortcutKey', 'ShortcutKey'),
  }
}

export function normalizeOutlookCategories(items: unknown): OutlookCategoryDto[] {
  return Array.isArray(items)
    ? items.map(normalizeOutlookCategory).filter((category) => category.name.trim())
    : []
}

async function getJson<T>(url: string): Promise<T> {
  const response = await fetch(url)
  if (!response.ok) throw new Error(`Request failed: ${response.status}`)
  return response.json() as Promise<T>
}

async function postJson<T>(url: string, body?: unknown): Promise<T> {
  const response = await fetch(url, {
    method: 'POST',
    headers: body ? { 'Content-Type': 'application/json' } : undefined,
    body: body ? JSON.stringify(body) : undefined,
  })
  if (!response.ok) throw new Error(`Request failed: ${response.status}`)
  return response.json() as Promise<T>
}

export const outlookApi = {
  getFolders: () => getJson<FolderSnapshotDto>('/api/outlook/folders'),
  getMails: async () => normalizeMailItems(await getJson<unknown>('/api/outlook/mails')),
  getMailSearchResults: async () => normalizeMailItems(await getJson<unknown>('/api/outlook/mail-search')),
  getRules: () => getJson<OutlookRuleDto[]>('/api/outlook/rules'),
  getCategories: async () => normalizeOutlookCategories(await getJson<unknown>('/api/outlook/categories')),
  getCalendar: () => getJson<CalendarEventDto[]>('/api/outlook/calendar'),
  getChat: () => getJson<ChatMessageDto[]>('/api/outlook/chat'),
  getAdminStatus: () => getJson<AddinStatusDto>('/api/outlook/admin/status'),
  getAdminLogs: () => getJson<AddinLogEntry[]>('/api/outlook/admin/logs'),
  getAttachmentExportSettings: () => getJson<AttachmentExportSettingsDto>('/api/outlook/attachment-export-settings'),

  requestFolders: () => postJson<CommandDispatchResponse>('/api/outlook/request-folders'),
  requestMails: (body: { folderPath: string; range: string; maxCount: number }) =>
    postJson<CommandDispatchResponse>('/api/outlook/request-mails', body),
  requestMailSearch: (body: SearchMailsRequest) =>
    postJson<CommandDispatchResponse>('/api/outlook/request-mail-search', body),
  requestMailBody: (body: { mailId: string; folderPath: string }) =>
    postJson<CommandDispatchResponse>('/api/outlook/request-mail-body', body),
  requestMailAttachments: (body: { mailId: string; folderPath: string }) =>
    postJson<CommandDispatchResponse>('/api/outlook/request-mail-attachments', body),
  requestExportMailAttachment: (body: {
    mailId: string
    folderPath: string
    attachmentId: string
    index: number
    name: string
    fileName: string
    displayName: string
  }) =>
    postJson<CommandDispatchResponse>('/api/outlook/request-export-mail-attachment', body),
  openExportedAttachment: (body: { exportedAttachmentId: string }) =>
    postJson('/api/outlook/open-exported-attachment', body),
  updateAttachmentExportSettings: (body: { rootPath: string }) =>
    postJson<AttachmentExportSettingsDto>('/api/outlook/attachment-export-settings', body),
  requestRules: () => postJson('/api/outlook/request-rules'),
  requestCategories: () => postJson<CommandDispatchResponse>('/api/outlook/request-categories'),
  requestSignalRPing: () => postJson('/api/outlook/request-signalr-ping'),
  requestCalendar: (body: { daysForward: number; startDate?: string; endDate?: string }) =>
    postJson('/api/outlook/request-calendar', body),
  sendChat: (body: { source: 'web'; text: string }) => postJson('/api/outlook/chat', body),

  requestMarkMailRead: (body: { mailId: string; folderPath: string }) =>
    postJson('/api/outlook/request-mark-mail-read', body),
  requestMarkMailUnread: (body: { mailId: string; folderPath: string }) =>
    postJson('/api/outlook/request-mark-mail-unread', body),
  requestMarkMailTask: (body: { mailId: string; folderPath: string }) =>
    postJson('/api/outlook/request-mark-mail-task', body),
  requestClearMailTask: (body: { mailId: string; folderPath: string }) =>
    postJson('/api/outlook/request-clear-mail-task', body),
  requestSetMailCategories: (body: { mailId: string; folderPath: string; categories: string }) =>
    postJson('/api/outlook/request-set-mail-categories', body),
  requestUpdateMailProperties: (body: MailPropertiesCommandRequest) =>
    postJson('/api/outlook/request-update-mail-properties', body),
  requestUpsertCategory: (body: CategoryCommandRequest) =>
    postJson('/api/outlook/request-upsert-category', body),
  requestCreateFolder: (body: { parentFolderPath: string; name: string }) =>
    postJson('/api/outlook/request-create-folder', body),
  requestDeleteFolder: (body: { folderPath: string }) => postJson('/api/outlook/request-delete-folder', body),
  requestMoveMail: (body: { mailId: string; sourceFolderPath: string; destinationFolderPath: string }) =>
    postJson('/api/outlook/request-move-mail', body),
  requestDeleteMail: (body: { mailId: string; folderPath: string }) =>
    postJson('/api/outlook/request-delete-mail', body),
}
