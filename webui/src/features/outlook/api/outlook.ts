import type {
  AddinLogEntry,
  AddinStatusDto,
  AddressBookContactDto,
  AddressBookGroupMembersResponse,
  AddressBookLookupResponse,
  AddressBookMergeSuggestionResponse,
  AddressBookResponse,
  AttachmentExportSettingsDto,
  CalendarEventDto,
  CalendarRoomDto,
  CategoryCommandRequest,
  ChatMessageDto,
  FolderSnapshotDto,
  FolderMailsRequest,
  ExportedMailAttachmentDto,
  MailAttachmentDto,
  MailAttachmentsDto,
  MailBodyDto,
  MailConversationDto,
  MailSearchProgressDto,
  MailPropertiesCommandRequest,
  MailItemDto,
  OutlookRecipientDto,
  OutlookCategoryDto,
  OutlookRuleDto,
  OutlookRuleCommandRequest,
  FetchResultRequest,
  FetchResultResponse,
  OutlookRequestResponse,
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

function readStringArray(value: unknown) {
  return Array.isArray(value) ? value.map((item) => String(item).trim()).filter(Boolean) : []
}

function readNumber(source: LooseRecord, camelName: string, pascalName: string, fallback = 0) {
  const value = source[camelName] ?? source[pascalName]
  return typeof value === 'number' ? value : typeof value === 'string' ? Number(value) || fallback : fallback
}

function normalizeRecipient(item: unknown, kind = ''): OutlookRecipientDto {
  const source = (item ?? {}) as LooseRecord
  const members = source.members ?? source.Members
  return {
    recipientKind: readString(source, 'recipientKind', 'RecipientKind', kind),
    displayName: readString(source, 'displayName', 'DisplayName'),
    smtpAddress: readString(source, 'smtpAddress', 'SmtpAddress'),
    rawAddress: readString(source, 'rawAddress', 'RawAddress'),
    addressType: readString(source, 'addressType', 'AddressType'),
    entryUserType: readString(source, 'entryUserType', 'EntryUserType'),
    isGroup: readBoolean(source, 'isGroup', 'IsGroup'),
    isResolved: readBoolean(source, 'isResolved', 'IsResolved'),
    members: Array.isArray(members) ? members.map((member) => normalizeRecipient(member, 'member')) : [],
  }
}

function normalizeRecipients(items: unknown, kind: string): OutlookRecipientDto[] {
  return Array.isArray(items) ? items.map((item) => normalizeRecipient(item, kind)) : []
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
    sender: normalizeRecipient(source.sender ?? source.Sender, 'sender'),
    toRecipients: normalizeRecipients(source.toRecipients ?? source.ToRecipients, 'to'),
    ccRecipients: normalizeRecipients(source.ccRecipients ?? source.CcRecipients, 'cc'),
    bccRecipients: normalizeRecipients(source.bccRecipients ?? source.BccRecipients, 'bcc'),
    receivedTime: readString(source, 'receivedTime', 'ReceivedTime'),
    body: readString(source, 'body', 'Body'),
    bodyHtml: readString(source, 'bodyHtml', 'BodyHtml'),
    folderPath: readString(source, 'folderPath', 'FolderPath'),
    messageClass: readString(source, 'messageClass', 'MessageClass'),
    conversationId: readString(source, 'conversationId', 'ConversationId'),
    conversationTopic: readString(source, 'conversationTopic', 'ConversationTopic'),
    conversationIndex: readString(source, 'conversationIndex', 'ConversationIndex'),
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

export function normalizeMailSearchProgress(item: unknown): MailSearchProgressDto {
  const source = (item ?? {}) as LooseRecord
  return {
    searchId: readString(source, 'searchId', 'SearchId'),
    commandId: readString(source, 'commandId', 'CommandId'),
    status: readString(source, 'status', 'Status'),
    phase: readString(source, 'phase', 'Phase'),
    processedStores: readNumber(source, 'processedStores', 'ProcessedStores'),
    totalStores: readNumber(source, 'totalStores', 'TotalStores'),
    processedFolders: readNumber(source, 'processedFolders', 'ProcessedFolders'),
    totalFolders: readNumber(source, 'totalFolders', 'TotalFolders'),
    resultCount: readNumber(source, 'resultCount', 'ResultCount'),
    currentStoreId: readString(source, 'currentStoreId', 'CurrentStoreId'),
    currentFolderPath: readString(source, 'currentFolderPath', 'CurrentFolderPath'),
    message: readString(source, 'message', 'Message'),
    timestamp: readString(source, 'timestamp', 'Timestamp'),
    percent: readNumber(source, 'percent', 'Percent'),
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

export function normalizeMailConversation(item: unknown): MailConversationDto {
  const source = (item ?? {}) as LooseRecord
  return {
    mailId: readString(source, 'mailId', 'MailId'),
    folderPath: readString(source, 'folderPath', 'FolderPath'),
    conversationId: readString(source, 'conversationId', 'ConversationId'),
    conversationTopic: readString(source, 'conversationTopic', 'ConversationTopic'),
    mails: normalizeMailItems(source.mails ?? source.Mails),
  }
}

export function normalizeCalendarEvent(item: unknown): CalendarEventDto {
  const source = (item ?? {}) as LooseRecord
  return {
    id: readString(source, 'id', 'Id'),
    subject: readString(source, 'subject', 'Subject'),
    start: readString(source, 'start', 'Start'),
    end: readString(source, 'end', 'End'),
    location: readString(source, 'location', 'Location'),
    organizer: normalizeRecipient(source.organizer ?? source.Organizer, 'organizer'),
    requiredAttendees: normalizeRecipients(source.requiredAttendees ?? source.RequiredAttendees, 'required'),
    isRecurring: readBoolean(source, 'isRecurring', 'IsRecurring'),
    busyStatus: readString(source, 'busyStatus', 'BusyStatus'),
    smartOfficeOwned: readBoolean(source, 'smartOfficeOwned', 'SmartOfficeOwned'),
    smartOfficeEventId: readString(source, 'smartOfficeEventId', 'SmartOfficeEventId'),
  }
}

export function normalizeCalendarEvents(items: unknown): CalendarEventDto[] {
  return Array.isArray(items) ? items.map(normalizeCalendarEvent) : []
}

export function normalizeCalendarRoom(item: unknown): CalendarRoomDto {
  const source = (item ?? {}) as LooseRecord
  return {
    id: readString(source, 'id', 'Id'),
    displayName: readString(source, 'displayName', 'DisplayName'),
    smtpAddress: readString(source, 'smtpAddress', 'SmtpAddress'),
    rawAddress: readString(source, 'rawAddress', 'RawAddress'),
    source: readString(source, 'source', 'Source'),
  }
}

export function normalizeCalendarRooms(items: unknown): CalendarRoomDto[] {
  return Array.isArray(items) ? items.map(normalizeCalendarRoom).filter((room) => room.displayName.trim()) : []
}

export function normalizeAddressBookContact(item: unknown): AddressBookContactDto {
  const source = (item ?? {}) as LooseRecord
  return {
    id: readString(source, 'id', 'Id'),
    displayName: readString(source, 'displayName', 'DisplayName'),
    smtpAddress: readString(source, 'smtpAddress', 'SmtpAddress'),
    rawAddress: readString(source, 'rawAddress', 'RawAddress'),
    addressType: readString(source, 'addressType', 'AddressType'),
    entryUserType: readString(source, 'entryUserType', 'EntryUserType'),
    source: readString(source, 'source', 'Source'),
    companyName: readString(source, 'companyName', 'CompanyName'),
    jobTitle: readString(source, 'jobTitle', 'JobTitle'),
    department: readString(source, 'department', 'Department'),
    officeLocation: readString(source, 'officeLocation', 'OfficeLocation'),
    businessTelephoneNumber: readString(source, 'businessTelephoneNumber', 'BusinessTelephoneNumber'),
    mobileTelephoneNumber: readString(source, 'mobileTelephoneNumber', 'MobileTelephoneNumber'),
    domain: readString(source, 'domain', 'Domain'),
    isKnown: readBoolean(source, 'isKnown', 'IsKnown'),
    isLikelySelf: readBoolean(source, 'isLikelySelf', 'IsLikelySelf'),
    isGroup: readBoolean(source, 'isGroup', 'IsGroup'),
    memberCount: readNumber(source, 'memberCount', 'MemberCount'),
    groupMembersLoaded: readBoolean(source, 'groupMembersLoaded', 'GroupMembersLoaded'),
    groupMembersLoading: readBoolean(source, 'groupMembersLoading', 'GroupMembersLoading'),
    groupMembersRequestId: readString(source, 'groupMembersRequestId', 'GroupMembersRequestId'),
    groupMembersUpdatedAt: readDate(source, 'groupMembersUpdatedAt', 'GroupMembersUpdatedAt'),
    relationScore: readNumber(source, 'relationScore', 'RelationScore'),
    mailCount: readNumber(source, 'mailCount', 'MailCount'),
    calendarCount: readNumber(source, 'calendarCount', 'CalendarCount'),
    senderCount: readNumber(source, 'senderCount', 'SenderCount'),
    recipientCount: readNumber(source, 'recipientCount', 'RecipientCount'),
    organizerCount: readNumber(source, 'organizerCount', 'OrganizerCount'),
    attendeeCount: readNumber(source, 'attendeeCount', 'AttendeeCount'),
    groupMemberCount: readNumber(source, 'groupMemberCount', 'GroupMemberCount'),
    firstSeen: readDate(source, 'firstSeen', 'FirstSeen'),
    lastSeen: readDate(source, 'lastSeen', 'LastSeen'),
    relationKinds: readStringArray(source.relationKinds ?? source.RelationKinds),
    sources: readStringArray(source.sources ?? source.Sources),
    memberSmtpAddresses: readStringArray(source.memberSmtpAddresses ?? source.MemberSmtpAddresses),
    memberGroupSmtpAddresses: readStringArray(source.memberGroupSmtpAddresses ?? source.MemberGroupSmtpAddresses),
    memberOfGroupSmtpAddresses: readStringArray(source.memberOfGroupSmtpAddresses ?? source.MemberOfGroupSmtpAddresses),
    folderPaths: readStringArray(source.folderPaths ?? source.FolderPaths),
    recentMailIds: readStringArray(source.recentMailIds ?? source.RecentMailIds),
    sampleSubjects: readStringArray(source.sampleSubjects ?? source.SampleSubjects),
  }
}

function normalizeAddressBookMergeSuggestionResponse(item: unknown): AddressBookMergeSuggestionResponse {
  const source = (item ?? {}) as LooseRecord
  const suggestions = source.suggestions ?? source.Suggestions
  return {
    state: readString(source, 'state', 'State'),
    suggestions: Array.isArray(suggestions)
      ? suggestions.map((suggestion) => {
        const suggestionSource = (suggestion ?? {}) as LooseRecord
        const coveredContacts = suggestionSource.coveredContacts ?? suggestionSource.CoveredContacts
        return {
          groupSmtpAddress: readString(suggestionSource, 'groupSmtpAddress', 'GroupSmtpAddress'),
          groupDisplayName: readString(suggestionSource, 'groupDisplayName', 'GroupDisplayName'),
          coveredContacts: Array.isArray(coveredContacts) ? coveredContacts.map(normalizeAddressBookContact) : [],
          coveredRecipientKeys: readStringArray(suggestionSource.coveredRecipientKeys ?? suggestionSource.CoveredRecipientKeys),
          message: readString(suggestionSource, 'message', 'Message'),
        }
      })
      : [],
  }
}

function normalizeAddressBookResponse(item: unknown): AddressBookResponse {
  const source = (item ?? {}) as LooseRecord
  const contacts = source.contacts ?? source.Contacts
  return {
    query: readString(source, 'query', 'Query'),
    totalCount: readNumber(source, 'totalCount', 'TotalCount'),
    contacts: Array.isArray(contacts) ? contacts.map(normalizeAddressBookContact) : [],
  }
}

function normalizeAddressBookLookupResponse(item: unknown): AddressBookLookupResponse {
  const source = (item ?? {}) as LooseRecord
  const suggestions = source.suggestions ?? source.Suggestions
  const contact = source.contact ?? source.Contact
  return {
    query: readString(source, 'query', 'Query'),
    state: readString(source, 'state', 'State'),
    message: readString(source, 'message', 'Message'),
    contact: contact ? normalizeAddressBookContact(contact) : null,
    suggestions: Array.isArray(suggestions) ? suggestions.map(normalizeAddressBookContact) : [],
  }
}

function normalizeAddressBookGroupMembersResponse(item: unknown): AddressBookGroupMembersResponse {
  const source = (item ?? {}) as LooseRecord
  const members = source.members ?? source.Members
  return {
    state: readString(source, 'state', 'State'),
    message: readString(source, 'message', 'Message'),
    groupKey: readString(source, 'groupKey', 'GroupKey'),
    groupSmtpAddress: readString(source, 'groupSmtpAddress', 'GroupSmtpAddress'),
    requestId: readString(source, 'requestId', 'RequestId'),
    totalCount: readNumber(source, 'totalCount', 'TotalCount'),
    updatedAt: readDate(source, 'updatedAt', 'UpdatedAt'),
    members: Array.isArray(members) ? members.map(normalizeAddressBookContact) : [],
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

export function normalizeOutlookRule(item: unknown): OutlookRuleDto {
  const source = (item ?? {}) as LooseRecord
  const rawRuleType = readString(source, 'ruleType', 'RuleType', 'receive').toLowerCase()
  const ruleType = rawRuleType.includes('send') ? 'send' : 'receive'
  return {
    storeId: readString(source, 'storeId', 'StoreId'),
    name: readString(source, 'name', 'Name'),
    enabled: readBoolean(source, 'enabled', 'Enabled'),
    executionOrder: readNumber(source, 'executionOrder', 'ExecutionOrder'),
    ruleType,
    isLocalRule: readBoolean(source, 'isLocalRule', 'IsLocalRule'),
    canModifyDefinition: readBoolean(source, 'canModifyDefinition', 'CanModifyDefinition', true),
    conditions: readStringArray(source.conditions ?? source.Conditions),
    actions: readStringArray(source.actions ?? source.Actions),
    exceptions: readStringArray(source.exceptions ?? source.Exceptions),
  }
}

export function normalizeOutlookRules(items: unknown): OutlookRuleDto[] {
  return Array.isArray(items) ? items.map(normalizeOutlookRule).filter((rule) => rule.name.trim()) : []
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
  getMailSearchProgress: async (searchId: string) =>
    normalizeMailSearchProgress(await getJson<unknown>(`/api/outlook/mail-search/progress/${encodeURIComponent(searchId)}`)),
  lookupAddressBookContact: async (email: string) =>
    normalizeAddressBookLookupResponse(await getJson<unknown>(`/api/outlook/address-book/lookup?email=${encodeURIComponent(email)}`)),
  suggestAddressBookMerges: async (recipients: string[]) =>
    normalizeAddressBookMergeSuggestionResponse(await postJson<unknown>('/api/outlook/address-book/merge-suggestions', { recipients })),
  requestAddressBookGroupMembers: async (body: { groupId?: string; groupSmtpAddress?: string; maxMembers?: number; forceRefresh?: boolean }) => {
    const response = await postJson<OutlookRequestResponse<unknown>>('/api/outlook/request-address-book-group-members', body)
    return { ...response, data: normalizeAddressBookGroupMembersResponse(response.data) }
  },
  normalizeAddressBookGroupMembersResponse,
  getChat: () => getJson<ChatMessageDto[]>('/api/outlook/chat'),
  getAdminStatus: () => getJson<AddinStatusDto>('/api/outlook/admin/status'),
  getAdminLogs: () => getJson<AddinLogEntry[]>('/api/outlook/admin/logs'),
  getAttachmentExportSettings: () => getJson<AttachmentExportSettingsDto>('/api/outlook/attachment-export-settings'),
  fetchResult: <TData = Record<string, unknown>>(endpoint: string, body: FetchResultRequest) =>
    postJson<FetchResultResponse<TData>>(`/api/outlook/${endpoint}`, body),

  requestFolders: () => postJson<OutlookRequestResponse>('/api/outlook/request-folders'),
  requestFolderChildren: (body: {
    storeId: string
    parentEntryId: string
    parentFolderPath: string
    maxDepth?: number
    maxChildren?: number
  }) => postJson<OutlookRequestResponse>('/api/outlook/request-folder-children', body),
  requestMails: (body: { folderPath: string; lookbackHours?: number; maxCount: number; receivedFrom?: string; receivedTo?: string }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-mails', body),
  requestFolderMails: (body: FolderMailsRequest & { maxCount?: number }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-folder-mails', body),
  requestMailSearch: (body: SearchMailsRequest) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-mail-search', body),
  requestMailBody: (body: { mailId: string; folderPath: string }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-mail-body', body),
  requestMailAttachments: (body: { mailId: string; folderPath: string }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-mail-attachments', body),
  requestMailConversation: (body: { mailId: string; folderPath: string; maxCount?: number; includeBody?: boolean }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-mail-conversation', body),
  requestExportMailAttachment: (body: {
    mailId: string
    folderPath: string
    attachmentId: string
    index: number
    name: string
    fileName: string
    displayName: string
  }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-export-mail-attachment', body),
  openExportedAttachment: (body: { exportedAttachmentId: string }) =>
    postJson('/api/outlook/open-exported-attachment', body),
  updateAttachmentExportSettings: (body: { rootPath: string }) =>
    postJson<AttachmentExportSettingsDto>('/api/outlook/attachment-export-settings', body),
  requestRules: () => postJson<OutlookRequestResponse>('/api/outlook/request-rules'),
  requestManageRule: (body: OutlookRuleCommandRequest) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-manage-rule', body),
  requestCategories: () => postJson<OutlookRequestResponse>('/api/outlook/request-categories'),
  requestSignalRPing: () => postJson<OutlookRequestResponse>('/api/outlook/request-signalr-ping'),
  requestCalendar: (body: { daysForward: number; startDate?: string; endDate?: string }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-calendar', body),
  requestCalendarRooms: () => postJson<OutlookRequestResponse>('/api/outlook/request-calendar-rooms'),
  requestCreateCalendarEvent: (body: {
    subject: string
    start: string
    end: string
    location: string
    body: string
    busyStatus: string
    requiredAttendees?: OutlookRecipientDto[]
    resources?: OutlookRecipientDto[]
  }) => postJson<OutlookRequestResponse>('/api/outlook/request-create-calendar-event', body),
  requestUpdateCalendarEvent: (body: {
    eventId: string
    smartOfficeEventId: string
    subject: string
    start: string
    end: string
    location: string
    body: string
    busyStatus: string
    requiredAttendees?: OutlookRecipientDto[]
    resources?: OutlookRecipientDto[]
  }) => postJson<OutlookRequestResponse>('/api/outlook/request-update-calendar-event', body),
  requestDeleteCalendarEvent: (body: { eventId: string; smartOfficeEventId: string }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-delete-calendar-event', body),
  requestAddressBook: (body: {
    includeOutlookContacts: boolean
    includeAddressLists: boolean
    maxContacts: number
    maxAddressEntriesPerList: number
    maxGroupMembers?: number
    maxGroupDepth?: number
  }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-address-book', body),
  sendChat: (body: { source: 'web'; text: string }) => postJson('/api/outlook/chat', body),

  requestUpdateMailProperties: (body: MailPropertiesCommandRequest) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-update-mail-properties', body),
  requestUpsertCategory: (body: CategoryCommandRequest) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-upsert-category', body),
  requestCreateFolder: (body: { parentFolderPath: string; name: string }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-create-folder', body),
  requestDeleteFolder: (body: { folderPath: string }) => postJson<OutlookRequestResponse>('/api/outlook/request-delete-folder', body),
  requestMoveMail: (body: { mailId: string; sourceFolderPath: string; destinationFolderPath: string }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-move-mail', body),
  requestMoveMails: (body: { mailIds: string[]; sourceFolderPath: string; sourceFolderPaths: string[]; destinationFolderPath: string; continueOnError: boolean }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-move-mails', body),
  requestDeleteMail: (body: { mailId: string; folderPath: string }) =>
    postJson<OutlookRequestResponse>('/api/outlook/request-delete-mail', body),
}
