import type {
  AddinLogEntry,
  AddinStatusDto,
  CalendarEventDto,
  CategoryCommandRequest,
  ChatMessageDto,
  CommandDispatchResponse,
  FolderSnapshotDto,
  MailPropertiesCommandRequest,
  MailItemDto,
  OutlookCategoryDto,
  OutlookRuleDto,
} from '../models/outlook'

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

export function normalizeCategoryColor(value: string) {
  const key = value.trim()
  if (!key) return 'olCategoryColorNone'

  const outlookColorMap: Record<string, string> = {
    '0': 'olCategoryColorNone',
    '1': 'olCategoryColorRed',
    '2': 'olCategoryColorOrange',
    '3': 'olCategoryColorPeach',
    '4': 'olCategoryColorYellow',
    '5': 'olCategoryColorGreen',
    '6': 'olCategoryColorTeal',
    '7': 'olCategoryColorOlive',
    '8': 'olCategoryColorBlue',
    '9': 'olCategoryColorPurple',
    '10': 'olCategoryColorMaroon',
    '11': 'olCategoryColorSteel',
    '12': 'olCategoryColorDarkSteel',
    '13': 'olCategoryColorGray',
    '14': 'olCategoryColorDarkGray',
    '15': 'olCategoryColorBlack',
    '16': 'olCategoryColorDarkRed',
    '17': 'olCategoryColorDarkOrange',
    '18': 'olCategoryColorDarkPeach',
    '19': 'olCategoryColorDarkYellow',
    '20': 'olCategoryColorDarkGreen',
    '21': 'olCategoryColorDarkTeal',
    '22': 'olCategoryColorDarkOlive',
    '23': 'olCategoryColorDarkBlue',
    '24': 'olCategoryColorDarkPurple',
    '25': 'olCategoryColorDarkMaroon',
    olcategorycolornone: 'olCategoryColorNone',
    olcategorycolorred: 'olCategoryColorRed',
    olcategorycolororange: 'olCategoryColorOrange',
    olcategorycolorpeach: 'olCategoryColorPeach',
    olcategorycoloryellow: 'olCategoryColorYellow',
    olcategorycolorgreen: 'olCategoryColorGreen',
    olcategorycolorteal: 'olCategoryColorTeal',
    olcategorycolorolive: 'olCategoryColorOlive',
    olcategorycolorblue: 'olCategoryColorBlue',
    olcategorycolorpurple: 'olCategoryColorPurple',
    olcategorycolormaroon: 'olCategoryColorMaroon',
    olcategorycolorsteel: 'olCategoryColorSteel',
    olcategorycolordarksteel: 'olCategoryColorDarkSteel',
    olcategorycolorgray: 'olCategoryColorGray',
    olcategorycolordarkgray: 'olCategoryColorDarkGray',
    olcategorycolorblack: 'olCategoryColorBlack',
    olcategorycolordarkred: 'olCategoryColorDarkRed',
    olcategorycolordarkorange: 'olCategoryColorDarkOrange',
    olcategorycolordarkpeach: 'olCategoryColorDarkPeach',
    olcategorycolordarkyellow: 'olCategoryColorDarkYellow',
    olcategorycolordarkgreen: 'olCategoryColorDarkGreen',
    olcategorycolordarkteal: 'olCategoryColorDarkTeal',
    olcategorycolordarkolive: 'olCategoryColorDarkOlive',
    olcategorycolordarkblue: 'olCategoryColorDarkBlue',
    olcategorycolordarkpurple: 'olCategoryColorDarkPurple',
    olcategorycolordarkmaroon: 'olCategoryColorDarkMaroon',
  }

  return outlookColorMap[key.toLowerCase()] ?? key
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

export function normalizeOutlookCategory(item: unknown): OutlookCategoryDto {
  const source = (item ?? {}) as LooseRecord
  return {
    name: readString(source, 'name', 'Name'),
    color: normalizeCategoryColor(readString(source, 'color', 'Color')),
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
  getRules: () => getJson<OutlookRuleDto[]>('/api/outlook/rules'),
  getCategories: async () => normalizeOutlookCategories(await getJson<unknown>('/api/outlook/categories')),
  getCalendar: () => getJson<CalendarEventDto[]>('/api/outlook/calendar'),
  getChat: () => getJson<ChatMessageDto[]>('/api/outlook/chat'),
  getAdminStatus: () => getJson<AddinStatusDto>('/api/outlook/admin/status'),
  getAdminLogs: () => getJson<AddinLogEntry[]>('/api/outlook/admin/logs'),

  requestFolders: () => postJson<CommandDispatchResponse>('/api/outlook/request-folders'),
  requestMails: (body: { folderPath: string; range: string; maxCount: number }) =>
    postJson('/api/outlook/request-mails', body),
  requestRules: () => postJson('/api/outlook/request-rules'),
  requestCategories: () => postJson('/api/outlook/request-categories'),
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
}
