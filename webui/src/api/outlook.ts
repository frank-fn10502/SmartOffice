import type {
  AddinLogEntry,
  AddinStatusDto,
  CalendarEventDto,
  CategoryCommandRequest,
  ChatMessageDto,
  FolderDto,
  MailPropertiesCommandRequest,
  MailItemDto,
  OutlookCategoryDto,
  OutlookRuleDto,
} from '../models/outlook'

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

export function pollUntil(check: () => Promise<boolean>, timeoutMs: number) {
  return new Promise<boolean>((resolve) => {
    const start = Date.now()
    const timer = window.setInterval(async () => {
      try {
        const done = await check()
        if (done || Date.now() - start >= timeoutMs) {
          window.clearInterval(timer)
          resolve(done)
        }
      } catch {
        if (Date.now() - start >= timeoutMs) {
          window.clearInterval(timer)
          resolve(false)
        }
      }
    }, 1200)
  })
}

export const outlookApi = {
  getFolders: () => getJson<FolderDto[]>('/api/outlook/folders'),
  getMails: () => getJson<MailItemDto[]>('/api/outlook/mails'),
  getRules: () => getJson<OutlookRuleDto[]>('/api/outlook/rules'),
  getCategories: () => getJson<OutlookCategoryDto[]>('/api/outlook/categories'),
  getCalendar: () => getJson<CalendarEventDto[]>('/api/outlook/calendar'),
  getChat: () => getJson<ChatMessageDto[]>('/api/outlook/chat'),
  getAdminStatus: () => getJson<AddinStatusDto>('/api/outlook/admin/status'),
  getAdminLogs: () => getJson<AddinLogEntry[]>('/api/outlook/admin/logs'),

  requestFolders: () => postJson('/api/outlook/request-folders'),
  requestMails: (body: { folderPath: string; range: string; maxCount: number }) =>
    postJson('/api/outlook/request-mails', body),
  requestRules: () => postJson('/api/outlook/request-rules'),
  requestCategories: () => postJson('/api/outlook/request-categories'),
  requestCalendar: (body: { daysForward: number }) => postJson('/api/outlook/request-calendar', body),
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
