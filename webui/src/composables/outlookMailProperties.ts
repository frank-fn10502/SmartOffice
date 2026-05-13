import type { MailItemDto, MailPropertiesDraft } from '../models/outlook'
import {
  dateInputToIso,
  defaultFlagRequest,
  isDefaultFlagRequest,
  splitCategories,
  toDateInput,
} from '../utils/outlookDashboardHelpers'

export function buildMailPropertiesDraft(mail: MailItemDto): MailPropertiesDraft {
  const flagInterval = mail.flagInterval || (mail.isMarkedAsTask ? 'today' : 'none')
  return {
    isRead: mail.isRead,
    flagInterval,
    flagRequest: isDefaultFlagRequest(mail.flagRequest) ? defaultFlagRequest(flagInterval) : mail.flagRequest,
    taskStartDate: toDateInput(mail.taskStartDate),
    taskDueDate: toDateInput(mail.taskDueDate),
    taskCompletedDate: toDateInput(mail.taskCompletedDate),
    categories: splitCategories(mail.categories),
  }
}

export function normalizeMailPropertiesDraft(draft: MailPropertiesDraft): MailPropertiesDraft {
  return {
    isRead: draft.isRead,
    flagInterval: draft.flagInterval || 'none',
    flagRequest: draft.flagInterval === 'none' ? '' : (draft.flagRequest || defaultFlagRequest(draft.flagInterval)).trim(),
    taskStartDate: draft.taskStartDate || '',
    taskDueDate: draft.taskDueDate || '',
    taskCompletedDate: draft.taskCompletedDate || '',
    categories: [...new Set(draft.categories.map((category) => category.trim()).filter(Boolean))]
      .sort((left, right) => left.localeCompare(right, undefined, { sensitivity: 'base' })),
  }
}

export function buildMailPropertiesPayload(mail: MailItemDto, draft: MailPropertiesDraft) {
  const normalized = normalizeMailPropertiesDraft(draft)
  const isCustomFlag = normalized.flagInterval === 'custom'
  return {
    mailId: mail.id,
    folderPath: mail.folderPath,
    isRead: normalized.isRead,
    flagInterval: normalized.flagInterval,
    flagRequest: normalized.flagRequest,
    taskStartDate: isCustomFlag ? dateInputToIso(normalized.taskStartDate) : undefined,
    taskDueDate: isCustomFlag ? dateInputToIso(normalized.taskDueDate) : undefined,
    taskCompletedDate: normalized.flagInterval === 'complete' ? dateInputToIso(normalized.taskCompletedDate) : undefined,
    categories: normalized.categories,
  }
}
