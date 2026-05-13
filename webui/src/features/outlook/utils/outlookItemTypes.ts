import type { MailItemDto } from '../models/outlook'

export function normalizedMessageClass(mail: Pick<MailItemDto, 'messageClass'>) {
  return mail.messageClass.trim()
}

export function isMeetingMessage(mail: Pick<MailItemDto, 'messageClass'>) {
  return normalizedMessageClass(mail).toLowerCase().startsWith('ipm.schedule.meeting')
}

export function isStandardMailMessage(mail: Pick<MailItemDto, 'messageClass'>) {
  const messageClass = normalizedMessageClass(mail).toLowerCase()
  return messageClass === '' || messageClass === 'ipm.note'
}

export function canUpdateMailProperties(mail: Pick<MailItemDto, 'messageClass'>) {
  return isStandardMailMessage(mail)
}

export function canMoveOutlookItem(_mail: Pick<MailItemDto, 'messageClass'>) {
  return true
}

export function outlookItemTypeLabel(mail: Pick<MailItemDto, 'messageClass'>) {
  if (isMeetingMessage(mail)) return '會議邀請'
  if (isStandardMailMessage(mail)) return ''
  return normalizedMessageClass(mail) || 'Outlook item'
}
