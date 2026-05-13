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

export function canUseMailMutation(mail: Pick<MailItemDto, 'messageClass'>) {
  return isStandardMailMessage(mail)
}

export function outlookItemTypeLabel(mail: Pick<MailItemDto, 'messageClass'>) {
  if (isMeetingMessage(mail)) return 'æè­°éè«'
  if (isStandardMailMessage(mail)) return ''
  return normalizedMessageClass(mail) || 'Outlook item'
}
