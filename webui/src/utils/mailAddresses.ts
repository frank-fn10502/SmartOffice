import type { MailItemDto, OutlookRecipientDto } from '../models/outlook'

const addressInAngleBracketsPattern = /<[^<>]*>/
const smtpAddressPattern = /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/

function cleanDisplayName(value: string) {
  const withoutAddress = value.replace(addressInAngleBracketsPattern, '').trim()
  return withoutAddress.replace(/^["']|["']$/g, '').trim()
}

export function formatMailSender(mail: Pick<MailItemDto, 'sender'>) {
  return formatRecipient(mail.sender, 'Unknown sender')
}

export function formatRecipient(recipient: OutlookRecipientDto, fallback = 'Unknown recipient') {
  const name = cleanDisplayName(recipient.displayName)
  if (name) return name

  const smtpName = cleanDisplayName(recipient.smtpAddress)
  if (smtpName) return smtpName

  const rawName = cleanDisplayName(recipient.rawAddress)
  return rawName || fallback
}

export function formatRecipients(recipients: OutlookRecipientDto[]) {
  return recipients.map((recipient) => formatRecipient(recipient)).join('、')
}

export function shouldShowRecipientSmtpAddress(recipient: OutlookRecipientDto) {
  const email = recipient.smtpAddress.trim()
  if (!email) return false
  if (!smtpAddressPattern.test(email)) return false

  return !recipient.displayName.includes(email) && formatRecipient(recipient) !== email
}
