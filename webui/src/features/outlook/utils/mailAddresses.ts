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

export function groupRecipientsForMail(mail: Pick<MailItemDto, 'toRecipients' | 'ccRecipients' | 'bccRecipients'>) {
  return [...mail.toRecipients, ...mail.ccRecipients, ...mail.bccRecipients]
    .filter((recipient) => recipient.isGroup)
    .filter((recipient, index, recipients) => {
      const key = (recipient.smtpAddress || recipient.rawAddress || recipient.displayName).trim().toLowerCase()
      return recipients.findIndex((item) => (item.smtpAddress || item.rawAddress || item.displayName).trim().toLowerCase() === key) === index
    })
}

export function formatGroupRecipientSummary(mail: Pick<MailItemDto, 'toRecipients' | 'ccRecipients' | 'bccRecipients'>) {
  return formatRecipients(groupRecipientsForMail(mail))
}
