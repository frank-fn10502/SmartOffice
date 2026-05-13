import type { MailItemDto } from '../models/outlook'

export function mergeMailSnapshot(current: MailItemDto, incoming: MailItemDto) {
  return {
    ...current,
    ...incoming,
    body: incoming.body || current.body,
    bodyHtml: incoming.bodyHtml || current.bodyHtml,
  }
}

export function patchMailSnapshotList(list: MailItemDto[], items: MailItemDto[]) {
  if (items.length === 0) return list
  const byId = new Map(items.filter((item) => item.id).map((item) => [item.id, item]))
  return list.map((mail) => {
    const patch = byId.get(mail.id)
    return patch ? mergeMailSnapshot(mail, patch) : mail
  })
}
