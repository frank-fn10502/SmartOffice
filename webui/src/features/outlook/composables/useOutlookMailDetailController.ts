import type { Ref } from 'vue'
import {
  normalizeMailItems,
  outlookApi,
} from '../api/outlook'
import type { MailAttachmentDto, MailAttachmentsDto, MailConversationDto, MailItemDto } from '../models/outlook'
import { canUpdateMailProperties } from '../utils/outlookItemTypes'
import { fetchResultEndpoint, requestIdFromResponse } from './outlookRequests'

type MailDetailControllerOptions = {
  exportingAttachmentIds: Ref<Set<string>>
  loadingAttachmentMailIds: Ref<Set<string>>
  loadingConversationMailIds: Ref<Set<string>>
  loadingMailBodyIds: Ref<Set<string>>
  mailAttachmentsByMailId: Ref<Record<string, MailAttachmentDto[]>>
  mailConversationsByMailId: Ref<Record<string, MailConversationDto>>
  attachmentKey: (mailId: string, attachmentId: string) => string
  completeAttachmentExport: (mailId: string, attachmentId: string) => void
  completeAttachmentLoad: (mailId: string) => void
  completeConversationLoad: (mailId: string) => void
  completeMailBodyLoad: (mailId: string) => void
  loadRequestMailItems: (response: { requestId?: string; request?: string }) => Promise<MailItemDto[]>
  patchMailAttachments: (payload: MailAttachmentsDto) => void
  patchMailConversation: (payload: MailConversationDto) => void
  patchMailSnapshots: (items: MailItemDto[]) => void
  waitForRequest: (response: { requestId?: string; request?: string }, timeoutMs?: number) => Promise<void>
}

export function useOutlookMailDetailController(options: MailDetailControllerOptions) {
  const {
    attachmentKey,
    completeAttachmentExport,
    completeAttachmentLoad,
    completeConversationLoad,
    completeMailBodyLoad,
    exportingAttachmentIds,
    loadRequestMailItems,
    loadingAttachmentMailIds,
    loadingConversationMailIds,
    loadingMailBodyIds,
    mailAttachmentsByMailId,
    mailConversationsByMailId,
    patchMailAttachments,
    patchMailConversation,
    patchMailSnapshots,
    waitForRequest,
  } = options

  function isMailBodyLoading(mail: MailItemDto) {
    return Boolean(mail.id && loadingMailBodyIds.value.has(mail.id))
  }

  function mailHasBody(mail: MailItemDto) {
    return Boolean(mail.body || mail.bodyHtml)
  }

  function isAttachmentListLoading(mail: MailItemDto) {
    return Boolean(mail.id && loadingAttachmentMailIds.value.has(mail.id))
  }

  function isAttachmentExporting(mail: MailItemDto, attachment: MailAttachmentDto) {
    return Boolean(mail.id && exportingAttachmentIds.value.has(attachmentKey(mail.id, attachment.attachmentId)))
  }

  function isConversationLoading(mail: MailItemDto) {
    return Boolean(mail.id && loadingConversationMailIds.value.has(mail.id))
  }

  async function requestMailBody(mail: MailItemDto) {
    if (!mail.id?.trim() || mailHasBody(mail) || isMailBodyLoading(mail)) return
    loadingMailBodyIds.value = new Set(loadingMailBodyIds.value).add(mail.id)
    try {
      const response = await outlookApi.requestMailBody({
        mailId: mail.id,
        folderPath: mail.folderPath,
      })
      await waitForRequest(response)
      patchMailSnapshots(await loadRequestMailItems(response))
      completeMailBodyLoad(mail.id)
    } catch {
      completeMailBodyLoad(mail.id)
    }
  }

  async function requestMailAttachments(mail: MailItemDto) {
    if (!mail.id?.trim() || isAttachmentListLoading(mail) || mailAttachmentsByMailId.value[mail.id]) return
    loadingAttachmentMailIds.value = new Set(loadingAttachmentMailIds.value).add(mail.id)
    try {
      const response = await outlookApi.requestMailAttachments({
        mailId: mail.id,
        folderPath: mail.folderPath,
      })
      await waitForRequest(response)
      patchMailAttachments(await outlookApi.getMailAttachments(mail.id))
    } catch {
      completeAttachmentLoad(mail.id)
    }
  }

  async function requestMailConversation(mail: MailItemDto) {
    if (!mail.id?.trim() || !canUpdateMailProperties(mail) || isConversationLoading(mail) || mailConversationsByMailId.value[mail.id]) return
    loadingConversationMailIds.value = new Set(loadingConversationMailIds.value).add(mail.id)
    try {
      const response = await outlookApi.requestMailConversation({
        mailId: mail.id,
        folderPath: mail.folderPath,
        maxCount: 100,
        includeBody: true,
      })
      await waitForRequest(response)
      const requestId = requestIdFromResponse(response)
      const items: MailItemDto[] = []
      let cursor = ''
      let conversation: MailConversationDto | null = null
      do {
        const state = await outlookApi.fetchResult<{
          mailId?: string
          folderPath?: string
          conversationId?: string
          conversationTopic?: string
          mails?: unknown[]
        }>(fetchResultEndpoint(response), {
          requestId,
          cursor,
          take: 100,
        })
        const data = state.data ?? {}
        conversation = {
          mailId: data.mailId || mail.id,
          folderPath: data.folderPath || mail.folderPath,
          conversationId: data.conversationId || mail.conversationId,
          conversationTopic: data.conversationTopic || mail.conversationTopic || mail.subject,
          mails: items,
        }
        items.push(...normalizeMailItems(data.mails))
        cursor = state.next.cursor
        if (!state.next.hasMore) break
      } while (cursor)

      if (conversation) {
        conversation.mails = items
        patchMailConversation(conversation)
      } else {
        patchMailConversation(await outlookApi.getMailConversation(mail.id))
      }
    } catch {
      completeConversationLoad(mail.id)
    }
  }

  async function exportMailAttachment(mail: MailItemDto, attachment: MailAttachmentDto) {
    if (!mail.id?.trim() || !attachment.attachmentId || isAttachmentExporting(mail, attachment)) return
    const key = attachmentKey(mail.id, attachment.attachmentId)
    const exportAttachmentId = attachment.index > 0 ? String(attachment.index) : attachment.attachmentId
    exportingAttachmentIds.value = new Set(exportingAttachmentIds.value).add(key)
    try {
      const response = await outlookApi.requestExportMailAttachment({
        mailId: mail.id,
        folderPath: mail.folderPath,
        attachmentId: exportAttachmentId,
        index: attachment.index,
        name: attachment.name,
        fileName: attachment.fileName,
        displayName: attachment.displayName,
      })
      await waitForRequest(response)
      patchMailAttachments(await outlookApi.getMailAttachments(mail.id))
      completeAttachmentExport(mail.id, attachment.attachmentId)
    } catch {
      completeAttachmentExport(mail.id, attachment.attachmentId)
    }
  }

  async function openExportedAttachment(attachment: MailAttachmentDto) {
    if (!attachment.exportedAttachmentId) return
    await outlookApi.openExportedAttachment({ exportedAttachmentId: attachment.exportedAttachmentId })
  }

  return {
    exportMailAttachment,
    isAttachmentExporting,
    isAttachmentListLoading,
    isConversationLoading,
    isMailBodyLoading,
    mailHasBody,
    openExportedAttachment,
    requestMailAttachments,
    requestMailBody,
    requestMailConversation,
  }
}
