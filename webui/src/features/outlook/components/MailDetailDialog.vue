<script setup lang="ts">
import { computed, nextTick, ref } from 'vue'
import MailPropertyPane from './MailPropertyPane.vue'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'
import { collectOutlookRequestData, waitForOutlookRequest } from '../composables/outlookRequests'
import { normalizeAddressBookContact, outlookApi } from '../api/outlook'
import type { AddressBookContactDto, AddressBookRecipientRelevanceDto, OutlookRecipientDto } from '../models/outlook'
import { formatDateTime } from '../utils/formatters'
import { formatMailSender, formatRecipient, shouldShowRecipientSmtpAddress } from '../utils/mailAddresses'
import { outlookItemTypeLabel } from '../utils/outlookItemTypes'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

interface RecipientTreeNode {
  key: string
  label: string
  email: string
  isGroup: boolean
  loaded: boolean
  loading: boolean
  children: RecipientTreeNode[]
  recipientRelevance?: AddressBookRecipientRelevanceDto
}

type RecipientTreeNodeHandle = {
  expanded?: boolean
  expand?: () => void
  collapse?: () => void
}

function formatAttachmentSize(size: number) {
  if (size >= 1024 * 1024) return `${(size / 1024 / 1024).toFixed(1)} MB`
  if (size >= 1024) return `${Math.round(size / 1024)} KB`
  return `${size} B`
}

function formatAttachmentMeta(contentType: string, size: number) {
  return `${contentType.trim() || 'unknown'} · ${formatAttachmentSize(size)}`
}

function mailPreviewText(body: string, bodyHtml: string) {
  const text = body || bodyHtml.replace(/<[^>]+>/g, ' ')
  return text.replace(/\s+/g, ' ').trim()
}

const groupTreeVisible = ref(false)
const groupTreeLoading = ref(false)
const groupTreeMessage = ref('')
const groupTreeRoot = ref<RecipientTreeNode | null>(null)
const groupTreeRef = ref<{
  updateKeyChildren?: (key: string, children: RecipientTreeNode[]) => void
  getNode?: (key: string) => RecipientTreeNodeHandle | undefined
} | null>(null)
const groupTreeProps = { label: 'label', children: 'children' }
const groupTreeNodes = computed(() => groupTreeRoot.value ? [groupTreeRoot.value] : [])
const groupTreeExpandedKeys = computed(() => groupTreeRoot.value ? [groupTreeRoot.value.key] : [])

const {
  addMailCategoryDraft,
  applyMailProperties,
  categories,
  categoryTagStyle,
  closeMailDialog,
  dialogLoading,
  dialogMail,
  dialogMailAttachments,
  dialogMailConversation,
  dialogMailConversationItems,
  dialogMailFolderName,
  dialogMailHasIdentity,
  exportMailAttachment,
  flagDisplayLabel,
  flagIntervalOptions,
  flagTagType,
  isAttachmentExporting,
  isAttachmentListLoading,
  isConversationLoading,
  isMailBodyLoading,
  loadingCategories,
  mailDialogHtml,
  mailDialogVisible,
  mailHasBody,
  mailHtmlSandbox,
  mailPropertiesChanged,
  mailPropertiesDraft,
  openCategoryManager,
  openExportedAttachment,
  operationLoading,
  outlookBusy,
  removeMailCategoryDraft,
  resetMailPropertiesDraft,
  setMailFlagDraft,
  splitCategories,
} = props.dashboard

function groupRecipientKey(recipient: OutlookRecipientDto) {
  return (recipient.smtpAddress || recipient.rawAddress || recipient.displayName).trim().toLowerCase()
}

function uniqueRecipients(recipients: OutlookRecipientDto[]) {
  return recipients
    .filter((recipient, index, items) => items.findIndex((item) => groupRecipientKey(item) === groupRecipientKey(recipient)) === index)
}

function recipientToTreeNode(recipient: OutlookRecipientDto): RecipientTreeNode {
  const label = formatRecipient(recipient)
  return {
    key: groupRecipientKey(recipient) || label,
    label,
    email: recipient.smtpAddress || recipient.rawAddress,
    isGroup: recipient.isGroup,
    loaded: false,
    loading: false,
    children: recipient.members.map((member) => recipientToTreeNode(member)),
  }
}

function contactToTreeNode(contact: AddressBookContactDto): RecipientTreeNode {
  const label = contact.displayName || contact.smtpAddress || contact.rawAddress || '(unknown)'
  return {
    key: (contact.smtpAddress || contact.rawAddress || contact.id || label).trim().toLowerCase(),
    label,
    email: contact.smtpAddress || contact.rawAddress,
    isGroup: contact.isGroup,
    loaded: false,
    loading: false,
    children: [],
  }
}

function normalizeRecipientRelevance(value: unknown): AddressBookRecipientRelevanceDto | undefined {
  const source = value as Partial<AddressBookRecipientRelevanceDto> | undefined
  if (!source || typeof source !== 'object') return undefined
  return {
    score: typeof source.score === 'number' ? source.score : 0,
    level: typeof source.level === 'string' ? source.level : 'unknown',
    summary: typeof source.summary === 'string' ? source.summary : '',
    routeDepth: typeof source.routeDepth === 'number' ? source.routeDepth : -1,
    directPersonCount: typeof source.directPersonCount === 'number' ? source.directPersonCount : 0,
    directGroupCount: typeof source.directGroupCount === 'number' ? source.directGroupCount : 0,
    audienceSize: typeof source.audienceSize === 'number' ? source.audienceSize : 0,
    includesSelf: Boolean(source.includesSelf),
    includesSelfDirectly: Boolean(source.includesSelfDirectly),
    reasons: Array.isArray(source.reasons) ? source.reasons.map((item) => String(item)) : [],
  }
}

function relevanceLabel(relevance?: AddressBookRecipientRelevanceDto) {
  if (!relevance) return ''
  return `${relevance.score} · ${relevance.level} · depth ${relevance.routeDepth}`
}

async function openGroupTree(recipient: OutlookRecipientDto) {
  const root = recipientToTreeNode(recipient)
  groupTreeRoot.value = root
  groupTreeVisible.value = true
  await expandGroupTreeNode(root)
}

async function toggleGroupTreeNode(data: RecipientTreeNode, treeNode: RecipientTreeNodeHandle) {
  if (!data.isGroup) return
  if (!data.loaded) {
    await expandGroupTreeNode(data)
    return
  }

  if (treeNode.expanded) {
    if (typeof treeNode.collapse === 'function') treeNode.collapse()
    else treeNode.expanded = false
    return
  }

  if (typeof treeNode.expand === 'function') treeNode.expand()
  else treeNode.expanded = true
}

async function expandGroupTreeNode(node: RecipientTreeNode) {
  if (!node.isGroup || node.loading || node.loaded) return
  node.loading = true
  groupTreeLoading.value = true
  groupTreeMessage.value = `查詢 ${node.label}...`
  try {
    const response = await outlookApi.requestAddressBookRelation({
      targetKind: 'group',
      groupSmtpAddress: node.email,
      groupId: node.email ? '' : node.key,
      take: 100,
    })
    await waitForOutlookRequest(response)
    const pages = await collectOutlookRequestData<Record<string, unknown>>(response, { take: 100 })
    const data = pages[0]?.data ?? {}
    const members = Array.isArray(data.members) ? data.members.map(normalizeAddressBookContact) : []
    node.recipientRelevance = normalizeRecipientRelevance(data.recipientRelevance)
    node.children = members.map(contactToTreeNode)
    node.loaded = true
    groupTreeRef.value?.updateKeyChildren?.(node.key, node.children)
    await nextTick()
    const treeNode = groupTreeRef.value?.getNode?.(node.key)
    if (treeNode) {
      if (typeof treeNode.expand === 'function') treeNode.expand()
      else treeNode.expanded = true
    }
    groupTreeMessage.value = node.children.length > 0
      ? `${node.label} 有 ${node.children.length} 個 direct members。`
      : `${node.label} 沒有可顯示的 direct members。`
  } catch (error) {
    groupTreeMessage.value = error instanceof Error ? error.message : 'Group 查詢失敗。'
  } finally {
    node.loading = false
    groupTreeLoading.value = false
  }
}
</script>

<template>
  <el-dialog
    v-model="mailDialogVisible"
    class="mail-detail-dialog"
    width="min(1160px, calc(100vw - 28px))"
    destroy-on-close
    @closed="closeMailDialog"
  >
    <template #header>
      <div v-if="dialogMail" class="dialog-mail-title">
        <strong>{{ dialogMail.subject || '(No subject)' }}</strong>
        <span>
          {{ formatMailSender(dialogMail) }} · {{ formatDateTime(dialogMail.receivedTime) }}
          <el-tag v-if="outlookItemTypeLabel(dialogMail)" size="small" type="success" effect="plain">
            {{ outlookItemTypeLabel(dialogMail) }}
          </el-tag>
          <el-tag v-if="dialogLoading" size="small" type="info" effect="plain">載入中</el-tag>
        </span>
      </div>
    </template>

    <div v-if="dialogMail" class="dialog-mail-layout">
      <div class="dialog-mail-content">
        <div class="dialog-mail-summary">
          <div class="dialog-mail-meta">
            <span>寄件者</span>
            <strong>{{ formatMailSender(dialogMail) }}</strong>
            <small v-if="shouldShowRecipientSmtpAddress(dialogMail.sender)">{{ dialogMail.sender.smtpAddress }}</small>
          </div>
          <div class="dialog-mail-meta">
            <span>收件者</span>
            <span v-if="dialogMail.toRecipients.length > 0" class="recipient-chip-list">
              <button
                v-for="recipient in uniqueRecipients(dialogMail.toRecipients)"
                :key="groupRecipientKey(recipient)"
                class="recipient-chip"
                :class="{ group: recipient.isGroup }"
                type="button"
                :disabled="!recipient.isGroup"
                @click="recipient.isGroup && openGroupTree(recipient)"
              >
                {{ formatRecipient(recipient) }}<span v-if="recipient.isGroup">Group</span>
              </button>
            </span>
            <strong v-else>-</strong>
          </div>
          <div v-if="dialogMail.ccRecipients.length > 0" class="dialog-mail-meta">
            <span>副本</span>
            <span class="recipient-chip-list">
              <button
                v-for="recipient in uniqueRecipients(dialogMail.ccRecipients)"
                :key="groupRecipientKey(recipient)"
                class="recipient-chip"
                :class="{ group: recipient.isGroup }"
                type="button"
                :disabled="!recipient.isGroup"
                @click="recipient.isGroup && openGroupTree(recipient)"
              >
                {{ formatRecipient(recipient) }}<span v-if="recipient.isGroup">Group</span>
              </button>
            </span>
          </div>
          <div class="dialog-mail-meta">
            <span>Folder</span>
            <strong>{{ dialogMail.folderPath }}</strong>
          </div>
        </div>
        <div class="mail-row-tags dialog-mail-tags">
          <el-tag v-if="outlookItemTypeLabel(dialogMail)" type="success" effect="plain">{{ outlookItemTypeLabel(dialogMail) }}</el-tag>
          <el-tag v-if="!dialogMail.isRead" type="warning" effect="plain">未讀</el-tag>
          <el-tag v-if="dialogMail.isMarkedAsTask" :type="flagTagType(dialogMail.flagInterval)" effect="plain">
            {{ flagDisplayLabel(dialogMail.flagInterval, dialogMail.flagRequest) }}
          </el-tag>
          <el-tag
            v-for="category in splitCategories(dialogMail.categories)"
            :key="category"
            effect="dark"
            :style="categoryTagStyle(category)"
          >
            {{ category }}
          </el-tag>
        </div>

        <div class="dialog-mail-section-head">
          <strong>內容</strong>
          <el-button v-if="mailHasBody(dialogMail)" size="small" @click="mailDialogHtml = !mailDialogHtml">
            {{ mailDialogHtml ? '切到文字' : '切到 HTML' }}
          </el-button>
        </div>
        <div class="dialog-mail-body-frame">
          <div v-if="isMailBodyLoading(dialogMail)" class="dialog-loading-block">
            <el-skeleton animated :rows="6" />
          </div>
          <iframe
            v-else-if="mailHasBody(dialogMail) && mailDialogHtml"
            class="mail-html dialog-mail-body"
            :sandbox="mailHtmlSandbox"
            referrerpolicy="no-referrer"
            :srcdoc="dialogMail.bodyHtml || dialogMail.body"
          />
          <pre v-else-if="mailHasBody(dialogMail)" class="mail-text dialog-mail-body">{{ dialogMail.body }}</pre>
          <p v-else class="hint dialog-mail-empty-body">目前沒有可顯示的 body。</p>
        </div>

        <div class="dialog-mail-attachments">
          <div class="attachment-header">
            <span class="attachment-header-title">
              <span>附件</span>
              <el-tag effect="plain">{{ dialogMailAttachments.length }}</el-tag>
            </span>
          </div>
          <div v-if="isAttachmentListLoading(dialogMail)" class="dialog-loading-block compact">
            <el-skeleton animated :rows="2" />
          </div>
          <p v-else-if="dialogMailAttachments.length === 0" class="hint">這封郵件沒有附件。</p>
          <div v-else class="attachment-list">
            <div v-for="attachment in dialogMailAttachments" :key="attachment.attachmentId" class="attachment-row">
              <span class="attachment-main">
                <strong>{{ attachment.name }}</strong>
                <span>{{ formatAttachmentMeta(attachment.contentType, attachment.size) }}</span>
              </span>
              <span class="attachment-actions">
                <el-button
                  size="small"
                  :loading="isAttachmentExporting(dialogMail, attachment)"
                  :disabled="isAttachmentExporting(dialogMail, attachment)"
                  @click="exportMailAttachment(dialogMail, attachment)"
                >
                  {{ attachment.isExported ? '重新匯出' : 'Export' }}
                </el-button>
                <el-button
                  size="small"
                  :disabled="!attachment.exportedAttachmentId"
                  @click="openExportedAttachment(attachment)"
                >
                  開啟
                </el-button>
              </span>
            </div>
          </div>
        </div>

        <div class="dialog-mail-conversation">
          <div class="attachment-header">
            <span class="attachment-header-title">
              <span>討論串</span>
              <el-tag effect="plain">{{ dialogMailConversationItems.length }}</el-tag>
            </span>
            <small v-if="dialogMailConversation?.conversationTopic">{{ dialogMailConversation.conversationTopic }}</small>
          </div>
          <div v-if="isConversationLoading(dialogMail)" class="dialog-loading-block compact">
            <el-skeleton animated :rows="2" />
          </div>
          <p v-else-if="dialogMailConversationItems.length === 0" class="hint">目前沒有可顯示的討論串。</p>
          <div v-else class="conversation-list">
            <article
              v-for="item in dialogMailConversationItems"
              :key="item.id"
              class="conversation-item"
              :class="{ current: item.id === dialogMail.id }"
            >
              <span class="conversation-marker" />
              <span class="conversation-main">
                <strong>{{ item.subject || '(No subject)' }}</strong>
                <span>{{ formatMailSender(item) }} · {{ formatDateTime(item.receivedTime) }} · {{ item.folderPath }}</span>
                <small v-if="mailPreviewText(item.body, item.bodyHtml)">{{ mailPreviewText(item.body, item.bodyHtml) }}</small>
              </span>
            </article>
          </div>
        </div>
      </div>
      <MailPropertyPane
        v-model:draft="mailPropertiesDraft"
        embedded
        :categories="categories"
        :category-tag-style="categoryTagStyle"
        :flag-interval-options="flagIntervalOptions"
        :loading-categories="loadingCategories"
        :mail-properties-changed="mailPropertiesChanged"
        :operation-loading="operationLoading"
        :outlook-busy="outlookBusy"
        :selected-mail="dialogMail"
        :selected-mail-folder-name="dialogMailFolderName"
        :selected-mail-has-identity="dialogMailHasIdentity"
        @add-category="addMailCategoryDraft"
        @apply="applyMailProperties"
        @open-category-manager="openCategoryManager"
        @remove-category="removeMailCategoryDraft"
        @reset="resetMailPropertiesDraft"
        @set-flag="setMailFlagDraft"
      />
    </div>

    <el-dialog
      v-model="groupTreeVisible"
      class="recipient-tree-dialog"
      width="min(620px, calc(100vw - 36px))"
      append-to-body
      destroy-on-close
      title="Group 收件者"
    >
      <div class="recipient-tree-summary">
        <strong>{{ groupTreeRoot?.label || 'Group' }}</strong>
        <span v-if="groupTreeRoot?.recipientRelevance">{{ groupTreeRoot.recipientRelevance.summary }}</span>
        <el-tag v-if="groupTreeRoot?.recipientRelevance" effect="plain">
          {{ relevanceLabel(groupTreeRoot.recipientRelevance) }}
        </el-tag>
        <small v-if="groupTreeMessage">{{ groupTreeMessage }}</small>
      </div>
      <el-tree
        ref="groupTreeRef"
        class="recipient-tree"
        :data="groupTreeNodes"
        :props="groupTreeProps"
        node-key="key"
        :default-expanded-keys="groupTreeExpandedKeys"
      >
        <template #default="{ data, node }">
          <div class="recipient-tree-node">
            <button
              v-if="data.isGroup"
              class="recipient-tree-toggle"
              :class="{ expanded: node.expanded }"
              type="button"
              :aria-label="node.expanded ? '收合 group' : '展開 group'"
              :disabled="data.loading || (!data.loaded && groupTreeLoading)"
              @click.stop="toggleGroupTreeNode(data, node)"
            >
              <span />
            </button>
            <span v-else class="recipient-tree-toggle-placeholder" />
            <span>
              <strong>{{ data.label }}</strong>
              <small v-if="data.email">{{ data.email }}</small>
            </span>
            <el-tag v-if="data.isGroup" size="small" effect="plain">Group</el-tag>
          </div>
        </template>
      </el-tree>
    </el-dialog>
  </el-dialog>
</template>
