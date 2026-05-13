<script setup lang="ts">
import type { CSSProperties } from 'vue'
import type { MailItemDto } from '../models/outlook'
import { formatDateTime } from '../utils/formatters'
import { formatMailSender } from '../utils/mailAddresses'
import { canMoveOutlookItem, outlookItemTypeLabel } from '../utils/outlookItemTypes'

defineProps<{
  mail: MailItemDto
  index: number
  selectedMailIds: Set<string>
  outlookBusy: boolean
  sourceLabel?: string
  categoryTagStyle: (name: string) => CSSProperties
  flagDisplayLabel: (interval: string, request: string) => string
  flagTagType: (interval: string) => string
  splitCategories: (categories: string) => string[]
}>()

defineEmits<{
  clearMailDrag: []
  openMailDialog: [index: number]
  selectMail: [index: number, event: MouseEvent]
  startMailPointerDrag: [mail: MailItemDto, index: number, event: PointerEvent]
}>()
</script>

<template>
  <article class="mail-card-row" :class="{ selected: selectedMailIds.has(mail.id), unread: !mail.isRead }">
    <div class="mail-row-shell">
      <button
        class="mail-row"
        type="button"
        :title="canMoveOutlookItem(mail) ? '按住拖曳到左側 folder 可移動；雙擊開啟郵件。' : '此 Outlook item 不能移動；雙擊可開啟郵件。'"
        @pointerdown="$emit('startMailPointerDrag', mail, index, $event)"
        @click="$emit('selectMail', index, $event)"
        @dblclick="$emit('openMailDialog', index)"
      >
        <span class="mail-row-head">
          <span class="mail-row-main">
            <strong>{{ mail.subject }}</strong>
            <span>{{ formatMailSender(mail) }} · {{ formatDateTime(mail.receivedTime) }}</span>
            <span v-if="sourceLabel" class="mail-source-label">{{ sourceLabel }}</span>
          </span>
          <el-tag v-if="mail.attachmentCount > 0" class="mail-attachment-tag" type="info" effect="plain">
            {{ mail.attachmentCount }} 個附件
          </el-tag>
        </span>
        <span class="mail-row-tags">
          <el-tag v-if="outlookItemTypeLabel(mail)" type="success" effect="plain">{{ outlookItemTypeLabel(mail) }}</el-tag>
          <el-tag v-if="!mail.isRead" type="warning" effect="plain">未讀</el-tag>
          <el-tag v-if="mail.isMarkedAsTask" :type="flagTagType(mail.flagInterval)" effect="plain">
            {{ flagDisplayLabel(mail.flagInterval, mail.flagRequest) }}<span v-if="mail.taskDueDate"> · {{ formatDateTime(mail.taskDueDate) }}</span>
          </el-tag>
          <el-tag
            v-for="category in splitCategories(mail.categories)"
            :key="category"
            effect="dark"
            :style="categoryTagStyle(category)"
          >
            {{ category }}
          </el-tag>
        </span>
      </button>
    </div>
  </article>
</template>
