<script setup lang="ts">
import { ArrowLeft, ArrowRight, Calendar, Delete, Edit, Plus, Refresh } from '@element-plus/icons-vue'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'
import { formatDateTime, formatTime } from '../utils/formatters'
import { formatRecipient, formatRecipients } from '../utils/mailAddresses'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const {
  beginCreateCalendarEvent,
  beginEditCalendarEvent,
  calendarDraft,
  calendarEditorMode,
  calendarEditorVisible,
  calendarEventDialogVisible,
  calendarEvents,
  calendarMonthLabel,
  calendarRooms,
  calendarWeekdays,
  calendarWeeks,
  changeCalendarMonth,
  deleteCalendarEvent,
  goToCurrentCalendarMonth,
  loadingCalendar,
  outlookBusy,
  requestCalendar,
  saveCalendarEvent,
  selectCalendarEvent,
  selectedCalendarEvent,
  setCalendarRoom,
} = props.dashboard
</script>

<template>
  <main class="calendar-layout">
    <section class="panel">
      <div class="panel-header">
        <div class="panel-title">
          <el-icon><Calendar /></el-icon>
          <span>月曆</span>
          <el-tag effect="plain">{{ calendarEvents.length }}</el-tag>
        </div>
        <div class="calendar-actions">
          <el-button :icon="ArrowLeft" :disabled="outlookBusy" @click="changeCalendarMonth(-1)" />
          <strong>{{ calendarMonthLabel }}</strong>
          <el-button :disabled="outlookBusy" @click="goToCurrentCalendarMonth">本月</el-button>
          <el-button :icon="ArrowRight" :disabled="outlookBusy" @click="changeCalendarMonth(1)" />
          <el-button :icon="Refresh" :loading="loadingCalendar" :disabled="outlookBusy && !loadingCalendar" @click="requestCalendar">
            同步整月
          </el-button>
          <el-button type="primary" :icon="Plus" :disabled="outlookBusy" @click="beginCreateCalendarEvent">
            新增
          </el-button>
        </div>
      </div>

      <div class="calendar-page">
        <div class="calendar-grid">
          <div v-for="day in calendarWeekdays" :key="day" class="calendar-weekday">{{ day }}</div>
          <div v-for="week in calendarWeeks" :key="week.key" class="calendar-week-row">
            <div class="calendar-week-days">
              <div
                v-for="day in week.days"
                :key="day.key"
                class="calendar-day"
                :class="{ muted: !day.inMonth, today: day.isToday }"
              >
                <div class="calendar-day-head">
                  <span class="calendar-day-number">{{ day.dayNumber }}</span>
                  <span v-if="day.eventCount > 1" class="calendar-day-count">{{ day.eventCount }} 項</span>
                </div>
              </div>
            </div>
            <div class="calendar-week-events">
              <button
                v-for="segment in week.segments"
                :key="`${segment.event.id || segment.event.start}-${segment.startColumn}`"
                class="calendar-event"
                :class="{ continued: segment.isMultiDay, 'continues-before': !segment.isStart, 'continues-after': !segment.isEnd }"
                type="button"
                :style="{ gridColumn: `${segment.startColumn} / span ${segment.span}` }"
                @click="selectCalendarEvent(segment.event)"
              >
                <span>{{ segment.isMultiDay ? `${formatDateTime(segment.event.start)} - ${formatDateTime(segment.event.end)}` : formatTime(segment.event.start) }}</span>
                <strong>{{ segment.event.subject }}</strong>
              </button>
            </div>
          </div>
        </div>

      </div>
    </section>

    <el-dialog
      v-model="calendarEventDialogVisible"
      width="min(560px, calc(100vw - 28px))"
      class="calendar-event-dialog"
      append-to-body
    >
      <template #header>
        <div v-if="selectedCalendarEvent" class="calendar-dialog-title">
          <span>Calendar Event</span>
          <strong>{{ selectedCalendarEvent.subject }}</strong>
        </div>
      </template>

      <div v-if="selectedCalendarEvent" class="calendar-dialog-content">
        <div class="calendar-dialog-time">
          {{ formatDateTime(selectedCalendarEvent.start) }} - {{ formatDateTime(selectedCalendarEvent.end) }}
        </div>
        <div class="rule-detail">
          <span>地點：{{ selectedCalendarEvent.location || '-' }}</span>
          <span>召集人：{{ formatRecipient(selectedCalendarEvent.organizer, '-') }}</span>
          <span>出席者：{{ formatRecipients(selectedCalendarEvent.requiredAttendees) || '-' }}</span>
        </div>
        <div class="marker-tags">
          <el-tag effect="plain">{{ selectedCalendarEvent.busyStatus || 'unknown' }}</el-tag>
          <el-tag :type="selectedCalendarEvent.smartOfficeOwned ? 'success' : 'info'" effect="plain">
            {{ selectedCalendarEvent.smartOfficeOwned ? 'SmartOffice event' : 'Outlook event' }}
          </el-tag>
          <el-tag v-if="selectedCalendarEvent.isRecurring" type="warning" effect="plain">週期性</el-tag>
        </div>
        <div class="calendar-dialog-actions">
          <el-button
            :icon="Edit"
            :disabled="outlookBusy || !selectedCalendarEvent.smartOfficeOwned"
            @click="beginEditCalendarEvent(selectedCalendarEvent)"
          >
            編輯
          </el-button>
          <el-button
            type="danger"
            :icon="Delete"
            :disabled="outlookBusy || !selectedCalendarEvent.smartOfficeOwned"
            @click="deleteCalendarEvent(selectedCalendarEvent)"
          >
            刪除
          </el-button>
        </div>
      </div>
    </el-dialog>

    <el-dialog
      v-model="calendarEditorVisible"
      width="min(560px, calc(100vw - 28px))"
      class="calendar-event-dialog"
      append-to-body
    >
      <template #header>
        <div class="calendar-dialog-title">
          <span>{{ calendarEditorMode === 'create' ? 'Create Event' : 'Edit Event' }}</span>
          <strong>{{ calendarDraft.subject || 'SmartOffice calendar event' }}</strong>
        </div>
      </template>

      <el-form label-position="top" class="calendar-editor-form">
        <el-form-item label="Subject">
          <el-input v-model="calendarDraft.subject" :disabled="outlookBusy" />
        </el-form-item>
        <el-form-item label="Time">
          <div class="calendar-editor-time">
            <el-date-picker v-model="calendarDraft.start" type="datetime" :disabled="outlookBusy" />
            <el-date-picker v-model="calendarDraft.end" type="datetime" :disabled="outlookBusy" />
          </div>
        </el-form-item>
        <el-form-item label="地點">
          <el-input v-model="calendarDraft.location" :disabled="outlookBusy" />
        </el-form-item>
        <el-form-item label="會議室 / Resource">
          <el-select
            v-model="calendarDraft.roomId"
            :disabled="outlookBusy"
            filterable
            clearable
            placeholder="選擇會議室"
            @change="setCalendarRoom"
          >
            <el-option
              v-for="room in calendarRooms"
              :key="room.id"
              :label="room.displayName"
              :value="room.id"
            >
              <span>{{ room.displayName }}</span>
              <small class="calendar-room-option">{{ room.smtpAddress || room.rawAddress }}</small>
            </el-option>
          </el-select>
        </el-form-item>
        <el-form-item label="內容">
          <el-input v-model="calendarDraft.body" type="textarea" :rows="4" :disabled="outlookBusy" />
        </el-form-item>
      </el-form>

      <template #footer>
        <el-button :disabled="outlookBusy" @click="calendarEditorVisible = false">取消</el-button>
        <el-button type="primary" :loading="loadingCalendar" :disabled="outlookBusy && !loadingCalendar" @click="saveCalendarEvent">
          儲存
        </el-button>
      </template>
    </el-dialog>
  </main>
</template>
