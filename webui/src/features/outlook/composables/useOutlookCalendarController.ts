import { ElMessage, ElMessageBox } from 'element-plus'
import { computed, ref } from 'vue'
import type { ComputedRef, Ref } from 'vue'
import { outlookApi } from '../api/outlook'
import type { AddressBookContactDto, CalendarEventDto, CalendarRoomDto, OutlookRecipientDto } from '../models/outlook'
import {
  addMonths,
  buildCalendarWeeks,
  monthEndExclusive,
  monthStart,
  toDateKey,
} from '../utils/outlookDashboardHelpers'

type CalendarControllerOptions = {
  loadingCalendar: Ref<boolean>
  loadCalendarFromRequest: (response: { requestId?: string; request?: string }) => Promise<void>
  loadCalendarAddressBookFromRequest: (response: { requestId?: string; request?: string }) => Promise<AddressBookContactDto[]>
  loadCalendarRoomsFromRequest: (response: { requestId?: string; request?: string }) => Promise<CalendarRoomDto[]>
  outlookBusy: ComputedRef<boolean>
  waitForRequest: (response: { requestId?: string; request?: string }, timeoutMs?: number) => Promise<void>
}

type CalendarAttendeeOption = {
  value: string
  label: string
  meta: string
  contact: AddressBookContactDto
}

type CalendarMergeHint = {
  groupLabel: string
  coveredLabels: string[]
}

type CalendarDraft = {
  subject: string
  start: Date
  end: Date
  location: string
  roomId: string
  room: string
  roomAddress: string
  attendeeKeys: string[]
  body: string
  busyStatus: string
}

export function useOutlookCalendarController(options: CalendarControllerOptions) {
  const { loadCalendarAddressBookFromRequest, loadCalendarFromRequest, loadCalendarRoomsFromRequest, loadingCalendar, outlookBusy, waitForRequest } = options
  const calendarEvents = ref<CalendarEventDto[]>([])
  const calendarAddressBook = ref<AddressBookContactDto[]>([])
  const calendarRooms = ref<CalendarRoomDto[]>([])
  const calendarMonthDate = ref(monthStart(new Date()))
  const selectedCalendarEvent = ref<CalendarEventDto | null>(null)
  const calendarEventDialogVisible = ref(false)
  const calendarEditorVisible = ref(false)
  const calendarEditorMode = ref<'create' | 'update'>('create')
  const calendarDraft = ref<CalendarDraft>(defaultCalendarDraft())
  const calendarWeekdays = ['日', '一', '二', '三', '四', '五', '六']

  const calendarMonthLabel = computed(() => {
    return calendarMonthDate.value.toLocaleDateString('zh-TW', { year: 'numeric', month: 'long' })
  })

  const calendarAttendeeOptions = computed(() => {
    return calendarAddressBook.value
      .filter((contact) => Boolean(contact.displayName || contact.smtpAddress || contact.rawAddress))
      .map((contact) => ({
        value: contactKey(contact),
        label: contact.displayName || contact.smtpAddress || contact.rawAddress,
        meta: attendeeMeta(contact),
        contact,
      }))
      .sort((left, right) => left.label.localeCompare(right.label, 'zh-TW'))
  })

  const calendarMergeHints = computed<CalendarMergeHint[]>(() => {
    const selected = new Set(calendarDraft.value.attendeeKeys.map(normalizeKey))
    const selectedOptions = calendarAttendeeOptions.value.filter((option) => selected.has(normalizeKey(option.value)))
    return selectedOptions
      .filter((option) => option.contact.isGroup)
      .map((group) => {
        const memberKeys = new Set(group.contact.memberSmtpAddresses.concat(group.contact.memberGroupSmtpAddresses).map(normalizeKey))
        const coveredLabels = selectedOptions
          .filter((option) => option.value !== group.value && memberKeys.has(normalizeKey(option.value)))
          .map((option) => option.label)
        return { groupLabel: group.label, coveredLabels }
      })
      .filter((hint) => hint.coveredLabels.length > 0)
  })

  const calendarWeeks = computed(() => buildCalendarWeeks(calendarMonthDate.value, calendarEvents.value))

  async function requestCalendar() {
    if (outlookBusy.value) return
    loadingCalendar.value = true
    try {
      const start = monthStart(calendarMonthDate.value)
      const end = monthEndExclusive(calendarMonthDate.value)
      const response = await outlookApi.requestCalendar({
        daysForward: Math.ceil((end.getTime() - start.getTime()) / 86400000),
        startDate: toDateKey(start),
        endDate: toDateKey(end),
      })
      await waitForRequest(response)
      await loadCalendarFromRequest(response)
      loadingCalendar.value = false
    } catch {
      loadingCalendar.value = false
    }
  }

  async function changeCalendarMonth(offset: number) {
    if (outlookBusy.value) return
    calendarMonthDate.value = addMonths(calendarMonthDate.value, offset)
    selectedCalendarEvent.value = null
    calendarEventDialogVisible.value = false
    await requestCalendar()
  }

  async function goToCurrentCalendarMonth() {
    if (outlookBusy.value) return
    calendarMonthDate.value = monthStart(new Date())
    selectedCalendarEvent.value = null
    calendarEventDialogVisible.value = false
    await requestCalendar()
  }

  function selectCalendarEvent(event: CalendarEventDto) {
    selectedCalendarEvent.value = event
    calendarEventDialogVisible.value = true
  }

  function beginCreateCalendarEvent() {
    calendarEditorMode.value = 'create'
    calendarDraft.value = defaultCalendarDraft()
    calendarEditorVisible.value = true
    void requestCalendarRooms()
    void requestCalendarAddressBook()
  }

  function beginEditCalendarEvent(event: CalendarEventDto) {
    if (!event.smartOfficeOwned) {
      ElMessage.warning('只能編輯 SmartOffice 建立的 calendar event。')
      return
    }
    calendarEditorMode.value = 'update'
    calendarDraft.value = {
      subject: event.subject,
      start: new Date(event.start),
      end: new Date(event.end),
      location: event.location,
      roomId: '',
      room: '',
      roomAddress: '',
      attendeeKeys: formatAttendeeKeys(event.requiredAttendees),
      body: '',
      busyStatus: normalizeBusyStatus(event.busyStatus),
    }
    calendarEditorVisible.value = true
    void requestCalendarRooms()
    void requestCalendarAddressBook()
  }

  async function requestCalendarRooms() {
    if (outlookBusy.value || calendarRooms.value.length > 0) return
    const response = await outlookApi.requestCalendarRooms()
    await waitForRequest(response)
    calendarRooms.value = await loadCalendarRoomsFromRequest(response)
  }

  async function requestCalendarAddressBook() {
    if (outlookBusy.value || calendarAddressBook.value.length > 0) return
    const response = await outlookApi.requestAddressBook({
      includeOutlookContacts: true,
      includeAddressLists: true,
      maxContacts: 1000,
      maxAddressEntriesPerList: 500,
      maxGroupMembers: 50,
      maxGroupDepth: 1,
    })
    await waitForRequest(response, 120000)
    calendarAddressBook.value = await loadCalendarAddressBookFromRequest(response)
  }

  function setCalendarRoom(roomId: string) {
    const room = calendarRooms.value.find((item) => item.id === roomId)
    calendarDraft.value.roomId = roomId
    calendarDraft.value.room = room?.displayName ?? ''
    calendarDraft.value.roomAddress = room?.smtpAddress || room?.rawAddress || ''
  }

  async function saveCalendarEvent() {
    if (outlookBusy.value) return
    const selected = selectedCalendarEvent.value
    const draft = calendarDraft.value
    if (!draft.subject.trim() || draft.start >= draft.end) {
      ElMessage.warning('請輸入標題，且結束時間必須晚於開始時間。')
      return
    }

    loadingCalendar.value = true
    try {
      const body = {
        subject: draft.subject.trim(),
        start: draft.start.toISOString(),
        end: draft.end.toISOString(),
        location: draft.location.trim(),
        body: draft.body,
        busyStatus: 'busy',
        requiredAttendees: draft.attendeeKeys.map((key) => attendeeRecipient(key, calendarAttendeeOptions.value)),
        resources: draft.room.trim()
          ? [{ recipientKind: 'resource', displayName: draft.room.trim(), smtpAddress: draft.roomAddress, rawAddress: draft.roomAddress || draft.room.trim(), addressType: '', entryUserType: '', isGroup: false, isResolved: false, members: [] }]
          : [],
      }
      const response = calendarEditorMode.value === 'create'
        ? await outlookApi.requestCreateCalendarEvent(body)
        : await outlookApi.requestUpdateCalendarEvent({
          ...body,
          eventId: selected?.id ?? '',
          smartOfficeEventId: selected?.smartOfficeEventId ?? '',
        })
      await waitForRequest(response)
      await loadCalendarFromRequest(response)
      calendarEditorVisible.value = false
      selectedCalendarEvent.value = null
      calendarEventDialogVisible.value = false
    } finally {
      loadingCalendar.value = false
    }
  }

  async function deleteCalendarEvent(event: CalendarEventDto) {
    if (outlookBusy.value) return
    if (!event.smartOfficeOwned) {
      ElMessage.warning('只能刪除 SmartOffice 建立的 calendar event。')
      return
    }
    const confirmed = await ElMessageBox.confirm('刪除這個 SmartOffice calendar event？', '刪除 calendar event', {
      confirmButtonText: '刪除',
      cancelButtonText: '取消',
      type: 'warning',
    }).then(() => true).catch(() => false)
    if (!confirmed) return
    loadingCalendar.value = true
    try {
      const response = await outlookApi.requestDeleteCalendarEvent({ eventId: event.id })
      await waitForRequest(response)
      await loadCalendarFromRequest(response)
      selectedCalendarEvent.value = null
      calendarEventDialogVisible.value = false
    } finally {
      loadingCalendar.value = false
    }
  }

  return {
    beginCreateCalendarEvent,
    beginEditCalendarEvent,
    calendarAttendeeOptions,
    calendarMergeHints,
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
    requestCalendar,
    saveCalendarEvent,
    selectCalendarEvent,
    selectedCalendarEvent,
    setCalendarRoom,
  }
}

function defaultCalendarDraft(): CalendarDraft {
  const start = new Date()
  start.setMinutes(0, 0, 0)
  start.setHours(start.getHours() + 1)
  const end = new Date(start.getTime() + 60 * 60 * 1000)
  return { subject: '', start, end, location: '', roomId: '', room: '', roomAddress: '', attendeeKeys: [], body: '', busyStatus: 'busy' }
}

function normalizeBusyStatus(value: string) {
  const normalized = value.trim().toLowerCase()
  return ['free', 'tentative', 'busy', 'out_of_office'].includes(normalized) ? normalized : 'busy'
}

function contactKey(contact: AddressBookContactDto) {
  return contact.smtpAddress || contact.rawAddress || contact.displayName || contact.id
}

function normalizeKey(value: string) {
  return value.trim().toLowerCase()
}

function attendeeMeta(contact: AddressBookContactDto) {
  const parts = [
    contact.isGroup ? '群組' : '聯絡人',
    contact.smtpAddress || contact.rawAddress,
    contact.department,
  ].filter(Boolean)
  return parts.join(' / ')
}

function attendeeRecipient(value: string, options: CalendarAttendeeOption[]): OutlookRecipientDto {
  const contact = options.find((option) => option.value === value)?.contact
  if (contact) {
    return {
      recipientKind: 'required',
      displayName: contact.displayName || contact.smtpAddress || contact.rawAddress,
      smtpAddress: contact.smtpAddress,
      rawAddress: contact.rawAddress || contact.smtpAddress || contact.displayName,
      addressType: contact.addressType,
      entryUserType: contact.entryUserType,
      isGroup: contact.isGroup,
      isResolved: true,
      members: [],
    }
  }
  const trimmed = value.trim()
  return {
    recipientKind: 'required',
    displayName: trimmed,
    smtpAddress: trimmed.includes('@') ? trimmed : '',
    rawAddress: trimmed,
    addressType: trimmed.includes('@') ? 'SMTP' : '',
    entryUserType: '',
    isGroup: false,
    isResolved: false,
    members: [],
  }
}

function formatAttendeeKeys(attendees: CalendarEventDto['requiredAttendees']) {
  return attendees
    .filter((attendee) => attendee.recipientKind !== 'resource')
    .map((attendee) => attendee.smtpAddress || attendee.rawAddress || attendee.displayName)
    .filter(Boolean)
}
