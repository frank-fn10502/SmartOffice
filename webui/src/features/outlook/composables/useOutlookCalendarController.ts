import { computed, ref } from 'vue'
import type { ComputedRef } from 'vue'
import type { Ref } from 'vue'
import { outlookApi } from '../api/outlook'
import type { CalendarEventDto } from '../models/outlook'
import {
  addMonths,
  buildCalendarWeeks,
  monthEndExclusive,
  monthStart,
  toDateKey,
} from '../utils/outlookDashboardHelpers'

type CalendarControllerOptions = {
  loadingCalendar: Ref<boolean>
  outlookBusy: ComputedRef<boolean>
  waitForRequest: (response: { requestId?: string; request?: string }, timeoutMs?: number) => Promise<void>
}

export function useOutlookCalendarController(options: CalendarControllerOptions) {
  const { loadingCalendar, outlookBusy, waitForRequest } = options
  const calendarEvents = ref<CalendarEventDto[]>([])
  const calendarMonthDate = ref(monthStart(new Date()))
  const selectedCalendarEvent = ref<CalendarEventDto | null>(null)
  const calendarWeekdays = ['日', '一', '二', '三', '四', '五', '六']

  const calendarMonthLabel = computed(() => {
    return calendarMonthDate.value.toLocaleDateString('zh-TW', { year: 'numeric', month: 'long' })
  })

  const calendarWeeks = computed(() => buildCalendarWeeks(calendarMonthDate.value, calendarEvents.value))

  async function loadCachedCalendar() {
    calendarEvents.value = await outlookApi.getCalendar()
  }

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
      await loadCachedCalendar()
      loadingCalendar.value = false
    } catch {
      loadingCalendar.value = false
    }
  }

  async function changeCalendarMonth(offset: number) {
    if (outlookBusy.value) return
    calendarMonthDate.value = addMonths(calendarMonthDate.value, offset)
    selectedCalendarEvent.value = null
    await requestCalendar()
  }

  async function goToCurrentCalendarMonth() {
    if (outlookBusy.value) return
    calendarMonthDate.value = monthStart(new Date())
    selectedCalendarEvent.value = null
    await requestCalendar()
  }

  function selectCalendarEvent(event: CalendarEventDto) {
    selectedCalendarEvent.value = event
  }

  return {
    calendarEvents,
    calendarMonthLabel,
    calendarWeekdays,
    calendarWeeks,
    changeCalendarMonth,
    goToCurrentCalendarMonth,
    loadCachedCalendar,
    loadingCalendar,
    requestCalendar,
    selectCalendarEvent,
    selectedCalendarEvent,
  }
}
