import type { CalendarEventDto, OutlookStoreDto } from '../models/outlook'

export const flagIntervalOptions = [
  { label: '不設定旗標', value: 'none' },
  { label: '今天', value: 'today' },
  { label: '明天', value: 'tomorrow' },
  { label: '本週', value: 'this_week' },
  { label: '下週', value: 'next_week' },
  { label: '無日期', value: 'no_date' },
  { label: '自訂日期', value: 'custom' },
  { label: '標示完成', value: 'complete' },
]

export function defaultFlagRequest(value: string) {
  return value === 'none' ? '' : flagIntervalLabel(value)
}

export function flagIntervalLabel(value?: string) {
  return flagIntervalOptions.find((option) => option.value === value)?.label ?? '旗標'
}

export function flagDisplayLabel(interval?: string, request?: string) {
  const label = flagIntervalLabel(interval)
  const trimmed = request?.trim() ?? ''
  if (interval !== 'custom') return label
  return trimmed && trimmed !== label ? `${label}(${trimmed})` : label
}

export function flagTagType(interval?: string, active = true) {
  if (!active || interval === 'none') return 'info'
  return interval === 'complete' ? 'success' : 'danger'
}

export function isDefaultFlagRequest(value: string, previousInterval = '') {
  const normalized = value.trim()
  return (
    !normalized
    || normalized === 'Follow up'
    || normalized === '旗標'
    || normalized === flagIntervalLabel(previousInterval)
    || flagIntervalOptions.some((option) => option.label === normalized)
  )
}

export function toDateInput(value?: string) {
  if (!value) return ''
  if (/^\d{4}-\d{2}-\d{2}/.test(value)) return value.slice(0, 10)
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return ''
  return date.toISOString().slice(0, 10)
}

export function dateInputToIso(value: string) {
  return value ? `${value}T00:00:00` : undefined
}

export function todayInputValue() {
  const now = new Date()
  const year = now.getFullYear()
  const month = `${now.getMonth() + 1}`.padStart(2, '0')
  const day = `${now.getDate()}`.padStart(2, '0')
  return `${year}-${month}-${day}`
}

export function sleep(ms: number) {
  return new Promise((resolve) => window.setTimeout(resolve, ms))
}

export function toDateKey(date: Date) {
  const year = date.getFullYear()
  const month = `${date.getMonth() + 1}`.padStart(2, '0')
  const day = `${date.getDate()}`.padStart(2, '0')
  return `${year}-${month}-${day}`
}

export function monthStart(date: Date) {
  return new Date(date.getFullYear(), date.getMonth(), 1)
}

export function monthEndExclusive(date: Date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 1)
}

export function addMonths(date: Date, count: number) {
  return new Date(date.getFullYear(), date.getMonth() + count, 1)
}

export function calendarEventSegment(event: CalendarEventDto, weekStart: Date, weekEnd: Date) {
  const start = new Date(event.start)
  const end = new Date(event.end)
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return null
  const eventEnd = new Date(end)
  if (eventEnd.getTime() > start.getTime()) eventEnd.setMilliseconds(eventEnd.getMilliseconds() - 1)
  const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate())
  const endDay = new Date(eventEnd.getFullYear(), eventEnd.getMonth(), eventEnd.getDate())
  if (endDay < weekStart || startDay > weekEnd) return null

  const segmentStart = startDay < weekStart ? weekStart : startDay
  const segmentEnd = endDay > weekEnd ? weekEnd : endDay
  const startColumn = Math.floor((segmentStart.getTime() - weekStart.getTime()) / 86400000) + 1
  const span = Math.floor((segmentEnd.getTime() - segmentStart.getTime()) / 86400000) + 1

  return {
    event,
    startColumn,
    span,
    isStart: startDay >= weekStart,
    isEnd: endDay <= weekEnd,
    isMultiDay: endDay.getTime() > startDay.getTime(),
  }
}

export function buildCalendarWeeks(calendarMonthDate: Date, calendarEvents: CalendarEventDto[]) {
  const first = monthStart(calendarMonthDate)
  const gridStart = new Date(first)
  gridStart.setDate(first.getDate() - first.getDay())
  const todayKey = toDateKey(new Date())

  return Array.from({ length: 6 }, (_, weekIndex) => {
    const weekStart = new Date(gridStart)
    weekStart.setDate(gridStart.getDate() + weekIndex * 7)
    const weekEnd = new Date(weekStart)
    weekEnd.setDate(weekStart.getDate() + 6)
    const days = Array.from({ length: 7 }, (_, dayIndex) => {
      const date = new Date(gridStart)
      date.setDate(gridStart.getDate() + weekIndex * 7 + dayIndex)
      const key = toDateKey(date)
      return {
        key,
        date,
        dayNumber: date.getDate(),
        inMonth: date.getMonth() === calendarMonthDate.getMonth(),
        isToday: key === todayKey,
      }
    })

    const segments = calendarEvents
      .map((event) => calendarEventSegment(event, weekStart, weekEnd))
      .filter((segment): segment is NonNullable<typeof segment> => Boolean(segment))
      .sort((a, b) => new Date(a.event.start).getTime() - new Date(b.event.start).getTime())

    return {
      key: days.map((day) => day.key).join('-'),
      days,
      segments,
    }
  })
}

export function splitCategories(value: string) {
  return value
    .split(',')
    .map((category) => category.trim())
    .filter(Boolean)
}

export function mergeStores(current: OutlookStoreDto[], incoming: OutlookStoreDto[]) {
  const stores = [...current]
  for (const next of incoming) {
    const index = stores.findIndex((store) => store.storeId === next.storeId)
    if (index < 0) stores.push(next)
    else stores[index] = next
  }
  return stores
}
