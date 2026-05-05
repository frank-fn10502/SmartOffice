import type { OutlookStoreDto } from '../models/outlook'

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
