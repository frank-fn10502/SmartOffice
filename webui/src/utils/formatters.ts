export function formatTime(value?: string) {
  if (!value) return '-'
  return new Date(value).toLocaleTimeString()
}

export function formatDateTime(value: string) {
  return new Date(value).toLocaleString()
}
