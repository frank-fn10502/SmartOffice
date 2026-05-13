import { outlookApi } from '../api/outlook'

export function isRequestResponse(value: unknown): value is { requestId?: string; request?: string; state: string; data?: unknown } {
  const response = value as { requestId?: unknown; request?: unknown; state?: unknown; data?: unknown }
  return typeof response?.requestId === 'string'
    && typeof response?.request === 'string'
    && typeof response?.state === 'string'
    && response.data !== undefined
}

export function requestIdFromResponse(response: { requestId?: string }) {
  return response.requestId || ''
}

export function fetchResultEndpoint(response: { request?: string }) {
  const request = response.request || ''
  return request.startsWith('request-')
    ? request.replace('request-', 'fetch-result-')
    : 'fetch-result-mails'
}

export async function waitForOutlookRequest(
  response: { requestId?: string; request?: string },
  options: { timeoutMs?: number; isUnmounted?: () => boolean } = {},
) {
  const requestId = requestIdFromResponse(response)
  if (!requestId) return
  const endpoint = fetchResultEndpoint(response)
  const timeoutMs = options.timeoutMs ?? 120000
  const started = Date.now()
  while (!options.isUnmounted?.() && Date.now() - started < timeoutMs) {
    try {
      const state = await outlookApi.fetchResult(endpoint, {
        requestId,
        take: 1,
      })
      if (state.state === 'completed') return
      if (state.state && !['accepted', 'running'].includes(state.state)) {
        throw new Error(state.message || 'Outlook operation failed')
      }
    } catch (error) {
      if (error instanceof Error && error.message !== 'Request failed: 404') throw error
    }
    await new Promise((resolve) => window.setTimeout(resolve, 300))
  }
  throw new Error('Outlook operation timed out')
}
