import { computed, type Ref } from 'vue'
import { normalizeMailSearchProgress, outlookApi } from '../api/outlook'
import type { MailItemDto, MailSearchProgressDto, OutlookStoreDto } from '../models/outlook'
import { fetchResultEndpoint, requestIdFromResponse } from './outlookRequests'

type MailSearchDraft = {
  keyword: string
  textFields: Array<'subject' | 'sender' | 'body'>
  categoryNames: string[]
  hasAttachments: boolean | undefined
  flagState: 'any' | 'flagged' | 'unflagged'
  readState: 'any' | 'unread' | 'read'
  receivedFrom: string
  receivedTo: string
  scopeMode: 'selected_folder' | 'selected_store' | 'global'
}

type SearchSummaryItem = { label: string; value: string; tone: 'active' | 'muted' | 'info' }

type SearchControllerOptions = {
  activeMailSearchSummary: Ref<SearchSummaryItem[]>
  collapsedSearchResultFolders: Ref<Set<string>>
  collapsedSearchResultStores: Ref<Set<string>>
  folderOptions: Ref<Array<{ folderPath: string; storeId: string; label: string; name: string }>>
  folderStores: Ref<OutlookStoreDto[]>
  loadRequestMailItems: (response: { requestId?: string; request?: string; data?: unknown }) => Promise<MailItemDto[]>
  loadingMailSearch: Ref<boolean>
  mailListMode: Ref<'folder' | 'search'>
  mailSearchDraft: Ref<MailSearchDraft>
  mailSearchProgress: Ref<MailSearchProgressDto | null>
  mailSearchResults: Ref<MailItemDto[]>
  selectedFolderPath: Ref<string>
  selectedMailIndex: Ref<number | null>
  clearSelectedMails: () => void
  folderNameForPath: (path: string) => string
  waitForRequest: (response: { requestId?: string; request?: string }, timeoutMs?: number) => Promise<void>
}

export function useOutlookSearchController(options: SearchControllerOptions) {
  const {
    activeMailSearchSummary,
    collapsedSearchResultFolders,
    collapsedSearchResultStores,
    folderNameForPath,
    folderOptions,
    folderStores,
    loadRequestMailItems,
    loadingMailSearch,
    mailListMode,
    mailSearchDraft,
    mailSearchProgress,
    mailSearchResults,
    selectedFolderPath,
    selectedMailIndex,
    clearSelectedMails,
    waitForRequest,
  } = options

  function storeForFolderPath(path: string) {
    if (!path) return undefined
    return folderStores.value.find((store) => {
      const root = store.rootFolderPath
      return root && (path === root || path.startsWith(`${root}/`) || path.startsWith(`${root}\\`))
    })
  }

  function folderLeafName(path: string) {
    const parts = path.split(/[\\/]+/).map((part) => part.trim()).filter(Boolean)
    return parts.at(-1) || path || 'Unknown folder'
  }

  function searchStoreLabel(store: OutlookStoreDto | undefined, fallbackStoreId = '') {
    if (!store) return fallbackStoreId || 'Unknown store'
    const kind = store.storeKind?.trim().toUpperCase() || 'STORE'
    if (kind === 'PST') {
      const fileName = store.storeFilePath.split(/[\\/]+/).filter(Boolean).at(-1)
      return `PST · ${fileName || store.displayName || store.storeId}`
    }
    if (kind === 'OST') return `OST · ${store.displayName || store.storeId}`
    return `${kind} · ${store.displayName || store.storeId}`
  }

  function mailSource(mail: MailItemDto) {
    const folder = folderOptions.value.find((item) => item.folderPath === mail.folderPath)
    const store = folderStores.value.find((item) => item.storeId === folder?.storeId) ?? storeForFolderPath(mail.folderPath)
    return {
      storeId: store?.storeId || folder?.storeId || '',
      storeLabel: searchStoreLabel(store, folder?.storeId),
      folderLabel: folder?.name || folderLeafName(mail.folderPath),
      folderPath: mail.folderPath,
    }
  }

  function mailSourceLabel(mail: MailItemDto) {
    const source = mailSource(mail)
    return source.folderPath ? `${source.storeLabel} / ${source.folderLabel}` : source.storeLabel
  }

  function compareMailSearchResults(left: MailItemDto, right: MailItemDto) {
    const sourceOrder = mailSourceLabel(left).localeCompare(mailSourceLabel(right), undefined, { sensitivity: 'base' })
    if (sourceOrder !== 0) return sourceOrder
    return new Date(right.receivedTime).getTime() - new Date(left.receivedTime).getTime()
  }

  function setMailSearchResults(items: MailItemDto[]) {
    mailSearchResults.value = [...items].sort(compareMailSearchResults)
  }

  const searchResultRows = computed(() => mailSearchResults.value.map((mail, index) => ({
    mail,
    index,
    sourceLabel: mailSourceLabel(mail),
  })))

  const searchResultGroups = computed(() => {
    const groups = new Map<string, {
      key: string
      label: string
      count: number
      collapsed: boolean
      folders: { key: string; label: string; path: string; count: number; collapsed: boolean; rows: typeof searchResultRows.value }[]
    }>()
    for (const row of searchResultRows.value) {
      const source = mailSource(row.mail)
      const storeKey = source.storeId || source.storeLabel
      const folderKey = `${storeKey}\n${source.folderPath || source.folderLabel}`
      let store = groups.get(storeKey)
      if (!store) {
        store = { key: storeKey, label: source.storeLabel, count: 0, collapsed: collapsedSearchResultStores.value.has(storeKey), folders: [] }
        groups.set(storeKey, store)
      }
      let folder = store.folders.find((item) => item.key === folderKey)
      if (!folder) {
        folder = { key: folderKey, label: source.folderLabel, path: source.folderPath, count: 0, collapsed: collapsedSearchResultFolders.value.has(folderKey), rows: [] }
        store.folders.push(folder)
      }
      store.count += 1
      folder.count += 1
      folder.rows.push(row)
    }
    return [...groups.values()]
  })

  const mailSearchProgressText = computed(() => {
    const progress = mailSearchProgress.value
    if (!progress || !loadingMailSearch.value) return ''
    const scopeText = progress.totalFolders > 0
      ? `${progress.processedFolders}/${progress.totalFolders} folders`
      : progress.totalStores > 0
        ? `${progress.processedStores}/${progress.totalStores} stores`
        : '準備中'
    const current = progress.currentFolderPath ? ` · ${folderNameForPath(progress.currentFolderPath)}` : ''
    return `${progress.percent}% · ${scopeText}${current}`
  })

  const mailSearchSummaryItems = computed(() => activeMailSearchSummary.value)

  function toggleSearchResultStore(key: string) {
    const next = new Set(collapsedSearchResultStores.value)
    if (next.has(key)) next.delete(key)
    else next.add(key)
    collapsedSearchResultStores.value = next
  }

  function toggleSearchResultFolder(key: string) {
    const next = new Set(collapsedSearchResultFolders.value)
    if (next.has(key)) next.delete(key)
    else next.add(key)
    collapsedSearchResultFolders.value = next
  }

  function localDateTimeToIso(value: string) {
    return value ? new Date(value).toISOString() : undefined
  }

  function selectedStoreIdForSearch() {
    return folderOptions.value.find((folder) => folder.folderPath === selectedFolderPath.value)?.storeId ?? ''
  }

  function searchScopeLabel(scopeMode: MailSearchDraft['scopeMode'], storeId: string, scopeFolderPaths: string[]) {
    if (scopeMode === 'global') return '全部信箱'
    if (scopeMode === 'selected_folder') return scopeFolderPaths[0] ? `${folderNameForPath(scopeFolderPaths[0])} + 子資料夾` : '目前資料夾未選擇'
    return searchStoreLabel(folderStores.value.find((item) => item.storeId === storeId), storeId)
  }

  function searchDateLabel(value: string) {
    return value ? value.replace('T', ' ') : ''
  }

  function searchReceivedCondition(from: string, to: string) {
    if (from && to) return `${searchDateLabel(from)} <= 時間 <= ${searchDateLabel(to)}`
    if (from) return `時間 >= ${searchDateLabel(from)}`
    if (to) return `時間 <= ${searchDateLabel(to)}`
    return ''
  }

  function buildMailSearchSummary(storeId: string, scopeFolderPaths: string[]) {
    const draft = mailSearchDraft.value
    const keyword = draft.keyword.trim()
    const receivedCondition = searchReceivedCondition(draft.receivedFrom, draft.receivedTo)
    const fieldLabels: Record<string, string> = { subject: '標題', sender: '寄件者', body: '內容' }
    const textFields = draft.textFields.map(field => fieldLabels[field] ?? field).join('、') || '標題'
    const summary: SearchSummaryItem[] = [
      { label: '範圍', value: searchScopeLabel(draft.scopeMode, storeId, scopeFolderPaths), tone: 'info' },
      { label: '文字範圍', value: textFields, tone: 'info' },
    ]
    if (draft.categoryNames.length > 0) summary.push({ label: '分類', value: draft.categoryNames.join('、'), tone: 'info' })
    if (draft.hasAttachments !== undefined) summary.push({ label: '附件', value: draft.hasAttachments ? '包含附件' : '不含附件', tone: 'info' })
    if (draft.flagState !== 'any') summary.push({ label: '旗標', value: draft.flagState === 'flagged' ? '有旗標' : '無旗標', tone: 'info' })
    if (draft.readState !== 'any') summary.push({ label: '狀態', value: draft.readState === 'unread' ? '未讀' : '已讀', tone: 'info' })
    if (receivedCondition) summary.unshift({ label: '時間', value: receivedCondition, tone: 'active' })
    if (keyword) summary.unshift({ label: '文字', value: `包含 "${keyword}"`, tone: 'active' })
    else summary.push({ label: '文字', value: '未使用', tone: 'muted' })
    return summary
  }

  async function requestMailSearch() {
    if (loadingMailSearch.value) return
    const searchId = window.crypto?.randomUUID?.() ?? `${Date.now()}`
    const scopeFolderPaths = mailSearchDraft.value.scopeMode === 'selected_folder' && selectedFolderPath.value ? [selectedFolderPath.value] : []
    const storeId = mailSearchDraft.value.scopeMode === 'global' ? '' : selectedStoreIdForSearch()
    activeMailSearchSummary.value = buildMailSearchSummary(storeId, scopeFolderPaths)
    loadingMailSearch.value = true
    mailSearchProgress.value = null
    mailListMode.value = 'search'
    mailSearchResults.value = []
    collapsedSearchResultStores.value = new Set()
    collapsedSearchResultFolders.value = new Set()
    selectedMailIndex.value = null
    clearSelectedMails()
    try {
      const response = await outlookApi.requestMailSearch({
        searchId,
        storeId,
        scopeFolderPaths,
        allowGlobalScope: mailSearchDraft.value.scopeMode === 'global',
        includeSubFolders: true,
        keyword: mailSearchDraft.value.keyword,
        textFields: mailSearchDraft.value.textFields,
        categoryNames: mailSearchDraft.value.categoryNames,
        hasAttachments: mailSearchDraft.value.hasAttachments,
        flagState: mailSearchDraft.value.flagState,
        readState: mailSearchDraft.value.readState,
        receivedFrom: localDateTimeToIso(mailSearchDraft.value.receivedFrom),
        receivedTo: localDateTimeToIso(mailSearchDraft.value.receivedTo),
      })
      await waitForRequest(response)
      try {
        const result = await outlookApi.fetchResult<{ searchId?: string }>(fetchResultEndpoint(response), {
          requestId: requestIdFromResponse(response),
          take: 1,
        })
        mailSearchProgress.value = result.data?.searchId
          ? normalizeMailSearchProgress(await outlookApi.getMailSearchProgress(result.data.searchId))
          : null
      } catch {
        // Search progress 不是每個失敗路徑都一定會留下 snapshot。
      }
      setMailSearchResults(await loadRequestMailItems(response))
      loadingMailSearch.value = false
    } catch {
      loadingMailSearch.value = false
    }
  }

  return {
    mailSearchProgressText,
    mailSearchSummaryItems,
    requestMailSearch,
    searchResultGroups,
    searchResultRows,
    setMailSearchResults,
    toggleSearchResultFolder,
    toggleSearchResultStore,
  }
}
