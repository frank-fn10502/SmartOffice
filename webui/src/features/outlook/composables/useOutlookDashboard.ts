import { computed } from 'vue'
import {
  normalizeCalendarRooms,
  normalizeMailAttachments,
  normalizeMailItems,
  outlookApi,
} from '../api/outlook'
import type {
  CalendarEventDto,
  MailAttachmentsDto,
  MailConversationDto,
  MailItemDto,
  OutlookCategoryDto,
  OutlookRuleDto,
} from '../models/outlook'
import {
  categoryColorOptions,
  categoryColorStyle,
} from '../utils/categoryColors'
import { collectFolderOptions, visibleRootFolders } from '../utils/folders'
import {
  flagDisplayLabel,
  flagIntervalOptions,
  flagTagType,
  splitCategories,
} from '../utils/outlookDashboardHelpers'
import { canUpdateMailProperties } from '../utils/outlookItemTypes'
import { patchMailSnapshotList } from './outlookMailSnapshots'
import { collectOutlookRequestData, fetchResultEndpoint, isRequestResponse, requestIdFromResponse, waitForOutlookRequest } from './outlookRequests'
import { useOutlookCalendarController } from './useOutlookCalendarController'
import { useOutlookDashboardState } from './useOutlookDashboardState'
import { useOutlookFolderMutationsController } from './useOutlookFolderMutationsController'
import { useOutlookFolderSelectionController } from './useOutlookFolderSelectionController'
import { useOutlookMailDetailController } from './useOutlookMailDetailController'
import { useOutlookMailListController } from './useOutlookMailListController'
import { useOutlookMailPropertiesController } from './useOutlookMailPropertiesController'
import { useOutlookRulesController } from './useOutlookRulesController'
import { useOutlookSearchController } from './useOutlookSearchController'
import { useOutlookFoldersController } from './useOutlookFoldersController'
import { useOutlookShellController } from './useOutlookShellController'

const manualOutlookDeleteMessage = 'SmartOffice API 不會永久刪除 Outlook 郵件或 folder。此項目已在 Outlook 刪除資料夾內；若要永久刪除，請到 Outlook 手動操作。'
export function useOutlookDashboard() {
  const {
    activeMailSearchSummary, activeView, addinLogs, addinStatus, attachmentExportRootDraft,
    attachmentExportSettings, categories, chatMessages, chatPanelRef, chatText,
    collapsedSearchResultFolders, collapsedSearchResultStores, creatingFolderName,
    creatingFolderParentPath, draggedMailId, dragOverFolderPath, expandedFolders,
    exportingAttachmentIds, fetchedMailFolderPath, folderContextMenu, folderMails,
    folderStores, folders, lastMailFetchAt, loadingAttachmentMailIds, loadingCalendar,
    loadingCategories, loadingConversationMailIds, loadingFolders, loadingMailBodyIds,
    loadingMails, loadingMailSearch, loadingRules, loadingSignalRPing, mailAttachmentsByMailId,
    mailConversationsByMailId, mailCount, mailDialogHtml, mailDialogIndex, mailDialogMailId,
    mailDialogVisible, mailFetchCountdownTick, mailListMode, mailLookbackHours,
    mailSearchDraft, mailSearchProgress, mailSearchResults, outlookFirstLoadCompleted,
    pendingMailFolderPath, requestLoading, ruleDraft, rules, savingAttachmentExportSettings,
    scheduledMailFetchAt, searchResultViewMode, selectedFolderPath, selectedMailIds,
    selectedMailIndex, selectedRuleIndex, signalRState,
  } = useOutlookDashboardState()
  const mailHtmlSandbox = 'allow-same-origin allow-popups allow-popups-to-escape-sandbox'
  let unmounted = false
  let initialFoldersFetchCompleted = false
  let initialMailsFetchCompleted = false
  let initialCategoriesFetchCompleted = false
  let startupSyncStarted = false
  let activeRequestId = ''
  let requestTimeoutId = 0
  const mailBodyTimeoutIds = new Map<string, number>()
  const attachmentTimeoutIds = new Map<string, number>()

  const visibleFolders = computed(() => visibleRootFolders(folders.value))
  const mails = computed(() => mailListMode.value === 'search' ? mailSearchResults.value : folderMails.value)

  const mailStats = computed(() => ({
    unread: mails.value.filter((mail) => !mail.isRead).length,
    flagged: mails.value.filter((mail) => mail.isMarkedAsTask).length,
    highImportance: mails.value.filter((mail) => mail.importance === 'high').length,
    categorized: mails.value.filter((mail) => Boolean(mail.categories)).length,
  }))
  const outlookBusy = computed(() => {
    return loadingFolders.value || loadingMails.value || loadingRules.value || loadingCategories.value || loadingCalendar.value || requestLoading.value
  })

  const outlookBusyText = computed(() => {
    if (loadingFolders.value) return 'Outlook folder 同步中...'
    if (loadingMails.value) return 'Outlook 郵件抓取中...'
    if (loadingRules.value) return 'Outlook rule 同步中...'
    if (loadingCategories.value) return 'Outlook category 同步中...'
    if (loadingCalendar.value) return 'Outlook calendar 同步中...'
    if (requestLoading.value) return 'Outlook 操作執行中...'
    return ''
  })

  const outlookDependentViewsLocked = computed(() => !outlookFirstLoadCompleted.value)
  const navOptions = computed(() => [
    { label: 'Outlook', value: 'outlook' },
    { label: 'Search', value: 'search', disabled: outlookDependentViewsLocked.value },
    { label: 'Rules', value: 'rules', disabled: outlookDependentViewsLocked.value },
    { label: 'Chat', value: 'chat', disabled: outlookDependentViewsLocked.value },
    { label: 'Calendar', value: 'calendar', disabled: outlookDependentViewsLocked.value },
    { label: 'Contacts', value: 'contacts', disabled: outlookDependentViewsLocked.value },
    { label: 'Admin', value: 'admin' },
  ])

  const folderOptions = computed(() => collectFolderOptions(visibleFolders.value))
  const {
    contextFolderName,
    deletedFolderForPath,
    fetchedMailFolderName,
    folderNameForPath,
    inferMailFolderPath,
    isInDeletedFolder,
    selectedFolderName,
  } = useOutlookFolderSelectionController({
    fetchedMailFolderPath,
    folderContextMenu,
    folderOptions,
    mailListMode,
    mailSearchResults,
    selectedFolderPath,
  })
  const {
    beginCreateCalendarEvent,
    beginEditCalendarEvent,
    calendarAttendeeOptions,
    calendarDraft,
    calendarEditorMode,
    calendarEditorVisible,
    calendarEventDialogVisible,
    calendarEvents,
    calendarMergeHints,
    calendarMonthLabel,
    calendarRooms,
    calendarWeekdays,
    calendarWeeks,
    changeCalendarMonth,
    deleteCalendarEvent,
    goToCurrentCalendarMonth,
    requestCalendar,
    saveCalendarEvent,
    selectCalendarEvent,
    selectedCalendarEvent,
    setCalendarRoom,
  } = useOutlookCalendarController({
    loadCalendarFromRequest,
    loadCalendarRoomsFromRequest,
    loadingCalendar,
    outlookBusy,
    waitForRequest,
  })

  const mailListNeedsFetch = computed(() => {
    if (mailListMode.value === 'search') return false
    return Boolean(selectedFolderPath.value && selectedFolderPath.value !== fetchedMailFolderPath.value)
  })

  const mailFetchCountdownSeconds = computed(() => {
    if (!scheduledMailFetchAt.value) return 0
    return Math.max(0, Math.ceil((scheduledMailFetchAt.value - mailFetchCountdownTick.value) / 100) / 10)
  })

  const mailFetchCountdownText = computed(() => {
    if (!scheduledMailFetchAt.value) return ''
    return `${mailFetchCountdownSeconds.value.toFixed(1)} 秒後自動抓取`
  })

  const showMailFetchWarning = computed(() => {
    return mailListNeedsFetch.value
      && !mailFetchCountdownText.value
      && !loadingMails.value
  })

  const mailFetchStatusText = computed(() => {
    if (mailFetchCountdownText.value) return `已選取 ${selectedFolderName.value}，${mailFetchCountdownText.value}；按「立即抓取」可提早執行。`
    if (loadingMails.value && pendingMailFolderPath.value) return `正在抓取：${folderNameForPath(pendingMailFolderPath.value)}`
    if (lastMailFetchAt.value) return `上次抓取：${lastMailFetchAt.value.toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}`
    return '尚未抓取郵件。'
  })

  const selectedMail = computed(() => {
    if (selectedMailIndex.value === null) return null
    return mails.value[selectedMailIndex.value] ?? null
  })

  const selectedMailCategories = computed(() => {
    if (!selectedMail.value?.categories) return []
    return selectedMail.value.categories
      .split(',')
      .map((category) => category.trim())
      .filter(Boolean)
  })

  const dialogMail = computed(() => {
    if (mailDialogMailId.value) {
      const mail = mails.value.find((item) => item.id === mailDialogMailId.value)
        ?? folderMails.value.find((item) => item.id === mailDialogMailId.value)
        ?? mailSearchResults.value.find((item) => item.id === mailDialogMailId.value)
      if (mail) return mail
    }
    if (mailDialogIndex.value === null) return null
    return mails.value[mailDialogIndex.value] ?? null
  })

  const dialogMailAttachments = computed(() => {
    return dialogMail.value?.id ? mailAttachmentsByMailId.value[dialogMail.value.id] ?? [] : []
  })

  const dialogMailConversation = computed(() => {
    return dialogMail.value?.id ? mailConversationsByMailId.value[dialogMail.value.id] ?? null : null
  })

  const dialogMailConversationItems = computed(() => dialogMailConversation.value?.mails ?? [])

  const selectedRule = computed(() => {
    return selectedRuleIndex.value === null ? null : rules.value[selectedRuleIndex.value] ?? null
  })

  const ruleDraftIsEditing = computed(() => Boolean(ruleDraft.value.originalRuleName))

  const dialogLoading = computed(() => {
    const mail = dialogMail.value
    return Boolean(mail && (isMailBodyLoading(mail) || isAttachmentListLoading(mail) || isConversationLoading(mail)))
  })

  const dialogMailFolderName = computed(() => {
    return dialogMail.value?.folderPath ? folderNameForPath(dialogMail.value.folderPath) : '未選擇'
  })

  const dialogMailHasIdentity = computed(() => Boolean(dialogMail.value?.id?.trim()))

  const activeMailForProperties = computed(() => dialogMail.value ?? selectedMail.value)

  function updateOutlookFirstLoadCompleted() {
    outlookFirstLoadCompleted.value = initialFoldersFetchCompleted
      && initialMailsFetchCompleted
      && initialCategoriesFetchCompleted
  }

  function setMails(items: MailItemDto[], preferredMailId = selectedMail.value?.id ?? '') {
    const sortedItems = [...items].sort((left, right) =>
      new Date(right.receivedTime).getTime() - new Date(left.receivedTime).getTime()
    )
    folderMails.value = sortedItems
    mailListMode.value = 'folder'
    pruneSelectedMailIds(sortedItems)

    if (sortedItems.length === 0) {
      clearSelectedMailIndex()
      return
    }

    const nextIndex = preferredMailId ? sortedItems.findIndex((mail) => mail.id === preferredMailId) : -1
    selectedMailIndex.value = nextIndex >= 0 ? nextIndex : 0
  }

  function patchMailAttachments(payload: MailAttachmentsDto) {
    if (!payload.mailId) return
    mailAttachmentsByMailId.value = {
      ...mailAttachmentsByMailId.value,
      [payload.mailId]: payload.attachments,
    }
    completeAttachmentLoad(payload.mailId)
  }

  function patchMailConversation(payload: MailConversationDto) {
    if (!payload.mailId) return
    mailConversationsByMailId.value = {
      ...mailConversationsByMailId.value,
      [payload.mailId]: payload,
    }
    completeConversationLoad(payload.mailId)
  }

  function patchMailSnapshots(items: MailItemDto[]) {
    if (items.length === 0) return
    folderMails.value = patchMailSnapshotList(folderMails.value, items)
    mailSearchResults.value = patchMailSnapshotList(mailSearchResults.value, items)
  }

  function markInitialFoldersComplete() {
    initialFoldersFetchCompleted = true
    updateOutlookFirstLoadCompleted()
  }

  function markInitialMailsComplete() {
    initialMailsFetchCompleted = true
    updateOutlookFirstLoadCompleted()
  }

  async function loadRequestMailItems(response: { requestId?: string; request?: string; data?: unknown }) {
    const pages = await collectOutlookRequestData<{ mails?: unknown[] }>(response, { isUnmounted: () => unmounted })
    return normalizeMailItems(pages.flatMap((page) => page.data?.mails ?? []))
  }

  async function loadRulesFromRequest(response: { requestId?: string; request?: string; data?: unknown }) {
    const pages = await collectOutlookRequestData<{ rules?: OutlookRuleDto[] }>(response, { isUnmounted: () => unmounted })
    rules.value = pages.flatMap((page) => page.data?.rules ?? [])
  }

  async function loadCategoriesFromRequest(response: { requestId?: string; request?: string; data?: unknown }) {
    const pages = await collectOutlookRequestData<{ categories?: OutlookCategoryDto[] }>(response, { isUnmounted: () => unmounted })
    categories.value = pages.flatMap((page) => page.data?.categories ?? [])
  }

  async function loadCalendarFromRequest(response: { requestId?: string; request?: string; data?: unknown }) {
    const pages = await collectOutlookRequestData<{ calendarEvents?: CalendarEventDto[] }>(response, { isUnmounted: () => unmounted })
    calendarEvents.value = pages.flatMap((page) => page.data?.calendarEvents ?? [])
  }

  async function loadCalendarRoomsFromRequest(response: { requestId?: string; request?: string; data?: unknown }) {
    const pages = await collectOutlookRequestData<{ rooms?: unknown[] }>(response, { isUnmounted: () => unmounted })
    return normalizeCalendarRooms(pages.flatMap((page) => page.data?.rooms ?? []))
  }

  const {
    beginCreateFolder,
    cancelCreateFolder,
    cancelScheduledMailFetch,
    closeFolderContextMenu,
    createFolderFromContext,
    ensureStartupInboxFolderLoaded,
    fetchMailsFromContext,
    loadFoldersFromRequest,
    openFolderContextMenu,
    requestFolders,
    requestMails,
    scheduleMailFetch,
    selectFolder,
    selectInboxFolder,
    toggleFolder,
  } = useOutlookFoldersController({
    clearSelectedMailIndex: () => { selectedMailIndex.value = null },
    clearSelectedMails: () => { selectedMailIds.value = new Set() },
    creatingFolderName,
    creatingFolderParentPath,
    expandedFolders,
    fetchedMailFolderPath,
    folderContextMenu,
    folderMails,
    folderOptions,
    folders,
    folderStores,
    inferMailFolderPath,
    lastMailFetchAt,
    loadRequestMailItems,
    loadingFolders,
    loadingMails,
    mailCount,
    mailFetchCountdownTick,
    mailListMode,
    mailLookbackHours,
    markInitialFoldersComplete,
    markInitialMailsComplete,
    outlookBusy,
    pendingMailFolderPath,
    scheduledMailFetchAt,
    selectedFolderPath,
    selectedMailIndex,
    setMails,
    visibleFolders,
    waitForRequest,
  })

  const {
    createFolder,
    deleteFolderFromContext,
  } = useOutlookFolderMutationsController({
    cancelCreateFolder,
    closeFolderContextMenu,
    creatingFolderName,
    creatingFolderParentPath,
    folderContextMenu,
    folderNameForPath,
    folderOptions,
    isInDeletedFolder,
    loadFoldersFromRequest,
    manualOutlookDeleteMessage,
    outlookBusy,
    requestLoading,
    selectedFolderPath,
    waitForRequest,
  })

  const {
    deleteRule,
    editRule,
    requestRules,
    resetRuleDraft,
    saveRule,
    toggleRuleEnabled,
    upsertCategory,
  } = useOutlookRulesController({
    loadCategoriesFromRequest,
    loadRulesFromRequest,
    loadingRules,
    outlookBusy,
    ruleDraft,
    rules,
    runMailOperation,
    selectedRuleIndex,
    waitForRequest,
  })

  const {
    activeMailPropertySections,
    addCategoryToMasterList,
    addMailCategoryDraft,
    applyMailProperties,
    categoryCreateColor,
    categoryCreateDraft,
    categoryManagerVisible,
    categoryTagStyle,
    flagEditorVisible,
    hiddenMasterCategoryCount,
    mailPropertiesChanged,
    mailPropertiesDraft,
    masterCategoryListExpanded,
    openCategoryManager,
    removeMailCategoryDraft,
    resetMailPropertiesDraft,
    setMailFlagDraft,
    toggleMasterCategoryList,
    updateCategoryColor,
    visibleMasterCategories,
  } = useOutlookMailPropertiesController({
    activeMailForProperties,
    categories,
    loadCategoriesFromRequest,
    loadRequestMailItems,
    outlookBusy,
    patchMailSnapshots,
    runMailOperation,
    upsertCategory,
    waitForRequest,
  })

  function showFolderMails() {
    mailListMode.value = 'folder'
    clearSelectedMailIndex()
    if (mailListNeedsFetch.value && !outlookBusy.value) scheduleMailFetch()
  }

  function openMailDialog(index: number) {
    const mail = mails.value[index]
    if (!mail) return
    selectOnlyMail(index)
    mailDialogIndex.value = index
    mailDialogMailId.value = mail.id
    mailDialogHtml.value = false
    mailDialogVisible.value = true
    void requestMailBody(mail)
    void requestMailAttachments(mail)
    if (canUpdateMailProperties(mail)) void requestMailConversation(mail)
  }

  function closeMailDialog() {
    mailDialogVisible.value = false
    mailDialogMailId.value = ''
    mailDialogHtml.value = false
  }

  async function requestCategories(force = false) {
    if (outlookBusy.value && !force) return
    loadingCategories.value = true
    try {
      const response = await outlookApi.requestCategories()
      await waitForRequest(response)
      await loadCategoriesFromRequest(response)
      initialCategoriesFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingCategories.value = false
      requestLoading.value = false
    } catch {
      initialCategoriesFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingCategories.value = false
    }
  }

  async function requestSignalRPing() {
    if (loadingSignalRPing.value) return
    loadingSignalRPing.value = true
    try {
      const response = await outlookApi.requestSignalRPing()
      await waitForRequest(response)
    } finally {
      loadingSignalRPing.value = false
    }
  }

  async function runStartupOutlookSync() {
    if (startupSyncStarted) return
    if (unmounted) return
    startupSyncStarted = true
    await requestFolders(true)
    if (!initialFoldersFetchCompleted) {
      initialFoldersFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingFolders.value = false
    }
    if (unmounted) return
    await ensureStartupInboxFolderLoaded()
    if (unmounted) return
    selectInboxFolder()

    await requestCategories(true)

    if (unmounted) return
    if (!selectedFolderPath.value) selectInboxFolder()
    mailLookbackHours.value = 168
    mailCount.value = 30
    await requestMails(true)
  }

  async function runMailOperation(action: () => Promise<unknown>, afterSuccess?: (response?: unknown) => Promise<void>) {
    if (outlookBusy.value) return false
    window.clearTimeout(requestTimeoutId)
    activeRequestId = ''
    requestLoading.value = true
    try {
      const response = await action()
      if (isRequestResponse(response)) {
        activeRequestId = requestIdFromResponse(response)
        if (!['accepted'].includes(response.state)) {
          requestLoading.value = false
          activeRequestId = ''
          return false
        }
        await waitForRequest(response)
      }
      if (afterSuccess) await afterSuccess(response)
      completeRequest(activeRequestId)
      return true
    } catch {
      requestLoading.value = false
      activeRequestId = ''
      window.clearTimeout(requestTimeoutId)
      return false
    }
  }

  async function waitForRequest(response: { requestId?: string; request?: string }, timeoutMs = 120000) {
    await waitForOutlookRequest(response, { timeoutMs, isUnmounted: () => unmounted })
  }

  function completeRequest(requestId = '') {
    if (requestId && activeRequestId && requestId !== activeRequestId) return
    requestLoading.value = false
    activeRequestId = ''
    window.clearTimeout(requestTimeoutId)
  }

  function completeMailBodyLoad(mailId: string) {
    const next = new Set(loadingMailBodyIds.value)
    next.delete(mailId)
    loadingMailBodyIds.value = next
    const timeoutId = mailBodyTimeoutIds.get(mailId)
    if (timeoutId) window.clearTimeout(timeoutId)
    mailBodyTimeoutIds.delete(mailId)
  }

  function clearMailBodyLoads() {
    loadingMailBodyIds.value = new Set()
    for (const timeoutId of mailBodyTimeoutIds.values()) window.clearTimeout(timeoutId)
    mailBodyTimeoutIds.clear()
  }

  function attachmentKey(mailId: string, attachmentId: string) {
    return `${mailId}\n${attachmentId}`
  }

  function completeAttachmentLoad(mailId: string) {
    const next = new Set(loadingAttachmentMailIds.value)
    next.delete(mailId)
    loadingAttachmentMailIds.value = next
    const timeoutId = attachmentTimeoutIds.get(mailId)
    if (timeoutId) window.clearTimeout(timeoutId)
    attachmentTimeoutIds.delete(mailId)
  }

  function completeAttachmentExport(mailId: string, attachmentId: string) {
    const key = attachmentKey(mailId, attachmentId)
    const next = new Set(exportingAttachmentIds.value)
    next.delete(key)
    exportingAttachmentIds.value = next
    const timeoutId = attachmentTimeoutIds.get(key)
    if (timeoutId) window.clearTimeout(timeoutId)
    attachmentTimeoutIds.delete(key)
  }

  function clearAttachmentLoads() {
    loadingAttachmentMailIds.value = new Set()
    exportingAttachmentIds.value = new Set()
    for (const timeoutId of attachmentTimeoutIds.values()) window.clearTimeout(timeoutId)
    attachmentTimeoutIds.clear()
  }

  function completeConversationLoad(mailId: string) {
    const next = new Set(loadingConversationMailIds.value)
    next.delete(mailId)
    loadingConversationMailIds.value = next
  }

  function clearConversationLoads() {
    loadingConversationMailIds.value = new Set()
  }

  const {
    exportMailAttachment,
    isAttachmentExporting,
    isAttachmentListLoading,
    isConversationLoading,
    isMailBodyLoading,
    mailHasBody,
    openExportedAttachment,
    requestMailAttachments,
    requestMailBody,
    requestMailConversation,
  } = useOutlookMailDetailController({
    attachmentKey,
    completeAttachmentExport,
    completeAttachmentLoad,
    completeConversationLoad,
    completeMailBodyLoad,
    exportingAttachmentIds,
    loadRequestMailItems,
    loadingAttachmentMailIds,
    loadingConversationMailIds,
    loadingMailBodyIds,
    mailAttachmentsByMailId,
    mailConversationsByMailId,
    patchMailAttachments,
    patchMailConversation,
    patchMailSnapshots,
    waitForRequest,
  })

  const {
    clearMailDrag,
    clearSelectedMailIndex,
    clearSelectedMails,
    deleteMail,
    deleteSelectedMails,
    moveDraggedMail,
    pruneSelectedMailIds,
    selectMail,
    selectOnlyMail,
    setDragOverFolder,
    startMailPointerDrag,
  } = useOutlookMailListController({
    deletedFolderForPath,
    draggedMailId,
    dragOverFolderPath,
    folderMails,
    folders,
    folderOptions,
    isInDeletedFolder,
    mailListMode,
    mailSearchResults,
    mails,
    manualOutlookDeleteMessage,
    outlookBusy,
    runMailOperation,
    selectedMailIds,
    selectedMailIndex,
  })

  const {
    mailSearchProgressText,
    mailSearchSummaryItems,
    requestMailSearch,
    searchResultGroups,
    searchResultRows,
    setMailSearchResults,
    toggleSearchResultFolder,
    toggleSearchResultStore,
  } = useOutlookSearchController({
    activeMailSearchSummary,
    clearSelectedMails,
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
    waitForRequest,
  })

  const {
    loadAttachmentExportSettings,
    refreshAdminData,
    resetAttachmentExportRoot,
    saveAttachmentExportSettings,
    sendChat,
    switchView,
  } = useOutlookShellController({
    activeView,
    addinLogs,
    addinStatus,
    attachmentExportRootDraft,
    attachmentExportSettings,
    cancelScheduledMailFetch,
    chatMessages,
    chatPanelRef,
    chatText,
    clearAttachmentLoads,
    clearConversationLoads,
    clearMailBodyLoads,
    clearSelectedMailIndex,
    closeFolderContextMenu,
    mailListMode,
    mailListNeedsFetch,
    outlookBusy,
    outlookDependentViewsLocked,
    requestCalendar,
    requestRules,
    runStartupOutlookSync,
    savingAttachmentExportSettings,
    scheduleMailFetch,
    signalRState,
    setUnmounted: (value) => { unmounted = value },
  })

  return {
    activeView, activeMailPropertySections, addCategoryToMasterList, addinLogs, addinStatus,
    attachmentExportRootDraft, attachmentExportSettings, addMailCategoryDraft, applyMailProperties,
    beginCreateCalendarEvent, beginEditCalendarEvent, calendarAttendeeOptions, calendarDraft, calendarEditorMode,
    calendarEditorVisible, calendarEventDialogVisible, calendarEvents, calendarMonthLabel,
    calendarMergeHints, calendarRooms,
    calendarWeekdays, calendarWeeks, cancelCreateFolder,
    categories, categoryManagerVisible, categoryColorOptions, categoryColorStyle, categoryTagStyle,
    categoryCreateColor, categoryCreateDraft, changeCalendarMonth, chatMessages, chatPanelRef,
    chatText, clearMailDrag, clearSelectedMails, contextFolderName, createFolder,
    createFolderFromContext, creatingFolderName, creatingFolderParentPath, deleteFolderFromContext,
    deleteCalendarEvent, deleteMail, deleteSelectedMails, dragOverFolderPath, draggedMailId, expandedFolders, exportMailAttachment,
    fetchMailsFromContext, flagDisplayLabel, flagIntervalOptions, flagTagType, folderContextMenu,
    folderStores, loadingCalendar, loadingCategories, loadAttachmentExportSettings, loadingFolders,
    loadingMails, loadingMailSearch, loadingSignalRPing, mailCount, mailHtmlSandbox, mailListMode,
    mailSearchDraft, mailSearchProgress, mailSearchProgressText, mailSearchSummaryItems,
    searchResultGroups, searchResultRows, searchResultViewMode, mailSearchResults,
    isAttachmentExporting, isAttachmentListLoading, isConversationLoading, isMailBodyLoading,
    mailHasBody, dialogMail, dialogMailAttachments, dialogMailConversation, dialogMailConversationItems,
    dialogMailFolderName, dialogMailHasIdentity, dialogLoading, mailDialogHtml, mailDialogVisible,
    mailPropertiesDraft, mailPropertiesChanged, mailLookbackHours, mailStats, masterCategoryListExpanded,
    mails, moveDraggedMail, navOptions, openFolderContextMenu, openExportedAttachment,
    openMailDialog, closeMailDialog,
    operationLoading: requestLoading,
    outlookBusy, outlookBusyText, outlookFirstLoadCompleted, openCategoryManager, refreshAdminData, requestCalendar,
    requestCategories, requestFolders, requestRules, requestSignalRPing, requestMails,
    requestMailSearch, resetMailPropertiesDraft, resetRuleDraft, resetAttachmentExportRoot,
    removeMailCategoryDraft, saveAttachmentExportSettings, savingAttachmentExportSettings,
    fetchedMailFolderName, mailListNeedsFetch, mailFetchCountdownText, showMailFetchWarning,
    mailFetchStatusText, selectedFolderName, selectedCalendarEvent, selectedFolderPath,
    selectedMail, selectedMailCategories, selectedMailIndex, selectedMailIds, selectedRule,
    selectedRuleIndex, selectFolder, selectCalendarEvent, selectMail, sendChat, saveCalendarEvent, saveRule,
    setCalendarRoom,
    deleteRule, editRule, showFolderMails, goToCurrentCalendarMonth, setDragOverFolder,
    setMailFlagDraft, signalRState, splitCategories, startMailPointerDrag, switchView, toggleFolder,
    toggleSearchResultFolder, toggleSearchResultStore, updateCategoryColor, toggleRuleEnabled,
    visibleFolders, folderOptions, ruleDraft, ruleDraftIsEditing, rules, loadingRules,
    visibleMasterCategories, hiddenMasterCategoryCount, toggleMasterCategoryList, flagEditorVisible,
  }
}

export type OutlookDashboardState = ReturnType<typeof useOutlookDashboard>
