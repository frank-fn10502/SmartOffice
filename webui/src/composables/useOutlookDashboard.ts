import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import * as signalR from '@microsoft/signalr'
import { outlookApi } from '../api/outlook'
import type {
  AddinLogEntry,
  AddinStatusDto,
  AppView,
  CalendarEventDto,
  ChatMessageDto,
  FolderDto,
  MailItemDto,
  MailPropertiesCommandRequest,
  OutlookCategoryDto,
  OutlookRuleDto,
  SignalRState,
} from '../models/outlook'
import { collectFolderOptions, folderType, visibleRootFolders } from '../utils/folders'

const flagIntervalOptions = [
  { label: '不設定旗標', value: 'none' },
  { label: '今天', value: 'today' },
  { label: '明天', value: 'tomorrow' },
  { label: '本週', value: 'this_week' },
  { label: '下週', value: 'next_week' },
  { label: '無日期', value: 'no_date' },
  { label: '自訂日期', value: 'custom' },
  { label: '標示完成', value: 'complete' },
]

const categoryColorOptions = [
  { label: '無色', value: 'preset0', color: '#eef2f7' },
  { label: '紅色', value: 'preset1', color: '#f87171' },
  { label: '橘色', value: 'preset2', color: '#fb923c' },
  { label: '黃色', value: 'preset3', color: '#facc15' },
  { label: '綠色', value: 'preset4', color: '#22c55e' },
  { label: '藍色', value: 'preset5', color: '#38bdf8' },
  { label: '紫色', value: 'preset6', color: '#a78bfa' },
  { label: '粉紅', value: 'preset7', color: '#f472b6' },
  { label: '深藍', value: 'preset8', color: '#2563eb' },
]

function defaultFlagRequest(value: string) {
  return value === 'none' ? '' : flagIntervalLabel(value)
}

function flagIntervalLabel(value?: string) {
  return flagIntervalOptions.find((option) => option.value === value)?.label ?? '旗標'
}

function isDefaultFlagRequest(value: string, previousInterval = '') {
  const normalized = value.trim()
  return (
    !normalized
    || normalized === 'Follow up'
    || normalized === '旗標'
    || normalized === flagIntervalLabel(previousInterval)
    || flagIntervalOptions.some((option) => option.label === normalized)
  )
}

function toDateInput(value?: string) {
  if (!value) return ''
  if (/^\d{4}-\d{2}-\d{2}/.test(value)) return value.slice(0, 10)
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return ''
  return date.toISOString().slice(0, 10)
}

function dateInputToIso(value: string) {
  return value ? `${value}T00:00:00` : undefined
}

function todayInputValue() {
  const now = new Date()
  const year = now.getFullYear()
  const month = `${now.getMonth() + 1}`.padStart(2, '0')
  const day = `${now.getDate()}`.padStart(2, '0')
  return `${year}-${month}-${day}`
}

function splitCategories(value: string) {
  return value
    .split(',')
    .map((category) => category.trim())
    .filter(Boolean)
}

function categoryColorStyle(value?: string) {
  const color = categoryColorOptions.find((option) => option.value === value)?.color ?? categoryColorOptions[0].color
  return { backgroundColor: color }
}

export function useOutlookDashboard() {
  const activeView = ref<AppView>('outlook')
  const signalRState = ref<SignalRState>('disconnected')
  const folders = ref<FolderDto[]>([])
  const mails = ref<MailItemDto[]>([])
  const rules = ref<OutlookRuleDto[]>([])
  const categories = ref<OutlookCategoryDto[]>([])
  const calendarEvents = ref<CalendarEventDto[]>([])
  const chatMessages = ref<ChatMessageDto[]>([])
  const addinStatus = ref<AddinStatusDto>({
    connected: false,
    lastCommand: '',
  })
  const addinLogs = ref<AddinLogEntry[]>([])
  const selectedFolderPath = ref('')
  const selectedMailIndex = ref<number | null>(null)
  const selectedMailHtml = ref(false)
  const activePropertyLibrarySections = ref(['property-library'])
  const activeMailPropertySections = ref(['set-mail-properties'])
  const expandedFolders = ref<Set<string>>(new Set())
  const openMailIndexes = ref<Set<number>>(new Set())
  const htmlMailIndexes = ref<Set<number>>(new Set())
  const mailRange = ref('1d')
  const mailCount = ref(10)
  const chatText = ref('')
  const loadingFolders = ref(false)
  const loadingMails = ref(false)
  const loadingRules = ref(false)
  const loadingCategories = ref(false)
  const loadingCalendar = ref(false)
  const loadingSignalRPing = ref(false)
  const operationLoading = ref(false)
  const mailPropertiesDraft = ref({
    isRead: false,
    flagInterval: 'none',
    flagRequest: '',
    taskStartDate: '',
    taskDueDate: '',
    taskCompletedDate: '',
    categories: [] as string[],
  })
  const categoryCreateDraft = ref('')
  const categoryCreateColor = ref('preset0')
  const creatingFolderParentPath = ref('')
  const creatingFolderName = ref('')
  const draggedMailId = ref('')
  const dragOverFolderPath = ref('')
  const folderContextMenu = ref({
    visible: false,
    x: 0,
    y: 0,
    folderPath: '',
  })
  const chatPanelRef = ref<HTMLElement | null>(null)
  const mailHtmlSandbox = 'allow-same-origin allow-popups allow-popups-to-escape-sandbox'
  let connection: signalR.HubConnection | null = null

  const visibleFolders = computed(() => visibleRootFolders(folders.value))

  const mailStats = computed(() => ({
    unread: mails.value.filter((mail) => !mail.isRead).length,
    flagged: mails.value.filter((mail) => mail.isMarkedAsTask).length,
    highImportance: mails.value.filter((mail) => mail.importance === 'high').length,
    categorized: mails.value.filter((mail) => Boolean(mail.categories)).length,
  }))

  const outlookBusy = computed(() => {
    return loadingFolders.value || loadingMails.value || loadingRules.value || loadingCategories.value || loadingCalendar.value || operationLoading.value
  })

  const outlookBusyText = computed(() => {
    if (loadingFolders.value) return 'Outlook folder 同步中...'
    if (loadingMails.value) return 'Outlook 郵件抓取中...'
    if (loadingRules.value) return 'Outlook rule 同步中...'
    if (loadingCategories.value) return 'Outlook category 同步中...'
    if (loadingCalendar.value) return 'Outlook calendar 同步中...'
    if (operationLoading.value) return 'Outlook 操作執行中...'
    return ''
  })

  const folderOptions = computed(() => collectFolderOptions(visibleFolders.value))

  const selectedFolderName = computed(() => {
    return folderOptions.value.find((folder) => folder.folderPath === selectedFolderPath.value)?.label.trim() ?? '未選擇'
  })

  const contextFolderName = computed(() => {
    return folderOptions.value.find((folder) => folder.folderPath === folderContextMenu.value.folderPath)?.label.trim() ?? '未選擇'
  })

  const selectedMail = computed(() => {
    if (selectedMailIndex.value === null) return null
    return mails.value[selectedMailIndex.value] ?? null
  })

  const selectedMailIsOpen = computed(() => {
    return selectedMailIndex.value !== null && openMailIndexes.value.has(selectedMailIndex.value)
  })

  const selectedMailCategories = computed(() => {
    if (!selectedMail.value?.categories) return []
    return selectedMail.value.categories
      .split(',')
      .map((category) => category.trim())
      .filter(Boolean)
  })

  function resetMailPropertiesDraft(mail: MailItemDto | null) {
    if (!mail) return
    const flagInterval = mail.flagInterval || (mail.isMarkedAsTask ? 'today' : 'none')
    mailPropertiesDraft.value = {
      isRead: mail.isRead,
      flagInterval,
      flagRequest: isDefaultFlagRequest(mail.flagRequest) ? defaultFlagRequest(flagInterval) : mail.flagRequest,
      taskStartDate: toDateInput(mail.taskStartDate),
      taskDueDate: toDateInput(mail.taskDueDate),
      taskCompletedDate: toDateInput(mail.taskCompletedDate),
      categories: splitCategories(mail.categories),
    }
  }

  function setMails(items: MailItemDto[], preferredMailId = selectedMail.value?.id ?? '') {
    const wasOpen = selectedMailIndex.value !== null && openMailIndexes.value.has(selectedMailIndex.value)
    mails.value = items

    if (items.length === 0) {
      selectedMailIndex.value = null
      openMailIndexes.value = new Set()
      htmlMailIndexes.value = new Set()
      return
    }

    const nextIndex = preferredMailId ? items.findIndex((mail) => mail.id === preferredMailId) : -1
    selectedMailIndex.value = nextIndex >= 0 ? nextIndex : 0

    if (wasOpen && selectedMailIndex.value !== null) {
      openMailIndexes.value = new Set([selectedMailIndex.value])
    }
  }

  async function loadCachedFolders() {
    folders.value = await outlookApi.getFolders()
    selectDefaultFolder()
  }

  async function requestFolders() {
    if (outlookBusy.value) return
    loadingFolders.value = true
    try {
      await outlookApi.requestFolders()
    } catch {
      loadingFolders.value = false
    }
  }

  function selectDefaultFolder() {
    if (selectedFolderPath.value || visibleFolders.value.length === 0) return
    const inbox = visibleFolders.value.find((folder) => folderType(folder.name) === 'inbox')
    selectedFolderPath.value = inbox?.folderPath ?? visibleFolders.value[0]?.folderPath ?? ''
  }

  function toggleFolder(path: string) {
    if (outlookBusy.value) return
    const next = new Set(expandedFolders.value)
    if (next.has(path)) next.delete(path)
    else next.add(path)
    expandedFolders.value = next
  }

  function selectFolder(path: string) {
    if (outlookBusy.value) return
    selectedFolderPath.value = path
    selectedMailIndex.value = null
    selectedMailHtml.value = false
  }

  function openFolderContextMenu(payload: { path: string; x: number; y: number }) {
    if (outlookBusy.value) return
    selectFolder(payload.path)
    folderContextMenu.value = {
      visible: true,
      x: payload.x,
      y: payload.y,
      folderPath: payload.path,
    }
  }

  function closeFolderContextMenu() {
    folderContextMenu.value.visible = false
  }

  async function createFolderFromContext() {
    beginCreateFolder(folderContextMenu.value.folderPath)
    closeFolderContextMenu()
  }

  async function deleteFolderFromContext() {
    const targetPath = folderContextMenu.value.folderPath
    closeFolderContextMenu()
    await deleteFolder(targetPath)
  }

  async function fetchMailsFromContext() {
    selectedFolderPath.value = folderContextMenu.value.folderPath
    closeFolderContextMenu()
    await requestMails()
  }

  async function loadCachedMails() {
    setMails(await outlookApi.getMails())
  }

  async function loadCachedRules() {
    rules.value = await outlookApi.getRules()
  }

  async function loadCachedCategories() {
    categories.value = await outlookApi.getCategories()
  }

  async function loadCachedCalendar() {
    calendarEvents.value = await outlookApi.getCalendar()
  }

  async function requestMails() {
    if (outlookBusy.value) return
    loadingMails.value = true
    openMailIndexes.value = new Set()
    htmlMailIndexes.value = new Set()
    selectedMailIndex.value = null
    selectedMailHtml.value = false
    try {
      await outlookApi.requestMails({
        folderPath: selectedFolderPath.value,
        range: mailRange.value,
        maxCount: mailCount.value,
      })
    } catch {
      loadingMails.value = false
    }
  }

  async function requestRules() {
    if (outlookBusy.value) return
    loadingRules.value = true
    try {
      await outlookApi.requestRules()
    } catch {
      loadingRules.value = false
    }
  }

  async function requestCategories() {
    if (outlookBusy.value) return
    loadingCategories.value = true
    try {
      await outlookApi.requestCategories()
    } catch {
      loadingCategories.value = false
    }
  }

  async function requestSignalRPing() {
    if (loadingSignalRPing.value) return
    loadingSignalRPing.value = true
    try {
      await outlookApi.requestSignalRPing()
    } finally {
      loadingSignalRPing.value = false
    }
  }

  async function requestCalendar() {
    if (outlookBusy.value) return
    loadingCalendar.value = true
    try {
      await outlookApi.requestCalendar({ daysForward: 14 })
    } catch {
      loadingCalendar.value = false
    }
  }

  async function runMailOperation(action: () => Promise<unknown>) {
    if (outlookBusy.value) return
    operationLoading.value = true
    try {
      await action()
    } catch {
      operationLoading.value = false
    }
  }

  async function applyMailProperties(mail: MailItemDto) {
    const selectedCategories = [...new Set(mailPropertiesDraft.value.categories.map((category) => category.trim()).filter(Boolean))]
    const existingCategoryNames = new Set(categories.value.map((category) => category.name.toLowerCase()))
    const newCategories = selectedCategories
      .filter((category) => !existingCategoryNames.has(category.toLowerCase()))
      .map((name) => ({ name, color: 'preset0', shortcutKey: '' }))
    const isCustomFlag = mailPropertiesDraft.value.flagInterval === 'custom'
    const body: MailPropertiesCommandRequest = {
      mailId: mail.id,
      folderPath: mail.folderPath,
      isRead: mailPropertiesDraft.value.isRead,
      flagInterval: mailPropertiesDraft.value.flagInterval,
      flagRequest: mailPropertiesDraft.value.flagRequest || defaultFlagRequest(mailPropertiesDraft.value.flagInterval),
      taskStartDate: isCustomFlag ? dateInputToIso(mailPropertiesDraft.value.taskStartDate) : undefined,
      taskDueDate: isCustomFlag ? dateInputToIso(mailPropertiesDraft.value.taskDueDate) : undefined,
      taskCompletedDate: mailPropertiesDraft.value.flagInterval === 'complete' ? dateInputToIso(mailPropertiesDraft.value.taskCompletedDate) : undefined,
      categories: selectedCategories,
      newCategories,
    }
    await runMailOperation(() => outlookApi.requestUpdateMailProperties(body))
  }

  async function upsertCategory(name: string, color: string, shortcutKey = '') {
    const categoryName = name.trim()
    if (!categoryName || outlookBusy.value) return
    operationLoading.value = true
    try {
      await outlookApi.requestUpsertCategory({
        name: categoryName,
        color: color || 'preset0',
        shortcutKey,
      })
    } catch {
      operationLoading.value = false
    }
  }

  async function addCategoryToMasterList() {
    const name = categoryCreateDraft.value.trim()
    if (!name) return
    await upsertCategory(name, categoryCreateColor.value)
    categoryCreateDraft.value = ''
    categoryCreateColor.value = 'preset0'
  }

  async function updateCategoryColor(category: OutlookCategoryDto, color: string) {
    await upsertCategory(category.name, color, category.shortcutKey)
  }

  function beginCreateFolder(parentPath: string) {
    if (!parentPath || outlookBusy.value) return
    creatingFolderParentPath.value = parentPath
    creatingFolderName.value = ''
    const next = new Set(expandedFolders.value)
    next.add(parentPath)
    expandedFolders.value = next
  }

  function cancelCreateFolder() {
    creatingFolderParentPath.value = ''
    creatingFolderName.value = ''
  }

  async function createFolder(parentPath = creatingFolderParentPath.value, name = creatingFolderName.value) {
    const folderName = name.trim()
    if (!parentPath || !folderName || outlookBusy.value) return
    operationLoading.value = true
    try {
      await outlookApi.requestCreateFolder({
        parentFolderPath: parentPath,
        name: folderName,
      })
      cancelCreateFolder()
    } catch {
      operationLoading.value = false
    }
  }

  async function deleteFolder(targetPath: string) {
    if (!targetPath || outlookBusy.value) return
    const targetName = folderOptions.value.find((folder) => folder.folderPath === targetPath)?.label.trim() ?? targetPath
    const confirmed = window.confirm(`刪除 Folder「${targetName}」？`)
    if (!confirmed) return
    operationLoading.value = true
    try {
      await outlookApi.requestDeleteFolder({
        folderPath: targetPath,
      })
      if (selectedFolderPath.value === targetPath) {
        selectedFolderPath.value = folderOptions.value[0]?.folderPath ?? ''
      }
    } catch {
      operationLoading.value = false
    }
  }

  function selectMail(index: number) {
    selectedMailIndex.value = index
    selectedMailHtml.value = false
    const next = new Set<number>()
    next.add(index)
    openMailIndexes.value = next
  }

  async function moveMailToFolder(mail: MailItemDto, destinationFolderPath: string) {
    if (!destinationFolderPath || destinationFolderPath === mail.folderPath) return
    await runMailOperation(() =>
      outlookApi.requestMoveMail({
        mailId: mail.id,
        sourceFolderPath: mail.folderPath,
        destinationFolderPath,
      }),
    )
  }

  function startMailDrag(mail: MailItemDto, index: number, event: DragEvent) {
    if (outlookBusy.value) return
    selectMail(index)
    draggedMailId.value = mail.id
    event.dataTransfer?.setData('text/plain', mail.id)
    if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move'
  }

  function clearMailDrag() {
    draggedMailId.value = ''
    dragOverFolderPath.value = ''
  }

  function setDragOverFolder(path: string) {
    if (!draggedMailId.value || outlookBusy.value) return
    dragOverFolderPath.value = path
  }

  async function moveDraggedMail(destinationFolderPath: string) {
    const mailId = draggedMailId.value
    clearMailDrag()
    if (!mailId || outlookBusy.value) return
    const mail = mails.value.find((item) => item.id === mailId)
    if (!mail) return
    await moveMailToFolder(mail, destinationFolderPath)
  }

  function toggleMail(index: number) {
    const next = new Set(openMailIndexes.value)
    if (next.has(index)) next.delete(index)
    else next.add(index)
    openMailIndexes.value = next
  }

  function toggleMailFormat(index: number) {
    const next = new Set(htmlMailIndexes.value)
    if (next.has(index)) next.delete(index)
    else next.add(index)
    htmlMailIndexes.value = next
  }

  async function loadChat() {
    chatMessages.value = await outlookApi.getChat()
    await scrollChatToBottom()
  }

  async function sendChat() {
    const text = chatText.value.trim()
    if (!text) return
    chatText.value = ''
    await outlookApi.sendChat({ source: 'web', text })
  }

  async function refreshAdminData() {
    const [status, logs] = await Promise.all([
      outlookApi.getAdminStatus(),
      outlookApi.getAdminLogs(),
    ])
    addinStatus.value = status
    addinLogs.value = logs
  }

  async function switchView(view: AppView) {
    activeView.value = view
    if (view === 'admin') await refreshAdminData()
  }

  async function scrollChatToBottom() {
    await nextTick()
    if (chatPanelRef.value) chatPanelRef.value.scrollTop = chatPanelRef.value.scrollHeight
  }

  async function connectSignalR() {
    connection = new signalR.HubConnectionBuilder()
      .withUrl('/hub/notifications')
      .withAutomaticReconnect()
      .build()

    connection.onreconnecting(() => {
      signalRState.value = 'reconnecting'
    })
    connection.onreconnected(() => {
      signalRState.value = 'connected'
    })
    connection.onclose(() => {
      signalRState.value = 'disconnected'
    })
    connection.on('FoldersUpdated', (items: FolderDto[]) => {
      folders.value = items
      selectDefaultFolder()
      loadingFolders.value = false
      operationLoading.value = false
    })
    connection.on('MailsUpdated', (items: MailItemDto[]) => {
      setMails(items)
      loadingMails.value = false
      operationLoading.value = false
    })
    connection.on('RulesUpdated', (items: OutlookRuleDto[]) => {
      rules.value = items
      loadingRules.value = false
      operationLoading.value = false
    })
    connection.on('CategoriesUpdated', (items: OutlookCategoryDto[]) => {
      categories.value = items
      loadingCategories.value = false
      operationLoading.value = false
    })
    connection.on('CalendarUpdated', (items: CalendarEventDto[]) => {
      calendarEvents.value = items
      loadingCalendar.value = false
      operationLoading.value = false
    })
    connection.on('NewChatMessage', async (message: ChatMessageDto) => {
      chatMessages.value = [...chatMessages.value, message]
      await scrollChatToBottom()
    })
    connection.on('AddinStatus', (status: AddinStatusDto) => {
      addinStatus.value = status
    })
    connection.on('AddinLog', (logs: AddinLogEntry[]) => {
      addinLogs.value = logs
    })

    try {
      await connection.start()
      signalRState.value = 'connected'
    } catch {
      signalRState.value = 'disconnected'
    }
  }

  watch(
    () => selectedMail.value?.id,
    () => resetMailPropertiesDraft(selectedMail.value),
  )

  watch(
    () => mailPropertiesDraft.value.flagInterval,
    (next, previous) => {
      if (isDefaultFlagRequest(mailPropertiesDraft.value.flagRequest, previous)) {
        mailPropertiesDraft.value.flagRequest = defaultFlagRequest(next)
      }
      if (next === 'custom') {
        const today = todayInputValue()
        mailPropertiesDraft.value.taskStartDate ||= today
        mailPropertiesDraft.value.taskDueDate ||= today
      }
    },
  )

  onMounted(async () => {
    window.addEventListener('click', closeFolderContextMenu)
    await connectSignalR()
    await Promise.allSettled([
      loadCachedFolders(),
      loadCachedMails(),
      loadCachedRules(),
      loadCachedCategories(),
      loadCachedCalendar(),
      loadChat(),
      refreshAdminData(),
    ])
    if (folders.value.length === 0) await requestFolders()
    if (categories.value.length === 0) await requestCategories()
  })

  onBeforeUnmount(() => {
    window.removeEventListener('click', closeFolderContextMenu)
    void connection?.stop()
  })

  return {
    activeView,
    activeMailPropertySections,
    activePropertyLibrarySections,
    addCategoryToMasterList,
    addinLogs,
    addinStatus,
    applyMailProperties,
    calendarEvents,
    cancelCreateFolder,
    categories,
    categoryColorOptions,
    categoryColorStyle,
    categoryCreateColor,
    categoryCreateDraft,
    chatMessages,
    chatPanelRef,
    chatText,
    clearMailDrag,
    contextFolderName,
    createFolder,
    createFolderFromContext,
    creatingFolderName,
    creatingFolderParentPath,
    deleteFolderFromContext,
    dragOverFolderPath,
    draggedMailId,
    expandedFolders,
    fetchMailsFromContext,
    flagIntervalLabel,
    flagIntervalOptions,
    folderContextMenu,
    htmlMailIndexes,
    loadingCalendar,
    loadingFolders,
    loadingMails,
    loadingSignalRPing,
    mailCount,
    mailHtmlSandbox,
    mailPropertiesDraft,
    mailRange,
    mailStats,
    mails,
    moveDraggedMail,
    openFolderContextMenu,
    openMailIndexes,
    operationLoading,
    outlookBusy,
    outlookBusyText,
    refreshAdminData,
    requestCalendar,
    requestFolders,
    requestSignalRPing,
    requestMails,
    resetMailPropertiesDraft,
    selectedFolderName,
    selectedFolderPath,
    selectedMail,
    selectedMailCategories,
    selectedMailHtml,
    selectedMailIndex,
    selectedMailIsOpen,
    selectFolder,
    selectMail,
    sendChat,
    setDragOverFolder,
    signalRState,
    startMailDrag,
    switchView,
    toggleFolder,
    toggleMail,
    toggleMailFormat,
    updateCategoryColor,
    visibleFolders,
  }
}
