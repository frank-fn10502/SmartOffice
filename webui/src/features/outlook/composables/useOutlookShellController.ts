import { nextTick, onBeforeUnmount, onMounted, type Ref } from 'vue'
import * as signalR from '@microsoft/signalr'
import { outlookApi } from '../api/outlook'
import type {
  AddinLogEntry,
  AddinStatusDto,
  AppView,
  AttachmentExportSettingsDto,
  ChatMessageDto,
  MailBodyDto,
  SignalRState,
} from '../models/outlook'

type ShellControllerOptions = {
  activeView: Ref<AppView>
  addinLogs: Ref<AddinLogEntry[]>
  addinStatus: Ref<AddinStatusDto>
  attachmentExportRootDraft: Ref<string>
  attachmentExportSettings: Ref<AttachmentExportSettingsDto>
  cancelScheduledMailFetch: () => void
  chatMessages: Ref<ChatMessageDto[]>
  chatPanelRef: Ref<HTMLElement | null>
  chatText: Ref<string>
  clearAttachmentLoads: () => void
  clearConversationLoads: () => void
  clearMailBodyLoads: () => void
  clearSelectedMailIndex: () => void
  closeFolderContextMenu: () => void
  mailListMode: Ref<'folder' | 'search'>
  mailListNeedsFetch: Ref<boolean>
  outlookBusy: Ref<boolean>
  outlookDependentViewsLocked: Ref<boolean>
  patchMailBody: (payload: MailBodyDto | unknown) => void
  requestCalendar: () => Promise<void>
  requestRules: () => Promise<void>
  runStartupOutlookSync: () => Promise<void>
  savingAttachmentExportSettings: Ref<boolean>
  scheduleMailFetch: () => void
  signalRState: Ref<SignalRState>
  setUnmounted: (value: boolean) => void
}

export function useOutlookShellController(options: ShellControllerOptions) {
  const {
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
    patchMailBody,
    requestCalendar,
    requestRules,
    runStartupOutlookSync,
    savingAttachmentExportSettings,
    scheduleMailFetch,
    signalRState,
    setUnmounted,
  } = options
  let connection: signalR.HubConnection | null = null
  const loadedViews = new Set<AppView>()

  async function scrollChatToBottom() {
    await nextTick()
    if (chatPanelRef.value) chatPanelRef.value.scrollTop = chatPanelRef.value.scrollHeight
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
    await loadChat()
  }

  async function refreshAdminData() {
    const [status, logs] = await Promise.all([
      outlookApi.getAdminStatus(),
      outlookApi.getAdminLogs(),
    ])
    addinStatus.value = status
    addinLogs.value = logs
  }

  async function loadAttachmentExportSettings() {
    const settings = await outlookApi.getAttachmentExportSettings()
    attachmentExportSettings.value = settings
    attachmentExportRootDraft.value = settings.rootPath
  }

  async function saveAttachmentExportSettings() {
    if (savingAttachmentExportSettings.value) return
    savingAttachmentExportSettings.value = true
    try {
      const settings = await outlookApi.updateAttachmentExportSettings({
        rootPath: attachmentExportRootDraft.value,
      })
      attachmentExportSettings.value = settings
      attachmentExportRootDraft.value = settings.rootPath
    } finally {
      savingAttachmentExportSettings.value = false
    }
  }

  async function resetAttachmentExportRoot() {
    attachmentExportRootDraft.value = attachmentExportSettings.value.defaultRootPath
    await saveAttachmentExportSettings()
  }

  async function switchView(view: AppView) {
    if (outlookDependentViewsLocked.value && ['search', 'rules', 'chat', 'calendar', 'contacts', 'profile'].includes(view)) return
    activeView.value = view
    if (view === 'admin') {
      cancelScheduledMailFetch()
      clearSelectedMailIndex()
      void refreshAdminData()
      void loadAttachmentExportSettings()
      return
    }
    if (view === 'outlook') {
      mailListMode.value = 'folder'
      clearSelectedMailIndex()
      if (mailListNeedsFetch.value && !outlookBusy.value) scheduleMailFetch()
      return
    }
    if (view === 'search') {
      cancelScheduledMailFetch()
      mailListMode.value = 'search'
      clearSelectedMailIndex()
    }
    void loadViewOnce(view)
  }

  async function loadViewOnce(view: AppView) {
    if (loadedViews.has(view)) return
    if (view === 'chat') {
      loadedViews.add(view)
      try {
        await loadChat()
      } catch {
        loadedViews.delete(view)
      }
      return
    }
    if (view === 'rules') {
      if (outlookBusy.value) return
      loadedViews.add(view)
      try {
        await requestRules()
      } catch {
        loadedViews.delete(view)
      }
      return
    }
    if (view === 'calendar') {
      if (outlookBusy.value) return
      loadedViews.add(view)
      try {
        await requestCalendar()
      } catch {
        loadedViews.delete(view)
      }
    }
  }

  async function connectSignalR() {
    connection = new signalR.HubConnectionBuilder()
      .withUrl('/hub/notifications')
      .withAutomaticReconnect()
      .build()

    connection.onreconnecting(() => { signalRState.value = 'reconnecting' })
    connection.onreconnected(() => {
      signalRState.value = 'connected'
      void refreshAdminData()
    })
    connection.onclose(() => { signalRState.value = 'disconnected' })
    connection.on('AddinStatus', (status: AddinStatusDto) => { addinStatus.value = status })
    connection.on('AddinLog', (logs: AddinLogEntry[]) => { addinLogs.value = logs })
    connection.on('MailBodyUpdated', (body: MailBodyDto | unknown) => { patchMailBody(body) })

    try {
      await connection.start()
      signalRState.value = 'connected'
    } catch {
      signalRState.value = 'disconnected'
    }
  }

  onMounted(async () => {
    setUnmounted(false)
    window.addEventListener('click', closeFolderContextMenu)
    void connectSignalR()
    await Promise.allSettled([
      refreshAdminData(),
      loadAttachmentExportSettings(),
    ])
    void runStartupOutlookSync()
  })

  onBeforeUnmount(() => {
    setUnmounted(true)
    window.removeEventListener('click', closeFolderContextMenu)
    cancelScheduledMailFetch()
    clearMailBodyLoads()
    clearAttachmentLoads()
    clearConversationLoads()
    void connection?.stop()
  })

  return {
    loadAttachmentExportSettings,
    refreshAdminData,
    resetAttachmentExportRoot,
    saveAttachmentExportSettings,
    sendChat,
    switchView,
  }
}
