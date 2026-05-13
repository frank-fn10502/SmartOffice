import type { ComputedRef, Ref } from 'vue'
import { ElMessage } from 'element-plus'
import { outlookApi } from '../api/outlook'
import type { FolderDto, FolderTreeNode, MailItemDto, OutlookStoreDto } from '../models/outlook'
import { buildFolderTree, findFolderByPath, folderType, isMailSelectableFolder } from '../utils/folders'
import { collectOutlookRequestData } from './outlookRequests'

type FolderOption = any

type FoldersControllerOptions = {
  clearSelectedMails: () => void
  creatingFolderName: Ref<string>
  creatingFolderParentPath: Ref<string>
  expandedFolders: Ref<Set<string>>
  fetchedMailFolderPath: Ref<string>
  folderContextMenu: Ref<{ visible: boolean; x: number; y: number; folderPath: string }>
  folderOptions: ComputedRef<FolderOption[]>
  folders: Ref<FolderTreeNode[]>
  folderStores: Ref<OutlookStoreDto[]>
  folderMails: Ref<MailItemDto[]>
  loadingFolders: Ref<boolean>
  loadingMails: Ref<boolean>
  lastMailFetchAt: Ref<Date | null>
  mailCount: Ref<number>
  mailFetchCountdownTick: Ref<number>
  mailListMode: Ref<'folder' | 'search'>
  mailLookbackHours: Ref<number>
  outlookBusy: Ref<boolean>
  pendingMailFolderPath: Ref<string>
  scheduledMailFetchAt: Ref<number>
  selectedFolderPath: Ref<string>
  selectedMailIndex: Ref<number | null>
  visibleFolders: ComputedRef<FolderTreeNode[]>
  clearSelectedMailIndex: () => void
  inferMailFolderPath: (items: MailItemDto[], fallback?: string) => string
  loadRequestMailItems: (response: { requestId?: string; request?: string }) => Promise<MailItemDto[]>
  markInitialFoldersComplete: () => void
  markInitialMailsComplete: () => void
  setMails: (items: MailItemDto[], preferredMailId?: string) => void
  waitForRequest: (response: { requestId?: string; request?: string }, timeoutMs?: number) => Promise<void>
}

const mailFetchDelayMs = 300
const mailFetchCountdownTickMs = 100

export function useOutlookFoldersController(options: FoldersControllerOptions) {
  const {
    clearSelectedMailIndex,
    clearSelectedMails,
    creatingFolderName,
    creatingFolderParentPath,
    expandedFolders,
    fetchedMailFolderPath,
    folderContextMenu,
    folderOptions,
    folders,
    folderStores,
    loadingFolders,
    loadingMails,
    lastMailFetchAt,
    mailCount,
    mailFetchCountdownTick,
    mailListMode,
    mailLookbackHours,
    outlookBusy,
    pendingMailFolderPath,
    scheduledMailFetchAt,
    selectedFolderPath,
    selectedMailIndex,
    visibleFolders,
    inferMailFolderPath,
    loadRequestMailItems,
    markInitialFoldersComplete,
    markInitialMailsComplete,
    setMails,
    waitForRequest,
  } = options
  let mailFetchTimeoutId = 0
  let mailFetchCountdownIntervalId = 0

  function folderStore(folder: { storeId: string }) {
    return folderStores.value.find((store) => store.storeId === folder.storeId)
  }

  function collectExistingFolderCounts() {
    const counts = new Map<string, number>()
    function visit(items: FolderTreeNode[]) {
      for (const folder of items) {
        counts.set(folder.folderPath, folder.itemCount)
        visit(folder.subFolders)
      }
    }
    visit(folders.value)
    return counts
  }

  async function loadFoldersFromRequest(response: { requestId?: string; request?: string }, options: { preserveExistingCounts?: boolean } = {}) {
    const pages = await collectOutlookRequestData<{ stores?: OutlookStoreDto[]; folders?: FolderDto[] }>(response)
    const snapshot = {
      stores: pages[0]?.data?.stores ?? [],
      folders: pages.flatMap((page) => page.data?.folders ?? []),
    }
    if (options.preserveExistingCounts) {
      const existingCounts = collectExistingFolderCounts()
      snapshot.folders = snapshot.folders.map((folder) => {
        const previousCount = existingCounts.get(folder.folderPath) ?? 0
        return previousCount > 0 && folder.itemCount === 0 ? { ...folder, itemCount: previousCount } : folder
      })
    }
    folderStores.value = snapshot.stores
    folders.value = buildFolderTree(snapshot)
    selectDefaultFolder()
  }

  async function requestFolders(force = false) {
    if (outlookBusy.value && !force) return
    loadingFolders.value = true
    try {
      const response = await outlookApi.requestFolders()
      await waitForRequest(response)
      await loadFoldersFromRequest(response)
      await loadVisibleRootChildren()
      markInitialFoldersComplete()
      loadingFolders.value = false
    } catch {
      markInitialFoldersComplete()
      loadingFolders.value = false
    }
  }

  function findLoadedInboxFolder() {
    const inboxes = folderOptions.value.filter((folder) => isMailSelectableFolder(folder) && folderType(folder.name) === 'inbox')
    return (
      inboxes.find((folder) => folderStore(folder)?.storeKind?.toLowerCase() === 'ost')
      ?? inboxes.find((folder) => ['exchange', 'ost'].includes(folderStore(folder)?.storeKind?.toLowerCase() ?? ''))
      ?? inboxes[0]
      ?? null
    )
  }

  function findPreferredInboxFolder() {
    return findLoadedInboxFolder() ?? folderOptions.value.find(isMailSelectableFolder) ?? null
  }

  function findPreferredInboxRootFolder() {
    return (
      visibleFolders.value.find((folder) => folderStore(folder)?.storeKind?.toLowerCase() === 'ost')
      ?? visibleFolders.value.find((folder) => ['exchange', 'ost'].includes(folderStore(folder)?.storeKind?.toLowerCase() ?? ''))
      ?? visibleFolders.value[0]
      ?? null
    )
  }

  async function ensureStartupInboxFolderLoaded() {
    if (findLoadedInboxFolder()) return
    const root = findPreferredInboxRootFolder()
    if (!root?.hasChildren || root.childrenLoaded) return

    loadingFolders.value = true
    try {
      const response = await outlookApi.requestFolderChildren({
        storeId: root.storeId,
        parentEntryId: root.entryId,
        parentFolderPath: root.folderPath,
        maxDepth: 1,
        maxChildren: 50,
      })
      await waitForRequest(response)
      await loadFoldersFromRequest(response, { preserveExistingCounts: true })
    } finally {
      loadingFolders.value = false
    }
  }

  async function loadVisibleRootChildren() {
    const roots = visibleFolders.value.filter((folder) => folder.hasChildren && !folder.childrenLoaded)
    for (const root of roots) {
      const next = new Set(expandedFolders.value)
      next.add(root.folderPath)
      expandedFolders.value = next

      const response = await outlookApi.requestFolderChildren({
        storeId: root.storeId,
        parentEntryId: root.entryId,
        parentFolderPath: root.folderPath,
        maxDepth: 1,
        maxChildren: 50,
      })
      await waitForRequest(response)
      await loadFoldersFromRequest(response, { preserveExistingCounts: true })
    }
  }

  function expandFolderAncestors(folderPath: string) {
    const next = new Set(expandedFolders.value)
    let current = folderOptions.value.find((folder) => folder.folderPath === folderPath)
    while (current?.parentFolderPath) {
      next.add(current.parentFolderPath)
      current = folderOptions.value.find((folder) => folder.folderPath === current?.parentFolderPath)
    }
    expandedFolders.value = next
  }

  function selectDefaultFolder() {
    if (selectedFolderPath.value || folderOptions.value.length === 0) return
    const folder = findPreferredInboxFolder()
    selectedFolderPath.value = folder?.folderPath ?? ''
    if (selectedFolderPath.value) expandFolderAncestors(selectedFolderPath.value)
  }

  function selectInboxFolder() {
    const folder = findPreferredInboxFolder()
    selectedFolderPath.value = folder?.folderPath ?? ''
    if (selectedFolderPath.value) expandFolderAncestors(selectedFolderPath.value)
    selectedMailIndex.value = null
  }

  async function toggleFolder(path: string) {
    if (outlookBusy.value) return
    const next = new Set(expandedFolders.value)
    if (next.has(path)) next.delete(path)
    else {
      next.add(path)
      const folder = findFolderByPath(folders.value, path)
      if (folder?.hasChildren && !folder.childrenLoaded) {
        loadingFolders.value = true
        try {
          const response = await outlookApi.requestFolderChildren({
            storeId: folder.storeId,
            parentEntryId: folder.entryId,
            parentFolderPath: folder.folderPath,
            maxDepth: 1,
            maxChildren: 50,
          })
          await waitForRequest(response)
          await loadFoldersFromRequest(response, { preserveExistingCounts: true })
        } finally {
          loadingFolders.value = false
        }
      }
    }
    expandedFolders.value = next
  }

  function selectFolder(path: string) {
    if (outlookBusy.value) return
    const folder = folderOptions.value.find((item) => item.folderPath === path)
    if (!folder || !isMailSelectableFolder(folder)) {
      void toggleFolder(path)
      return
    }
    selectedFolderPath.value = path
    selectedMailIndex.value = null
    scheduleMailFetch()
  }

  function scheduleMailFetch() {
    cancelScheduledMailFetch()
    if (!selectedFolderPath.value) return
    const folder = folderOptions.value.find((item) => item.folderPath === selectedFolderPath.value)
    if (!folder || !isMailSelectableFolder(folder)) return
    scheduledMailFetchAt.value = Date.now() + mailFetchDelayMs
    mailFetchCountdownTick.value = Date.now()
    mailFetchCountdownIntervalId = window.setInterval(() => {
      mailFetchCountdownTick.value = Date.now()
    }, mailFetchCountdownTickMs)
    mailFetchTimeoutId = window.setTimeout(() => {
      void requestMails()
    }, mailFetchDelayMs)
  }

  function cancelScheduledMailFetch() {
    if (mailFetchTimeoutId) window.clearTimeout(mailFetchTimeoutId)
    if (mailFetchCountdownIntervalId) window.clearInterval(mailFetchCountdownIntervalId)
    mailFetchTimeoutId = 0
    mailFetchCountdownIntervalId = 0
    scheduledMailFetchAt.value = 0
  }

  function openFolderContextMenu(payload: { path: string; x: number; y: number }) {
    if (outlookBusy.value) return
    const folder = folderOptions.value.find((item) => item.folderPath === payload.path)
    if (!folder || !isMailSelectableFolder(folder)) return
    selectedFolderPath.value = payload.path
    selectedMailIndex.value = null
    folderContextMenu.value = { visible: true, x: payload.x, y: payload.y, folderPath: payload.path }
  }

  function closeFolderContextMenu() {
    folderContextMenu.value.visible = false
  }

  function beginCreateFolder(parentPath: string) {
    if (!parentPath) return
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

  async function requestMails(force = false) {
    cancelScheduledMailFetch()
    const selectedFolder = folderOptions.value.find((folder) => folder.folderPath === selectedFolderPath.value)
    if ((outlookBusy.value && !force) || !selectedFolderPath.value || !selectedFolder || !isMailSelectableFolder(selectedFolder)) {
      if (!selectedFolderPath.value) {
        markInitialMailsComplete()
        loadingMails.value = false
      }
      return
    }
    loadingMails.value = true
    mailListMode.value = 'folder'
    pendingMailFolderPath.value = selectedFolderPath.value
    clearSelectedMailIndex()
    clearSelectedMails()
    try {
      const receivedTo = new Date()
      const receivedFrom = new Date(receivedTo.getTime() - mailLookbackHours.value * 60 * 60 * 1000)
      const response = await outlookApi.requestFolderMails({
        folderPath: selectedFolderPath.value,
        includeSubFolders: false,
        receivedFrom: receivedFrom.toISOString(),
        receivedTo: receivedTo.toISOString(),
        maxCount: mailCount.value,
      })
      await waitForRequest(response)
      const items = await loadRequestMailItems(response)
      setMails(items)
      fetchedMailFolderPath.value = inferMailFolderPath(items, pendingMailFolderPath.value)
      lastMailFetchAt.value = new Date()
      pendingMailFolderPath.value = ''
      markInitialMailsComplete()
      loadingMails.value = false
    } catch {
      pendingMailFolderPath.value = ''
      markInitialMailsComplete()
      loadingMails.value = false
    }
  }

  function createFolderFromContext() {
    beginCreateFolder(folderContextMenu.value.folderPath)
    closeFolderContextMenu()
  }

  async function fetchMailsFromContext() {
    const folder = folderOptions.value.find((item) => item.folderPath === folderContextMenu.value.folderPath)
    if (!folder || !isMailSelectableFolder(folder)) {
      closeFolderContextMenu()
      return
    }
    selectedFolderPath.value = folderContextMenu.value.folderPath
    closeFolderContextMenu()
    await requestMails()
  }

  return {
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
  }
}
