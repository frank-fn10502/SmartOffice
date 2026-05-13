import type { ComputedRef, Ref } from 'vue'
import { ElMessage } from 'element-plus'
import { outlookApi } from '../api/outlook'
import type { FolderTreeNode } from '../models/outlook'
import { isMailSelectableFolder } from '../utils/folders'

type FolderOption = FolderTreeNode & { label: string }

type FolderMutationsControllerOptions = {
  creatingFolderName: Ref<string>
  creatingFolderParentPath: Ref<string>
  folderContextMenu: Ref<{ visible: boolean; x: number; y: number; folderPath: string }>
  folderOptions: ComputedRef<FolderOption[]>
  outlookBusy: ComputedRef<boolean>
  requestLoading: Ref<boolean>
  selectedFolderPath: Ref<string>
  cancelCreateFolder: () => void
  closeFolderContextMenu: () => void
  folderNameForPath: (path: string) => string
  isInDeletedFolder: (path: string) => boolean
  loadCachedFolders: () => Promise<void>
  manualOutlookDeleteMessage: string
  waitForRequest: (response: { requestId?: string; request?: string }, timeoutMs?: number) => Promise<void>
}

export function useOutlookFolderMutationsController(options: FolderMutationsControllerOptions) {
  const {
    cancelCreateFolder,
    closeFolderContextMenu,
    creatingFolderName,
    creatingFolderParentPath,
    folderContextMenu,
    folderNameForPath,
    folderOptions,
    isInDeletedFolder,
    loadCachedFolders,
    manualOutlookDeleteMessage,
    outlookBusy,
    requestLoading,
    selectedFolderPath,
    waitForRequest,
  } = options

  async function createFolder(parentPath = creatingFolderParentPath.value, name = creatingFolderName.value) {
    const folderName = name.trim()
    if (!parentPath || !folderName || outlookBusy.value) return
    const parent = folderOptions.value.find((folder) => folder.folderPath === parentPath)
    if (!parent || !isMailSelectableFolder(parent)) return
    requestLoading.value = true
    try {
      const response = await outlookApi.requestCreateFolder({
        parentFolderPath: parentPath,
        name: folderName,
      })
      await waitForRequest(response)
      await loadCachedFolders()
      cancelCreateFolder()
      requestLoading.value = false
    } catch {
      requestLoading.value = false
    }
  }

  async function deleteFolder(targetPath: string) {
    if (!targetPath || outlookBusy.value) return
    const targetFolder = folderOptions.value.find((folder) => folder.folderPath === targetPath)
    if (!targetFolder || !isMailSelectableFolder(targetFolder)) return
    const targetName = folderOptions.value.find((folder) => folder.folderPath === targetPath)?.label.trim() ?? targetPath
    if (isInDeletedFolder(targetPath)) {
      ElMessage.warning(manualOutlookDeleteMessage)
      return
    }
    const confirmed = window.confirm(`將 Folder「${targetName}」移到 Outlook 刪除資料夾？`)
    if (!confirmed) return
    requestLoading.value = true
    try {
      const response = await outlookApi.requestDeleteFolder({
        folderPath: targetPath,
      })
      await waitForRequest(response)
      await loadCachedFolders()
      if (selectedFolderPath.value === targetPath) {
        selectedFolderPath.value = folderOptions.value[0]?.folderPath ?? ''
      }
      requestLoading.value = false
    } catch {
      requestLoading.value = false
    }
  }

  async function deleteFolderFromContext() {
    const targetPath = folderContextMenu.value.folderPath
    closeFolderContextMenu()
    await deleteFolder(targetPath)
  }

  return {
    createFolder,
    deleteFolderFromContext,
  }
}
