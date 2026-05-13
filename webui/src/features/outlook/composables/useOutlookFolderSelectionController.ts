import { computed } from 'vue'
import type { ComputedRef, Ref } from 'vue'
import type { FolderTreeNode, MailItemDto } from '../models/outlook'

type FolderOption = FolderTreeNode & { label: string }

type FolderSelectionControllerOptions = {
  fetchedMailFolderPath: Ref<string>
  folderContextMenu: Ref<{ visible: boolean; x: number; y: number; folderPath: string }>
  folderOptions: ComputedRef<FolderOption[]>
  mailListMode: Ref<'folder' | 'search'>
  mailSearchResults: Ref<MailItemDto[]>
  selectedFolderPath: Ref<string>
}

export function useOutlookFolderSelectionController(options: FolderSelectionControllerOptions) {
  const {
    fetchedMailFolderPath,
    folderContextMenu,
    folderOptions,
    mailListMode,
    mailSearchResults,
    selectedFolderPath,
  } = options

  const contextFolderName = computed(() => {
    return folderOptions.value.find((folder) => folder.folderPath === folderContextMenu.value.folderPath)?.label.trim() ?? '未選擇'
  })

  const selectedFolderName = computed(() => {
    return folderNameForPath(selectedFolderPath.value)
  })

  const fetchedMailFolderName = computed(() => {
    if (mailListMode.value === 'search') return `搜尋結果：${mailSearchResults.value.length}`
    return fetchedMailFolderPath.value ? folderNameForPath(fetchedMailFolderPath.value) : '尚未抓取郵件'
  })

  function folderNameForPath(path: string) {
    if (!path) return '未選擇'
    return folderOptions.value.find((folder) => folder.folderPath === path)?.label.trim() ?? path
  }

  function inferMailFolderPath(items: MailItemDto[], fallback = '') {
    const paths = [...new Set(items.map((mail) => mail.folderPath).filter(Boolean))]
    return paths.length === 1 ? paths[0] : fallback
  }

  function deletedFolderForPath(path: string) {
    if (!path) return null
    const exactFolder = folderOptions.value.find((folder) => folder.folderPath === path)
    const storeId = exactFolder?.storeId
    return folderOptions.value.find((folder) =>
      folder.folderType === 'Deleted'
      && (!storeId || folder.storeId === storeId)
      && (path === folder.folderPath || path.startsWith(`${folder.folderPath}/`))
    ) ?? null
  }

  function isInDeletedFolder(path: string) {
    return deletedFolderForPath(path) !== null
  }

  return {
    contextFolderName,
    deletedFolderForPath,
    fetchedMailFolderName,
    folderNameForPath,
    inferMailFolderPath,
    isInDeletedFolder,
    selectedFolderName,
  }
}
