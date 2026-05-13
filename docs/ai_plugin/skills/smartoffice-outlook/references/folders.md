# Folder API Reference

讀這份文件處理 folder discovery、Inbox 定位、folder children 與 `folderPath` 規則。共通 request/fetch-result envelope 見 `http-api.md`。

## 目錄

- `Folder Path Rules`: 對外 folder path 格式與禁忌。
- `request-folders`: 讀 stores 與 root folders。
- `request-folder-children`: 載入 children 的 request 欄位。
- `request-find-folder`: 以 name/path/type 定位唯一 folder。
- `FolderDto`: folder result 欄位與 Inbox 選取規則。
- `OutlookStoreDto`: store 欄位。

## Folder Path Rules

HTTP API 的 folder path 一律使用普通斜線，例如 `/主要信箱 - User/收件匣`。

不要自行組 folder path；一律使用 folder fetch result 回傳的 `folderPath`。`Inbox` 是範例名稱，不是穩定 path；中文 Outlook 常見路徑是 `/主要信箱 - User/收件匣`。

## `POST /api/outlook/request-folders`

要求 Outlook stores 與 root folders。完成後用 `POST /api/outlook/fetch-result-folders` 讀 `data.stores` 與 `data.folders`。

## `POST /api/outlook/request-folder-children`

Request:

```json
{
  "storeId": "store-id",
  "parentEntryId": "folder-entry-id",
  "parentFolderPath": "/主要信箱 - User",
  "maxDepth": 1,
  "maxChildren": 50
}
```

API 會 clamp `maxDepth` 到 1-3、`maxChildren` 到 1-200，並設定 `reset=false`。

主要 store root 來自 `fetch-result-folders` 的 `data.folders`：

- 主要 store：預設使用 `stores[0]`。
- root folder：同 `storeId`、`isStoreRoot=true` 的 `FolderDto`。
- `request-folder-children.storeId` 使用 root 的 `storeId`。
- `request-folder-children.parentEntryId` 使用 root 的 `entryId`。
- `request-folder-children.parentFolderPath` 使用 root 的 `folderPath`。

注意 request 欄位名稱必須是 `parentEntryId` 與 `parentFolderPath`；`entryId` 與 `folderPath` 是 folder data 欄位，不是此 endpoint 的 request 欄位。

## `POST /api/outlook/request-find-folder`

封裝 folder discovery 與查找。一般 caller 要取得 `folderAAA` 的正式 `folderPath` 時，優先使用這個 endpoint，不需要自行逐層呼叫 `request-folder-children`。

Request:

```json
{
  "name": "folderAAA",
  "folderPath": "",
  "folderType": "",
  "storeId": "",
  "includeHidden": false,
  "maxResults": 20
}
```

若 caller 已經有完整 path，可改用：

```json
{
  "name": "",
  "folderPath": "/主要信箱 - User/Projects/folderAAA",
  "folderType": "",
  "storeId": "",
  "includeHidden": false,
  "maxResults": 20
}
```

若要取得主要 store 的 Inbox，可用：

```json
{
  "name": "",
  "folderPath": "",
  "folderType": "Inbox",
  "storeId": "",
  "includeHidden": false,
  "maxResults": 20
}
```

`storeId` 空白且 `folderType` 有值時，SmartOffice API 只在主要 store 查找該 folder type；若要查特定 store，請填入該 store 的 `storeId`。

完成後讀 `POST /api/outlook/fetch-result-find-folder`。

`fetch-result-find-folder.data` 包含：

- `query`: 本次查找條件。
- `matchCount`: 符合條件的 folder 數。
- `isAmbiguous`: `matchCount > 1` 時為 true；caller 必須請使用者確認。
- `discoveryComplete`: 目前 folder tree 是否已完成可用範圍載入。
- `pendingDiscoveryTargets`: 仍待載入的 folder discovery target 數量。
- `folders`: 候選 `FolderDto[]`，其 `folderPath` 是後續 API 要使用的正式 path。

查找規則：

- `folderPath` 有值時做大小寫不敏感完全比對。
- `folderPath` 空白且 `folderType` 有值時，用 `folderType` 比對，例如 `Inbox`、`Deleted`、`Sent`。
- `folderPath` 與 `folderType` 都空白時，用 `name` 做大小寫不敏感完全比對。
- `storeId` 有值時只搜尋該 store。
- 找不到時停止並回報，不要自行猜 path 或改用 Inbox。
- 找到多筆同名 folder 時，列出必要的 `folderPath` 與 store 顯示名稱請使用者確認。

## `FolderDto`

`name`, `entryId`, `folderPath`, `parentEntryId`, `parentFolderPath`, `itemCount`, `storeId`, `isStoreRoot`, `folderType`, `defaultItemType`, `isHidden`, `isSystem`, `hasChildren`, `childrenLoaded`, `discoveryState`。

`folderType`: `Unknown`, `StoreRoot`, `Mail`, `Inbox`, `Sent`, `Drafts`, `Deleted`, `Junk`, `Archive`, `Outbox`, `SyncIssues`, `Conflicts`, `LocalFailures`, `ServerFailures`, `Calendar`, `Contacts`, `Tasks`, `Notes`, `Journal`, `RssFeeds`, `ConversationHistory`, `ConversationActionSettings`, `OtherSystem`。

主要 Inbox 選取規則：

1. 從 `FolderSnapshotDto.stores[0]` 取得主要 `storeId`。
2. 若主要 store root 的 `childrenLoaded=false`，先用 root `FolderDto` 要求載入 children。
3. 在同一個 `storeId` 底下優先選 `folderType="Inbox"`。
4. 若 `folderType` 不可靠，才 fallback 到 `name="收件匣"` 或 `name="Inbox"`。
5. 後續 request 使用該 folder 的完整 `folderPath`，不要使用 `name` 或自行組路徑。

## `OutlookStoreDto`

`storeId`, `displayName`, `storeKind`, `storeFilePath`, `rootFolderPath`。
