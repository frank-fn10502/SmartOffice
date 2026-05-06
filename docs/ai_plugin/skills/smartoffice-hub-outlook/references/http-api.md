# SmartOffice Outlook HTTP API Reference

## Base URL

Default: `http://localhost:2805`.

所有 endpoint 都在 `/api/outlook` 底下。

## Dispatch Response

一般 `request-*` 回應：

```json
{
  "commandId": "command-id",
  "status": "completed",
  "message": ""
}
```

`request-mail-search` 另有：

```json
{
  "commandId": "parent-command-id",
  "searchId": "search-id",
  "status": "completed",
  "message": "",
  "sliceCount": 3
}
```

常見 `status`：`completed`、`mocked`、`timeout`、`failed`、`folder_cache_unavailable`、`no_searchable_folder`。遇到其他非完成狀態時，回報原始 `status` 與 `message`。

`request-*` response 不是資料本體。取得 `commandId` 後仍要查 `GET /api/outlook/command-results/{commandId}`；完成後讀對應 snapshot endpoint：

- folder request -> `GET /api/outlook/folders`
- mail list / body request -> `GET /api/outlook/mails`
- mail attachment request -> `GET /api/outlook/mail-attachments?mailId={mailId}`
- mail search request -> `GET /api/outlook/mail-search`
- calendar request -> `GET /api/outlook/calendar`
- rules / categories request -> `GET /api/outlook/rules` 或 `GET /api/outlook/categories`

## Command Results

- `GET /api/outlook/command-results/{commandId}`：查單一 command。
- `GET /api/outlook/command-results`：查最近 commands。

`OutlookCommandStatusDto`：

- `commandId`: string
- `type`: string
- `status`: `pending`、`completed`、`failed`，也可能出現其他 API 回傳的狀態字串。
- `success`: boolean 或 null
- `message`: string
- `payload`: string；只當簡短診斷，不要期待完整 mail body。
- `dispatchTimestamp`: DateTime
- `resultTimestamp`: DateTime 或 null

## API Status

- `GET /api/outlook/admin/status`
- `GET /api/outlook/admin/logs`

Status fields：

- `connected`: boolean
- `lastPollTime`: DateTime 或 null
- `lastPushTime`: DateTime 或 null
- `lastCommand`: string

## Cached Snapshot Endpoints

這些 GET 不會送出新 request，只讀 memory cache：

- `GET /api/outlook/folders` -> `FolderSnapshotDto`
- `GET /api/outlook/mails` -> `MailItemDto[]`
- `GET /api/outlook/mail-attachments?mailId={mailId}` -> `MailAttachmentsDto`
- `GET /api/outlook/mail-search` -> `MailItemDto[]`
- `GET /api/outlook/rules` -> `OutlookRuleDto[]`
- `GET /api/outlook/categories` -> `OutlookCategoryDto[]`
- `GET /api/outlook/calendar` -> `CalendarEventDto[]`
- `GET /api/outlook/chat` -> `ChatMessageDto[]`

服務 restart 後 cache 會清空，需要重新 request。

HTTP API 的 folder path 一律使用普通斜線，例如 `/主要信箱 - User/收件匣`。

## Folder Endpoints

### `POST /api/outlook/request-folders`

要求 Outlook stores 與 root folders。完成後讀 `GET /api/outlook/folders`。

### `POST /api/outlook/request-folder-children`

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
若要尋找預設 Inbox，先對主要 store root 呼叫此 endpoint，再從 `GET /api/outlook/folders` 找 `folderType="Inbox"` 或 localized folder name。不要假設 folder path 一定是英文 `/Mailbox - User/Inbox`。

主要 store root 來自 `GET /api/outlook/folders`：

- 主要 store：預設使用 `stores[0]`。
- root folder：同 `storeId`、`isStoreRoot=true` 的 `FolderDto`。
- `request-folder-children.storeId` 使用 root 的 `storeId`。
- `request-folder-children.parentEntryId` 使用 root 的 `entryId`。
- `request-folder-children.parentFolderPath` 使用 root 的 `folderPath`。

## Mail List / Body / Attachment Endpoints

### `POST /api/outlook/request-mails`

Request:

```json
{
  "folderPath": "/主要信箱 - User/收件匣",
  "range": "1m",
  "maxCount": 30
}
```

`folderPath` 必須取自 `GET /api/outlook/folders` snapshot。`range`: `1d`、`1w`、`1m`。完成後讀 `GET /api/outlook/mails`。mail list 只應包含 metadata，完整 body 需另請求。

### `POST /api/outlook/request-mail-body`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/Mailbox - User/Inbox"
}
```

完成後讀 `GET /api/outlook/mails`，找同一封 mail 的 `body` / `bodyHtml`。

### `POST /api/outlook/request-mail-attachments`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/Mailbox - User/Inbox"
}
```

完成後 attachment metadata 會進入 server state；Web UI 會收到更新。

讀取 cached attachment metadata：

```text
GET /api/outlook/mail-attachments?mailId={mailId}
```

### `POST /api/outlook/request-export-mail-attachment`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/Mailbox - User/Inbox",
  "attachmentId": "1",
  "index": 1,
  "name": "sample.pdf",
  "fileName": "sample.pdf",
  "displayName": "sample.pdf",
  "exportRootPath": ""
}
```

`exportRootPath` 空白時 API 會使用目前 attachment export settings。完成後使用 `exportedAttachmentId` 呼叫 open endpoint。

### Attachment Settings / Open

- `GET /api/outlook/attachment-export-settings`
- `POST /api/outlook/attachment-export-settings` with `{ "rootPath": "..." }`
- `POST /api/outlook/open-exported-attachment` with `{ "exportedAttachmentId": "..." }`

只接受已記錄的 exported attachment id，不接受任意本機路徑。

## Mail Search

### `POST /api/outlook/request-mail-search`

Request:

```json
{
  "searchId": "optional-client-search-id",
  "storeId": "",
  "scopeFolderPaths": ["/主要信箱 - User/收件匣"],
  "includeSubFolders": true,
  "keyword": "customer",
  "textFields": ["subject"],
  "categoryNames": [],
  "hasAttachments": null,
  "flagState": "any",
  "readState": "any",
  "receivedFrom": null,
  "receivedTo": null
}
```

- `scopeFolderPaths` 空陣列代表指定 store 或全部 store 內目前已載入的可搜尋 mail folders；AI agent 不可在使用者未要求全域搜尋時送空陣列。
- 使用者未指定 folder 時，使用主要 mailbox 的 Inbox，並設為 `scopeFolderPaths` 第一個值；此值必須取自 folder snapshot，不能硬寫英文 `Inbox`。
- 若回應 `no_searchable_folder`，代表指定 scope 沒有對上目前 cached folders；先重新載入/展開 folders 並改用 snapshot 裡的實際 `folderPath`。
- `textFields`: `subject`、`sender`、`body`；API 會 normalize，不合法時回到 `subject`。
- `flagState`: `any`、`flagged`、`unflagged`。
- `readState`: `any`、`unread`、`read`。
- `hasAttachments`: true / false / null。若只想用「存在附件」搜尋，設定 `hasAttachments=true` 並讓 `keyword` 保持空字串。

Agent 預設 request 範例應像這樣指定單一 Inbox path：

```json
{
  "searchId": "",
  "storeId": "",
  "scopeFolderPaths": ["/主要信箱 - User/收件匣"],
  "includeSubFolders": true,
  "keyword": "customer",
  "textFields": ["subject"],
  "categoryNames": [],
  "hasAttachments": null,
  "flagState": "any",
  "readState": "any",
  "receivedFrom": null,
  "receivedTo": null
}
```

只有使用者明確要求全信箱或全部已載入 mail folders 時，才可使用：

```json
{
  "scopeFolderPaths": []
}
```

使用空 scope 時，回覆使用者必須說明範圍是「目前 Hub folder cache 中已載入的可搜尋 mail folders」，不是保證完整 Outlook mailbox。

Progress：

- `GET /api/outlook/mail-search/progress/{searchId}`
- `GET /api/outlook/mail-search/progress/by-command/{commandId}`

Result：

- `GET /api/outlook/mail-search`

## Calendar / Rules / Categories

- `POST /api/outlook/request-calendar` with `{ "daysForward": 31, "startDate": null, "endDate": null }` -> `GET /api/outlook/calendar`
- `POST /api/outlook/request-rules` -> `GET /api/outlook/rules`
- `POST /api/outlook/request-categories` -> `GET /api/outlook/categories`
- `POST /api/outlook/request-upsert-category` with category object -> wait -> `GET /api/outlook/categories`

`CategoryCommandRequest`：

```json
{
  "name": "Project",
  "color": "olCategoryColorGreen",
  "colorValue": 5,
  "shortcutKey": ""
}
```

## Mail / Folder Mutations

### `POST /api/outlook/request-update-mail-properties`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/Mailbox - User/Inbox",
  "isRead": true,
  "flagInterval": "today",
  "flagRequest": "今天",
  "taskStartDate": null,
  "taskDueDate": null,
  "taskCompletedDate": null,
  "categories": ["Customer"],
  "newCategories": []
}
```

`flagInterval`: `none`、`today`、`tomorrow`、`this_week`、`next_week`、`no_date`、`custom`、`complete`。

### `POST /api/outlook/request-create-folder`

```json
{
  "parentFolderPath": "/Mailbox - User/Projects",
  "name": "Sample Folder"
}
```

### `POST /api/outlook/request-delete-folder`

```json
{
  "folderPath": "/Mailbox - User/Projects/Sample Folder"
}
```

### `POST /api/outlook/request-move-mail`

```json
{
  "mailId": "mail-id",
  "sourceFolderPath": "/Mailbox - User/Inbox",
  "destinationFolderPath": "/Mailbox - User/Projects"
}
```

### `POST /api/outlook/request-delete-mail`

```json
{
  "mailId": "mail-id",
  "folderPath": "/Mailbox - User/Inbox"
}
```

語意是移到 Deleted Items，不是永久刪除。

## Chat

- `POST /api/outlook/chat`
- `GET /api/outlook/chat`

Request:

```json
{
  "source": "web",
  "text": "message"
}
```

Chat text 可能含敏感 business data。

## DTO 速查

### `MailItemDto`

`id`, `subject`, `senderName`, `senderEmail`, `receivedTime`, `body`, `bodyHtml`, `folderPath`, `categories`, `isRead`, `isMarkedAsTask`, `attachmentCount`, `attachmentNames`, `flagRequest`, `flagInterval`, `taskStartDate`, `taskDueDate`, `taskCompletedDate`, `importance`, `sensitivity`。

### `FolderDto`

`name`, `entryId`, `folderPath`, `parentEntryId`, `parentFolderPath`, `itemCount`, `storeId`, `isStoreRoot`, `folderType`, `defaultItemType`, `isHidden`, `isSystem`, `hasChildren`, `childrenLoaded`, `discoveryState`。

`folderType`: `Unknown`, `StoreRoot`, `Mail`, `Inbox`, `Sent`, `Drafts`, `Deleted`, `Junk`, `Archive`, `Outbox`, `SyncIssues`, `Conflicts`, `LocalFailures`, `ServerFailures`, `Calendar`, `Contacts`, `Tasks`, `Notes`, `Journal`, `RssFeeds`, `ConversationHistory`, `ConversationActionSettings`, `OtherSystem`。

主要 Inbox 選取規則：

1. 從 `FolderSnapshotDto.stores[0]` 取得主要 `storeId`。
2. 若主要 store root 的 `childrenLoaded=false`，先用 root `FolderDto` 展開 children。
3. 在同一個 `storeId` 底下優先選 `folderType="Inbox"`。
4. 若 `folderType` 不可靠，才 fallback 到 `name="收件匣"` 或 `name="Inbox"`。
5. 後續 request 使用該 folder 的完整 `folderPath`，不要使用 `name` 或自行組路徑。

### `OutlookStoreDto`

`storeId`, `displayName`, `storeKind`, `storeFilePath`, `rootFolderPath`。

### `CalendarEventDto`

`id`, `subject`, `start`, `end`, `location`, `organizer`, `requiredAttendees`, `isRecurring`, `busyStatus`。
