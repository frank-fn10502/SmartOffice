# SmartOffice Outlook HTTP API Reference

## Base URL

Default: `http://localhost:2805`.

所有 endpoint 都在 `/api/outlook` 底下。

## Request / Fetch Result Pattern

所有 Outlook 工作都走同一個對外模式：

1. 呼叫 `POST /api/outlook/request-*` 發起動作。
2. 從 response 取得 `requestId`。
3. 持續呼叫該 request 配對的 `POST /api/outlook/fetch-result-*`，直到 `state=completed`，或遇到 `failed`、`unavailable`、`timeout`。

一般 `request-*` response：

```json
{
  "requestId": "request-id",
  "request": "request-mails",
  "state": "accepted",
  "message": "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
  "data": {}
}
```

`request-mail-search` 與 `request-folder-mails` 另有：

```json
{
  "requestId": "request-id",
  "request": "request-mail-search",
  "state": "accepted",
  "message": "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
  "data": {
    "searchId": "search-id"
  }
}
```

`request-*` response 的固定欄位是 `requestId`、`request`、`state`、`message`、`data`。`data` 是該 request 自己的 struct；沒有客製化資訊時是 `{}`。

### `POST /api/outlook/fetch-result-*`

Request:

```json
{
  "requestId": "request-id",
  "cursor": "",
  "take": 100
}
```

Response:

```json
{
  "requestId": "request-id",
  "request": "request-folder-mails",
  "state": "running",
  "message": "",
  "next": {
    "cursor": "100",
    "hasMore": true
  },
  "data": {
    "searchId": "search-id",
    "mails": []
  }
}
```

`fetch-result-*` response 的固定欄位是 `requestId`、`request`、`state`、`message`、`next`、`data`。

- `state`: `accepted`、`running`、`completed`、`failed`、`unavailable`、`timeout`。
- `next.cursor`: 下一段 result 的 cursor。
- `next.hasMore`: 是否還有下一段 result。
- `data`: 該 `fetch-result-*` 自己的 domain struct，例如 `mails`、`folders`、`stores`、`rules`、`categories`、`calendarEvents`。

`take` 只控制此次 `fetch-result-*` response 最多回傳多少筆 result；它不改變原本 `request-*` 的工作範圍。`take` 會 clamp 到 1-500，建議 AI/MCP 使用 100。

若 `request-*` 回 HTTP 409 / 400 / 502 / 504，body 通常仍包含 `state`、`message` 與可能存在的 `requestId`。Caller 應回報該狀態；不要自行擴大 folder scope、改成空 `scopeFolderPaths`，或猜測 folder path 重試。

### Paired Fetch Result Endpoints

| Request endpoint | Fetch result endpoint | 主要 `data` 欄位 |
| --- | --- | --- |
| `POST /api/outlook/request-folders` | `POST /api/outlook/fetch-result-folders` | `stores`, `folders` |
| `POST /api/outlook/request-folder-children` | `POST /api/outlook/fetch-result-folder-children` | `stores`, `folders` |
| `POST /api/outlook/request-mails` | `POST /api/outlook/fetch-result-mails` | `mails` |
| `POST /api/outlook/request-folder-mails` | `POST /api/outlook/fetch-result-folder-mails` | `searchId`, `mails` |
| `POST /api/outlook/request-mail-search` | `POST /api/outlook/fetch-result-mail-search` | `searchId`, `mails` |
| `POST /api/outlook/request-mail-body` | `POST /api/outlook/fetch-result-mail-body` | `mails` |
| `POST /api/outlook/request-mail-attachments` | `POST /api/outlook/fetch-result-mail-attachments` | `mailId`, `folderPath`, `attachments` |
| `POST /api/outlook/request-export-mail-attachment` | `POST /api/outlook/fetch-result-export-mail-attachment` | `{}`；匯出後的 attachment id 目前需從 attachment metadata 讀取 |
| `POST /api/outlook/request-rules` | `POST /api/outlook/fetch-result-rules` | `rules` |
| `POST /api/outlook/request-categories` | `POST /api/outlook/fetch-result-categories` | `categories` |
| `POST /api/outlook/request-calendar` | `POST /api/outlook/fetch-result-calendar` | `calendarEvents` |
| `POST /api/outlook/request-update-mail-properties` | `POST /api/outlook/fetch-result-update-mail-properties` | `mails` |
| `POST /api/outlook/request-upsert-category` | `POST /api/outlook/fetch-result-upsert-category` | `categories` |
| `POST /api/outlook/request-create-folder` | `POST /api/outlook/fetch-result-create-folder` | `stores`, `folders` |
| `POST /api/outlook/request-delete-folder` | `POST /api/outlook/fetch-result-delete-folder` | `stores`, `folders` |
| `POST /api/outlook/request-move-mail` | `POST /api/outlook/fetch-result-move-mail` | `mails` |
| `POST /api/outlook/request-move-mails` | `POST /api/outlook/fetch-result-move-mails` | `mails` |
| `POST /api/outlook/request-delete-mail` | `POST /api/outlook/fetch-result-delete-mail` | `mails` |

## Diagnostic Command Results

`command-results` 是 Hub/AddIn 診斷入口，不是 AI / MCP / Web UI 的主要 workflow。

- `GET /api/outlook/command-results/{commandId}`：查單一內部 command。
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

## Legacy Data Endpoints

這些 GET 保留給診斷與舊工具；正式 caller 優先使用 paired `fetch-result-*`：

- `GET /api/outlook/folders` -> `FolderSnapshotDto`
- `GET /api/outlook/mails` -> `MailItemDto[]`
- `GET /api/outlook/folder-mails` -> `MailItemDto[]`
- `GET /api/outlook/mail-attachments?mailId={mailId}` -> `MailAttachmentsDto`
- `GET /api/outlook/mail-search` -> `MailItemDto[]`
- `GET /api/outlook/rules` -> `OutlookRuleDto[]`
- `GET /api/outlook/categories` -> `OutlookCategoryDto[]`
- `GET /api/outlook/calendar` -> `CalendarEventDto[]`
- `GET /api/outlook/chat` -> `ChatMessageDto[]`

服務 restart 後需要重新送出相關 `request-*` 才會有最新資料。

HTTP API 的 folder path 一律使用普通斜線，例如 `/主要信箱 - User/收件匣`。

## Folder Endpoints

### `POST /api/outlook/request-folders`

要求 Outlook stores 與 root folders。完成後用 `POST /api/outlook/fetch-result-folders` 讀 `data.stores` 與 `data.folders`。

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
若要尋找預設 Inbox，先對主要 store root 呼叫此 endpoint，再從 `fetch-result-folder-children` 的 `data.folders` 找 `folderType="Inbox"` 或 localized folder name。不要假設 folder path 一定是英文 `/Mailbox - User/Inbox`。

主要 store root 來自 `fetch-result-folders` 的 `data.folders`：

- 主要 store：預設使用 `stores[0]`。
- root folder：同 `storeId`、`isStoreRoot=true` 的 `FolderDto`。
- `request-folder-children.storeId` 使用 root 的 `storeId`。
- `request-folder-children.parentEntryId` 使用 root 的 `entryId`。
- `request-folder-children.parentFolderPath` 使用 root 的 `folderPath`。

注意 request 欄位名稱必須是 `parentEntryId` 與 `parentFolderPath`；`entryId` 與 `folderPath` 是 folder data 欄位，不是此 endpoint 的 request 欄位。

## Mail List / Body / Attachment Endpoints

### `POST /api/outlook/request-mails`

Request:

```json
{
  "folderPath": "/主要信箱 - User/收件匣",
  "lookbackHours": 168,
  "maxCount": 30
}
```

`folderPath` 必須取自 folder result 的 `data.folders[].folderPath`。`lookbackHours` 是以小時為單位的簡易相對時間，例如 `12` 代表過去 12 小時、`24` 代表過去 1 天、`168` 代表過去 7 天。也可直接傳入 `receivedFrom` / `receivedTo` date-time 邊界；Hub 會在 dispatch 前補齊給 AddIn。完成後用 `POST /api/outlook/fetch-result-mails` 讀 `data.mails`。mail list 只應包含 metadata，完整 body 需另請求。

### `POST /api/outlook/request-folder-mails`

列出指定 folder 範圍內的所有 mail metadata。這是 Web UI、AI 與 MCP 要做批次操作時的簡單入口；不要用近期 mail list API 來枚舉整個 folder。

Request:

```json
{
  "folderPath": "/主要信箱 - User/Projects/folderA",
  "includeSubFolders": true,
  "receivedFrom": null,
  "receivedTo": null
}
```

完成後讀：

```text
POST /api/outlook/fetch-result-folder-mails
```

`folderPath` 必須取自 folder result 的 `data.folders[].folderPath`。`includeSubFolders=true` 時，Hub 會負責規劃 folder 範圍；caller 不需要理解 Hub 內部如何收集結果。

### `POST /api/outlook/request-mail-body`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/Mailbox - User/Inbox"
}
```

完成後用 `POST /api/outlook/fetch-result-mail-body` 讀 `data.mails`，找同一封 mail 的 `body` / `bodyHtml`。

### `POST /api/outlook/request-mail-attachments`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/Mailbox - User/Inbox"
}
```

完成後用 `POST /api/outlook/fetch-result-mail-attachments` 讀 `data.attachments`。

診斷或舊工具也可讀取 cached attachment metadata：

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

`exportRootPath` 空白時 API 會使用目前 attachment export settings。`fetch-result-export-mail-attachment` 目前只回空 `data`；完成後請重新讀同一封 mail 的 attachment metadata，從 `attachments[].exportedAttachmentId` 取得已匯出的 attachment id，再呼叫 open endpoint。

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
- 使用者未指定 folder 時，使用主要 mailbox 的 Inbox，並設為 `scopeFolderPaths` 第一個值；此值必須取自 folder result 的 `data.folders[].folderPath`，不能硬寫英文 `Inbox`。
- 若 `state=failed` 且 message 或 progress 顯示 `no_searchable_folder`，代表指定 scope 目前無法搜尋；先重新讀 folders 並改用回傳的實際 `folderPath`。
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

使用空 scope 時，回覆使用者必須說明範圍是「目前 SmartOffice API 已知的可搜尋 mail folders」，不是保證完整 Outlook mailbox。

Progress：

- `GET /api/outlook/mail-search/progress/{searchId}`
- `GET /api/outlook/mail-search/progress/by-command/{commandId}`：診斷用；正式 caller 優先讀 paired `fetch-result-*` 的 `data`。

Result：

- `POST /api/outlook/fetch-result-mail-search`

## Calendar / Rules / Categories

- `POST /api/outlook/request-calendar` with `{ "daysForward": 31, "startDate": null, "endDate": null }` -> `POST /api/outlook/fetch-result-calendar`
- `POST /api/outlook/request-rules` -> `POST /api/outlook/fetch-result-rules`
- `POST /api/outlook/request-categories` -> `POST /api/outlook/fetch-result-categories`
- `POST /api/outlook/request-upsert-category` with category object -> `POST /api/outlook/fetch-result-upsert-category`

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

語意是將 folder 移到 Outlook default Deleted Items folder，不是永久刪除。HTTP request 不接受目的 folder；AddIn 必須用 Outlook default folder identity 定位 Deleted Items，不得依賴 `Deleted Items`、`刪除的郵件` 或其他本地化顯示名稱。若目標已經在 default Deleted Items folder 內，paired fetch result 會回 `state=failed`，`message=manual_delete_required`；請使用者自行到 Outlook 永久刪除。

### `POST /api/outlook/request-move-mail`

```json
{
  "mailId": "mail-id",
  "sourceFolderPath": "/Mailbox - User/Inbox",
  "destinationFolderPath": "/Mailbox - User/Projects"
}
```

### `POST /api/outlook/request-move-mails`

單次最多 500 個 `mailIds`；更多郵件必須由 caller 分批呼叫。

```json
{
  "mailIds": ["mail-id-1", "mail-id-2"],
  "sourceFolderPath": "/Mailbox - User/Inbox",
  "sourceFolderPaths": ["/Mailbox - User/Inbox"],
  "destinationFolderPath": "/Mailbox - User/Projects",
  "continueOnError": true
}
```

超過限制時回 `400`：

```json
{
  "status": "too_many_mail_ids",
  "maxBatchSize": 500,
  "actualCount": 8000
}
```

### `POST /api/outlook/request-delete-mail`

```json
{
  "mailId": "mail-id",
  "folderPath": "/Mailbox - User/Inbox"
}
```

語意是移到 Outlook default Deleted Items folder，不是永久刪除。HTTP request 不接受 `destinationFolderPath`；AddIn 必須用 Outlook default folder identity 定位目的 folder，不得用 `Deleted Items`、`刪除的郵件` 或其他本地化顯示名稱猜測。若目標已經在 default Deleted Items folder 內，paired fetch result 會回 `state=failed`，`message=manual_delete_required`；請使用者自行到 Outlook 永久刪除。

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

`id`, `subject`, `sender`, `toRecipients`, `ccRecipients`, `bccRecipients`, `receivedTime`, `body`, `bodyHtml`, `folderPath`, `categories`, `isRead`, `isMarkedAsTask`, `attachmentCount`, `attachmentNames`, `flagRequest`, `flagInterval`, `taskStartDate`, `taskDueDate`, `taskCompletedDate`, `importance`, `sensitivity`。

`sender` 與 recipients 使用 `OutlookRecipientDto`：`recipientKind`, `displayName`, `smtpAddress`, `rawAddress`, `addressType`, `entryUserType`, `isGroup`, `isResolved`, `members`。Web UI / client 應以 `displayName` 作為預設顯示名稱；`rawAddress` 可能是 Exchange legacyDN，不應直接當人名顯示。group 可用 `isGroup=true` 表示，若 AddIn 已展開成員則放在 `members`。

### `FolderDto`

`name`, `entryId`, `folderPath`, `parentEntryId`, `parentFolderPath`, `itemCount`, `storeId`, `isStoreRoot`, `folderType`, `defaultItemType`, `isHidden`, `isSystem`, `hasChildren`, `childrenLoaded`, `discoveryState`。

`folderType`: `Unknown`, `StoreRoot`, `Mail`, `Inbox`, `Sent`, `Drafts`, `Deleted`, `Junk`, `Archive`, `Outbox`, `SyncIssues`, `Conflicts`, `LocalFailures`, `ServerFailures`, `Calendar`, `Contacts`, `Tasks`, `Notes`, `Journal`, `RssFeeds`, `ConversationHistory`, `ConversationActionSettings`, `OtherSystem`。

主要 Inbox 選取規則：

1. 從 `FolderSnapshotDto.stores[0]` 取得主要 `storeId`。
2. 若主要 store root 的 `childrenLoaded=false`，先用 root `FolderDto` 要求載入 children。
3. 在同一個 `storeId` 底下優先選 `folderType="Inbox"`。
4. 若 `folderType` 不可靠，才 fallback 到 `name="收件匣"` 或 `name="Inbox"`。
5. 後續 request 使用該 folder 的完整 `folderPath`，不要使用 `name` 或自行組路徑。

### `OutlookStoreDto`

`storeId`, `displayName`, `storeKind`, `storeFilePath`, `rootFolderPath`。

### `CalendarEventDto`

`id`, `subject`, `start`, `end`, `location`, `organizer`, `requiredAttendees`, `isRecurring`, `busyStatus`。`organizer` 是 `OutlookRecipientDto`，`requiredAttendees` 是 `OutlookRecipientDto[]`。
