# SmartOffice.Hub Outlook HTTP API Reference

## Base URL

Use `${SMARTOFFICE_HUB_URL:-http://localhost:2805}`.

所有 endpoint 都在 `/api/outlook` 底下。外部 AI agent 只呼叫 HTTP API；SignalR `/hub/outlook-addin` 是工作機 Outlook AddIn 的 channel。

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

常見 `status`：`completed`、`mocked`、`timeout`、`failed`、`addin_unavailable`、`folder_cache_unavailable`、`no_searchable_folder`。

## Command Results

- `GET /api/outlook/command-results/{commandId}`：查單一 command。
- `GET /api/outlook/command-results`：查最近 commands。

`OutlookCommandStatusDto`：

- `commandId`: string
- `type`: string
- `status`: `pending`、`completed`、`failed`、`addin_unavailable`
- `success`: boolean 或 null
- `message`: string
- `payload`: string；只當簡短診斷，不要期待完整 mail body。
- `dispatchTimestamp`: DateTime
- `resultTimestamp`: DateTime 或 null

## Diagnostics

- `GET /api/outlook/admin/status` -> `AddinStatusDto`
- `GET /api/outlook/admin/logs` -> `AddinLogEntry[]`
- `POST /api/outlook/request-signalr-ping` -> dispatch `ping`

`AddinStatusDto`：

- `connected`: boolean
- `lastPollTime`: DateTime 或 null
- `lastPushTime`: DateTime 或 null
- `lastCommand`: string

## Cached Snapshot Endpoints

這些 GET 不會觸發 Outlook automation，只讀 Hub memory cache：

- `GET /api/outlook/folders` -> `FolderSnapshotDto`
- `GET /api/outlook/mails` -> `MailItemDto[]`
- `GET /api/outlook/mail-attachments?mailId={mailId}` -> `MailAttachmentsDto`
- `GET /api/outlook/mail-search` -> `MailItemDto[]`
- `GET /api/outlook/rules` -> `OutlookRuleDto[]`
- `GET /api/outlook/categories` -> `OutlookCategoryDto[]`
- `GET /api/outlook/calendar` -> `CalendarEventDto[]`
- `GET /api/outlook/chat` -> `ChatMessageDto[]`

Hub restart 後 cache 會清空，需要重新 request。

## Folder Endpoints

### `POST /api/outlook/request-folders`

要求 Outlook stores 與 root folders。完成後讀 `GET /api/outlook/folders`。

### `POST /api/outlook/request-folder-children`

Request:

```json
{
  "storeId": "store-id",
  "parentEntryId": "folder-entry-id",
  "parentFolderPath": "\\\\Mailbox - User\\Inbox",
  "maxDepth": 1,
  "maxChildren": 50
}
```

Hub 會 clamp `maxDepth` 到 1-3、`maxChildren` 到 1-200，並設定 `reset=false`。

## Mail List / Body / Attachment Endpoints

### `POST /api/outlook/request-mails`

Request:

```json
{
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "range": "1m",
  "maxCount": 30
}
```

`range`: `1d`、`1w`、`1m`。完成後讀 `GET /api/outlook/mails`。mail list 只應包含 metadata，完整 body 需另請求。

### `POST /api/outlook/request-mail-body`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "\\\\Mailbox - User\\Inbox"
}
```

完成後讀 `GET /api/outlook/mails`，找同一封 mail 的 `body` / `bodyHtml`。

### `POST /api/outlook/request-mail-attachments`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "\\\\Mailbox - User\\Inbox"
}
```

完成後 attachment metadata 會進入 Hub state；Web UI 會收到 SignalR 更新。

讀取 cached attachment metadata：

```text
GET /api/outlook/mail-attachments?mailId={mailId}
```

### `POST /api/outlook/request-export-mail-attachment`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "attachmentId": "1",
  "index": 1,
  "name": "sample.pdf",
  "fileName": "sample.pdf",
  "displayName": "sample.pdf",
  "exportRootPath": ""
}
```

`exportRootPath` 空白時 Hub 會使用目前 attachment export settings。完成後使用 `exportedAttachmentId` 呼叫 open endpoint。

### Attachment Settings / Open

- `GET /api/outlook/attachment-export-settings`
- `POST /api/outlook/attachment-export-settings` with `{ "rootPath": "..." }`
- `POST /api/outlook/open-exported-attachment` with `{ "exportedAttachmentId": "..." }`

只接受 Hub 已記錄的 exported attachment id，不接受任意本機路徑。

## Mail Search

### `POST /api/outlook/request-mail-search`

Request:

```json
{
  "searchId": "optional-client-search-id",
  "storeId": "",
  "scopeFolderPaths": ["\\\\Mailbox - User\\Inbox"],
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

- `scopeFolderPaths` 空陣列代表指定 store 或全部 store 內目前已載入的可搜尋 mail folders。
- 使用者未指定 folder 時，建議先從 `GET /api/outlook/folders` 選擇主要 mailbox 的 Inbox，並設為 `scopeFolderPaths` 第一個值。
- `textFields`: `subject`、`sender`、`body`；Hub 會 normalize，不合法時回到 `subject`。
- `flagState`: `any`、`flagged`、`unflagged`。
- `readState`: `any`、`unread`、`read`。
- `hasAttachments`: true / false / null。若只想用「存在附件」搜尋，設定 `hasAttachments=true` 並讓 `keyword` 保持空字串。

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
  "folderPath": "\\\\Mailbox - User\\Inbox",
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
  "parentFolderPath": "\\\\Mailbox - User\\Projects",
  "name": "Sample Folder"
}
```

### `POST /api/outlook/request-delete-folder`

```json
{
  "folderPath": "\\\\Mailbox - User\\Projects\\Sample Folder"
}
```

### `POST /api/outlook/request-move-mail`

```json
{
  "mailId": "mail-id",
  "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
  "destinationFolderPath": "\\\\Mailbox - User\\Projects"
}
```

### `POST /api/outlook/request-delete-mail`

```json
{
  "mailId": "mail-id",
  "folderPath": "\\\\Mailbox - User\\Inbox"
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

### `OutlookStoreDto`

`storeId`, `displayName`, `storeKind`, `storeFilePath`, `rootFolderPath`。

### `CalendarEventDto`

`id`, `subject`, `start`, `end`, `location`, `organizer`, `requiredAttendees`, `isRecurring`, `busyStatus`。
