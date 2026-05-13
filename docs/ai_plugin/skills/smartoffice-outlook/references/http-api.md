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

`request-folder-mails` 另有：

```json
{
  "requestId": "request-id",
  "request": "request-folder-mails",
  "state": "accepted",
  "message": "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
  "data": {
    "folderMailsId": "folder-mails-id"
  }
}
```

`request-mail-search` 另有：

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
    "folderMailsId": "folder-mails-id",
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

### Request Body 格式錯誤

SmartOffice API 不接受未文件化欄位。若 request body 使用錯欄位名，會回 HTTP 400：

```json
{
  "status": "invalid_request_body",
  "state": "failed",
  "message": "Request JSON does not match this endpoint schema. Remove unknown fields and use the exact property names documented in Swagger.",
  "errors": {
    "$.query": [
      "The JSON property 'query' could not be mapped to any .NET member contained in type 'SmartOffice.Hub.Contracts.SearchMailsRequest'."
    ]
  }
}
```

遇到 `invalid_request_body` 時，修正欄位名稱後重送同一 endpoint；不要改成其他 endpoint，也不要擴大搜尋範圍。常見錯誤：

- `request-mail-search` 使用 `keyword`，不是 `query`、`text` 或 `searchText`。
- `request-mail-search.scopeFolderPaths` 必須是 string array，例如 `["/主要信箱 - User/收件匣"]`。
- `request-calendar` 使用 `daysForward` 或 `startDate` / `endDate`，不是 `lookaheadDays` / `lookbackDays`。
- `request-folder-children` 使用 `parentEntryId` / `parentFolderPath`，不是 folder data 裡的 `entryId` / `folderPath`。
- mail body、attachment、conversation、mutation endpoints 都需要從 `data.mails[]` 取得的 `mailId` 與同筆 `folderPath`。

缺少必要欄位時會回：

```json
{
  "status": "missing_required_fields",
  "state": "failed",
  "message": "Missing required request field(s): folderPath.",
  "requiredFields": ["folderPath"]
}
```

遇到 `missing_required_fields` 時，只補齊 `requiredFields` 列出的欄位後重送。

### Paired Fetch Result Endpoints

| Request endpoint | Fetch result endpoint | 主要 `data` 欄位 |
| --- | --- | --- |
| `POST /api/outlook/request-folders` | `POST /api/outlook/fetch-result-folders` | `stores`, `folders` |
| `POST /api/outlook/request-folder-children` | `POST /api/outlook/fetch-result-folder-children` | `stores`, `folders` |
| `POST /api/outlook/request-find-folder` | `POST /api/outlook/fetch-result-find-folder` | `query`, `matchCount`, `isAmbiguous`, `folders` |
| `POST /api/outlook/request-mails` | `POST /api/outlook/fetch-result-mails` | `mails` |
| `POST /api/outlook/request-folder-mails` | `POST /api/outlook/fetch-result-folder-mails` | `folderMailsId`, `mails` |
| `POST /api/outlook/request-mail-search` | `POST /api/outlook/fetch-result-mail-search` | `searchId`, `mails` |
| `POST /api/outlook/request-mail-body` | `POST /api/outlook/fetch-result-mail-body` | `mails` |
| `POST /api/outlook/request-mail-attachments` | `POST /api/outlook/fetch-result-mail-attachments` | `mailId`, `folderPath`, `attachments` |
| `POST /api/outlook/request-mail-conversation` | `POST /api/outlook/fetch-result-mail-conversation` | `mailId`, `folderPath`, `conversationId`, `conversationTopic`, `mails` |
| `POST /api/outlook/request-export-mail-attachment` | `POST /api/outlook/fetch-result-export-mail-attachment` | `{}`；匯出後的 attachment id 目前需從 attachment metadata 讀取 |
| `POST /api/outlook/request-rules` | `POST /api/outlook/fetch-result-rules` | `rules` |
| `POST /api/outlook/request-categories` | `POST /api/outlook/fetch-result-categories` | `categories` |
| `POST /api/outlook/request-calendar` | `POST /api/outlook/fetch-result-calendar` | `calendarEvents` |
| `POST /api/outlook/request-address-book` | `POST /api/outlook/fetch-result-address-book` | `contacts` |
| `POST /api/outlook/request-update-mail-properties` | `POST /api/outlook/fetch-result-update-mail-properties` | `mails` |
| `POST /api/outlook/request-upsert-category` | `POST /api/outlook/fetch-result-upsert-category` | `categories` |
| `POST /api/outlook/request-create-folder` | `POST /api/outlook/fetch-result-create-folder` | `stores`, `folders` |
| `POST /api/outlook/request-delete-folder` | `POST /api/outlook/fetch-result-delete-folder` | `stores`, `folders` |
| `POST /api/outlook/request-move-mail` | `POST /api/outlook/fetch-result-move-mail` | `mails` |
| `POST /api/outlook/request-move-mails` | `POST /api/outlook/fetch-result-move-mails` | `mails` |
| `POST /api/outlook/request-delete-mail` | `POST /api/outlook/fetch-result-delete-mail` | `mails` |

## Diagnostic Command Results

`command-results` 是 SmartOffice API / AddIn 診斷入口，不是一般 client 的主要 workflow。

- `GET /api/outlook/command-results/{commandId}`：查單一 request 狀態。
- `GET /api/outlook/command-results`：查最近 commands。

`OutlookCommandStatusDto`：

- `commandId`: string
- `type`: string
- `status`: `pending`、`completed`、`failed`，也可能出現其他 API 回傳的狀態字串。
- `success`: boolean 或 null
- `message`: string
- `payload`: string；只當簡短診斷，不要期待完整 mail body。
- `createdAt`: DateTime
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
- `GET /api/outlook/mail-conversation?mailId={mailId}` -> `MailConversationDto`
- `GET /api/outlook/mail-search` -> `MailItemDto[]`
- `GET /api/outlook/rules` -> `OutlookRuleDto[]`
- `GET /api/outlook/categories` -> `OutlookCategoryDto[]`
- `GET /api/outlook/calendar` -> `CalendarEventDto[]`
- `GET /api/outlook/address-book?query={text}&take={count}` -> `{ query, totalCount, contacts }`
- `GET /api/outlook/address-book/lookup?email={email}` -> `{ query, state, message, contact, suggestions }`
- `GET /api/outlook/chat` -> `ChatMessageDto[]`

服務 restart 後需要重新送出相關 `request-*` 才會有最新資料。

HTTP API 的 folder path 一律使用普通斜線，例如 `/主要信箱 - User/收件匣`。

## Address Book Cache

通訊錄 cache 是 Hub 的關聯視圖。資料來源包含 cached mails 的 sender / to / cc / bcc / group members、cached calendar events 的 organizer / attendees，以及 `request-address-book` 從 Outlook Contacts folder / AddressLists / GAL 同步回來的 metadata。它只暴露 metadata、mail ids 與少量 subject sample，不會讀取或回傳完整 mail body。

使用方式：

- 想檢查一個收件者是否和使用者有已知互動：呼叫 `GET /api/outlook/address-book/lookup?email={email}`。
- `state=known` 代表目前 Hub cache 找到 mail 或 calendar 關聯；`state=unknown` 只代表目前 cache 未知，不代表 Outlook 裡一定沒有。
- 若使用者要求同步真正 Outlook 通訊錄，呼叫 `POST /api/outlook/request-address-book`，再用 `POST /api/outlook/fetch-result-address-book` 讀 `data.contacts`。
- `contact.relationKinds` 會指出 `sender`、`to`、`cc`、`bcc`、`organizer`、`attendee` 或 `group_member` 等關聯。
- `contact.isLikelySelf=true` 代表該地址看起來是自己的寄件地址；Hub 主要從 Sent folder 的 sender 推斷。

`request-address-book` body：

```json
{
  "includeOutlookContacts": true,
  "includeAddressLists": true,
  "maxContacts": 1000,
  "maxAddressEntriesPerList": 500
}
```

`maxContacts` 與 `maxAddressEntriesPerList` 是負載上限；不要要求無限制 GAL 枚舉。

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

### `POST /api/outlook/request-find-folder`

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

完成後讀：

```text
POST /api/outlook/fetch-result-find-folder
```

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

`folderPath` 必須取自 folder result 的 `data.folders[].folderPath`。`lookbackHours` 是以小時為單位的簡易相對時間，例如 `12` 代表過去 12 小時、`24` 代表過去 1 天、`168` 代表過去 7 天。也可直接傳入 `receivedFrom` / `receivedTo` date-time 邊界。完成後用 `POST /api/outlook/fetch-result-mails` 讀 `data.mails`。mail list 只應包含 metadata，完整 body 需另請求。

### `POST /api/outlook/request-folder-mails`

列出指定 folder 範圍內的所有 mail metadata。這是 client 要做批次操作時的簡單入口；不要用近期 mail list API 來枚舉整個 folder。

Request:

```json
{
  "folderPath": "/主要信箱 - User/Projects/folderA",
  "includeSubFolders": true,
  "receivedFrom": null,
  "receivedTo": null,
  "maxCount": 500
}
```

完成後讀：

```text
POST /api/outlook/fetch-result-folder-mails
```

`folderPath` 必須取自 folder result 的 `data.folders[].folderPath`。AI agent 預設應使用 `includeSubFolders=true`，讓指定 folder 底下的子資料夾也納入範圍；只有使用者明確排除 subfolders 時才設為 `false`。`maxCount` 會套用到每個 folder slice，Hub / AddIn 會 clamp 到 1-500；需要搬空大型 folder tree 時應分批規劃，不要一次要求無上限。這是直接列出 folder mails 的 API，不是文字搜尋；不要改用 `request-mail-search` 取代。

`fetch-result-folder-mails.data` 包含：

- `folderMailsId`: 本次 folder mails request 的 id。
- `mails`: 指定 folder 範圍內的 mail metadata；每筆 `id` 是後續 move / delete / update 使用的 mail id。

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

診斷時也可讀取 attachment metadata：

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

`request-mail-search` 是 mail 操作的主要定位與篩選 API。Caller 可用它先取得符合條件的 `data.mails[].id` 與 `folderPath`，再進行 body 讀取、附件讀取、category 更新、move 或 delete。它不只用於文字搜尋，也用於日期、category、附件、已讀與旗標條件。

Request:

```json
{
  "searchId": "optional-client-search-id",
  "storeId": "",
  "scopeFolderPaths": ["/主要信箱 - User/收件匣"],
  "allowGlobalScope": false,
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

- `scopeFolderPaths` 空陣列且 `storeId` 有值時，代表指定 store 內目前已載入的可搜尋 mail folders。
- `scopeFolderPaths` 空陣列且 `storeId` 也空白時，必須同時設定 `allowGlobalScope=true`，代表全部目前已載入的可搜尋 mail folders；AI agent 不可在使用者未要求全域搜尋時設定此值。
- 使用者未指定 folder 時，使用主要 mailbox 的 Inbox，並設為 `scopeFolderPaths` 第一個值；此值必須取自 folder result 的 `data.folders[].folderPath`，不能硬寫英文 `Inbox`。預設 `includeSubFolders=true`，只有使用者明確排除 subfolders 時才設為 `false`。
- 若 `state=failed` 且 message 或 progress 顯示 `no_searchable_folder`，代表指定 scope 目前無法搜尋；先重新讀 folders 並改用回傳的實際 `folderPath`。
- `textFields`: `subject`、`sender`、`body`；API 會 normalize，不合法時回到 `subject`。
- `textFields` 包含 `body` 且 `keyword` 有值時，會使用 Outlook 內容搜尋。其他 subject、sender、category、attachment、flag、read state 與 received time 條件屬於 metadata filter。
- `categoryNames`: Outlook category names；多個值代表任一 category 符合即可。只用 category 搜尋時，讓 `keyword` 保持空字串。
- `flagState`: `any`、`flagged`、`unflagged`。
- `readState`: `any`、`unread`、`read`。
- `hasAttachments`: true / false / null。若只想用「存在附件」搜尋，設定 `hasAttachments=true` 並讓 `keyword` 保持空字串。

Agent 預設 request 範例應像這樣指定單一 Inbox path：

```json
{
  "searchId": "",
  "storeId": "",
  "scopeFolderPaths": ["/主要信箱 - User/收件匣"],
  "allowGlobalScope": false,
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
  "scopeFolderPaths": [],
  "allowGlobalScope": true
}
```

使用空 scope 時，回覆使用者必須說明範圍是「目前 SmartOffice API 已知的可搜尋 mail folders」，不是保證完整 Outlook mailbox。

常用 request 型態：

```json
{
  "scopeFolderPaths": ["/主要信箱 - User/folderA"],
  "includeSubFolders": true,
  "keyword": "",
  "textFields": ["subject"],
  "categoryNames": ["待處理"],
  "hasAttachments": null,
  "flagState": "any",
  "readState": "any",
  "receivedFrom": null,
  "receivedTo": null
}
```

上述代表在 folderA 與其 subfolders 中搜尋 category 為 `待處理` 的 mails，可接續用 `request-move-mails` 做批次搬移。

Progress：

- `GET /api/outlook/mail-search/progress/{searchId}`
- `GET /api/outlook/mail-search/progress/by-command/{commandId}`：診斷用；正式 caller 優先讀 paired `fetch-result-*` 的 `data`。

Result：

- `POST /api/outlook/fetch-result-mail-search`

## Calendar / Rules / Categories

- `POST /api/outlook/request-calendar` with `{ "daysForward": 31, "startDate": null, "endDate": null }` -> `POST /api/outlook/fetch-result-calendar`
- `POST /api/outlook/request-rules` -> `POST /api/outlook/fetch-result-rules`
- `POST /api/outlook/request-manage-rule` with `OutlookRuleCommandRequest` -> `POST /api/outlook/fetch-result-manage-rule`
- `POST /api/outlook/request-categories` -> `POST /api/outlook/fetch-result-categories`
- `POST /api/outlook/request-upsert-category` with category object -> `POST /api/outlook/fetch-result-upsert-category`

`OutlookRuleCommandRequest` 常用欄位：

```json
{
  "operation": "upsert",
  "storeId": "",
  "ruleName": "客戶郵件標記",
  "originalRuleName": "",
  "originalExecutionOrder": null,
  "ruleType": "receive",
  "enabled": true,
  "executionOrder": null,
  "conditions": {
    "subjectContains": ["報價"],
    "bodyContains": [],
    "bodyOrSubjectContains": [],
    "messageHeaderContains": [],
    "senderAddressContains": ["example.com"],
    "recipientAddressContains": [],
    "categories": ["客戶"],
    "hasAttachment": true,
    "importance": "high",
    "toMe": false,
    "toOrCcMe": false,
    "onlyToMe": false,
    "meetingInviteOrUpdate": false
  },
  "actions": {
    "moveToFolderPath": "\\\\主要信箱 - User\\Inbox\\客戶",
    "copyToFolderPath": "",
    "assignCategories": ["客戶"],
    "clearCategories": false,
    "markAsTask": true,
    "markAsTaskInterval": "this_week",
    "delete": false,
    "desktopAlert": true,
    "stopProcessingMoreRules": true
  }
}
```

`importance` 可用 `any`、`low`、`normal`、`high`；`markAsTaskInterval` 可用 `today`、`tomorrow`、`this_week`、`next_week`、`no_date`。`hasAttachment=false` 不支援，因 Outlook Rules object model 只能建立「有附件」條件。

`CategoryCommandRequest`：

```json
{
  "name": "Project",
  "color": "olCategoryColorGreen",
  "colorValue": 5,
  "shortcutKey": ""
}
```

常用 Outlook category color：

| 使用者顏色 | `color` | `colorValue` |
| --- | --- | --- |
| 無色 | `olCategoryColorNone` | `0` |
| 紅色 | `olCategoryColorRed` | `1` |
| 橘色 | `olCategoryColorOrange` | `2` |
| 桃色 | `olCategoryColorPeach` | `3` |
| 黃色 | `olCategoryColorYellow` | `4` |
| 綠色 | `olCategoryColorGreen` | `5` |
| 青色 | `olCategoryColorTeal` | `6` |
| 橄欖 | `olCategoryColorOlive` | `7` |
| 藍色 | `olCategoryColorBlue` | `8` |
| 紫色 | `olCategoryColorPurple` | `9` |
| 栗色 | `olCategoryColorMaroon` | `10` |
| 鋼藍 | `olCategoryColorSteel` | `11` |
| 深鋼藍 | `olCategoryColorDarkSteel` | `12` |
| 灰色 | `olCategoryColorGray` | `13` |
| 深灰 | `olCategoryColorDarkGray` | `14` |
| 黑色 | `olCategoryColorBlack` | `15` |

若使用者要求黑色 category，request body 應使用：

```json
{
  "name": "xxxx",
  "color": "olCategoryColorBlack",
  "colorValue": 15,
  "shortcutKey": ""
}
```

Agent 處理 category 時應先用 `request-categories` / `fetch-result-categories` 讀取 master category list。Category name 比對大小寫不敏感；若已有同名 category，`request-upsert-category` 視為更新該 category。使用者指定顏色不在上表時，請使用者改選，不要猜 `OlCategoryColor` enum 或 numeric value。套用 category 到 mail 前，必須先從 mail fetch result 定位唯一 `mailId` 與同筆 mail 的 `folderPath`。

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

語意是將 folder 移到 Outlook default Deleted Items folder，不是永久刪除。HTTP request 不接受目的 folder；AddIn 必須用 Outlook default folder identity 定位 Deleted Items，不得依賴 `Deleted Items`、`刪除的郵件` 或其他本地化顯示名稱。完成後告知使用者 folder 已移到刪除資料夾。若目標 folder 已經位於 default Deleted Items folder 或其子層，paired fetch result 會回 `state=failed` / `message=manual_delete_required`；agent 必須停止，並請使用者自行到 Outlook 操作。

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

語意是移到 Outlook default Deleted Items folder，不是永久刪除。HTTP request 不接受 `destinationFolderPath`；AddIn 必須用 Outlook default folder identity 定位目的 folder，不得用 `Deleted Items`、`刪除的郵件` 或其他本地化顯示名稱猜測。完成後告知使用者 mail 已移到刪除資料夾；若使用者要永久刪除，請使用者自行到 Outlook 操作。

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

`sender` 與 recipients 使用 `OutlookRecipientDto`：`recipientKind`, `displayName`, `smtpAddress`, `rawAddress`, `addressType`, `entryUserType`, `isGroup`, `isResolved`, `members`。client 應以 `displayName` 作為預設顯示名稱；`rawAddress` 可能是 Exchange legacyDN，不應直接當人名顯示。group 可用 `isGroup=true` 表示，若 AddIn 已展開成員則放在 `members`。

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
