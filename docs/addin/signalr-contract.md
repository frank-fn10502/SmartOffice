# Outlook AddIn SignalR 溝通介面

本文件整理目前 `SmartOffice.Hub` 與工作機 Outlook AddIn 的正式 SignalR-only contract。Web UI 每個功能會送出的 command 與工作機實作 checklist 請先看 `docs/addin/features-checklist.md`；本文件只保留 SignalR method、request object、DTO 與 payload 細節。

## 適用範圍

- 本 repository 只定義 Hub/Web UI/contract，不包含真正的 Outlook AddIn 實作。
- Outlook AddIn 實作應在工作機完整 SmartOffice solution 中完成。
- Hub 負責 HTTP API、SignalR、command dispatch 與 temporary in-memory state。
- AddIn 負責 Outlook object model / Office automation，並把結果轉成本文件的 DTO。
- Mail body、folder name、category name、chat message 都可能含有敏感 business data；測試回報請匿名化。

## 通訊模型

1. Outlook AddIn 連線到 `/hub/outlook-addin`。
2. AddIn invoke `RegisterOutlookAddin(info)` 完成註冊。
3. Web UI、AI 或 MCP client 呼叫 Hub 的 HTTP request endpoint。
4. Hub 透過 SignalR client event `OutlookCommand` dispatch command 給 AddIn。
5. AddIn 執行 Outlook automation。
6. AddIn 透過 SignalR server method `Push*`、`ReportAddinLog` 或 `ReportCommandResult` 回報結果。
7. Hub 更新 cache，並透過 `/hub/notifications` 通知 Web UI。

正式 AddIn endpoint：

```text
/hub/outlook-addin
```

Web UI notification endpoint：

```text
/hub/notifications
```

舊的 `/api/outlook/poll` 與 `/api/outlook/push-*` 不再是 AddIn contract。

## AddIn 連線註冊

AddIn 連到 `/hub/outlook-addin` 後，先 invoke：

```text
RegisterOutlookAddin(info)
```

```json
{
  "clientName": "Outlook VSTO AddIn",
  "workstation": "WORKSTATION-01",
  "version": "0.1.0"
}
```

Hub 會將此 connection 放入 Outlook AddIn group，後續 command 會送到這個 group。

## AddIn 接收 Command

AddIn 需要 listen：

```text
OutlookCommand
```

Payload：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_mails",
  "mailsRequest": {
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "range": "1m",
    "maxCount": 30
  },
  "searchMailsRequest": null,
  "mailBodyRequest": null,
  "calendarRequest": null,
  "mailMarkerRequest": null,
  "mailPropertiesRequest": null,
  "categoryRequest": null,
  "createFolderRequest": null,
  "deleteFolderRequest": null,
  "moveMailRequest": null
}
```

目前 command type：

| Command type | 對應 request object | 說明 |
| --- | --- | --- |
| `fetch_folders` | 無 | 讀取 folder tree |
| `fetch_mails` | `mailsRequest` | 讀取指定 folder 的 mail metadata，不應包含完整 body |
| `search_mails` | `searchMailsRequest` | 依 store / folder / 時間 / keyword 搜尋 mail metadata，分批回推結果 |
| `fetch_mail_body` | `mailBodyRequest` | 使用者點開單封 mail 後，讀取該 mail body |
| `fetch_mail_attachments` | `mailAttachmentsRequest` | 使用者或 AI/MCP 要求單封 mail 的附件 metadata |
| `export_mail_attachment` | `exportMailAttachmentRequest` | 將指定 attachment 匯出到 Hub 約定的本機 attachment root |
| `fetch_rules` | 無 | 讀取 Outlook rules |
| `fetch_categories` | 無 | 讀取 Outlook master category list |
| `ping` | 無 | 測試正式 SignalR AddIn channel |
| `fetch_calendar` | `calendarRequest` | 讀取 calendar events |
| `mark_mail_read` | `mailMarkerRequest` | 標記單封 mail 已讀 |
| `mark_mail_unread` | `mailMarkerRequest` | 標記單封 mail 未讀 |
| `mark_mail_task` | `mailMarkerRequest` | 將單封 mail 標記 follow-up/task |
| `clear_mail_task` | `mailMarkerRequest` | 清除單封 mail follow-up/task |
| `set_mail_categories` | `mailMarkerRequest` | 設定單封 mail categories |
| `update_mail_properties` | `mailPropertiesRequest` | 一次更新已讀、flag、category 與新 category |
| `upsert_category` | `categoryRequest` | 新增或更新 master category |
| `create_folder` | `createFolderRequest` | 建立 folder |
| `delete_folder` | `deleteFolderRequest` | 刪除 folder |
| `move_mail` | `moveMailRequest` | 移動單封 mail |

本 Hub contract 不提供 `delete_mail` command。任何 Web UI 上的「刪除郵件」語意都必須實作為 `move_mail` 到 Outlook 的「刪除的郵件 / Deleted Items」folder；AddIn 不得呼叫 Outlook `MailItem.Delete()` 或永久刪除郵件。

## AddIn 回報結果

AddIn 可 invoke 下列 server method：

| Method | Payload | 用途 |
| --- | --- | --- |
| `BeginFolderSync` | `FolderSyncBeginDto` | 開始 folder 增量同步；通常會清空 Hub folder cache 並 broadcast `FolderSyncStarted` |
| `PushFolderBatch` | `FolderSyncBatchDto` | 推送一批 stores / folders，Hub merge 到 cache 並 broadcast `FoldersPatched` |
| `CompleteFolderSync` | `FolderSyncCompleteDto` | 結束 folder 增量同步並 broadcast `FolderSyncCompleted` |
| `PushMails` | `MailItemDto[]` | 取代 cached mails 並 broadcast `MailsUpdated`；`fetch_mails` 的回傳應只含 metadata，`body` / `bodyHtml` 留空 |
| `PushMail` | `MailItemDto` | 更新 cached mails 中同 id 的單封 mail 並 broadcast `MailUpdated`；用於 `update_mail_properties` 這類不應重抓列表的單封 mutation |
| `BeginMailSearch` | `MailSearchBatchDto` | 開始 mail search；通常 reset Hub search result cache 並 broadcast `MailSearchStarted` |
| `PushMailSearchBatch` | `MailSearchBatchDto` | 推送一批 search result metadata，Hub merge 到 search cache 並 broadcast `MailSearchPatched` |
| `CompleteMailSearch` | `MailSearchCompleteDto` | 結束 mail search 並 broadcast `MailSearchCompleted` |
| `PushMailBody` | `MailBodyDto` | 更新 cached mails 中同 id 的 body 並 broadcast `MailBodyUpdated`；用於 `fetch_mail_body` |
| `PushMailAttachments` | `MailAttachmentsDto` | 更新 cached attachment metadata 並 broadcast `MailAttachmentsUpdated`；用於 `fetch_mail_attachments` |
| `PushExportedMailAttachment` | `ExportedMailAttachmentDto` | 記錄已匯出的 attachment path 並 broadcast `MailAttachmentExported`；用於 `export_mail_attachment` |
| `PushRules` | `OutlookRuleDto[]` | 取代 cached rules 並 broadcast `RulesUpdated` |
| `PushCategories` | `OutlookCategoryDto[]` | 取代 cached categories 並 broadcast `CategoriesUpdated` |
| `PushCalendar` | `CalendarEventDto[]` | 取代 cached calendar events 並 broadcast `CalendarUpdated` |
| `SendChatMessage` | `ChatMessageDto` | AddIn 透過 SignalR 送出 chat message，Hub 會 broadcast `NewChatMessage` |
| `ReportAddinLog` | `AddinLogEntry` | 回報診斷 log |
| `ReportCommandResult` | `OutlookCommandResult` | 回報 command 成敗 |

每個 command 完成後，建議至少 invoke `ReportCommandResult`。如果 command 會改變畫面資料，請同時 invoke 對應 `Push*` method。Folder tree 只使用 `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync`。Search result 只使用 `BeginMailSearch`、`PushMailSearchBatch`、`CompleteMailSearch`，不要用 `PushMails` 覆蓋目前 folder list。單封屬性更新請使用 `PushMail`，不要為了更新一封 mail 重新 `PushMails`。郵件 body 請只在 `fetch_mail_body` 後以 `PushMailBody` 回推。附件採 `fetch_mail_attachments` 先看 metadata、有需要才 `export_mail_attachment` 匯出到本機路徑；Web UI Host 會透過 `/api/outlook/open-exported-attachment` 開啟已匯出的檔案，AddIn 不負責開檔。

`fetch_folders` 不應只回 `ReportCommandResult(success=true)`。AddIn 必須在成功結果前至少完成 `PushFolderBatch` 或 `CompleteFolderSync`；如果 Outlook store 尚未 ready 或列舉結果為 0，請回報 `success=false` 與可診斷訊息，避免 Hub/Web UI 把空 folder tree 視為有效成功同步。

AddIn 不應使用 HTTP `/api/outlook/chat` 送 chat；請改用 `/hub/outlook-addin` 上的 `SendChatMessage(message)`。

`OutlookCommandResult` sample：

```json
{
  "commandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "success": true,
  "message": "fetch_mails completed",
  "payload": "",
  "timestamp": "2026-05-04T09:30:06+08:00"
}
```

## Request Object 格式

### FetchMailsRequest

```json
{
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "range": "1m",
  "maxCount": 30
}
```

`range` 目前預期值：

- `1d`
- `1w`
- `1m`

### FetchCalendarRequest

```json
{
  "daysForward": 31,
  "startDate": "2026-05-01",
  "endDate": "2026-06-01"
}
```

Web UI 的月曆介面會帶目前月份的 `startDate` / `endDate`。`startDate` 含當日，`endDate` 不含當日；`daysForward` 保留作為舊 AddIn fallback。

### SearchMailsRequest

```json
{
  "searchId": "6fb66d3a-7f4f-4a6d-9b3f-7e1e8c2f2d84",
  "storeId": "mock-store-primary",
  "scopeFolderPaths": ["\\\\主要信箱 - User\\Inbox"],
  "includeSubFolders": true,
  "keyword": "客戶xxxx",
  "matchMode": "contains",
  "fields": ["subject", "sender"],
  "receivedFrom": "2026-05-01T00:00:00+08:00",
  "receivedTo": "2026-05-04T23:59:59+08:00",
  "exactReceivedTime": null,
  "exactReceivedToleranceSeconds": 60,
  "maxCount": 50
}
```

- `searchId`: Web UI / AI 產生的 search correlation id；AddIn 回推 batch 時必須沿用。
- `storeId`: 指定單一 Outlook Store。空字串代表全域搜尋，AddIn 必須依 store / folder 分批查找，不得一次做高成本全量搜尋。
- `scopeFolderPaths`: 指定 folder scope；空陣列代表使用整個 store 或全域分批。
- `includeSubFolders`: true 時包含 scope folder 底下子 folder。
- `keyword`: 片段關鍵字，可為空；空值代表只用時間 / folder / store 條件查找。
- `matchMode`: `contains`、`exact`、`regex`。`regex` 僅能對 bounded result 做後篩，不應造成全量 body 掃描。
- `fields`: `subject`、`sender`、`categories`、`body`。body search 必須受 `storeId`、folder/time scope 與 `maxCount` 約束。
- `receivedFrom` / `receivedTo`: 時間區段。
- `exactReceivedTime`: 單一時間查找；若有值，AddIn 可優先用它與 `exactReceivedToleranceSeconds` 篩選。
- `maxCount`: 回傳上限；AddIn 端應再 clamp，建議不超過 200。

Search 回推 sample：

```json
{
  "searchId": "6fb66d3a-7f4f-4a6d-9b3f-7e1e8c2f2d84",
  "sequence": 1,
  "reset": true,
  "isFinal": false,
  "mails": [],
  "message": ""
}
```

### MailMarkerCommandRequest

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "categories": "Customer, Follow-up"
}
```

`mailId` 不可為空。Web UI 以 `MailItemDto.id` 填入此欄位；工作機 AddIn 的 `PushMails` 必須提供 Outlook `MailItem.EntryID` 或其他可由 AddIn 找回該 mail 的穩定識別。`categories` 只在 `set_mail_categories` 使用；其他 marker command 可留空。

### MailPropertiesCommandRequest

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "isRead": true,
  "flagInterval": "today",
  "flagRequest": "今天",
  "taskStartDate": "2026-05-04T00:00:00+08:00",
  "taskDueDate": "2026-05-04T00:00:00+08:00",
  "taskCompletedDate": null,
  "categories": ["Customer", "Follow-up"],
  "newCategories": [
    {
      "name": "Customer",
      "color": "olCategoryColorGreen",
      "colorValue": 5,
      "shortcutKey": ""
    }
  ]
}
```

`mailId` 不可為空。若工作機 AddIn push 回來的 mail 沒有 `id`，Web UI 會停用修改與移動，避免送出 `missing mail id` 的 command。

`flagInterval` 目前預期值：

- `none`
- `today`
- `tomorrow`
- `this_week`
- `next_week`
- `no_date`
- `custom`
- `complete`

### CategoryCommandRequest

```json
{
  "name": "Project",
  "color": "olCategoryColorGreen",
  "colorValue": 5,
  "shortcutKey": ""
}
```

### CreateFolderRequest

```json
{
  "parentFolderPath": "\\\\Mailbox - User\\Projects",
  "name": "Sample Folder"
}
```

### DeleteFolderRequest

```json
{
  "folderPath": "\\\\Mailbox - User\\Projects\\Sample Folder"
}
```

### MoveMailRequest

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
  "destinationFolderPath": "\\\\Mailbox - User\\Projects\\Sample Folder"
}
```

`move_mail` 只有在目前 mail 有非空 `id` 時才會由 Web UI 送出。AddIn 應用 `mailId` 找到 Outlook item，將 `destinationFolderPath` 解析成 Outlook `Folder` object，呼叫 Outlook `MailItem.Move(destinationFolder)`，完成後回推最新 `PushMails`，並用 folder 增量同步更新 folder count。

若 `destinationFolderPath` 是 Outlook 的「刪除的郵件 / Deleted Items」folder，這仍然只是移動郵件到該 folder，不是永久刪除。AddIn 必須沿用同一個 `MailItem.Move(destinationFolder)` 流程。

注意：Microsoft 文件說 Outlook `MailItem.EntryID` 在 item save 或 send 後才會存在，跨 store 移動時可能改變。因此 AddIn 若使用 EntryID 當 `MailItemDto.id`，移動後應重新讀取並回推最新 mail snapshot。相關官方依據與 Web UI 操作對照請看 `docs/addin/features-checklist.md`。

## Push Payload Sample

### FolderSyncBatchDto

```json
{
  "syncId": "folder-sync-001",
  "sequence": 1,
  "reset": true,
  "isFinal": false,
  "stores": [
    {
      "storeId": "[redacted primary store id]",
      "displayName": "主要信箱 - User",
      "storeKind": "ost",
      "storeFilePath": "C:\\Users\\User\\AppData\\Local\\Microsoft\\Outlook\\user@example.com.ost",
      "rootFolderPath": "\\\\主要信箱 - User"
    }
  ],
  "folders": [
    {
      "name": "主要信箱 - User",
      "folderPath": "\\\\主要信箱 - User",
      "parentFolderPath": "",
      "itemCount": 0,
      "storeId": "[redacted primary store id]",
      "isStoreRoot": true
    },
    {
      "name": "Inbox",
      "folderPath": "\\\\主要信箱 - User\\Inbox",
      "parentFolderPath": "\\\\主要信箱 - User",
      "itemCount": 18,
      "storeId": "[redacted primary store id]",
      "isStoreRoot": false
    }
  ]
}
```

### MailItemDto

```json
[
  {
    "id": "[redacted Outlook EntryID or stable id]",
    "subject": "[redacted] sample subject",
    "senderName": "Sample Sender",
    "senderEmail": "sender@example.invalid",
    "receivedTime": "2026-05-04T09:30:00+08:00",
    "body": "",
    "bodyHtml": "",
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "categories": "Customer, Follow-up",
    "isRead": false,
    "isMarkedAsTask": true,
    "flagRequest": "今天",
    "flagInterval": "today",
    "taskStartDate": "2026-05-04T00:00:00+08:00",
    "taskDueDate": "2026-05-04T00:00:00+08:00",
    "taskCompletedDate": null,
    "importance": "high",
    "sensitivity": "normal"
  }
]
```

`fetch_mails` 回推的 `MailItemDto` 應只包含列表與屬性面板需要的 metadata，`body` / `bodyHtml` 留空。使用者點開單封 mail 後，Web UI 會送 `fetch_mail_body`，AddIn 再用 `PushMailBody` 回推內容：

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "body": "[redacted plain text body]",
  "bodyHtml": "<p>[redacted html body]</p>"
}
```

### OutlookRuleDto

```json
[
  {
    "name": "Move customer mail",
    "enabled": true,
    "executionOrder": 1,
    "ruleType": "receive",
    "conditions": ["sender contains example.com"],
    "actions": ["move to \\\\Mailbox - User\\Customers"],
    "exceptions": []
  }
]
```

### OutlookCategoryDto

```json
[
  {
    "name": "Customer",
    "color": "olCategoryColorGreen",
    "colorValue": 5,
    "shortcutKey": ""
  }
]
```

### CalendarEventDto

```json
[
  {
    "id": "[redacted appointment id if available]",
    "subject": "[redacted meeting subject]",
    "start": "2026-05-04T10:00:00+08:00",
    "end": "2026-05-04T11:00:00+08:00",
    "location": "Meeting Room",
    "organizer": "Sample Organizer",
    "requiredAttendees": "Sample Attendee",
    "isRecurring": false,
    "busyStatus": "busy"
  }
]
```

### ChatMessageDto

```json
{
  "id": "[optional client-generated id]",
  "source": "outlook",
  "text": "AddIn message",
  "timestamp": "2026-05-04T10:00:00+08:00"
}
```

AddIn 透過 `SendChatMessage(message)` 送出時，Hub 會覆寫 `timestamp` 為收到訊息的時間；`source` 空白時會設為 `outlook`。

## Web UI / AI Request Endpoint 摘要

這些 endpoint 由 Web UI、AI 或 MCP client 呼叫；Hub 收到後會 dispatch `OutlookCommand` 給 SignalR AddIn。

| Method | Path | Command |
| --- | --- | --- |
| `POST` | `/api/outlook/request-folders` | `fetch_folders` |
| `POST` | `/api/outlook/request-mails` | `fetch_mails` |
| `POST` | `/api/outlook/request-mail-body` | `fetch_mail_body` |
| `POST` | `/api/outlook/request-mail-attachments` | `fetch_mail_attachments` |
| `POST` | `/api/outlook/request-export-mail-attachment` | `export_mail_attachment` |
| `POST` | `/api/outlook/open-exported-attachment` | Hub Host 開啟已匯出 attachment，不 dispatch 給 AddIn |
| `POST` | `/api/outlook/request-rules` | `fetch_rules` |
| `POST` | `/api/outlook/request-categories` | `fetch_categories` |
| `POST` | `/api/outlook/request-signalr-ping` | `ping` |
| `POST` | `/api/outlook/request-calendar` | `fetch_calendar` |
| `POST` | `/api/outlook/request-mark-mail-read` | `mark_mail_read` |
| `POST` | `/api/outlook/request-mark-mail-unread` | `mark_mail_unread` |
| `POST` | `/api/outlook/request-mark-mail-task` | `mark_mail_task` |
| `POST` | `/api/outlook/request-clear-mail-task` | `clear_mail_task` |
| `POST` | `/api/outlook/request-set-mail-categories` | `set_mail_categories` |
| `POST` | `/api/outlook/request-update-mail-properties` | `update_mail_properties` |
| `POST` | `/api/outlook/request-upsert-category` | `upsert_category` |
| `POST` | `/api/outlook/request-create-folder` | `create_folder` |
| `POST` | `/api/outlook/request-delete-folder` | `delete_folder` |
| `POST` | `/api/outlook/request-move-mail` | `move_mail` |

沒有 AddIn SignalR connection 時，Hub 會回 `409 Conflict`：

```json
{
  "commandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "status": "addin_unavailable"
}
```

AI / MCP client 可用下列 endpoint 查詢 command 執行狀態：

| Method | Path | 說明 |
| --- | --- | --- |
| `GET` | `/api/outlook/command-results/{commandId}` | 查詢指定 command 的 `pending` / `completed` / `failed` / `addin_unavailable` 狀態 |
| `GET` | `/api/outlook/command-results` | 查詢最近 command 執行狀態 |

## DTO 欄位速查

### MailItemDto

- `id`: string
- `subject`: string
- `senderName`: string
- `senderEmail`: string
- `receivedTime`: DateTime
- `body`: string，`fetch_mails` 時應留空；只在單封內容載入後填入。
- `bodyHtml`: string，`fetch_mails` 時應留空；只在單封內容載入後填入。
- `folderPath`: string
- `categories`: string
- `isRead`: boolean
- `isMarkedAsTask`: boolean
- `flagRequest`: string
- `flagInterval`: string
- `taskStartDate`: DateTime 或 `null`
- `taskDueDate`: DateTime 或 `null`
- `taskCompletedDate`: DateTime 或 `null`
- `importance`: string，預設 `normal`
- `sensitivity`: string，預設 `normal`

### FolderDto

- `name`: string
- `folderPath`: string
- `parentFolderPath`: string，store root 可為空字串。
- `itemCount`: number
- `storeId`: string，Outlook Store ID 或 AddIn 內可追蹤的 store identifier。
- `isStoreRoot`: boolean，folder 是否是該 store 的 root folder。

`FolderDto` 不再包含 `subFolders`，也不再重複保存 store display name / type / file path。tree 由 `parentFolderPath` 與 `storeId` 組回。

### OutlookStoreDto

- `storeId`: string，Outlook Store ID 或 AddIn 內可追蹤的 store identifier。
- `displayName`: string，Outlook store 顯示名稱。
- `storeKind`: string，目前預期 `ost`、`pst`、`exchange` 或 `other`。
- `storeFilePath`: string，`.pst` 或 `.ost` 的完整檔案路徑；沒有檔案路徑時可空字串。
- `rootFolderPath`: string，該 store root folder 的 `folderPath`。

### OutlookRuleDto

- `name`: string
- `enabled`: boolean
- `executionOrder`: number
- `ruleType`: string，預設 `receive`
- `conditions`: string[]
- `actions`: string[]
- `exceptions`: string[]

### OutlookCategoryDto

- `name`: string
- `color`: string，Outlook `OlCategoryColor` enum name，例如 `olCategoryColorGreen`。
- `colorValue`: number，Outlook `OlCategoryColor` enum numeric value，例如 `5`。Add-in 寫入 Outlook 時應優先使用此欄位。
- `shortcutKey`: string

### CalendarEventDto

- `id`: string
- `subject`: string
- `start`: DateTime
- `end`: DateTime
- `location`: string
- `organizer`: string
- `requiredAttendees`: string
- `isRecurring`: boolean
- `busyStatus`: string

### ChatMessageDto

- `id`: string，可空；未填時 Hub DTO 會產生預設 id。
- `source`: string，AddIn 建議填 `outlook`。
- `text`: string
- `timestamp`: DateTime，AddIn 送出時可空；Hub 會覆寫為收到訊息的時間。

### AddinStatusDto

- `connected`: boolean
- `lastPollTime`: DateTime 或 `null`；SignalR-only 後代表最近 AddIn connection time。
- `lastPushTime`: DateTime 或 `null`
- `lastCommand`: string

### AddinLogEntry

- `level`: string，預期 `info`、`warn` 或 `error`
- `message`: string
- `timestamp`: DateTime
