# Outlook AddIn SignalR 溝通介面

本文件整理工作機 Outlook AddIn 的正式 SignalR-only contract。AddIn 實作順序與驗收請先看 `docs/addin/features-checklist.md`；本文件只保留 SignalR method、request object、DTO 與 payload 細節。

## 適用範圍

- Outlook AddIn 實作應在工作機完整 SmartOffice solution 中完成。
- AddIn 只負責 Outlook object model / Office automation，並把結果轉成本文件的 DTO。
- AddIn 不負責 request endpoint、cache、跨 folder 搜尋排程、progress 推算、資料 merge 或外部 client workflow。
- 不相容舊版 AddIn channel；不要實作或保留 `/api/outlook/poll`、`/api/outlook/push-*`、HTTP chat 或未列於本文件的 legacy command。
- 不維護未使用功能。若 command、欄位或 handler 沒被本 contract 使用，請刪除或不要新增。
- Mail body、folder name、category name、chat message 都可能含有敏感 business data；測試回報請匿名化。

## 通訊模型

1. Outlook AddIn 連線到 `/hub/outlook-addin`。
2. AddIn invoke `RegisterOutlookAddin(info)` 完成註冊。
3. AddIn 透過 SignalR client event `OutlookCommand` 收到 command。
4. AddIn 執行 Outlook automation。
5. AddIn 透過 SignalR server method `Push*`、`ReportAddinLog` 或 `ReportCommandResult` 回報結果。

正式 AddIn endpoint：

```text
/hub/outlook-addin
```

舊的 `/api/outlook/poll` 與 `/api/outlook/push-*` 不再是 AddIn contract，也不需要任何 fallback。

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
  "mailSearchSliceRequest": null,
  "mailBodyRequest": null,
  "calendarRequest": null,
  "mailPropertiesRequest": null,
  "categoryRequest": null,
  "createFolderRequest": null,
  "deleteFolderRequest": null,
  "moveMailRequest": null,
  "deleteMailRequest": null
}
```

目前 command type：

| Command type | 對應 request object | 說明 |
| --- | --- | --- |
| `fetch_folder_roots` | `folderDiscoveryRequest` | 只讀取 Outlook stores 與各 store root folder；不得遞迴 subfolders。 |
| `fetch_folder_children` | `folderDiscoveryRequest` | 只讀取指定 parent folder 的直接 children；Hub 會逐段排程。 |
| `fetch_mails` | `mailsRequest` | 讀取指定 folder 的 mail metadata，不應包含完整 body |
| `fetch_mail_search_slice` | `mailSearchSliceRequest` | 讀取指定單一 folder 的 mail search slice |
| `fetch_mail_body` | `mailBodyRequest` | 使用者點開單封 mail 後，讀取該 mail body |
| `fetch_mail_attachments` | `mailAttachmentsRequest` | 讀取單封 mail 的附件 metadata |
| `export_mail_attachment` | `exportMailAttachmentRequest` | 將指定 attachment 匯出到約定的本機 attachment root |
| `fetch_rules` | 無 | 讀取 Outlook rules |
| `fetch_categories` | 無 | 讀取 Outlook master category list |
| `ping` | 無 | readiness probe；只有 Outlook object model 可正常呼叫時才回成功 |
| `fetch_calendar` | `calendarRequest` | 讀取 calendar events |
| `update_mail_properties` | `mailPropertiesRequest` | 一次更新已讀、flag、category 與新 category |
| `upsert_category` | `categoryRequest` | 新增或更新 master category |
| `create_folder` | `createFolderRequest` | 建立 folder |
| `delete_folder` | `deleteFolderRequest` | 刪除 folder |
| `move_mail` | `moveMailRequest` | 移動單封 mail |
| `delete_mail` | `deleteMailRequest` | 將單封 mail 移到「刪除的郵件 / Deleted Items」folder |

`delete_mail` 是獨立 command；但它的唯一允許實作仍是將 mail 移到 Outlook 的「刪除的郵件 / Deleted Items」folder。AddIn 不得呼叫 Outlook `MailItem.Delete()` 或永久刪除郵件。

`ping` 不是單純 SignalR echo。Hub 會在排隊處理 Web UI / AI request 前先送 `ping`，用來確認 Outlook AddIn 已連線且 Outlook object model 可正常呼叫。若 Outlook 剛啟動、profile 尚未 ready、COM object 暫時 busy，AddIn 應回 `ReportCommandResult(success=false)` 或等到可判斷後再回覆；不要在 Outlook 尚不可操作時回成功。

## AddIn 回報結果

AddIn 可 invoke 下列 server method：

| Method | Payload | 用途 |
| --- | --- | --- |
| `BeginFolderSync` | `FolderSyncBeginDto` | 開始 folder 增量同步 |
| `PushFolderBatch` | `FolderSyncBatchDto` | 推送一批 stores / folders |
| `CompleteFolderSync` | `FolderSyncCompleteDto` | 結束 folder 增量同步 |
| `PushMails` | `MailItemDto[]` | 回推目前 mail snapshot；`fetch_mails` 的回傳應只含 metadata，`body` / `bodyHtml` 留空 |
| `PushMail` | `MailItemDto` | 回推同 id 的單封 mail；用於 `update_mail_properties` 這類不應重抓列表的單封 mutation |
| `BeginMailSearch` | `MailSearchSliceResultDto` | 開始 mail search slice |
| `PushMailSearchSliceResult` | `MailSearchSliceResultDto` | 推送 Outlook 內建搜尋結果 |
| `CompleteMailSearchSlice` | `MailSearchCompleteDto` | 結束 mail search slice |
| `PushMailBody` | `MailBodyDto` | 回推同 id 的 body；用於 `fetch_mail_body` |
| `PushMailAttachments` | `MailAttachmentsDto` | 回推 attachment metadata；用於 `fetch_mail_attachments` |
| `PushExportedMailAttachment` | `ExportedMailAttachmentDto` | 回推已匯出的 attachment path；用於 `export_mail_attachment` |
| `PushRules` | `OutlookRuleDto[]` | 回推 Outlook rules snapshot |
| `PushCategories` | `OutlookCategoryDto[]` | 回推 Outlook master category snapshot |
| `PushCalendar` | `CalendarEventDto[]` | 回推 calendar events snapshot |
| `SendChatMessage` | `ChatMessageDto` | AddIn 透過 SignalR 送出 chat message |
| `ReportAddinLog` | `AddinLogEntry` | 回報診斷 log |
| `ReportCommandResult` | `OutlookCommandResult` | 回報 command 成敗 |

每個 command 完成後，建議至少 invoke `ReportCommandResult`。如果 command 會改變 Outlook snapshot，請同時 invoke 對應 `Push*` method。Folder discovery 只使用 `fetch_folder_roots` 與 `fetch_folder_children`，並透過 `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync` 回推增量結果；AddIn 不得實作一次遞迴整棵樹的 command。Mail search slice 只使用 `BeginMailSearch`、`PushMailSearchSliceResult`、`CompleteMailSearchSlice`，不要用 `PushMails` 覆蓋目前 folder list。單封屬性更新請使用 `PushMail`，不要為了更新一封 mail 重新 `PushMails`。郵件 body 請只在 `fetch_mail_body` 後以 `PushMailBody` 回推。附件採 `fetch_mail_attachments` 先看 metadata、有需要才 `export_mail_attachment` 匯出到本機路徑；AddIn 不負責開檔。

`fetch_folder_roots` 與 `fetch_folder_children` 不應只回 `ReportCommandResult(success=true)`。AddIn 必須在成功結果前至少完成 `PushFolderBatch` 或 `CompleteFolderSync`；如果 Outlook store 尚未 ready、parent folder 無效或列舉結果異常，請回報 `success=false` 與可診斷訊息。

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

`startDate` 含當日，`endDate` 不含當日。`daysForward` 不再作為舊 AddIn fallback；若同時收到 date range 與 `daysForward`，以 `startDate` / `endDate` 為準。

### FolderDiscoveryRequest

`fetch_folder_roots` sample：

```json
{
  "syncId": "folder-sync-001",
  "storeId": "",
  "parentEntryId": "",
  "parentFolderPath": "",
  "maxDepth": 0,
  "maxChildren": 50,
  "reset": true
}
```

`fetch_folder_children` sample：

```json
{
  "syncId": "folder-sync-001",
  "storeId": "[redacted store id]",
  "parentEntryId": "[redacted folder entry id]",
  "parentFolderPath": "\\\\主要信箱 - User\\Inbox",
  "maxDepth": 1,
  "maxChildren": 50,
  "reset": false
}
```

- `fetch_folder_roots` 只允許列出 `Application.Session.Stores`、每個 `Store.GetRootFolder()`，以及必要的 root metadata。
- `fetch_folder_children` 必須使用 `storeId` + `parentEntryId` 優先定位 parent folder；若 `parentEntryId` 空白才可用 `parentFolderPath` fallback。
- `maxDepth` 預設與正式值都是 `1`；AddIn 不得自行遞迴超過 Hub 指定深度。
- `maxChildren` 是單次 command 上限，AddIn 必須 clamp 到合理值。
- 每個 children command 應回推 parent folder 本身，並將該 parent 的 `childrenLoaded=true`、`discoveryState="loaded"`。

### MailSearchSliceRequest

```json
{
  "searchId": "6fb66d3a-7f4f-4a6d-9b3f-7e1e8c2f2d84",
  "commandId": "slice-command-id",
  "parentCommandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "storeId": "[redacted store id]",
  "folderPath": "\\\\主要信箱 - User\\Inbox",
  "keyword": "客戶",
  "textFields": ["subject"],
  "categoryNames": ["Customer"],
  "hasAttachments": true,
  "flagState": "any",
  "readState": "unread",
  "receivedFrom": "2026-05-01T00:00:00+08:00",
  "receivedTo": "2026-05-04T23:59:59+08:00",
  "sliceIndex": 0,
  "sliceCount": 30,
  "resetSearchResults": true,
  "completeSearchOnSlice": false
}
```

- `searchId`: search correlation id；AddIn 回推 slice result 時必須沿用。
- `commandId`: 此 slice command id；AddIn 回推 slice result 時必須沿用。
- `parentCommandId`: 原始 `request-mail-search` 的 command id。
- `storeId`: 指定單一 Outlook Store，必須非空。
- `folderPath`: 指定單一 Outlook folder，必須非空。
- `keyword`: 文字搜尋關鍵字；空白時只套用篩選條件。
- `textFields`: keyword 文字搜尋欄位；目前正式值為 `subject`、`sender`、`body`。預設只有 `subject`。
- `categoryNames`: 分類篩選；任一分類符合即可。
- `hasAttachments`: 附件篩選；`true` 表示包含附件，`false` 表示不含附件，省略表示不限。
- `flagState`: 旗標篩選；`any`、`flagged` 或 `unflagged`。
- `readState`: 已讀篩選；`any`、`unread` 或 `read`。
- `receivedFrom` / `receivedTo`: 收到時間區段，兩者可獨立使用。
- `sliceIndex` / `sliceCount`: folder slice 序號與總數，可用於 progress message。
- `resetSearchResults`: 只有第一個 slice 是 `true`；AddIn 呼叫 `BeginMailSearch` 或第一批 `PushMailSearchSliceResult` 時應沿用。
- `completeSearchOnSlice`: 只有最後一個 slice 是 `true`；AddIn 只有最後一個 slice 才應呼叫 `CompleteMailSearchSlice` 或送 `isFinal=true`。

AddIn 若收到空 `storeId` 或空 `folderPath`，應使用 `CompleteMailSearchSlice(success=false)` 結束該 slice，不得自行展開整個 store 或全域搜尋。

AddIn 必須在指定單一 folder 內依 Microsoft Outlook `AdvancedSearch` / DASL 這類內建搜尋流程，把 `keyword`、`textFields`、分類、附件、旗標、已讀狀態與時間組成 filter，再回傳符合的 metadata。Hub 不做主要 keyword 後篩；這不是 typo-tolerant fuzzy search。

Mail search slice result sample：

```json
{
  "searchId": "6fb66d3a-7f4f-4a6d-9b3f-7e1e8c2f2d84",
  "commandId": "slice-command-id",
  "parentCommandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "sequence": 1,
  "sliceIndex": 0,
  "sliceCount": 30,
  "reset": true,
  "isFinal": false,
  "isSliceComplete": true,
  "mails": [],
  "message": ""
}
```

AddIn 不需要回報 search progress。整體進度由 Hub 依照 dispatch 的 slice、收到的 `PushMailSearchSliceResult` / `CompleteMailSearchSlice` 與 timeout 推算。

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

`mailId` 不可為空。若工作機 AddIn push 回來的 mail 沒有 `id`，後續 mail mutation command 無法可靠執行。

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

AddIn 應用 `mailId` 找到 Outlook item，將 `destinationFolderPath` 解析成 Outlook `Folder` object，呼叫 Outlook `MailItem.Move(destinationFolder)`，完成後回推最新 `PushMails`，並用 folder 增量同步更新 folder count。

若 `destinationFolderPath` 是 Outlook 的「刪除的郵件 / Deleted Items」folder，這仍然只是移動郵件到該 folder，不是永久刪除。AddIn 必須沿用同一個 `MailItem.Move(destinationFolder)` 流程。

注意：Microsoft 文件說 Outlook `MailItem.EntryID` 在 item save 或 send 後才會存在，跨 store 移動時可能改變。因此 AddIn 若使用 EntryID 當 `MailItemDto.id`，移動後應重新讀取並回推最新 mail snapshot。相關官方依據請看 `docs/addin/features-checklist.md`。

### DeleteMailRequest

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox"
}
```

AddIn 應用 `mailId` 找到 Outlook item，再用目前 store 的「刪除的郵件 / Deleted Items」folder 作為 destination，呼叫 Outlook `MailItem.Move(destinationFolder)`。完成後回推最新 `PushMails`，並用 folder 增量同步更新 folder count。這個 command 不得呼叫 Outlook `MailItem.Delete()`。

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

`fetch_mails` 回推的 `MailItemDto` 應只包含 metadata，`body` / `bodyHtml` 留空。收到 `fetch_mail_body` 後，AddIn 再用 `PushMailBody` 回推內容：

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

## Request Endpoint 背景

AddIn 不需要呼叫下列 endpoint。這些 endpoint 只列作 command 來源對照，方便工作機測試時確認「哪個 HTTP request 會變成哪個 `OutlookCommand`」。

| Method | Path | Command |
| --- | --- | --- |
| `POST` | `/api/outlook/request-folders` | `fetch_folder_roots` |
| `POST` | `/api/outlook/request-folder-children` | `fetch_folder_children` |
| `POST` | `/api/outlook/request-mails` | `fetch_mails` |
| `POST` | `/api/outlook/request-mail-body` | `fetch_mail_body` |
| `POST` | `/api/outlook/request-mail-attachments` | `fetch_mail_attachments` |
| `POST` | `/api/outlook/request-export-mail-attachment` | `export_mail_attachment` |
| `POST` | `/api/outlook/open-exported-attachment` | 不 dispatch 給 AddIn |
| `POST` | `/api/outlook/request-rules` | `fetch_rules` |
| `POST` | `/api/outlook/request-categories` | `fetch_categories` |
| `POST` | `/api/outlook/request-signalr-ping` | `ping` |
| `POST` | `/api/outlook/request-calendar` | `fetch_calendar` |
| `POST` | `/api/outlook/request-update-mail-properties` | `update_mail_properties` |
| `POST` | `/api/outlook/request-upsert-category` | `upsert_category` |
| `POST` | `/api/outlook/request-create-folder` | `create_folder` |
| `POST` | `/api/outlook/request-delete-folder` | `delete_folder` |
| `POST` | `/api/outlook/request-move-mail` | `move_mail` |
| `POST` | `/api/outlook/request-delete-mail` | `delete_mail` |

沒有 AddIn SignalR connection 時，Hub 會回 `409 Conflict`：

```json
{
  "commandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "status": "addin_unavailable"
}
```

外部 client 可用下列 endpoint 查詢 command 執行狀態；AddIn 不需要呼叫：

| Method | Path | 說明 |
| --- | --- | --- |
| `GET` | `/api/outlook/command-results/{commandId}` | 查詢指定 command 的 `pending` / `completed` / `failed` / `addin_unavailable` 狀態 |
| `GET` | `/api/outlook/command-results` | 查詢最近 command 執行狀態 |
| `GET` | `/api/outlook/mail-search/progress/{searchId}` | 查詢指定 search id 的進度 |
| `GET` | `/api/outlook/mail-search/progress/by-command/{commandId}` | 用 command id 查詢對應 search 進度 |

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
- `attachmentCount`: number，附件數 metadata；未知時可為 `0`，完整 metadata 仍以 `fetch_mail_attachments` / `PushMailAttachments` 為準。
- `attachmentNames`: string，附件名稱摘要；多個附件名稱建議以 `、` 或 `, ` 串接；避免放入檔案內容或本機路徑。
- `flagRequest`: string
- `flagInterval`: string
- `taskStartDate`: DateTime 或 `null`
- `taskDueDate`: DateTime 或 `null`
- `taskCompletedDate`: DateTime 或 `null`
- `importance`: string，預設 `normal`
- `sensitivity`: string，預設 `normal`

### MailAttachmentDto

AddIn 處理 `fetch_mail_attachments` 時，請從 Outlook `MailItem.Attachments` 逐筆建立 metadata；依 Microsoft Outlook Interop 文件，附件名稱應優先使用 `Attachment.FileName`，沒有檔名時再用 `Attachment.DisplayName`，大小使用 `Attachment.Size`。附件識別請使用同一封 mail 內穩定可 round-trip 的值；Office COM collection 為 1-based，因此可用 `Attachment.Index.ToString()` 作為 `attachmentId`，export 時再用此值取回 `Attachments.Item(index)`。

- `mailId`: string，必須等於 request 的 `mailId`。
- `attachmentId`: string，必填；建議使用 Outlook `Attachment.Index` 的字串值，或 AddIn 自己能在同一封 mail 內穩定查回的 id。
- `index`: number，可填 Outlook `Attachment.Index`，Hub 會在 `attachmentId` 空白時以此補值。
- `name`: string，顯示名稱；建議填 `Attachment.FileName`，沒有時填 `Attachment.DisplayName`。
- `fileName`: string，可填 Outlook `Attachment.FileName`。
- `displayName`: string，可填 Outlook `Attachment.DisplayName`。
- `contentType`: string，可空；Outlook Object Model 沒有直接暴露 MIME type 時不要硬猜。
- `size`: number，可填 Outlook `Attachment.Size`；Microsoft 文件說部分情況可能拿不到實際大小而回 `0`。
- `isExported`: boolean，尚未匯出時為 `false`。
- `exportedAttachmentId`: string，尚未匯出時空白。
- `exportedPath`: string，尚未匯出時空白。

### ExportMailAttachmentRequest

收到 `export_mail_attachment` 時會帶下列欄位。對 Outlook COM/VSTO AddIn，請優先使用 `index` 或可解析為整數的 `attachmentId` 取回 `mail.Attachments.Item(index)`；這是配合 Microsoft Outlook Object Model 的 1-based `Attachments` collection。若 `attachmentId` 不是數字，AddIn 可用自己在 metadata 階段建立的 mapping 查回附件。

- `mailId`: string，目標 mail id。
- `folderPath`: string，目標 mail 所在 folder path。
- `attachmentId`: string；若 metadata 有 `index`，request 會帶 `index.ToString()`，方便 AddIn 直接呼叫 `Attachments.Item(index)`。
- `index`: number；Outlook `Attachment.Index`。
- `name`: string；目前顯示的附件名稱。
- `fileName`: string；metadata 中的 Outlook `Attachment.FileName`。
- `displayName`: string；metadata 中的 Outlook `Attachment.DisplayName`。
- `exportRootPath`: string；允許的 attachment export root。AddIn 輸出檔案必須放在此 root 底下。

### ExportedMailAttachmentDto

AddIn 處理 `export_mail_attachment` 時，請用 request 的 `attachmentId` 找回同一個 Outlook `Attachment`，將檔案儲存到 Hub 約定的 attachment root 底下，呼叫 Outlook `Attachment.SaveAsFile(path)` 後再 `PushExportedMailAttachment`。

- `mailId`: string，必須等於 request 的 `mailId`。
- `folderPath`: string。
- `attachmentId`: string，必須等於 request 的 `attachmentId`。
- `exportedAttachmentId`: string，可由 AddIn 產生；空白時 Hub 會補一個 GUID。
- `name`: string，建議與 metadata 階段相同。
- `fileName`: string，可填 Outlook `Attachment.FileName`。
- `displayName`: string，可填 Outlook `Attachment.DisplayName`。
- `contentType`: string，可空。
- `size`: number，建議填實際輸出檔案長度；拿不到時可用 Outlook `Attachment.Size`。
- `exportedPath`: string，必填；必須是 `SaveAsFile(path)` 實際輸出的完整本機路徑，Hub 之後會用這個路徑開檔。
- `exportedAt`: DateTime。

### FolderDto

- `name`: string
- `entryId`: string，Outlook folder `EntryID`，Hub 後續會搭配 `storeId` 指定 parent。
- `folderPath`: string
- `parentEntryId`: string，store root 可為空字串。
- `parentFolderPath`: string，store root 可為空字串。
- `itemCount`: number
- `storeId`: string，Outlook Store ID 或 AddIn 內可追蹤的 store identifier。
- `isStoreRoot`: boolean，folder 是否是該 store 的 root folder。
- `hasChildren`: boolean，該 folder 是否可能有直接 children。
- `childrenLoaded`: boolean，該 folder 的直接 children 是否已由 Hub 指定 command 載入。
- `discoveryState`: string，預期 `partial`、`loaded` 或 `failed`。

`FolderDto` 不再包含 `subFolders`，也不再重複保存 store display name / type / file path。tree 由 `parentFolderPath` 與 `storeId` 組回。Folder discovery 由 Hub 主導；AddIn 不得保留全量 folder tree command。

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

### AddinLogEntry

- `level`: string，預期 `info`、`warn` 或 `error`
- `message`: string
- `timestamp`: DateTime
