# Outlook AddIn 溝通介面

本文件整理目前 `SmartOffice.Hub` 與工作機 Outlook AddIn 的 HTTP/JSON contract。線上文件入口請看 `docs/ai/office2016-addin-references.md`；實測資料、差異與錯誤回報格式請看 `docs/ai/office2016-test-report.md`。

## 適用範圍

- 本 repository 只定義 Hub/Web UI/contract/mock，不包含真正的 Outlook AddIn 實作。
- Outlook AddIn 實作應在工作機完整 SmartOffice solution 中完成。
- Hub 負責 HTTP API、SignalR、command routing 與 temporary in-memory state。
- AddIn 負責 Outlook object model / Office automation，並把結果轉成本文件的 DTO。
- HTTP JSON 使用 ASP.NET Core 預設 camelCase field。
- Mail body、folder name、category name、chat message 都可能含有敏感 business data；測試回報請匿名化。

## 通訊模型

Outlook AddIn 目前使用 polling protocol：

1. Web UI、AI 或 MCP client 呼叫 Hub 的 request endpoint。
2. Hub 將 command 放入 in-memory queue。
3. Outlook AddIn 呼叫 `GET /api/outlook/poll` long-poll 取得一筆 command。
4. AddIn 在本機執行 Outlook automation。
5. AddIn 將結果 push 回 Hub。
6. Hub 更新 cache，並透過 SignalR 通知 Web UI。

目前 Outlook AddIn 沒有連線到 SignalR hub，也沒有使用 SignalR server-to-client method 直接接收 command。SignalR 在目前實作中的角色是 Hub 對 Web UI broadcast cache/status/log 更新；AddIn 與 Hub 的溝通仍是 HTTP `poll`、`push-*` 與 `admin/log`。

這個設計犧牲了 SignalR 雙向即時 command dispatch，但對 Office 2016 / VSTO / 受限企業環境比較保守：AddIn 只需要主動發出 outbound HTTP request，不需要維持 WebSocket 或讓 Hub 連入工作機 process。如果未來確定工作機環境允許穩定 SignalR client 連線，才建議另開一版 AddIn protocol，把 command dispatch 改成 SignalR 並保留 HTTP polling 作 fallback。

Route prefix：

```text
/api/outlook
```

SignalR endpoint：

```text
/hub/notifications
```

## AddIn 需要呼叫的 Endpoint

| Method | Path | 用途 |
| --- | --- | --- |
| `GET` | `/api/outlook/poll` | long-poll 取得 pending command，timeout 為 30 秒 |
| `POST` | `/api/outlook/push-folders` | 回傳 Outlook folder tree |
| `POST` | `/api/outlook/push-mails` | 回傳 mail list |
| `POST` | `/api/outlook/push-rules` | 回傳 Outlook rules |
| `POST` | `/api/outlook/push-categories` | 回傳 Outlook master category list |
| `POST` | `/api/outlook/push-calendar` | 回傳 calendar events |
| `POST` | `/api/outlook/admin/log` | 回報簡短診斷 log |

## Poll Command

Request：

```http
GET /api/outlook/poll
```

沒有 command 時：

```json
{
  "type": "none"
}
```

有 command 時：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_mails",
  "mailsRequest": {
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "range": "1d",
    "maxCount": 10
  },
  "calendarRequest": null,
  "mailMarkerRequest": null,
  "mailPropertiesRequest": null,
  "categoryRequest": null,
  "createFolderRequest": null,
  "deleteFolderRequest": null,
  "moveMailRequest": null
}
```

AddIn 應依 `type` 讀取對應的 request object；其他 request object 可以是 `null` 或缺省。

目前 command type：

| Command type | 對應 request object | 說明 |
| --- | --- | --- |
| `fetch_folders` | 無 | 讀取 folder tree |
| `fetch_mails` | `mailsRequest` | 讀取指定 folder 的 mails |
| `fetch_rules` | 無 | 讀取 Outlook rules |
| `fetch_categories` | 無 | 讀取 Outlook master category list |
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

## Request Object 格式

### FetchMailsRequest

```json
{
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "range": "1d",
  "maxCount": 10
}
```

`range` 目前預期值：

- `1d`
- `1w`
- `1m`

### FetchCalendarRequest

```json
{
  "daysForward": 14
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

`categories` 只在 `set_mail_categories` 使用；其他 marker command 可留空。

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
      "color": "preset4",
      "shortcutKey": ""
    }
  ]
}
```

`flagInterval` 目前預期值：

- `none`
- `today`
- `tomorrow`
- `this_week`
- `next_week`
- `no_date`
- `custom`
- `complete`

`newCategories` 用來讓 AddIn 先建立不存在的 Outlook master category，再套用到單封 mail。

### CategoryCommandRequest

```json
{
  "name": "Project",
  "color": "preset4",
  "shortcutKey": ""
}
```

此 command 修改 Outlook master category list，不是單封 mail 的 category assignment。AddIn 應以 category name 找既有 category；存在時更新 color / shortcut key，不存在時新增。

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

## Push Folders

Request：

```http
POST /api/outlook/push-folders
Content-Type: application/json
```

```json
[
  {
    "name": "Mailbox - User",
    "folderPath": "\\\\Mailbox - User",
    "itemCount": 42,
    "subFolders": [
      {
        "name": "Inbox",
        "folderPath": "\\\\Mailbox - User\\Inbox",
        "itemCount": 18,
        "subFolders": []
      }
    ]
  }
]
```

Response：

```json
{
  "count": 1
}
```

`folderPath` 必須保留 AddIn 後續可再定位或穩定比對的路徑。

## Push Mails

Request：

```http
POST /api/outlook/push-mails
Content-Type: application/json
```

```json
[
  {
    "id": "[redacted Outlook EntryID or stable id]",
    "subject": "[redacted] sample subject",
    "senderName": "Sample Sender",
    "senderEmail": "sender@example.invalid",
    "receivedTime": "2026-05-04T09:30:00+08:00",
    "body": "[redacted plain text body]",
    "bodyHtml": "<p>[redacted html body]</p>",
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

Response：

```json
{
  "count": 1
}
```

注意事項：

- `receivedTime`、task dates 建議回傳 ISO 8601 字串。
- `id` 建議使用 Outlook `EntryID` 或其他可重新定位 item 的穩定值。
- `body` 與 `bodyHtml` 可能含有敏感 business data；除非已獲授權，測試回報只放匿名化 sample。
- `categories` 是 Outlook 對單封 mail 的 category assignment，格式目前是逗號分隔字串。

## Push Rules

Request：

```http
POST /api/outlook/push-rules
Content-Type: application/json
```

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

Response：

```json
{
  "count": 1
}
```

## Push Categories

Request：

```http
POST /api/outlook/push-categories
Content-Type: application/json
```

```json
[
  {
    "name": "Customer",
    "color": "preset4",
    "shortcutKey": ""
  }
]
```

Response：

```json
{
  "count": 1
}
```

`color` 目前是 Hub/Web UI 傳遞用字串；AddIn 實作需在工作機端映射到 Outlook category color。

## Push Calendar

Request：

```http
POST /api/outlook/push-calendar
Content-Type: application/json
```

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

Response：

```json
{
  "count": 1
}
```

## Admin Log

工作機可以用既有 admin log 回報簡短事件。這不是完整測試回報包，只適合 dashboard 即時診斷。

```http
POST /api/outlook/admin/log
Content-Type: application/json
```

```json
{
  "level": "error",
  "message": "fetch_mails failed: MailItem.HTMLBody returned empty for selected folder.",
  "timestamp": "2026-05-04T09:35:00+08:00"
}
```

`level` 目前預期值：

- `info`
- `warn`
- `error`

## Web UI / AI Request Endpoint 摘要

這些 endpoint 主要由 Web UI、AI 或 MCP client 呼叫，用來 enqueue command；AddIn 不需要主動呼叫它們。

| Method | Path | Enqueued command |
| --- | --- | --- |
| `POST` | `/api/outlook/request-folders` | `fetch_folders` |
| `POST` | `/api/outlook/request-mails` | `fetch_mails` |
| `POST` | `/api/outlook/request-rules` | `fetch_rules` |
| `POST` | `/api/outlook/request-categories` | `fetch_categories` |
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

## Cache Read Endpoint 摘要

| Method | Path | 用途 |
| --- | --- | --- |
| `GET` | `/api/outlook/folders` | 讀取 cached folders |
| `GET` | `/api/outlook/mails` | 讀取 cached mails |
| `GET` | `/api/outlook/rules` | 讀取 cached rules |
| `GET` | `/api/outlook/categories` | 讀取 cached categories |
| `GET` | `/api/outlook/calendar` | 讀取 cached calendar events |
| `GET` | `/api/outlook/chat` | 讀取 cached chat messages |
| `POST` | `/api/outlook/chat` | 新增並 broadcast chat message |
| `GET` | `/api/outlook/admin/status` | 讀取 AddIn status |
| `GET` | `/api/outlook/admin/logs` | 讀取 AddIn logs |

## SignalR Event

Hub 在 AddIn poll、push 或寫入 log 後，會 broadcast 下列事件給 Web UI：

| Event | Payload |
| --- | --- |
| `FoldersUpdated` | `FolderDto[]` |
| `MailsUpdated` | `MailItemDto[]` |
| `RulesUpdated` | `OutlookRuleDto[]` |
| `CategoriesUpdated` | `OutlookCategoryDto[]` |
| `CalendarUpdated` | `CalendarEventDto[]` |
| `NewChatMessage` | `ChatMessageDto` |
| `AddinStatus` | `AddinStatusDto` |
| `AddinLog` | `AddinLogEntry[]` |

## DTO 欄位速查

### MailItemDto

- `id`: string
- `subject`: string
- `senderName`: string
- `senderEmail`: string
- `receivedTime`: DateTime
- `body`: string
- `bodyHtml`: string
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
- `itemCount`: number
- `subFolders`: `FolderDto[]`

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
- `color`: string
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

### AddinStatusDto

- `connected`: boolean
- `lastPollTime`: DateTime 或 `null`
- `lastPushTime`: DateTime 或 `null`
- `lastCommand`: string

### AddinLogEntry

- `level`: string，預期 `info`、`warn` 或 `error`
- `message`: string
- `timestamp`: DateTime

## 實作注意事項

- AddIn 應持續 poll；Hub 以最近 90 秒內有 poll 視為 connected。
- 每次 AddIn 完成 fetch 或 command 後，建議 push 最新完整 list，而不是只 push delta。
- `push-*` 目前會取代 Hub cache 內對應資料。
- Folder 與 mail 操作完成後，建議重新 push 受影響的 folders/mails，讓 Web UI 狀態同步。
- Category 操作完成後，建議重新 push master category list；若同時修改 mail category，也要 push mails。
- Hub state 是 process-local memory；重啟 Hub 後 cache、queue、logs 都會清空。
