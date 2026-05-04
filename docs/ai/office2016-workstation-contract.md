# Outlook AddIn SignalR 溝通介面

本文件整理目前 `SmartOffice.Hub` 與工作機 Outlook AddIn 的正式 SignalR-only contract。HTTP polling 到 SignalR 的差異與過渡方式請看 `docs/ai/outlook-signalr-migration.md`。

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

目前 command type：

| Command type | 對應 request object | 說明 |
| --- | --- | --- |
| `fetch_folders` | 無 | 讀取 folder tree |
| `fetch_mails` | `mailsRequest` | 讀取指定 folder 的 mails |
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

## AddIn 回報結果

AddIn 可 invoke 下列 server method：

| Method | Payload | 用途 |
| --- | --- | --- |
| `PushFolders` | `FolderDto[]` | 取代 cached folders 並 broadcast `FoldersUpdated` |
| `PushMails` | `MailItemDto[]` | 取代 cached mails 並 broadcast `MailsUpdated` |
| `PushRules` | `OutlookRuleDto[]` | 取代 cached rules 並 broadcast `RulesUpdated` |
| `PushCategories` | `OutlookCategoryDto[]` | 取代 cached categories 並 broadcast `CategoriesUpdated` |
| `PushCalendar` | `CalendarEventDto[]` | 取代 cached calendar events 並 broadcast `CalendarUpdated` |
| `ReportAddinLog` | `AddinLogEntry` | 回報診斷 log |
| `ReportCommandResult` | `OutlookCommandResult` | 回報 command 成敗 |

每個 command 完成後，建議至少 invoke `ReportCommandResult`。如果 command 會改變畫面資料，請同時 invoke 對應 `Push*` method。

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

### CategoryCommandRequest

```json
{
  "name": "Project",
  "color": "preset4",
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

## Push Payload Sample

### FolderDto

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

### MailItemDto

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
    "color": "preset4",
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

## Web UI / AI Request Endpoint 摘要

這些 endpoint 由 Web UI、AI 或 MCP client 呼叫；Hub 收到後會 dispatch `OutlookCommand` 給 SignalR AddIn。

| Method | Path | Command |
| --- | --- | --- |
| `POST` | `/api/outlook/request-folders` | `fetch_folders` |
| `POST` | `/api/outlook/request-mails` | `fetch_mails` |
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
- `lastPollTime`: DateTime 或 `null`；SignalR-only 後代表最近 AddIn connection time。
- `lastPushTime`: DateTime 或 `null`
- `lastCommand`: string

### AddinLogEntry

- `level`: string，預期 `info`、`warn` 或 `error`
- `message`: string
- `timestamp`: DateTime
