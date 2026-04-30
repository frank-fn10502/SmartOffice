# Office 2016 工作機格式交接

本文件只說明工作機上的 AI 或開發人員需要傳送與接收的目前格式。線上文件入口請看 `docs/ai/office2016-addin-references.md`；實測資料、差異與錯誤回報格式請看 `docs/ai/office2016-test-report.md`。

## 適用範圍

- Office 2016 desktop 是主要目標環境。
- 本文件位於開發機的 `SmartOffice.Hub` repository，只定義 Hub 與工作機 AddIn 的 HTTP/JSON contract。
- 工作機上的完整 SmartOffice solution 會參考 `..\SmartOffice.Hub\SmartOffice.Hub.csproj`；Outlook AddIn 實作請在工作機 SmartOffice solution 中完成，不要在 Hub repository 假裝實作 AddIn。
- Hub 維持本機 HTTP API、SignalR、command routing 與 temporary state。
- Add-in 負責 Office automation、Outlook object model / Office API interaction，以及將實測結果轉成 Hub DTO。
- 目前主要 Office surface 是 Outlook；未來 Word、Excel、PowerPoint 請建立各自的 protocol boundary。

## 工作機交接流程

工作機開始測試前，請先讀：

1. `docs/ai/protocols.md`：Hub route、polling protocol 與 SignalR event。
2. 本文件：目前 Hub contract 與 JSON sample。
3. `Models/Dtos.cs`：目前 C# DTO 的來源；HTTP JSON 預期使用 camelCase field。

建議流程：

1. 在開發機啟動 Hub，確認 `http://localhost:2805/swagger` 可用。
2. 工作機 Add-in 只呼叫 `/api/outlook/poll`、`/api/outlook/push-folders`、`/api/outlook/push-mails`、`/api/outlook/push-rules`、`/api/outlook/push-calendar`、`/api/outlook/admin/log`。
3. Web UI、AI 或 MCP client 只透過 `/api/outlook/request-folders`、`/api/outlook/request-mails`、`/api/outlook/request-rules`、`/api/outlook/request-calendar` enqueue command。
4. 每次 Add-in 收到 command 都記錄 `commandId`、`type`、Office API call、轉換後 JSON sample。
5. 格式不符合，或開發機需要真實資料校準 Web UI、mock、Add-in mapping、檔案寫入或 protocol 時，請依 `docs/ai/office2016-test-report.md` 回傳工作機測試回報。

## Route Prefix

```text
/api/outlook
```

## Enqueue Folder Fetch

Request：

```http
POST /api/outlook/request-folders
```

Response：

```json
{
  "status": "queued"
}
```

Add-in 後續會在 `GET /api/outlook/poll` 收到：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_folders",
  "mailsRequest": null
}
```

## Enqueue Mail Fetch

Request：

```http
POST /api/outlook/request-mails
Content-Type: application/json
```

```json
{
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "range": "1d",
  "maxCount": 10
}
```

Response：

```json
{
  "commandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "status": "queued"
}
```

Add-in 後續會在 `GET /api/outlook/poll` 收到：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_mails",
  "mailsRequest": {
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "range": "1d",
    "maxCount": 10
  }
}
```

`range` 目前預期值：

- `1d`
- `1w`
- `1m`

## Enqueue Rule Fetch

Request：

```http
POST /api/outlook/request-rules
```

Response：

```json
{
  "status": "queued"
}
```

Add-in 後續會在 `GET /api/outlook/poll` 收到：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_rules",
  "mailsRequest": null,
  "calendarRequest": null
}
```

## Enqueue Calendar Fetch

Request：

```http
POST /api/outlook/request-calendar
Content-Type: application/json
```

```json
{
  "daysForward": 14
}
```

Response：

```json
{
  "commandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "status": "queued"
}
```

Add-in 後續會在 `GET /api/outlook/poll` 收到：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_calendar",
  "mailsRequest": null,
  "calendarRequest": {
    "daysForward": 14
  }
}
```

## Enqueue Mail Marker Commands

Request：

```http
POST /api/outlook/request-mark-mail-read
POST /api/outlook/request-mark-mail-unread
POST /api/outlook/request-mark-mail-task
POST /api/outlook/request-clear-mail-task
POST /api/outlook/request-set-mail-categories
Content-Type: application/json
```

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "categories": "Sample Category"
}
```

Add-in 後續會在 `GET /api/outlook/poll` 收到對應 command type：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "set_mail_categories",
  "mailMarkerRequest": {
    "mailId": "[redacted Outlook EntryID or stable id]",
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "categories": "Sample Category"
  }
}
```

目前 mail marker command type：

- `mark_mail_read`
- `mark_mail_unread`
- `mark_mail_task`
- `clear_mail_task`
- `set_mail_categories`

## Enqueue Master Category Commands

Request：

```http
POST /api/outlook/request-upsert-category
Content-Type: application/json
```

```json
{
  "name": "Project",
  "color": "preset4",
  "shortcutKey": ""
}
```

Add-in 後續會在 `GET /api/outlook/poll` 收到：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "upsert_category",
  "categoryRequest": {
    "name": "Project",
    "color": "preset4",
    "shortcutKey": ""
  }
}
```

此 command 修改的是 Outlook master category list，不是單封郵件的 category assignment。工作機 Add-in 應以 category name 找既有 category；存在時更新 color / shortcut key，不存在時新增 category。完成後請重新 push master category list。

## Enqueue Folder / Move Mail Commands

Request：

```http
POST /api/outlook/request-create-folder
Content-Type: application/json
```

```json
{
  "parentFolderPath": "\\\\Mailbox - User\\Projects",
  "name": "Sample Folder"
}
```

Add-in poll command：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "create_folder",
  "createFolderRequest": {
    "parentFolderPath": "\\\\Mailbox - User\\Projects",
    "name": "Sample Folder"
  }
}
```

Request：

```http
POST /api/outlook/request-delete-folder
Content-Type: application/json
```

```json
{
  "folderPath": "\\\\Mailbox - User\\Projects\\Sample Folder"
}
```

Add-in poll command：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "delete_folder",
  "deleteFolderRequest": {
    "folderPath": "\\\\Mailbox - User\\Projects\\Sample Folder"
  }
}
```

Request：

```http
POST /api/outlook/request-move-mail
Content-Type: application/json
```

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
  "destinationFolderPath": "\\\\Mailbox - User\\Projects\\Sample Folder"
}
```

Add-in poll command：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "move_mail",
  "moveMailRequest": {
    "mailId": "[redacted Outlook EntryID or stable id]",
    "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
    "destinationFolderPath": "\\\\Mailbox - User\\Projects\\Sample Folder"
  }
}
```

## Poll No Command

Request：

```http
GET /api/outlook/poll
```

如果 30 秒內沒有 command，Hub 回傳：

```json
{
  "type": "none"
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

實測重點：

- `folderPath` 必須保留 Outlook 實際可再定位的路徑或可穩定比對的路徑。
- `name` 可能包含公司、使用者或客戶資訊；回報時請匿名化。
- `itemCount` 如果 Office API 取不到，請回報正確 API、錯誤訊息與可替代欄位。

## Push Mails

Request：

```http
POST /api/outlook/push-mails
Content-Type: application/json
```

```json
[
  {
    "subject": "[redacted] sample subject",
    "id": "[redacted stable Outlook entry id if available]",
    "senderName": "Sample Sender",
    "senderEmail": "sender@example.invalid",
    "receivedTime": "2026-04-29T09:30:00+08:00",
    "body": "[redacted plain text body]",
    "bodyHtml": "<p>[redacted html body]</p>",
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "categories": "Customer, Follow-up",
    "isRead": false,
    "isMarkedAsTask": true,
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

實測重點：

- `receivedTime` 請回傳 ISO 8601 字串；如果 Office 2016 interop 只能取得 local `DateTime`，請在測試回報寫明 timezone 與轉換方式。
- `body` 與 `bodyHtml` 可能含有敏感 business data；除非已獲授權，測試回報只放已匿名化的最小可用 sample。
- 如果 `MailItem.Body`、`HTMLBody`、`SenderEmailAddress` 或 `ReceivedTime` 在特定帳號型態下行為不同，請附上觀察結果與官方 API link。
- `id` 如果可以取得，建議使用 Outlook `EntryID` 或其他可重新定位 item 的穩定值；如果無法穩定取得，可以留空。
- `categories`、`isRead`、`isMarkedAsTask`、`importance`、`sensitivity` 是 Web UI 與 AI 操作建議需要的 metadata；如果 Office 2016 工作機取不到，請留預設值並在測試回報說明。

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
    "start": "2026-04-30T10:00:00+08:00",
    "end": "2026-04-30T11:00:00+08:00",
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

## Add-in Log

工作機可以用既有 admin log 回報簡短事件。這不是完整測試回報包，只適合 dashboard 即時診斷。

```http
POST /api/outlook/admin/log
Content-Type: application/json
```

```json
{
  "level": "error",
  "message": "fetch_mails failed: MailItem.HTMLBody returned empty for selected folder.",
  "timestamp": "2026-04-29T09:35:00+08:00"
}
```

`level` 目前預期值：

- `info`
- `warn`
- `error`
