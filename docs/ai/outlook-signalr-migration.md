# Outlook AddIn SignalR 過渡說明

本文件說明舊 HTTP polling protocol 與新 SignalR-only protocol 的差異，目標是讓工作機 Outlook AddIn 一次性切換到 SignalR，不再保留舊連線方式。

## 結論

新的正式 AddIn channel 是：

```text
/hub/outlook-addin
```

工作機 AddIn 不應再呼叫：

- `GET /api/outlook/poll`
- `POST /api/outlook/push-folders`
- `POST /api/outlook/push-mails`
- `POST /api/outlook/push-rules`
- `POST /api/outlook/push-categories`
- `POST /api/outlook/push-calendar`

Web UI、AI 或 MCP client 對 Hub 的 request endpoint 仍保留，因為它們是外部 caller 的穩定入口；改變的是 Hub 與 Outlook AddIn 之間的 command/result channel。

## 新舊差異

| 面向 | 舊 HTTP polling | 新 SignalR-only |
| --- | --- | --- |
| AddIn 連線 | AddIn 週期性 `GET /api/outlook/poll` | AddIn 啟動後連 `/hub/outlook-addin` |
| Command dispatch | Hub queue command，AddIn 下次 poll 才取到 | Hub 立即送 `OutlookCommand` client event |
| Result 回報 | AddIn 呼叫 `POST /api/outlook/push-*` | AddIn invoke `PushFolders`、`PushMails` 等 server method |
| Online 判斷 | 用最近 poll 時間推測 | 用 SignalR connection lifecycle |
| 無 AddIn 時 | command 留在 queue | request endpoint 直接回 `409 addin_unavailable` |
| Web UI 等待 | 送 request 後反覆讀 cache | 送 request 後等待 SignalR update event |
| Hub 狀態 | 需要 `CommandQueue` | 不需要 command queue |

## AddIn 端一次性修改清單

1. 移除 `/api/outlook/poll` 迴圈。
2. 移除所有 `/api/outlook/push-*` HTTP 呼叫。
3. 新增 ASP.NET Core SignalR client，連到 Hub 的 `/hub/outlook-addin`。
4. 連線成功後 invoke `RegisterOutlookAddin(info)`。
5. listen `OutlookCommand`。
6. 依 `command.type` 執行既有 Outlook automation。
7. 用 SignalR server method 回報結果：
   - folder list：`PushFolders(folders)`
   - mail list：`PushMails(mails)`
   - rule list：`PushRules(rules)`
   - master category list：`PushCategories(categories)`
   - calendar events：`PushCalendar(events)`
   - log：`ReportAddinLog(entry)`
   - command 成敗：`ReportCommandResult(result)`
8. 實作 reconnect 後重新 invoke `RegisterOutlookAddin(info)`。

## Method 對照

| 舊 HTTP AddIn 呼叫 | 新 SignalR AddIn 行為 |
| --- | --- |
| `GET /api/outlook/poll` | listen `OutlookCommand` |
| `POST /api/outlook/push-folders` | invoke `PushFolders(folders)` |
| `POST /api/outlook/push-mails` | invoke `PushMails(mails)` |
| `POST /api/outlook/push-rules` | invoke `PushRules(rules)` |
| `POST /api/outlook/push-categories` | invoke `PushCategories(categories)` |
| `POST /api/outlook/push-calendar` | invoke `PushCalendar(events)` |
| `POST /api/outlook/admin/log` | invoke `ReportAddinLog(entry)` |

## Command Payload

新 protocol 沿用 `PendingCommand` payload，不需要重寫 command mapping。

AddIn listen：

```text
OutlookCommand
```

Sample：

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

## Register

AddIn 連線後 invoke：

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

## Result 與 Log

每個 command 完成後建議至少 invoke `ReportCommandResult(result)`，讓 dashboard 能看到成功或失敗。

```json
{
  "commandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "success": true,
  "message": "fetch_mails completed",
  "payload": "",
  "timestamp": "2026-05-04T09:30:06+08:00"
}
```

如果 command 會更新畫面資料，請同時 invoke 對應 `Push*` method。Hub 會更新 cache 並 broadcast Web UI event。

## Reconnect

AddIn 應使用 SignalR automatic reconnect。每次 reconnect 成功後都要重新 invoke `RegisterOutlookAddin(info)`，確保 Hub 重新把 connection 放入 Outlook AddIn group。

## 測試順序

1. 連 `/hub/outlook-addin`。
2. invoke `RegisterOutlookAddin(info)`，Web UI Admin 應顯示 AddIn online。
3. 在 Web UI Admin 按 `SignalR Ping`，AddIn 應收到 `type: "ping"` 的 `OutlookCommand`。
4. AddIn invoke `ReportCommandResult(result)` 回 pong，Web UI Admin logs 應顯示結果。
5. 在 Web UI 按 `Fetch Folders`，AddIn 應收到 `fetch_folders`。
6. AddIn invoke `PushFolders(folders)`，Web UI 應即時更新 folder tree。
7. 依序測 `fetch_mails`、`fetch_rules`、`fetch_categories`、`fetch_calendar` 與 mail/folder 操作 command。
