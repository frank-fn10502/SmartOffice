# Protocol Notes for AI Agents

Office 2016 AddIn 功能 checklist 請參考 `docs/addin/features-checklist.md`。工作機傳送與接收格式請參考 `docs/addin/signalr-contract.md`。Office 2016 AddIn 線上文件請參考 `docs/addin/outlook-references.md`。工作機測試回報格式請參考 `docs/addin/test-report.md`。

## AddIn Protocol

Outlook AddIn 正式 protocol 已改為 SignalR-only：

1. Outlook AddIn 連線到 `/hub/outlook-addin`。
2. AddIn invoke `RegisterOutlookAddin(info)` 完成註冊。
3. Web UI、AI 或 MCP client 透過 Hub HTTP request endpoint 發出要求。
4. Hub 透過 SignalR client event `OutlookCommand` 即時 dispatch command 給 AddIn。
5. AddIn 在本機執行 Outlook automation。
6. AddIn 透過 SignalR server method `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync`、`PushMails`、`PushRules`、`PushCategories`、`PushCalendar`、`SendChatMessage`、`ReportAddinLog` 或 `ReportCommandResult` 回報結果。
7. Hub 更新 cache，並透過 `/hub/notifications` broadcast 給 Web UI。

目前不保留舊 AddIn HTTP long-poll / push channel；工作機 AddIn 不應再呼叫 `/api/outlook/poll` 或 `/api/outlook/push-*`。

## SignalR Endpoint

正式 Outlook AddIn channel：

```text
/hub/outlook-addin
```

Web UI notification channel：

```text
/hub/notifications
```

## Web UI / AI Request Endpoint

這些 endpoint 仍是 Web UI、AI 或 MCP client 對 Hub 的入口。Hub 收到 request 後會透過 `/hub/outlook-addin` dispatch `OutlookCommand`。

- `POST /api/outlook/request-folders`：dispatch folder fetch command。
- `POST /api/outlook/request-mails`：dispatch mail fetch command。
- `POST /api/outlook/request-rules`：dispatch Outlook rule fetch command。
- `POST /api/outlook/request-categories`：dispatch Outlook master category fetch command。
- `POST /api/outlook/request-signalr-ping`：透過正式 AddIn channel dispatch `ping` 測試 command。
- `POST /api/outlook/request-calendar`：dispatch Outlook calendar fetch command。
- `POST /api/outlook/request-mark-mail-read`：dispatch 單封郵件標記已讀 command。
- `POST /api/outlook/request-mark-mail-unread`：dispatch 單封郵件標記未讀 command。
- `POST /api/outlook/request-mark-mail-task`：dispatch 單封郵件 flag/follow-up command。
- `POST /api/outlook/request-clear-mail-task`：dispatch 單封郵件清除 flag/follow-up command。
- `POST /api/outlook/request-set-mail-categories`：dispatch 單封郵件 category command。
- `POST /api/outlook/request-update-mail-properties`：dispatch 單封郵件屬性整批更新 command。
- `POST /api/outlook/request-upsert-category`：dispatch Outlook master category 新增或更新顏色 command。
- `POST /api/outlook/request-create-folder`：dispatch 建立 folder command。
- `POST /api/outlook/request-delete-folder`：dispatch 刪除 folder command。
- `POST /api/outlook/request-move-mail`：dispatch 移動單封郵件 command。
- `GET /api/outlook/folders`：讀取 cached folder snapshot，格式是 `FolderSnapshotDto`。
- `GET /api/outlook/mails`：讀取 cached mails。
- `GET /api/outlook/rules`：讀取 cached Outlook rules。
- `GET /api/outlook/categories`：讀取 cached Outlook master category list。
- `GET /api/outlook/calendar`：讀取 cached Outlook calendar events。
- `POST /api/outlook/chat`：新增並 broadcast chat message。
- `GET /api/outlook/chat`：讀取 cached chat messages。
- `GET /api/outlook/command-results/{commandId}`：查詢指定 command 的執行狀態，供 AI / MCP client 等待 AddIn 回報。
- `GET /api/outlook/command-results`：查詢最近 command 執行狀態。

如果沒有 Outlook AddIn SignalR connection，request endpoint 回傳：

```json
{
  "commandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "status": "addin_unavailable"
}
```

HTTP status 是 `409 Conflict`。

AI / MCP client 的完整建議流程請參考 `docs/ai/mcp-skill-integration.md`。

## AddIn SignalR Server Method

AddIn 連到 `/hub/outlook-addin` 後可以 invoke：

- `RegisterOutlookAddin(info)`：註冊 AddIn connection。
- `BeginFolderSync(info)`：開始 folder 增量同步並 broadcast `FolderSyncStarted`。
- `PushFolderBatch(batch)`：merge stores / folders 小批次並 broadcast `FoldersPatched`。
- `CompleteFolderSync(info)`：結束 folder 增量同步並 broadcast `FolderSyncCompleted`。
- `PushMails(mails)`：取代 cached mails 並 broadcast update。
- `PushRules(rules)`：取代 cached Outlook rules 並 broadcast update。
- `PushCategories(categories)`：取代 cached Outlook master category list 並 broadcast update。
- `PushCalendar(events)`：取代 cached Outlook calendar events 並 broadcast update。
- `SendChatMessage(message)`：AddIn 透過 SignalR 送出 chat message，Hub 會 broadcast `NewChatMessage`。
- `ReportAddinLog(entry)`：回報 AddIn log。
- `ReportCommandResult(result)`：回報 command 執行結果。

## AddIn SignalR Client Event

AddIn 需要 listen：

- `OutlookCommand`

Payload 是 `PendingCommand`，command type 與 request object 請看 `docs/addin/signalr-contract.md`。

## Admin Endpoint

- `GET /api/outlook/admin/status`
- `GET /api/outlook/admin/logs`
- `POST /api/outlook/admin/log`

## Web UI SignalR Event

Web UI notification endpoint 是 `/hub/notifications`。

目前事件：

- `FolderSyncStarted`
- `FoldersPatched`
- `FolderSyncCompleted`
- `MailsUpdated`
- `RulesUpdated`
- `CategoriesUpdated`
- `CalendarUpdated`
- `NewChatMessage`
- `AddinStatus`
- `AddinLog`
