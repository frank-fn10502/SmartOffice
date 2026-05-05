# Protocol Notes for AI Agents

Office 2016 AddIn 功能 checklist 請參考 `docs/addin/features-checklist.md`。工作機傳送與接收格式請參考 `docs/addin/signalr-contract.md`。Office 2016 AddIn 線上文件請參考 `docs/addin/outlook-references.md`。工作機測試回報格式請參考 `docs/addin/test-report.md`。

## AddIn Protocol

Outlook AddIn 正式 protocol 已改為 SignalR-only：

1. Outlook AddIn 連線到 `/hub/outlook-addin`。
2. AddIn invoke `RegisterOutlookAddin(info)` 完成註冊。
3. Web UI、AI 或 MCP client 透過 Hub HTTP request endpoint 發出要求。
4. Hub 透過 SignalR client event `OutlookCommand` 即時 dispatch command 給 AddIn。
5. AddIn 在本機執行 Outlook automation。
6. AddIn 透過 SignalR server method `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync`、`PushMails`、`PushMail`、`PushMailBody`、`PushMailAttachments`、`PushExportedMailAttachment`、`PushRules`、`PushCategories`、`PushCalendar`、`SendChatMessage`、`ReportAddinLog` 或 `ReportCommandResult` 回報結果；mail search progress 由 Hub 自行推算。
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
- `POST /api/outlook/request-mail-search`：dispatch mail search command；Hub 必須先確保 folder cache、展開 store/folder scope、切成單 folder slices，並節流 dispatch 給 AddIn。搜尋預設只看 subject。
- `POST /api/outlook/request-rules`：dispatch Outlook rule fetch command。
- `POST /api/outlook/request-categories`：dispatch Outlook master category fetch command。
- `POST /api/outlook/request-signalr-ping`：透過正式 AddIn channel dispatch `ping` 測試 command。
- `POST /api/outlook/request-calendar`：dispatch Outlook calendar fetch command。
- `POST /api/outlook/request-update-mail-properties`：dispatch 單封郵件屬性整批更新 command。
- `POST /api/outlook/request-upsert-category`：dispatch Outlook master category 新增或更新顏色 command。
- `POST /api/outlook/request-create-folder`：dispatch 建立 folder command。
- `POST /api/outlook/request-delete-folder`：dispatch 刪除 folder command。
- `POST /api/outlook/request-move-mail`：dispatch 移動單封郵件 command。
- `POST /api/outlook/request-delete-mail`：dispatch `delete_mail` command；AddIn 必須實作為移到 Deleted Items，不可永久刪除。
- `GET /api/outlook/folders`：讀取 cached folder snapshot，格式是 `FolderSnapshotDto`。
- `GET /api/outlook/mails`：讀取 cached mails。
- `GET /api/outlook/mail-search/progress/{searchId}`：查詢 mail search 進度。
- `GET /api/outlook/mail-search/progress/by-command/{commandId}`：以 command id 查詢 mail search 進度。
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

目前郵件已讀、flag 與 category mutation 以 `request-update-mail-properties` / `update_mail_properties` 為正式入口。舊的 marker-style endpoint 若仍存在於程式碼中，只能視為待移除的 transitional surface，不應再寫入 `docs/addin/` 或要求工作機 AddIn 維護相容 handler。

## Hub Mail Search Planning

Hub 是 mail search 的負載控管者；AddIn 只處理 Hub 指定的單一 folder slice。Hub 收到 `request-mail-search` 時必須：

1. 若 folder cache 為空，先 dispatch `fetch_folders` 並等待 cached folder tree 可用。
2. 若 folder cache 無法建立，讓原始 search command 失敗並回 `folder_cache_unavailable`。
3. 使用 cached folder tree 展開原始 request 的 `storeId` / `scopeFolderPaths` / `includeSubFolders`。
4. 將搜尋計畫切成單 folder slices；每個 slice 都要有非空 `storeId`、單一 `scopeFolderPaths[0]`、`includeSubFolders=false`、`isHubSlice=true`。
5. 依序 dispatch slices，slice 之間保留短暫 delay，避免大量 Outlook search 同時啟動。
6. 用 `MailSearchProgress` broadcast 整體進度；MCP / Skill 可用 progress endpoint 主動查詢。

原始 request scope 展開規則：

| 原始 `storeId` | 原始 `scopeFolderPaths` | Hub 規劃方式 |
| --- | --- | --- |
| 非空 | 非空 | 展開指定 store 內的指定 folder；`includeSubFolders=true` 時包含子 folder。 |
| 非空 | 空陣列 | 展開該 store 底下所有可搜尋 mail folders。 |
| 空字串 | 非空 | 用 cached folder tree 找出 folder 所在 store，再分成單 folder slices。 |
| 空字串 | 空陣列 | 展開所有 stores 底下所有可搜尋 mail folders。 |

## AddIn SignalR Server Method

AddIn 連到 `/hub/outlook-addin` 後可以 invoke：

- `RegisterOutlookAddin(info)`：註冊 AddIn connection。
- `BeginFolderSync(info)`：開始 folder 增量同步並 broadcast `FolderSyncStarted`。
- `PushFolderBatch(batch)`：merge stores / folders 小批次並 broadcast `FoldersPatched`。
- `CompleteFolderSync(info)`：結束 folder 增量同步並 broadcast `FolderSyncCompleted`。
- `PushMails(mails)`：取代 cached mails 並 broadcast update；`fetch_mails` 只應回 metadata。
- `PushMail(mail)`：只更新 cached mails 中同 id 的單封 mail 並 broadcast update；`update_mail_properties` 應使用這個方法。
- `PushMailBody(body)`：只更新 cached mails 中同 id 的 body 並 broadcast update；`fetch_mail_body` 應使用這個方法。
- `PushMailAttachments(attachments)`：回推單封 mail 的附件 metadata；`fetch_mail_attachments` 應使用這個方法。
- `PushExportedMailAttachment(exported)`：回推已匯出附件的本機路徑與識別；`export_mail_attachment` 應使用這個方法。
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
- `MailUpdated`
- `MailSearchStarted`
- `MailSearchPatched`
- `MailSearchProgress`
- `MailSearchCompleted`
- `RulesUpdated`
- `CategoriesUpdated`
- `CalendarUpdated`
- `NewChatMessage`
- `AddinStatus`
- `AddinLog`
