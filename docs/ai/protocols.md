# Protocol Notes for AI Agents

Office 2016 Outlook AddIn 功能 checklist 請參考 `../SmartOffice/docs/outlook-addin/features-checklist.md`。工作機傳送與接收格式請參考 `../SmartOffice/docs/outlook-addin/signalr-contract.md`。Office 2016 Outlook AddIn 線上文件請參考 `../SmartOffice/docs/outlook-addin/outlook-references.md`。工作機測試回報格式請參考 `../SmartOffice/docs/outlook-addin/test-report.md`。

## AddIn Protocol

Outlook AddIn 正式 protocol 已改為 SignalR-only：

1. Outlook AddIn 連線到 `/hub/outlook-addin`。
2. AddIn invoke `RegisterOutlookAddin(info)` 完成註冊。
3. Web UI、AI 或 MCP client 透過 Hub HTTP request endpoint 發出要求。
4. Hub 透過 SignalR client event `OutlookCommand` 即時 dispatch command 給 AddIn。
5. AddIn 在本機執行 Outlook automation。
6. AddIn 透過 SignalR server method `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync`、`PushMails`、`PushMail`、`PushMailBody`、`PushMailAttachments`、`PushExportedMailAttachment`、`PushRules`、`PushCategories`、`PushCalendar`、`SendChatMessage`、`ReportAddinLog` 或 `ReportCommandResult` 回報結果；mail search progress 由 Hub 自行推算。
7. Hub 更新內部狀態，並透過 `/hub/notifications` broadcast diagnostic / progress notification。Web UI 的主要資料驗證路徑必須仍是 HTTP：送出 request endpoint，再用 paired `POST /api/outlook/fetch-result-*` 查狀態與分頁資料。

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

Web UI 的定位是 Hub HTTP API 的手動測試與檢視工具。除 AddIn status、log 與非資料面的 progress/diagnostic 顯示外，不應依賴 `/hub/notifications` 直接修改 folders、mails、folder mail results、mail search results、rules、categories、calendar、chat 或 attachment data；主要資料應由 `request-*` 發起，再由 paired `fetch-result-*` 更新。

## Web UI / AI Request Endpoint

這些 endpoint 是 Web UI、AI 或 MCP client 對 Hub 發起 Outlook 工作的入口。HTTP caller 不需要理解 Hub 內部 command；每個 `request-*` 都用固定 envelope 回傳 `requestId`、`request`、`state`、`message`、`data`，後續只查 paired `POST /api/outlook/fetch-result-*`。`data` 是各 request 自己的 struct，例如 mail search 的 `searchId`。

API 本身必須可被 Web UI 以外的 caller 直接理解。新增或修改 endpoint 時，請用 Swagger、`curl` 或 `.http` 檔檢查 raw JSON：caller 應能只靠 response 判斷目前狀態、下一步、資料欄位語意、錯誤原因與是否有下一頁。如果 raw API 很難理解，就要修正 API contract，而不是只在 Web UI 裡補 UI 文案或轉換邏輯。

HTTP API 對外的 folder path 使用 `/主要信箱 - User/收件匣`；Hub 在 dispatch `OutlookCommand` 前會轉成 AddIn SignalR contract 使用的 Outlook folder path，例如 `\\主要信箱 - User\收件匣`。反向讀取 `GET /api/outlook/folders`、`mails`、`folder-mails`、`mail-search` 與 search progress 時，Hub 也會把 Outlook path 轉回 HTTP API path。這是 Hub API 邊界的實作細節，不應出現在外部 SKILL 或 Swagger 說明中。

- `POST /api/outlook/request-folders`：建立載入 stores 與 root folders 的 operation。
- `POST /api/outlook/request-folder-children`：建立載入單一 parent folder children 的 operation。
- `POST /api/outlook/request-mails`：建立 mail fetch operation。
- `POST /api/outlook/request-folder-mails`：列出指定 folder 範圍內的所有 mail metadata；Hub 負責規劃 folder scope，底層可重用 mail search slice。
- `POST /api/outlook/request-mail-search`：建立 mail search operation；Hub 必須先確保 folder data 可用、展開 store/folder scope、切成單 folder slices，並節流送給 AddIn。搜尋由 Outlook 內建搜尋執行，條件包含文字搜尋與篩選條件。
- `POST /api/outlook/request-mail-conversation`：建立單封郵件所屬 Outlook conversation 載入 operation；AddIn 應回推同一討論串的 mail metadata，`includeBody=true` 時可一併包含 body/bodyHtml，方便 Web UI 一次性檢視討論串。
- `POST /api/outlook/request-rules`：建立 Outlook rule fetch operation。
- `POST /api/outlook/request-manage-rule`：建立 Outlook rule mutation operation；只支援 Microsoft Rules object model 可建立的條件與動作，特殊 rule 只允許啟用/停用或刪除。
- `POST /api/outlook/request-categories`：建立 Outlook master category fetch operation。
- `POST /api/outlook/request-signalr-ping`：透過正式 AddIn channel 建立 `ping` 測試 operation。
- `POST /api/outlook/request-calendar`：建立 Outlook calendar fetch operation。
- `POST /api/outlook/request-update-mail-properties`：建立單封郵件屬性整批更新 operation。
- `POST /api/outlook/request-upsert-category`：建立 Outlook master category 新增或更新顏色 operation。
- `POST /api/outlook/request-create-folder`：建立 folder creation operation。
- `POST /api/outlook/request-delete-folder`：建立 `delete_folder` operation；AddIn 必須實作為將 folder 移到 Outlook default Deleted Items folder，不可永久刪除，也不可用顯示名稱或本地化名稱猜目的 folder。
- `POST /api/outlook/request-move-mail`：建立移動單封郵件 operation。
- `POST /api/outlook/request-move-mails`：建立移動多封郵件 operation；`mailIds` 必須來自 Hub data endpoint，單次最多 500 封，AddIn 逐封移動並回報結果。
- `POST /api/outlook/request-delete-mail`：建立 `delete_mail` operation；AddIn 必須實作為移到 Outlook default Deleted Items folder，不可永久刪除，也不可用顯示名稱或本地化名稱猜目的 folder。
- 若 `request-delete-folder` 的目標已經位於 Outlook default Deleted Items folder 或其子層，HTTP API 應回 `manual_delete_required` 與說明文字；使用者必須自行到 Outlook 永久刪除。`request-delete-mail` 仍只代表移到 default Deleted Items folder，不要求 Hub 以 folder 顯示名稱或本地化名稱阻擋。
- `GET /api/outlook/folders`：讀取 folder data，格式是 `FolderSnapshotDto`。
- `GET /api/outlook/mails`：讀取 recent mail list data。
- `GET /api/outlook/folder-mails`：讀取上次 folder mail request 的 results。
- `GET /api/outlook/mail-search/progress/{searchId}`：查詢 mail search 進度。
- `GET /api/outlook/mail-search/progress/by-command/{commandId}`：診斷用，以內部 command id 查詢 mail search 進度。
- `GET /api/outlook/mail-conversation?mailId=...`：讀取上次載入的單封 mail conversation。
- `GET /api/outlook/rules`：讀取 Outlook rules。
- `GET /api/outlook/categories`：讀取 Outlook master category list。
- `GET /api/outlook/calendar`：讀取 Outlook calendar events。
- `POST /api/outlook/chat`：新增並 broadcast chat message。
- `GET /api/outlook/chat`：讀取 chat messages。
- `POST /api/outlook/fetch-result-*`：正式 client workflow 使用的狀態與分頁資料入口；每個 request endpoint 都有對應 fetch-result endpoint。
- `GET /api/outlook/command-results/{commandId}`：診斷用，查詢指定內部 command 的執行狀態。
- `GET /api/outlook/command-results`：診斷用，查詢最近 command 執行狀態。

如果沒有 Outlook AddIn SignalR connection，request endpoint 回傳：

```json
{
  "requestId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "request": "request-mails",
  "state": "accepted",
  "message": "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
  "data": {}
}
```

HTTP status 是 `200 OK`；實際的 `unavailable`、`timeout` 或其他失敗狀態由 paired `fetch-result-*` 回報。request body 格式錯誤或 destructive safety rule 擋下時，仍可直接回 `400` 或 `409`。

AI / MCP client 與 Agents SKILL 的完整建議流程請參考 `docs/ai_plugin/mcp-agents-skill-integration.md`。

目前郵件已讀、flag 與 category mutation 以 `request-update-mail-properties` / `update_mail_properties` 為正式入口。舊的 marker-style endpoint 已從 Hub 公開 API 移除，不應再寫入 `../SmartOffice/docs/outlook-addin/` 或要求工作機 AddIn 維護相容 handler。

## Hub Mail Search Planning

Hub 是 mail search 的負載控管者；AddIn 只處理 Hub 指定的單一 folder slice。Hub 收到 `request-mail-search` 時必須：

1. 若 folder data 尚不可用，先由 Hub dispatch `fetch_folder_roots`，取得 store/root folder data；mail search 只使用目前 Hub 已知的 folder data，不得為了搜尋自動展開整棵 folder tree。
2. 若 folder data 無法建立，讓原始 search operation 失敗並回 `folder_cache_unavailable`。
3. 使用 folder tree data 展開原始 request 的 `storeId` / `scopeFolderPaths` / `includeSubFolders`；Hub 只允許 `defaultItemType == 0`、`isHidden == false`、`isSystem == false`、非 store root，且 `folderType` 是可操作 mail enum 的 folder 進入搜尋計畫。
4. 將搜尋計畫切成單 folder slices；每個 slice 都要有非空 `storeId`、`folderEntryId` 與單一 `folderPath`。
5. 依序 dispatch slices，slice 之間保留短暫 delay，避免大量 Outlook search 同時啟動。
6. 用 `MailSearchProgress` broadcast 整體進度；MCP / Agents SKILL 可用 progress endpoint 主動查詢。

搜尋條件：

- Hub 把 folder data 的 `entryId`、`folderPath`、`keyword`、`textFields` 與分類、附件、旗標、已讀狀態、時間等篩選條件傳給每個 `MailSearchSliceRequest`。
- AddIn 在單一 folder 內依 Microsoft Outlook `AdvancedSearch` / DASL 這類內建搜尋流程組合 filter，再回傳符合條件的 metadata。
- AddIn 必須在同一個 folder slice 內以多次 `PushMailSearchSliceResult` 分段回推結果，每批約 `3` 到 `5` 封 mail metadata，最後一批才標 `isSliceComplete=true`。
- `keyword` 預設只套用在 `subject`，使用者可在 Web UI 選擇 `sender` 或 `body`。這不是 typo-tolerant fuzzy search。

原始 request scope 展開規則：

| 原始 `storeId` | 原始 `scopeFolderPaths` | Hub 規劃方式 |
| --- | --- | --- |
| 非空 | 非空 | 展開指定 store 內的指定 folder；`includeSubFolders=true` 時包含子 folder。 |
| 非空 | 空陣列 | 展開該 store 底下所有可搜尋 mail folders。 |
| 空字串 | 非空 | 用 folder tree data 找出 folder 所在 store，再分成單 folder slices。 |
| 空字串 | 空陣列 | 展開所有 stores 底下所有可搜尋 mail folders。 |

## Hub / AddIn 負載邊界

任何新功能都要先問：能否用少量 Hub command 表達使用者意圖，並由 Hub 控制 pagination、slice、batch、progress 與 result cache。除非 Outlook object model 本身要求逐項操作，Hub 不應一次 dispatch 數百或數千個 `OutlookCommand` 給 AddIn。

設計原則：

- 使用者發出一個操作時，HTTP API 儘量對應一個 parent request；需要多 folder 或多 item 時，由 Hub 規劃 slice/batch 與節流。
- AddIn 只處理目前 command 的 Outlook object model 工作，不負責跨 command 排程、跨 folder 負載管理或 Web UI state 合併。
- 大量讀取應優先 metadata-only、分批回推、paired `fetch-result-*` 分頁讀取；完整 body、attachments、conversation body 等昂貴資料只在使用者打開或明確要求時載入。
- Mock backend 必須能模擬足夠資料量與邊界情境，讓 Hub 的 slicing、paging、empty/loading/error state 在沒有真 Outlook 時先被檢查。
- 若真實 Outlook API 只能逐項處理，Hub contract 也要保留批次 request 的語意，AddIn 在單一 command 內逐項處理並回報 progress/result，而不是由 Hub 對每一項各送一個 command。

## AddIn SignalR Server Method

AddIn 連到 `/hub/outlook-addin` 後可以 invoke：

- `RegisterOutlookAddin(info)`：註冊 AddIn connection。
- `BeginFolderSync(info)`：開始 folder 增量同步並 broadcast `FolderSyncStarted`。
- `PushFolderBatch(batch)`：merge stores / folders 小批次並 broadcast `FoldersPatched`。
- `CompleteFolderSync(info)`：結束 folder 增量同步並 broadcast `FolderSyncCompleted`。
- `PushMails(mails)`：取代目前 mail list 並 broadcast update；`fetch_mails` 只應回 metadata。
- `PushMail(mail)`：只更新目前 mail list 中同 id 的單封 mail 並 broadcast update；`update_mail_properties` 應使用這個方法。
- `PushMailBody(body)`：只更新目前 mail list 中同 id 的 body 並 broadcast update；`fetch_mail_body` 應使用這個方法。
- `PushMailAttachments(attachments)`：回推單封 mail 的附件 metadata；`fetch_mail_attachments` 應使用這個方法。
- `PushExportedMailAttachment(exported)`：回推已匯出附件的本機路徑與識別；`export_mail_attachment` 應使用這個方法。
- `PushRules(rules)`：取代目前 Outlook rules 並 broadcast update。
- `PushCategories(categories)`：取代目前 Outlook master category list 並 broadcast update。
- `PushCalendar(events)`：取代目前 Outlook calendar events 並 broadcast update。
- `SendChatMessage(message)`：AddIn 透過 SignalR 送出 chat message，Hub 會 broadcast `NewChatMessage`。
- `ReportAddinLog(entry)`：回報 AddIn log。
- `ReportCommandResult(result)`：回報 command 執行結果。

## AddIn SignalR Client Event

AddIn 需要 listen：

- `OutlookCommand`

Payload 是 `PendingCommand`，command type 與 request object 請看 `../SmartOffice/docs/outlook-addin/signalr-contract.md`。

## Admin Endpoint

- `GET /api/outlook/admin/status`
- `GET /api/outlook/admin/logs`
- `POST /api/outlook/admin/log`

## Notification SignalR Event

Notification endpoint 是 `/hub/notifications`。這些事件可供 diagnostics 或外部 client 使用；Web UI 不應把它們當成主要資料來源。

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
