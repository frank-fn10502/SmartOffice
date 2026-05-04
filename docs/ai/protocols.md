# Protocol Notes for AI Agents

Office 2016 Add-in 線上文件請參考 `docs/ai/office2016-addin-references.md`。工作機傳送與接收格式請參考 `docs/ai/office2016-workstation-contract.md`。工作機測試回報格式請參考 `docs/ai/office2016-test-report.md`。

## Add-in Protocol

Outlook Add-in 目前使用 polling protocol：

1. Web UI、AI 或 MCP client 透過 Hub enqueue command。
2. Outlook Add-in 呼叫 `GET /api/outlook/poll`。
3. Hub 回傳一筆 command，或 `{ "type": "none" }`。
4. Add-in 在本機執行 Office automation。
5. Add-in 將結果 push 回 Hub。
6. Hub 更新 cache 並 broadcast SignalR event。

除非要同步替換所有 caller，否則請保持這個 pattern。

注意：目前 SignalR 只用於 Hub 對 Web UI 的 real-time update。Outlook Add-in 沒有連到 `/hub/notifications`，也沒有透過 SignalR 接收 command；Add-in 與 Hub 的 command/result 溝通是 HTTP long-poll 與 push endpoint。

## Add-in Mocks

預設 runtime 不啟用 Hub 端 Add-in mock，讓工作電腦的 Outlook Add-in 可以直接透過 polling protocol 連進 Hub。離線開發或沒有 Office Add-in 可用時，使用 `http-mock` launch profile 啟用 Hub 端 Add-in mock：

```bash
dotnet run --launch-profile http-mock
```

`Mock` 環境使用下列設定：

```json
{
  "AddinMocks": {
    "Enabled": true,
    "ResponseDelayMilliseconds": 400,
    "Outlook": {
      "Enabled": true
    }
  }
}
```

這些 mock 位於 SmartOffice.Hub 內，模擬的是 Add-in connection，不是前端假資料。

- `OutlookMockAddinWorker`：目前的 Outlook mock，會從 Outlook 專用 `CommandQueue` 取出 command、產生 mock folders/mails/rules/calendar events、寫入 Hub cache，並 broadcast `FoldersUpdated`、`MailsUpdated`、`RulesUpdated`、`CalendarUpdated`、`AddinStatus` 與 `AddinLog`。

新增 Word、PTT 或其他 Add-in mock 時，請新增獨立 worker 與自己的 protocol boundary。不要抽共同 Add-in interface，因為不同 Office / tool Add-in 的 command、state、payload 與 execution model 可能不同。Web UI 必須仍然呼叫既有 Hub API，不要在前端硬塞 mock data。

## Outlook Route Prefix

```text
/api/outlook
```

## Web UI / AI Request Endpoint

- `POST /request-folders`：enqueue folder fetch command。
- `POST /request-mails`：enqueue mail fetch command。
- `POST /request-rules`：enqueue Outlook rule fetch command。
- `POST /request-categories`：enqueue Outlook master category fetch command。
- `POST /request-calendar`：enqueue Outlook calendar fetch command。
- `POST /request-mark-mail-read`：enqueue 單封郵件標記已讀 command。
- `POST /request-mark-mail-unread`：enqueue 單封郵件標記未讀 command。
- `POST /request-mark-mail-task`：enqueue 單封郵件 flag/follow-up command。
- `POST /request-clear-mail-task`：enqueue 單封郵件清除 flag/follow-up command。
- `POST /request-set-mail-categories`：enqueue 單封郵件 category command。
- `POST /request-update-mail-properties`：enqueue 單封郵件屬性整批更新 command。
- `POST /request-upsert-category`：enqueue Outlook master category 新增或更新顏色 command。
- `POST /request-create-folder`：enqueue 建立 folder command。
- `POST /request-delete-folder`：enqueue 刪除 folder command。
- `POST /request-move-mail`：enqueue 移動單封郵件 command。
- `GET /folders`：讀取 cached folders。
- `GET /mails`：讀取 cached mails。
- `GET /rules`：讀取 cached Outlook rules。
- `GET /categories`：讀取 cached Outlook master category list。
- `GET /calendar`：讀取 cached Outlook calendar events。
- `POST /chat`：新增並 broadcast chat message。
- `GET /chat`：讀取 cached chat messages。

## Add-in Endpoint

- `GET /poll`：long-poll 取得一筆 pending command。
- `POST /push-folders`：取代 cached folders 並 broadcast update。
- `POST /push-mails`：取代 cached mails 並 broadcast update。
- `POST /push-rules`：取代 cached Outlook rules 並 broadcast update。
- `POST /push-categories`：取代 cached Outlook master category list 並 broadcast update。
- `POST /push-calendar`：取代 cached Outlook calendar events 並 broadcast update。

## Admin Endpoint

- `GET /admin/status`
- `GET /admin/logs`
- `POST /admin/log`

## SignalR

SignalR endpoint：

```text
/hub/notifications
```

目前事件：

- `FoldersUpdated`
- `MailsUpdated`
- `RulesUpdated`
- `CategoriesUpdated`
- `CalendarUpdated`
- `NewChatMessage`
- `AddinStatus`
- `AddinLog`
