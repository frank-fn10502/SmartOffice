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

- `OutlookMockAddinWorker`：目前的 Outlook mock，會從 Outlook 專用 `CommandQueue` 取出 command、產生 mock folders/mails、寫入 Hub cache，並 broadcast `FoldersUpdated`、`MailsUpdated`、`AddinStatus` 與 `AddinLog`。

新增 Word、PTT 或其他 Add-in mock 時，請新增獨立 worker 與自己的 protocol boundary。不要抽共同 Add-in interface，因為不同 Office / tool Add-in 的 command、state、payload 與 execution model 可能不同。Web UI 必須仍然呼叫既有 Hub API，不要在前端硬塞 mock data。

## Outlook Route Prefix

```text
/api/outlook
```

## Web UI / AI Request Endpoint

- `POST /request-folders`：enqueue folder fetch command。
- `POST /request-mails`：enqueue mail fetch command。
- `GET /folders`：讀取 cached folders。
- `GET /mails`：讀取 cached mails。
- `POST /chat`：新增並 broadcast chat message。
- `GET /chat`：讀取 cached chat messages。

## Add-in Endpoint

- `GET /poll`：long-poll 取得一筆 pending command。
- `POST /push-folders`：取代 cached folders 並 broadcast update。
- `POST /push-mails`：取代 cached mails 並 broadcast update。

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
- `NewChatMessage`
- `AddinStatus`
- `AddinLog`
