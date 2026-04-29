# Protocol Notes for AI Agents

## Add-in Protocol

Outlook Add-in 目前使用 polling protocol：

1. Web UI、AI 或 MCP client 透過 Hub enqueue command。
2. Outlook Add-in 呼叫 `GET /api/outlook/poll`。
3. Hub 回傳一筆 command，或 `{ "type": "none" }`。
4. Add-in 在本機執行 Office automation。
5. Add-in 將結果 push 回 Hub。
6. Hub 更新 cache 並 broadcast SignalR event。

除非要同步替換所有 caller，否則請保持這個 pattern。

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
