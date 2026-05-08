# MCP / Agents SKILL Integration Notes

本文件說明 AI client 如何透過 MCP 或 Agents SKILL 呼叫 SmartOffice API，再由 SmartOffice API 操作工作機 Outlook AddIn。

## 建議架構

AI 不應直接連 Outlook、Office COM 或內部 AddIn channel。正式路徑應保持：

```text
AI client
  |
  | MCP tool 或 Agents SKILL helper
  v
SmartOffice API
  |
  | SignalR OutlookCommand
  v
Outlook AddIn
  |
  | SignalR Push* / ReportCommandResult
  v
SmartOffice API fetch-result-* 
```

這樣做可以把 Office automation 留在工作機 AddIn，AI 只看見小而明確的 HTTP contract。

## 呼叫模式

每個會操作 Outlook 的 AI tool 建議使用同一個流程：

1. 呼叫 `/api/outlook/request-*` 發起 request。
2. 取得 request response 裡的固定欄位：`requestId`、`request`、`state`、`message`、`data`。`data` 是各 request 自己的 struct；response 沒有 `success` 欄位，`accepted` 只代表 SmartOffice API 已收下 request。
3. 呼叫 paired `POST /api/outlook/fetch-result-*`，直到 `state=completed`。
4. 若 `next.hasMore=true`，下一次 request 帶 `next.cursor`；AI/MCP 建議 `take=100`，避免單次 payload 太大。

`fetch-result-*` request 範例：

```json
{
  "requestId": "request-id",
  "cursor": "",
  "take": 100
}
```

`fetch-result-*` 回應範例：

```json
{
  "requestId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "request": "request-mails",
  "state": "completed",
  "message": "fetch_mails completed",
  "next": {
    "cursor": "",
    "hasMore": false
  },
  "data": {
    "mails": []
  }
}
```

`state` 目前預期值：

- `running`：SmartOffice API 已建立 request，正在等待 Outlook AddIn 或背景 queue 完成。
- `completed`：AddIn 回報成功。
- `failed`：AddIn 回報失敗。
- `unavailable`：沒有可用的 Outlook AddIn SignalR connection。
- `timeout`：等待 Outlook AddIn 結果逾時。

## MCP Tool 建議

MCP server 可以很薄，只負責把 tool call 轉成 SmartOffice API HTTP call。建議第一批 tools：

| Tool | SmartOffice API 呼叫 | 回傳 |
| --- | --- | --- |
| `outlook_status` | `GET /api/outlook/admin/status` | AddIn 是否 online、last command |
| `outlook_get_folders` | `POST /api/outlook/request-folders` -> `POST /api/outlook/fetch-result-*` | folder tree |
| `outlook_get_mails` | `POST /api/outlook/request-mails` -> `POST /api/outlook/fetch-result-*` | recent folder mails |
| `outlook_get_folder_mails` | `POST /api/outlook/request-folder-mails` -> `POST /api/outlook/fetch-result-*` | folder scope mail metadata |
| `outlook_get_mail_attachments` | `POST /api/outlook/request-mail-attachments` -> `POST /api/outlook/fetch-result-*` | attachment metadata |
| `outlook_get_calendar` | `POST /api/outlook/request-calendar` -> `POST /api/outlook/fetch-result-*` | calendar events |
| `outlook_get_categories` | `POST /api/outlook/request-categories` -> `POST /api/outlook/fetch-result-*` | master categories |
| `outlook_update_mail` | `POST /api/outlook/request-update-mail-properties` -> `POST /api/outlook/fetch-result-*` | updated mails |
| `outlook_move_mail` | `POST /api/outlook/request-move-mail` -> `POST /api/outlook/fetch-result-*` | updated data |
| `outlook_delete_mail` | `POST /api/outlook/request-delete-mail` -> `POST /api/outlook/fetch-result-*` | move to Outlook default Deleted Items folder |
| `outlook_create_folder` | `POST /api/outlook/request-create-folder` -> `POST /api/outlook/fetch-result-*` | folder tree |
| `outlook_delete_folder` | `POST /api/outlook/request-delete-folder` -> `POST /api/outlook/fetch-result-*` | folder moved under Outlook default Deleted Items folder |

MCP tool schema 應盡量保守。例如 `outlook_get_mails` 只收：

```json
{
  "folderPath": "/Mailbox - User/Inbox",
  "lookbackHours": 168,
  "maxCount": 30
}
```

`outlook_update_mail` 則直接沿用 `MailPropertiesCommandRequest`，避免 MCP adapter 發明另一套 field。

## Agents SKILL 建議

Agents SKILL 適合做輕量版 AI 操作手冊，不一定要啟動 MCP server。Agents SKILL 內容建議包含：

- API base URL，例如 `SMARTOFFICE_API_URL=http://localhost:2805`。
- 不直接呼叫 SignalR；一律走 `/api/outlook/request-*` 與 `POST /api/outlook/fetch-result-*`。
- 修改郵件前必須先取得 `MailItemDto.id`，不可用 subject 猜測 mail。
- 讀取 mail body、folder name、chat message 時，視為敏感 business data，不在回覆中大量外洩。
- 每次 request 後都要查 `fetch-result-*`，不要只看 HTTP 200。

Agents SKILL 可提供 helper script 或 prompt workflow；最小可用版可用 `curl` 驗證 folder request：

```bash
curl -sS -X POST "$SMARTOFFICE_API_URL/api/outlook/request-folders"
curl -sS -X POST "$SMARTOFFICE_API_URL/api/outlook/fetch-result-folders" \
  -H 'Content-Type: application/json' \
  -d '{"requestId":"{requestId}","take":100}'
```

## 需要注意的限制

- SmartOffice API 重啟後需要重新送出相關 `request-*` 才會有最新資料。
- 多個 AI client 同時操作同一個 mailbox 時，仍需要由上層流程避免衝突。
- 目前沒有 authentication / authorization；SmartOffice API 只適合可信任 localhost 或受控 intranet。
- `ReportCommandResult.payload` 是 SmartOffice API / AddIn 內部診斷欄位，不建議塞完整 mail body。

## 實作優先順序

1. 先完成 Outlook AddIn 的 SignalR command handling 與 `ReportCommandResult`。
2. 用 Agents SKILL 透過 HTTP endpoint 驗證基本流程。
3. 再做 MCP adapter，把穩定的 HTTP workflow 包成 tool。
4. 若要讓多個 AI client 長期共用，後續再補 authentication、audit log 與更完整的 command history。
