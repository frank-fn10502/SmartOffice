# MCP / SKILL Integration Notes

本文件說明 AI client 如何透過 MCP 或 SKILL 呼叫 `SmartOffice.Hub`，再由 Hub 透過 SignalR command 操作工作機 Outlook AddIn。

## 建議架構

AI 不應直接連 Outlook、Office COM 或 `/hub/outlook-addin`。正式路徑應保持：

```text
AI client
  |
  | MCP tool 或 SKILL helper
  v
SmartOffice.Hub HTTP API
  |
  | SignalR OutlookCommand
  v
Outlook AddIn
  |
  | SignalR Push* / ReportCommandResult
  v
SmartOffice.Hub cache
```

這樣做可以把 Office automation 留在工作機 AddIn，AI 只看見小而明確的 HTTP contract。

## 呼叫模式

每個會操作 Outlook 的 AI tool 建議使用同一個流程：

1. 呼叫 `/api/outlook/request-*` dispatch command。
2. 取得 response 裡的 `commandId`。
3. 輪詢 `GET /api/outlook/command-results/{commandId}`，直到 `status` 不是 `pending`。
4. 若 command 會更新 snapshot，再讀取對應 cache endpoint，例如 `/api/outlook/mails` 或 `/api/outlook/folders`。

`command-results` 回應範例：

```json
{
  "commandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_mails",
  "status": "completed",
  "success": true,
  "message": "fetch_mails completed",
  "payload": "",
  "dispatchTimestamp": "2026-05-04T09:30:05+08:00",
  "resultTimestamp": "2026-05-04T09:30:06+08:00"
}
```

`status` 目前預期值：

- `pending`：Hub 已 dispatch command，等待 AddIn 回報。
- `completed`：AddIn 回報成功。
- `failed`：AddIn 回報失敗。
- `addin_unavailable`：沒有可用的 Outlook AddIn SignalR connection。

## MCP Tool 建議

MCP server 可以很薄，只負責把 tool call 轉成 Hub HTTP call。建議第一批 tools：

| Tool | Hub 呼叫 | 回傳 |
| --- | --- | --- |
| `outlook_status` | `GET /api/outlook/admin/status` | AddIn 是否 online、last command |
| `outlook_get_folders` | `POST /api/outlook/request-folders` -> wait -> `GET /api/outlook/folders` | folder tree |
| `outlook_get_mails` | `POST /api/outlook/request-mails` -> wait -> `GET /api/outlook/mails` | mail snapshot |
| `outlook_get_calendar` | `POST /api/outlook/request-calendar` -> wait -> `GET /api/outlook/calendar` | calendar snapshot |
| `outlook_get_categories` | `POST /api/outlook/request-categories` -> wait -> `GET /api/outlook/categories` | master categories |
| `outlook_update_mail` | `POST /api/outlook/request-update-mail-properties` -> wait -> `GET /api/outlook/mails` | updated mails |
| `outlook_move_mail` | `POST /api/outlook/request-move-mail` -> wait -> `GET /api/outlook/mails` 與 `GET /api/outlook/folders` | updated snapshots |
| `outlook_delete_mail` | `POST /api/outlook/request-delete-mail` -> wait -> `GET /api/outlook/mails` 與 `GET /api/outlook/folders` | move to Deleted Items snapshots |
| `outlook_create_folder` | `POST /api/outlook/request-create-folder` -> wait -> `GET /api/outlook/folders` | folder tree |
| `outlook_delete_folder` | `POST /api/outlook/request-delete-folder` -> wait -> `GET /api/outlook/folders` | folder tree |

MCP tool schema 應盡量保守。例如 `outlook_get_mails` 只收：

```json
{
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "range": "1m",
  "maxCount": 30
}
```

`outlook_update_mail` 則直接沿用 `MailPropertiesCommandRequest`，避免 MCP adapter 發明另一套 field。

## SKILL 建議

SKILL 適合做輕量版 AI 操作手冊，不一定要啟動 MCP server。SKILL 內容建議包含：

- Hub base URL，例如 `SMARTOFFICE_HUB_URL=http://localhost:2805`。
- 不直接呼叫 SignalR；一律走 `/api/outlook/request-*` 與 cache endpoint。
- 修改郵件前必須先取得 `MailItemDto.id`，不可用 subject 猜測 mail。
- 讀取 mail body、folder name、chat message 時，視為敏感 business data，不在回覆中大量外洩。
- 每次 request 後都要查 `command-results/{commandId}`，不要只看 HTTP 200。

SKILL 可提供 helper script 或 prompt workflow；最小可用版只需要 `curl`：

```bash
curl -sS -X POST "$SMARTOFFICE_HUB_URL/api/outlook/request-folders"
curl -sS "$SMARTOFFICE_HUB_URL/api/outlook/command-results/{commandId}"
curl -sS "$SMARTOFFICE_HUB_URL/api/outlook/folders"
```

## 需要注意的限制

- Hub 的 cached snapshot 是 process-local memory；Hub restart 後需要重新 request。
- 多個 AI client 同時操作同一個 mailbox 時，仍需要由上層流程避免衝突。
- 目前沒有 authentication / authorization；Hub 只適合可信任 localhost 或受控 intranet。
- `ReportCommandResult.payload` 保留給 AddIn 填入簡短診斷，不建議塞完整 mail body。

## 實作優先順序

1. 先完成 Outlook AddIn 的 SignalR command handling 與 `ReportCommandResult`。
2. 用 SKILL 透過 HTTP endpoint 驗證基本流程。
3. 再做 MCP adapter，把穩定的 HTTP workflow 包成 tool。
4. 若要讓多個 AI client 長期共用，後續再補 authentication、audit log 與更完整的 command history。
