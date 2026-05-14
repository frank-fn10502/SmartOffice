# SmartOffice Outlook HTTP API Reference

本文件只放共通 HTTP contract 與任務文件索引。需要特定 endpoint 欄位時，依任務讀同層 reference。

## Base URL

Default: `http://localhost:2805`.

所有 endpoint 都在 `/api/outlook` 底下。

## 目錄

- `依任務讀取`: 選擇下一份 reference。
- `Request / Fetch Result Pattern`: 共通 request/fetch-result 流程。
- `Error Envelopes`: 錯誤 response 形狀。
- `Paired Fetch Result Endpoints`: request endpoint 與 fetch-result endpoint 對照。
- `Diagnostic And Lightweight Endpoints`: 診斷與輕量輔助 endpoint。

## 依任務讀取

- Folder discovery、Inbox、children、folder path：讀 `folders.md`。
- Mail list、mail search、body、conversation、attachments：讀 `mail.md`。
- Address book、calendar、rules、categories、mail/folder mutation、chat、DTO 欄位：讀 `organizing.md`。
- 多步驟操作、日期範圍、批次搬移與判斷規則：讀 `workflows.md`。
- 大量搬移、搬空 folder tree、分批 `request-move-mails`：讀 `bulk-move.md`。

## Request / Fetch Result Pattern

所有 Outlook 工作都走同一個對外模式：

1. 呼叫 `POST /api/outlook/request-*` 發起動作。
2. 從 response 取得 `requestId` 與 `data.fetchResultEndpoint`。
3. 持續呼叫該 request 配對的 `POST /api/outlook/fetch-result-*`。遇到 `failed`、`unavailable`、`timeout` 時停止並回報；遇到 `state=completed` 仍要檢查 `next.hasMore`，只有 `state=completed` 且 `next.hasMore=false` 才代表資料取完。

一般 `request-*` response：

```json
{
  "requestId": "request-id",
  "request": "request-mails",
  "state": "accepted",
  "message": "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
  "data": {
    "fetchResultEndpoint": "/api/outlook/fetch-result-mails"
  }
}
```

`request-folder-mails` 另有：

```json
{
  "requestId": "request-id",
  "request": "request-folder-mails",
  "state": "accepted",
  "message": "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
  "data": {
    "fetchResultEndpoint": "/api/outlook/fetch-result-folder-mails",
    "folderMailsId": "folder-mails-id"
  }
}
```

`request-mail-search` 另有：

```json
{
  "requestId": "request-id",
  "request": "request-mail-search",
  "state": "accepted",
  "message": "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
  "data": {
    "fetchResultEndpoint": "/api/outlook/fetch-result-mail-search",
    "searchId": "search-id"
  }
}
```

`request-*` response 的固定欄位是 `requestId`、`request`、`state`、`message`、`data`。`data.fetchResultEndpoint` 直接指出下一步要呼叫的 paired endpoint；其他 `data` 欄位是該 request 自己的 struct。

## `POST /api/outlook/fetch-result-*`

Request:

```json
{
  "requestId": "request-id",
  "cursor": "",
  "take": 100
}
```

Response:

```json
{
  "requestId": "request-id",
  "request": "request-folder-mails",
  "state": "running",
  "message": "",
  "next": {
    "cursor": "100",
    "hasMore": true
  },
  "data": {
    "folderMailsId": "folder-mails-id",
    "mails": []
  }
}
```

`fetch-result-*` response 的固定欄位是 `requestId`、`request`、`state`、`message`、`next`、`data`。

- `state`: `accepted`、`running`、`completed`、`failed`、`unavailable`、`timeout`。
- `next.cursor`: 下一段 result 的 cursor。
- `next.hasMore`: 是否還有下一段 result。
- `data`: 該 `fetch-result-*` 自己的 domain struct，例如 `mails`、`folders`、`stores`、`rules`、`categories`、`calendarEvents`。

`completed` 表示該 request 的工作已完成，不表示目前 response 已包含全部資料。`fetch-result-*` 會用 `next.hasMore` 與 `next.cursor` 做資料分頁；caller 必須持續抓到 `next.hasMore=false`。

`take` 只控制此次 `fetch-result-*` response 最多回傳多少筆 result；它不改變原本 `request-*` 的工作範圍。`take` 會 clamp 到 1-500，建議 AI/MCP 使用 100。

Skill folder 內的 `scripts/outlook-api.sh` 與 `scripts/outlook-api.ps1` 已封裝這個 loop。AI 可以直接使用：

```bash
./scripts/outlook-api.sh request-fetch /api/outlook/request-calendar '{"daysForward":31,"startDate":null,"endDate":null}'
```

```powershell
pwsh ./scripts/outlook-api.ps1 request-fetch /api/outlook/request-calendar '{ "daysForward": 31, "startDate": null, "endDate": null }'
```

## Error Envelopes

若 `request-*` 回 HTTP 409 / 400 / 502 / 504，body 通常仍包含 `state`、`message` 與可能存在的 `requestId`。Caller 應回報該狀態；不要自行擴大 folder scope、改成空 `scopeFolderPaths`，或猜測 folder path 重試。

Request body 使用錯欄位名時：

```json
{
  "request": "request-mail-search",
  "status": "invalid_request_body",
  "state": "failed",
  "message": "Request JSON does not match this endpoint schema. Remove unknown fields and use the exact property names documented in Swagger.",
  "errors": {
    "$.query": [
      "The JSON property 'query' could not be mapped to any .NET member contained in type 'SmartOffice.Hub.Contracts.SearchMailsRequest'."
    ]
  },
  "data": {}
}
```

遇到 `invalid_request_body` 時，修正欄位名稱後重送同一 endpoint；不要改成其他 endpoint，也不要擴大搜尋範圍。

缺少必要欄位時：

```json
{
  "request": "request-mails",
  "status": "missing_required_fields",
  "state": "failed",
  "message": "Missing required request field(s): folderPath.",
  "requiredFields": ["folderPath"],
  "data": {}
}
```

遇到 `missing_required_fields` 時，只補齊 `requiredFields` 列出的欄位後重送。

## Paired Fetch Result Endpoints

| Request endpoint | Fetch result endpoint | 主要 `data` 欄位 | 細節 |
| --- | --- | --- | --- |
| `POST /api/outlook/request-folders` | `POST /api/outlook/fetch-result-folders` | `stores`, `folders` | `folders.md` |
| `POST /api/outlook/request-folder-children` | `POST /api/outlook/fetch-result-folder-children` | `stores`, `folders` | `folders.md` |
| `POST /api/outlook/request-find-folder` | `POST /api/outlook/fetch-result-find-folder` | `query`, `matchCount`, `isAmbiguous`, `folders` | `folders.md` |
| `POST /api/outlook/request-mails` | `POST /api/outlook/fetch-result-mails` | `mails` | `mail.md` |
| `POST /api/outlook/request-folder-mails` | `POST /api/outlook/fetch-result-folder-mails` | `folderMailsId`, `mails` | `mail.md` |
| `POST /api/outlook/request-mail-search` | `POST /api/outlook/fetch-result-mail-search` | `searchId`, `mails` | `mail.md` |
| `POST /api/outlook/request-mail-body` | `POST /api/outlook/fetch-result-mail-body` | `mails` | `mail.md` |
| `POST /api/outlook/request-mail-attachments` | `POST /api/outlook/fetch-result-mail-attachments` | `mailId`, `folderPath`, `attachments` | `mail.md` |
| `POST /api/outlook/request-mail-conversation` | `POST /api/outlook/fetch-result-mail-conversation` | `mailId`, `folderPath`, `conversationId`, `conversationTopic`, `mails` | `mail.md` |
| `POST /api/outlook/request-export-mail-attachment` | `POST /api/outlook/fetch-result-export-mail-attachment` | `{}` | `mail.md` |
| `POST /api/outlook/request-rules` | `POST /api/outlook/fetch-result-rules` | `rules` | `organizing.md` |
| `POST /api/outlook/request-categories` | `POST /api/outlook/fetch-result-categories` | `categories` | `organizing.md` |
| `POST /api/outlook/request-calendar` | `POST /api/outlook/fetch-result-calendar` | `calendarEvents` | `organizing.md` |
| `POST /api/outlook/request-calendar-rooms` | `POST /api/outlook/fetch-result-calendar-rooms` | `rooms` | `organizing.md` |
| `POST /api/outlook/request-create-calendar-event` | `POST /api/outlook/fetch-result-create-calendar-event` | `calendarEvents` | `organizing.md` |
| `POST /api/outlook/request-update-calendar-event` | `POST /api/outlook/fetch-result-update-calendar-event` | `calendarEvents` | `organizing.md` |
| `POST /api/outlook/request-delete-calendar-event` | `POST /api/outlook/fetch-result-delete-calendar-event` | `calendarEvents` | `organizing.md` |
| `POST /api/outlook/request-address-book` | `POST /api/outlook/fetch-result-address-book` | `contacts` | `organizing.md` |
| `POST /api/outlook/request-update-mail-properties` | `POST /api/outlook/fetch-result-update-mail-properties` | `mails` | `organizing.md` |
| `POST /api/outlook/request-upsert-category` | `POST /api/outlook/fetch-result-upsert-category` | `categories` | `organizing.md` |
| `POST /api/outlook/request-create-folder` | `POST /api/outlook/fetch-result-create-folder` | `stores`, `folders` | `organizing.md` |
| `POST /api/outlook/request-delete-folder` | `POST /api/outlook/fetch-result-delete-folder` | `stores`, `folders` | `organizing.md` |
| `POST /api/outlook/request-move-mail` | `POST /api/outlook/fetch-result-move-mail` | `mails` | `organizing.md` |
| `POST /api/outlook/request-move-mails` | `POST /api/outlook/fetch-result-move-mails` | `mails` | `organizing.md` |
| `POST /api/outlook/request-delete-mail` | `POST /api/outlook/fetch-result-delete-mail` | `mails` | `organizing.md` |

## Diagnostic And Lightweight Endpoints

一般 client 優先使用 `request-*` / `fetch-result-*`。以下 endpoint 只做診斷或輕量輔助：

- `GET /api/outlook/admin/status`
- `GET /api/outlook/admin/logs`
- `GET /api/outlook/command-results/{commandId}`
- `GET /api/outlook/command-results`
- `GET /api/outlook/address-book/lookup?email={email}`
- `GET /api/outlook/chat`

服務重啟後，重新送出相關 `request-*` 取得資料。
