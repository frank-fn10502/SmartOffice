# SmartOffice Outlook Workflows

本文件描述操作順序與判斷規則，不指定工具。Agent 可使用任何 HTTP client、MCP tool、SDK helper 或 shell，只要能送出 HTTP request 並以結構化方式解析 JSON。

## 通用規則

- Base URL 預設為 `http://localhost:2805`；只有使用者明確提供其他 Hub URL 時才改用該 URL。
- 呼叫任何 `POST /api/outlook/request-*` 後，從 response 取出 `commandId`，查 `GET /api/outlook/command-results/{commandId}` 直到不是 `pending`。
- `completed` 且 `success=true` 才進入下一步；`failed`、`folder_cache_unavailable`、`timeout` 或其他狀態要回報使用者。
- HTTP 200 只代表 Hub 已處理 request 流程；真正資料要讀對應 snapshot endpoint。
- HTTP API 的 folder path 一律使用普通斜線，例如 `/主要信箱 - User/收件匣`。
- 大型 JSON 可暫存到 skill folder 的 `tmp/<run>/`，但預設只保存 metadata，不保存完整 mail body。

## API Status

讀取：

- `GET /api/outlook/admin/status`

若 `connected=false`，回報 Outlook API 目前無法完成即時 request，並附上可用的簡短診斷。

## Folder Scope

未指定 folder 時，預設範圍是主要 mailbox 的 Inbox，但 `Inbox` 不是穩定 path。必須從 folder snapshot 找實際 `folderPath`：

1. `POST /api/outlook/request-folders`
2. 等待 command result。
3. `GET /api/outlook/folders`
4. 以 `stores[0]` 作為主要 store；使用同 `storeId` 且 `isStoreRoot=true` 的 folder 作為 root。
5. 若主要 store root 的 `childrenLoaded=false`，用 root folder 的 `storeId`、`entryId`、`folderPath` 呼叫 `POST /api/outlook/request-folder-children`。
6. 等待 command result。
7. 再讀 `GET /api/outlook/folders`。
8. 在主要 store 底下找 `folderType="Inbox"`；若 AddIn 未回報 folder type，再用 localized folder name，例如 `收件匣` 或 `Inbox`。

Folder snapshot 形狀範例：

```json
{
  "stores": [
    {
      "storeId": "store-primary",
      "displayName": "主要信箱 - User",
      "storeKind": "exchange",
      "storeFilePath": "",
      "rootFolderPath": "/主要信箱 - User"
    }
  ],
  "folders": [
    {
      "name": "主要信箱 - User",
      "entryId": "root-entry-id",
      "folderPath": "/主要信箱 - User",
      "parentEntryId": "",
      "parentFolderPath": "",
      "itemCount": 0,
      "storeId": "store-primary",
      "isStoreRoot": true,
      "folderType": "StoreRoot",
      "defaultItemType": -1,
      "isHidden": false,
      "isSystem": true,
      "hasChildren": true,
      "childrenLoaded": false,
      "discoveryState": "partial"
    }
  ]
}
```

展開主要 store root 的 request body：

```json
{
  "storeId": "store-primary",
  "parentEntryId": "root-entry-id",
  "parentFolderPath": "/主要信箱 - User",
  "maxDepth": 1,
  "maxChildren": 50
}
```

展開後，應從同一個 `storeId` 找 Inbox：

```json
{
  "name": "收件匣",
  "entryId": "inbox-entry-id",
  "folderPath": "/主要信箱 - User/收件匣",
  "parentEntryId": "root-entry-id",
  "parentFolderPath": "/主要信箱 - User",
  "storeId": "store-primary",
  "isStoreRoot": false,
  "folderType": "Inbox",
  "defaultItemType": 0,
  "hasChildren": true,
  "childrenLoaded": false,
  "discoveryState": "partial"
}
```

後續任何 HTTP API `folderPath` 都使用這個完整值：`/主要信箱 - User/收件匣`。不要自行改寫或只傳 folder name。

找不到 Inbox 時，不要改成全域搜尋；回報目前 folder snapshot 無法定位主要 Inbox。

## Recent Mails

用 snapshot 中的實際 Inbox `folderPath` 呼叫：

- `POST /api/outlook/request-mails`
- `GET /api/outlook/command-results/{commandId}`
- `GET /api/outlook/mails`

Request body 重點欄位：

```json
{
  "folderPath": "/主要信箱 - User/收件匣",
  "range": "1m",
  "maxCount": 30
}
```

回覆使用者時摘要 `subject`、`senderName`、`receivedTime`、`categories`、`flagInterval` 等 metadata，並說明「範圍：主要 mailbox 的 Inbox」或實際指定 folder。

## Mail Search

使用者未指定 folder 時，`scopeFolderPaths` 只放主要 mailbox 的實際 Inbox path。只有使用者明確要求其他 folder、子資料夾或全域搜尋時，才擴大 scope。

禁止把空陣列當作「預設 Inbox」。`scopeFolderPaths: []` 代表指定 store 或全部 store 內目前已載入的可搜尋 mail folders；這會放大搜尋範圍。

呼叫：

- `POST /api/outlook/request-mail-search`
- `GET /api/outlook/mail-search/progress/{searchId}` 或 `GET /api/outlook/command-results/{commandId}`
- `GET /api/outlook/mail-search`

Request body 重點欄位：

```json
{
  "searchId": "",
  "storeId": "",
  "scopeFolderPaths": ["/主要信箱 - User/收件匣"],
  "includeSubFolders": true,
  "keyword": "",
  "textFields": ["subject", "sender"],
  "categoryNames": [],
  "hasAttachments": null,
  "flagState": "any",
  "readState": "any",
  "receivedFrom": "2026-05-01T00:00:00+08:00",
  "receivedTo": "2026-06-01T00:00:00+08:00"
}
```

`hasAttachments=true` 代表只找有附件；`false` 代表只找無附件；`null` 代表不限。`keyword` 可為空字串，表示只套用日期、附件、已讀、旗標等條件。

若 response status 是 `no_searchable_folder`，先檢查 `scopeFolderPaths` 是否完全等於 snapshot 中的真實 `folderPath`，並在必要時展開 root children；不要用猜測路徑重試，也不要自行改成全域搜尋。

若使用者明確要求全信箱搜尋，才可以送空 `scopeFolderPaths`；回覆時必須說明本次搜尋範圍是目前 cache 中已載入的可搜尋 mail folders，而不是保證完整 Outlook mailbox。

## Read Body Or Attachments

先從 `GET /api/outlook/mails` 或 `GET /api/outlook/mail-search` 找到目標 `id` 與 `folderPath`。

使用者只要求最近郵件、郵件清單、統計或 Markdown metadata 報告時，停在 metadata；不要為每封 mail 自動讀 body。只有使用者明確要求內容摘要、內文關鍵字判讀，或 metadata 不足以完成任務時才讀 body。

讀 body：

- `POST /api/outlook/request-mail-body`
- `GET /api/outlook/command-results/{commandId}`
- `GET /api/outlook/mails`

若同一封 mail 的 body request 已完成，但 snapshot 中同 id 的 `body` 與 `bodyHtml` 仍為空，不要重複呼叫同一 request；視為 Outlook/AddIn 目前未提供可讀內容，回報限制或改用 metadata。

讀 attachment metadata：

- `POST /api/outlook/request-mail-attachments`
- `GET /api/outlook/command-results/{commandId}`
- `GET /api/outlook/mail-attachments?mailId={mailId}`

只取任務需要的內容片段；不要把完整 body、attachment path 或大量敏感資料貼回對話。

## Mutations

修改、移動或刪除 mail 前，必須使用 snapshot 中的 `mailId` 與 `folderPath`，並確認目標就是使用者指定的 mail。

若 snapshot 中有多封 mail 符合同一 subject 或 sender，不要任選一封。先向使用者列出必要 metadata，例如 `receivedTime`、`senderName`、短 subject，請使用者確認目標。

常用 endpoint：

- `POST /api/outlook/request-update-mail-properties`
- `POST /api/outlook/request-move-mail`
- `POST /api/outlook/request-move-mails`
- `POST /api/outlook/request-delete-mail`

`request-delete-mail` 代表移到 Deleted Items，不是永久刪除。mutation 完成後重新讀 `mails` 與必要的 `folders` snapshot，確認結果。

### Bulk Move Folder Tree

當使用者要求「將 folderA 與底下 subfolder 的郵件都搬到 folderOther」或「搬空某個 folder tree」時，使用這個流程。不要用 `request-mails` 蒐集目標郵件，因為它是近期列表 API，受 `range` / `maxCount` 限制。

1. 取得最新 folders snapshot，定位來源 folderA 與 destination folderOther 的真實 `folderPath`；不要自行組 path。
2. 若 folderA 的子 folder 尚未載入，先用 `request-folder-children` 展開到足以涵蓋使用者要求的層級，再重新讀 folders snapshot。
3. 用 `request-mail-search` 枚舉來源範圍：

```json
{
  "searchId": "",
  "storeId": "",
  "scopeFolderPaths": ["/Mailbox - User/folderA"],
  "includeSubFolders": true,
  "keyword": "",
  "textFields": ["subject"],
  "categoryNames": [],
  "hasAttachments": null,
  "flagState": "any",
  "readState": "any",
  "receivedFrom": null,
  "receivedTo": null
}
```

4. 等待 command result 完成後讀 `GET /api/outlook/mail-search`，只取 `id` 與 `folderPath` 等必要 metadata。
5. 將結果依最多 500 封切 batch。`POST /api/outlook/request-move-mails` 單次超過 500 會回 `400 too_many_mail_ids`。
6. 逐批呼叫 `request-move-mails`，每批都等待 `command-results/{commandId}` 成功後再送下一批：

```json
{
  "mailIds": ["id-1", "id-2"],
  "sourceFolderPath": "",
  "sourceFolderPaths": ["/Mailbox - User/folderA", "/Mailbox - User/folderA/Subfolder"],
  "destinationFolderPath": "/Mailbox - User/folderOther",
  "continueOnError": true
}
```

`sourceFolderPath` 只有單一來源 folder 時才填；跨 subfolders 時可留空並填 `sourceFolderPaths` 去幫助 folder count 更新。

7. 回報進度，例如 `500/8000`、`1000/8000`。若某批失敗，停止後回報失敗批次與已完成數量；不要假裝整批成功。
8. 全部批次完成後，重新讀必要的 folders snapshot 確認來源與目的 folder count。若需要確認目前 UI folder list，再針對相關 folder request mails。

## Calendar

呼叫：

- `POST /api/outlook/request-calendar`
- `GET /api/outlook/command-results/{commandId}`
- `GET /api/outlook/calendar`

Request body 可使用 `daysForward`，或指定 `startDate` / `endDate`。

## API Contract Reflection Checklist

若 workflow 讓 agent 必須猜測或重複試錯，回報使用者：

- request response 是否足以指出下一個 cache endpoint。
- mutation endpoint 是否要求穩定 id，而不是只靠顯示名稱。
- folder scope 是否能從 snapshot 穩定取得。
- sensitive fields 是否有必要回傳。
- mail search 的進度、結果與 command result 是否能清楚串接。
