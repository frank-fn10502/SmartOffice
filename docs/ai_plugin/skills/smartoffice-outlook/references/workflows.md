# SmartOffice Outlook Workflows

本文件描述操作順序與判斷規則，不指定工具。Agent 可使用任何 HTTP client、MCP tool、SDK helper 或 shell，只要能送出 HTTP request 並以結構化方式解析 JSON。

## 通用規則

- Base URL 預設為 `http://localhost:2805`；只有使用者明確提供其他 API URL 時才改用該 URL。
- 呼叫任何 `POST /api/outlook/request-*` 後，從 request response 取出 `requestId`。response 沒有 `success` 欄位，`accepted` 只代表 SmartOffice API 已收下 request，不可在這一步判斷已成功。
- 查 `POST /api/outlook/fetch-result-*`，直到 `state=completed`；若 `next.hasMore=true`，下一次 request 帶 `next.cursor`。
- `fetch-result-* .state` 可能是 `running`、`completed`、`failed`、`unavailable`、`timeout` 等；失敗狀態要回報使用者。
- HTTP 409 / 400 / 502 / 504 的 response body 也可能有 `requestId`、`state` 與 `message`；解析後回報，不要靠猜測重試。
- HTTP API 的 folder path 一律使用普通斜線，例如 `/主要信箱 - User/收件匣`。
- 大型 JSON 可暫存到 skill folder 的 `tmp/<run>/`，但預設只保存 metadata，不保存完整 mail body。

## API Status

讀取：

- `GET /api/outlook/admin/status`

若 `connected=false`，回報 Outlook API 目前無法完成即時 request，並附上可用的簡短診斷。

## Folder Scope

未指定 folder 時，預設範圍是主要 mailbox 的 Inbox，但 `Inbox` 不是穩定 path。必須從 folder data 找實際 `folderPath`：

1. `POST /api/outlook/request-folders`
2. 用 `fetch-result-folders` 等到 `state=completed`，讀取 `data.stores` 與 `data.folders`。
3. 以 `data.stores[0]` 作為主要 store。
4. 以 `stores[0]` 作為主要 store；使用同 `storeId` 且 `isStoreRoot=true` 的 folder 作為 root。
5. 若主要 store root 的 `childrenLoaded=false`，呼叫 `POST /api/outlook/request-folder-children`，request 欄位必須是 `storeId=root.storeId`、`parentEntryId=root.entryId`、`parentFolderPath=root.folderPath`。
6. 用 `fetch-result-folder-children` 等到 `state=completed`。
7. 再用 `fetch-result-folder-children` 讀 `data.folders`。
8. 在主要 store 底下找 `folderType="Inbox"`；若 AddIn 未回報 folder type，再用 localized folder name，例如 `收件匣` 或 `Inbox`。

Folder data 形狀範例：

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

不要把 folder data 欄位 `entryId` / `folderPath` 原樣當成 request 欄位送出；此 endpoint 只接受 `parentEntryId` / `parentFolderPath`。

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

找不到 Inbox 時，不要改成全域搜尋；回報目前 folder data 無法定位主要 Inbox。

## Locate Named Folders

使用者指定 folder 名稱或路徑時，先定位正式 `folderPath`，不要自行組 path。

1. 先執行 Folder Scope 的 folder 載入流程，至少取得主要 store root 與第一層 children。
2. 若使用者提供完整路徑，優先用 `folderPath` 做大小寫不敏感的完全比對。
3. 若使用者只提供顯示名稱，例如 `folderA`，在已載入的 `data.folders` 中用 `name` 做大小寫不敏感完全比對。
4. 若找不到且相關 parent folder `childrenLoaded=false`，逐層呼叫 `request-folder-children` 載入 children，再重新比對。
5. 若找到多個同名 folder，不要任選；列出必要的 `folderPath` 與 store 顯示名稱，請使用者確認。
6. 若仍找不到，回報目前 folder data 無法定位該 folder；不要改用 Inbox、空 scope 或自行猜測 `/Mailbox/folderA`。

目的 folder 也必須同樣定位成唯一 `folderPath`。移動郵件時，來源與目的 folder 必須不同；若相同，停止並回報。

## Relative Date Ranges

使用者用相對日期時，轉成明確 `receivedFrom` / `receivedTo` 後再送 request，並在回覆中說明實際範圍。日期以 SmartOffice API / 使用者所在地時區為準；目前工作環境是 Asia/Taipei。

- `今天`：今天 00:00 到目前時間。
- `昨天`：昨天 00:00 到今天 00:00。
- `這週` / `本週`：本週一 00:00 到目前時間。
- `上週`：上週一 00:00 到本週一 00:00。
- `這個月` / `本月`：本月 1 日 00:00 到目前時間。
- `最近 N 天`：目前時間往前推 N 天到目前時間。
- `最近 N 個月`：目前時間往前推 N 個月到目前時間。
- `這兩個月`：若使用者未補充，預設按「最近兩個月」處理。

若使用者要求完整自然週或完整自然月，而不是「到目前為止」，才把 `receivedTo` 設為下一個週期開始時間。回覆時使用具體日期，例如「範圍：2026-05-04 00:00 到目前時間」。

## Recent Mails

用 folder data 中的實際 Inbox `folderPath` 呼叫：

- `POST /api/outlook/request-mails`
- `POST /api/outlook/fetch-result-*`

Request body 重點欄位：

```json
{
  "folderPath": "/主要信箱 - User/收件匣",
  "lookbackHours": 168,
  "maxCount": 30
}
```

回覆使用者時摘要 `subject`、`sender.displayName`、`receivedTime`、`categories`、`flagInterval` 等 metadata，並說明「範圍：主要 mailbox 的 Inbox」或實際指定 folder。

## Date Range Mail Lookup

使用者要求「這週的郵件」、「這兩個月有附件的郵件」或任何需要完整日期範圍的查找時，使用 `request-mail-search`，不要用 `request-mails`。`request-mails` 是最近列表 API，可能受 `maxCount` 限制而漏掉符合日期的郵件。

呼叫：

- `POST /api/outlook/request-mail-search`
- `POST /api/outlook/fetch-result-mail-search`

Request body 範例：

```json
{
  "searchId": "",
  "storeId": "",
  "scopeFolderPaths": ["/主要信箱 - User/收件匣"],
  "includeSubFolders": true,
  "keyword": "",
  "textFields": ["subject"],
  "categoryNames": [],
  "hasAttachments": null,
  "flagState": "any",
  "readState": "any",
  "receivedFrom": "2026-05-04T00:00:00+08:00",
  "receivedTo": null
}
```

有附件條件時設定 `hasAttachments=true`。若使用者未指定 folder，仍只查主要 mailbox 的 Inbox，並在回覆中說明範圍。

## Mail Search

使用者未指定 folder 時，`scopeFolderPaths` 只放主要 mailbox 的實際 Inbox path。只有使用者明確要求其他 folder、子資料夾或全域搜尋時，才擴大 scope。

禁止把空陣列當作「預設 Inbox」。`scopeFolderPaths: []` 代表指定 store 或全部 store 內目前已載入的可搜尋 mail folders；這會放大搜尋範圍。

呼叫：

- `POST /api/outlook/request-mail-search`
- `POST /api/outlook/fetch-result-*`

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

若 `state=failed` 且 message 或 progress 顯示 `no_searchable_folder`，先檢查 `scopeFolderPaths` 是否完全等於 folder result 中的真實 `folderPath`，並在必要時要求載入 children；不要用猜測路徑重試，也不要自行改成全域搜尋。

若使用者明確要求全信箱搜尋，才可以送空 `scopeFolderPaths`；回覆時必須說明本次搜尋範圍是目前 SmartOffice API 已知的可搜尋 mail folders，而不是保證完整 Outlook mailbox。

## Read Body Or Attachments

先從 `fetch-result-* data.mails` 找到目標 `id` 與 `folderPath`。

使用者只要求最近郵件、郵件清單、統計或 Markdown metadata 報告時，停在 metadata；不要為每封 mail 自動讀 body。只有使用者明確要求內容摘要、內文關鍵字判讀，或 metadata 不足以完成任務時才讀 body。

讀 body：

- `POST /api/outlook/request-mail-body`
- `POST /api/outlook/fetch-result-*`

若同一封 mail 的 body request 已完成，但資料 endpoint 中同 id 的 `body` 與 `bodyHtml` 仍為空，不要重複呼叫同一 request；視為 Outlook/AddIn 目前未提供可讀內容，回報限制或改用 metadata。

讀 attachment metadata：

- `POST /api/outlook/request-mail-attachments`
- `POST /api/outlook/fetch-result-*`
- `GET /api/outlook/mail-attachments?mailId={mailId}`

只取任務需要的內容片段；不要把完整 body、attachment path 或大量敏感資料貼回對話。

## Mutations

修改、移動或刪除 mail 前，必須使用 `fetch-result-* data.mails` 中的 `id` 與 `folderPath`，並把該 `id` 作為 mutation request 的 `mailId`，確認目標就是使用者指定的 mail。

若 `fetch-result-* data.mails` 中有多封 mail 符合同一 subject 或 sender，不要任選一封。先向使用者列出必要 metadata，例如 `receivedTime`、`sender.displayName`、短 subject，請使用者確認目標。

使用者用 subject 指定刪除目標，例如「刪除 `Re:xxxx` 這封郵件」時：

1. 未指定 folder 時，先用 Golden Path 定位主要 Inbox。
2. 用 `request-mail-search` 搜尋 subject，`keyword` 放使用者提供的 subject 片段，`textFields=["subject"]`。
3. 從 `fetch-result-mail-search data.mails` 中再做精準比對；優先找 `subject` 完整等於使用者提供文字的 mail。
4. 若只有一封唯一目標，才呼叫 `request-delete-mail`。
5. 若有多封或沒有精準命中，列出必要 metadata 請使用者確認或回報找不到，不要猜。

常用 endpoint：

- `POST /api/outlook/request-update-mail-properties`
- `POST /api/outlook/request-move-mail`
- `POST /api/outlook/request-move-mails`
- `POST /api/outlook/request-delete-mail`

`request-delete-mail` 代表移到 Outlook default Deleted Items folder，不是永久刪除。HTTP request 不提供目的 folder；AddIn 必須用 Outlook default folder identity 定位 Deleted Items，不得依賴顯示名稱或本地化名稱。若目標已經在 default Deleted Items folder 內，paired fetch result 會回 `state=failed` / `message=manual_delete_required`，此時請使用者自行到 Outlook 永久刪除。mutation 完成後用 paired fetch result 或重新送出必要 request 確認結果。

### Bulk Move Folder Mails

當使用者要求「將 folderA 的郵件都搬到 folderB」時，預設只處理 folderA 直接包含的郵件，不包含 subfolders；只有使用者明確說「包含子資料夾」、「底下所有 folder」或「folder tree」時才包含 subfolders。不要用 `request-mails` 蒐集目標郵件，因為它是近期列表 API，受 `lookbackHours` / `maxCount` 限制。

1. 用 Locate Named Folders 定位來源 folderA 與 destination folderB 的唯一 `folderPath`；不要自行組 path。
2. 用 `request-folder-mails` 取得來源範圍的 mail metadata：

```json
{
  "folderPath": "/Mailbox - User/folderA",
  "includeSubFolders": false,
  "receivedFrom": null,
  "receivedTo": null
}
```

3. 用 `fetch-result-folder-mails` loop 到 `state=completed`，只取 `data.mails[].id` 與 `data.mails[].folderPath` 等必要 metadata。
4. 若沒有 mails，回報來源 folder 沒有可搬移郵件並停止。
5. 將結果依最多 500 封切 batch。`POST /api/outlook/request-move-mails` 單次超過 500 會回 `400 too_many_mail_ids`。
6. 逐批呼叫 `request-move-mails`，每批都用 `fetch-result-move-mails` 等到 `state=completed` 後再送下一批：

```json
{
  "mailIds": ["id-1", "id-2"],
  "sourceFolderPath": "/Mailbox - User/folderA",
  "sourceFolderPaths": ["/Mailbox - User/folderA"],
  "destinationFolderPath": "/Mailbox - User/folderB",
  "continueOnError": true
}
```

7. 回報進度與完成數量。若某批失敗，停止後回報失敗批次與已完成數量。
8. 全部批次完成後，重新讀必要的 folders 或 folder mails 確認結果。

### Bulk Move Folder Tree

當使用者要求「將 folderA 與底下 subfolder 的郵件都搬到 folderOther」或「搬空某個 folder tree」時，使用這個流程。不要用 `request-mails` 蒐集目標郵件，因為它是近期列表 API，受 `lookbackHours` / `maxCount` 限制。

1. 用 Locate Named Folders 定位來源 folderA 與 destination folderOther 的真實 `folderPath`；不要自行組 path。
2. 用 `request-folder-mails` 取得來源範圍的 mail metadata：

```json
{
  "folderPath": "/Mailbox - User/folderA",
  "includeSubFolders": true,
  "receivedFrom": null,
  "receivedTo": null
}
```

3. 用 `fetch-result-folder-mails` loop 到 `state=completed`，只取 `data.mails[].id` 與 `data.mails[].folderPath` 等必要 metadata。
4. 將結果依最多 500 封切 batch。`POST /api/outlook/request-move-mails` 單次超過 500 會回 `400 too_many_mail_ids`。
5. 逐批呼叫 `request-move-mails`，每批都用 `fetch-result-move-mails` 等到 `state=completed` 後再送下一批：

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

6. 回報進度，例如 `500/8000`、`1000/8000`。若某批失敗，停止後回報失敗批次與已完成數量；不要假裝整批成功。
7. 全部批次完成後，重新讀必要的 folders 確認來源與目的 folder count。若需要確認目前 UI folder list，再針對相關 folder request mails。

## Calendar

呼叫：

- `POST /api/outlook/request-calendar`
- `POST /api/outlook/fetch-result-*`

Request body 可使用 `daysForward`，或指定 `startDate` / `endDate`。

## API Contract Reflection Checklist

若 workflow 讓 agent 必須猜測或重複試錯，回報使用者：

- request response 是否足以提供 `requestId`。
- mutation endpoint 是否要求穩定 id，而不是只靠顯示名稱。
- folder scope 是否能從 folder data 穩定取得。
- sensitive fields 是否有必要回傳。
- fetch-result-* 的進度、結果與分頁是否能清楚串接。
