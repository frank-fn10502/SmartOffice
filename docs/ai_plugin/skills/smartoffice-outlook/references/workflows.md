# SmartOffice Outlook Workflows

本文件描述操作順序與判斷規則，不指定工具。Agent 可使用任何 HTTP client、MCP tool、SDK helper 或 shell，只要能送出 HTTP request 並以結構化方式解析 JSON。

## 目錄

- `通用規則`: base URL、request/fetch loop、錯誤處理與暫存。
- `API Status`: 服務狀態檢查。
- `Folder Scope`: 預設 Inbox 定位與低階 fallback。
- `Locate Named Folders`: 使用者指定 folder 的定位流程。
- `Relative Date Ranges`: 相對日期轉換規則。
- `Recent Mails`: 最近郵件流程。
- `Date Range Mail Lookup`: 日期範圍郵件查找。
- `Mail Search`: 多條件搜尋策略。
- `Categories`: category 建立、更新與套用。
- `Read Body Or Attachments`: body、conversation、attachment 讀取判斷。
- `Mutations`: 修改、移動、刪除單封 mail 或 folder。
- `Bulk Move`: 大量搬移、搬空 folder tree 與分批操作請讀 `bulk-move.md`。
- `Calendar`: calendar request。
- `API Contract Reflection Checklist`: 回報 API contract 問題。

## 通用規則

- Base URL 預設為 `http://localhost:2805`；只有使用者明確提供其他 API URL 時才改用該 URL。
- 呼叫任何 `POST /api/outlook/request-*` 後，從 request response 取出 `requestId`。response 沒有 `success` 欄位，`accepted` 只代表 SmartOffice API 已收下 request，不可在這一步判斷已成功。
- 查 `POST /api/outlook/fetch-result-*`。若回 `failed`、`unavailable`、`timeout`，停止並回報；若 `next.hasMore=true`，下一次 request 帶 `next.cursor`，即使本頁已是 `state=completed` 也要繼續。只有 `state=completed` 且 `next.hasMore=false` 才代表資料取完。
- `fetch-result-* .state` 可能是 `running`、`completed`、`failed`、`unavailable`、`timeout` 等；失敗狀態要回報使用者。
- HTTP 409 / 400 / 502 / 504 的 response body 也可能有 `requestId`、`state` 與 `message`；解析後回報，不要靠猜測重試。
- HTTP API 的 folder path 一律使用普通斜線，例如 `/主要信箱 - User/收件匣`。
- 大型 JSON 可暫存到 skill folder 的 `tmp/<run>/`，但預設只保存 metadata，不保存完整 mail body。

## API Status

讀取：

- `GET /api/outlook/admin/status`

若 `connected=false`，回報 Outlook API 目前無法完成即時 request，並附上可用的簡短診斷。

可選 helper：

```bash
./scripts/outlook-api.sh status
```

## Folder Scope

未指定 folder 時，預設範圍是主要 mailbox 的 Inbox 與其 subfolders，但 `Inbox` 不是穩定 path。優先使用 `request-find-folder` 以 `folderType="Inbox"` 取得實際 `folderPath`：

1. `POST /api/outlook/request-find-folder`

```json
{
  "name": "",
  "folderPath": "",
  "folderType": "Inbox",
  "storeId": "",
  "includeHidden": false,
  "maxResults": 20
}
```

2. 用 `fetch-result-find-folder` 等到 `state=completed` 且 `next.hasMore=false`。
3. 若 `data.matchCount=1`，使用 `data.folders[0].folderPath` 作為主要 Inbox path。
4. 若 `data.matchCount=0` 或 `data.isAmbiguous=true`，停止並回報目前無法唯一定位主要 Inbox；不要改成全域搜尋。

可選 helper：

```bash
./scripts/outlook-api.sh inbox
```

低階 fallback 流程只用於診斷：

1. `POST /api/outlook/request-folders`
2. 用 `fetch-result-folders` 等到 `state=completed` 且 `next.hasMore=false`，讀取每頁 `data.stores` 與 `data.folders`。
3. 以 `data.stores[0]` 作為主要 store。
4. 以 `stores[0]` 作為主要 store；使用同 `storeId` 且 `isStoreRoot=true` 的 folder 作為 root。
5. 若主要 store root 的 `childrenLoaded=false`，呼叫 `POST /api/outlook/request-folder-children`，request 欄位必須是 `storeId=root.storeId`、`parentEntryId=root.entryId`、`parentFolderPath=root.folderPath`。
6. 用 `fetch-result-folder-children` 等到 `state=completed` 且 `next.hasMore=false`。
7. 再用 `fetch-result-folder-children` 讀 `data.folders`。
8. 在主要 store 底下找 `folderType="Inbox"`；若 SmartOffice 未回報 folder type，再用 localized folder name，例如 `收件匣` 或 `Inbox`。

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

使用者指定 folder 名稱或路徑時，優先使用 `request-find-folder` 定位正式 `folderPath`，不要自行組 path，也不要讓 caller 自己重寫每個分支的 traversal。

建議流程：

1. 呼叫 `POST /api/outlook/request-find-folder`：

```json
{
  "name": "folderA",
  "folderPath": "",
  "folderType": "",
  "storeId": "",
  "includeHidden": false,
  "maxResults": 20
}
```

2. 用 `fetch-result-find-folder` 等到 `state=completed` 且 `next.hasMore=false`。
3. 若 `data.matchCount=1`，使用 `data.folders[0].folderPath` 作為正式 `folderPath`。
4. 若 `data.isAmbiguous=true`，列出必要的 `folderPath` 與 store 顯示名稱，請使用者確認。
5. 若 `data.matchCount=0`，回報目前 folder data 無法定位該 folder；不要改用 Inbox、空 scope 或自行猜測 `/Mailbox/folderA`。

使用者提供完整路徑時，request body 改用：

```json
{
  "name": "",
  "folderPath": "/主要信箱 - User/Projects/folderA",
  "folderType": "",
  "storeId": "",
  "includeHidden": false,
  "maxResults": 20
}
```

低階 fallback 流程只用於診斷或需要精細控制載入範圍：

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

可選 helper：

```bash
./scripts/outlook-api.sh recent-mails --lookback-hours 168 --max-count 30
```

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

有附件條件時設定 `hasAttachments=true`。若使用者未指定 folder，仍只查主要 mailbox 的 Inbox 與其 subfolders，並在回覆中說明範圍。

## Mail Search

`request-mail-search` 是 mail 操作的主要定位與篩選入口，不只是文字搜尋。當使用者用 subject、sender、日期、category、附件、已讀狀態、旗標狀態或 folder scope 描述一批 mails 時，先用 search 取得候選 metadata，再進行摘要、讀 body、更新、移動或刪除。

使用者未指定 folder 時，`scopeFolderPaths` 只放主要 mailbox 的實際 Inbox path，並預設 `includeSubFolders=true`。只有使用者明確要求其他 folder 或全域搜尋時，才擴大 scope；只有使用者明確排除 subfolders 時，才把 `includeSubFolders` 設為 `false`。

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

常見組合：

- 只用 category：`keyword=""`、`categoryNames=["待處理"]`。
- 指定 folder 裡某 category：`scopeFolderPaths` 放該 folder 的真實 path、`includeSubFolders=true`、`categoryNames=["待處理"]`。
- 最近兩個月有附件：`keyword=""`、`hasAttachments=true`、`receivedFrom` 設為目前時間往前兩個月。
- 未讀且有旗標：`readState="unread"`、`flagState="flagged"`。
- subject 或 sender 文字：`keyword` 放使用者詞彙，`textFields` 依使用者語意設為 `["subject"]`、`["sender"]` 或兩者。

若 `state=failed` 且 message 或 progress 顯示 `no_searchable_folder`，先檢查 `scopeFolderPaths` 是否完全等於 folder result 中的真實 `folderPath`，並在必要時要求載入 children；不要用猜測路徑重試，也不要自行改成全域搜尋。

若使用者明確要求全信箱搜尋，才可以送空 `scopeFolderPaths`；回覆時必須說明本次搜尋範圍是目前 SmartOffice API 已知的可搜尋 mail folders，而不是保證完整 Outlook mailbox。

## Categories

讀取 master category list：

- `POST /api/outlook/request-categories`
- `POST /api/outlook/fetch-result-categories`

新增或更新 master category：

- `POST /api/outlook/request-upsert-category`
- `POST /api/outlook/fetch-result-upsert-category`

Agent 必須用 Outlook `OlCategoryColor` enum name 與 numeric value 建立或更新 category；若使用者指定的顏色不在 `organizing.md` color table 中，請使用者改選，不要猜 enum。Category name 比對大小寫不敏感；若 master category list 已有同名 category，視為更新。套用 category 到 mail 時，先定位唯一 mail，再用 `request-update-mail-properties.categories`；若需要同時建立新 category，可先 `request-upsert-category` 或在 mail update request 使用 `newCategories`。

## Read Body Or Attachments

先從 `fetch-result-* data.mails` 找到目標 `id` 與 `folderPath`。

使用者只要求最近郵件、郵件清單、統計或 Markdown metadata 報告時，停在 metadata；不要為每封 mail 自動讀 body。只有使用者明確要求內容摘要、內文關鍵字判讀，或 metadata 不足以完成任務時才讀 body。

讀 body：

- `POST /api/outlook/request-mail-body`
- `POST /api/outlook/fetch-result-*`

若同一封 mail 的 body request 已完成，但 fetch result 中同 id 的 `body` 與 `bodyHtml` 仍為空，不要重複呼叫同一 request；視為 Outlook 目前未提供可讀內容，回報限制或改用 metadata。

讀 conversation/thread：

- `POST /api/outlook/request-mail-conversation`
- `POST /api/outlook/fetch-result-mail-conversation`

```json
{
  "mailId": "mail-entry-id",
  "folderPath": "/主要信箱 - User/收件匣",
  "maxCount": 100,
  "includeBody": true
}
```

只有使用者要求「討論串」、「conversation」、「thread」或需要一次看到同一封郵件的往返脈絡時，才讀 conversation。一般郵件清單、統計、分類、移動或刪除任務不要自動讀整串 body。

讀每頁的 `data.conversationTopic` 與 `data.mails[]`。若 `next.hasMore=true`，下一次 request 帶 `next.cursor`；直到 `state=completed` 且 `next.hasMore=false` 才停止。回覆使用者時說明這是哪一個 conversation、共取得幾封 mail，並只摘要任務必要內容；不要貼出完整討論串或大量 mail body。

讀 attachment metadata：

- `POST /api/outlook/request-mail-attachments`
- `POST /api/outlook/fetch-result-mail-attachments`

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

`request-delete-mail` 代表移到 Outlook default Deleted Items folder，不是永久刪除。HTTP request 不提供目的 folder；SmartOffice 必須用 Outlook default folder identity 定位 Deleted Items，不得依賴顯示名稱或本地化名稱。完成後告知使用者 mail 已移到刪除資料夾；若使用者要永久刪除，請使用者自行到 Outlook 操作。mutation 完成後用 paired fetch result 或重新送出必要 request 確認結果。

刪除 folder 前，先定位唯一 `folderPath`，再呼叫 `request-delete-folder` 並讀 `fetch-result-delete-folder`。`request-delete-folder` 也是移到 Outlook default Deleted Items folder，不是永久刪除；但若目標 folder 已經位於 default Deleted Items folder 或其子層，SmartOffice API 會阻擋並回 `message=manual_delete_required`。此時停止流程，回報使用者需要自行到 Outlook 操作。Agent 不可用 `Deleted Items`、`刪除的郵件` 或其他本地化顯示名稱自行判斷是否可刪。

## Bulk Move

大量搬移、搬空 folder tree 與分批 `request-move-mails` 流程請讀 `bulk-move.md`。一般單封或少量 mutation 規則仍在本文件的 `Mutations`。

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
