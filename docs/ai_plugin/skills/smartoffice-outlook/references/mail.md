# Mail API Reference

讀這份文件處理 mail list、mail search、body、conversation、attachments 與 attachment export。共通 request/fetch-result envelope 見 `http-api.md`；folder path 規則見 `folders.md`。

## 目錄

- `Mail List`: 最近郵件與指定 folder mails。
- `Mail Search`: subject、sender、日期、category、附件、已讀與旗標搜尋。
- `Body / Conversation / Attachments`: 讀 body、討論串、附件與匯出附件。
- `MailItemDto`: mail metadata 與 recipient 欄位。

## Mail List

### `POST /api/outlook/request-mails`

Request:

```json
{
  "folderPath": "/主要信箱 - User/收件匣",
  "lookbackHours": 168,
  "maxCount": 30
}
```

`folderPath` 必須取自 folder result 的 `data.folders[].folderPath`。`lookbackHours` 是以小時為單位的簡易相對時間，例如 `12` 代表過去 12 小時、`24` 代表過去 1 天、`168` 代表過去 7 天。也可直接傳入 `receivedFrom` / `receivedTo` date-time 邊界。完成後用 `POST /api/outlook/fetch-result-mails` 讀 `data.mails`。mail list 只應包含 metadata，完整 body 需另請求。

### `POST /api/outlook/request-folder-mails`

列出指定 folder 範圍內的所有 mail metadata。這是 client 要做批次操作時的簡單入口；不要用近期 mail list API 來枚舉整個 folder。

Request:

```json
{
  "folderPath": "/主要信箱 - User/Projects/folderA",
  "includeSubFolders": true,
  "receivedFrom": null,
  "receivedTo": null,
  "maxCount": 500
}
```

完成後讀 `POST /api/outlook/fetch-result-folder-mails`。

`folderPath` 必須取自 folder result 的 `data.folders[].folderPath`。AI agent 預設應使用 `includeSubFolders=true`，讓指定 folder 底下的子資料夾也納入範圍；只有使用者明確排除 subfolders 時才設為 `false`。`maxCount` 會套用到每個 folder slice，SmartOffice 會 clamp 到 1-500；需要搬空大型 folder tree 時應分批規劃，不要一次要求無上限。這是直接列出 folder mails 的 API，不是文字搜尋；不要改用 `request-mail-search` 取代。

`fetch-result-folder-mails.data` 包含：

- `folderMailsId`: 本次 folder mails request 的 id。
- `mails`: 指定 folder 範圍內的 mail metadata；每筆 `id` 是後續 move / delete / update 使用的 mail id。

## Mail Search

### `POST /api/outlook/request-mail-search`

`request-mail-search` 是 mail 操作的主要定位與篩選 API。Caller 可用它先取得符合條件的 `data.mails[].id` 與 `folderPath`，再進行 body 讀取、附件讀取、category 更新、move 或 delete。它不只用於文字搜尋，也用於日期、category、附件、已讀與旗標條件。

Request:

```json
{
  "searchId": "optional-client-search-id",
  "storeId": "",
  "scopeFolderPaths": ["/主要信箱 - User/收件匣"],
  "allowGlobalScope": false,
  "includeSubFolders": true,
  "keyword": "customer",
  "textFields": ["subject"],
  "categoryNames": [],
  "hasAttachments": null,
  "flagState": "any",
  "readState": "any",
  "receivedFrom": null,
  "receivedTo": null
}
```

- `scopeFolderPaths` 空陣列且 `storeId` 有值時，代表指定 store 內目前已載入的可搜尋 mail folders。
- `scopeFolderPaths` 空陣列且 `storeId` 也空白時，必須同時設定 `allowGlobalScope=true`，代表全部目前已載入的可搜尋 mail folders；AI agent 不可在使用者未要求全域搜尋時設定此值。
- 使用者未指定 folder 時，使用主要 mailbox 的 Inbox，並設為 `scopeFolderPaths` 第一個值；此值必須取自 folder result 的 `data.folders[].folderPath`，不能硬寫英文 `Inbox`。預設 `includeSubFolders=true`，只有使用者明確排除 subfolders 時才設為 `false`。
- 若 `state=failed` 且 message 或 progress 顯示 `no_searchable_folder`，代表指定 scope 目前無法搜尋；先重新讀 folders 並改用回傳的實際 `folderPath`。
- `textFields`: `subject`、`sender`、`body`；API 會 normalize，不合法時回到 `subject`。
- `textFields` 包含 `body` 且 `keyword` 有值時，會使用 Outlook 內容搜尋。其他 subject、sender、category、attachment、flag、read state 與 received time 條件屬於 metadata filter。
- `categoryNames`: Outlook category names；多個值代表任一 category 符合即可。只用 category 搜尋時，讓 `keyword` 保持空字串。
- `flagState`: `any`、`flagged`、`unflagged`。
- `readState`: `any`、`unread`、`read`。
- `hasAttachments`: true / false / null。若只想用「存在附件」搜尋，設定 `hasAttachments=true` 並讓 `keyword` 保持空字串。

只有使用者明確要求全信箱或全部已載入 mail folders 時，才可使用：

```json
{
  "scopeFolderPaths": [],
  "allowGlobalScope": true
}
```

使用空 scope 時，回覆使用者必須說明範圍是「目前 SmartOffice API 已知的可搜尋 mail folders」，不是保證完整 Outlook mailbox。

常用 request 型態：

```json
{
  "scopeFolderPaths": ["/主要信箱 - User/folderA"],
  "includeSubFolders": true,
  "keyword": "",
  "textFields": ["subject"],
  "categoryNames": ["待處理"],
  "hasAttachments": null,
  "flagState": "any",
  "readState": "any",
  "receivedFrom": null,
  "receivedTo": null
}
```

上述代表在 folderA 與其 subfolders 中搜尋 category 為 `待處理` 的 mails，可接續用 `request-move-mails` 做批次搬移。

Progress:

- `GET /api/outlook/mail-search/progress/{searchId}`

Result:

- `POST /api/outlook/fetch-result-mail-search`

## Body / Conversation / Attachments

### `POST /api/outlook/request-mail-body`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/主要信箱 - User/收件匣"
}
```

完成後用 `POST /api/outlook/fetch-result-mail-body` 讀 `data.mails`，找同一封 mail 的 `body` / `bodyHtml`。

### `POST /api/outlook/request-mail-conversation`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/主要信箱 - User/收件匣",
  "maxCount": 100,
  "includeBody": true
}
```

完成後用 `POST /api/outlook/fetch-result-mail-conversation` 讀 `data.conversationTopic` 與 `data.mails`。只有使用者需要一次性查看討論串時才讀 conversation；若包含 body，摘要必要內容即可。

### `POST /api/outlook/request-mail-attachments`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/主要信箱 - User/收件匣"
}
```

完成後用 `POST /api/outlook/fetch-result-mail-attachments` 讀 `data.attachments`。

### `POST /api/outlook/request-export-mail-attachment`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/主要信箱 - User/收件匣",
  "attachmentId": "1",
  "index": 1,
  "name": "sample.pdf",
  "fileName": "sample.pdf",
  "displayName": "sample.pdf",
  "exportRootPath": ""
}
```

`exportRootPath` 空白時 API 會使用 attachment export settings。`fetch-result-export-mail-attachment` 回傳空 `data`；完成後請重新呼叫 `request-mail-attachments` / `fetch-result-mail-attachments` 讀同一封 mail 的 attachment metadata，從 `attachments[].exportedAttachmentId` 取得已匯出的 attachment id，再呼叫 open endpoint。

### Attachment Settings / Open

- `GET /api/outlook/attachment-export-settings`
- `POST /api/outlook/attachment-export-settings` with `{ "rootPath": "..." }`
- `POST /api/outlook/open-exported-attachment` with `{ "exportedAttachmentId": "..." }`

只接受已記錄的 exported attachment id，不接受任意本機路徑。

## `MailItemDto`

`id`, `subject`, `sender`, `toRecipients`, `ccRecipients`, `bccRecipients`, `receivedTime`, `body`, `bodyHtml`, `folderPath`, `categories`, `isRead`, `isMarkedAsTask`, `attachmentCount`, `attachmentNames`, `flagRequest`, `flagInterval`, `taskStartDate`, `taskDueDate`, `taskCompletedDate`, `importance`, `sensitivity`。

`sender` 與 recipients 使用 `OutlookRecipientDto`：`recipientKind`, `displayName`, `smtpAddress`, `rawAddress`, `addressType`, `entryUserType`, `isGroup`, `isResolved`, `members`。client 應以 `displayName` 作為預設顯示名稱；`rawAddress` 可能是 Exchange legacyDN，不應直接當人名顯示。group 可用 `isGroup=true` 表示，若 SmartOffice 已展開成員則放在 `members`。
