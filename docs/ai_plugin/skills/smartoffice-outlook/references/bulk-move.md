# Bulk Move Workflows

讀這份文件處理大量搬移郵件、搬空 folder tree、依條件搬移 mails，以及分批呼叫 `request-move-mails`。一般 folder 定位規則見 `workflows.md` 的 `Locate Named Folders`；endpoint 欄位見 `mail.md` 與 `organizing.md`。

## 目錄

- `共通規則`: 大量搬移前的安全規則。
- `Bulk Move Folder Mails`: 搬移指定 folder 內 mails。
- `Bulk Move Folder Tree`: 搬空 folder tree。
- `Bulk Move Filtered Mails`: 依 category、日期、附件、已讀或旗標搬移。

## 共通規則

- 不要用 `request-mails` 蒐集大量目標郵件，因為它是近期列表 API，受 `lookbackHours` / `maxCount` 限制。
- 來源與目的 folder 都必須先用 `request-find-folder` 定位唯一 `folderPath`；不要自行組 path。
- `POST /api/outlook/request-move-mails` 單次最多 500 個 `mailIds`；更多郵件必須切 batch。
- 每批都要用 paired `fetch-result-move-mails` 等到 `state=completed` 後再送下一批。
- 若某批失敗，停止並回報失敗批次與已完成數量；不要假裝整批成功。
- 全部批次完成後，重新查必要的 folders、folder mails 或 search 結果確認。

## Bulk Move Folder Mails

當使用者要求「將 folderA 的郵件都搬到 folderB」時，預設包含 folderA 底下的 subfolders；只有使用者明確說「不包含子資料夾」或「只處理這個 folder 直層郵件」時才排除 subfolders。

1. 用 Locate Named Folders 定位來源 folderA 與 destination folderB 的唯一 `folderPath`。
2. 用 `request-folder-mails` 取得來源範圍的 mail metadata：

```json
{
  "folderPath": "/Mailbox - User/folderA",
  "includeSubFolders": true,
  "receivedFrom": null,
  "receivedTo": null,
  "maxCount": 500
}
```

3. 用 `fetch-result-folder-mails` loop 到 `state=completed`，只取 `data.mails[].id` 與 `data.mails[].folderPath` 等必要 metadata。
4. 若沒有 mails，回報來源 folder 沒有可搬移郵件並停止。
5. 將結果依最多 500 封切 batch。
6. 逐批呼叫 `request-move-mails`：

```json
{
  "mailIds": ["id-1", "id-2"],
  "sourceFolderPath": "/Mailbox - User/folderA",
  "sourceFolderPaths": ["/Mailbox - User/folderA"],
  "destinationFolderPath": "/Mailbox - User/folderB",
  "continueOnError": true
}
```

7. 回報進度與完成數量。

## Bulk Move Folder Tree

當使用者要求「將 folderA 與底下 subfolder 的郵件都搬到 folderOther」或「搬空某個 folder tree」時，使用這個流程。

1. 用 Locate Named Folders 定位來源 folderA 與 destination folderOther 的真實 `folderPath`。
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
4. 將結果依最多 500 封切 batch。
5. 逐批呼叫 `request-move-mails`：

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

6. 回報進度，例如 `500/8000`、`1000/8000`。
7. 全部批次完成後，重新讀必要的 folders 確認來源與目的 folder count。若需要確認目前 UI folder list，再針對相關 folder request mails。

## Bulk Move Filtered Mails

當使用者要求「將 folderA 的 `待處理` 類別郵件搬到 folder_staged」、「搬移最近兩個月有附件的 mails」或任何有 category、日期、附件、已讀、旗標等條件的批次搬移時，使用 `request-mail-search` 篩選，不要用 `request-folder-mails` 後自行猜條件。

1. 用 Locate Named Folders 定位來源 folder 與 destination folder 的真實 `folderPath`。
2. 用 `request-mail-search` 在來源 scope 內取得符合條件的 mail metadata：

```json
{
  "searchId": "",
  "storeId": "",
  "scopeFolderPaths": ["/Mailbox - User/folderA"],
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

3. 用 `fetch-result-mail-search` loop 到 `state=completed`，只取 `data.mails[].id` 與 `data.mails[].folderPath` 等必要 metadata。
4. 若沒有符合條件的 mails，回報找不到符合條件的郵件並停止。
5. 將結果依最多 500 封切 batch。若結果跨 subfolders，`sourceFolderPath` 留空，`sourceFolderPaths` 放本批 mails 實際出現過的 source folder paths：

```json
{
  "mailIds": ["id-1", "id-2"],
  "sourceFolderPath": "",
  "sourceFolderPaths": ["/Mailbox - User/folderA", "/Mailbox - User/folderA/Subfolder"],
  "destinationFolderPath": "/Mailbox - User/folder_staged",
  "continueOnError": true
}
```

6. 每批都用 `fetch-result-move-mails` 等到 `state=completed` 後再送下一批。全部批次完成後，重新查必要的 folders 或 search 結果確認。
