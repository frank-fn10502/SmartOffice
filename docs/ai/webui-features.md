# Web UI 功能與工作機 AddIn 實作對照

本文件描述目前 `webui/` 實際提供的功能、會送出的 Hub command、工作機 Outlook AddIn 應如何處理，以及相關官方文件依據。工作機 AddIn 請以本文件搭配 `docs/ai/office2016-workstation-contract.md` 實作；不要用 mock data 反推 Outlook object model 行為。

## 官方文件依據

- Outlook `MailItem.EntryID`：Microsoft 文件說明 `EntryID` 是 item 的唯一 Entry ID，但 item 需 save 或 send 後才會有值，而且跨 store 移動可能改變。工作機 AddIn push mail 時必須提供 `id`；建議使用可由 AddIn 找回該 item 的 Outlook `EntryID`，必要時搭配 store id。參考：https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.entryid
- Outlook `NameSpace.GetItemFromID`：可用 EntryID 取回 Outlook item；文件也說用 MAPI IDs 取 item 時通常需要提供 StoreID。參考：https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getitemfromid
- Outlook `MailItem.UnRead`：已讀/未讀是 `UnRead` boolean；標記已讀應設 `UnRead = false`，標記未讀應設 `UnRead = true`。參考：https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.unread
- Outlook `MailItem.Move`：移動 mail 需要目的地 `Folder` object，並回傳移動後的 item object。參考：https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.move
- Outlook `Folder` / `Folders.Add`：folder tree 來自 Outlook folder hierarchy；建立 folder 使用 `Folders.Add(Name, Type)`，未指定 Type 時預設採用父 folder 類型。參考：https://learn.microsoft.com/en-us/office/vba/api/outlook.folder 與 https://learn.microsoft.com/en-us/office/vba/api/outlook.folders.add
- Outlook `MailItem.Categories`：官方文件說 categories 是可讀寫字串，且多分類 delimiter 由 Windows regional setting 的 `sList` 決定；工作機 AddIn 轉成 Hub DTO 時目前使用逗號字串。參考：https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.categories
- Outlook `Categories.Add`：新增 master category 使用 `Categories.Add(Name, Color, ShortcutKey)`。參考：https://learn.microsoft.com/en-us/office/vba/api/outlook.categories.add
- HTML Drag and Drop：Web UI 的拖放使用標準 `draggable`、`dragstart`、`dragover`、`drop` 與 `DataTransfer.setData("text/plain", mail.id)`；這是 browser 內建 API，不依賴外部 CSS 或 CDN。參考：https://developer.mozilla.org/en-US/docs/Web/API/HTML_Drag_and_Drop_API/Drag_operations 與 https://developer.mozilla.org/en-US/docs/Web/API/DataTransfer/setData
- CSS `:hover`：folder hover / drop highlight 是本 repo `webui/src/styles.css` 的本地 CSS pseudo-class 與 class 樣式。參考：https://developer.mozilla.org/en-US/docs/Web/CSS/:hover

## 重要前提：Mail `id` 不可省略

Web UI 對單封 mail 的所有修改都依賴 `MailItemDto.id`：

- 修改已讀/未讀、flag、category 時，Web UI 送 `mailPropertiesRequest.mailId`。
- 拖曳或選單移動郵件時，Web UI 送 `moveMailRequest.mailId`。
- 如果 AddIn push 的 mail 沒有 `id`，Web UI 會顯示警告並停用修改與移動，避免送出空 `mailId` 造成工作機誤判。

工作機 AddIn 必須在 `PushMails` 的每筆 mail 填入可找回 Outlook item 的 `id`。建議做法：

1. 讀 mail 時使用 Outlook `MailItem.EntryID` 填入 `MailItemDto.id`。
2. 若用 `NameSpace.GetItemFromID` 查找 item 時需要 store，AddIn 內部可維護 `mailId -> StoreID` map，或用 folder path / store traversal 找回 item。
3. 移動 mail 後，因 `MailItem.Move` 回傳移動後的 object，且 `EntryID` 在跨 store 移動可能改變，AddIn 應重新讀取目前 folder 或回推更新後的 mail list / folder tree。

## Web UI 功能清單

### Folder tree

- 初始讀取 cached `GET /api/outlook/folders`。
- 若 cache 為空，呼叫 `POST /api/outlook/request-folders`。
- AddIn 收到 `fetch_folders` 後，讀 Outlook folder hierarchy，invoke `PushFolders(folders)`。
- 建立 folder：右鍵 folder 後輸入名稱，Web UI 呼叫 `POST /api/outlook/request-create-folder`，command type 是 `create_folder`。AddIn 建立後應 `PushFolders`。
- 刪除 folder：右鍵 folder 後刪除，Web UI 呼叫 `POST /api/outlook/request-delete-folder`，command type 是 `delete_folder`。AddIn 刪除後應 `PushFolders`，必要時也 `PushMails` 清掉目前畫面已不存在的 mail。

### Mail list

- 選 folder 後按抓取郵件，Web UI 呼叫 `POST /api/outlook/request-mails`。
- Payload 包含 `folderPath`、`range`、`maxCount`。
- AddIn 收到 `fetch_mails` 後讀該 folder 的 mail，排序建議新到舊，invoke `PushMails(mails)`。
- 每筆 `MailItemDto.id` 必填；缺 id 的 mail 只能顯示，不能修改或移動。

### Mail detail

- 點中央 mail row 會展開文字內容。
- 可切換 HTML 顯示；若 `bodyHtml` 是空字串，Web UI 會 fallback 顯示 `body`。
- HTML 內容放在 sandboxed iframe，Web UI 不會主動載入外部資源；但如果 AddIn 放入的 `bodyHtml` 本身含外部圖片或 CSS，受限企業網路可能無法載入。

### 修改郵件屬性

- Web UI 目前用單一表單送 `POST /api/outlook/request-update-mail-properties`，command type 是 `update_mail_properties`。
- Payload 包含 `mailId`、`folderPath`、`isRead`、`flagInterval`、`flagRequest`、日期欄位、`categories` 與 `newCategories`。
- AddIn 應依 `isRead` 設定 Outlook `MailItem.UnRead` 的反向值：`isRead = true` 時設 `UnRead = false`。
- AddIn 應套用 flag 與 categories，儲存 item，然後至少 invoke `ReportCommandResult`。若畫面資料會變更，應再 `PushMails`，若新增 master category，應 `PushCategories`。

### 移動郵件

- Web UI 只提供 drag/drop 入口：拖曳中央 mail row 到左側 folder。
- Drop 成功時呼叫 `POST /api/outlook/request-move-mail`，command type 是 `move_mail`。
- Payload：

```json
{
  "mailId": "[Outlook EntryID or stable id]",
  "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
  "destinationFolderPath": "\\\\Mailbox - User\\Projects"
}
```

- AddIn 應用 `mailId` 找到 mail item，解析 `destinationFolderPath` 成 Outlook `Folder` object，呼叫 `MailItem.Move(destinationFolder)`。
- 完成後應 invoke `ReportCommandResult`，並回推：
  - `PushMails`：目前 source folder 畫面應移除該 mail，或重新推送目前 folder 的最新 mail list。
  - `PushFolders`：更新 source 與 destination folder 的 item count。

拖放視覺效果由本地 `webui/src/styles.css` 的 `.folder-row.drop-target`、`.folder-row.drop-active` 與 `:hover` 控制；不需要外部 CSS。若工作機看不到拖放效果，優先檢查該 mail 是否有 `id`，因為缺 `id` 時 Web UI 不會進入可拖曳狀態，也不會送出 `move_mail`。

### Master categories

- 初始讀取 cached `GET /api/outlook/categories`。
- 按同步會呼叫 `POST /api/outlook/request-categories`，command type 是 `fetch_categories`。
- 新增或更新顏色會呼叫 `POST /api/outlook/request-upsert-category`，command type 是 `upsert_category`。
- AddIn 應更新 Outlook master category list，然後 `PushCategories`。

### Rules

- 初始讀取 cached `GET /api/outlook/rules`。
- 按同步會呼叫 `POST /api/outlook/request-rules`，command type 是 `fetch_rules`。
- AddIn 目前只需回推可顯示的 rules snapshot；Web UI 不修改 rules。

### Calendar

- 初始讀取 cached `GET /api/outlook/calendar`。
- 按同步會呼叫 `POST /api/outlook/request-calendar`，command type 是 `fetch_calendar`。
- Payload 目前是 `{ "daysForward": 14 }`。
- AddIn 回推 `PushCalendar(events)`。

### Chat

- Web UI 呼叫 `POST /api/outlook/chat` 儲存與 broadcast chat message。
- 目前 chat 不要求 AddIn automation；Mock backend 會自動回覆，真 AddIn 可依後續需求實作。

### Admin

- Web UI 讀 `GET /api/outlook/admin/status` 與 `GET /api/outlook/admin/logs`。
- `SignalR Ping` 呼叫 `POST /api/outlook/request-signalr-ping`，command type 是 `ping`。
- AddIn 收到任何 command 後，建議在執行開始與結束時用 `ReportAddinLog` 或 `ReportCommandResult` 回報，這樣 Admin 才看得到工作機目前處理到哪一步。

## Docker / 離線資源說明

- Web UI 使用 Vue、Element Plus 與 `@microsoft/signalr` npm package；production build 會 bundle 到 `wwwroot/assets`。
- `webui/src/main.ts` 匯入的是本機 npm package 的 `element-plus/dist/index.css` 與本 repo 的 `webui/src/styles.css`，不是 CDN。
- 拖放 hover / drop highlight 沒有外部 CSS 特效；如果工作機離線，只要 Docker image 內的 Web UI build 是最新的，就應該能看到本地樣式。
