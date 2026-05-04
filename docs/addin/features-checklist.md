# Web UI 功能實作 Checklist

本文件是工作機 AI 實作 Outlook AddIn 功能的入口。請先用本 checklist 對照缺口；只有需要 payload 細節或 Outlook object model 依據時，再查看後面的參考文件。

## 先讀這份，必要時再查

- SignalR / DTO / payload 細節：`docs/addin/signalr-contract.md`
- Office 2016 / Outlook 官方文件入口：`docs/addin/outlook-references.md`
- 工作機測試回報格式：`docs/addin/test-report.md`

## 完成定義

工作機 AddIn 視為完成 Web UI 支援時，必須同時符合：

- [ ] AddIn 可連線 `/hub/outlook-addin`，並 invoke `RegisterOutlookAddin(info)`。
- [ ] AddIn 可 listen `OutlookCommand` 並處理本文件列出的 command。
- [ ] 每個 command 完成後都有 `ReportCommandResult`，失敗時 message 可讓 Admin logs 看出原因。
- [ ] 會改變 Web UI 資料的 command 會 invoke 對應 `Push*`。
- [ ] 所有 `MailItemDto.id` 都是非空，且 AddIn 可用它找回該 Outlook item。
- [ ] Folder tree 第一層是 Outlook Store：主要 OST 與每個 PST 都是獨立頂層。
- [ ] Web UI 不需要真實敏感資料；測試資料與錯誤回報需匿名化。

## 功能總覽

| 功能 | Web UI 操作 | Command | AddIn 完成後必須回推 |
| --- | --- | --- | --- |
| AddIn 連線 | Admin / 啟動狀態 | `ping` | `ReportCommandResult` |
| Store-first folder tree | Folders refresh | `fetch_folders` | `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync` |
| 讀取郵件 | 選 folder 後抓取郵件 | `fetch_mails` | `PushMails` metadata |
| 讀取郵件內容 | 點開 mail row | `fetch_mail_body` | `PushMailBody` |
| 讀取附件清單 | 點開 mail row / AI 需要附件 | `fetch_mail_attachments` | `PushMailAttachments` |
| 匯出附件 | Web UI / AI 選定附件 | `export_mail_attachment` | `PushExportedMailAttachment` |
| 修改郵件屬性 | 右側屬性面板送出 | `update_mail_properties` | `PushMail`，必要時 `PushCategories` |
| 拖曳移動郵件 | Drag mail row 到 folder | `move_mail` | `PushMails`、folder 增量同步 |
| Master categories | Category refresh / 新增 / 改色 | `fetch_categories`、`upsert_category` | `PushCategories` |
| Rules snapshot | Rules refresh | `fetch_rules` | `PushRules` |
| 月曆 | Calendar 同步整月 | `fetch_calendar` | `PushCalendar` |
| Chat | AddIn 主動送訊息 | `SendChatMessage` SignalR server method | Hub broadcast `NewChatMessage` |
| Folder 建立 / 刪除 | Folder 右鍵選單 | `create_folder`、`delete_folder` | folder 增量同步，必要時 `PushMails` |

## 必做 Checklist

### 1. SignalR 基礎

- [ ] AddIn 啟動後連線 Hub `/hub/outlook-addin`。
- [ ] 連線成功後 invoke `RegisterOutlookAddin(info)`。
- [ ] 註冊後 Admin status 顯示 online。
- [ ] Web UI 按 `SignalR Ping` 時，AddIn 收到 `type: "ping"`。
- [ ] AddIn 對 `ping` invoke `ReportCommandResult`。

驗收：

- [ ] Admin `Last Command` 可看到 `ping`。
- [ ] Admin logs 可看到 AddIn 收到並完成 ping。

### 2. Store-first Folder Tree

Web UI 要求：第一層必須是 Outlook Store，而不是直接顯示 Inbox/Sent。

- [ ] AddIn 收到 `fetch_folders`。
- [ ] 使用 Outlook `Application.Session.Stores` 列出目前 profile 的所有 stores。
- [ ] 每個 store 使用 `Store.GetRootFolder()` 取得 root folder。
- [ ] invoke `BeginFolderSync` 開始 folder 增量同步。
- [ ] 用 `PushFolderBatch` 分批送回 `OutlookStoreDto[]` 與 flat `FolderDto[]`。
- [ ] invoke `CompleteFolderSync` 結束 folder 增量同步。
- [ ] 主要 OST 作為第一個 store root，`storeKind = "ost"`。
- [ ] 每個 PST 各自作為一個 store root，`storeKind = "pst"`。
- [ ] 每個 `OutlookStoreDto` 都填入：
  - `storeId`
  - `displayName`
  - `storeKind`
  - `storeFilePath`
  - `rootFolderPath`
- [ ] 每個 `FolderDto` 都填入：
  - `name`
  - `folderPath`
  - `parentFolderPath`
  - `itemCount`
  - `storeId`
  - `isStoreRoot`
- [ ] Store root folder 的 `parentFolderPath = ""` 且 `isStoreRoot = true`，底下 folder 都是 `false`。
- [ ] `FolderDto` 不再包含 `subFolders`，也不重複傳 store metadata。
- [ ] `.pst` / `.ost` 的真實位置填在 `storeFilePath`，讓 Web UI hover 顯示。

驗收：

- [ ] Web UI Folders 第一層可以看到主要 OST 與每個 PST。
- [ ] Hover store 或 folder 時可看到 PST/OST 真實檔案路徑。
- [ ] 展開 PST 後可看到該 PST 底下 folder。

### 3. Mail List 與 Mail Identity

- [ ] AddIn 收到 `fetch_mails`。
- [ ] 依 `mailsRequest.folderPath` 讀取該 folder 的 mail。
- [ ] 支援 `range`：`1d`、`1w`、`1m`。
- [ ] 支援 `maxCount`。
- [ ] 回推 `PushMails(mails)`，只包含 metadata。
- [ ] 每筆 mail 的 `id` 必填，建議使用 Outlook `MailItem.EntryID` 或 AddIn 可穩定找回 item 的識別。
- [ ] `folderPath` 必須對應目前 mail 所在 folder。
- [ ] `body` 與 `bodyHtml` 在 `fetch_mails` 回應中必須留空，避免一次載入大量郵件內容。
- [ ] AddIn 收到 `fetch_mail_body`。
- [ ] 依 `mailBodyRequest.mailId` 與 `folderPath` 找回單封 mail。
- [ ] 回推 `PushMailBody(body)`，包含 `mailId`、`folderPath`、`body` 與 `bodyHtml`。
- [ ] AddIn 收到 `fetch_mail_attachments`。
- [ ] 回推 `PushMailAttachments(attachments)`，只包含附件 metadata。
- [ ] AddIn 收到 `export_mail_attachment`。
- [ ] 將指定附件匯出到 Hub 約定的 attachment root，回推 `PushExportedMailAttachment(exported)`。
- [ ] AddIn 只負責 export，不負責開啟附件；Web UI Host 會開啟已匯出的檔案。

驗收：

- [ ] Web UI 中每封 mail 都可被選取。
- [ ] Web UI 點開 mail row 後才載入該封內容。
- [ ] Web UI 點開 mail row 後可看到附件清單；按 `Export` 後可按「開啟」由 Hub Host 開檔。
- [ ] 沒有出現「缺少 id」警告。
- [ ] 已讀、flag、category、drag/drop move 都送出有效 `mailId`。

### 4. 修改郵件屬性

Web UI 目前使用 `update_mail_properties` 一次送出已讀、flag、category 與新 category。

- [ ] AddIn 收到 `update_mail_properties`。
- [ ] 用 `mailPropertiesRequest.mailId` 找回 Outlook mail item。
- [ ] 套用 `isRead`：`isRead = true` 時 Outlook `UnRead = false`。
- [ ] 套用 flag：
  - `flagInterval = "none"`：清除 task/follow-up flag。
  - `today`、`tomorrow`、`this_week`、`next_week`、`no_date`：標記 task/follow-up，並設定 `FlagRequest` 與日期。
  - `custom`：使用 payload 的 `taskStartDate` / `taskDueDate`。
  - `complete`：設定完成狀態與 `taskCompletedDate`。
- [ ] 套用 mail categories：把 `categories` 寫回 Outlook mail item。
- [ ] 若 `newCategories` 不存在於 master category list，先建立或更新 master category。
- [ ] 儲存 mail item。
- [ ] invoke `ReportCommandResult`。
- [ ] invoke `PushMail` 更新畫面中的同一封 mail；不要重新抓取整個 mail list。
- [ ] 若 master category 有變更，invoke `PushCategories`。

驗收：

- [ ] Web UI 修改已讀/未讀後狀態會更新。
- [ ] Web UI 修改 flag 後，flag label / due date 會更新。
- [ ] Web UI 修改 categories 後，mail tag 會更新。
- [ ] 若新增 category，Master categories 清單也會更新。

### 5. Drag/Drop 移動郵件

Web UI 只支援 drag/drop 移動，不提供額外移動表單。

刪除郵件的唯一允許語意是移動到 Outlook 的「刪除的郵件 / Deleted Items」folder。AddIn 不得直接呼叫 `MailItem.Delete()` 或永久刪除郵件；若 Web UI 要刪除郵件，也必須送 `move_mail` 並指定刪除的郵件 folder 作為 `destinationFolderPath`。

- [ ] AddIn 收到 `move_mail`。
- [ ] 用 `moveMailRequest.mailId` 找回 mail item。
- [ ] 用 `destinationFolderPath` 找到 Outlook destination `Folder`。
- [ ] 呼叫 Outlook `MailItem.Move(destinationFolder)`。
- [ ] 若 destination 是「刪除的郵件 / Deleted Items」，仍只呼叫 `Move(destinationFolder)`，不可呼叫 `Delete()`。
- [ ] 移動後重新讀取目前 source folder 或以正確方式移除已移動 mail。
- [ ] invoke `ReportCommandResult`。
- [ ] invoke `PushMails`，讓目前 mail list 反映移動後結果。
- [ ] 用 folder 增量同步更新 source 與 destination folder item count。

驗收：

- [ ] 拖曳 mail 到左側 folder 時，Admin `Last Command` 顯示 `move_mail`。
- [ ] Web UI 中 source folder mail list 不再顯示已移動 mail。
- [ ] 目的 folder item count 增加，source folder item count 減少。
- [ ] 跨 PST / OST 移動若 EntryID 改變，AddIn 仍會回推最新 mail snapshot。

### 6. Master Categories

- [ ] AddIn 收到 `fetch_categories`。
- [ ] 從 Outlook session master category list 讀取所有 category。
- [ ] 回推 `PushCategories(categories)`。
- [ ] AddIn 收到 `upsert_category`。
- [ ] 若 category 不存在，建立 category。
- [ ] 若 category 已存在，更新 color / shortcut key。
- [ ] 回推 `PushCategories(categories)`。

驗收：

- [ ] Web UI category 清單不是空白，除非 Outlook profile 真的沒有任何 category。
- [ ] 新增 category 後清單立即出現。
- [ ] 修改 category 顏色後清單立即更新。

### 7. Calendar 月曆

- [ ] AddIn 收到 `fetch_calendar`。
- [ ] 使用 `calendarRequest.startDate` 與 `calendarRequest.endDate` 讀取整個月份。
- [ ] `startDate` 含當日，`endDate` 不含當日。
- [ ] 回推區間內所有 calendar events。
- [ ] invoke `PushCalendar(events)`。

驗收：

- [ ] Web UI Calendar 顯示月曆格狀介面。
- [ ] 同步整月後，事件出現在對應日期。
- [ ] 點選事件可看到 subject、時間、location、organizer、attendees、busy status。

### 8. Rules Snapshot

- [ ] AddIn 收到 `fetch_rules`。
- [ ] 讀取 Outlook rules snapshot。
- [ ] 回推 `PushRules(rules)`。

驗收：

- [ ] Web UI Rules 清單可顯示 rule name、enabled、order、conditions、actions、exceptions。

### 9. Chat

AddIn 送 chat 必須使用 `/hub/outlook-addin` 的 SignalR method，不要再用 HTTP `/api/outlook/chat`。

- [ ] AddIn 要送 chat message 時 invoke `SendChatMessage(message)`。
- [ ] `message.source` 建議填 `outlook`。
- [ ] `message.text` 填入要顯示的訊息。
- [ ] 不需要自行呼叫 Web UI notification hub；Hub 會 broadcast `NewChatMessage`。

驗收：

- [ ] AddIn invoke `SendChatMessage` 後，Web UI Chat 頁面立即出現該訊息。
- [ ] AddIn code 不再呼叫 HTTP `/api/outlook/chat`。

### 10. Folder 建立與刪除

- [ ] AddIn 收到 `create_folder`。
- [ ] 用 `parentFolderPath` 找到 parent folder。
- [ ] 建立 `name` 指定的新 folder。
- [ ] 用 folder 增量同步回推 folder 變更。
- [ ] AddIn 收到 `delete_folder`。
- [ ] 用 `folderPath` 找到 folder 並刪除。
- [ ] 用 folder 增量同步回推 folder 變更。
- [ ] 若目前 mail list 指向已刪除 folder，回推 `PushMails` 清掉或更新畫面。

驗收：

- [ ] Web UI 新增子 folder 後 folder tree 更新。
- [ ] Web UI 刪除 folder 後 folder tree 更新。

## 常見失敗對照

| 現象 | 優先檢查 |
| --- | --- |
| 已讀/未讀出現 missing mail id | `PushMails` 是否填 `MailItemDto.id` |
| Drag/drop 沒送出 `move_mail` | mail 是否有 `id`，以及 folder 是否是可 drop target |
| Admin 看不到 command | AddIn 是否有 SignalR connection，Web UI request 是否回 409 |
| Category 空白 | AddIn 是否處理 `fetch_categories` 並 `PushCategories` |
| Flag 修改沒效果 | AddIn 是否處理 `update_mail_properties` 的 flag 欄位並儲存 item |
| Calendar 空白 | AddIn 是否處理 `fetch_calendar` 的 `startDate/endDate` 並 `PushCalendar` |
| AddIn chat 沒出現在 Web UI | 是否 invoke `/hub/outlook-addin` 的 `SendChatMessage`，而不是 HTTP `/api/outlook/chat` |
| Folder 沒分 OST/PST | `OutlookStoreDto` 是否正確填入 `storeId`、`storeKind` 與 `rootFolderPath` |

## 需要時查看的官方文件

- Outlook Stores / PST / OST：
  - `NameSpace.Stores`: https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.stores
  - `Store.GetRootFolder`: https://learn.microsoft.com/en-us/office/vba/api/outlook.store.getrootfolder
  - `Store.FilePath`: https://learn.microsoft.com/office/vba/api/Outlook.store.filepath
  - `Store.ExchangeStoreType`: https://learn.microsoft.com/en-us/office/vba/api/outlook.store.exchangestoretype
- Mail identity / lookup：
  - `MailItem.EntryID`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.entryid
  - `NameSpace.GetItemFromID`: https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getitemfromid
- Mail read / move：
  - `MailItem.UnRead`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.unread
  - `MailItem.Move`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.move
- Folder：
  - `Folder`: https://learn.microsoft.com/en-us/office/vba/api/outlook.folder
  - `Folders.Add`: https://learn.microsoft.com/en-us/office/vba/api/outlook.folders.add
- Category：
  - `MailItem.Categories`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.categories
  - `Categories.Add`: https://learn.microsoft.com/en-us/office/vba/api/outlook.categories.add
- Flag / task：
  - `MailItem.FlagRequest`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.flagrequest
  - `MailItem.MarkAsTask`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.markastask
  - `MailItem.ClearTaskFlag`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.cleartaskflag
  - `MailItem.TaskDueDate`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.taskduedate
- Web drag/drop UI 行為：
  - HTML Drag and Drop: https://developer.mozilla.org/en-US/docs/Web/API/HTML_Drag_and_Drop_API/Drag_operations
  - `DataTransfer.setData`: https://developer.mozilla.org/en-US/docs/Web/API/DataTransfer/setData

## Contract 對照

實作時請以 `docs/addin/signalr-contract.md` 作為 JSON / DTO 欄位準則。本文件只描述「要做哪些功能」與「怎樣算做對」；payload 範例、DTO 欄位速查與 SignalR method 名稱以 contract 文件為準。
