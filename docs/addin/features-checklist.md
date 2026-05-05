# Outlook AddIn 功能實作 Checklist

本文件是工作機 AI 實作 Outlook AddIn 功能的入口。請先用本 checklist 對照缺口；只有需要 payload 細節或 Outlook object model 依據時，再查看後面的參考文件。

AddIn 的角色必須保持單純：listen `OutlookCommand`、呼叫 Outlook object model / Office automation、invoke `Push*` 與 `ReportCommandResult`。不要在 AddIn 裡實作 Hub 已負責的排程、跨 folder 負載管理、progress 推算、cache merge、Web UI fallback 或 AI/MCP 對外流程。

## 先讀這份，必要時再查

- SignalR / DTO / payload 細節：`docs/addin/signalr-contract.md`
- Office 2016 / Outlook 官方文件入口：`docs/addin/outlook-references.md`
- 工作機測試回報格式：`docs/addin/test-report.md`

## 完成定義

工作機 AddIn 視為完成時，必須同時符合：

- [ ] AddIn 可連線 `/hub/outlook-addin`，並 invoke `RegisterOutlookAddin(info)`。
- [ ] AddIn 可 listen `OutlookCommand` 並處理本文件列出的 command。
- [ ] 每個 command 完成後都有 `ReportCommandResult`，失敗時 message 可診斷且不得含敏感資料。
- [ ] 會改變 Outlook snapshot 的 command 會 invoke 對應 `Push*`。
- [ ] 所有 `MailItemDto.id` 都是非空，且 AddIn 可用它找回該 Outlook item。
- [ ] Folder tree 第一層是 Outlook Store：主要 OST 與每個 PST 都是獨立頂層。
- [ ] 測試資料與錯誤回報需匿名化。
- [ ] 不保留舊版或未使用功能：不要實作 `/api/outlook/poll`、`/api/outlook/push-*`、HTTP chat 或未列於 contract 的 legacy command。
- [ ] 效能最佳化必須有 Microsoft 官方文件或工作機實測依據；沒有依據時，選擇最薄、最直接、最容易診斷的 Outlook API 呼叫流程。

## 功能總覽

| 功能 | Command / Method | AddIn 完成後必須回推 |
| --- | --- | --- |
| AddIn 連線 | `ping` | `ReportCommandResult` |
| Store-first folder tree | `fetch_folders` | `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync` |
| 讀取郵件 | `fetch_mails` | `PushMails` metadata |
| 搜尋郵件 candidates | `fetch_mail_search_slice` | `BeginMailSearch`、`PushMailSearchSliceResult`、`CompleteMailSearchSlice` |
| 讀取郵件內容 | `fetch_mail_body` | `PushMailBody` |
| 讀取附件清單 | `fetch_mail_attachments` | `PushMailAttachments` |
| 匯出附件 | `export_mail_attachment` | `PushExportedMailAttachment` |
| 修改郵件屬性 | `update_mail_properties` | `PushMail`，必要時 `PushCategories` |
| 移動郵件 | `move_mail` | `PushMails`、folder 增量同步 |
| 刪除郵件 | `delete_mail` | `PushMails`、folder 增量同步；實作為 Move to Deleted Items |
| Master categories | `fetch_categories`、`upsert_category` | `PushCategories` |
| Rules snapshot | `fetch_rules` | `PushRules` |
| 月曆 | `fetch_calendar` | `PushCalendar` |
| Chat | `SendChatMessage` SignalR server method | `ReportCommandResult` 不適用；method invoke 成功即可 |
| Folder 建立 / 刪除 | `create_folder`、`delete_folder` | folder 增量同步，必要時 `PushMails` |

## 必做 Checklist

### 1. SignalR 基礎

- [ ] AddIn 啟動後連線 Hub `/hub/outlook-addin`。
- [ ] 連線成功後 invoke `RegisterOutlookAddin(info)`。
- [ ] AddIn 收到 `type: "ping"`。
- [ ] AddIn 對 `ping` invoke `ReportCommandResult`；只有 Outlook object model 可正常呼叫時才回 `success=true`。
- [ ] AddIn code 不再包含 HTTP poll、HTTP push 或 HTTP chat fallback。

驗收：

- [ ] Hub 可看到 AddIn 已註冊。
- [ ] Hub 可收到 `ping` 的 command result。

### 2. Store-first Folder Tree

第一層必須是 Outlook Store，而不是直接回 Inbox/Sent。

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
- [ ] `.pst` / `.ost` 的真實位置填在 `storeFilePath`；回報文件中必須匿名化路徑。

驗收：

- [ ] `OutlookStoreDto[]` 至少包含目前 profile 可見的 store。
- [ ] 每個 folder 的 `storeId` 可對回同一批 `OutlookStoreDto`。
- [ ] 展開 PST/OST 時，folder path 與 parent path 可組回正確 tree。

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
- [ ] AddIn 只負責 export，不負責開啟附件。

驗收：

- [ ] 每封 mail 都有非空 `id`。
- [ ] `fetch_mails` 不載入完整 body。
- [ ] `fetch_mail_body` 可用 `mailId` 找回並回推同一封 mail 的內容。
- [ ] `fetch_mail_attachments` 與 `export_mail_attachment` 可用同一組附件識別 round-trip。
- [ ] 沒有出現「缺少 id」警告。
- [ ] 已讀、flag、category、move/delete 都能用有效 `mailId` 執行。

### 4. Mail Search Candidates

AddIn 收到的是 Hub 已規劃好的單一 Outlook folder candidate 讀取 slice；AddIn 不負責 keyword、contains、fuzzy、regex、跨 folder 搜尋語意、全域排程或整體 progress。

- [ ] AddIn 收到 `fetch_mail_search_slice`。
- [ ] AddIn 使用 `mailSearchSliceRequest.storeId` 與 `folderPath` 定位單一 Outlook folder。
- [ ] AddIn 若收到空 `storeId` 或空 `folderPath`，必須用 `CompleteMailSearchSlice(success=false)` 結束該 slice；不得自行全域掃描。
- [ ] AddIn 只套用 Outlook 原生且低成本的限制：單一 folder、`receivedFrom` / `receivedTo`、`maxCount`。
- [ ] `includeBody=false` 時只回 metadata，`body` / `bodyHtml` 留空。
- [ ] `includeBody=true` 時才讀取 body，供 Hub 做 body keyword 後篩。
- [ ] `maxCount` 必須有上限，建議 AddIn 端 clamp 到 200 以內。
- [ ] 使用 `BeginMailSearch`、`PushMailSearchSliceResult`、`CompleteMailSearchSlice` 回傳 candidates；不要用 `PushMails` 覆蓋目前 folder list。
- [ ] `PushMailSearchSliceResult` 必須帶回 `commandId`、`parentCommandId`、`sliceIndex`、`sliceCount` 與 `isSliceComplete=true`，讓 Hub 自行推算 slice 完成與整體進度。
- [ ] AddIn 不需要自行做跨 folder 排程；單一 folder 搜尋仍應避免 blocking Outlook UI。
- [ ] 發生 Outlook busy、search timeout 或使用者取消時，使用 `CompleteMailSearchSlice(success=false)` 並以匿名化 message 說明。

驗收：

- [ ] AddIn 不會自行處理 regex / fuzzy / keyword contains。
- [ ] AddIn 不會收到空 scope 後自行全域掃描。
- [ ] 單一 slice 只回 bounded candidates。
- [ ] Outlook busy、timeout 或取消時，有失敗結果而不是卡住。

### 5. 修改郵件屬性

目前只實作 `update_mail_properties` 作為郵件屬性 mutation 入口。不要再新增或維護舊的單一 marker command handler，除非 contract 明確恢復使用。

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

- [ ] Outlook mail item 的 read / flag / category 狀態正確保存。
- [ ] 回推的 `PushMail` 包含最新 snapshot。
- [ ] 若新增 category，回推的 `PushCategories` 包含最新 master category list。

### 6. 移動與刪除郵件

刪除郵件有獨立 `delete_mail` command；但唯一允許實作仍是移動到 Outlook 的「刪除的郵件 / Deleted Items」folder。AddIn 不得直接呼叫 `MailItem.Delete()` 或永久刪除郵件。

- [ ] AddIn 收到 `move_mail`。
- [ ] AddIn 收到 `delete_mail` 時，用同一套移動流程移到 Deleted Items。
- [ ] 用 `moveMailRequest.mailId` 找回 mail item。
- [ ] 用 `destinationFolderPath` 找到 Outlook destination `Folder`。
- [ ] 呼叫 Outlook `MailItem.Move(destinationFolder)`。
- [ ] 若 command 是 `delete_mail` 或 destination 是「刪除的郵件 / Deleted Items」，仍只呼叫 `Move(destinationFolder)`，不可呼叫 `Delete()`。
- [ ] 移動後重新讀取目前 source folder 或以正確方式移除已移動 mail。
- [ ] invoke `ReportCommandResult`。
- [ ] invoke `PushMails`，讓目前 mail list 反映移動後結果。
- [ ] 用 folder 增量同步更新 source 與 destination folder item count。

驗收：

- [ ] `move_mail` 可把 mail 移到指定 folder。
- [ ] `delete_mail` 只會 move to Deleted Items，不會永久刪除。
- [ ] Source folder mail snapshot 不再包含已移動 mail。
- [ ] 目的 folder item count 增加，source folder item count 減少。
- [ ] 跨 PST / OST 移動若 EntryID 改變，AddIn 仍會回推最新 mail snapshot。

### 7. Master Categories

- [ ] AddIn 收到 `fetch_categories`。
- [ ] 從 Outlook session master category list 讀取所有 category。
- [ ] 回推 `PushCategories(categories)`。
- [ ] AddIn 收到 `upsert_category`。
- [ ] 若 category 不存在，建立 category。
- [ ] 若 category 已存在，更新 color / shortcut key。
- [ ] 回推 `PushCategories(categories)`。

驗收：

- [ ] `PushCategories` 能反映 Outlook master category list。
- [ ] 新增或更新 category 後，回推最新 category snapshot。

### 8. Calendar 月曆

- [ ] AddIn 收到 `fetch_calendar`。
- [ ] 使用 `calendarRequest.startDate` 與 `calendarRequest.endDate` 讀取整個月份。
- [ ] `startDate` 含當日，`endDate` 不含當日。
- [ ] 回推區間內所有 calendar events。
- [ ] invoke `PushCalendar(events)`。

驗收：

- [ ] 回推的 event 落在 requested date range 內。
- [ ] Event 欄位包含 subject、時間、location、organizer、attendees、busy status。

### 9. Rules Snapshot

- [ ] AddIn 收到 `fetch_rules`。
- [ ] 讀取 Outlook rules snapshot。
- [ ] 回推 `PushRules(rules)`。

驗收：

- [ ] 回推的 rules 包含 rule name、enabled、order、conditions、actions、exceptions。

### 10. Chat

AddIn 送 chat 必須使用 `/hub/outlook-addin` 的 SignalR method，不要再用 HTTP `/api/outlook/chat`。

- [ ] AddIn 要送 chat message 時 invoke `SendChatMessage(message)`。
- [ ] `message.source` 建議填 `outlook`。
- [ ] `message.text` 填入要顯示的訊息。
- [ ] 不需要自行呼叫其他 SignalR hub 或 HTTP endpoint。

驗收：

- [ ] AddIn invoke `SendChatMessage` 成功。
- [ ] AddIn code 不再呼叫 HTTP `/api/outlook/chat`。

### 11. Folder 建立與刪除

- [ ] AddIn 收到 `create_folder`。
- [ ] 用 `parentFolderPath` 找到 parent folder。
- [ ] 建立 `name` 指定的新 folder。
- [ ] 用 folder 增量同步回推 folder 變更。
- [ ] AddIn 收到 `delete_folder`。
- [ ] 用 `folderPath` 找到 folder 並刪除。
- [ ] 用 folder 增量同步回推 folder 變更。
- [ ] 若目前 mail list 指向已刪除 folder，回推 `PushMails` 清掉或更新畫面。

驗收：

- [ ] 新增子 folder 後，folder snapshot 包含新 folder。
- [ ] 刪除 folder 後，folder snapshot 不再包含該 folder。

## 常見失敗對照

| 現象 | 優先檢查 |
| --- | --- |
| 已讀/未讀出現 missing mail id | `PushMails` 是否填 `MailItemDto.id` |
| `move_mail` 沒有執行 | mail 是否有 `id`，以及 destination folder path 是否可解析 |
| Hub 看不到 command | AddIn 是否有 SignalR connection |
| Category 空白 | AddIn 是否處理 `fetch_categories` 並 `PushCategories` |
| Flag 修改沒效果 | AddIn 是否處理 `update_mail_properties` 的 flag 欄位並儲存 item |
| Calendar 空白 | AddIn 是否處理 `fetch_calendar` 的 `startDate/endDate` 並 `PushCalendar` |
| AddIn chat 沒送出 | 是否 invoke `/hub/outlook-addin` 的 `SendChatMessage`，而不是 HTTP `/api/outlook/chat` |
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
## Contract 對照

實作時請以 `docs/addin/signalr-contract.md` 作為 JSON / DTO 欄位準則。本文件只描述「要做哪些功能」與「怎樣算做對」；payload 範例、DTO 欄位速查與 SignalR method 名稱以 contract 文件為準。
