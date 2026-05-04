# Outlook AddIn 目前所需功能

## 目的

本文件先列出工作機 Outlook AddIn 需要完成與實測的功能範圍。這裡不是最終切分任務；後續應依工作機實測結果、官方文件限制與 Hub contract 是否需要調整，再拆成 `Plan/NNN-*.md` 任務。

## 工作原則

- 實作位置是工作機完整 SmartOffice solution，不是本 repository。
- Hub repository 只提供 HTTP API、SignalR、command routing、temporary state 與 contract。
- AddIn 程式碼必須以 Microsoft 官方文件與工作機 Outlook 2016 實測結果為準。
- 每個功能完成時都要輸出真實結果：Office API call、成功或失敗、匿名化 JSON sample、錯誤訊息與是否符合 Hub DTO。
- 測試回報不得包含真實 mail body、folder name、rule name、calendar subject、attendee、客戶名稱或公司內部資訊。

## 必讀文件

- `..\SmartOffice.Hub\AGENTS.md`
- `..\SmartOffice.Hub\docs\ai\workstation-solution.md`
- `..\SmartOffice.Hub\docs\ai\protocols.md`
- `..\SmartOffice.Hub\docs\addin\features-checklist.md`
- `..\SmartOffice.Hub\docs\addin\signalr-contract.md`
- `..\SmartOffice.Hub\docs\addin\outlook-references.md`
- `..\SmartOffice.Hub\docs\addin\test-report.md`
- `..\SmartOffice.Hub\Models\Dtos.cs`

## 連線與診斷

1. AddIn 可設定 Hub URL，預設連到 `http://localhost:2805`。
2. AddIn 連線到 `/hub/outlook-addin`。
3. AddIn 連線成功後 invoke `RegisterOutlookAddin(info)`。
4. AddIn listen `OutlookCommand`，不可造成 Outlook 卡頓。
5. 每次收到 command 都記錄 `commandId`、`type`、開始時間、結束時間、成功或失敗。
6. 發生錯誤時 invoke `ReportAddinLog(entry)` 與 `ReportCommandResult(result)`，message 必須匿名化。
7. 工作機維護 `Plan\WORKSTATION-STATUS.md`，記錄 Hub URL、AddIn 專案名稱、SignalR connection handler、command handler 位置、主要 automation class、blocker 與驗證紀錄。

## 資料讀取功能

1. 讀取 folder tree，透過 SignalR invoke `PushFolders(folders)`。
2. 讀取指定 folder 的 mails，支援 `folderPath`、`range` 與 `maxCount`，透過 SignalR invoke `PushMails(mails)`。
3. 郵件資料需盡量取得 `EntryID` 或其他可重新定位的 stable id。
4. 郵件 metadata 需實測 `categories`、`isRead`、`flagRequest`、`flagInterval`、`taskStartDate`、`taskDueDate`、`taskCompletedDate`、`importance`、`sensitivity`。
5. 讀取 Outlook rules，透過 SignalR invoke `PushRules(rules)`。
6. 讀取 master category list，透過 SignalR invoke `PushCategories(categories)`。
7. 讀取 calendar events，支援 `daysForward`，透過 SignalR invoke `PushCalendar(events)`。

## 郵件操作功能

1. 依 `mailId` 與 `folderPath` 重新定位單封郵件。
2. 將郵件標記為已讀。
3. 將郵件標記為未讀。
4. 設定 follow-up flag。
5. 清除 follow-up flag。
6. 設定或覆蓋單封郵件 categories。
7. 使用較新的 `mailPropertiesRequest` 時，可一次更新 read state、flag、task dates、categories 與需要新增的 master categories。
8. 移動單封郵件到指定 destination folder。
9. 每個 mutation 完成後，至少重新 invoke `PushMails` 回報受影響 folder 的 mails；若 folder count 或 categories 受影響，也要 invoke `PushFolders` 或 `PushCategories`。

## Folder 與 Category 操作功能

1. 在指定 parent folder 建立子 folder。
2. 刪除指定 folder。
3. 新增或更新 Outlook master category。
4. category color 與 shortcut key 必須記錄 Office 2016 實測支援狀態。
5. 每個操作完成後重新 invoke `PushFolders` 或 `PushCategories`。

## 工作機實測重點

1. 確認 Outlook 2016 使用 VSTO / COM AddIn 或 Office.js 的實際架構；若是 Office.js，先確認 requirement set 是否支援所需 API。
2. 確認 folder path 是否能穩定 round-trip：Hub request 的 `folderPath` 能否在 Outlook object model 中重新定位。
3. 確認 `MailItem.EntryID` 在 move、restart Outlook、不同 store 或 Exchange cached mode 下是否穩定。
4. 確認 `MailItem.Body`、`HTMLBody`、`SenderEmailAddress`、`ReceivedTime` 在工作帳號上的真實行為。
5. 確認 rules 與 master categories 的讀寫 API 在 Outlook 2016 是否可用。
6. 確認 mutation command 失敗時不會造成未回報狀態或重複執行。
7. 每次實測都產出匿名化 JSON sample，並標註是否可直接符合 `Models/Dtos.cs`。

## 暫定優先順序

1. SignalR 連線、Hub URL 設定、`OutlookCommand` dispatch、log/result 回報。
2. Folder tree 讀取與 folderPath round-trip 驗證。
3. Mail 讀取與 stable mail id 驗證。
4. Mail metadata 讀取。
5. Mail mutation：read/unread、flag、categories。
6. Folder mutation：create/delete folder。
7. Move mail。
8. Master category list 讀取與 upsert。
9. Rules 讀取。
10. Calendar 讀取。

## 後續拆分方式

等工作機確認 AddIn 架構與第一輪官方文件對照後，再依 `docs/ai/plan-splitting.md` 把本清單拆成單一目標、單一驗證的 `Plan/NNN-*.md` 任務。
