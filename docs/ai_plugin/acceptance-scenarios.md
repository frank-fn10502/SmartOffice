# SmartOffice Outlook Skill Acceptance Scenarios

本文件是 SmartOffice Outlook Agents SKILL 的驗收情境清單。每個情境都應以使用者角度模擬：agent 必須照 `SKILL.md`、`references/workflows.md` 與 `references/http-api.md` 操作，不可跳過 request / fetch-result pattern，也不可猜 folder path、mail id 或 Outlook category color。

驗收時不一定要真的操作 Outlook；若是模擬，請逐步寫出 agent 會呼叫的 endpoint、request body 重點、判斷條件、使用者回覆內容，以及遇到不確定資料時是否停止詢問。

## 通用驗收規則

- 每個 `request-*` 後必須解析 `requestId`、`request`、`state`、`message`、`data`。
- 每個 Outlook operation 必須呼叫 paired `fetch-result-*` 到 `state=completed`，或在 `failed`、`unavailable`、`timeout` 時停止並回報。
- 回覆使用者時必須說明 folder scope，例如主要 Inbox、指定 folder、預設包含 subfolders 或使用者明確排除 subfolders、目前已載入的可搜尋 folders。
- 未指定 folder 時，預設查主要 mailbox 的 Inbox 與其 subfolders，不可改成全信箱或空 `scopeFolderPaths`。
- 所有 `folderPath` 必須來自 folder fetch result，不可自行組路徑。
- 所有 mail mutation 必須使用 `data.mails[].id` 作為 `mailId`，且 `folderPath` 必須來自同一筆 mail。
- 對 subject、sender、folder name 或 category name 有多個候選時，必須列出必要 metadata 請使用者確認，不可任選。
- 使用者只要求 metadata、清單或統計時，不可讀 mail body。
- 大量資料只能摘要必要欄位；mail body、folder name、category name、attachment path 與 chat message 都視為敏感資料。

## API / Connection

- [ ] API online：讀 `GET /api/outlook/admin/status`，`connected=true` 時繼續工作。
- [ ] API unavailable：`connected=false` 或 fetch result `state=unavailable` 時，停止並回報目前無法完成即時 Outlook request。
- [ ] Request accepted：確認 `accepted` 只代表 SmartOffice API 收下 request，不代表 Outlook 操作成功。
- [ ] Fetch result pagination：`next.hasMore=true` 時使用 `next.cursor` 讀下一頁，直到沒有下一頁。
- [ ] Wrong fetch-result endpoint：`requestId does not match this fetch-result endpoint` 時回報流程錯誤，不改用猜測資料。
- [ ] Unknown request id：fetch result 回 `request not found` 時回報 request 已不存在或服務重啟可能造成資料失效。
- [ ] Timeout：`state=timeout` 時停止，不重複送出 mutation request。

## Folder Scope

- [ ] 封裝查找：使用者要求拿 `folderAAA` 或指定 folder 名稱時，優先使用 `request-find-folder` -> `fetch-result-find-folder`，不可讓 caller 自行組 path。
- [ ] Find folder 唯一：`matchCount=1` 時使用 `data.folders[0].folderPath` 作為後續 API scope。
- [ ] Find Inbox：未指定 folder 時，優先用 `request-find-folder` with `folderType="Inbox"`、`storeId=""` 定位主要 store Inbox，不硬寫 `Inbox` 或 localized name。
- [ ] Find folder 同名：`isAmbiguous=true` 時列出必要路徑與 store 顯示名稱請使用者確認。
- [ ] Find folder 找不到：`matchCount=0` 時停止並回報，不改用 Inbox、空 scope 或猜測 path。
- [ ] 未指定 folder：載入 folders，定位主要 store root，再展開 children 找主要 Inbox。
- [ ] 中文 Inbox：能用 `folderType="Inbox"` 找到 `/主要信箱 - User/收件匣`，不硬寫英文 Inbox。
- [ ] Root children 未載入：看到 `childrenLoaded=false` 時呼叫 `request-folder-children`。
- [ ] 指定完整 folder path：用 folder fetch result 做完全比對後才操作。
- [ ] 指定 folder 顯示名稱：用 `name` 比對並定位唯一 `folderPath`。
- [ ] 同名 folder：找到多個 `folderA` 時列出必要路徑請使用者確認。
- [ ] 找不到 folder：停止並回報無法定位，不改用 Inbox、空 scope 或猜測路徑。
- [ ] 目的 folder 與來源 folder 相同：移動任務停止並回報。
- [ ] 使用者要求包含 subfolders：`request-folder-mails.includeSubFolders=true`。
- [ ] 使用者未提 subfolders：`request-folder-mails.includeSubFolders=true`，並在回覆中說明預設包含 subfolders。
- [ ] 使用者明確排除 subfolders：`request-folder-mails.includeSubFolders=false`。

## Mail List / Date Lookup

- [ ] 最近 N 封郵件：使用 `request-mails`，取得主要 Inbox path，設定 `maxCount=N`。
- [ ] 最近郵件但未給數量：使用合理 `maxCount`，回覆中說明只列 metadata。
- [ ] 這週的郵件：使用 `request-mail-search` 空 keyword，`receivedFrom` 為本週一 00:00，`receivedTo` 為目前時間。
- [ ] 本月郵件：使用 `request-mail-search` 空 keyword，`receivedFrom` 為本月 1 日 00:00。
- [ ] 最近 N 天郵件：使用 `request-mail-search`，不要用受 `maxCount` 限制的 `request-mails`。
- [ ] 最近兩個月郵件：使用 `request-mail-search`，`receivedFrom` 為目前時間往前推兩個月。
- [ ] 使用者說「這兩個月」：預設按最近兩個月，並在回覆中說明實際日期範圍。
- [ ] 指定日期區間：直接使用 `receivedFrom` / `receivedTo`，並回覆具體日期。
- [ ] 空結果：回覆指定範圍內沒有找到符合條件的郵件，不擴大 scope。

## Mail Search

- [ ] Search-first workflow：當使用者用 subject、sender、日期、category、附件、已讀、旗標或 folder scope 描述目標 mails 時，先用 `request-mail-search` 取得候選 metadata，再讀 body 或 mutation。
- [ ] Subject 搜尋：`keyword` 放使用者詞彙，`textFields=["subject"]`。
- [ ] Sender 搜尋：`textFields=["sender"]`，結果摘要用 `sender.displayName`，避免暴露 raw address。
- [ ] Body 搜尋：只有使用者明確要求內文關鍵字判讀時才用 `textFields=["body"]`。
- [ ] 全信箱搜尋：只有使用者明確要求全信箱或所有已載入 mail folders 時，才送 `scopeFolderPaths=[]`。
- [ ] 全信箱搜尋回覆：必須說明範圍是目前 SmartOffice API 已知的可搜尋 mail folders，不保證完整 Outlook mailbox。
- [ ] 指定 folder 搜尋：`scopeFolderPaths` 放該 folder 的真實 `folderPath`，預設 `includeSubFolders=true`。
- [ ] 指定 folder + subfolders 搜尋：`includeSubFolders=true`。
- [ ] 指定 folder 且明確不含 subfolders 搜尋：`includeSubFolders=false`。
- [ ] 有附件搜尋：`hasAttachments=true`，`keyword` 可空。
- [ ] 無附件搜尋：`hasAttachments=false`。
- [ ] 未讀搜尋：`readState="unread"`。
- [ ] 已讀搜尋：`readState="read"`。
- [ ] 旗標搜尋：`flagState="flagged"`。
- [ ] 分類搜尋：`categoryNames` 放使用者指定 category；多個 category 任一符合即可。
- [ ] 條件式批次搬移：例如「將 folderA 的 `待處理` 類別郵件搬到 folder_staged」時，用 `request-mail-search` 搭配 source `scopeFolderPaths`、`includeSubFolders=true`、`categoryNames=["待處理"]` 取得 mail ids，再分批 `request-move-mails`。
- [ ] `no_searchable_folder`：重新檢查 folder path 是否來自 folder result，不自行改成全域搜尋。

## Mail Body

- [ ] 使用者要求摘要內容：先以 metadata 定位唯一 mail，再呼叫 `request-mail-body`。
- [ ] 使用者只要清單：不呼叫 `request-mail-body`。
- [ ] 同 subject 多封：先列 `receivedTime`、`sender.displayName`、短 subject 請使用者確認。
- [ ] Body 仍為空：body request 完成後同 id 的 `body` / `bodyHtml` 仍空，停止重試並回報目前內容不可用。
- [ ] 大量摘要：只摘要必要內容，不貼完整 body。

## Attachments

- [ ] 讀附件 metadata：先由 mail result 取得 `id` 與 `folderPath`，再呼叫 `request-mail-attachments`。
- [ ] 有附件郵件清單：使用 mail search `hasAttachments=true`，不需要讀每封附件 metadata。
- [ ] 匯出附件：使用 `attachmentId` / `index` / `fileName` 等 metadata 呼叫 `request-export-mail-attachment`。
- [ ] 開啟附件：只能用已記錄的 `exportedAttachmentId` 呼叫 `open-exported-attachment`，不可傳任意本機路徑。
- [ ] Export result 空 data：知道 `fetch-result-export-mail-attachment` 目前不直接回 `exportedAttachmentId`，需重新讀 attachment metadata。
- [ ] 附件路徑敏感：回覆不貼完整本機路徑，除非使用者明確需要。

## Categories

- [ ] 讀 master categories：`request-categories` -> `fetch-result-categories`，列出 `name`、必要時列 color。
- [ ] 新增 category：`request-upsert-category` -> `fetch-result-upsert-category`。
- [ ] 更新 category 顏色：以 category `name` 為識別，送 `color` 與 `colorValue`。
- [ ] 黑色 category：使用 `olCategoryColorBlack` / `15`。
- [ ] 未知顏色：若使用者指定顏色不在 color table 中，請使用者改選，不猜 enum。
- [ ] Category 名稱重複：比對時大小寫不敏感；若已有同名 category，視為更新。
- [ ] 套用 category 到 mail：先定位唯一 mail，再用 `request-update-mail-properties.categories`。
- [ ] 新 category 套用到 mail：需要時用 `newCategories` 或先 `request-upsert-category`。

## Mail Mutations

- [ ] 標記已讀：定位唯一 mail 後呼叫 `request-update-mail-properties`，`isRead=true`。
- [ ] 標記未讀：`isRead=false`。
- [ ] 設定今日 follow-up：`flagInterval="today"`，必要時填 `flagRequest`。
- [ ] 完成 follow-up：`flagInterval="complete"`。
- [ ] 清除 flag：`flagInterval="none"`。
- [ ] 移動單封 mail：使用 `request-move-mail`，source / destination folder path 都必須已定位。
- [ ] 批次移動 mails：使用 `request-move-mails`，單批最多 500。
- [ ] 批次移動 500 以上：分批送 request，每批完成後再送下一批。
- [ ] 批次部分失敗：停止並回報已完成數與失敗批次，不假裝全部成功。
- [ ] 刪除 mail：只使用 `request-delete-mail`，語意是移到 Outlook default Deleted Items folder；完成後告知使用者已移到刪除資料夾，不代表永久刪除。
- [ ] 使用者要求永久刪除 mail：回覆 SmartOffice API 不做永久刪除，若要永久刪除請使用者自行到 Outlook 操作。
- [ ] 以 subject 刪除 mail：先 subject search，再精準比對唯一 mail；多封時要求確認。
- [ ] Mutation 後確認：用 paired fetch result 或重新送出必要 request 確認結果。

## Folder Mutations

- [ ] 建立 folder：先定位 parent folder path，再 `request-create-folder`。
- [ ] 刪除 folder：先定位唯一 folder path，再 `request-delete-folder`，語意是移到 Outlook default Deleted Items folder；完成後告知使用者已移到刪除資料夾，不代表永久刪除。
- [ ] 刪除 Deleted Items 內 folder：fetch result `message=manual_delete_required` 時停止，回報 SmartOffice API 會阻擋此操作；若要永久刪除請使用者自行到 Outlook 操作。
- [ ] 使用者要求永久刪除 folder：回覆 SmartOffice API 不做永久刪除，若要永久刪除請使用者自行到 Outlook 操作。
- [ ] Folder delete 後確認：讀 folder result 確認 folder tree 或刪除資料夾相關結果。
- [ ] 建立同名 folder：若 API / AddIn 回失敗，回報原因，不自行改名。

## Calendar / Rules / Chat

- [ ] 查 calendar 未給範圍：使用 `request-calendar` 預設或 `daysForward=31`。
- [ ] 查指定日期 calendar：使用 `startDate` / `endDate`。
- [ ] Calendar 結果：摘要 `subject`、`start`、`end`、`location`，避免大量參與者資訊。
- [ ] 查 rules：`request-rules` -> `fetch-result-rules`。
- [ ] Chat：chat message 視為敏感資料；非必要不大量回放 chat history。

## Error / Safety

- [ ] HTTP 400 invalid request：解析 body 的 `state` / `message`，回報欄位問題。
- [ ] HTTP 409：解析 body，不猜測重試。
- [ ] HTTP 502 / failed：回報 AddIn 或 Outlook automation 失敗訊息。
- [ ] HTTP 504 / timeout：停止並回報 timeout。
- [ ] Service restart：舊 request id 查不到時，回報需要重新送出 request。
- [ ] Sensitive output：不輸出完整 mail body、raw Exchange address、大量 folder path 或 attachment path。
- [ ] User asks for global action：先確認 scope 與 destructive nature，尤其 delete / move 大量郵件。
- [ ] Ambiguous destructive request：要求使用者確認唯一 mail / folder / scope 後再 mutation。

## Regression Scenarios From Earlier Reviews

- [ ] `fetch-result-*` response 使用 `state`，不是 `status`。
- [ ] `request-*` response 沒有 `success` 欄位。
- [ ] Folder discovery 讀 `data.stores` / `data.folders`，不是 `data.mails`。
- [ ] `request-folder-children` body 使用 `parentEntryId` / `parentFolderPath`，不是 `entryId` / `folderPath`。
- [ ] `request-folder-mails` 使用 `fetch-result-folder-mails`，不是 legacy `GET /api/outlook/folder-mails`。
- [ ] `request-mail-search` 使用 `fetch-result-mail-search`，不是 legacy `GET /api/outlook/mail-search`。
- [ ] `request-export-mail-attachment` 的 fetch result 不直接回 `exportedAttachmentId`。
- [ ] Swagger 中所有 `fetch-result-*` endpoint 分類清楚，不落到預設 `Outlook` tag。
- [ ] Skill 名稱、folder、installer 與 external docs 不包含內部實作術語。
- [ ] 若 API / DTO / route / workflow / error semantics 有變更，`skills/smartoffice-outlook/SKILL.md` 與 `references/http-api.md`、`references/workflows.md` 已同步；外部 AI 不需要讀 repo `AGENTS.md` 也能理解操作方式。
