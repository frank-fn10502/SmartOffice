---
name: smartoffice-outlook
description: Use when an AI agent needs to operate Outlook through the SmartOffice Outlook HTTP API, including reading folders, mails, mail body, mail conversations, attachments, calendar, categories, rules, mail search, and safe mail or folder mutations. Default base URL is http://localhost:2805. Use this skill for Outlook client workflows, not for implementing the API itself.
metadata:
  owner: "SmartOffice"
  skill_id: "smartoffice-outlook.skill.smartoffice.2026-05"
---

# SmartOffice Outlook

## 邊界

- 一律透過 Outlook HTTP API 操作 Outlook；不要繞過 API 使用其他通道。
- 預設 base URL 是 `http://localhost:2805`；若使用者明確提供其他 API URL，才改用該 URL。
- 可用任何可呼叫 HTTP 與解析 JSON 的工具；不要把 `curl`、PowerShell 或特定 shell 視為必要條件。
- 每次 `request-*` 後，先解析 request response 的固定欄位：`requestId`、`request`、`state`、`message`、`data`。`data.fetchResultEndpoint` 是下一步要呼叫的 paired endpoint；response 沒有 `success` 欄位，`accepted` 只代表 SmartOffice API 已收下 request，不代表 Outlook 操作已成功。
- 取得 `requestId` 後呼叫 `data.fetchResultEndpoint`；回應固定提供 `requestId`、`request`、`state`、`message`、`next`、`data`。
- 使用者或 AI 只需要 loop paired `fetch-result-*`。每次都先檢查 `state`；若是 `failed`、`unavailable` 或 `timeout` 就停止並回報。若 `next.hasMore=true`，即使本頁 `state=completed`，仍要用 `cursor=next.cursor` 取下一頁；只有 `state=completed` 且 `next.hasMore=false` 才代表資料取完。
- 若 `request-*` 回 HTTP 409 / 400 / 502 / 504，仍先解析 body 的 `state`、`message` 與可能存在的 `requestId`；不要靠猜測重試，也不要改用更大的搜尋範圍。
- HTTP request body 不接受未文件化欄位；未知欄位會回 `400 invalid_request_body`。看到這個錯誤時，改用 Swagger 或任務對應 reference 裡的精確欄位名稱，不要猜 alias，例如 `request-mail-search` 用 `keyword` 而不是 `query`，`request-calendar` 用 `daysForward` 而不是 `lookaheadDays`。
- 看到 `400 missing_required_fields` 時，只補齊 `requiredFields` 列出的欄位後重送；不要更換 endpoint、擴大搜尋範圍或省略 `folderPath`。
- `request-mail-search` 是多數 mail workflow 的前置定位與篩選入口，不只是「文字搜尋」。凡是使用者用 subject、sender、日期、category、附件、已讀、旗標或 folder scope 描述目標 mails，優先用 `request-mail-search` 找出候選 metadata，再決定是否讀 body 或 mutation。
- 修改郵件前必須先從 `fetch-result-*` 的 `data.mails` 取得 `MailItemDto.id` 與 `folderPath`，不可只用 subject、sender 或 folder name 猜目標。
- 修改屬性前必須確認 `MailItemDto.messageClass`。空字串或 `IPM.Note` 可視為一般郵件；`IPM.Schedule.Meeting.*` 是會議邀請/更新，可讀 metadata/body/attachments，也可移動或移到 Deleted Items，但不可呼叫 `request-update-mail-properties`。API 若回 `unsupported_outlook_item_type`，停止並告知使用者需用 Outlook 會議/行事曆流程處理該屬性操作。
- `mail body`、`folderPath`、`category`、`attachment path`、`chat message` 都可能含敏感 business data；只摘要必要資訊，不在回覆中大量外洩。
- 使用者只要求最近郵件、郵件清單、統計或 Markdown metadata 報告時，不要呼叫 `request-mail-body`；只有使用者明確要求內容摘要、內文關鍵字判讀，或 metadata 不足以完成任務時才讀 body。
- `request-delete-mail` 的語意是移到 Outlook default Deleted Items folder，不是永久刪除；目的 folder 由 SmartOffice 用 Outlook default folder identity 定位，不靠顯示名稱猜測。完成後告知使用者 mail 已移到刪除資料夾；若使用者要永久刪除，請使用者自行到 Outlook 操作。
- `request-delete-folder` 的語意也是移到 Outlook default Deleted Items folder；若目標 folder 已經位於 default Deleted Items folder 或其子層，SmartOffice API 會以 `manual_delete_required` 阻擋，agent 必須停止並請使用者自行到 Outlook 操作。不要用 `Deleted Items`、`刪除的郵件` 或其他本地化顯示名稱自行判斷。
- 使用者未指定 folder 時，預設查主要 mailbox 的 Inbox 與其 subfolders；Inbox path 優先用 `request-find-folder` 搭配 `folderType="Inbox"` 取得，不可硬寫英文 `Inbox`。
- HTTP API 的 `folderPath` 一律使用 `/主要信箱 - User/收件匣` 這種普通斜線格式。
- `scopeFolderPaths: []` 代表目前已載入的全部可搜尋 mail folders；使用者未明確要求全域搜尋時禁止送空陣列。
- 回覆使用者時必須說明本次 folder 範圍，避免使用者誤以為已搜尋全信箱。
- 使用者指定 folder 顯示名稱或要求「拿 folderAAA」時，優先用 `request-find-folder` / `fetch-result-find-folder` 定位候選 folder。若 `matchCount=1`，後續使用該 folder 的真實 `folderPath`；若 `isAmbiguous=true`，列出必要路徑請使用者確認，不可任選。
- 相對日期要轉成明確 `receivedFrom` / `receivedTo` 並在回覆中說明，例如「這週」代表本週一 00:00 到目前時間；「最近兩個月」代表從目前時間往前推兩個月到目前時間。若使用者說「這兩個月」且語意不明，預設按最近兩個月處理。
- 大型 response 可暫存於 skill folder 的 `tmp/<run>/`，但不可提交或長期保留。

## 可選 Helper Scripts

Skill 內附兩類 script。`scripts/outlook-api.sh` 與 `scripts/outlook-api.ps1` 是底層 helper，只封裝 HTTP 呼叫、request/fetch-result loop 與分頁；`scripts/recipes/` 放已經形成穩定流程的實際操作案例。所有 script 都輸出 JSON 到 stdout，方便 AI 直接解析。

- bash: `./scripts/outlook-api.sh status`
- PowerShell: `pwsh ./scripts/outlook-api.ps1 status`
- 指定 API URL: 設定 `SMARTOFFICE_OUTLOOK_BASE_URL`，或 bash 使用 `--base-url URL`，PowerShell 使用 `-BaseUrl URL`。
- 取得主要 Inbox: `./scripts/recipes/inbox.sh`
- 最近郵件: `./scripts/recipes/recent-mails.sh --lookback-hours 168 --max-count 30`
- 通用 request/fetch: `./scripts/outlook-api.sh request-fetch /api/outlook/request-calendar '{"daysForward":31,"startDate":null,"endDate":null}'`

PowerShell recipe 對應為 `pwsh ./scripts/recipes/inbox.ps1` 與 `pwsh ./scripts/recipes/recent-mails.ps1 -LookbackHours 168 -MaxCount 30`。只有當流程已經有實際案例、包含多步驟或容易做錯時，才新增 recipe；單一步驟或仍在探索中的呼叫請用底層 helper。

使用 helper 或 recipe 時仍需遵守本文件規則：修改前先確認唯一 `mailId` / `folderPath`，只摘要必要資料，並在回覆使用者時說明 folder 範圍。

## Golden Path

使用者未指定 folder 時，請照這個最短正確路徑取得主要 mailbox 的 Inbox：

1. 確認 API status。
2. `POST /api/outlook/request-find-folder`，body 使用 `folderType="Inbox"`、`storeId=""`。未指定 `storeId` 時，SmartOffice API 會在主要 store 查找該 folder type。
3. 取得 `requestId` 後 loop `POST /api/outlook/fetch-result-find-folder`，直到 `state=completed` 且 `next.hasMore=false`。
4. 若 `data.matchCount=1`，使用 `data.folders[0].folderPath` 呼叫 `request-mails` 或放入 `request-mail-search.scopeFolderPaths[0]`。
5. 若 `matchCount=0` 或 `isAmbiguous=true`，停止並回報目前無法唯一定位主要 Inbox；不要自行改成全信箱或空 scope 搜尋。
6. 取得後續 request 的 `requestId` 後 loop paired `fetch-result-*`，從 `data` 讀資料。
7. 摘要必要 metadata；只有任務需要時才讀 body 或 attachment。
8. mutation 後用該次 `fetch-result-*` 或重新送出必要 request 確認結果。

找不到主要 Inbox 時，停止並回報目前 folder 資料無法定位 Inbox；不要自行改成全信箱或空 scope 搜尋。

## 操作流程速查

- 最近 N 封郵件：Golden Path 取得 Inbox path -> `request-mails` -> `fetch-result-mails`。
- 日期範圍郵件：Golden Path 取得 Inbox path -> `request-mail-search`，用空 `keyword` 搭配 `receivedFrom` / `receivedTo` -> `fetch-result-mail-search`。
- 指定 folder：用 `request-find-folder` -> `fetch-result-find-folder` 定位唯一 `folderPath`。
- 指定 folder 內所有 mails：先 `request-find-folder` 取得 folder path -> `request-folder-mails` -> `fetch-result-folder-mails`。
- 郵件定位與篩選：Golden Path 或 `request-find-folder` 取得 scope path -> `request-mail-search`，用 `keyword`、`textFields`、`categoryNames`、`hasAttachments`、`flagState`、`readState`、`receivedFrom` / `receivedTo` 組合條件 -> `fetch-result-mail-search`。
- 讀 body：先從 `fetch-result-* data.mails` 取 `id` 與 `folderPath` -> `request-mail-body` -> `fetch-result-mail-body data.mails` 找同 id 的 `body` / `bodyHtml`。
- 讀 conversation：先從 `fetch-result-* data.mails` 取 `id` 與 `folderPath` -> `request-mail-conversation` -> `fetch-result-mail-conversation data.mails`。只有使用者需要一次性查看討論串時才讀；若包含 body，摘要必要內容即可。
- 讀附件：先從 `fetch-result-* data.mails` 取 `id` 與 `folderPath` -> `request-mail-attachments` -> `fetch-result-mail-attachments`。
- 查 group 與個人關聯：優先使用 `request-address-book-relation` -> `fetch-result-address-book-relation`。輸入 `groupSmtpAddress` / `groupId` 可反查 group members、nested groups、上層 groups 與是否和自己相關；輸入 `email` / `smtpAddress` / `displayName` / `id` 可反查個人所屬 groups。若結果 `state=ambiguous`，用 `matches[]` 請使用者縮小目標；不要自行任選。詳細欄位見 `references/organizing.md`。
- 修改、移動、刪除郵件：先從 `fetch-result-* data.mails` 確認唯一目標的 `id` 與 `folderPath` -> mutation endpoint -> `fetch-result-*`。
- 大量搬移 folder 內全部郵件：先定位來源與目的 `folderPath`，用 `request-folder-mails` 取得 ids，再以每批最多 500 封逐批呼叫 `request-move-mails`。預設包含 subfolders；只有使用者明確排除 subfolders 時才設定 `includeSubFolders=false`。
- 大量搬移符合條件的郵件，例如 category、日期、附件、已讀或旗標：先定位來源與目的 `folderPath`，用 `request-mail-search` 篩出目標 mails，再分批 `request-move-mails`。詳細流程見 `references/bulk-move.md`。

## 何時讀 references

- 需要共通 request/fetch-result envelope、錯誤格式或 endpoint 配對表時，讀 `references/http-api.md`。
- 需要 folder discovery、Inbox 定位、folder children 或 `folderPath` 規則時，讀 `references/folders.md`。
- 需要 mail list、mail search、body、conversation、attachment 或 attachment export 時，讀 `references/mail.md`。
- 需要 calendar、rules、categories、mail/folder mutation、chat、通訊錄 group/個人關聯或 DTO 欄位時，讀 `references/organizing.md`。
- 需要跨 endpoint 的操作順序、folder scope 判斷、日期範圍或批次搬移流程時，讀 `references/workflows.md`。
- 需要搬移大量郵件、搬空 folder tree 或分批 `request-move-mails` 時，讀 `references/bulk-move.md`。

## 常見陷阱

- `Inbox` 是範例名稱，不是穩定 folder path；中文 Outlook 常見路徑是 `/主要信箱 - User/收件匣`。操作前一定要從 folder fetch result 取實際 `folderPath`。
- 一般 folder 查找先用 `request-find-folder`，不要讓 caller 自己重寫每個分支的 traversal。只有需要精細控制載入範圍或診斷 folder discovery 時，才直接使用 `request-folders` / `request-folder-children`。
- `request-find-folder.folderType="Inbox"` 可用來取得主要 store 的 Inbox；只有需要精細控制載入範圍或診斷時，才用 `request-folders` / `request-folder-children` 展開後再找 Inbox。
- `request-folder-children` 的 request 欄位是 `storeId`、`parentEntryId`、`parentFolderPath`；值分別取自 folder fetch result 內 root folder 的 `storeId`、`entryId`、`folderPath`。不要送 `entryId` 或 `folderPath` 這兩個錯誤欄位名，也不要只傳 folder display name。
- `request-mails.folderPath`、`request-folder-mails.folderPath` 與 `request-mail-search.scopeFolderPaths[]` 必須完整等於 folder fetch result 裡的 `folderPath`。
- `request-mail-search` 或 `request-folder-mails` 回 `no_searchable_folder` 時，通常代表指定 folder path 目前無法搜尋；此時不要改成全域搜尋，應先重新讀 folders 並改用回傳的真實路徑。
- 不要自行組 folder path；一律使用 folder fetch result 回傳的 `/Mailbox/Inbox` 形式。
- 對同一封 mail 呼叫 `request-mail-body` 完成後，若同 id 的 `body` 與 `bodyHtml` 仍為空，不要重複呼叫同一 endpoint；將該封內容視為目前不可用，回報限制或改用 metadata。
- `request-move-mails` 單次最多 500 個 `mailIds`。遇到「搬移 folderA 和所有 subfolder」這類任務時，必須分批慢慢送，不可把 8000+ ids 放進單一 request。
- Outlook master category 顏色必須使用 Outlook `OlCategoryColor` enum name 與 numeric value；黑色是 `olCategoryColorBlack` / `15`。若使用者指定的顏色不在 color table 中，先請使用者改選，不要猜 enum。
- Category name 比對大小寫不敏感；若 master category 已有同名項目，視為更新而不是新增第二個。套用 category 到 mail 前，仍必須先定位唯一 mail。
