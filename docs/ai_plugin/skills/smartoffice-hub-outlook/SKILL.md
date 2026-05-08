---
name: smartoffice-hub-outlook
description: Use when an AI agent needs to operate Outlook through the SmartOffice Outlook HTTP API, including reading folders, mails, mail body, attachments, calendar, categories, rules, mail search, and safe mail or folder mutations. Default base URL is http://localhost:2805. Use this skill for Outlook client workflows, not for implementing the API itself.
metadata:
  owner: "SmartOffice"
  skill_id: "smartoffice-hub-outlook.skill.smartoffice-hub.2026-05"
---

# SmartOffice Outlook

## 邊界

- 一律透過 Outlook HTTP API 操作 Outlook；不要繞過 API 使用其他通道。
- 預設 base URL 是 `http://localhost:2805`；若使用者明確提供其他 Hub URL，才改用該 URL。
- 可用任何可呼叫 HTTP 與解析 JSON 的工具；不要把 `curl`、PowerShell 或特定 shell 視為必要條件。
- 每次 `request-*` 後，先解析 request response 的固定欄位：`requestId`、`request`、`state`、`message`、`data`。`data` 是各 request 自己的 struct；response 沒有 `success` 欄位，`accepted` 只代表 Hub 已收下 request，不代表 Outlook 操作已成功。
- 取得 `requestId` 後呼叫配對的 `POST /api/outlook/fetch-result-*`；回應固定提供 `requestId`、`request`、`state`、`message`、`next`、`data`。
- 使用者或 AI 只需要 loop paired `fetch-result-*` 到 `state=completed`；若 `next.hasMore=true`，下一次 body 帶 `cursor=next.cursor`。
- 若 `request-*` 回 HTTP 409 / 400 / 502 / 504，仍先解析 body 的 `state`、`message` 與可能存在的 `requestId`；不要靠猜測重試，也不要改用更大的搜尋範圍。
- 修改郵件前必須先從 `fetch-result-*` 的 `data.mails` 取得 `MailItemDto.id` 與 `folderPath`，不可只用 subject、sender 或 folder name 猜目標。
- `mail body`、`folderPath`、`category`、`attachment path`、`chat message` 都可能含敏感 business data；只摘要必要資訊，不在回覆中大量外洩。
- 使用者只要求最近郵件、郵件清單、統計或 Markdown metadata 報告時，不要呼叫 `request-mail-body`；只有使用者明確要求內容摘要、內文關鍵字判讀，或 metadata 不足以完成任務時才讀 body。
- `request-delete-mail` 的語意是移到 Outlook default Deleted Items folder，不是永久刪除；目的 folder 由 AddIn 用 Outlook default folder identity 定位，不靠顯示名稱猜測；仍需先確認 `mailId` 來自 `fetch-result-* data.mails[].id`，且 `folderPath` 來自同一筆 mail。若目標已經在 Deleted Items 內，paired fetch result 會回 `state=failed` / `message=manual_delete_required`，此時要求使用者自行到 Outlook 永久刪除。
- 使用者未指定 folder 時，只查主要 mailbox 的 Inbox；但 Inbox path 必須取自 `fetch-result-folders` / `fetch-result-folder-children` 的 `data.folders`，不可硬寫英文 `Inbox`。
- HTTP API 的 `folderPath` 一律使用 `/主要信箱 - User/收件匣` 這種普通斜線格式。
- `scopeFolderPaths: []` 代表目前已載入的全部可搜尋 mail folders；使用者未明確要求全域搜尋時禁止送空陣列。
- 回覆使用者時必須說明本次 folder 範圍，避免使用者誤以為已搜尋全信箱。
- 大型 response 可暫存於 skill folder 的 `tmp/<run>/`，但不可提交或長期保留。

## Golden Path

使用者未指定 folder 時，請照這個最短正確路徑取得主要 mailbox 的 Inbox：

1. 確認 API status。
2. `POST /api/outlook/request-folders`，取得 `requestId` 後 loop `POST /api/outlook/fetch-result-*` 到 `state=completed`。
3. 從 `fetch-result-folders .data.stores` 與 `fetch-result-folders .data.folders` 找主要 store/root folder。
4. 若 root folder `childrenLoaded=false`，呼叫 `POST /api/outlook/request-folder-children`，request 欄位必須是 `storeId=root.storeId`、`parentEntryId=root.entryId`、`parentFolderPath=root.folderPath`。
5. 再用 `fetch-result-folder-children` 讀 `data.folders`，在主要 store 底下選 `folderType="Inbox"` 的 folder；若沒有，才用 `name` 等於 `收件匣` 或 `Inbox` fallback。
6. 使用該 folder 的完整 `folderPath` 呼叫 `request-mails` 或放入 `request-mail-search.scopeFolderPaths[0]`。
7. 取得 `requestId` 後 loop `fetch-result-*`，從 `data` 讀資料。
8. 摘要必要 metadata；只有任務需要時才讀 body 或 attachment。
9. mutation 後用該次 `fetch-result-*` 或重新送出必要 request 確認結果。

找不到主要 Inbox 時，停止並回報目前 folder 資料無法定位 Inbox；不要自行改成全信箱或空 scope 搜尋。

## 操作流程速查

- 最近郵件：Golden Path 取得 Inbox path -> `request-mails` -> `fetch-result-*`。
- 指定 folder 內所有 mails：取得 folder path -> `request-folder-mails` -> `fetch-result-*`。
- 郵件搜尋：Golden Path 取得 Inbox path -> `request-mail-search` 且 `scopeFolderPaths` 放該 path -> `fetch-result-*`。
- 讀 body：先從 `fetch-result-* data.mails` 取 `id` 與 `folderPath` -> `request-mail-body` -> `fetch-result-mail-body data.mails` 找同 id 的 `body` / `bodyHtml`。
- 讀附件：先從 `fetch-result-* data.mails` 取 `id` 與 `folderPath` -> `request-mail-attachments` -> `fetch-result-*`，必要時讀 `mail-attachments?mailId={id}`。
- 修改、移動、刪除郵件：先從 `fetch-result-* data.mails` 確認唯一目標的 `id` 與 `folderPath` -> mutation endpoint -> `fetch-result-*`。
- 大量搬移整個 folder 或含 subfolders 的郵件：用 `request-folder-mails` 取得 ids，再以每批最多 500 封逐批呼叫 `request-move-mails`。詳細流程見 `references/workflows.md` 的「Bulk Move Folder Tree」。

## 何時讀 references

- 需要 endpoint、request/response 欄位或 enum 時，讀 `references/http-api.md`。
- 需要操作順序、folder scope 或搜尋流程時，讀 `references/workflows.md`。

## 常見陷阱

- `Inbox` 是範例名稱，不是穩定 contract；中文 Outlook 常見路徑是 `/主要信箱 - User/收件匣`。操作前一定要從 folder fetch result 取實際 `folderPath`。
- `request-folders` 只保證載入 stores/root folders；若 root 的 `childrenLoaded=false`，要用 `request-folder-children` 展開後再找 Inbox。
- `request-folder-children` 的 request 欄位是 `storeId`、`parentEntryId`、`parentFolderPath`；值分別取自 folder fetch result 內 root folder 的 `storeId`、`entryId`、`folderPath`。不要送 `entryId` 或 `folderPath` 這兩個錯誤欄位名，也不要只傳 folder display name。
- `request-mails.folderPath`、`request-folder-mails.folderPath` 與 `request-mail-search.scopeFolderPaths[]` 必須完整等於 folder fetch result 裡的 `folderPath`。
- `request-mail-search` 或 `request-folder-mails` 回 `no_searchable_folder` 時，通常代表指定 folder path 目前無法搜尋；此時不要改成全域搜尋，應先重新讀 folders 並改用回傳的真實路徑。
- 不要自行組 folder path；一律使用 folder fetch result 回傳的 `/Mailbox/Inbox` 形式。
- 對同一封 mail 呼叫 `request-mail-body` 完成後，若同 id 的 `body` 與 `bodyHtml` 仍為空，不要重複呼叫同一 endpoint；將該封內容視為目前不可用，回報限制或改用 metadata。
- `request-move-mails` 單次最多 500 個 `mailIds`。遇到「搬移 folderA 和所有 subfolder」這類任務時，必須分批慢慢送，不可把 8000+ ids 放進單一 request。

## API 設計反思

使用此 skill 時若發現 contract 讓 agent 必須猜測、重複查詢、或暴露敏感資料，先把問題回報給使用者；除非使用者明確要求修改 API contract，否則不要擅自更動 API。特別留意：

- `request-*` response 是否足以提供 `requestId`，且 `fetch-result-*` 是否足以表達狀態與資料頁。
- mutation endpoint 是否都要求穩定 id，而不是只靠顯示名稱。
- destructive 或本機檔案操作是否有足夠明確的識別與限制。
- mail search 的進度、結果與 fetch result 是否能被 agent 清楚串起來。
