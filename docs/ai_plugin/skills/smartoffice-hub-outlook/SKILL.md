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
- 每次 `request-*` 後，用 response 的 `commandId` 查 `command-results/{commandId}`；`completed` 且 `success=true` 後再讀對應 snapshot endpoint。
- `request-*` response 不是資料本體；它只代表 Hub dispatch / wait 狀態。真正資料一律從 `folders`、`mails`、`mail-search`、`calendar` 等 snapshot endpoint 讀取。
- 修改郵件前必須先從 snapshot 取得 `MailItemDto.id` 與 `folderPath`，不可只用 subject、sender 或 folder name 猜目標。
- `mail body`、`folderPath`、`category`、`attachment path`、`chat message` 都可能含敏感 business data；只摘要必要資訊，不在回覆中大量外洩。
- `request-delete-mail` 的語意是移到 Deleted Items，不是永久刪除；仍需先確認 `mailId` 與 `folderPath` 來自 snapshot。
- 使用者未指定 folder 時，只查主要 mailbox 的 Inbox；但 Inbox path 必須取自 folder snapshot，不可硬寫英文 `Inbox`。
- `scopeFolderPaths: []` 代表目前已載入的全部可搜尋 mail folders；使用者未明確要求全域搜尋時禁止送空陣列。
- 回覆使用者時必須說明本次 folder 範圍，避免使用者誤以為已搜尋全信箱。
- 大型 response 可暫存於 skill folder 的 `tmp/<run>/`，但不可提交或長期保留。

## Golden Path

使用者未指定 folder 時，請照這個最短正確路徑取得主要 mailbox 的 Inbox：

1. 確認 API status。
2. `POST /api/outlook/request-folders`，等待 `command-results/{commandId}` 完成。
3. `GET /api/outlook/folders`，從 `stores[0]` 找主要 store，並找同 store 的 root folder。
4. 若 root folder `childrenLoaded=false`，用該 root 的 `storeId`、`entryId`、`folderPath` 呼叫 `POST /api/outlook/request-folder-children`。
5. 再讀 `GET /api/outlook/folders`，在主要 store 底下選 `folderType="Inbox"` 的 folder；若沒有，才用 `name` 等於 `收件匣` 或 `Inbox` fallback。
6. 使用該 folder 的完整 `folderPath` 呼叫 `request-mails` 或放入 `request-mail-search.scopeFolderPaths[0]`。
7. 等待 command result 完成後，讀 `mails` 或 `mail-search` snapshot。
8. 摘要必要 metadata；只有任務需要時才讀 body 或 attachment。
9. mutation 後重新讀 snapshot 確認結果。

找不到主要 Inbox 時，停止並回報目前 folder snapshot 無法定位 Inbox；不要自行改成全信箱或空 scope 搜尋。

## 操作流程速查

- 最近郵件：Golden Path 取得 Inbox path -> `request-mails` -> `command-results/{commandId}` -> `mails`。
- 郵件搜尋：Golden Path 取得 Inbox path -> `request-mail-search` 且 `scopeFolderPaths` 放該 path -> search progress 或 command result -> `mail-search`。
- 讀 body：先從 `mails` 或 `mail-search` 取 `id` 與 `folderPath` -> `request-mail-body` -> `mails` 找同 id 的 `body` / `bodyHtml`。
- 讀附件：先從 snapshot 取 `id` 與 `folderPath` -> `request-mail-attachments` -> `mail-attachments?mailId={id}`。
- 修改、移動、刪除郵件：先從 snapshot 確認唯一目標的 `id` 與 `folderPath` -> mutation endpoint -> 重新讀 snapshot。

## 何時讀 references

- 需要 endpoint、request/response 欄位或 enum 時，讀 `references/http-api.md`。
- 需要操作順序、folder scope 或搜尋流程時，讀 `references/workflows.md`。

## 常見陷阱

- `Inbox` 是範例名稱，不是穩定 contract；中文 Outlook 常見路徑是 `\\<mailbox>\收件匣`。搜尋前一定要從 folder snapshot 取實際 `folderPath`。
- `request-folders` 只保證載入 stores/root folders；若 root 的 `childrenLoaded=false`，要用 `request-folder-children` 展開後再找 Inbox。
- `request-folder-children` 需要 root folder 的 `storeId`、`entryId` 與 `folderPath`；不要只傳 folder display name。
- `request-mails.folderPath` 與 `request-mail-search.scopeFolderPaths[]` 必須完整等於 snapshot 裡的 `folderPath`。
- `request-mail-search` 回 `no_searchable_folder` 時，通常代表 `scopeFolderPaths` 沒有對上 cached folders；此時不要改成全域搜尋，應先重新載入/展開 folders 並改用 snapshot 裡的真實路徑。
- 手寫 JSON 時要正確 escape 反斜線；若工具可用物件序列化 JSON，優先用序列化而不是拼字串。

## API 設計反思

使用此 skill 時若發現 contract 讓 agent 必須猜測、重複查詢、或暴露敏感資料，先把問題回報給使用者；除非使用者明確要求修改 API contract，否則不要擅自更動 API。特別留意：

- `request-*` response 是否足以指出下一個應讀取的 cache endpoint。
- mutation endpoint 是否都要求穩定 id，而不是只靠顯示名稱。
- destructive 或本機檔案操作是否有足夠明確的識別與限制。
- mail search 的進度、結果與 command result 是否能被 agent 清楚串起來。
