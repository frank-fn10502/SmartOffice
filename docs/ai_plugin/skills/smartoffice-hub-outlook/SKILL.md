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
- 預設 base URL 是 `http://localhost:2805`；若環境有不同位置，可用 `SMARTOFFICE_OUTLOOK_URL` 覆寫。
- 每次 `POST /api/outlook/request-*` 後，都要用 response 的 `commandId` 查 `GET /api/outlook/command-results/{commandId}`，直到 `status` 不是 `pending`。
- HTTP 200 只代表 API 完成 dispatch / wait 流程，不保證 cached snapshot 已是呼叫者想像中的最新資料；完成後再讀對應 cache endpoint。
- 修改郵件前必須先從 snapshot 取得 `MailItemDto.id`，不可只用 subject、sender 或 folder name 猜目標。
- `mail body`、`folderPath`、`category`、`attachment path`、`chat message` 都可能含敏感 business data；只摘要必要資訊，不在回覆中大量外洩。
- `request-delete-mail` 的語意是移到 Deleted Items，不是永久刪除；仍需先確認 `mailId` 與 `folderPath` 來自 snapshot。
- 使用者沒有明確指定 folder / 全域搜尋 / 移動或刪除 folder 時，一律以目前主要 mailbox 的 Inbox 為範圍；但不可硬寫英文 `Inbox` 路徑，必須先從 `GET /api/outlook/folders` 的 snapshot 找到實際 `folderType=Inbox` 或使用者語系中的收件匣路徑。
- 不要主動搜尋其他 folder；只有使用者明確要求其他 folder、子資料夾或全域搜尋時，才擴大 scope。
- 回覆使用者時必須說明本次使用的 folder 範圍，例如「範圍：主要 mailbox 的 Inbox」或「範圍：指定 folder `...`」，避免使用者誤以為已搜尋全信箱。
- API response 太大時，可在本 skill folder 底下建立 `tmp/<timestamp>-<shortid>/` 暫存 JSON，方便查找與過濾；此資料可能含敏感 business data，不可提交、不可長期保留。

## 操作流程

1. 讀 `GET /api/outlook/admin/status` 確認 Outlook API 是否可用。
2. 使用者未指定 folder 時，先 `request-folders` 載入 stores/root，再展開主要 mailbox root children，從 snapshot 找出實際 Inbox folderPath；不要直接猜 `\\Mailbox - User\Inbox`。
3. 查詢 Inbox 郵件時，呼叫 `POST /api/outlook/request-mails`，再用 response 的 `commandId` 查 `GET /api/outlook/command-results/{commandId}`。
4. 需要條件搜尋時，呼叫 `POST /api/outlook/request-mail-search`，但 `scopeFolderPaths` 預設只放主要 Inbox；只有使用者明確要求時才設定其他 folder 或全域搜尋。
5. 讀單封 body、attachment 或執行 mutation 前，必須使用 snapshot 中的 `mailId` 與 `folderPath`。
6. 回覆使用者時列出本次 folder 範圍；若只查 Inbox，要明確說不是全域搜尋。
7. 對 mutation 類操作，在回覆使用者前重新讀 snapshot，確認變更結果。

## 暫存 JSON workspace

- 暫存路徑格式：`tmp/<yyyyMMdd-HHmmss>-<shortid>/`，例如 `tmp/20260506-143012-a1b2c3/`。
- 建議檔名：`status.json`、`folders.json`、`mails.json`、`mail-search.json`、`command-<commandId>.json`、`attachments-<mailId>.json`。
- 預設只暫存 metadata response；只有使用者要求摘要或檢視內容時，才暫存 mail body。
- 不要把完整 mail body、attachment path 或 chat message 貼回對話；只摘要必要片段。
- 任務完成後若不再需要，刪除該次 `tmp/<run>/`；若工具環境不方便刪除，至少不可把 `tmp/` 內容加入 commit 或對外分享。
- 回覆使用者時不需要提暫存檔路徑，除非使用者要求 debug 或追溯。

## 何時讀 references

- 需要 endpoint、request/response 欄位或 enum 時，讀 `references/http-api.md`。
- 需要具體 curl workflow、等待 command 或 mail search progress 範例時，讀 `references/workflows.md`。

## 常見陷阱

- `Inbox` 是範例名稱，不是穩定 contract；中文 Outlook 常見路徑是 `\\<mailbox>\收件匣`。搜尋前一定要從 folder snapshot 取實際 `folderPath`。
- `request-folders` 只保證載入 stores/root folders；若 root 的 `childrenLoaded=false`，要用 `request-folder-children` 展開後再找 Inbox。
- `request-mail-search` 回 `no_searchable_folder` 時，通常代表 `scopeFolderPaths` 沒有對上 cached folders；此時不要改成全域搜尋，應先重新載入/展開 folders 並改用 snapshot 裡的真實路徑。
- 在 PowerShell 中不要手寫含反斜線的 JSON 字串給 `curl -d`；用 hashtable/object 搭配 `ConvertTo-Json`，並優先呼叫 `curl.exe` 避免 alias 差異。

## API 設計反思

使用此 skill 時若發現 contract 讓 agent 必須猜測、重複查詢、或暴露敏感資料，先把問題回報給使用者；除非使用者明確要求修改 API contract，否則不要擅自更動 API。特別留意：

- `request-*` response 是否足以指出下一個應讀取的 cache endpoint。
- mutation endpoint 是否都要求穩定 id，而不是只靠顯示名稱。
- destructive 或本機檔案操作是否有足夠明確的識別與限制。
- mail search 的進度、結果與 command result 是否能被 agent 清楚串起來。
