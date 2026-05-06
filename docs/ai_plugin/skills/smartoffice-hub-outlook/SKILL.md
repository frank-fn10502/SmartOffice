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
- 使用者沒有明確指定 folder / 全域搜尋 / 移動或刪除 folder 時，一律以目前主要 mailbox 的 Inbox 為範圍。
- 不要主動搜尋其他 folder；只有使用者明確要求其他 folder、子資料夾或全域搜尋時，才擴大 scope。
- 回覆使用者時必須說明本次使用的 folder 範圍，例如「範圍：主要 mailbox 的 Inbox」或「範圍：指定 folder `...`」，避免使用者誤以為已搜尋全信箱。

## 操作流程

1. 讀 `GET /api/outlook/admin/status` 確認 Outlook API 是否可用。
2. 使用者未指定 folder 時，使用目前主要 mailbox 的 Inbox；不要主動 request 或搜尋其他 folder。
3. 查詢 Inbox 郵件時，呼叫 `POST /api/outlook/request-mails`，再用 response 的 `commandId` 查 `GET /api/outlook/command-results/{commandId}`。
4. 需要條件搜尋時，呼叫 `POST /api/outlook/request-mail-search`，但 `scopeFolderPaths` 預設只放主要 Inbox；只有使用者明確要求時才設定其他 folder 或全域搜尋。
5. 讀單封 body、attachment 或執行 mutation 前，必須使用 snapshot 中的 `mailId` 與 `folderPath`。
6. 回覆使用者時列出本次 folder 範圍；若只查 Inbox，要明確說不是全域搜尋。
7. 對 mutation 類操作，在回覆使用者前重新讀 snapshot，確認變更結果。

## 何時讀 references

- 需要 endpoint、request/response 欄位或 enum 時，讀 `references/http-api.md`。
- 需要具體 curl workflow、等待 command 或 mail search progress 範例時，讀 `references/workflows.md`。

## API 設計反思

使用此 skill 時若發現 contract 讓 agent 必須猜測、重複查詢、或暴露敏感資料，先把問題回報給使用者；除非使用者明確要求修改 API contract，否則不要擅自更動 API。特別留意：

- `request-*` response 是否足以指出下一個應讀取的 cache endpoint。
- mutation endpoint 是否都要求穩定 id，而不是只靠顯示名稱。
- destructive 或本機檔案操作是否有足夠明確的識別與限制。
- mail search 的進度、結果與 command result 是否能被 agent 清楚串起來。
