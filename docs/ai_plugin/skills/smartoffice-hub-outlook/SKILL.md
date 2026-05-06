---
name: smartoffice-hub-outlook
description: Use when an AI agent needs to operate Outlook through SmartOffice.Hub HTTP APIs, including reading folders, mails, mail body, attachments, calendar, categories, rules, mail search, and safe mail or folder mutations. Use this skill for SmartOffice.Hub client workflows; do not use it for implementing the Outlook AddIn SignalR automation itself.
metadata:
  owner: "SmartOffice.Hub"
  skill_id: "smartoffice-hub-outlook.skill.smartoffice-hub.2026-05"
---

# SmartOffice Hub Outlook

## 邊界

- 一律透過 SmartOffice.Hub HTTP API 操作 Outlook；不要直接連 Outlook COM、Office automation 或 `/hub/outlook-addin`。
- Hub 預設 base URL 是 `SMARTOFFICE_HUB_URL`；未設定時可用 `http://localhost:2805`。
- 每次 `POST /api/outlook/request-*` 後，都要用 response 的 `commandId` 查 `GET /api/outlook/command-results/{commandId}`，直到 `status` 不是 `pending`。
- HTTP 200 只代表 Hub 完成 dispatch / wait 流程，不保證 cached snapshot 已是呼叫者想像中的最新資料；完成後再讀對應 cache endpoint。
- 修改郵件前必須先從 Hub snapshot 取得 `MailItemDto.id`，不可只用 subject、sender 或 folder name 猜目標。
- `mail body`、`folderPath`、`category`、`attachment path`、`chat message` 都可能含敏感 business data；只摘要必要資訊，不在回覆中大量外洩。
- `request-delete-mail` 的語意是移到 Deleted Items，不是永久刪除；仍需先確認 `mailId` 與 `folderPath` 來自 snapshot。
- 使用者沒有指定 folder 時，以目前主要 mailbox 的 Inbox 作為起點，並預設 `includeSubFolders=true`。

## 操作流程

1. 讀 `GET /api/outlook/admin/status` 確認 AddIn 是否 connected。
2. 若需要 folder scope，先 `POST /api/outlook/request-folders`，等待 command 完成，再讀 `GET /api/outlook/folders`。
3. 使用者未指定 folder 時，從 folder snapshot 選擇主要 store 的 Inbox；若有多個 Inbox 且無法判斷主要 mailbox，回報候選並請使用者指定。
4. 發出任務所需的 `request-*` endpoint。
5. 輪詢 `command-results/{commandId}`；mail search 可同時查 progress endpoint。
6. 依 command 類型讀取 cache：`mails`、`mail-attachments`、`mail-search`、`folders`、`calendar`、`categories` 或 `rules`。
7. 對 mutation 類操作，在回覆使用者前重新讀 snapshot，確認變更結果。

## 何時讀 references

- 需要 endpoint、request/response 欄位或 enum 時，讀 `references/http-api.md`。
- 需要具體 curl workflow、等待 command 或 mail search progress 範例時，讀 `references/workflows.md`。

## API 設計反思

使用此 skill 時若發現 contract 讓 agent 必須猜測、重複查詢、或暴露敏感資料，先把問題回報給使用者；除非使用者明確要求修改 Hub contract，否則不要擅自更動 Hub API。特別留意：

- `request-*` response 是否足以指出下一個應讀取的 cache endpoint。
- mutation endpoint 是否都要求穩定 id，而不是只靠顯示名稱。
- destructive 或本機檔案操作是否有足夠明確的識別與限制。
- mail search 的進度、結果與 command result 是否能被 agent 清楚串起來。
