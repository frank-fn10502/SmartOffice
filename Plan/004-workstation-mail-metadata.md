# Task 004：補齊 Mail Metadata

## 新 Session 起手

本任務可以在全新 session 單獨執行。請先讀：

1. `AGENTS.md`
2. `Plan/000-session-handoff.md`
3. `docs/ai/office2016-workstation-contract.md`
4. `Models/Dtos.cs`
5. 本檔

不要假設工作機已完成 rules 或 calendar。只補 mail metadata mapping。

## 目標

讓工作機推送 mails 時，除了 body，也回傳 Web UI 與 AI 操作需要的 metadata。

## 需要補的欄位

`MailItemDto` 目前包含：

```json
{
  "id": "...",
  "subject": "...",
  "senderName": "...",
  "senderEmail": "...",
  "receivedTime": "...",
  "body": "...",
  "bodyHtml": "...",
  "folderPath": "...",
  "categories": "Customer, Follow-up",
  "isRead": false,
  "isMarkedAsTask": true,
  "importance": "high",
  "sensitivity": "normal"
}
```

## 建議實作步驟

1. 在工作機既有 fetch mail flow 找到轉換 `MailItem` 的地方。
2. `id` 優先使用 Outlook `EntryID`；如果不穩定或取不到，先留空。
3. `categories` 對應 Outlook `Categories`。
4. `isRead` 可由 `UnRead` 反向轉換。
5. `isMarkedAsTask` 對應 `IsMarkedAsTask`。
6. `importance` 轉成 `low`、`normal`、`high`。
7. `sensitivity` 轉成 `normal`、`personal`、`private`、`confidential`。
8. 保持既有 `body` 與 `bodyHtml` 欄位，不要 rename。

## 注意事項

- 如果某欄位取不到，使用預設值，不要讓整批 mails 失敗。
- 不要把完整 mail body 寫進 log。
- `bodyHtml` 是 best-effort，不保證保留 Outlook 視覺樣式。

## 驗證

1. Web UI Fetch Mails。
2. 切到 `Outlook` 分頁。
3. 確認 unread、flagged、importance、categories 統計有合理數字。
4. 檢查 Hub admin log 是否有 mapping error。

## 完成回報

請回報每個 metadata 欄位對應的 Outlook property、取不到時的預設值，以及匿名化測試結果。
